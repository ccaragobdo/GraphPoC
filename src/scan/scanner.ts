import { GraphClient } from "../graph/client";
import {
  FileRecord,
  ScanConfig,
  ScanProgress,
  ScanResult,
  ScanSummary
} from "./types";
import { summarizeInventory } from "../analysis/aggregate";

interface GraphSite {
  id: string;
  name?: string;
  webUrl?: string;
}

interface GraphDrive {
  id: string;
  name?: string;
}

interface GraphDriveItem {
  id: string;
  name: string;
  size?: number;
  createdDateTime?: string;
  lastModifiedDateTime?: string;
  file?: { mimeType?: string };
  folder?: { childCount?: number };
}

interface WorkerContext {
  site: GraphSite;
  drives: GraphDrive[];
}

function nowIso(): string {
  return new Date().toISOString();
}

function getExtension(fileName: string): string {
  const idx = fileName.lastIndexOf(".");
  return idx === -1 ? "" : fileName.slice(idx + 1).toLowerCase();
}

function safeDate(value?: string): string {
  return value || "";
}

function clamp(value: number, min: number, max: number): number {
  return Math.max(min, Math.min(max, value));
}

export async function getSitesToScan(
  client: GraphClient,
  config: ScanConfig,
  shouldContinue: () => boolean,
  notes: string[]
): Promise<GraphSite[]> {
  const terms = config.searchTerms.length > 0 ? config.searchTerms : ["team", "project", "shared"];
  const dedupe = new Map<string, GraphSite>();

  for (const term of terms) {
    if (!shouldContinue()) {
      break;
    }

    try {
      const sites = await client.getAllPages<GraphSite>(
        `/sites?search=${encodeURIComponent(term)}&$top=25`,
        (payload) => payload.value ?? []
      );

      for (const site of sites) {
        if (site.id && !dedupe.has(site.id)) {
          dedupe.set(site.id, site);
        }

        if (dedupe.size >= config.maxSites) {
          break;
        }
      }
    } catch (error) {
      notes.push(`Site discovery term '${term}' failed: ${(error as Error).message}`);
    }

    if (dedupe.size >= config.maxSites) {
      break;
    }
  }

  return Array.from(dedupe.values()).slice(0, config.maxSites);
}

export async function getDrivesForSite(client: GraphClient, siteId: string): Promise<GraphDrive[]> {
  const drives = await client.getAllPages<GraphDrive>(
    `/sites/${encodeURIComponent(siteId)}/drives?$top=50`,
    (payload) => payload.value ?? []
  );

  return drives;
}

function buildProgress(
  phase: ScanProgress["phase"],
  startMs: number,
  maxRuntimeSeconds: number,
  filesScanned: number,
  sitesProcessed: number,
  timeLimitReached: boolean,
  currentSite?: string,
  currentLibrary?: string
): ScanProgress {
  const elapsedSeconds = (Date.now() - startMs) / 1000;
  const remainingSeconds = Math.max(0, maxRuntimeSeconds - elapsedSeconds);
  const filesPerSecond = elapsedSeconds > 0 ? filesScanned / elapsedSeconds : 0;

  return {
    phase,
    elapsedSeconds,
    remainingSeconds,
    sitesProcessed,
    filesScanned,
    filesPerSecond,
    currentSite,
    currentLibrary,
    timeLimitReached
  };
}

async function scanChildren(
  client: GraphClient,
  driveId: string,
  path: "root" | string,
  depth: number,
  ctx: {
    siteName: string;
    siteUrl: string;
    libraryName: string;
    config: ScanConfig;
    records: FileRecord[];
    siteFileCounter: { value: number };
    globalFileCounter: { value: number };
    shouldContinue: () => boolean;
    progressTick: () => void;
  }
): Promise<void> {
  if (!ctx.shouldContinue()) {
    return;
  }

  if (ctx.siteFileCounter.value >= ctx.config.maxFilesPerSite || ctx.globalFileCounter.value >= ctx.config.maxTotalFiles) {
    return;
  }

  const endpoint =
    path === "root"
      ? `/drives/${encodeURIComponent(driveId)}/root/children?$top=200`
      : `/drives/${encodeURIComponent(driveId)}/items/${encodeURIComponent(path)}/children?$top=200`;

  let nextUrl: string | undefined = endpoint;

  while (nextUrl && ctx.shouldContinue()) {
    if (ctx.siteFileCounter.value >= ctx.config.maxFilesPerSite || ctx.globalFileCounter.value >= ctx.config.maxTotalFiles) {
      return;
    }

    const payload = await client.getJson<{ value?: GraphDriveItem[]; "@odata.nextLink"?: string }>(nextUrl);
    const items = payload.value ?? [];

    for (const item of items) {
      if (!ctx.shouldContinue()) {
        return;
      }

      if (item.file) {
        if (ctx.siteFileCounter.value >= ctx.config.maxFilesPerSite || ctx.globalFileCounter.value >= ctx.config.maxTotalFiles) {
          return;
        }

        ctx.records.push({
          siteName: ctx.siteName,
          siteUrl: ctx.siteUrl,
          libraryName: ctx.libraryName,
          fileName: item.name,
          extension: getExtension(item.name),
          size: item.size ?? 0,
          createdDateTime: safeDate(item.createdDateTime),
          lastModifiedDateTime: safeDate(item.lastModifiedDateTime)
        });

        ctx.siteFileCounter.value += 1;
        ctx.globalFileCounter.value += 1;
        ctx.progressTick();
        continue;
      }

      if (item.folder && depth < ctx.config.maxFolderDepth) {
        await scanChildren(client, driveId, item.id, depth + 1, ctx);
      }
    }

    nextUrl = payload["@odata.nextLink"];
  }
}

async function runWithConcurrency<T>(
  items: T[],
  concurrency: number,
  worker: (item: T) => Promise<void>
): Promise<void> {
  const queue = [...items];
  const workerCount = clamp(concurrency, 1, 5);

  const runners = Array.from({ length: workerCount }).map(async () => {
    while (queue.length > 0) {
      const next = queue.shift();
      if (!next) {
        return;
      }

      await worker(next);
    }
  });

  await Promise.all(runners);
}

export async function runSamplingScan(
  client: GraphClient,
  config: ScanConfig,
  onProgress: (progress: ScanProgress) => void
): Promise<ScanResult> {
  const startedAt = nowIso();
  const startMs = Date.now();
  const maxRuntimeSeconds = config.maxRuntimeMinutes * 60;
  const notes: string[] = [];

  const records: FileRecord[] = [];
  const filesScanned = { value: 0 };
  const sitesProcessed = { value: 0 };
  const timeLimitReached = { value: false };
  let phase: ScanProgress["phase"] = "Discovery";
  let lastCurrentSite = "";
  let lastCurrentLibrary = "";
  let lastProgressEmit = 0;

  const shouldContinue = (): boolean => {
    const elapsedSeconds = (Date.now() - startMs) / 1000;

    if (elapsedSeconds > maxRuntimeSeconds) {
      timeLimitReached.value = true;
      return false;
    }

    if (filesScanned.value >= config.maxTotalFiles) {
      return false;
    }

    return true;
  };

  const emitProgress = (force = false): void => {
    const now = Date.now();
    if (!force && now - lastProgressEmit < 2000) {
      return;
    }

    lastProgressEmit = now;
    onProgress(
      buildProgress(
        phase,
        startMs,
        maxRuntimeSeconds,
        filesScanned.value,
        sitesProcessed.value,
        timeLimitReached.value,
        lastCurrentSite,
        lastCurrentLibrary
      )
    );
  };

  emitProgress(true);

  let sites: GraphSite[] = [];
  try {
    sites = await getSitesToScan(client, config, shouldContinue, notes);
  } catch (error) {
    notes.push(`Site discovery failed: ${(error as Error).message}`);
  }

  if (sites.length === 0) {
    if (config.includeOneDrive) {
      notes.push("No SharePoint sites discovered; proceeding with OneDrive-only scan if accessible.");
    } else {
      notes.push("No sites found for the provided discovery search terms.");
    }
  }

  const contexts: WorkerContext[] = [];

  for (const site of sites.slice(0, config.maxSites)) {
    if (!shouldContinue()) {
      break;
    }

    try {
      const drives = await getDrivesForSite(client, site.id);
      contexts.push({ site, drives });
    } catch (error) {
      notes.push(`Drive enumeration failed for site '${site.name ?? site.id}': ${(error as Error).message}`);
    }
  }

  if (config.includeOneDrive && shouldContinue()) {
    try {
      const meDrive = await client.getJson<{ id: string; name?: string; webUrl?: string }>("/me/drive");
      contexts.push({
        site: { id: "me-drive", name: "OneDrive", webUrl: meDrive.webUrl ?? "https://onedrive.live.com" },
        drives: [{ id: meDrive.id, name: meDrive.name ?? "OneDrive" }]
      });
    } catch (error) {
      notes.push(`OneDrive enumeration failed: ${(error as Error).message}`);
    }
  }

  phase = "Scanning";
  emitProgress(true);

  await runWithConcurrency(contexts, config.siteConcurrency, async (ctx) => {
    if (!shouldContinue()) {
      return;
    }

    const siteFileCounter = { value: 0 };
    const siteName = ctx.site.name ?? ctx.site.id;
    const siteUrl = ctx.site.webUrl ?? "";

    for (const drive of ctx.drives) {
      if (!shouldContinue()) {
        break;
      }

      if (siteFileCounter.value >= config.maxFilesPerSite || filesScanned.value >= config.maxTotalFiles) {
        break;
      }

      lastCurrentSite = siteName;
      lastCurrentLibrary = drive.name ?? drive.id;
      emitProgress();

      try {
        await scanChildren(client, drive.id, "root", 0, {
          siteName,
          siteUrl,
          libraryName: drive.name ?? drive.id,
          config,
          records,
          siteFileCounter,
          globalFileCounter: filesScanned,
          shouldContinue,
          progressTick: () => emitProgress(false)
        });
      } catch (error) {
        notes.push(
          `Drive scan failed for ${siteName} / ${drive.name ?? drive.id}: ${(error as Error).message}`
        );
      }
    }

    sitesProcessed.value += 1;
    emitProgress();
  });

  if (timeLimitReached.value) {
    notes.push("Time limit reached — returning partial results.");
  }

  phase = "Aggregating";
  emitProgress(true);

  const summary: ScanSummary = summarizeInventory(records);

  phase = "Complete";
  const finalProgress = buildProgress(
    phase,
    startMs,
    maxRuntimeSeconds,
    filesScanned.value,
    sitesProcessed.value,
    timeLimitReached.value,
    lastCurrentSite,
    lastCurrentLibrary
  );
  onProgress(finalProgress);

  return {
    inventory: records,
    summary,
    progress: finalProgress,
    startedAt,
    finishedAt: nowIso(),
    notes
  };
}
