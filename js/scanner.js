// ── scanner.js ───────────────────────────────────────────────
// Time-boxed, sampling-based Microsoft 365 content scanner.
// Hard stops at time limit OR file caps — always returns partial results.

import { createGraphClient } from "./graph.js";

function ext(name) {
  const i = name.lastIndexOf(".");
  return i === -1 ? "" : name.slice(i + 1).toLowerCase();
}

// Recursively scan a drive folder up to maxFolderDepth levels.
// Stops immediately when any cap is exceeded or time is up.
async function scanFolder(client, driveId, itemId, depth, ctx) {
  if (!ctx.ok()) return;

  const url = itemId === "root"
    ? `/drives/${encodeURIComponent(driveId)}/root/children?$top=200`
    : `/drives/${encodeURIComponent(driveId)}/items/${encodeURIComponent(itemId)}/children?$top=200`;

  let nextUrl = url;

  while (nextUrl) {
    if (!ctx.ok()) return;
    if (ctx.siteCount.v >= ctx.maxPerSite) return;
    if (ctx.total.v >= ctx.maxFiles) return;

    const payload = await client.getJson(nextUrl);
    const items = payload.value || [];

    for (const item of items) {
      if (!ctx.ok()) return;
      if (ctx.siteCount.v >= ctx.maxPerSite) return;
      if (ctx.total.v >= ctx.maxFiles) return;

      if (item.file) {
        ctx.records.push({
          siteName:             ctx.siteName,
          siteUrl:              ctx.siteUrl,
          libraryName:          ctx.libName,
          fileName:             item.name,
          extension:            ext(item.name),
          size:                 item.size || 0,
          createdDateTime:      item.createdDateTime || "",
          lastModifiedDateTime: item.lastModifiedDateTime || ""
        });
        ctx.siteCount.v++;
        ctx.total.v++;
        ctx.tick();
      } else if (item.folder && depth < ctx.cfg.maxFolderDepth) {
        await scanFolder(client, driveId, item.id, depth + 1, ctx);
      }
    }

    nextUrl = payload["@odata.nextLink"] || null;
  }
}

// ── Main entry ────────────────────────────────────────────────
export async function runScan(accessToken, config, onProgress) {
  const client  = createGraphClient(accessToken);
  const startMs = Date.now();
  const maxMs   = config.skipTimeLimit ? Infinity : (config.maxRuntimeMinutes * 60 * 1000);
  const maxSites = config.unlimitedSites ? Infinity : config.maxSites;
  const maxFiles = config.unlimitedFiles ? Infinity : config.maxTotalFiles;
  const maxPerSite = config.unlimitedFiles ? Infinity : config.maxFilesPerSite;
  
  const records = [];
  const notes   = [];
  const total   = { v: 0 };
  let   sites   = 0;
  let   timeLimitReached = false;
  let   lastEmit = 0;

  // Returns false when we must stop
  function ok() {
    if (!config.skipTimeLimit && Date.now() - startMs > maxMs) { timeLimitReached = true; return false; }
    if (total.v >= maxFiles) return false;
    return true;
  }

  function emit(phase, site = "", lib = "", force = false) {
    const now = Date.now();
    if (!force && now - lastEmit < 1500) return;
    lastEmit = now;
    const elapsed = (now - startMs) / 1000;
    onProgress({
      phase,
      elapsedSeconds:   elapsed,
      remainingSeconds: Math.max(0, maxMs / 1000 - elapsed),
      sitesProcessed:   sites,
      filesScanned:     total.v,
      filesPerSecond:   elapsed > 0 ? total.v / elapsed : 0,
      currentSite:      site,
      currentLibrary:   lib,
      timeLimitReached
    });
  }

  // ── Phase 1: Site Discovery ───────────────────────────────
  emit("Discovery", "", "", true);

  const terms = config.searchTerms && config.searchTerms.length
    ? config.searchTerms
    : ["team", "project", "shared"];

  const siteMap = new Map();

  for (const term of terms) {
    if (!ok()) break;
    try {
      const found = await client.getAllPages(
        `/sites?search=${encodeURIComponent(term)}&$top=25`,
        p => p.value || []
      );
      for (const s of found) {
        if (s.id && !siteMap.has(s.id)) siteMap.set(s.id, s);
        if (siteMap.size >= maxSites) break;
      }
    } catch (e) {
      notes.push(`Discovery term '${term}' failed: ${e.message}`);
    }
    if (siteMap.size >= maxSites) break;
  }

  const siteList = [...siteMap.values()].slice(0, maxSites);
  if (!siteList.length) {
    notes.push(config.includeOneDrive
      ? "No SharePoint sites found — will try OneDrive only."
      : "No SharePoint sites found for the given search terms.");
  }

  // ── Phase 2: Scan Each Site ───────────────────────────────
  emit("Scanning", "", "", true);

  for (const site of siteList) {
    if (!ok()) break;

    const siteName = site.displayName || site.name || site.id;
    const siteUrl  = site.webUrl || "";
    let   drives   = [];

    try {
      drives = await client.getAllPages(
        `/sites/${encodeURIComponent(site.id)}/drives?$top=50`,
        p => p.value || []
      );
    } catch (e) {
      notes.push(`Could not list drives for '${siteName}': ${e.message}`);
    }

    for (const drive of drives) {
      if (!ok()) break;
      const libName   = drive.name || drive.id;
      const siteCount = { v: 0 };
      emit("Scanning", siteName, libName);

      try {
        await scanFolder(client, drive.id, "root", 0, {
          siteName, siteUrl, libName,
          cfg: config, records, siteCount, total,
          maxPerSite, maxFiles,
          ok, tick: () => emit("Scanning", siteName, libName)
        });
      } catch (e) {
        notes.push(`Scan failed for '${siteName}/${libName}': ${e.message}`);
      }
    }

    sites++;
    emit("Scanning", siteName, "", true);
  }

  // ── Optional: OneDrive ────────────────────────────────────
  if (config.includeOneDrive && ok()) {
    try {
      const meDrive   = await client.getJson("/me/drive");
      const siteName  = "OneDrive";
      const siteUrl   = meDrive.webUrl || "";
      const libName   = meDrive.name || "OneDrive";
      const siteCount = { v: 0 };
      emit("Scanning", siteName, libName, true);

      await scanFolder(client, meDrive.id, "root", 0, {
        siteName, siteUrl, libName,
        cfg: config, records, siteCount, total,
        maxPerSite, maxFiles,
        ok, tick: () => emit("Scanning", siteName, libName)
      });
      sites++;
    } catch (e) {
      notes.push(`OneDrive scan failed: ${e.message}`);
    }
  }

  if (timeLimitReached) notes.push("Time limit reached — returning partial results.");

  emit("Aggregating", "", "", true);

  return {
    records,
    notes,
    sitesProcessed: sites,
    timeLimitReached,
    startedAt:  new Date(startMs).toISOString(),
    finishedAt: new Date().toISOString()
  };
}
