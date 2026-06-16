import { createGraphClient } from "./graph.js";
import { discoverSites } from "./scanner.js?v=20260616b";

const DEFAULT_UNSUPPORTED_EXTENSIONS = new Set(["exe", "dll", "iso", "tmp", "zip", "rar", "7z", "bak", "pst"]);

function normalizeConfig(config = {}) {
  return {
    maxSites: config.maxSites || 15,
    unlimitedSites: !!config.unlimitedSites,
    includeOneDrive: !!config.includeOneDrive,
    maxSampleFiles: config.maxSampleFiles || 2000,
    lifecycleDays: config.lifecycleDays || 180
  };
}

function percent(part, total) {
  return total > 0 ? Number(((part / total) * 100).toFixed(1)) : 0;
}

async function getSites(accessToken, config, notes) {
  const siteData = await discoverSites(accessToken, config);
  notes.push(...(siteData.notes || []));
  return siteData.siteList || [];
}

export async function runStep4Sensitivity(accessToken, config, onProgress) {
  const client = createGraphClient(accessToken);
  const notes = [];
  const cfg = normalizeConfig(config);

  onProgress({ phase: "Step 4: Sensitivity", sitesProcessed: 0, filesScanned: 0, currentSite: "", currentLibrary: "" });

  let labels = [];
  let labelsSupported = false;

  try {
    const res = await client.getJson("/security/informationProtection/sensitivityLabels?$top=200");
    labels = res.value || [];
    labelsSupported = true;
    notes.push(`Sensitivity label catalog available via Graph (${labels.length} labels).`);
  } catch (e) {
    notes.push("Sensitivity label catalog endpoint unavailable for this token or tenant. File-level label coverage cannot be measured with current Graph access.");
    notes.push(`Graph response: ${e.message}`);
  }

  return {
    step: 4,
    title: "Sensitivity Labels & Purview Integration",
    feasible: labelsSupported,
    summary: labelsSupported
      ? "Label catalog retrieved. File-level applied-label coverage and auto/manual attribution are not available from this scanner path."
      : "Not fully feasible with current Graph token path. Requires Purview/compliance data plane for coverage and policy attribution.",
    metrics: {
      labelCatalogAvailable: labelsSupported ? "Y" : "N",
      labelCount: labels.length
    },
    rows: labels.map(l => ({
      LabelId: l.id || "",
      LabelName: l.name || l.displayName || "",
      Priority: l.priority || ""
    })),
    notes
  };
}

export async function runStep5SIT(accessToken, config, onProgress) {
  onProgress({ phase: "Step 5: SIT", sitesProcessed: 0, filesScanned: 0, currentSite: "", currentLibrary: "" });
  return {
    step: 5,
    title: "Sensitive Information Types (SITs)",
    feasible: false,
    summary: "SIT detection and policy match data are not exposed through this delegated Graph scanning path.",
    metrics: {
      supportedViaCurrentGraphPath: "N"
    },
    rows: [],
    notes: [
      "Requires Purview DLP/classification telemetry (compliance data plane) to report SIT matches in files and enforcement outcomes.",
      "Graph file metadata endpoints do not return detected SIT hits for arbitrary SharePoint/OneDrive files."
    ]
  };
}

export async function runStep6DLP(accessToken, config, onProgress) {
  onProgress({ phase: "Step 6: DLP", sitesProcessed: 0, filesScanned: 0, currentSite: "", currentLibrary: "" });
  return {
    step: 6,
    title: "Data Loss Prevention (DLP)",
    feasible: false,
    summary: "DLP policy coverage and incidents are not available from this scanner's delegated Graph path.",
    metrics: {
      supportedViaCurrentGraphPath: "N"
    },
    rows: [],
    notes: [
      "Requires Purview DLP policy and incident APIs/reports.",
      "Current scanner can infer oversharing risk, but not DLP enforcement actions."
    ]
  };
}

export async function runStep7InformationArchitecture(accessToken, config, onProgress) {
  const client = createGraphClient(accessToken);
  const notes = [];
  const cfg = normalizeConfig(config);
  const siteList = await getSites(accessToken, cfg, notes);

  let processedSites = 0;
  let librariesChecked = 0;
  let librariesFolderOnly = 0;
  let librariesDeepNesting = 0;
  let customColumnCount = 0;
  let contentTypeEnabledCount = 0;

  const rows = [];

  async function estimateDepth(driveId, itemId, depth, maxDepth) {
    if (depth >= maxDepth) return depth;
    const payload = await client.getJson(`/drives/${encodeURIComponent(driveId)}/items/${encodeURIComponent(itemId)}/children?$top=200`);
    const children = payload.value || [];
    let deepest = depth;
    for (const child of children) {
      if (child.folder) {
        const d = await estimateDepth(driveId, child.id, depth + 1, maxDepth);
        if (d > deepest) deepest = d;
      }
    }
    return deepest;
  }

  for (const site of siteList) {
    const siteName = site.displayName || site.name || site.id;
    processedSites++;
    onProgress({ phase: "Step 7: Information Architecture", sitesProcessed: processedSites, currentSite: siteName, currentLibrary: "" });

    let drives = [];
    try {
      drives = await client.getAllPages(`/sites/${encodeURIComponent(site.id)}/drives?$top=50`, p => p.value || []);
    } catch (e) {
      notes.push(`Could not list drives for '${siteName}': ${e.message}`);
      continue;
    }

    for (const drive of drives) {
      const libName = drive.name || drive.id;
      librariesChecked++;
      onProgress({ phase: "Step 7: Information Architecture", sitesProcessed: processedSites, currentSite: siteName, currentLibrary: libName });

      let topItems = [];
      try {
        topItems = await client.getAllPages(`/drives/${encodeURIComponent(drive.id)}/root/children?$top=200`, p => p.value || []);
      } catch {
        topItems = [];
      }

      const topFiles = topItems.filter(i => i.file).length;
      const topFolders = topItems.filter(i => i.folder).length;
      const folderOnly = topFolders > 0 && topFiles === 0;
      if (folderOnly) librariesFolderOnly++;

      let deepest = 0;
      for (const item of topItems) {
        if (item.folder) {
          try {
            const depth = await estimateDepth(drive.id, item.id, 1, 6);
            if (depth > deepest) deepest = depth;
          } catch {
            // Ignore depth failures for individual folders.
          }
        }
      }
      if (deepest > 5) librariesDeepNesting++;

      const listId = drive.list?.id;
      let customCols = 0;
      let contentTypesEnabled = false;

      if (listId) {
        try {
          const columns = await client.getAllPages(
            `/sites/${encodeURIComponent(site.id)}/lists/${encodeURIComponent(listId)}/columns?$top=200`,
            p => p.value || []
          );
          customCols = columns.filter(c => !c.hidden && !c.readOnly).length;
          customColumnCount += customCols;
        } catch {
          // Column endpoint may be blocked.
        }

        try {
          const cts = await client.getAllPages(
            `/sites/${encodeURIComponent(site.id)}/lists/${encodeURIComponent(listId)}/contentTypes?$top=200`,
            p => p.value || []
          );
          contentTypesEnabled = cts.length > 1;
          if (contentTypesEnabled) contentTypeEnabledCount++;
        } catch {
          // Content type endpoint may be blocked.
        }
      }

      rows.push({
        SiteName: siteName,
        LibraryName: libName,
        TopLevelFiles: topFiles,
        TopLevelFolders: topFolders,
        FolderOnlyPattern: folderOnly ? "Y" : "N",
        DeepestFolderDepthObserved: deepest,
        CustomColumnsObserved: customCols,
        MultipleContentTypes: contentTypesEnabled ? "Y" : "N"
      });
    }
  }

  return {
    step: 7,
    title: "SharePoint Information Architecture",
    feasible: true,
    summary: "Library structure signals collected using SharePoint/Drive metadata from Microsoft Graph.",
    metrics: {
      sitesAssessed: processedSites,
      librariesAssessed: librariesChecked,
      librariesFolderOnly,
      librariesDeepNestingOver5: librariesDeepNesting,
      customColumnsObserved: customColumnCount,
      librariesWithMultipleContentTypes: contentTypeEnabledCount
    },
    rows,
    notes
  };
}

export async function runStep8SearchReadiness(accessToken, config, onProgress) {
  const client = createGraphClient(accessToken);
  const notes = [];
  const cfg = normalizeConfig(config);
  const siteList = await getSites(accessToken, cfg, notes);

  let processedSites = 0;
  let sampledFiles = 0;
  let unsupportedCount = 0;
  let ambiguousCount = 0;
  let missingTitleLikeCount = 0;
  const fileNameCounts = new Map();
  const rows = [];

  const titleLikePattern = /^(document|doc|file|scan|image|copy)\d*\.(docx|xlsx|pptx|pdf|txt)$/i;

  for (const site of siteList) {
    const siteName = site.displayName || site.name || site.id;
    processedSites++;
    onProgress({ phase: "Step 8: Search Readiness", sitesProcessed: processedSites, filesScanned: sampledFiles, currentSite: siteName, currentLibrary: "" });

    let drives = [];
    try {
      drives = await client.getAllPages(`/sites/${encodeURIComponent(site.id)}/drives?$top=20`, p => p.value || []);
    } catch {
      continue;
    }

    for (const drive of drives) {
      if (sampledFiles >= cfg.maxSampleFiles) break;
      const libName = drive.name || drive.id;
      onProgress({ phase: "Step 8: Search Readiness", sitesProcessed: processedSites, filesScanned: sampledFiles, currentSite: siteName, currentLibrary: libName });

      let items = [];
      try {
        items = await client.getAllPages(`/drives/${encodeURIComponent(drive.id)}/root/children?$top=200`, p => p.value || []);
      } catch {
        items = [];
      }

      for (const item of items) {
        if (sampledFiles >= cfg.maxSampleFiles) break;
        if (!item.file) continue;

        sampledFiles++;
        const name = item.name || "";
        const ext = name.includes(".") ? name.split(".").pop().toLowerCase() : "";
        const key = name.toLowerCase();
        fileNameCounts.set(key, (fileNameCounts.get(key) || 0) + 1);

        const looksUntitled = titleLikePattern.test(name);
        if (looksUntitled) missingTitleLikeCount++;
        if (DEFAULT_UNSUPPORTED_EXTENSIONS.has(ext)) unsupportedCount++;

        rows.push({
          SiteName: siteName,
          LibraryName: libName,
          FileName: name,
          LooksGenericName: looksUntitled ? "Y" : "N",
          PotentiallyNonIndexedType: DEFAULT_UNSUPPORTED_EXTENSIONS.has(ext) ? "Y" : "N"
        });
      }
    }

    if (sampledFiles >= cfg.maxSampleFiles) {
      notes.push(`Stopped sampling at maxSampleFiles=${cfg.maxSampleFiles}.`);
      break;
    }
  }

  for (const count of fileNameCounts.values()) {
    if (count > 1) ambiguousCount += count;
  }

  notes.push("Graph does not reliably expose file title/description completeness for every library without expensive list-item field expansion.");

  return {
    step: 8,
    title: "Search Readiness Signals",
    feasible: true,
    summary: "Search-readiness heuristics calculated from sampled file names and extensions.",
    metrics: {
      sitesAssessed: processedSites,
      sampledFiles,
      genericFileNameCount: missingTitleLikeCount,
      duplicateNameCount: ambiguousCount,
      potentiallyNonIndexedTypeCount: unsupportedCount,
      genericFileNamePct: percent(missingTitleLikeCount, sampledFiles),
      duplicateNamePct: percent(ambiguousCount, sampledFiles)
    },
    rows,
    notes
  };
}

export async function runStep9Lifecycle(accessToken, config, onProgress) {
  const client = createGraphClient(accessToken);
  const notes = [];
  const cfg = normalizeConfig(config);
  const siteList = await getSites(accessToken, cfg, notes);

  const rows = [];
  let staleSites = 0;
  let processedSites = 0;
  const thresholdMs = cfg.lifecycleDays * 24 * 60 * 60 * 1000;
  const now = Date.now();

  for (const site of siteList) {
    processedSites++;
    const siteName = site.displayName || site.name || site.id;
    onProgress({ phase: "Step 9: Lifecycle", sitesProcessed: processedSites, currentSite: siteName, currentLibrary: "" });

    let modified = site.lastModifiedDateTime || "";
    if (!modified) {
      try {
        const fullSite = await client.getJson(`/sites/${encodeURIComponent(site.id)}`);
        modified = fullSite.lastModifiedDateTime || "";
      } catch {
        modified = "";
      }
    }

    const ageDays = modified ? Math.floor((now - new Date(modified).getTime()) / (24 * 60 * 60 * 1000)) : "";
    const stale = typeof ageDays === "number" && (now - new Date(modified).getTime()) > thresholdMs;
    if (stale) staleSites++;

    rows.push({
      SiteName: siteName,
      SiteUrl: site.webUrl || "",
      LastModifiedDateTime: modified,
      DaysSinceLastModified: ageDays,
      StaleByThreshold: stale ? "Y" : "N"
    });
  }

  notes.push("Retention policy coverage and record enforcement require Purview/Records Management APIs not included in this scanner path.");
  notes.push("Teams inactivity requires Teams/report telemetry beyond the current SharePoint file scan scope.");

  return {
    step: 9,
    title: "Lifecycle & Records Management",
    feasible: true,
    summary: "Site lifecycle staleness assessed from site last-modified signals; retention/records controls require Purview integration.",
    metrics: {
      sitesAssessed: processedSites,
      staleSitesByThreshold: staleSites,
      staleSitesPct: percent(staleSites, processedSites),
      staleThresholdDays: cfg.lifecycleDays
    },
    rows,
    notes
  };
}
