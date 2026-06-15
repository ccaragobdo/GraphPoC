// ── scanner.js ───────────────────────────────────────────────
// Time-boxed, sampling-based Microsoft 365 content scanner.
// Hard stops at time limit OR file caps — always returns partial results.

import { createGraphClient } from "./graph.js";

function ext(name) {
  const i = name.lastIndexOf(".");
  return i === -1 ? "" : name.slice(i + 1).toLowerCase();
}

function makeSiteKey(siteName, siteUrl) {
  return `${siteName}||${siteUrl}`;
}

function ensureSiteExposure(exposureMap, siteName, siteUrl) {
  const key = makeSiteKey(siteName, siteUrl);
  if (!exposureMap.has(key)) {
    exposureMap.set(key, {
      siteName,
      siteUrl,
      siteScopesChecked: 0,
      libraryScopesChecked: 0,
      folderScopesChecked: 0,
      permissionsChecked: 0,
      anyoneLinks: 0,
      organizationLinks: 0,
      everyoneAllUsersGrants: 0,
      resourcesWithAnyoneLinks: 0,
      resourcesWithEveryoneAllUsers: 0
    });
  }
  return exposureMap.get(key);
}

function collectPrincipalNames(permission) {
  const names = [];
  const identities = [];

  if (Array.isArray(permission.grantedToIdentitiesV2)) {
    identities.push(...permission.grantedToIdentitiesV2);
  }
  if (permission.grantedToV2) {
    identities.push(permission.grantedToV2);
  }
  if (permission.grantedTo) {
    identities.push(permission.grantedTo);
  }

  for (const identity of identities) {
    const directName = identity?.displayName;
    const userName = identity?.user?.displayName;
    const groupName = identity?.group?.displayName;
    const siteGroupName = identity?.siteGroup?.displayName;

    if (directName) names.push(directName);
    if (userName) names.push(userName);
    if (groupName) names.push(groupName);
    if (siteGroupName) names.push(siteGroupName);
  }

  return names;
}

function isEveryoneOrAllUsersName(name) {
  const value = String(name || "").toLowerCase();
  return value.includes("everyone") || value.includes("all users");
}

async function collectExposureForEndpoint(client, endpoint, exposure) {
  const permissions = await client.getAllPages(
    `${endpoint}?$top=200`,
    p => p.value || []
  );

  let resourceHasAnyoneLink = false;
  let resourceHasEveryoneGroup = false;

  for (const permission of permissions) {
    exposure.permissionsChecked += 1;

    const scope = String(permission?.link?.scope || "").toLowerCase();
    if (scope === "anonymous") {
      exposure.anyoneLinks += 1;
      resourceHasAnyoneLink = true;
    }
    if (scope === "organization") {
      exposure.organizationLinks += 1;
    }

    const names = collectPrincipalNames(permission);
    if (names.some(isEveryoneOrAllUsersName)) {
      exposure.everyoneAllUsersGrants += 1;
      resourceHasEveryoneGroup = true;
    }
  }

  if (resourceHasAnyoneLink) {
    exposure.resourcesWithAnyoneLinks += 1;
  }
  if (resourceHasEveryoneGroup) {
    exposure.resourcesWithEveryoneAllUsers += 1;
  }
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
        if (depth === 0) {
          try {
            await collectExposureForEndpoint(
              client,
              `/drives/${encodeURIComponent(driveId)}/items/${encodeURIComponent(item.id)}/permissions`,
              ctx.siteExposure
            );
            ctx.siteExposure.folderScopesChecked += 1;
          } catch (e) {
            ctx.notes.push(
              `Top-level folder permission check failed for '${ctx.siteName}/${ctx.libName}/${item.name}': ${e.message}`
            );
          }
        }

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
  const exposureBySite = new Map();
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

  const siteMap = new Map();

  async function collectSitesFromEndpoint(url) {
    const found = await client.getAllPages(url, p => p.value || []);
    for (const s of found) {
      if (s.id && !siteMap.has(s.id)) {
        siteMap.set(s.id, s);
      }
      if (siteMap.size >= maxSites) {
        break;
      }
    }
  }

  let usedGetAllSites = false;
  let usedFallback = "";

  try {
    await collectSitesFromEndpoint("/sites/getAllSites?$top=999");
    usedGetAllSites = true;
  } catch {
    notes.push("getAllSites is unavailable for this token; using fallback site discovery.");
  }

  if (!usedGetAllSites && siteMap.size < maxSites) {
    const fallbackEndpoints = [
      { name: "sites-list", url: "/sites?$top=999" },
      { name: "followed-sites", url: "/me/followedSites?$top=999" },
      { name: "site-search", url: "/sites?search=*&$top=999" }
    ];

    for (const fallback of fallbackEndpoints) {
      if (siteMap.size >= maxSites) {
        break;
      }

      try {
        await collectSitesFromEndpoint(fallback.url);
        usedFallback = fallback.name;
      } catch {
        // Try the next fallback endpoint.
      }
    }

    if (usedFallback) {
      notes.push(`Used ${usedFallback} fallback discovery path.`);
    }
  }

  const siteList = [...siteMap.values()].slice(0, maxSites);
  notes.push(`Discovered ${siteList.length} SharePoint site(s) before scanning drives.`);
  if (!siteList.length) {
    notes.push(config.includeOneDrive
      ? "No SharePoint sites found — will try OneDrive only."
      : "No SharePoint sites found for this token.");
  }

  // ── Phase 2: Scan Each Site ───────────────────────────────
  emit("Scanning", "", "", true);

  for (const site of siteList) {
    if (!ok()) break;

    const siteName = site.displayName || site.name || site.id;
    const siteUrl  = site.webUrl || "";
    const siteExposure = ensureSiteExposure(exposureBySite, siteName, siteUrl);
    let   drives   = [];

    try {
      await collectExposureForEndpoint(
        client,
        `/sites/${encodeURIComponent(site.id)}/permissions`,
        siteExposure
      );
      siteExposure.siteScopesChecked += 1;
    } catch (e) {
      notes.push(`Site permission check failed for '${siteName}': ${e.message}`);
    }

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
        await collectExposureForEndpoint(
          client,
          `/drives/${encodeURIComponent(drive.id)}/root/permissions`,
          siteExposure
        );
        siteExposure.libraryScopesChecked += 1;
      } catch (e) {
        notes.push(`Library root permission check failed for '${siteName}/${libName}': ${e.message}`);
      }

      try {
        await scanFolder(client, drive.id, "root", 0, {
          siteName, siteUrl, libName,
          cfg: config, records, notes, siteExposure, siteCount, total,
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
      const siteExposure = ensureSiteExposure(exposureBySite, siteName, siteUrl);
      const libName   = meDrive.name || "OneDrive";
      const siteCount = { v: 0 };
      emit("Scanning", siteName, libName, true);

      try {
        await collectExposureForEndpoint(
          client,
          `/drives/${encodeURIComponent(meDrive.id)}/root/permissions`,
          siteExposure
        );
        siteExposure.libraryScopesChecked += 1;
      } catch (e) {
        notes.push(`OneDrive library root permission check failed: ${e.message}`);
      }

      await scanFolder(client, meDrive.id, "root", 0, {
        siteName, siteUrl, libName,
        cfg: config, records, notes, siteExposure, siteCount, total,
        maxPerSite, maxFiles,
        ok, tick: () => emit("Scanning", siteName, libName)
      });
      sites++;
    } catch (e) {
      notes.push(`OneDrive scan failed: ${e.message}`);
    }
  }

  if (timeLimitReached) notes.push("Time limit reached — returning partial results.");
  notes.push("Sharing exposure checks are limited to site scope, library root scope, and top-level folders.");

  emit("Aggregating", "", "", true);

  return {
    records,
    notes,
    siteExposureBySite: [...exposureBySite.values()],
    sitesProcessed: sites,
    timeLimitReached,
    startedAt:  new Date(startMs).toISOString(),
    finishedAt: new Date().toISOString()
  };
}
