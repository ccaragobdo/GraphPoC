// ── analysis.js ──────────────────────────────────────────────
// Staleness aggregation and CSV/JSON export.
// No dependencies — pure browser APIs.

const DAY_MS = 86_400_000;

function staleness(lastModifiedDateTime) {
  if (!lastModifiedDateTime) return { s90: 0, s180: 0, s365: 0 };
  const age = (Date.now() - new Date(lastModifiedDateTime).getTime()) / DAY_MS;
  return {
    s90:  age > 90  ? 1 : 0,
    s180: age > 180 ? 1 : 0,
    s365: age > 365 ? 1 : 0
  };
}

export function summarize(records) {
  const bySiteMap = new Map();
  const byLibMap  = new Map();
  const byExtMap  = new Map();
  let totalSize = 0, s90 = 0, s180 = 0, s365 = 0;

  for (const r of records) {
    totalSize += r.size;
    const st = staleness(r.lastModifiedDateTime);
    s90 += st.s90; s180 += st.s180; s365 += st.s365;

    // By site
    const sk = `${r.siteName}||${r.siteUrl}`;
    if (!bySiteMap.has(sk)) bySiteMap.set(sk, { siteName: r.siteName, siteUrl: r.siteUrl, files: 0, totalSize: 0, stale90: 0, stale180: 0, stale365: 0 });
    const sv = bySiteMap.get(sk);
    sv.files++; sv.totalSize += r.size; sv.stale90 += st.s90; sv.stale180 += st.s180; sv.stale365 += st.s365;

    // By library
    const lk = `${r.siteName}||${r.siteUrl}||${r.libraryName}`;
    if (!byLibMap.has(lk)) byLibMap.set(lk, { siteName: r.siteName, siteUrl: r.siteUrl, libraryName: r.libraryName, files: 0, totalSize: 0, stale90: 0, stale180: 0, stale365: 0 });
    const lv = byLibMap.get(lk);
    lv.files++; lv.totalSize += r.size; lv.stale90 += st.s90; lv.stale180 += st.s180; lv.stale365 += st.s365;

    // By extension
    const ex = r.extension || "(none)";
    if (!byExtMap.has(ex)) byExtMap.set(ex, { extension: ex, files: 0, totalSize: 0 });
    const ev = byExtMap.get(ex);
    ev.files++; ev.totalSize += r.size;
  }

  const n = records.length;
  const pct = v => n ? (v / n * 100).toFixed(1) : "0.0";

  return {
    totalFiles: n,
    totalSize,
    stale90Pct:  pct(s90),
    stale180Pct: pct(s180),
    stale365Pct: pct(s365),
    bySite:      [...bySiteMap.values()].sort((a, b) => b.stale365 - a.stale365),
    byLibrary:   [...byLibMap.values()].sort((a, b) => b.stale365 - a.stale365),
    byExtension: [...byExtMap.values()].sort((a, b) => b.files - a.files)
  };
}

export function mergeSiteExposureIntoSummary(summary, siteExposureBySite) {
  const exposureMap = new Map();

  for (const site of siteExposureBySite || []) {
    const key = `${site.siteName}||${site.siteUrl}`;
    exposureMap.set(key, site);
  }

  summary.bySite = summary.bySite.map((site) => {
    const exposure = exposureMap.get(`${site.siteName}||${site.siteUrl}`);
    return {
      ...site,
      siteScopesChecked: exposure?.siteScopesChecked ?? 0,
      libraryScopesChecked: exposure?.libraryScopesChecked ?? 0,
      folderScopesChecked: exposure?.folderScopesChecked ?? 0,
      anyoneLinks: exposure?.anyoneLinks ?? 0,
      organizationLinks: exposure?.organizationLinks ?? 0,
      everyoneAllUsersGrants: exposure?.everyoneAllUsersGrants ?? 0,
      resourcesWithAnyoneLinks: exposure?.resourcesWithAnyoneLinks ?? 0,
      resourcesWithEveryoneAllUsers: exposure?.resourcesWithEveryoneAllUsers ?? 0,
      permissionsChecked: exposure?.permissionsChecked ?? 0
    };
  });

  return summary;
}

// ── CSV / JSON helpers ────────────────────────────────────────
function escCsv(v) {
  const s = String(v ?? "");
  return s.includes(",") || s.includes("\n") || s.includes('"')
    ? `"${s.replaceAll('"', '""')}"` : s;
}

function toCsv(rows) {
  if (!rows.length) return "";
  const headers = Object.keys(rows[0]);
  const lines   = rows.map(r => headers.map(h => escCsv(r[h])).join(","));
  return [headers.join(","), ...lines].join("\n");
}

function download(filename, content, type) {
  const url = URL.createObjectURL(new Blob([content], { type }));
  const a   = document.createElement("a");
  a.href     = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}

export function exportSummaryBySiteCsv(summary) {
  const rows = summary.bySite.map(s => ({
    SiteName: s.siteName, SiteUrl: s.siteUrl,
    Files: s.files, TotalSize: s.totalSize,
    Stale90: s.stale90, Stale180: s.stale180, Stale365: s.stale365,
    SiteScopesChecked: s.siteScopesChecked ?? 0,
    LibraryScopesChecked: s.libraryScopesChecked ?? 0,
    FolderScopesChecked: s.folderScopesChecked ?? 0,
    AnyoneLinks: s.anyoneLinks ?? 0,
    OrganizationLinks: s.organizationLinks ?? 0,
    EveryoneOrAllUsersGrants: s.everyoneAllUsersGrants ?? 0,
    ResourcesWithAnyoneLinks: s.resourcesWithAnyoneLinks ?? 0,
    ResourcesWithEveryoneOrAllUsers: s.resourcesWithEveryoneAllUsers ?? 0,
    PermissionsChecked: s.permissionsChecked ?? 0
  }));
  download("summary_by_site.csv", toCsv(rows), "text/csv;charset=utf-8");
}

export function exportSiteExposureCsv(siteExposureBySite) {
  const rows = (siteExposureBySite || []).map((s) => ({
    SiteName: s.siteName,
    SiteUrl: s.siteUrl,
    SiteScopesChecked: s.siteScopesChecked ?? 0,
    LibraryScopesChecked: s.libraryScopesChecked ?? 0,
    FolderScopesChecked: s.folderScopesChecked ?? 0,
    AnyoneLinks: s.anyoneLinks ?? 0,
    OrganizationLinks: s.organizationLinks ?? 0,
    EveryoneOrAllUsersGrants: s.everyoneAllUsersGrants ?? 0,
    ResourcesWithAnyoneLinks: s.resourcesWithAnyoneLinks ?? 0,
    ResourcesWithEveryoneOrAllUsers: s.resourcesWithEveryoneAllUsers ?? 0,
    PermissionsChecked: s.permissionsChecked ?? 0
  }));
  download("summary_site_exposure.csv", toCsv(rows), "text/csv;charset=utf-8");
}

export function exportFullJson(result) {
  download("full_results.json", JSON.stringify(result, null, 2), "application/json;charset=utf-8");
}

export function exportRowsCsv(filename, rows) {
  download(filename, toCsv(rows || []), "text/csv;charset=utf-8");
}
