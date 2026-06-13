import { FileRecord, LibrarySummary, ScanSummary, SiteSummary, ExtensionSummary } from "../scan/types";

const DAY_MS = 24 * 60 * 60 * 1000;

interface Staleness {
  stale90: number;
  stale180: number;
  stale365: number;
}

function getStaleness(lastModifiedDateTime: string): Staleness {
  if (!lastModifiedDateTime) {
    return { stale90: 0, stale180: 0, stale365: 0 };
  }

  const modified = new Date(lastModifiedDateTime).getTime();
  if (Number.isNaN(modified)) {
    return { stale90: 0, stale180: 0, stale365: 0 };
  }

  const ageDays = (Date.now() - modified) / DAY_MS;
  return {
    stale90: ageDays > 90 ? 1 : 0,
    stale180: ageDays > 180 ? 1 : 0,
    stale365: ageDays > 365 ? 1 : 0
  };
}

function pct(value: number, total: number): number {
  return total === 0 ? 0 : (value / total) * 100;
}

export function summarizeInventory(records: FileRecord[]): ScanSummary {
  const bySiteMap = new Map<string, SiteSummary>();
  const byLibraryMap = new Map<string, LibrarySummary>();
  const byExtensionMap = new Map<string, ExtensionSummary>();

  let totalSize = 0;
  let stale90 = 0;
  let stale180 = 0;
  let stale365 = 0;

  for (const record of records) {
    totalSize += record.size;
    const staleness = getStaleness(record.lastModifiedDateTime);
    stale90 += staleness.stale90;
    stale180 += staleness.stale180;
    stale365 += staleness.stale365;

    const siteKey = `${record.siteName}|${record.siteUrl}`;
    const currentSite = bySiteMap.get(siteKey) ?? {
      siteName: record.siteName,
      siteUrl: record.siteUrl,
      files: 0,
      totalSize: 0,
      stale90: 0,
      stale180: 0,
      stale365: 0
    };

    currentSite.files += 1;
    currentSite.totalSize += record.size;
    currentSite.stale90 += staleness.stale90;
    currentSite.stale180 += staleness.stale180;
    currentSite.stale365 += staleness.stale365;
    bySiteMap.set(siteKey, currentSite);

    const libKey = `${record.siteName}|${record.siteUrl}|${record.libraryName}`;
    const currentLibrary = byLibraryMap.get(libKey) ?? {
      siteName: record.siteName,
      siteUrl: record.siteUrl,
      libraryName: record.libraryName,
      files: 0,
      totalSize: 0,
      stale90: 0,
      stale180: 0,
      stale365: 0
    };

    currentLibrary.files += 1;
    currentLibrary.totalSize += record.size;
    currentLibrary.stale90 += staleness.stale90;
    currentLibrary.stale180 += staleness.stale180;
    currentLibrary.stale365 += staleness.stale365;
    byLibraryMap.set(libKey, currentLibrary);

    const ext = record.extension || "(none)";
    const currentExt = byExtensionMap.get(ext) ?? {
      extension: ext,
      files: 0,
      totalSize: 0
    };
    currentExt.files += 1;
    currentExt.totalSize += record.size;
    byExtensionMap.set(ext, currentExt);
  }

  return {
    totalFiles: records.length,
    totalSize,
    stale90Pct: pct(stale90, records.length),
    stale180Pct: pct(stale180, records.length),
    stale365Pct: pct(stale365, records.length),
    bySite: Array.from(bySiteMap.values()).sort((a, b) => b.stale365 - a.stale365),
    byLibrary: Array.from(byLibraryMap.values()).sort((a, b) => b.stale365 - a.stale365),
    byExtension: Array.from(byExtensionMap.values()).sort((a, b) => b.files - a.files)
  };
}
