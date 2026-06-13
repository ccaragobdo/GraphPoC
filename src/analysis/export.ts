import { FileRecord, ScanResult } from "../scan/types";

function escapeCsv(value: string | number): string {
  const stringValue = String(value ?? "");
  if (stringValue.includes(",") || stringValue.includes("\n") || stringValue.includes("\"")) {
    return `"${stringValue.replaceAll("\"", "\"\"")}"`;
  }

  return stringValue;
}

function recordsToCsv(records: Record<string, string | number>[]): string {
  if (records.length === 0) {
    return "";
  }

  const headers = Object.keys(records[0]);
  const rows = records.map((record) => headers.map((h) => escapeCsv(record[h] ?? "")).join(","));
  return [headers.join(","), ...rows].join("\n");
}

function downloadBlob(fileName: string, content: string, contentType: string): void {
  const blob = new Blob([content], { type: contentType });
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement("a");
  anchor.href = url;
  anchor.download = fileName;
  anchor.click();
  URL.revokeObjectURL(url);
}

export function exportInventoryCsv(inventory: FileRecord[]): void {
  const rows = inventory.map((item) => ({
    SiteName: item.siteName,
    SiteUrl: item.siteUrl,
    LibraryName: item.libraryName,
    FileName: item.fileName,
    FileExtension: item.extension,
    FileSize: item.size,
    CreatedDateTime: item.createdDateTime,
    LastModifiedDateTime: item.lastModifiedDateTime
  }));

  downloadBlob("file_inventory.csv", recordsToCsv(rows), "text/csv;charset=utf-8");
}

export function exportSummaryBySiteCsv(result: ScanResult): void {
  const siteRows = result.summary.bySite.map((site) => ({
    SiteName: site.siteName,
    SiteUrl: site.siteUrl,
    Files: site.files,
    TotalSize: site.totalSize,
    Stale90: site.stale90,
    Stale180: site.stale180,
    Stale365: site.stale365
  }));

  downloadBlob("summary_by_site.csv", recordsToCsv(siteRows), "text/csv;charset=utf-8");
}

export function exportSummaryByLibraryCsv(result: ScanResult): void {
  const libraryRows = result.summary.byLibrary.map((library) => ({
    SiteName: library.siteName,
    SiteUrl: library.siteUrl,
    LibraryName: library.libraryName,
    Files: library.files,
    TotalSize: library.totalSize,
    Stale90: library.stale90,
    Stale180: library.stale180,
    Stale365: library.stale365
  }));

  downloadBlob("summary_by_library.csv", recordsToCsv(libraryRows), "text/csv;charset=utf-8");
}

export function exportFullJson(result: ScanResult): void {
  downloadBlob("full_results.json", JSON.stringify(result, null, 2), "application/json;charset=utf-8");
}
