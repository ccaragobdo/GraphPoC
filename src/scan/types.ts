export type ScanPhase = "Idle" | "Discovery" | "Scanning" | "Aggregating" | "Complete";

export interface ScanConfig {
  maxRuntimeMinutes: number;
  maxSites: number;
  maxFilesPerSite: number;
  maxTotalFiles: number;
  maxFolderDepth: number;
  includeOneDrive: boolean;
  searchTerms: string[];
  siteConcurrency: number;
}

export interface FileRecord {
  siteName: string;
  siteUrl: string;
  libraryName: string;
  fileName: string;
  extension: string;
  size: number;
  createdDateTime: string;
  lastModifiedDateTime: string;
}

export interface ScanProgress {
  phase: ScanPhase;
  elapsedSeconds: number;
  remainingSeconds: number;
  sitesProcessed: number;
  filesScanned: number;
  filesPerSecond: number;
  currentSite?: string;
  currentLibrary?: string;
  timeLimitReached: boolean;
}

export interface SiteSummary {
  siteName: string;
  siteUrl: string;
  files: number;
  totalSize: number;
  stale90: number;
  stale180: number;
  stale365: number;
}

export interface LibrarySummary {
  siteName: string;
  siteUrl: string;
  libraryName: string;
  files: number;
  totalSize: number;
  stale90: number;
  stale180: number;
  stale365: number;
}

export interface ExtensionSummary {
  extension: string;
  files: number;
  totalSize: number;
}

export interface ScanSummary {
  totalFiles: number;
  totalSize: number;
  stale90Pct: number;
  stale180Pct: number;
  stale365Pct: number;
  bySite: SiteSummary[];
  byLibrary: LibrarySummary[];
  byExtension: ExtensionSummary[];
}

export interface ScanResult {
  inventory: FileRecord[];
  summary: ScanSummary;
  progress: ScanProgress;
  startedAt: string;
  finishedAt: string;
  notes: string[];
}
