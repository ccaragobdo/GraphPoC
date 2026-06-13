import { ScanProgress } from "../scan/types";

interface ProgressPanelProps {
  progress: ScanProgress;
  maxRuntimeMinutes: number;
}

function formatSeconds(totalSeconds: number): string {
  const value = Math.max(0, Math.floor(totalSeconds));
  const minutes = Math.floor(value / 60)
    .toString()
    .padStart(2, "0");
  const seconds = Math.floor(value % 60)
    .toString()
    .padStart(2, "0");
  return `${minutes}:${seconds}`;
}

export function ProgressPanel({ progress, maxRuntimeMinutes }: ProgressPanelProps) {
  const total = maxRuntimeMinutes * 60;
  const pct = Math.min(100, (progress.elapsedSeconds / total) * 100);

  return (
    <section className="card panel">
      <h2>Progress</h2>
      <div className="stats-grid">
        <p><span>Phase</span><strong>{progress.phase}</strong></p>
        <p><span>Elapsed</span><strong>{formatSeconds(progress.elapsedSeconds)}</strong></p>
        <p><span>Remaining</span><strong>~{formatSeconds(progress.remainingSeconds)}</strong></p>
        <p><span>Sites Scanned</span><strong>{progress.sitesProcessed}</strong></p>
        <p><span>Files Scanned</span><strong>{progress.filesScanned}</strong></p>
        <p><span>Files / Sec</span><strong>{progress.filesPerSecond.toFixed(2)}</strong></p>
        <p><span>Current Site</span><strong>{progress.currentSite || "-"}</strong></p>
        <p><span>Current Library</span><strong>{progress.currentLibrary || "-"}</strong></p>
      </div>

      <div className="meter" role="progressbar" aria-valuemin={0} aria-valuemax={100} aria-valuenow={pct}>
        <div className="meter-fill" style={{ width: `${pct}%` }} />
      </div>

      {progress.timeLimitReached && (
        <p className="warning">Time limit reached — returning partial results.</p>
      )}
    </section>
  );
}
