import {
  exportFullJson,
  exportInventoryCsv,
  exportSummaryByLibraryCsv,
  exportSummaryBySiteCsv
} from "../analysis/export";
import { ScanResult } from "../scan/types";

interface ResultsPanelProps {
  result: ScanResult | null;
}

function pct(value: number): string {
  return `${value.toFixed(1)}%`;
}

export function ResultsPanel({ result }: ResultsPanelProps) {
  if (!result) {
    return (
      <section className="card panel">
        <h2>Results</h2>
        <p>No scan results yet.</p>
      </section>
    );
  }

  const topSites = result.summary.bySite.slice(0, 5);
  const topLibraries = result.summary.byLibrary.slice(0, 5);

  return (
    <section className="card panel">
      <h2>Results</h2>
      <div className="stats-grid">
        <p><span>Total Files</span><strong>{result.summary.totalFiles}</strong></p>
        <p><span>Total Size (Bytes)</span><strong>{result.summary.totalSize}</strong></p>
        <p><span>Stale &gt; 90 Days</span><strong>{pct(result.summary.stale90Pct)}</strong></p>
        <p><span>Stale &gt; 180 Days</span><strong>{pct(result.summary.stale180Pct)}</strong></p>
        <p><span>Stale &gt; 365 Days</span><strong>{pct(result.summary.stale365Pct)}</strong></p>
      </div>

      <h3>Top Stale Sites</h3>
      <ul className="rank-list">
        {topSites.map((site) => (
          <li key={`${site.siteName}-${site.siteUrl}`}>
            <strong>{site.siteName}</strong>
            <span>{site.files} files | stale&gt;365: {site.stale365}</span>
          </li>
        ))}
      </ul>

      <h3>Top Stale Libraries</h3>
      <ul className="rank-list">
        {topLibraries.map((library) => (
          <li key={`${library.siteName}-${library.libraryName}`}>
            <strong>{library.siteName} / {library.libraryName}</strong>
            <span>{library.files} files | stale&gt;365: {library.stale365}</span>
          </li>
        ))}
      </ul>

      {result.notes.length > 0 && (
        <>
          <h3>Scan Notes</h3>
          <ul className="notes-list">
            {result.notes.map((note, idx) => (
              <li key={`${idx}-${note}`}>{note}</li>
            ))}
          </ul>
        </>
      )}

      <div className="button-row">
        <button className="btn" onClick={() => exportInventoryCsv(result.inventory)}>
          Export CSV (File Inventory)
        </button>
        <button className="btn ghost" onClick={() => exportSummaryBySiteCsv(result)}>
          Export CSV (Summary by Site)
        </button>
        <button className="btn ghost" onClick={() => exportSummaryByLibraryCsv(result)}>
          Export CSV (Summary by Library)
        </button>
        <button className="btn ghost" onClick={() => exportFullJson(result)}>
          Export JSON
        </button>
      </div>
    </section>
  );
}
