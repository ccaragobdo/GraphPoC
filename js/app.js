// ── app.js ───────────────────────────────────────────────────
// React app loaded via ESM CDN — no build step required.
// Uses htm for JSX-like syntax without a compiler.

import { createElement, Component, useState, useCallback, useEffect } from "react";
import { createRoot }                           from "react-dom/client";
import htm                                      from "htm";
import { runScan }                              from "./scanner.js?v=20260615a";
import {
  summarize,
  mergeSiteExposureIntoSummary,
  exportSummaryBySiteCsv,
  exportSiteExposureCsv,
  exportFullJson
} from "./analysis.js";

// Bind htm to React.createElement — gives us html`` template literals
const html = htm.bind(createElement);

// ── Defaults ──────────────────────────────────────────────────
const DEFAULT_CONFIG = {
  maxRuntimeMinutes:    5,
  skipTimeLimit:        false,
  maxSites:             15,
  unlimitedSites:       false,
  maxFilesPerSite:      400,
  maxTotalFiles:        5000,
  unlimitedFiles:       false,
  maxFolderDepth:       2,
  includeOneDrive:      true
};

const IDLE_PROGRESS = {
  phase: "Idle", elapsedSeconds: 0, remainingSeconds: 300,
  sitesProcessed: 0, filesScanned: 0, filesPerSecond: 0,
  currentSite: "", currentLibrary: "", timeLimitReached: false
};

// ── Helpers ───────────────────────────────────────────────────
function fmtTime(seconds) {
  const v = Math.max(0, Math.floor(seconds));
  return `${String(Math.floor(v / 60)).padStart(2, "0")}:${String(v % 60).padStart(2, "0")}`;
}

// ── Components ────────────────────────────────────────────────

function Header() {
  return html`
    ${html`<img
      src="https://www.bdo.com/getmedia/4e64ef6e-396a-4245-b942-b3c74eb04f8f/BDO-USA_web-01.svg"
      alt="BDO"
      className="bdo-logo"
    />`}
    <header className="card hero">
      <p className="kicker">Microsoft 365 FAST Assessment</p>
      <h1>Content Staleness Sampling Scanner</h1>
      <p className="caption">Scans only content accessible to the delegated token.</p>
      <p className="caption strong">This is a time-based sampling assessment, not a full inventory.</p>
    </header>`;
}

function TokenPanel({ token, onChange, disabled }) {
  const ok = token.trim().length > 20;
  return html`
    <section className="card panel">
      <h2>Step 1 — Paste your Graph Access Token</h2>
      <p className="caption">
        Get a token from${" "}
        <a href="https://developer.microsoft.com/en-us/graph/graph-explorer" target="_blank" rel="noreferrer">
          Graph Explorer
        </a>${" "}
        (sign in → click the token icon → copy access token), or from PowerShell / Azure CLI.
      </p>
      <p className="caption">
        Required scopes: <code>User.Read</code> <code>Files.Read.All</code> <code>Sites.Read.All</code>
      </p>
      <textarea
        className="token-input"
        rows=${6}
        value=${token}
        disabled=${disabled}
        onChange=${e => onChange(e.target.value)}
        placeholder="Paste access token here — do NOT include the word 'Bearer'"
      />
      <p className=${ok ? "caption token-status-ok" : "caption token-status-bad"}>
        ${ok ? "✓ Token looks good" : "⚠ Paste a token above to enable the scan"}
      </p>
    </section>`;
}

function ConfigPanel({ config, onChange, disabled, onStart, hasToken }) {
  const set = (key, val) => onChange({ ...config, [key]: val });

  return html`
    <section className="card panel">
      <h2>Step 2 — Configure &amp; Start Scan</h2>
      <div className="config-grid">
        <label>
          <input type="checkbox" checked=${config.unlimitedSites} disabled=${disabled}
            onChange=${e => set("unlimitedSites", e.target.checked)} />
          Unlimited Sites
        </label>
        ${!config.unlimitedSites ? html`
          <label>
            Max Sites
            <input type="number" min="1" max="100" value=${config.maxSites} disabled=${disabled}
              onChange=${e => set("maxSites", Number(e.target.value))} />
          </label>
        ` : null}
        
        <label>
          <input type="checkbox" checked=${config.unlimitedFiles} disabled=${disabled}
            onChange=${e => set("unlimitedFiles", e.target.checked)} />
          Unlimited Files
        </label>
        ${!config.unlimitedFiles ? html`
          <label>
            Max Total Files
            <input type="number" min="100" max="100000" value=${config.maxTotalFiles} disabled=${disabled}
              onChange=${e => set("maxTotalFiles", Number(e.target.value))} />
          </label>
        ` : null}
        
        ${!config.unlimitedFiles ? html`
          <label>
            Max Files Per Site
            <input type="number" min="50" max="10000" value=${config.maxFilesPerSite} disabled=${disabled}
              onChange=${e => set("maxFilesPerSite", Number(e.target.value))} />
          </label>
        ` : null}

        <label>
          <input type="checkbox" checked=${config.skipTimeLimit} disabled=${disabled}
            onChange=${e => set("skipTimeLimit", e.target.checked)} />
          Skip Time Limit (no timeout)
        </label>
        
        ${!config.skipTimeLimit ? html`
          <label>
            Max Runtime (minutes)
            <input type="number" min="1" max="120" value=${config.maxRuntimeMinutes} disabled=${disabled}
              onChange=${e => set("maxRuntimeMinutes", Number(e.target.value))} />
          </label>
        ` : null}
        
        <label>
          Max Folder Depth
          <input type="number" min="1" max="5" value=${config.maxFolderDepth} disabled=${disabled}
            onChange=${e => set("maxFolderDepth", Number(e.target.value))} />
        </label>

        <label className="inline-check">
          <input type="checkbox" checked=${config.includeOneDrive} disabled=${disabled}
            onChange=${e => set("includeOneDrive", e.target.checked)} />
          Include OneDrive
        </label>
        <button className="btn" type="button" onClick=${onStart} disabled=${disabled || !hasToken}>
          ${disabled ? "Scan Running…" : hasToken ? "▶ Start Scan" : "Paste a token first"}
        </button>
      </div>
    </section>`;
}

function ProgressPanel({ progress, maxRuntimeMinutes, visible }) {
  if (!visible) return null;
  const pct = Math.min(100, (progress.elapsedSeconds / (maxRuntimeMinutes * 60)) * 100);
  return html`
    <section className="card panel">
      <h2>Scan Progress</h2>
      <div className="stats-grid">
        <p><span>Phase</span><strong>${progress.phase}</strong></p>
        <p><span>Elapsed</span><strong>${fmtTime(progress.elapsedSeconds)}</strong></p>
        <p><span>~Remaining</span><strong>${fmtTime(progress.remainingSeconds)}</strong></p>
        <p><span>Sites Processed</span><strong>${progress.sitesProcessed}</strong></p>
        <p><span>Files Scanned</span><strong>${progress.filesScanned.toLocaleString()}</strong></p>
        <p><span>Files / Sec</span><strong>${progress.filesPerSecond.toFixed(2)}</strong></p>
        <p><span>Current Site</span><strong>${progress.currentSite || "—"}</strong></p>
        <p><span>Current Library</span><strong>${progress.currentLibrary || "—"}</strong></p>
      </div>
      <div className="meter" role="progressbar" aria-valuenow=${Math.round(pct)} aria-valuemin="0" aria-valuemax="100">
        <div className="meter-fill" style=${{ width: pct + "%" }} />
      </div>
      ${progress.timeLimitReached
        ? html`<p className="warning">⏱ Time limit reached — returning partial results.</p>`
        : null}
    </section>`;
}

function ResultsPanel({ result }) {
  if (!result) return null;
  const { summary, notes, siteExposureBySite } = result;

  return html`
    <section className="card panel">
      <h2>Results</h2>
      <div className="stats-grid">
        <p><span>Total Files Scanned</span><strong>${summary.totalFiles.toLocaleString()}</strong></p>
        <p><span>Total Size (bytes)</span><strong>${summary.totalSize.toLocaleString()}</strong></p>
        <p><span>Stale &gt; 90 days</span><strong>${summary.stale90Pct}%</strong></p>
        <p><span>Stale &gt; 180 days</span><strong>${summary.stale180Pct}%</strong></p>
        <p><span>Stale &gt; 365 days</span><strong>${summary.stale365Pct}%</strong></p>
        <p><span>Sites Processed</span><strong>${result.sitesProcessed}</strong></p>
      </div>

      <h3>Top Stale Sites (by files &gt;365 days)</h3>
      <ul className="rank-list">
        ${summary.bySite.slice(0, 5).map(s => html`
          <li key=${s.siteName + s.siteUrl}>
            <strong>${s.siteName}</strong>
            <span>
              ${s.files} files | stale &gt;365: ${s.stale365} | anyone links: ${(s.anyoneLinks ?? 0)}
              | everyone/all users grants: ${(s.everyoneAllUsersGrants ?? 0)}
              | scopes checked: ${(s.permissionsChecked ?? 0)}
            </span>
          </li>`)}
      </ul>

      <h3>Top Stale Libraries</h3>
      <ul className="rank-list">
        ${summary.byLibrary.slice(0, 5).map(l => html`
          <li key=${l.siteName + l.libraryName}>
            <strong>${l.siteName} / ${l.libraryName}</strong>
            <span>${l.files} files | stale &gt;365: ${l.stale365}</span>
          </li>`)}
      </ul>

      ${notes.length ? html`
        <h3>Scan Notes</h3>
        <ul className="notes-list">
          ${notes.map((n, i) => html`<li key=${i}>${n}</li>`)}
        </ul>` : null}

      <p className="caption strong">
        ⚠ This is a time-based sampling assessment, not a full inventory.
      </p>

      <div className="button-row">
        <button className="btn ghost" onClick=${() => exportSummaryBySiteCsv(summary)}>⬇ summary_by_site.csv</button>
        <button className="btn ghost" onClick=${() => exportSiteExposureCsv(siteExposureBySite)}>⬇ summary_site_exposure.csv</button>
        <button className="btn ghost" onClick=${() => exportFullJson(result)}>⬇ full_results.json</button>
      </div>
    </section>`;
}

// ── Root App ──────────────────────────────────────────────────
function App() {
  const [token,      setToken]      = useState("");
  const [config,     setConfig]     = useState(DEFAULT_CONFIG);
  const [isScanning, setIsScanning] = useState(false);
  const [progress,   setProgress]   = useState(IDLE_PROGRESS);
  const [result,     setResult]     = useState(null);
  const [error,      setError]      = useState("");
  const [fatalError, setFatalError] = useState("");

  const hasToken = token.trim().length > 20;

  const startScan = useCallback(async () => {
    const t = token.trim();
    if (!t) { setError("Paste a Graph bearer token first."); return; }

    setError("");
    setResult(null);
    setIsScanning(true);
    setProgress({ ...IDLE_PROGRESS, phase: "Discovery", remainingSeconds: config.maxRuntimeMinutes * 60 });

    try {
      const scanData = await runScan(t, config, p => setProgress({ ...p }));
      const summary  = mergeSiteExposureIntoSummary(
        summarize(scanData.records),
        scanData.siteExposureBySite || []
      );
      setResult({ ...scanData, summary });
      setProgress(prev => ({ ...prev, phase: "Complete" }));
    } catch (e) {
      setError(`Scan error: ${e.message}`);
    } finally {
      setIsScanning(false);
    }
  }, [token, config]);

  useEffect(() => {
    function onWindowError(event) {
      const message = event?.error?.message || event?.message || "Unexpected runtime error";
      setFatalError(String(message));
    }

    function onUnhandledRejection(event) {
      const reason = event?.reason;
      const message = reason?.message || String(reason || "Unexpected promise rejection");
      setFatalError(message);
    }

    window.addEventListener("error", onWindowError);
    window.addEventListener("unhandledrejection", onUnhandledRejection);

    return () => {
      window.removeEventListener("error", onWindowError);
      window.removeEventListener("unhandledrejection", onUnhandledRejection);
    };
  }, []);

  useEffect(() => {
    function guardNavigation(event) {
      if (!isScanning) return;
      event.preventDefault();
      event.returnValue = "";
    }

    window.addEventListener("beforeunload", guardNavigation);
    return () => {
      window.removeEventListener("beforeunload", guardNavigation);
    };
  }, [isScanning]);

  const showProgress = isScanning || (progress.phase !== "Idle" && progress.phase !== "Complete") || progress.phase === "Complete";

  return html`
    <main className="app-shell">
      <${Header} />
      <${TokenPanel} token=${token} onChange=${setToken} disabled=${isScanning} />
      <${ConfigPanel}
        config=${config}
        onChange=${setConfig}
        disabled=${isScanning}
        onStart=${startScan}
        hasToken=${hasToken}
      />
      <${ProgressPanel}
        progress=${progress}
        maxRuntimeMinutes=${config.maxRuntimeMinutes}
        visible=${showProgress}
      />
      <${ResultsPanel} result=${result} />
      ${error ? html`<section className="card panel error">⚠ ${error}</section>` : null}
      ${fatalError ? html`<section className="card panel error">⚠ Runtime error: ${fatalError}</section>` : null}
    </main>`;
}

class AppErrorBoundary extends Component {
  constructor(props) {
    super(props);
    this.state = { hasError: false };
  }

  static getDerivedStateFromError() {
    return { hasError: true };
  }

  componentDidCatch(error) {
    // eslint-disable-next-line no-console
    console.error("Unhandled UI error", error);
  }

  render() {
    if (this.state.hasError) {
      return html`
        <main className="app-shell">
          <section className="card panel error">
            The app hit an unexpected error. Refresh and try again.
          </section>
        </main>
      `;
    }

    return this.props.children;
  }
}

// ── Mount ─────────────────────────────────────────────────────
createRoot(document.getElementById("root")).render(
  html`<${AppErrorBoundary}><${App} /><//>`
);
