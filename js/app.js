// ── app.js ───────────────────────────────────────────────────
// React app loaded via ESM CDN — no build step required.
// Uses htm for JSX-like syntax without a compiler.

import { createElement, Component, useState, useCallback, useEffect } from "react";
import { createRoot }                           from "react-dom/client";
import htm                                      from "htm";
import { runScan, runPermissionAnalysis }       from "./scanner.js?v=20260616b";
import {
  summarize,
  mergeSiteExposureIntoSummary,
  exportSummaryBySiteCsv,
  exportSiteExposureCsv,
  exportFullJson,
  exportRowsCsv
} from "./analysis.js";
import {
  runStep4Sensitivity,
  runStep5SIT,
  runStep6DLP,
  runStep7InformationArchitecture,
  runStep8SearchReadiness,
  runStep9Lifecycle
} from "./assessments.js?v=20260616b";
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
  includeOneDrive:      true,
  maxSampleFiles:       2000,
  lifecycleDays:        180
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
      <p className="kicker">BDO Data Governance Assessment</p>
      <h1>Data Architecture Scanner</h1>
      <p className="caption">Scans only content accessible to the delegated token.</p>
      <p className="caption strong">This is a time-based sampling assessment, not a full inventory.</p>
      <a href="about.html" target="_blank" rel="noreferrer" className="caption" style=${{ marginTop: ".4rem", display: "inline-block" }}>ℹ What does each step do? →</a>
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

        <label>
          Max Sample Files (Step 8)
          <input type="number" min="200" max="20000" value=${config.maxSampleFiles} disabled=${disabled}
            onChange=${e => set("maxSampleFiles", Number(e.target.value))} />
        </label>

        <label>
          Lifecycle Threshold Days (Step 9)
          <input type="number" min="30" max="3650" value=${config.lifecycleDays} disabled=${disabled}
            onChange=${e => set("lifecycleDays", Number(e.target.value))} />
        </label>

        <label className="inline-check">
          <input type="checkbox" checked=${config.includeOneDrive} disabled=${disabled}
            onChange=${e => set("includeOneDrive", e.target.checked)} />
          Include OneDrive
        </label>
        <button className="btn" type="button" onClick=${onStart} disabled=${disabled || !hasToken}>
          ${disabled ? "Scan Running…" : hasToken ? "▶ Start Scan" : "Paste a token first"}
        </button>
        <p className="caption">
          Steps 3-9 can run independently below using the same token.
        </p>
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
        <div className="meter-fill" style=${{ width: `${pct}%` }} />
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

function IndependentStepsPanel({ hasToken, runningStep, onRunStep, stepResults }) {
  const steps = [
    { id: 3, title: "Step 3 — Sharing & Permissions Exposure",          runnable: true  },
    { id: 4, title: "Step 4 — Sensitivity Labels & Purview Integration", runnable: true  },
    { id: 5, title: "Step 5 — Sensitive Information Types (SITs)",       runnable: false },
    { id: 6, title: "Step 6 — Data Loss Prevention (DLP)",               runnable: false },
    { id: 7, title: "Step 7 — SharePoint Information Architecture",      runnable: true  },
    { id: 8, title: "Step 8 — Search Readiness Signals",                 runnable: true  },
    { id: 9, title: "Step 9 — Lifecycle & Records Management",           runnable: true  }
  ];

  function metricEntries(metrics) {
    return Object.entries(metrics || {}).slice(0, 8);
  }

  return html`
    <section className="card panel">
      <h2>Independent Assessment Steps</h2>
      <p className="caption">Run any step independently with the same Graph token. Each step has its own output and export.</p>

      ${steps.map(step => {
        const stepResult = stepResults[step.id];
        const isRunning = runningStep === step.id;
        const canRun = hasToken && !runningStep;

        return html`
          <div className="card panel" key=${step.id}>
            <h3>${step.title}</h3>
            <div className="button-row">
              ${step.runnable ? html`
                <button className="btn" onClick=${() => onRunStep(step.id)} disabled=${!canRun}>
                  ${isRunning ? "Running..." : `▶ Run ${step.title.split("—")[0].trim()}`}
                </button>
              ` : html`<p className="caption">Requires Purview — not runnable via this Graph token path.</p>`}
              ${stepResult ? html`
                <button className="btn ghost" onClick=${() => onRunStep(step.id, "json")}>⬇ step_${step.id}_results.json</button>
                ${stepResult.rows?.length ? html`
                  <button className="btn ghost" onClick=${() => onRunStep(step.id, "csv")}>⬇ step_${step.id}_results.csv</button>
                ` : null}
              ` : null}
            </div>

            ${stepResult ? html`
              <p className="caption strong">${stepResult.feasible ? "Feasible via current Graph access" : "Not fully feasible via current Graph access"}</p>
              <p className="caption">${stepResult.summary || ""}</p>

              ${metricEntries(stepResult.metrics).length ? html`
                <div className="stats-grid">
                  ${metricEntries(stepResult.metrics).map(([k, v]) => html`
                    <p key=${k}><span>${k}</span><strong>${String(v)}</strong></p>
                  `)}
                </div>
              ` : null}

              ${stepResult.notes?.length ? html`
                <ul className="notes-list">
                  ${stepResult.notes.slice(0, 6).map((n, i) => html`<li key=${i}>${n}</li>`)}
                </ul>
              ` : null}
            ` : html`<p className="caption">No output yet.</p>`}
          </div>
        `;
      })}
    </section>
  `;
}

// ── Root App ──────────────────────────────────────────────────
function App() {
  const [token,      setToken]      = useState("");
  const [config,     setConfig]     = useState(DEFAULT_CONFIG);
  const [isScanning, setIsScanning] = useState(false);
  const [runningStep, setRunningStep] = useState(null);
  const [stepResults, setStepResults] = useState(() => {
    try {
      const saved = sessionStorage.getItem("dga_stepResults");
      return saved ? JSON.parse(saved) : {};
    } catch { return {}; }
  });
  const [progress,   setProgress]   = useState(IDLE_PROGRESS);
  const [result,     setResult]     = useState(null);
  const [error,      setError]      = useState("");
  const [fatalError, setFatalError] = useState("");

  const hasToken = token.trim().length > 20;

  const exportStepJson = useCallback((stepId) => {
    const data = stepResults[stepId];
    if (!data) return;
    exportRowsCsv(`step_${stepId}_results.csv`, data.rows || []);
  }, [stepResults]);

  const exportStepStructuredJson = useCallback((stepId) => {
    const data = stepResults[stepId];
    if (!data) return;
    const blob = new Blob([JSON.stringify(data, null, 2)], { type: "application/json;charset=utf-8" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `step_${stepId}_results.json`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  }, [stepResults]);

  const runStep = useCallback(async (stepId, action = "run") => {
    if (action === "json") {
      exportStepStructuredJson(stepId);
      return;
    }
    if (action === "csv") {
      exportStepJson(stepId);
      return;
    }

    const t = token.trim();
    if (!t) { setError("Paste a Graph bearer token first."); return; }

    setError("");
    setRunningStep(stepId);
    setProgress({ ...IDLE_PROGRESS, phase: `Running Step ${stepId}` });

    const runners = {
      3: () => runPermissionAnalysis(t, config, p => setProgress(prev => ({ ...prev, ...p }))),
      4: () => runStep4Sensitivity(t, config, p => setProgress(prev => ({ ...prev, ...p }))),
      5: () => runStep5SIT(t, config, p => setProgress(prev => ({ ...prev, ...p }))),
      6: () => runStep6DLP(t, config, p => setProgress(prev => ({ ...prev, ...p }))),
      7: () => runStep7InformationArchitecture(t, config, p => setProgress(prev => ({ ...prev, ...p }))),
      8: () => runStep8SearchReadiness(t, config, p => setProgress(prev => ({ ...prev, ...p }))),
      9: () => runStep9Lifecycle(t, config, p => setProgress(prev => ({ ...prev, ...p })))
    };

    try {
      const data = await runners[stepId]();

      if (stepId === 3 && result) {
        const mergedSummary = mergeSiteExposureIntoSummary(
          summarize(result.records || []),
          data.siteExposureBySite || []
        );
        setResult(prev => prev ? {
          ...prev,
          siteExposureBySite: data.siteExposureBySite || [],
          summary: mergedSummary,
          notes: [...(prev.notes || []), ...(data.notes || [])]
        } : prev);
      }

      setStepResults(prev => {
        const next = { ...prev, [stepId]: data };
        try { sessionStorage.setItem("dga_stepResults", JSON.stringify(next)); } catch {}
        return next;
      });
      setProgress(prev => ({ ...prev, phase: `Step ${stepId} Complete` }));
    } catch (e) {
      setError(`Step ${stepId} failed: ${e.message}`);
    } finally {
      setRunningStep(null);
    }
  }, [token, config, result, exportStepJson, exportStepStructuredJson]);

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

  const showProgress = isScanning || !!runningStep || (progress.phase !== "Idle" && progress.phase !== "Complete") || progress.phase === "Complete";

  return html`
    <main className="app-shell">
      <${Header} />
      <${TokenPanel} token=${token} onChange=${setToken} disabled=${isScanning} />
      <${ConfigPanel}
        config=${config}
        onChange=${setConfig}
        disabled=${isScanning || !!runningStep}
        onStart=${startScan}
        hasToken=${hasToken}
      />
      <${IndependentStepsPanel}
        hasToken=${hasToken}
        runningStep=${runningStep}
        onRunStep=${runStep}
        stepResults=${stepResults}
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
  html`<${AppErrorBoundary}><${App} /></${AppErrorBoundary}>`
);
