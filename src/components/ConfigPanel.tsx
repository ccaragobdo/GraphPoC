import { FormEvent } from "react";
import { ScanConfig } from "../scan/types";

interface ConfigPanelProps {
  config: ScanConfig;
  disabled: boolean;
  onConfigChange: (next: ScanConfig) => void;
  onStart: () => void;
}

export function ConfigPanel({ config, disabled, onConfigChange, onStart }: ConfigPanelProps) {
  function update<K extends keyof ScanConfig>(key: K, value: ScanConfig[K]): void {
    onConfigChange({ ...config, [key]: value });
  }

  function submit(event: FormEvent): void {
    event.preventDefault();
    onStart();
  }

  return (
    <section className="card panel">
      <h2>Scan Configuration</h2>
      <form onSubmit={submit} className="config-grid">
        <label>
          Max Sites
          <input
            type="number"
            min={1}
            max={30}
            value={config.maxSites}
            disabled={disabled}
            onChange={(e) => update("maxSites", Number(e.target.value))}
          />
        </label>

        <label>
          Max Files Total
          <input
            type="number"
            min={100}
            max={10000}
            value={config.maxTotalFiles}
            disabled={disabled}
            onChange={(e) => update("maxTotalFiles", Number(e.target.value))}
          />
        </label>

        <label>
          Max Files Per Site
          <input
            type="number"
            min={50}
            max={1000}
            value={config.maxFilesPerSite}
            disabled={disabled}
            onChange={(e) => update("maxFilesPerSite", Number(e.target.value))}
          />
        </label>

        <label>
          Max Runtime (Minutes)
          <input
            type="number"
            min={1}
            max={10}
            value={config.maxRuntimeMinutes}
            disabled={disabled}
            onChange={(e) => update("maxRuntimeMinutes", Number(e.target.value))}
          />
        </label>

        <label>
          Search Terms (comma-separated)
          <input
            type="text"
            value={config.searchTerms.join(",")}
            disabled={disabled}
            onChange={(e) =>
              update(
                "searchTerms",
                e.target.value
                  .split(",")
                  .map((x) => x.trim())
                  .filter((x) => x.length > 0)
              )
            }
          />
        </label>

        <label className="inline-check">
          <input
            type="checkbox"
            checked={config.includeOneDrive}
            disabled={disabled}
            onChange={(e) => update("includeOneDrive", e.target.checked)}
          />
          Include OneDrive
        </label>

        <button className="btn" type="submit" disabled={disabled}>
          {disabled ? "Scan Running..." : "Start Scan"}
        </button>
      </form>
    </section>
  );
}
