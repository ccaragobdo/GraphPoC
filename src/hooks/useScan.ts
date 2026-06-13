import { useMemo, useState } from "react";
import { createGraphClient } from "../graph/client";
import { runSamplingScan } from "../scan/scanner";
import { ScanConfig, ScanProgress, ScanResult } from "../scan/types";

const defaultProgress: ScanProgress = {
  phase: "Idle",
  elapsedSeconds: 0,
  remainingSeconds: 300,
  sitesProcessed: 0,
  filesScanned: 0,
  filesPerSecond: 0,
  currentSite: "",
  currentLibrary: "",
  timeLimitReached: false
};

export function useScan() {
  const [isRunning, setIsRunning] = useState(false);
  const [progress, setProgress] = useState<ScanProgress>(defaultProgress);
  const [result, setResult] = useState<ScanResult | null>(null);
  const [error, setError] = useState<string>("");

  async function start(accessToken: string, config: ScanConfig): Promise<void> {
    setIsRunning(true);
    setError("");
    setResult(null);
    setProgress({ ...defaultProgress, remainingSeconds: config.maxRuntimeMinutes * 60, phase: "Discovery" });

    try {
      const client = createGraphClient(accessToken);
      const scanResult = await runSamplingScan(client, config, (nextProgress) => {
        setProgress(nextProgress);
      });
      setResult(scanResult);
    } catch (scanError) {
      setError((scanError as Error).message);
    } finally {
      setIsRunning(false);
    }
  }

  const state = useMemo(
    () => ({ isRunning, progress, result, error }),
    [isRunning, progress, result, error]
  );

  return {
    ...state,
    start
  };
}
