import { useEffect, useMemo, useState } from "react";
import { AccountInfo, PublicClientApplication } from "@azure/msal-browser";
import { getAccessToken, signIn, signOut } from "./auth/authService";
import { msalConfig } from "./auth/msalConfig";
import { ConfigPanel } from "./components/ConfigPanel";
import { HeaderBar } from "./components/HeaderBar";
import { ProgressPanel } from "./components/ProgressPanel";
import { ResultsPanel } from "./components/ResultsPanel";
import { TokenPanel } from "./components/TokenPanel";
import { useScan } from "./hooks/useScan";
import { ScanConfig } from "./scan/types";

const defaultConfig: ScanConfig = {
  maxRuntimeMinutes: 5,
  maxSites: 15,
  maxFilesPerSite: 400,
  maxTotalFiles: 5000,
  maxFolderDepth: 2,
  includeOneDrive: true,
  searchTerms: ["team", "project", "shared"],
  siteConcurrency: 2
};

export default function App() {
  const isMsalEnabled = Boolean((import.meta.env.VITE_MSAL_CLIENT_ID || "").trim());
  const [instance] = useState<PublicClientApplication | null>(() =>
    isMsalEnabled ? new PublicClientApplication(msalConfig) : null
  );
  const [isSigningIn, setIsSigningIn] = useState(false);
  const [authError, setAuthError] = useState("");
  const [account, setAccount] = useState<AccountInfo | null>(null);
  const [manualToken, setManualToken] = useState("");
  const [config, setConfig] = useState<ScanConfig>(defaultConfig);
  const scan = useScan();

  useEffect(() => {
    if (!instance) {
      setAccount(null);
      return;
    }

    const active = instance.getActiveAccount();
    if (active) {
      setAccount(active);
      return;
    }

    const all = instance.getAllAccounts();
    if (all.length > 0) {
      instance.setActiveAccount(all[0]);
      setAccount(all[0]);
    }
  }, [instance]);

  async function handleSignIn(): Promise<void> {
    if (!instance) {
      setAuthError("MSAL is not configured. Use manual token mode.");
      return;
    }

    setAuthError("");
    setIsSigningIn(true);
    try {
      const signedIn = await signIn(instance);
      instance.setActiveAccount(signedIn);
      setAccount(signedIn);
    } catch (error) {
      setAuthError((error as Error).message);
    } finally {
      setIsSigningIn(false);
    }
  }

  async function handleSignOut(): Promise<void> {
    if (!instance || !account) {
      return;
    }

    await signOut(instance, account);
    setAccount(null);
  }

  async function handleStartScan(): Promise<void> {
    setAuthError("");

    const manualTokenValue = manualToken.trim();
    if (!isMsalEnabled && !manualTokenValue) {
      setAuthError("Paste a Graph bearer token before starting the scan.");
      return;
    }

    try {
      let accessToken = manualTokenValue;

      if (!accessToken) {
        if (!instance || !account) {
          setAuthError("Sign in before starting the scan.");
          return;
        }

        accessToken = await getAccessToken(instance, account);
      }

      await scan.start(accessToken, config);
    } catch (error) {
      setAuthError(`Authentication error: ${(error as Error).message}`);
    }
  }

  const combinedError = useMemo(() => scan.error || authError, [scan.error, authError]);
  const hasManualToken = manualToken.trim().length > 0;
  const canRunScan = isMsalEnabled ? Boolean(account) : hasManualToken;

  return (
    <main className="app-shell">
      <HeaderBar
        account={account}
        isMsalEnabled={isMsalEnabled}
        hasManualToken={hasManualToken}
        isSigningIn={isSigningIn}
        onSignIn={handleSignIn}
        onSignOut={handleSignOut}
      />

      {!isMsalEnabled && (
        <TokenPanel token={manualToken} onTokenChange={setManualToken} disabled={scan.isRunning} />
      )}

      <ConfigPanel
        config={config}
        onConfigChange={setConfig}
        disabled={!canRunScan || scan.isRunning}
        onStart={handleStartScan}
      />

      <ProgressPanel progress={scan.progress} maxRuntimeMinutes={config.maxRuntimeMinutes} />

      <ResultsPanel result={scan.result} />

      {combinedError && <section className="card panel error">{combinedError}</section>}
    </main>
  );
}
