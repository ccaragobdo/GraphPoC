import { AccountInfo } from "@azure/msal-browser";

interface HeaderBarProps {
  account: AccountInfo | null;
  isMsalEnabled: boolean;
  isMsalReady: boolean;
  hasManualToken: boolean;
  isSigningIn: boolean;
  onSignIn: () => void;
  onSignOut: () => void;
}

export function HeaderBar({
  account,
  isMsalEnabled,
  isMsalReady,
  hasManualToken,
  isSigningIn,
  onSignIn,
  onSignOut
}: HeaderBarProps) {
  return (
    <header className="card hero">
      <div>
        <p className="kicker">Microsoft 365 FAST Assessment</p>
        <h1>Content Staleness Sampling Scanner</h1>
        <p className="caption">
          This scan only includes content accessible to the signed-in user.
        </p>
        <p className="caption strong">
          This is a time-based sampling assessment, not a full inventory.
        </p>
      </div>

      <div className="auth-box">
        {isMsalEnabled ? (
          <>
            <p className="label">Current User</p>
            <p className="value">{isMsalReady ? account?.username ?? "Not signed in" : "Initializing auth..."}</p>
            {account ? (
              <button className="btn ghost" onClick={onSignOut}>
                Sign Out
              </button>
            ) : (
              <button className="btn" onClick={onSignIn} disabled={isSigningIn || !isMsalReady}>
                {!isMsalReady ? "Initializing..." : isSigningIn ? "Signing In..." : "Sign In"}
              </button>
            )}
          </>
        ) : (
          <>
            <p className="label">Auth Mode</p>
            <p className="value">Manual Graph Token</p>
            <p className="caption">Token loaded: {hasManualToken ? "Yes" : "No"}</p>
          </>
        )}
      </div>
    </header>
  );
}
