interface TokenPanelProps {
  token: string;
  disabled: boolean;
  onTokenChange: (value: string) => void;
}

export function TokenPanel({ token, disabled, onTokenChange }: TokenPanelProps) {
  return (
    <section className="card panel">
      <h2>Graph Access Token</h2>
      <p className="caption">
        No Entra app mode is enabled. Paste a delegated Microsoft Graph bearer token with
        User.Read, Files.Read.All, and Sites.Read.All.
      </p>
      <textarea
        className="token-input"
        rows={5}
        value={token}
        disabled={disabled}
        onChange={(e) => onTokenChange(e.target.value)}
        placeholder="Paste bearer token here (without the word 'Bearer')"
      />
    </section>
  );
}
