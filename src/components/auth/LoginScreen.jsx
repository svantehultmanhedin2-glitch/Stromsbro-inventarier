/* ================= Login Screen ================= */

export default function LoginScreen({ users, onLogin }) {
  const [username, setUsername] = useState(users?.[0]?.username ?? "");
  const [pin, setPin] = useState("");
  const [err, setErr] = useState("");

  useEffect(() => {
    if (users.length && !users.some((u) => u.username === username)) {
      setUsername(users[0].username);
    }
  }, [users, username]);

  const selected = users.find((u) => u.username === username);

  const tryLogin = () => {
    setErr("");
    const u = selected;
    if (!u) return setErr("Välj användare.");
    if (!pinLooksOk(pin)) return setErr("PIN måste vara 4–8 siffror.");
    if (!u.pinSalt || !u.pinHash) return setErr("Användaren saknar PIN. Be Admin sätta en PIN.");

    const h = hashPin(pin, u.pinSalt);
    if (h !== u.pinHash) return setErr("Fel PIN.");

    setPin("");
    onLogin({ id: u.id, name: u.name, username: u.username, role: u.role, lag: u.lag ?? null });
  };

  return (
    <div className="loginWrap">
      <div className="card">
        <div className="card__top">
          <div className="card__title" style={{ fontSize: 20 }}>
            🔐 Logga in
          </div>
          <Pill tone="neutral">PIN</Pill>
        </div>

        <div className="formGrid" style={{ marginTop: 12 }}>
          <label className="field">
            <span>Användare</span>
            <select value={username} onChange={(e) => setUsername(e.target.value)}>
              {users.map((u) => (
                <option key={u.id} value={u.username}>
                  {u.name} ({u.role}
                  {u.role === "Ledare" ? ` • ${u.lag || "Okänt"}` : ""})
                </option>
              ))}
            </select>
          </label>

          <label className="field">
            <span>PIN</span>
            <input
              type="password"
              inputMode="numeric"
              value={pin}
              onChange={(e) => setPin(e.target.value)}
              placeholder="4–8 siffror"
              onKeyDown={(e) => e.key === "Enter" && tryLogin()}
            />
          </label>

          {err ? <div className="banner banner--error">❌ {err}</div> : null}

          <PrimaryButton tone="primary" onClick={tryLogin}>
            Logga in
          </PrimaryButton>

          <div className="muted" style={{ marginTop: 8 }}>
            Första gången efter uppdatering: standard-PIN är <strong>1234</strong> (Admin kan byta).
          </div>
        </div>
      </div>
    </div>
  );
}
