import React, { useCallback, useEffect, useMemo, useRef, useState } from "react";
import * as XLSX from "xlsx";

/* ================= ErrorBoundary ================= */
class ErrorBoundary extends React.Component {
  constructor(props) {
    super(props);
    this.state = { hasError: false, error: null };
  }
  static getDerivedStateFromError(error) {
    return { hasError: true, error };
  }
  render() {
    if (this.state.hasError) {
      return (
        <div style={{ padding: 20, fontFamily: "Arial" }}>
          <h2>⚠️ Något gick fel</h2>
          <p>Feltext:</p>
          <pre
            style={{
              whiteSpace: "pre-wrap",
              background: "#111827",
              color: "#e5e7eb",
              padding: 12,
              borderRadius: 8,
            }}
          >
            {String(this.state.error)}
          </pre>
          <button
            style={{
              marginTop: 12,
              padding: "10px 14px",
              borderRadius: 10,
              border: "1px solid #ccc",
              cursor: "pointer",
            }}
            onClick={() => window.location.reload()}
          >
            Ladda om
          </button>
        </div>
      );
    }
    return this.props.children;
  }
}

/* ================= Storage ================= */
const STORAGE_KEY = "stromsbro-inventarie-v6";

function loadState() {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    return raw ? JSON.parse(raw) : null;
  } catch {
    return null;
  }
}
function saveState(state) {
  try {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(state));
  } catch {}
}

/* ================= Utils ================= */
function normalizeHeader(h) {
  return String(h ?? "")
    .trim()
    .toLowerCase()
    .replace(/\s+/g, " ")
    .replace(/å/g, "a")
    .replace(/ä/g, "a")
    .replace(/ö/g, "o");
}
function toInt(v, fallback = 0) {
  const n = Number(String(v ?? "").replace(",", "."));
  return Number.isFinite(n) ? Math.trunc(n) : fallback;
}
function truthyCell(v) {
  const s = String(v ?? "").trim().toLowerCase();
  if (!s) return false;
  return !(s === "nej" || s === "no" || s === "false" || s === "0");
}
function normKeyPart(v) {
  return String(v ?? "")
    .trim()
    .toLowerCase()
    .replace(/\s+/g, " ")
    .replace(/å/g, "a")
    .replace(/ä/g, "a")
    .replace(/ö/g, "o");
}
function makeKey(huvudgrupp, produkt) {
  return `${normKeyPart(huvudgrupp)}::${normKeyPart(produkt)}`;
}
function downloadXlsx(workbook, filename) {
  const data = XLSX.write(workbook, { bookType: "xlsx", type: "array" });
  const blob = new Blob([data], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  a.rel = "noopener";
  document.body.appendChild(a);
  a.click();
  a.remove();
  URL.revokeObjectURL(url);
}

/* ================= PIN hashing (local-only) =================
   OBS: inte kryptografiskt starkt (offline/local). Bra nog för PIN i localStorage.
*/
function randomSalt(len = 10) {
  const chars = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789";
  let s = "";
  for (let i = 0; i < len; i++) s += chars[Math.floor(Math.random() * chars.length)];
  return s;
}
function fnv1a32(str) {
  let h = 0x811c9dc5;
  for (let i = 0; i < str.length; i++) {
    h ^= str.charCodeAt(i);
    h = (h + ((h << 1) + (h << 4) + (h << 7) + (h << 8) + (h << 24))) >>> 0;
  }
  return ("00000000" + h.toString(16)).slice(-8);
}
function hashPin(pin, salt) {
  return fnv1a32(`${salt}:${String(pin)}`);
}
function pinLooksOk(pin) {
  return /^\d{4,8}$/.test(String(pin ?? "").trim());
}

/* ================= Inköpslogik ================= */
function buildInkopLista(produkter, prevList = []) {
  const prevByKey = new Map((prevList ?? []).map((x) => [x.key, x]));

  const auto = (produkter ?? [])
    .filter((p) => Number.isFinite(p.antal) && Number.isFinite(p.beställningspunkt))
    .filter((p) => toInt(p.beställningspunkt, 0) > 0)
    .filter((p) => toInt(p.antal, 0) <= toInt(p.beställningspunkt, 0))
    .map((p) => {
      const target = toInt(p.minAntal, 0) > 0 ? toInt(p.minAntal, 0) : toInt(p.beställningspunkt, 0);
      const rekommenderat = Math.max(0, target - toInt(p.antal, 0));
      const key = makeKey(p.huvudgrupp, p.produkt);

      const prev = prevByKey.get(key);
      const manuell = prev?.manuell === true;
      const antal = manuell ? Math.max(0, toInt(prev?.antal, 0)) : rekommenderat;

      return {
        id: key,
        key,
        huvudgrupp: p.huvudgrupp ?? "",
        produkt: p.produkt ?? "",
        antal,
        rekommenderat,
        manuell,
        bestalld: prev?.bestalld === true,
        bestalldTid: prev?.bestalldTid ?? "",
        bestalldAv: prev?.bestalldAv ?? "",
        onskemalTid: prev?.onskemalTid ?? "",
        onskemalAv: prev?.onskemalAv ?? "",
        onskemalKommentar: prev?.onskemalKommentar ?? "",

        // ✅ Inleveransdata
        mottagetAntal: Math.max(0, toInt(prev?.mottagetAntal, 0)),
        levererad: prev?.levererad === true,
        levereradTid: prev?.levereradTid ?? "",
        levereradAv: prev?.levereradAv ?? "",
        leveranser: Array.isArray(prev?.leveranser) ? prev.leveranser : [],
      };
    });

  const bestalldaExtras = (prevList ?? [])
    .filter((x) => x?.bestalld === true)
    .filter((x) => !auto.some((a) => a.key === x.key))
    .map((x) => ({
      id: x.key,
      key: x.key,
      huvudgrupp: x.huvudgrupp ?? "",
      produkt: x.produkt ?? "",
      antal: Math.max(0, toInt(x.antal, 0)),
      rekommenderat: 0,
      manuell: x.manuell === true,
      bestalld: true,
      bestalldTid: x.bestalldTid ?? "",
      bestalldAv: x.bestalldAv ?? "",
      onskemalTid: x.onskemalTid ?? "",
      onskemalAv: x.onskemalAv ?? "",
      onskemalKommentar: x.onskemalKommentar ?? "",

      mottagetAntal: Math.max(0, toInt(x?.mottagetAntal, 0)),
      levererad: x?.levererad === true,
      levereradTid: x?.levereradTid ?? "",
      levereradAv: x?.levereradAv ?? "",
      leveranser: Array.isArray(x?.leveranser) ? x.leveranser : [],
    }));

  return [...auto, ...bestalldaExtras];
}

/* ================= UI bits ================= */
function Pill({ tone = "neutral", children }) {
  return <span className={`pill pill--${tone}`}>{children}</span>;
}
function PrimaryButton({ onClick, children, tone = "primary", type = "button", disabled = false, style }) {
  return (
    <button className={`btn btn--${tone}`} onClick={onClick} type={type} disabled={disabled} style={style}>
      {children}
    </button>
  );
}
function Modal({ open, title, onClose, children, footer }) {
  if (!open) return null;
  return (
    <div className="modalOverlay" role="dialog" aria-modal="true">
      <div className="modalSheet">
        <div className="modalHeader">
          <div className="modalTitle">{title}</div>
          <button className="iconBtn" onClick={onClose} aria-label="Stäng" type="button">
            ✕
          </button>
        </div>
        <div className="modalBody">{children}</div>
        {footer ? <div className="modalFooter">{footer}</div> : null}
      </div>
    </div>
  );
}

/* ================= Toast Hook ================= */
function useToast(timeoutMs = 2500) {
  const [info, setInfo] = useState("");
  const tRef = useRef(null);

  const show = useCallback(
    (msg) => {
      setInfo(msg);
      if (tRef.current) window.clearTimeout(tRef.current);
      tRef.current = window.setTimeout(() => setInfo(""), timeoutMs);
    },
    [timeoutMs]
  );

  useEffect(() => {
    return () => {
      if (tRef.current) window.clearTimeout(tRef.current);
    };
  }, []);

  return { info, show, setInfo };
}

/* ================= Login Screen (render-only) ================= */
function LoginScreen({ users, onLogin }) {
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

/* ================= App ================= */
export default function App() {
  const saved = loadState();

  /* ===== Auth/User state ===== */
  const [users, setUsers] = useState(() => {
    const base =
      Array.isArray(saved?.users) && saved.users.length
        ? saved.users
        : [
            { id: 1, name: "Admin", username: "admin", role: "Admin", lag: null },
            { id: 2, name: "Materialansvarig", username: "ma", role: "Materialansvarig", lag: null },
            { id: 3, name: "Ledare P12", username: "p12", role: "Ledare", lag: "P12" },
          ];

    return base.map((u) => {
      if (u.pinSalt && u.pinHash) return u;
      const salt = randomSalt();
      return { ...u, pinSalt: salt, pinHash: hashPin("1234", salt) };
    });
  });

  const [currentUser, setCurrentUser] = useState(() => saved?.currentUser ?? null);

  useEffect(() => {
    if (currentUser && (!currentUser.role || !currentUser.name)) {
      console.error("Korrupt currentUser:", currentUser);
      setCurrentUser(null);
    }
  }, [currentUser]);

  /* ===== Behörigheter ===== */
  const isAdmin = currentUser?.role === "Admin";
  const isMA = currentUser?.role === "Materialansvarig";
  const isLedare = currentUser?.role === "Ledare";

  const canEditUsers = isAdmin;
  const canSeeInkop = isAdmin || isMA;
  const canSeeInleverans = isAdmin || isMA;
  const canMoveToInkop = isAdmin || isMA;
  const canUtlamna = isAdmin || isMA;
  const canEditHuvudlager = isAdmin || isMA;
  const canImportExport = isAdmin || isMA;

  /* ===== App state ===== */
  const { info, show: showInfo } = useToast(2500);
  const [fel, setFel] = useState("");
  const [vy, setVy] = useState(saved?.vy ?? "lager");
  const [sok, setSok] = useState("");

  const [produkter, setProdukter] = useState(
    Array.isArray(saved?.produkter)
      ? saved.produkter
      : [
          {
            id: 1,
            huvudgrupp: "Bollar",
            produkt: "Fotboll strl 5",
            antal: 12,
            lagerplats: "Huvudförråd",
            minAntal: 10,
            beställningspunkt: 15,
          },
          {
            id: 2,
            huvudgrupp: "Västar",
            produkt: "Träningsvästar gula",
            antal: 6,
            lagerplats: "Lilla förrådet",
            minAntal: 8,
            beställningspunkt: 10,
          },
        ]
  );

  const [inkopslista, setInkopslista] = useState(() => {
    if (Array.isArray(saved?.inkopslista)) return saved.inkopslista;
    return buildInkopLista(Array.isArray(saved?.produkter) ? saved.produkter : [], []);
  });

  const [onskemal, setOnskemal] = useState(Array.isArray(saved?.onskemal) ? saved.onskemal : []);
  const [historik, setHistorik] = useState(Array.isArray(saved?.historik) ? saved.historik : []);

  const [lagLager, setLagLager] = useState(Array.isArray(saved?.lagLager) ? saved.lagLager : []);
  const [aktivtLag, setAktivtLag] = useState(saved?.aktivtLag ?? "P12");

  const [recentComments, setRecentComments] = useState(
    Array.isArray(saved?.recentComments) ? saved.recentComments : []
  );
  const rememberComment = (c) => {
    const s = String(c ?? "").trim();
    if (!s) return;
    setRecentComments((prev) => [s, ...prev.filter((x) => x !== s)].slice(0, 8));
  };

  /* ===== Qty per product ===== */
  const [qtyMap, setQtyMap] = useState(() =>
    saved?.qtyMap && typeof saved.qtyMap === "object" ? saved.qtyMap : {}
  );
  const getQty = (id) => Math.max(1, toInt(qtyMap[id], 1));
  const setQty = (id, value) => {
    const n = Math.max(1, toInt(value, 1));
    setQtyMap((prev) => ({ ...prev, [id]: n }));
  };

  /* ===== Inleverans state ===== */
  const [leveransKommentar, setLeveransKommentar] = useState(saved?.leveransKommentar ?? "");
  const [leveransLagerplats, setLeveransLagerplats] = useState(saved?.leveransLagerplats ?? "Huvudförråd");
  const [leveransQtyMap, setLeveransQtyMap] = useState(() =>
    saved?.leveransQtyMap && typeof saved.leveransQtyMap === "object" ? saved.leveransQtyMap : {}
  );
  const getLevQty = (key, fallback = 1) => Math.max(1, toInt(leveransQtyMap[key] ?? fallback, 1));
  const setLevQty = (key, value) => {
    const n = Math.max(1, toInt(value, 1));
    setLeveransQtyMap((prev) => ({ ...prev, [key]: n }));
  };

  const lagerplatsOptions = useMemo(() => {
    const set = new Set(["Huvudförråd", "Lilla förrådet", "Materialrum", "Lagret"]);
    (produkter ?? []).forEach((p) => {
      const lp = String(p.lagerplats ?? "").trim();
      if (lp) set.add(lp);
    });
    return Array.from(set).filter(Boolean).sort((a, b) => a.localeCompare(b, "sv"));
  }, [produkter]);

  /* ===== Forms ===== */
  const [nyttOnskemal, setNyttOnskemal] = useState({ huvudgrupp: "", produkt: "", antal: 1, kommentar: "" });

  const [flytt, setFlytt] = useState({
    lag: saved?.aktivtLag ?? "P12",
    produktId: "",
    antal: 1,
    utlamningsdatum: "",
    kommentar: "",
  });

  const [nyLagRad, setNyLagRad] = useState({
    lag: saved?.aktivtLag ?? "P12",
    huvudgrupp: "",
    produkt: "",
    antal: 1,
    utlamningsdatum: "",
    kommentar: "",
  });

  /* ===== Persist ALL state ===== */
  useEffect(() => {
    saveState({
      vy,
      users,
      currentUser,
      produkter,
      inkopslista,
      onskemal,
      historik,
      lagLager,
      aktivtLag,
      recentComments,
      qtyMap,
      leveransKommentar,
      leveransLagerplats,
      leveransQtyMap,
    });
  }, [
    vy,
    users,
    currentUser,
    produkter,
    inkopslista,
    onskemal,
    historik,
    lagLager,
    aktivtLag,
    recentComments,
    qtyMap,
    leveransKommentar,
    leveransLagerplats,
    leveransQtyMap,
  ]);

  /* ===== View guarding ===== */
  useEffect(() => {
    if (!currentUser) return;
    if (isLedare && (vy === "inkop" || vy === "admin" || vy === "inleverans")) setVy("lagmaterial");
    if (!canSeeInkop && vy === "inkop") setVy("lager");
    if (!canEditUsers && vy === "admin") setVy("lager");
    if (!canSeeInleverans && vy === "inleverans") setVy("lager");
  }, [vy, currentUser, isLedare, canSeeInkop, canEditUsers, canSeeInleverans]);

  /* ===== Derived ===== */
  const status = (p) => {
    const antal = toInt(p.antal, 0);
    const min = toInt(p.minAntal, 0);
    const bp = toInt(p.beställningspunkt, 0);
    if (min <= 0 && bp <= 0) return { text: "OK", tone: "ok" };
    if (min > 0 && antal <= min) return { text: "Beställ", tone: "danger" };
    if (bp > 0 && antal <= bp) return { text: "Bevaka", tone: "warn" };
    return { text: "OK", tone: "ok" };
  };

  const filtreradeProdukter = useMemo(() => {
    const q = sok.trim().toLowerCase();
    if (!q) return produkter;
    return produkter.filter((p) => `${p.produkt} ${p.huvudgrupp} ${p.lagerplats}`.toLowerCase().includes(q));
  }, [produkter, sok]);

  const filtreradInkop = useMemo(() => {
    const q = sok.trim().toLowerCase();
    if (!q) return inkopslista;
    return inkopslista.filter((r) => `${r.produkt} ${r.huvudgrupp}`.toLowerCase().includes(q));
  }, [inkopslista, sok]);

  const inkopAttVisa = useMemo(() => filtreradInkop.filter((r) => !r.bestalld), [filtreradInkop]);

  const totalInkopAntal = useMemo(() => {
    return filtreradInkop
      .filter((r) => !r.bestalld)
      .reduce((sum, r) => sum + Math.max(0, toInt(r.antal, 0)), 0);
  }, [filtreradInkop]);

  const lagLista = useMemo(() => {
    const set = new Set(["P12", "F13", "A-lag", "U-lag"]);
    (lagLager ?? []).forEach((r) => set.add(r.lag || "Okänt"));
    return Array.from(set).filter(Boolean).sort();
  }, [lagLager]);

  const aktivtLagEff = isLedare ? (currentUser?.lag ?? "Okänt") : aktivtLag;

  const lagRaderVisning = useMemo(() => {
    const q = sok.trim().toLowerCase();
    const base = (lagLager ?? []).filter((r) => (r.lag || "Okänt") === aktivtLagEff);
    if (!q) return base;
    return base.filter((r) => `${r.produkt} ${r.huvudgrupp} ${r.kommentar} ${r.utlamningsdatum}`.toLowerCase().includes(q));
  }, [lagLager, aktivtLagEff, sok]);

  const lagTotalAntal = useMemo(
    () => lagRaderVisning.reduce((sum, r) => sum + Math.max(0, toInt(r.antal, 0)), 0),
    [lagRaderVisning]
  );

  /* ================= Single Tx Modal (lager) ================= */
  const [txOpen, setTxOpen] = useState(false);
  const [txMode, setTxMode] = useState("uttag");
  const [txProductId, setTxProductId] = useState(null);
  const [txQty, setTxQty] = useState(1);
  const [txComment, setTxComment] = useState("");

  const openTx = (mode, product) => {
    setFel("");
    setTxMode(mode);
    setTxProductId(product.id);
    setTxQty(getQty(product.id));
    setTxComment("");
    setTxOpen(true);
  };
  const closeTx = () => {
    setTxOpen(false);
    setTxProductId(null);
    setTxQty(1);
    setTxComment("");
  };

  const andringDirekt = (produkt, delta, kommentar) => {
    setFel("");

    if (delta < 0 && toInt(produkt.antal, 0) + delta < 0) {
      setFel("Kan inte ta ut fler än det som finns i lager.");
      return;
    }

    const nyaProdukter = produkter.map((p) =>
      p.id === produkt.id ? { ...p, antal: Math.max(0, toInt(p.antal, 0) + delta) } : p
    );
    setProdukter(nyaProdukter);
    setInkopslista((prev) => buildInkopLista(nyaProdukter, prev));

    const kom = kommentar?.trim() ? kommentar.trim() : "(ingen kommentar)";
    rememberComment(kom);

    setHistorik((prev) => [
      {
        tid: new Date().toLocaleString("sv-SE"),
        typ: delta < 0 ? "Uttag" : "Inleverans",
        produkt: produkt.produkt,
        huvudgrupp: produkt.huvudgrupp,
        lagerplats: produkt.lagerplats,
        antal: delta,
        användare: currentUser?.name ?? "Okänd",
        kommentar: kom,
      },
      ...prev,
    ]);
  };

  /* ================= Massuttag ================= */
  const [selectMode, setSelectMode] = useState(false);
  const [selectedIds, setSelectedIds] = useState(() => new Set());
  const selectedCount = selectedIds.size;

  const toggleSelected = (id) => {
    setSelectedIds((prev) => {
      const next = new Set(prev);
      if (next.has(id)) next.delete(id);
      else next.add(id);
      return next;
    });
  };
  const clearSelected = () => setSelectedIds(new Set());
  const selectAllVisible = () => setSelectedIds(new Set(filtreradeProdukter.map((p) => p.id)));

  useEffect(() => {
    if (vy !== "lager" && selectMode) {
      setSelectMode(false);
      clearSelected();
    }
  }, [vy]); // eslint-disable-line react-hooks/exhaustive-deps

  const [batchOpen, setBatchOpen] = useState(false);
  const [batchMode, setBatchMode] = useState("uttag");
  const [batchUsePerItemQty, setBatchUsePerItemQty] = useState(true);
  const [batchQty, setBatchQty] = useState(1);
  const [batchComment, setBatchComment] = useState("");

  const openBatch = (mode) => {
    setFel("");
    if (selectedIds.size === 0) {
      setFel("Markera minst en produkt först.");
      return;
    }
    setBatchMode(mode);
    setBatchOpen(true);
  };
  const closeBatch = () => {
    setBatchOpen(false);
    setBatchComment("");
    setBatchQty(1);
    setBatchUsePerItemQty(true);
  };

  const applyBatch = () => {
    setFel("");

    const ids = Array.from(selectedIds);
    const items = produkter.filter((p) => ids.includes(p.id));
    if (items.length === 0) {
      setFel("Inga produkter valda.");
      return;
    }

    const qtyFor = (p) => {
      if (batchUsePerItemQty) return Math.max(1, toInt(getQty(p.id), 1));
      return Math.max(1, toInt(batchQty, 1));
    };

    const deltas = items.map((p) => {
      const q = qtyFor(p);
      const delta = batchMode === "uttag" ? -q : +q;
      return { p, delta };
    });

    if (batchMode === "uttag") {
      const bad = deltas.find(({ p, delta }) => toInt(p.antal, 0) + delta < 0);
      if (bad) {
        setFel(`Kan inte ta ut ${Math.abs(bad.delta)} st av "${bad.p.produkt}" (i lager: ${bad.p.antal}).`);
        return;
      }
    }

    const deltaById = new Map(deltas.map(({ p, delta }) => [p.id, delta]));
    const nyaProdukter = produkter.map((p) =>
      deltaById.has(p.id) ? { ...p, antal: Math.max(0, toInt(p.antal, 0) + deltaById.get(p.id)) } : p
    );

    setProdukter(nyaProdukter);
    setInkopslista((prev) => buildInkopLista(nyaProdukter, prev));

    const tid = new Date().toLocaleString("sv-SE");
    const typ = batchMode === "uttag" ? "Uttag (mass)" : "Inleverans (mass)";
    const kommentar = batchComment?.trim() ? batchComment.trim() : "(ingen kommentar)";
    rememberComment(kommentar);

    const historikRader = deltas.map(({ p, delta }) => ({
      tid,
      typ,
      produkt: p.produkt,
      huvudgrupp: p.huvudgrupp,
      lagerplats: p.lagerplats,
      antal: delta,
      användare: currentUser?.name ?? "Okänd",
      kommentar,
    }));
    setHistorik((prev) => [...historikRader, ...prev]);

    showInfo(batchMode === "uttag" ? "Massuttag registrerat." : "Massinleverans registrerad.");
    closeBatch();
    clearSelected();
    setSelectMode(false);
  };

  const [showUtlamning, setShowUtlamning] = useState(false);
  const [showAddLagRad, setShowAddLagRad] = useState(false);

  /* ================= Lagmaterial ================= */
  const flyttaTillLag = () => {
    if (!canUtlamna) return;
    const prodId = flytt.produktId;
    const antal = Math.max(1, toInt(flytt.antal, 1));
    const lag = (flytt.lag ?? aktivtLagEff ?? "Okänt").trim() || "Okänt";

    const p = produkter.find((x) => String(x.id) === String(prodId));
    if (!p) {
      setFel("Välj en produkt i huvudlagret.");
      return;
    }
    if (toInt(p.antal, 0) < antal) {
      setFel("Finns inte så många i huvudlagret.");
      return;
    }

    setFel("");

    const nyaProdukter = produkter.map((x) => (x.id === p.id ? { ...x, antal: toInt(x.antal, 0) - antal } : x));
    setProdukter(nyaProdukter);
    setInkopslista((prev) => buildInkopLista(nyaProdukter, prev));

    setLagLager((prev) => {
      const exists = prev.find(
        (r) =>
          (r.lag || "Okänt") === lag &&
          normKeyPart(r.produkt) === normKeyPart(p.produkt) &&
          normKeyPart(r.huvudgrupp) === normKeyPart(p.huvudgrupp)
      );

      const now = new Date().toLocaleString("sv-SE");
      const ut = (flytt.utlamningsdatum ?? "").trim();
      const kom = (flytt.kommentar ?? "").trim();

      if (exists) {
        return prev.map((r) =>
          r.id === exists.id
            ? {
                ...r,
                antal: Math.max(0, toInt(r.antal, 0)) + antal,
                utlamningsdatum: ut || r.utlamningsdatum || "",
                kommentar: kom || r.kommentar || "",
              }
            : r
        );
      }

      return [
        {
          id: Date.now(),
          lag,
          huvudgrupp: p.huvudgrupp ?? "",
          produkt: p.produkt ?? "",
          antal,
          utlamningsdatum: ut,
          kommentar: kom,
          skapadTid: now,
          skapadAv: currentUser?.name ?? "Okänd",
        },
        ...prev,
      ];
    });

    setHistorik((prev) => [
      {
        tid: new Date().toLocaleString("sv-SE"),
        typ: "Utlämnat till lag",
        produkt: p.produkt,
        huvudgrupp: p.huvudgrupp,
        lagerplats: p.lagerplats,
        antal: -antal,
        användare: currentUser?.name ?? "Okänd",
        kommentar: `Till ${lag}${flytt.kommentar ? ` • ${flytt.kommentar}` : ""}`,
      },
      ...prev,
    ]);

    showInfo(`Flyttade ${antal} st till ${lag}.`);
    setFlytt((f) => ({ ...f, antal: 1, kommentar: "" }));
  };

  const uppdateraLagRadAntal = (id, newValue) => {
    const n = Math.max(0, toInt(newValue, 0));
    setLagLager((prev) => prev.map((r) => (r.id === id ? { ...r, antal: n } : r)));
    showInfo("Antal uppdaterat.");
  };

  const laggTillLagRad = () => {
    if (isLedare) {
      setFel("Ledare kan inte lägga till nya rader i lagmaterial (ändra antal på befintliga istället).");
      return;
    }
    const prod = (nyLagRad.produkt ?? "").trim();
    if (!prod) return;

    const item = {
      id: Date.now(),
      lag: (nyLagRad.lag ?? aktivtLagEff ?? "Okänt").trim() || "Okänt",
      huvudgrupp: (nyLagRad.huvudgrupp ?? "").trim(),
      produkt: prod,
      antal: Math.max(0, toInt(nyLagRad.antal, 1)),
      utlamningsdatum: (nyLagRad.utlamningsdatum ?? "").trim(),
      kommentar: (nyLagRad.kommentar ?? "").trim(),
      skapadTid: new Date().toLocaleString("sv-SE"),
      skapadAv: currentUser?.name ?? "Okänd",
    };

    setLagLager((prev) => [item, ...prev]);
    setNyLagRad({ lag: item.lag, huvudgrupp: "", produkt: "", antal: 1, utlamningsdatum: "", kommentar: "" });
    showInfo("Rad lades till i lagmaterial.");
  };

  const taBortLagRad = (id) => {
    if (isLedare) {
      setFel("Ledare kan inte ta bort rader (ändra antal istället).");
      return;
    }
    setLagLager((prev) => prev.filter((x) => x.id !== id));
    showInfo("Rad borttagen.");
  };

  const [returOpen, setReturOpen] = useState(false);
  const [returRad, setReturRad] = useState(null);
  const [returQty, setReturQty] = useState(1);

  const openRetur = (rad) => {
    if (!canUtlamna) return;
    setFel("");
    setReturRad(rad);
    setReturQty(1);
    setReturOpen(true);
  };
  const closeRetur = () => {
    setReturOpen(false);
    setReturRad(null);
    setReturQty(1);
  };

  const returTillHuvudlager = (rad, antalInput) => {
    if (!canUtlamna) return;
    const antal = Math.max(1, toInt(antalInput, 1));
    if (antal > (rad.antal ?? 0)) {
      setFel("Kan inte returnera fler än lagets antal.");
      return;
    }
    setFel("");

    setProdukter((prev) => {
      const idx = prev.findIndex(
        (p) => normKeyPart(p.produkt) === normKeyPart(rad.produkt) && normKeyPart(p.huvudgrupp) === normKeyPart(rad.huvudgrupp)
      );

      if (idx >= 0) {
        const nya = [...prev];
        nya[idx] = { ...nya[idx], antal: Math.max(0, toInt(nya[idx].antal, 0)) + antal };
        setInkopslista((inkPrev) => buildInkopLista(nya, inkPrev));
        return nya;
      }

      const ny = {
        id: Date.now(),
        huvudgrupp: rad.huvudgrupp ?? "",
        produkt: rad.produkt ?? "",
        antal,
        lagerplats: "Huvudförråd",
        minAntal: 0,
        beställningspunkt: 0,
      };
      const nya = [ny, ...prev];
      setInkopslista((inkPrev) => buildInkopLista(nya, inkPrev));
      return nya;
    });

    setLagLager((prev) => {
      const kvar = prev.map((r) => (r.id === rad.id ? { ...r, antal: Math.max(0, toInt(r.antal, 0) - antal) } : r));
      return kvar.filter((r) => toInt(r.antal, 0) > 0);
    });

    setHistorik((prev) => [
      {
        tid: new Date().toLocaleString("sv-SE"),
        typ: "Retur från lag",
        produkt: rad.produkt,
        huvudgrupp: rad.huvudgrupp,
        lagerplats: "",
        antal: +antal,
        användare: currentUser?.name ?? "Okänd",
        kommentar: `Från ${rad.lag || "Okänt"}`,
      },
      ...prev,
    ]);

    showInfo(`Returnerade ${antal} st till huvudlager.`);
  };

  /* ================= Önskemål ================= */
  const skapaOnskemal = () => {
    const prod = nyttOnskemal.produkt.trim();
    if (!prod) return;

    const item = {
      id: Date.now(),
      huvudgrupp: nyttOnskemal.huvudgrupp.trim(),
      produkt: prod,
      antal: Math.max(1, toInt(nyttOnskemal.antal, 1)),
      kommentar: (nyttOnskemal.kommentar ?? "").trim(),
      skapadTid: new Date().toLocaleString("sv-SE"),
      skapadAv: currentUser?.name ?? "Okänd",
      lag: isLedare ? (currentUser?.lag ?? "Okänt") : "",
    };

    setOnskemal((prev) => [item, ...prev]);
    setNyttOnskemal({ huvudgrupp: "", produkt: "", antal: 1, kommentar: "" });
    showInfo("Önskemålet lades till.");
  };

  const laggOnskemalTillInkop = (o) => {
    if (!canMoveToInkop) return;

    const key = makeKey(o.huvudgrupp, o.produkt);

    setInkopslista((prev) => {
      const finns = prev.find((r) => r.key === key);
      if (finns) {
        return prev.map((r) =>
          r.key === key
            ? {
                ...r,
                antal: Math.max(0, toInt(r.antal, 0)) + Math.max(1, toInt(o.antal, 1)),
                manuell: true,
                bestalld: false,
                onskemalTid: r.onskemalTid || o.skapadTid || "",
                onskemalAv: r.onskemalAv || o.skapadAv || "",
                onskemalKommentar: r.onskemalKommentar || o.kommentar || "",
              }
            : r
        );
      }
      return [
        ...prev,
        {
          id: key,
          key,
          huvudgrupp: o.huvudgrupp ?? "",
          produkt: o.produkt ?? "",
          antal: Math.max(1, toInt(o.antal, 1)),
          rekommenderat: 0,
          manuell: true,
          bestalld: false,
          bestalldTid: "",
          bestalldAv: "",
          onskemalTid: o.skapadTid ?? "",
          onskemalAv: o.skapadAv ?? "",
          onskemalKommentar: o.kommentar ?? "",
          mottagetAntal: 0,
          levererad: false,
          levereradTid: "",
          levereradAv: "",
          leveranser: [],
        },
      ];
    });

    setOnskemal((prev) => prev.filter((x) => x.id !== o.id));
    showInfo("Önskemål flyttat till inköp.");
  };

  /* ================= Inköp ================= */
  const sättInköpsantalManuellt = (key, value) => {
    setInkopslista((prev) => prev.map((r) => (r.key === key ? { ...r, antal: Math.max(0, toInt(value, 0)), manuell: true } : r)));
  };
  const återställRekommenderat = (key) => {
    setInkopslista((prev) => prev.map((r) => (r.key === key ? { ...r, antal: r.rekommenderat, manuell: false } : r)));
  };
  const markeraSomBestalld = (key) => {
    const tid = new Date().toLocaleString("sv-SE");
    setInkopslista((prev) =>
      prev.map((r) =>
        r.key === key
          ? {
              ...r,
              bestalld: true,
              bestalldTid: tid,
              bestalldAv: currentUser?.name ?? "Okänd",
              mottagetAntal: Math.max(0, toInt(r.mottagetAntal, 0)),
              levererad: r.levererad === true,
              leveranser: Array.isArray(r.leveranser) ? r.leveranser : [],
            }
          : r
      )
    );
    showInfo("Markerad som beställd.");
  };
  const angraBestalld = (key) => {
    setInkopslista((prev) => prev.map((r) => (r.key === key ? { ...r, bestalld: false, bestalldTid: "", bestalldAv: "" } : r)));
    showInfo("Ångrade beställd.");
  };

  /* ================= Inleverans: motta delleverans ================= */
const mottaLeverans = (rowKey, qtyInput) => {
  if (!canSeeInleverans) return;

  const row = inkopslista.find((r) => r.key === rowKey);
  if (!row || !row.bestalld) {
    setFel("Välj en beställd rad.");
    return;
  }

  const ordered = Math.max(0, toInt(row.antal, 0));
  const received = Math.max(0, toInt(row.mottagetAntal, 0));
  const kvar = Math.max(0, ordered - received);

  const qtyReq = Math.max(1, toInt(qtyInput, 1));
  const qty = Math.min(kvar, qtyReq);
  if (qty <= 0) {
    setFel("Inget kvar att ta emot.");
    return;
  }

  setFel("");
  const tid = new Date().toLocaleString("sv-SE");
  const av = currentUser?.name ?? "Okänd";
  const kom = String(leveransKommentar ?? "").trim();

  /* ✅ 1. Uppdatera HUVUDLAGER – EN gång */
  setProdukter((prev) => {
    const idx = prev.findIndex(
      (p) =>
        normKeyPart(p.produkt) === normKeyPart(row.produkt) &&
        normKeyPart(p.huvudgrupp) === normKeyPart(row.huvudgrupp)
    );

    if (idx >= 0) {
      const nya = [...prev];
      nya[idx] = {
        ...nya[idx],
        antal: Math.max(0, toInt(nya[idx].antal, 0)) + qty,
        lagerplats: leveransLagerplats,
      };
      return nya;
    }

    return [
      {
        id: Date.now(),
        huvudgrupp: row.huvudgrupp ?? "",
        produkt: row.produkt ?? "",
        antal: qty,
        lagerplats: leveransLagerplats,
        minAntal: 0,
        beställningspunkt: 0,
      },
      ...prev,
    ];
  });

  /* ✅ 2. Uppdatera INKÖPSRAD – UTAN rebuild */
  setInkopslista((prev) =>
    prev.map((r) => {
      if (r.key !== rowKey) return r;

      const newReceived = received + qty;
      const fully = ordered > 0 && newReceived >= ordered;

      return {
        ...r,
        mottagetAntal: newReceived,
        levererad: fully,
        levereradTid: fully ? tid : r.levereradTid,
        levereradAv: fully ? av : r.levereradAv,
        leveranser: [
          {
            tid,
            av,
            antal: qty,
            lagerplats: leveransLagerplats,
            kommentar: kom,
          },
          ...(r.leveranser || []),
        ],
      };
    })
  );

  /* ✅ 3. Historik */
  const kommentarText = `Inleverans → ${leveransLagerplats}${kom ? ` • ${kom}` : ""}`;
  setHistorik((prev) => [
    {
      tid,
      typ: "Inleverans (beställning)",
      produkt: row.produkt,
      huvudgrupp: row.huvudgrupp,
      lagerplats: leveransLagerplats,
      antal: qty,
      användare: av,
      kommentar: kommentarText,
    },
    ...prev,
  ]);

  showInfo(`Tog emot ${qty} st: ${row.produkt}`);
};

  /* ================= Admin: användare + PIN ================= */
  const [userDraft, setUserDraft] = useState({ name: "", username: "", role: "Ledare", lag: "P12", pin: "" });

  const addUser = () => {
    if (!canEditUsers) return;

    const name = userDraft.name.trim();
    const username = userDraft.username.trim().toLowerCase();
    const role = userDraft.role;
    const lag = role === "Ledare" ? (userDraft.lag || "").trim() : null;
    const pin = (userDraft.pin || "").trim();

    if (!name || !username) return setFel("Namn och användarnamn krävs.");
    if (users.some((u) => u.username === username)) return setFel("Användarnamn finns redan.");
    if (!pinLooksOk(pin)) return setFel("PIN måste vara 4–8 siffror.");

    const salt = randomSalt();
    const u = {
      id: Date.now(),
      name,
      username,
      role,
      lag: lag || null,
      pinSalt: salt,
      pinHash: hashPin(pin, salt),
    };

    setUsers((prev) => [u, ...prev]);
    setUserDraft({ name: "", username: "", role: "Ledare", lag: "P12", pin: "" });
    showInfo("Användare skapad.");
  };

  const deleteUser = (id) => {
    if (!canEditUsers) return;
    if (currentUser?.id === id) {
      setFel("Du kan inte ta bort dig själv.");
      return;
    }
    setUsers((prev) => prev.filter((u) => u.id !== id));
    showInfo("Användare borttagen.");
  };

  const [pinEdit, setPinEdit] = useState({ userId: "", newPin: "" });

  const setUserPin = () => {
    if (!canEditUsers) return;
    const id = Number(pinEdit.userId);
    const newPin = (pinEdit.newPin || "").trim();
    if (!id) return setFel("Välj användare.");
    if (!pinLooksOk(newPin)) return setFel("PIN måste vara 4–8 siffror.");

    setUsers((prev) =>
      prev.map((u) => {
        if (u.id !== id) return u;
        const salt = randomSalt();
        return { ...u, pinSalt: salt, pinHash: hashPin(newPin, salt) };
      })
    );

    showInfo("PIN uppdaterad.");
    setPinEdit({ userId: "", newPin: "" });
  };

  /* ================= Import/Export ================= */
  const fileInputRef = useRef(null);

  const importeraExcel = async (file) => {
    if (!canImportExport) {
      setFel("Du saknar behörighet för import.");
      return;
    }
    setFel("");
    try {
      const buf = await file.arrayBuffer();
      const wb = XLSX.read(buf, { type: "array" });

      const sheetName = wb.SheetNames.includes("Huvudlager") ? "Huvudlager" : wb.SheetNames[0];
      const ws = wb.Sheets[sheetName];
      const rows = XLSX.utils.sheet_to_json(ws, { defval: "" });
      if (!rows.length) throw new Error(`Bladet "${sheetName}" verkar vara tomt.`);

      const headerMap = {};
      Object.keys(rows[0]).forEach((h) => (headerMap[normalizeHeader(h)] = h));
      const col = (name) => headerMap[normalizeHeader(name)];
      const pick = (r, ...names) => {
        for (const name of names) {
          const c = col(name);
          if (c && r[c] !== undefined) return r[c];
        }
        return "";
      };

      const produkterNy = rows
        .map((r, idx) => {
          const prod = pick(r, "Produkt", "produkt");
          if (!String(prod).trim()) return null;
          return {
            id: Date.now() + idx,
            huvudgrupp: String(pick(r, "Huvudgrupp") ?? "").trim(),
            produkt: String(prod).trim(),
            antal: toInt(pick(r, "Antal"), 0),
            lagerplats: String(pick(r, "Lagerplats") ?? "").trim(),
            beställningspunkt: toInt(pick(r, "Beställningspunkt", "Bestallningspunkt"), 0),
            minAntal: toInt(pick(r, "Min antal", "MinAntal", "Min"), 0),
          };
        })
        .filter(Boolean);

      let prevInkop = [];
      if (wb.SheetNames.includes("MaterialÖnskemål")) {
        const wsM = wb.Sheets["MaterialÖnskemål"];
        const rowsM = XLSX.utils.sheet_to_json(wsM, { defval: "" });
        if (rowsM.length) {
          const hm = {};
          Object.keys(rowsM[0]).forEach((h) => (hm[normalizeHeader(h)] = h));
          const cm = (name) => hm[normalizeHeader(name)];

          const manuellCol = cm("Manuell");
          const hasManuellCol = Boolean(manuellCol);

          prevInkop = rowsM
            .map((r) => {
              const prod = r[cm("Produkt")] ?? "";
              if (!String(prod).trim()) return null;
              const hg = String(r[cm("Huvudgrupp")] ?? "").trim();
              const key = makeKey(hg, prod);

              return {
                id: key,
                key,
                huvudgrupp: hg,
                produkt: String(prod).trim(),
                antal: toInt(r[cm("Antal")], 0),
                rekommenderat: 0,
                manuell: hasManuellCol ? truthyCell(r[manuellCol]) : true,
                bestalld: truthyCell(r[cm("Beställd")] ?? r[cm("Bestalld")]),
                bestalldTid: String(r[cm("Beställd tid")] ?? r[cm("Bestalld tid")] ?? "").trim(),
                bestalldAv: String(r[cm("Beställd av")] ?? r[cm("Bestalld av")] ?? "").trim(),

                mottagetAntal: toInt(r[cm("Mottaget")], 0),
                levererad: truthyCell(r[cm("Levererad")]),
                levereradTid: String(r[cm("Levererad tid")] ?? "").trim(),
                levereradAv: String(r[cm("Levererad av")] ?? "").trim(),
                leveranser: [],
              };
            })
            .filter(Boolean);
        }
      }

      let lagNy = [];
      if (wb.SheetNames.includes("Lagmaterial")) {
        const wsL = wb.Sheets["Lagmaterial"];
        const rowsL = XLSX.utils.sheet_to_json(wsL, { defval: "" });
        if (rowsL.length) {
          const hl = {};
          Object.keys(rowsL[0]).forEach((h) => (hl[normalizeHeader(h)] = h));
          const cl = (name) => hl[normalizeHeader(name)];

          lagNy = rowsL
            .map((r, idx) => {
              const prod = r[cl("Produkt")] ?? "";
              if (!String(prod).trim()) return null;
              return {
                id: Date.now() + 5000 + idx,
                lag: String(r[cl("Lag")] ?? r[cl("lag")] ?? aktivtLagEff ?? "Okänt").trim() || (aktivtLagEff ?? "Okänt"),
                huvudgrupp: String(r[cl("Huvudgrupp")] ?? "").trim(),
                produkt: String(prod).trim(),
                antal: toInt(r[cl("Antal")], 0),
                utlamningsdatum: String(r[cl("Utlämningsdatum")] ?? r[cl("Utlamningsdatum")] ?? "").trim(),
                kommentar: String(r[cl("Kommentar")] ?? "").trim(),
                skapadTid: "",
                skapadAv: "",
              };
            })
            .filter(Boolean);
        }
      }

      setProdukter(produkterNy);
      setInkopslista(buildInkopLista(produkterNy, prevInkop));
      if (lagNy.length) setLagLager(lagNy);

      showInfo(`Import klar (${produkterNy.length} produkter)${lagNy.length ? ` + Lagmaterial (${lagNy.length})` : ""}`);
      if (fileInputRef.current) fileInputRef.current.value = "";
      setVy("lager");
    } catch (e) {
      setFel(`Import misslyckades: ${e.message}`);
    }
  };

  const exporteraExcel = () => {
    if (!canImportExport) {
      setFel("Du saknar behörighet för export.");
      return;
    }
    setFel("");
    try {
      const lagerRows = produkter.map((p) => ({
        Huvudgrupp: p.huvudgrupp,
        Produkt: p.produkt,
        Antal: p.antal,
        Lagerplats: p.lagerplats,
        Beställningspunkt: p.beställningspunkt,
        "Min antal": p.minAntal,
        Status: status(p).text,
      }));

      const inkopRows = inkopslista.map((r) => ({
        Huvudgrupp: r.huvudgrupp,
        Produkt: r.produkt,
        Antal: r.antal,
        Rekommenderat: r.rekommenderat,
        Manuell: r.manuell ? "Ja" : "Nej",
        Beställd: r.bestalld ? "Ja" : "Nej",
        "Beställd tid": r.bestalldTid || "",
        "Beställd av": r.bestalldAv || "",
        Mottaget: toInt(r.mottagetAntal, 0),
        Levererad: r.levererad ? "Ja" : "Nej",
        "Levererad tid": r.levereradTid || "",
        "Levererad av": r.levereradAv || "",
      }));

      const onskemalRows = onskemal.map((o) => ({
        Produkt: o.produkt,
        Huvudgrupp: o.huvudgrupp,
        Antal: o.antal,
        Kommentar: o.kommentar,
        Datum: o.skapadTid,
        "Önskad av": o.skapadAv,
        Lag: o.lag || "",
      }));

      const lagRows = lagLager.map((r) => ({
        Lag: r.lag,
        Huvudgrupp: r.huvudgrupp,
        Produkt: r.produkt,
        Antal: r.antal,
        Utlämningsdatum: r.utlamningsdatum,
        Kommentar: r.kommentar,
      }));

      const historikRows = historik.map((h) => ({
        Tid: h.tid,
        Typ: h.typ,
        Produkt: h.produkt,
        Huvudgrupp: h.huvudgrupp,
        Lagerplats: h.lagerplats,
        Antal: h.antal,
        Användare: h.användare,
        Kommentar: h.kommentar,
      }));

      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(lagerRows), "Huvudlager");
      XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(lagRows), "Lagmaterial");
      XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(inkopRows), "MaterialÖnskemål");
      XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(onskemalRows), "Önskemål");
      XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(historikRows), "Historik");

      const filnamn = `Inventarie_export_${new Date().toISOString().slice(0, 10)}.xlsx`;
      downloadXlsx(wb, filnamn);
      showInfo(`Exporterade Excel: ${filnamn}`);
    } catch (e) {
      setFel(`Export misslyckades: ${e.message}`);
    }
  };

  /* ================= Render ================= */
  return (
    <ErrorBoundary>
      <div className="app">
        <style>{css}</style>

        {!currentUser ? (
          <LoginScreen users={users} onLogin={(u) => setCurrentUser(u)} />
        ) : (
          <>
            <header className="topbar">
              <div className="topbar__row">
                <div className="brand">
                  <div className="brand__logo">S</div>
                  <div className="brand__text">
                    <div className="title">Strömsbro IF – Inventarie</div>
                    <div className="subtitle">
                      Inloggad: <strong>{currentUser.name}</strong> • {currentUser.role}
                      {currentUser.role === "Ledare" ? ` • ${currentUser.lag || "Okänt"}` : ""}
                    </div>
                  </div>
                </div>

                <div className="seg">
                  <button className={`seg__btn ${vy === "lager" ? "seg__btn--active" : ""}`} onClick={() => setVy("lager")} type="button">
                    Lager
                  </button>
                  <button className={`seg__btn ${vy === "lagmaterial" ? "seg__btn--active" : ""}`} onClick={() => setVy("lagmaterial")} type="button">
                    Lagmaterial
                  </button>
                  <button className={`seg__btn ${vy === "onskemal" ? "seg__btn--active" : ""}`} onClick={() => setVy("onskemal")} type="button">
                    Önskemål
                  </button>

                  {canSeeInkop ? (
                    <button className={`seg__btn ${vy === "inkop" ? "seg__btn--active" : ""}`} onClick={() => setVy("inkop")} type="button">
                      Inköp
                    </button>
                  ) : null}

                  {canSeeInleverans ? (
                    <button className={`seg__btn ${vy === "inleverans" ? "seg__btn--active" : ""}`} onClick={() => setVy("inleverans")} type="button">
                      Inleverans
                    </button>
                  ) : null}

                  <button className={`seg__btn ${vy === "historik" ? "seg__btn--active" : ""}`} onClick={() => setVy("historik")} type="button">
                    Historik
                  </button>

                  {canEditUsers ? (
                    <button className={`seg__btn ${vy === "admin" ? "seg__btn--active" : ""}`} onClick={() => setVy("admin")} type="button">
                      Admin
                    </button>
                  ) : null}
                </div>
              </div>

              <div className="topbar__row topbar__row--actions">
                <div className="searchWrap">
                  <div className="searchBox">
                    <input className="search" value={sok} onChange={(e) => setSok(e.target.value)} placeholder="Sök…" inputMode="search" />
                    {sok ? (
                      <button className="clearSearch" type="button" onClick={() => setSok("")} aria-label="Rensa sök">
                        ✕
                      </button>
                    ) : null}
                  </div>
                </div>

                <div className="actions">
                  <PrimaryButton
                    onClick={() => {
                      setCurrentUser(null);
                      setVy("lager");
                      setSok("");
                    }}
                    tone="ghost"
                  >
                    🚪 Logga ut
                  </PrimaryButton>

                  {canImportExport ? (
                    <>
                      <input
                        ref={fileInputRef}
                        type="file"
                        accept=".xlsx,.xls"
                        className="hidden"
                        onChange={(e) => {
                          const f = e.target.files?.[0];
                          if (f) importeraExcel(f);
                        }}
                      />
                      <PrimaryButton onClick={() => fileInputRef.current?.click()} tone="ghost">
                        📥 Import
                      </PrimaryButton>
                      <PrimaryButton onClick={exporteraExcel} tone="ghost">
                        📤 Export
                      </PrimaryButton>
                    </>
                  ) : null}

                  {canUtlamna ? (
                    <PrimaryButton
                      onClick={() => {
                        setSelectMode((v) => !v);
                        clearSelected();
                      }}
                      tone={selectMode ? "primary" : "ghost"}
                      disabled={vy !== "lager"}
                    >
                      {selectMode ? "✅ Klar" : "🧺 Massläge"}
                    </PrimaryButton>
                  ) : null}
                </div>
              </div>

              {fel ? <div className="banner banner--error">❌ {fel}</div> : null}
              {info ? <div className="banner banner--ok">✅ {info}</div> : null}
            </header>

            <main className="content">
              {/* Lager */}
              {vy === "lager" ? (
                <>
                  {selectMode && canUtlamna ? (
                    <div className="massBar">
                      <div className="massBar__left">
                        <div className="massBar__title">Massläge</div>
                        <div className="massBar__sub">{selectedCount} markerade</div>
                      </div>
                      <div className="massBar__actions">
                        <PrimaryButton tone="ghost" onClick={selectAllVisible}>
                          Markera alla
                        </PrimaryButton>
                        <PrimaryButton tone="ghost" onClick={clearSelected} disabled={selectedCount === 0}>
                          Rensa
                        </PrimaryButton>
                        <PrimaryButton tone="danger" onClick={() => openBatch("uttag")} disabled={selectedCount === 0}>
                          ➖ Uttag
                        </PrimaryButton>
                        <PrimaryButton tone="ok" onClick={() => openBatch("in")} disabled={selectedCount === 0}>
                          ➕ In
                        </PrimaryButton>
                      </div>
                    </div>
                  ) : null}

                  <div className="grid">
                    {filtreradeProdukter.length === 0 ? <div className="empty">Inga produkter matchar sökningen.</div> : null}

                    {filtreradeProdukter.map((p) => {
                      const st = status(p);
                      const qty = getQty(p.id);
                      const selected = selectMode && selectedIds.has(p.id);

                      return (
                        <section
                          className={`card ${selected ? "card--selected" : ""}`}
                          key={p.id}
                          onClick={() => {
                            if (!selectMode) return;
                            toggleSelected(p.id);
                          }}
                        >
                          <div className="card__top">
                            <div className="card__title">{p.produkt}</div>
                            <div style={{ display: "flex", gap: 8, alignItems: "center" }}>
                              {selectMode && canUtlamna ? (
                                <label className="selectCtl" title="Markera" onClick={(e) => e.stopPropagation()}>
                                  <input type="checkbox" checked={selectedIds.has(p.id)} onChange={() => toggleSelected(p.id)} />
                                  <span>Markera</span>
                                </label>
                              ) : null}
                              <Pill tone={st.tone}>{st.text}</Pill>
                            </div>
                          </div>

                          <div className="meta">
                            <div className="meta__row">
                              <span className="meta__label">Grupp</span>
                              <span className="meta__value">{p.huvudgrupp || "—"}</span>
                            </div>
                            <div className="meta__row">
                              <span className="meta__label">Plats</span>
                              <span className="meta__value">{p.lagerplats || "—"}</span>
                            </div>
                          </div>

                          <div className="qtyRow">
                            <div className="qty">
                              <div className="qty__label">I lager</div>
                              <div className="qty__value">{toInt(p.antal, 0)} st</div>
                            </div>
                            <div className="miniMeta">
                              <div>
                                Min: <strong>{toInt(p.minAntal, 0)}</strong>
                              </div>
                              <div>
                                Beställ: <strong>{toInt(p.beställningspunkt, 0)}</strong>
                              </div>
                            </div>
                          </div>

                          <div className="qtyPick" onClick={(e) => e.stopPropagation()}>
                            <div className="qtyPick__label">Antal (st)</div>
                            <input className="qtyInput" type="number" min={1} inputMode="numeric" value={qty} onChange={(e) => setQty(p.id, e.target.value)} />
                            <div className="qtyChips">
                              <button className="chip" type="button" onClick={() => setQty(p.id, 1)}>
                                1
                              </button>
                              <button className="chip" type="button" onClick={() => setQty(p.id, 5)}>
                                5
                              </button>
                              <button className="chip" type="button" onClick={() => setQty(p.id, 10)}>
                                10
                              </button>
                            </div>
                          </div>

                          <div className="btnRow" onClick={(e) => e.stopPropagation()}>
                            <PrimaryButton onClick={() => openTx("uttag", p)} tone="danger" disabled={!canUtlamna || selectMode}>
                              ➖ Uttag
                            </PrimaryButton>
                            <PrimaryButton onClick={() => openTx("in", p)} tone="ok" disabled={!canEditHuvudlager || selectMode}>
                              ➕ In
                            </PrimaryButton>
                          </div>

                          {isLedare ? <div className="muted" style={{ marginTop: 8 }}>Ledare kan inte göra uttag/inleverans i huvudlager.</div> : null}
                        </section>
                      );
                    })}
                  </div>
                </>
              ) : null}

              {/* Lagmaterial */}
              {vy === "lagmaterial" ? (
                <>
                  <div className="summaryCard">
                    <div className="summaryTitle">Lagmaterial</div>
                    <div className="summarySub">
                      {isLedare ? `Du ser endast material för ${aktivtLagEff}.` : "Välj lag, se material och gör åtgärder vid behov."}
                    </div>
                  </div>

                  {!isLedare ? (
                    <section className="card">
                      <div className="formGrid">
                        <label className="field">
                          <span>Välj lag</span>
                          <select
                            value={aktivtLag}
                            onChange={(e) => {
                              const lag = e.target.value;
                              setAktivtLag(lag);
                              setFlytt((f) => ({ ...f, lag }));
                              setNyLagRad((r) => ({ ...r, lag }));
                            }}
                          >
                            {lagLista.map((l) => (
                              <option key={l} value={l}>
                                {l}
                              </option>
                            ))}
                          </select>
                        </label>
                      </div>
                    </section>
                  ) : null}

                  <div className="summaryCard" style={{ marginTop: 12 }}>
                    <div className="summaryTitle">{aktivtLagEff} – material</div>
                    <div className="summaryValue">{lagTotalAntal} st</div>
                    <div className="summarySub">{lagRaderVisning.length} rader</div>
                  </div>

                  {lagRaderVisning.length === 0 ? (
                    <div className="empty">Inga rader ännu.</div>
                  ) : (
                    <div className="grid">
                      {lagRaderVisning.map((r) => (
                        <section className="card" key={r.id}>
                          <div className="card__top">
                            <div className="card__title">{r.produkt}</div>
                            <Pill tone="neutral">{toInt(r.antal, 0)} st</Pill>
                          </div>

                          <div className="meta">
                            <div className="meta__row">
                              <span className="meta__label">Grupp</span>
                              <span className="meta__value">{r.huvudgrupp || "—"}</span>
                            </div>
                            <div className="meta__row">
                              <span className="meta__label">Utlämningsdatum</span>
                              <span className="meta__value">{r.utlamningsdatum || "—"}</span>
                            </div>
                            {r.kommentar ? (
                              <div className="meta__row">
                                <span className="meta__label">Kommentar</span>
                                <span className="meta__value">{r.kommentar}</span>
                              </div>
                            ) : null}
                          </div>

                          <div className="qtyPick" style={{ marginTop: 10 }}>
                            <div className="qtyPick__label">Ändra antal</div>
                            <input className="qtyInput" type="number" min={0} inputMode="numeric" value={toInt(r.antal, 0)} onChange={(e) => uppdateraLagRadAntal(r.id, e.target.value)} />
                          </div>

                          <div className="btnRow" style={{ marginTop: 10 }}>
                            <PrimaryButton tone="ok" onClick={() => openRetur(r)} disabled={!canUtlamna}>
                              ↩ Retur
                            </PrimaryButton>
                            <PrimaryButton tone="danger" onClick={() => taBortLagRad(r.id)} disabled={isLedare}>
                              🗑️ Ta bort
                            </PrimaryButton>
                          </div>
                        </section>
                      ))}
                    </div>
                  )}

                  {!isLedare ? (
                    <div className="btnRow" style={{ marginTop: 16 }}>
                      <PrimaryButton
                        tone="primary"
                        onClick={() => {
                          setShowUtlamning((v) => !v);
                          setShowAddLagRad(false);
                        }}
                      >
                        📤 Utlämna till lag
                      </PrimaryButton>

                      <PrimaryButton
                        tone="ghost"
                        onClick={() => {
                          setShowAddLagRad((v) => !v);
                          setShowUtlamning(false);
                        }}
                      >
                        ➕ Lägg till rad
                      </PrimaryButton>
                    </div>
                  ) : null}

                  {!isLedare && showUtlamning ? (
                    <section className="card" style={{ marginTop: 12 }}>
                      <div className="card__top">
                        <div className="card__title">📤 Utlämna till {aktivtLagEff}</div>
                        <Pill tone="neutral">Minskar huvudlager</Pill>
                      </div>

                      <div className="formGrid" style={{ marginTop: 10 }}>
                        <label className="field">
                          <span>Produkt (från huvudlager) *</span>
                          <select value={flytt.produktId} onChange={(e) => setFlytt((f) => ({ ...f, produktId: e.target.value }))}>
                            <option value="">— Välj —</option>
                            {produkter
                              .slice()
                              .sort((a, b) => {
                                const ah = (a.huvudgrupp || "").localeCompare(b.huvudgrupp || "");
                                if (ah !== 0) return ah;
                                return (a.produkt || "").localeCompare(b.produkt || "");
                              })
                              .map((p) => (
                                <option key={p.id} value={String(p.id)}>
                                  {p.produkt} ({toInt(p.antal, 0)} st) • {p.huvudgrupp || "—"}
                                </option>
                              ))}
                          </select>
                        </label>

                        <label className="field">
                          <span>Antal</span>
                          <input type="number" min={1} inputMode="numeric" value={flytt.antal} onChange={(e) => setFlytt((f) => ({ ...f, antal: Math.max(1, toInt(e.target.value, 1)) }))} />
                        </label>

                        <label className="field">
                          <span>Utlämningsdatum</span>
                          <input type="date" value={flytt.utlamningsdatum} onChange={(e) => setFlytt((f) => ({ ...f, utlamningsdatum: e.target.value }))} />
                        </label>

                        <label className="field">
                          <span>Kommentar</span>
                          <input value={flytt.kommentar} onChange={(e) => setFlytt((f) => ({ ...f, kommentar: e.target.value }))} placeholder="t.ex. match, cup, träning" list="recent-comments" />
                          <datalist id="recent-comments">
                            {recentComments.map((c) => (
                              <option key={c} value={c} />
                            ))}
                          </datalist>
                        </label>
                      </div>

                      <div className="btnRow" style={{ marginTop: 12 }}>
                        <PrimaryButton tone="primary" onClick={flyttaTillLag} disabled={!flytt.produktId}>
                          ✅ Utlämna
                        </PrimaryButton>

                        <PrimaryButton
                          tone="ghost"
                          onClick={() =>
                            setFlytt((f) => ({
                              ...f,
                              produktId: "",
                              antal: 1,
                              utlamningsdatum: "",
                              kommentar: "",
                            }))
                          }
                        >
                          ↺ Rensa
                        </PrimaryButton>
                      </div>

                      <div className="muted" style={{ marginTop: 8 }}>Tips: Antalet dras från huvudlagret och hamnar i lagets materiallista.</div>
                    </section>
                  ) : null}

                  {!isLedare && showAddLagRad ? (
                    <section className="card" style={{ marginTop: 12 }}>
                      <div className="card__top">
                        <div className="card__title">➕ Lägg till rad (manuellt)</div>
                        <Pill tone="neutral">Påverkar ej huvudlager</Pill>
                      </div>

                      <div className="formGrid" style={{ marginTop: 10 }}>
                        <label className="field">
                          <span>Lag</span>
                          <select value={nyLagRad.lag} onChange={(e) => setNyLagRad((r) => ({ ...r, lag: e.target.value }))}>
                            {lagLista.map((l) => (
                              <option key={l} value={l}>
                                {l}
                              </option>
                            ))}
                          </select>
                        </label>

                        <label className="field">
                          <span>Produkt *</span>
                          <input value={nyLagRad.produkt} onChange={(e) => setNyLagRad((r) => ({ ...r, produkt: e.target.value }))} placeholder="t.ex. Koner, Första hjälpen-kit…" />
                        </label>

                        <label className="field">
                          <span>Huvudgrupp</span>
                          <input value={nyLagRad.huvudgrupp} onChange={(e) => setNyLagRad((r) => ({ ...r, huvudgrupp: e.target.value }))} placeholder="t.ex. Koner, Medicin, Västar…" />
                        </label>

                        <label className="field">
                          <span>Antal</span>
                          <input type="number" min={0} inputMode="numeric" value={nyLagRad.antal} onChange={(e) => setNyLagRad((r) => ({ ...r, antal: Math.max(0, toInt(e.target.value, 0)) }))} />
                        </label>

                        <label className="field">
                          <span>Utlämningsdatum</span>
                          <input type="date" value={nyLagRad.utlamningsdatum} onChange={(e) => setNyLagRad((r) => ({ ...r, utlamningsdatum: e.target.value }))} />
                        </label>

                        <label className="field">
                          <span>Kommentar</span>
                          <input value={nyLagRad.kommentar} onChange={(e) => setNyLagRad((r) => ({ ...r, kommentar: e.target.value }))} placeholder="valfritt" list="recent-comments" />
                        </label>
                      </div>

                      <div className="btnRow" style={{ marginTop: 12 }}>
                        <PrimaryButton tone="primary" onClick={laggTillLagRad} disabled={!String(nyLagRad.produkt || "").trim()}>
                          ✅ Lägg till rad
                        </PrimaryButton>

                        <PrimaryButton
                          tone="ghost"
                          onClick={() =>
                            setNyLagRad((r) => ({
                              ...r,
                              huvudgrupp: "",
                              produkt: "",
                              antal: 1,
                              utlamningsdatum: "",
                              kommentar: "",
                            }))
                          }
                        >
                          ↺ Rensa
                        </PrimaryButton>
                      </div>

                      <div className="muted" style={{ marginTop: 8 }}>Detta skapar en rad i lagets material utan att ändra huvudlagret.</div>
                    </section>
                  ) : null}
                </>
              ) : null}

              {/* Önskemål */}
              {vy === "onskemal" ? (
                <>
                  <div className="summaryCard">
                    <div className="summaryTitle">Nytt inköpsönskemål</div>
                    <div className="summarySub">
                      {isLedare ? "Ledare kan lägga önskemål men kan inte flytta dem till inköp." : "Lägg in önskemål och flytta till inköp vid behov."}
                    </div>
                  </div>

                  <section className="card">
                    <div className="formGrid">
                      <label className="field">
                        <span>Produkt *</span>
                        <input value={nyttOnskemal.produkt} onChange={(e) => setNyttOnskemal((o) => ({ ...o, produkt: e.target.value }))} />
                      </label>
                      <label className="field">
                        <span>Huvudgrupp</span>
                        <input value={nyttOnskemal.huvudgrupp} onChange={(e) => setNyttOnskemal((o) => ({ ...o, huvudgrupp: e.target.value }))} />
                      </label>
                      <label className="field">
                        <span>Antal</span>
                        <input type="number" min={1} inputMode="numeric" value={nyttOnskemal.antal} onChange={(e) => setNyttOnskemal((o) => ({ ...o, antal: Math.max(1, toInt(e.target.value, 1)) }))} />
                      </label>
                      <label className="field">
                        <span>Kommentar</span>
                        <input value={nyttOnskemal.kommentar} onChange={(e) => setNyttOnskemal((o) => ({ ...o, kommentar: e.target.value }))} />
                      </label>
                    </div>
                    <div className="btnRow" style={{ marginTop: 12 }}>
                      <PrimaryButton tone="primary" onClick={skapaOnskemal}>
                        ➕ Lägg till önskemål
                      </PrimaryButton>
                      <div />
                    </div>
                  </section>

                  <div className="grid" style={{ marginTop: 12 }}>
                    {onskemal
                      .filter((o) => {
                        if (!isLedare) return true;
                        return o.lag === (currentUser?.lag ?? "") || o.skapadAv === currentUser?.name;
                      })
                      .filter((o) => {
                        const q = sok.trim().toLowerCase();
                        if (!q) return true;
                        return `${o.produkt} ${o.huvudgrupp} ${o.skapadAv} ${o.kommentar}`.toLowerCase().includes(q);
                      })
                      .map((o) => (
                        <section className="card" key={o.id}>
                          <div className="card__top">
                            <div className="card__title">{o.produkt}</div>
                            <Pill tone="neutral">{toInt(o.antal, 1)} st</Pill>
                          </div>
                          <div className="meta">
                            <div className="meta__row">
                              <span className="meta__label">Grupp</span>
                              <span className="meta__value">{o.huvudgrupp || "—"}</span>
                            </div>
                            <div className="meta__row">
                              <span className="meta__label">Skapad</span>
                              <span className="meta__value">
                                {o.skapadTid || "—"} • {o.skapadAv || "—"}
                              </span>
                            </div>
                            {o.lag ? (
                              <div className="meta__row">
                                <span className="meta__label">Lag</span>
                                <span className="meta__value">{o.lag}</span>
                              </div>
                            ) : null}
                            {o.kommentar ? (
                              <div className="meta__row">
                                <span className="meta__label">Kommentar</span>
                                <span className="meta__value">{o.kommentar}</span>
                              </div>
                            ) : null}
                          </div>

                          <div className="btnRow" style={{ marginTop: 10 }}>
                            <PrimaryButton tone="ok" onClick={() => laggOnskemalTillInkop(o)} disabled={!canMoveToInkop}>
                              ✅ Lägg till i inköp
                            </PrimaryButton>
                            <PrimaryButton tone="danger" onClick={() => setOnskemal((prev) => prev.filter((x) => x.id !== o.id))}>
                              🗑️ Ta bort
                            </PrimaryButton>
                          </div>

                          {!canMoveToInkop ? <div className="muted" style={{ marginTop: 8 }}>Ledare kan inte flytta önskemål till inköp.</div> : null}
                        </section>
                      ))}
                  </div>
                </>
              ) : null}

              {/* Inköp */}
              {vy === "inkop" && canSeeInkop ? (
                <>
                  <div className="summaryCard">
                    <div className="summaryTitle">Totalt att beställa</div>
                    <div className="summaryValue">{totalInkopAntal} st</div>
                    <div className="summarySub">{inkopAttVisa.length} produkter (ej beställda)</div>
                  </div>

                  <div className="grid">
                    {inkopAttVisa.length === 0 ? <div className="empty">Inget att beställa just nu.</div> : null}

                    {inkopAttVisa.map((r) => (
                      <section className="card" key={r.key}>
                        <div className="card__top">
                          <div className="card__title">{r.produkt}</div>
                          <Pill tone={r.manuell ? "neutral" : "ok"}>{r.manuell ? "Manuell" : "Auto"}</Pill>
                        </div>

                        <div className="meta">
                          <div className="meta__row">
                            <span className="meta__label">Grupp</span>
                            <span className="meta__value">{r.huvudgrupp || "—"}</span>
                          </div>
                          <div className="meta__row">
                            <span className="meta__label">Rekommenderat</span>
                            <span className="meta__value">{toInt(r.rekommenderat, 0)}</span>
                          </div>
                        </div>

                        <div className="qtyPick">
                          <div className="qtyPick__label">Inköpsantal</div>
                          <input className="qtyInput" type="number" min={0} inputMode="numeric" value={toInt(r.antal, 0)} onChange={(e) => sättInköpsantalManuellt(r.key, e.target.value)} />
                          <div className="qtyChips">
                            <button className="chip" type="button" onClick={() => återställRekommenderat(r.key)}>
                              ↺
                            </button>
                          </div>
                        </div>

                        <div className="btnRow" style={{ marginTop: 10 }}>
                          <PrimaryButton tone="ok" onClick={() => markeraSomBestalld(r.key)} disabled={r.bestalld}>
                            ✅ Beställt
                          </PrimaryButton>
                          <PrimaryButton tone="ghost" onClick={() => angraBestalld(r.key)} disabled={!r.bestalld}>
                            ↩ Ångra
                          </PrimaryButton>
                        </div>
                      </section>
                    ))}
                  </div>
                </>
              ) : null}

              {/* Inleverans */}
              {vy === "inleverans" && canSeeInleverans ? (
                <>
                  {(() => {
                    const q = sok.trim().toLowerCase();
                    const bestallda = inkopslista.filter((r) => r.bestalld);
                    const filtrerade = !q ? bestallda : bestallda.filter((r) => `${r.produkt} ${r.huvudgrupp}`.toLowerCase().includes(q));

                    const ejKompletta = filtrerade.filter((r) => !r.levererad);
                    const kompletta = filtrerade.filter((r) => r.levererad);

                    const totalKvar = ejKompletta.reduce((sum, r) => {
                      const ordered = Math.max(0, toInt(r.antal, 0));
                      const received = Math.max(0, toInt(r.mottagetAntal, 0));
                      return sum + Math.max(0, ordered - received);
                    }, 0);

                    return (
                      <>
                        <div className="summaryCard">
                          <div className="summaryTitle">Inleverans – beställda artiklar</div>
                          <div className="summaryValue">{totalKvar} st kvar att ta emot</div>
                          <div className="summarySub">{ejKompletta.length} rader ej kompletta • {kompletta.length} levererade</div>
                        </div>

                        <section className="card">
                          <div className="formGrid">
                            <label className="field">
                              <span>Lagerplats</span>
                              <select value={leveransLagerplats} onChange={(e) => setLeveransLagerplats(e.target.value)}>
                                {lagerplatsOptions.map((lp) => (
                                  <option key={lp} value={lp}>
                                    {lp}
                                  </option>
                                ))}
                              </select>
                            </label>

                            <label className="field">
                              <span>Kommentar (valfri, gemensam)</span>
                              <input value={leveransKommentar} onChange={(e) => setLeveransKommentar(e.target.value)} placeholder="t.ex. leverans från Stadium / faktura 123" list="recent-comments" />
                              <datalist id="recent-comments">
                                {recentComments.map((c) => (
                                  <option key={c} value={c} />
                                ))}
                              </datalist>
                            </label>
                          </div>
                          <div className="muted" style={{ marginTop: 8 }}>Du kan göra delleveranser. Varje mottagning ökar huvudlagret direkt.</div>
                        </section>

                        <div className="grid" style={{ marginTop: 12 }}>
                          {ejKompletta.length === 0 ? <div className="empty">Inga beställda rader kvar att leverera.</div> : null}

                          {ejKompletta.map((r) => {
                            const ordered = Math.max(0, toInt(r.antal, 0));
                            const received = Math.max(0, toInt(r.mottagetAntal, 0));
                            const kvar = Math.max(0, ordered - received);
                            const qtyNow = getLevQty(r.key, kvar > 0 ? kvar : 1);

                            return (
                              <section className="card" key={r.key}>
                                <div className="card__top">
                                  <div className="card__title">{r.produkt}</div>
                                  <Pill tone="neutral">Beställd</Pill>
                                </div>

                                <div className="meta">
                                  <div className="meta__row">
                                    <span className="meta__label">Grupp</span>
                                    <span className="meta__value">{r.huvudgrupp || "—"}</span>
                                  </div>
                                  <div className="meta__row">
                                    <span className="meta__label">Beställt</span>
                                    <span className="meta__value">{ordered} st</span>
                                  </div>
                                  <div className="meta__row">
                                    <span className="meta__label">Mottaget</span>
                                    <span className="meta__value">{received} st</span>
                                  </div>
                                  <div className="meta__row">
                                    <span className="meta__label">Kvar</span>
                                    <span className="meta__value">{kvar} st</span>
                                  </div>
                                  {r.bestalldTid ? (
                                    <div className="meta__row">
                                      <span className="meta__label">Beställd</span>
                                      <span className="meta__value">{r.bestalldTid} • {r.bestalldAv || "—"}</span>
                                    </div>
                                  ) : null}
                                </div>

                                <div className="qtyPick">
                                  <div className="qtyPick__label">Ta emot nu</div>
                                  <input className="qtyInput" type="number" min={1} max={Math.max(1, kvar)} inputMode="numeric" value={qtyNow} onChange={(e) => setLevQty(r.key, e.target.value)} />
                                  <div className="qtyChips">
                                    <button className="chip" type="button" onClick={() => setLevQty(r.key, 1)}>1</button>
                                    <button className="chip" type="button" onClick={() => setLevQty(r.key, Math.min(5, Math.max(1, kvar)))}>5</button>
                                    <button className="chip" type="button" onClick={() => setLevQty(r.key, Math.min(10, Math.max(1, kvar)))}>10</button>
                                    <button className="chip" type="button" onClick={() => setLevQty(r.key, Math.max(1, kvar))} title="Sätt till kvar">⇢</button>
                                  </div>
                                </div>

                                <div className="btnRow" style={{ marginTop: 10 }}>
                                  <PrimaryButton tone="ok" onClick={() => mottaLeverans(r.key, qtyNow)} disabled={kvar <= 0}>
                                    ✅ Mottagen
                                  </PrimaryButton>
                                  <PrimaryButton tone="ghost" onClick={() => setLevQty(r.key, Math.max(1, kvar))} disabled={kvar <= 0}>
                                    Sätt = kvar
                                  </PrimaryButton>
                                </div>

                                {Array.isArray(r.leveranser) && r.leveranser.length ? (
                                  <div className="muted" style={{ marginTop: 10 }}>
                                    Senaste: {r.leveranser[0].tid} • {r.leveranser[0].av} • +{r.leveranser[0].antal}
                                    {r.leveranser[0].lagerplats ? ` • ${r.leveranser[0].lagerplats}` : ""}
                                    {r.leveranser[0].kommentar ? ` • ${r.leveranser[0].kommentar}` : ""}
                                  </div>
                                ) : null}
                              </section>
                            );
                          })}
                        </div>

                        {kompletta.length ? (
                          <div style={{ marginTop: 14 }}>
                            <div className="summaryCard">
                              <div className="summaryTitle">Levererade (klara)</div>
                              <div className="summarySub">{kompletta.length} rader</div>
                            </div>

                            <div className="grid">
                              {kompletta.map((r) => (
                                <section className="card" key={r.key}>
                                  <div className="card__top">
                                    <div className="card__title">{r.produkt}</div>
                                    <Pill tone="ok">Levererad</Pill>
                                  </div>
                                  <div className="meta">
                                    <div className="meta__row">
                                      <span className="meta__label">Grupp</span>
                                      <span className="meta__value">{r.huvudgrupp || "—"}</span>
                                    </div>
                                    <div className="meta__row">
                                      <span className="meta__label">Beställt</span>
                                      <span className="meta__value">{toInt(r.antal, 0)} st</span>
                                    </div>
                                    <div className="meta__row">
                                      <span className="meta__label">Mottaget</span>
                                      <span className="meta__value">{toInt(r.mottagetAntal, 0)} st</span>
                                    </div>
                                    <div className="meta__row">
                                      <span className="meta__label">Klarmarkerad</span>
                                      <span className="meta__value">{r.levereradTid || "—"} • {r.levereradAv || "—"}</span>
                                    </div>
                                  </div>
                                </section>
                              ))}
                            </div>
                          </div>
                        ) : null}
                      </>
                    );
                  })()}
                </>
              ) : null}

              {/* Historik */}
              {vy === "historik" ? (
                <div className="history">
                  {historik.length === 0 ? <div className="empty">Ingen historik ännu.</div> : null}
                  {historik
                    .filter((h) => {
                      if (!isLedare) return true;
                      const lag = currentUser?.lag ?? "";
                      return String(h.kommentar || "").includes(`Till ${lag}`) || String(h.kommentar || "").includes(`Från ${lag}`);
                    })
                    .filter((h) => {
                      const q = sok.trim().toLowerCase();
                      if (!q) return true;
                      return `${h.typ} ${h.produkt} ${h.huvudgrupp} ${h.kommentar} ${h.användare}`.toLowerCase().includes(q);
                    })
                    .map((h, i) => (
                      <div className="historyRow" key={i}>
                        <div className="historyRow__left">
                          <div className="historyRow__title">
                            <strong>{h.typ}</strong> • {h.produkt}
                          </div>
                          <div className="historyRow__sub">
                            {h.tid} • {h.användare} • {h.huvudgrupp ? `${h.huvudgrupp} • ` : ""}
                            {h.lagerplats || ""}
                          </div>
                          <div className="historyRow__comment">📝 {h.kommentar}</div>
                        </div>
                        <div className={`historyRow__delta ${h.antal < 0 ? "neg" : "pos"}`}>
                          {h.antal > 0 ? "+" : ""}
                          {h.antal}
                        </div>
                      </div>
                    ))}
                </div>
              ) : null}

              {/* Admin */}
              {vy === "admin" && canEditUsers ? (
                <>
                  <div className="summaryCard">
                    <div className="summaryTitle">Admin</div>
                    <div className="summarySub">Skapa och hantera användare (lokalt) + PIN.</div>
                  </div>

                  <section className="card">
                    <div className="card__top">
                      <div className="card__title">➕ Lägg till användare</div>
                      <Pill tone="neutral">Endast Admin</Pill>
                    </div>

                    <div className="formGrid" style={{ marginTop: 10 }}>
                      <label className="field">
                        <span>Namn *</span>
                        <input value={userDraft.name} onChange={(e) => setUserDraft((d) => ({ ...d, name: e.target.value }))} />
                      </label>
                      <label className="field">
                        <span>Användarnamn *</span>
                        <input value={userDraft.username} onChange={(e) => setUserDraft((d) => ({ ...d, username: e.target.value }))} />
                      </label>
                      <label className="field">
                        <span>Roll</span>
                        <select value={userDraft.role} onChange={(e) => setUserDraft((d) => ({ ...d, role: e.target.value }))}>
                          <option value="Admin">Admin</option>
                          <option value="Materialansvarig">Materialansvarig</option>
                          <option value="Ledare">Ledare</option>
                        </select>
                      </label>

                      <label className="field">
                        <span>PIN (4–8 siffror)</span>
                        <input type="password" inputMode="numeric" value={userDraft.pin} onChange={(e) => setUserDraft((d) => ({ ...d, pin: e.target.value }))} />
                      </label>

                      {userDraft.role === "Ledare" ? (
                        <label className="field">
                          <span>Lag (för Ledare)</span>
                          <input value={userDraft.lag} onChange={(e) => setUserDraft((d) => ({ ...d, lag: e.target.value }))} />
                        </label>
                      ) : null}
                    </div>

                    <div className="btnRow" style={{ marginTop: 12 }}>
                      <PrimaryButton tone="primary" onClick={addUser}>
                        ✅ Skapa användare
                      </PrimaryButton>
                      <div />
                    </div>
                  </section>

                  <section className="card" style={{ marginTop: 12 }}>
                    <div className="card__top">
                      <div className="card__title">🔁 Återställ PIN</div>
                      <Pill tone="neutral">Admin</Pill>
                    </div>

                    <div className="formGrid" style={{ marginTop: 10 }}>
                      <label className="field">
                        <span>Användare</span>
                        <select value={pinEdit.userId} onChange={(e) => setPinEdit((p) => ({ ...p, userId: e.target.value }))}>
                          <option value="">— Välj —</option>
                          {users.map((u) => (
                            <option key={u.id} value={u.id}>
                              {u.name} ({u.role}
                              {u.role === "Ledare" ? ` • ${u.lag || "Okänt"}` : ""})
                            </option>
                          ))}
                        </select>
                      </label>

                      <label className="field">
                        <span>Ny PIN</span>
                        <input type="password" inputMode="numeric" value={pinEdit.newPin} onChange={(e) => setPinEdit((p) => ({ ...p, newPin: e.target.value }))} />
                      </label>
                    </div>

                    <div className="btnRow" style={{ marginTop: 12 }}>
                      <PrimaryButton tone="primary" onClick={setUserPin}>
                        ✅ Spara PIN
                      </PrimaryButton>
                      <div />
                    </div>
                  </section>

                  <div className="grid" style={{ marginTop: 12 }}>
                    {users.map((u) => (
                      <section className="card" key={u.id}>
                        <div className="card__top">
                          <div className="card__title">{u.name}</div>
                          <Pill tone="neutral">
                            {u.role}
                            {u.role === "Ledare" ? ` • ${u.lag || "Okänt"}` : ""}
                          </Pill>
                        </div>

                        <div className="meta">
                          <div className="meta__row">
                            <span className="meta__label">Användarnamn</span>
                            <span className="meta__value">{u.username}</span>
                          </div>
                          <div className="meta__row">
                            <span className="meta__label">PIN satt</span>
                            <span className="meta__value">{u.pinHash ? "Ja" : "Nej"}</span>
                          </div>
                        </div>

                        <div className="btnRow" style={{ marginTop: 10 }}>
                          <PrimaryButton
                            tone="ghost"
                            onClick={() => setCurrentUser({ id: u.id, name: u.name, username: u.username, role: u.role, lag: u.lag ?? null })}
                          >
                            Logga in som
                          </PrimaryButton>
                          <PrimaryButton tone="danger" onClick={() => deleteUser(u.id)} disabled={currentUser.id === u.id}>
                            Ta bort
                          </PrimaryButton>
                        </div>
                      </section>
                    ))}
                  </div>
                </>
              ) : null}
            </main>

            {/* ===== Modal: Enstaka uttag/in ===== */}
            <Modal
              open={txOpen}
              title={txMode === "uttag" ? "➖ Uttag från lager" : "➕ Inleverans till lager"}
              onClose={closeTx}
              footer={
                <>
                  <PrimaryButton tone="ghost" onClick={closeTx}>
                    Avbryt
                  </PrimaryButton>
                  <PrimaryButton
                    tone={txMode === "uttag" ? "danger" : "ok"}
                    onClick={() => {
                      const p = produkter.find((x) => x.id === txProductId);
                      if (!p) return;
                      const qty = Math.max(1, toInt(txQty, 1));
                      const delta = txMode === "uttag" ? -qty : +qty;
                      andringDirekt(p, delta, txComment);
                      closeTx();
                      showInfo(txMode === "uttag" ? "Uttag registrerat." : "Inleverans registrerad.");
                    }}
                    disabled={txMode === "uttag" ? !canUtlamna : !canEditHuvudlager}
                  >
                    Bekräfta
                  </PrimaryButton>
                </>
              }
            >
              {(() => {
                const p = produkter.find((x) => x.id === txProductId);
                if (!p) return <div className="empty">Ingen produkt vald.</div>;
                return (
                  <div className="formGrid">
                    <div style={{ display: "grid", gap: 6 }}>
                      <div style={{ fontWeight: 950, fontSize: 16 }}>{p.produkt}</div>
                      <div style={{ color: "var(--muted)", fontSize: 12 }}>
                        {p.huvudgrupp || "—"} • {p.lagerplats || "—"} • I lager: <strong>{toInt(p.antal, 0)}</strong> st
                      </div>
                    </div>

                    <label className="field">
                      <span>Antal</span>
                      <input className="qtyInput" type="number" min={1} inputMode="numeric" value={txQty} onChange={(e) => setTxQty(e.target.value)} />
                    </label>

                    <label className="field">
                      <span>Kommentar</span>
                      <input value={txComment} onChange={(e) => setTxComment(e.target.value)} placeholder="t.ex. P12 träning" list="recent-comments" />
                      <datalist id="recent-comments">
                        {recentComments.map((c) => (
                          <option key={c} value={c} />
                        ))}
                      </datalist>
                    </label>
                  </div>
                );
              })()}
            </Modal>

            {/* ===== Modal: Mass ===== */}
            <Modal
              open={batchOpen}
              title={batchMode === "uttag" ? "➖ Massuttag" : "➕ Massinleverans"}
              onClose={closeBatch}
              footer={
                <>
                  <PrimaryButton tone="ghost" onClick={closeBatch}>
                    Avbryt
                  </PrimaryButton>
                  <PrimaryButton tone={batchMode === "uttag" ? "danger" : "ok"} onClick={applyBatch}>
                    Bekräfta
                  </PrimaryButton>
                </>
              }
            >
              <div className="formGrid">
                <div style={{ display: "grid", gap: 6 }}>
                  <div style={{ fontWeight: 950, fontSize: 16 }}>{selectedCount} produkter markerade</div>
                  <div style={{ color: "var(--muted)", fontSize: 12 }}>
                    {batchUsePerItemQty ? "Använder respektive produkts Antal (st) i kortet." : `Använder samma antal för alla: ${Math.max(1, toInt(batchQty, 1))} st`}
                  </div>
                </div>

                <label className="field">
                  <span>Antal</span>
                  <div style={{ display: "grid", gap: 8 }}>
                    <label style={{ display: "flex", gap: 10, alignItems: "center", color: "var(--text)", fontWeight: 900 }}>
                      <input type="checkbox" checked={batchUsePerItemQty} onChange={(e) => setBatchUsePerItemQty(e.target.checked)} />
                      <span>Använd antal per produkt</span>
                    </label>

                    {!batchUsePerItemQty ? <input className="qtyInput" type="number" min={1} inputMode="numeric" value={batchQty} onChange={(e) => setBatchQty(e.target.value)} /> : null}
                  </div>
                </label>

                <label className="field">
                  <span>Kommentar (gemensam)</span>
                  <input value={batchComment} onChange={(e) => setBatchComment(e.target.value)} placeholder="t.ex. P12 träning" list="recent-comments" />
                </label>
              </div>
            </Modal>

            {/* ===== Modal: Retur ===== */}
            <Modal
              open={returOpen}
              title="↩ Retur till huvudlager"
              onClose={closeRetur}
              footer={
                <>
                  <PrimaryButton tone="ghost" onClick={closeRetur}>
                    Avbryt
                  </PrimaryButton>
                  <PrimaryButton
                    tone="ok"
                    onClick={() => {
                      if (!returRad) return;
                      returTillHuvudlager(returRad, returQty);
                      closeRetur();
                    }}
                    disabled={!canUtlamna}
                  >
                    Bekräfta retur
                  </PrimaryButton>
                </>
              }
            >
              {!returRad ? (
                <div className="empty">Ingen rad vald.</div>
              ) : (
                <div className="formGrid">
                  <div style={{ display: "grid", gap: 6 }}>
                    <div style={{ fontWeight: 950, fontSize: 16 }}>{returRad.produkt}</div>
                    <div style={{ color: "var(--muted)", fontSize: 12 }}>
                      Lag: <strong>{returRad.lag || "Okänt"}</strong> • Tillgängligt: <strong>{toInt(returRad.antal, 0)}</strong> st
                    </div>
                  </div>

                  <label className="field">
                    <span>Antal att returnera</span>
                    <input type="number" min={1} max={Math.max(1, toInt(returRad.antal, 1))} inputMode="numeric" value={returQty} onChange={(e) => setReturQty(e.target.value)} />
                  </label>
                </div>
              )}
            </Modal>
          </>
        )}
      </div>
    </ErrorBoundary>
  );
}

/* ================= CSS ================= */
const css = `
:root{
  --primary:#0b3a82;
  --primary-strong:#1e5bbf;
  --ok:#1e5bbf;
  --warn:#f59e0b;
  --danger:#ef4444;

  --bg:#071426;
  --muted:#9db3d8;
  --text:#e6ecf7;
  --line:rgba(157,179,216,.18);

  --shadow: 0 8px 30px rgba(0,0,0,.28);
  --radius:14px;
}
*{ box-sizing:border-box; }
html,body{ height:100%; }
body{ margin:0; background: linear-gradient(180deg, #050b14, var(--bg)); color:var(--text); font-family: ui-sans-serif, system-ui, -apple-system, Segoe UI, Roboto, Arial; }
.hidden{ display:none; }
.app{ min-height:100vh; }

.loginWrap{ max-width: 520px; margin: 70px auto; padding: 12px; }

.topbar{
  position: sticky; top: 0; z-index: 10;
  padding: 12px;
  background: rgba(7,20,38,.90);
  backdrop-filter: blur(10px);
  border-bottom: 1px solid var(--line);
}
.topbar__row{ display:flex; gap: 10px; align-items:center; justify-content: space-between; flex-wrap: wrap; }
.topbar__row--actions{ margin-top: 10px; }

.brand{ display:flex; align-items:center; gap:10px; min-width: 260px; }
.brand__logo{
  width: 40px; height: 40px; border-radius: 12px;
  display:flex; align-items:center; justify-content:center;
  background: rgba(255,255,255,.06);
  border: 1px solid var(--line);
  font-weight: 950;
}
.brand__text{ display:flex; flex-direction:column; gap:2px; }
.title{ font-size: 18px; font-weight: 900; letter-spacing: .2px; }
.subtitle{ font-size: 12px; color: var(--muted); }

.seg{ display:flex; gap: 6px; background: rgba(157,179,216,.10); padding: 6px; border-radius: 999px; }
.seg__btn{
  min-height: 40px; padding: 0 14px; border: 0; border-radius: 999px;
  background: transparent; color: var(--text); font-weight: 800; cursor: pointer;
}
.seg__btn--active{ background: rgba(11,58,130,.35); outline: 1px solid rgba(11,58,130,.55); }

.searchWrap{ flex: 1 1 240px; }
.searchBox{ position: relative; width:100%; }
.search{
  width: 100%; min-height: 44px; padding: 0 44px 0 12px;
  border-radius: 12px; border: 1px solid var(--line);
  background: rgba(15,23,42,.55); color: var(--text); outline: none;
}
.search::placeholder{ color: rgba(157,179,216,.85); }
.clearSearch{
  position:absolute; right:8px; top:50%; transform: translateY(-50%);
  min-width: 36px; min-height: 36px; border-radius: 999px;
  border: 1px solid var(--line); background: rgba(157,179,216,.18);
  color: var(--text); font-weight: 900; cursor:pointer;
}

.actions{ display:flex; gap: 8px; flex-wrap: wrap; }

.btn{
  min-height: 44px; padding: 0 14px; border-radius: 12px; border: 1px solid var(--line);
  background: rgba(15,23,42,.55); color: var(--text);
  font-weight: 900; cursor:pointer; display:inline-flex; align-items:center; justify-content:center;
}
.btn:disabled{ opacity: .55; cursor: not-allowed; }
.btn--primary{ background: var(--primary); border-color: var(--primary); color:#fff; }
.btn--ghost{ background: rgba(15,23,42,.25); }
.btn--ok{ background: var(--primary-strong); border-color: var(--primary-strong); color:#fff; }
.btn--danger{ background: rgba(239,68,68,.95); border-color: rgba(239,68,68,.95); color:#160707; }

.banner{ margin-top: 10px; border-radius: 12px; padding: 10px 12px; border: 1px solid var(--line); }
.banner--error{ background: rgba(239,68,68,.12); }
.banner--ok{ background: rgba(30,91,191,.18); }

.content{
  padding: 14px 12px 26px;
  max-width: 1040px; margin: 0 auto;
}
.grid{ display:grid; grid-template-columns: 1fr; gap: 12px; }

.card{ background: rgba(15,23,42,.70); border: 1px solid var(--line); border-radius: var(--radius); padding: 12px; box-shadow: var(--shadow); }
.card__top{ display:flex; align-items:flex-start; justify-content: space-between; gap: 10px; }
.card__title{ font-size: 16px; font-weight: 950; line-height: 1.2; }

.card--selected{
  outline: 2px solid rgba(30,91,191,.75);
  box-shadow: 0 0 0 4px rgba(30,91,191,.18), var(--shadow);
}

.pill{ display:inline-flex; align-items:center; justify-content:center; padding: 6px 10px; border-radius: 999px; font-size: 12px; font-weight: 950; border: 1px solid var(--line); }
.pill--ok{ background: rgba(30,91,191,.18); border-color: rgba(30,91,191,.35); }
.pill--warn{ background: rgba(245,158,11,.16); border-color: rgba(245,158,11,.35); }
.pill--danger{ background: rgba(239,68,68,.16); border-color: rgba(239,68,68,.35); }
.pill--neutral{ background: rgba(157,179,216,.12); }

.selectCtl{
  display:flex; gap: 8px; align-items:center;
  color: var(--muted); font-size: 12px; font-weight: 900;
}
.selectCtl input{ width: 18px; height: 18px; accent-color: var(--primary-strong); }

.meta{ margin-top: 10px; display:grid; gap: 6px; }
.meta__row{ display:flex; justify-content: space-between; gap: 10px; color: var(--muted); font-size: 13px; }
.meta__value{ color: var(--text); font-weight: 800; }

.qtyRow{ margin-top: 12px; display:flex; justify-content: space-between; gap: 12px; align-items:flex-end; }
.qty__label{ font-size: 12px; color: var(--muted); }
.qty__value{ font-size: 22px; font-weight: 950; }
.miniMeta{ font-size: 12px; color: var(--muted); display:grid; gap: 4px; text-align:right; }

.qtyPick{ margin-top: 12px; display:flex; gap: 10px; align-items:center; flex-wrap: wrap; }
.qtyPick__label{ font-size: 12px; color: var(--muted); font-weight: 900; }
.qtyInput{ min-height: 44px; width: 110px; padding: 0 12px; border-radius: 12px; border: 1px solid var(--line); background: rgba(2,6,23,.35); color: var(--text); outline: none; font-weight: 950; }
.qtyChips{ display:flex; gap: 8px; }
.chip{ min-height: 44px; min-width: 44px; padding: 0 12px; border-radius: 999px; border: 1px solid var(--line); background: rgba(15,23,42,.35); color: var(--text); font-weight: 950; cursor: pointer; }

.btnRow{ margin-top: 12px; display:grid; grid-template-columns: 1fr 1fr; gap: 10px; }

.iconBtn{ min-height: 44px; min-width: 44px; padding: 0 12px; border-radius: 12px; border: 1px solid var(--line); background: rgba(15,23,42,.35); color: var(--text); font-weight: 950; cursor:pointer; }
.iconBtn:disabled{ opacity:.55; cursor:not-allowed; }

.empty{ color: var(--muted); padding: 18px 10px; text-align:center; }
.muted{ color: var(--muted); }

.history{ display:grid; gap: 10px; }
.historyRow{ display:flex; justify-content: space-between; gap: 12px; padding: 12px; border-radius: var(--radius); border: 1px solid var(--line); background: rgba(15,23,42,.55); }
.historyRow__title{ font-size: 14px; font-weight: 950; }
.historyRow__sub{ font-size: 12px; color: var(--muted); margin-top: 2px; }
.historyRow__comment{ font-size: 12px; color: var(--text); opacity: .9; margin-top: 6px; }
.historyRow__delta{ min-width: 64px; text-align:right; font-weight: 950; font-size: 18px; }
.historyRow__delta.pos{ color: #86efac; }
.historyRow__delta.neg{ color: #fca5a5; }

.summaryCard{ border: 1px solid var(--line); background: rgba(11,58,130,.18); border-radius: 14px; padding: 12px; margin-bottom: 12px; }
.summaryTitle{ color: var(--muted); font-weight: 950; font-size: 12px; }
.summaryValue{ font-weight: 950; font-size: 26px; margin-top: 4px; }
.summarySub{ color: var(--muted); font-size: 12px; margin-top: 4px; }

.modalOverlay{ position: fixed; inset: 0; background: rgba(0,0,0,.55); display:flex; align-items:flex-end; justify-content:center; padding: 10px; z-index: 50; }
.modalSheet{ width: 100%; max-width: 720px; max-height: 90vh; background: rgba(15,23,42,.98); border: 1px solid var(--line); border-radius: 18px; box-shadow: var(--shadow); overflow:hidden; }
.modalHeader{ display:flex; align-items:center; justify-content: space-between; padding: 12px; border-bottom: 1px solid var(--line); }
.modalTitle{ font-weight: 950; font-size: 16px; }
.modalBody{ padding: 12px; overflow:auto; }
.modalFooter{ padding: 12px; display:flex; justify-content:flex-end; gap: 10px; border-top: 1px solid var(--line); }

.formGrid{ display:grid; grid-template-columns: 1fr; gap: 10px; }
.field{ display:grid; gap: 6px; font-size: 13px; color: var(--muted); }
.field input, .field select{ min-height: 44px; padding: 0 12px; border-radius: 12px; border: 1px solid var(--line); background: rgba(2,6,23,.35); color: var(--text); outline: none; }

.massBar{
  display:flex; gap: 10px; align-items:center; justify-content: space-between;
  padding: 12px; border: 1px solid var(--line);
  background: rgba(157,179,216,.10);
  border-radius: 14px; margin-bottom: 12px;
}
.massBar__title{ font-weight: 950; }
.massBar__sub{ color: var(--muted); font-size: 12px; margin-top: 2px; }
.massBar__left{ display:grid; gap: 2px; }
.massBar__actions{ display:flex; gap: 8px; flex-wrap: wrap; justify-content:flex-end; }

@media (min-width: 760px){
  .grid{ grid-template-columns: 1fr 1fr; }
  .modalOverlay{ align-items:center; }
}
`;
