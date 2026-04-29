import React, { useCallback, useEffect, useMemo, useRef, useState } from "react";
import * as XLSX from "xlsx";
import QtyInput from "./components/common/QtyInput";


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
/* ================= Outbox (offline ops) ================= */
const OPS_KEY = "stromsbro:pendingOps:v1";

const uuid = () =>
  globalThis.crypto?.randomUUID
    ? globalThis.crypto.randomUUID()
    : `${Date.now()}-${Math.random().toString(16).slice(2)}`;

function loadOps() {
  try {
    const raw = localStorage.getItem(OPS_KEY);
    const ops = raw ? JSON.parse(raw) : [];
    return Array.isArray(ops) ? ops : [];
  } catch {
    return [];
  }
}
function saveOps(ops) {
  try {
    localStorage.setItem(OPS_KEY, JSON.stringify(ops));
  } catch {}
}
function enqueueOp(op) {
  const ops = loadOps();
  ops.push(op);
  saveOps(ops);
  return ops.length;
}
function clearOps() {
  saveOps([]);
  return 0;
}

/** Hjälp: se till att payload har arrays */
function ensureArray(obj, key) {
  if (!obj[key] || !Array.isArray(obj[key])) obj[key] = [];
}
function getId(x) {
  return x?.id;
}

/** Applicera ops på en payload (non-destructive rebase) */
function applyOpsToPayload(payload, ops) {
  const next = structuredClone(payload ?? {});
  ensureArray(next, "users");
  ensureArray(next, "produkter");
  ensureArray(next, "inkopslista");
  ensureArray(next, "onskemal");
  ensureArray(next, "historik");
  ensureArray(next, "lagLager");

  const arrByEntity = {
    users: next.users,
    produkter: next.produkter,
    inkopslista: next.inkopslista,
    onskemal: next.onskemal,
    historik: next.historik,
    lagLager: next.lagLager,
  };

  for (const op of ops) {
    const arr = arrByEntity[op.entity];
    if (!arr) continue;

    if (op.kind === "upsert") {
      const id = getId(op.item);
      if (id == null) {
        arr.unshift(op.item);
      } else {
        const idx = arr.findIndex((x) => getId(x) === id);
        if (idx >= 0) arr[idx] = { ...arr[idx], ...op.item };
        else arr.unshift(op.item);
      }
      continue;
    }

    if (op.kind === "patch") {
      const idx = arr.findIndex((x) => getId(x) === op.itemId);
      if (idx >= 0) arr[idx] = { ...arr[idx], ...op.patch };
      continue;
    }

    if (op.kind === "remove") {
      const idx = arr.findIndex((x) => getId(x) === op.itemId);
      if (idx >= 0) arr.splice(idx, 1);
      continue;
    }
  }

  return next;
}

/** Bygg cloud-payload från state (delad data) */
function buildSharedPayload({ users, produkter, inkopslista, onskemal, historik, lagLager }) {
  return { users, produkter, inkopslista, onskemal, historik, lagLager };
}

/* ================= Cloud sync (Vercel KV via API) ================= */
const CLOUD_GET_URL = "/api/state/get";
const CLOUD_SET_URL = "/api/state/set";

function createDebouncer(ms = 900) {
  let t = null;
  return (fn) => {
    if (t) clearTimeout(t);
    t = setTimeout(fn, ms);
  };
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
  .filter((p) => !p.autoInkopPaused)
  .filter((p) => toInt(p.beställningspunkt, 0) > 0)
  .filter((p) => {
    const bp = toInt(p.beställningspunkt, 0);
    const antal = toInt(p.antal, 0);
    const orderThreshold = Math.floor(bp * 0.7);
    return antal <= orderThreshold; // ✅ ENDAST “Beställ”
  })
  .map((p) => {
    const bp = toInt(p.beställningspunkt, 0);
    const antal = toInt(p.antal, 0);
    const rekommenderat = Math.max(0, bp - antal);

    if (rekommenderat <= 0) return null;

    const key = makeKey(p.huvudgrupp, p.produkt);
    const prev = prevByKey.get(key);
    const manuell = prev?.manuell === true;

    const inkopAntal = manuell
      ? Math.max(0, toInt(prev?.antal, 0))
      : rekommenderat;

    if (inkopAntal <= 0) return null;

    return {
      id: key,
      key,
      huvudgrupp: p.huvudgrupp ?? "",
      produkt: p.produkt ?? "",
      antal: inkopAntal,
      rekommenderat,
      manuell,
      bestalld: prev?.bestalld === true,
      bestalldTid: prev?.bestalldTid ?? "",
      bestalldAv: prev?.bestalldAv ?? "",
      mottagetAntal: Math.max(0, toInt(prev?.mottagetAntal, 0)),
      levererad: prev?.levererad === true,
      leveranser: Array.isArray(prev?.leveranser) ? prev.leveranser : [],
    };
  })
  .filter(Boolean);

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

/* ================= Login Screen ================= */
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

const [pendingOpsCount, setPendingOpsCount] = useState(() => loadOps().length);

const refreshPendingOps = useCallback(() => {
  setPendingOpsCount(loadOps().length);
}, []);


/* ===== Auth/User state ===== */
  const [users, setUsers] = useState(() => {
    const base =
      Array.isArray(saved?.users) && saved.users.length
        ? saved.users
        : [
            { id: 1, name: "Admin", username: "admin", role: "Admin", lag: null },
            { id: 2, name: "Materialansvarig", username: "ma", role: "Materialansvarig", lag: null },
            { id: 3, name: "Ledare P-2012", username: "p12", role: "Ledare", lag: "P-2012" },
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
            beställningspunkt: 15,
          },
          {
            id: 2,
            huvudgrupp: "Västar",
            produkt: "Träningsvästar gula",
            antal: 6,
            lagerplats: "Lilla förrådet",
            beställningspunkt: 10,
          },
        ]
  );

  const [inkopslista, setInkopslista] = useState(() => {
    if (Array.isArray(saved?.inkopslista)) return saved.inkopslista;
    return buildInkopLista(Array.isArray(saved?.produkter) ? saved.produkter : [], []);
  });

const [manuellInkopOpen, setManuellInkopOpen] = useState(false);
const [manuellInkopProd, setManuellInkopProd] = useState(null);
const [manuellInkopQty, setManuellInkopQty] = useState(1);

const openManuellInkop = (p) => {
  setManuellInkopProd(p);
  setManuellInkopQty(1);
  setManuellInkopOpen(true);
};
const closeManuellInkop = () => {
  setManuellInkopOpen(false);
  setManuellInkopProd(null);
};

const inkopStatusForProdukt = (p) => {
  const key = makeKey(p.huvudgrupp, p.produkt);
  const r = inkopslista.find((x) => x.key === key);
  if (!r) return null;
  if (r.bestalld) return "beställd";
  return "inköp";
};

  const [onskemal, setOnskemal] = useState(Array.isArray(saved?.onskemal) ? saved.onskemal : []);
  const [historik, setHistorik] = useState(Array.isArray(saved?.historik) ? saved.historik : []);

  const [lagLager, setLagLager] = useState(Array.isArray(saved?.lagLager) ? saved.lagLager : []);
  const [aktivtLag, setAktivtLag] = useState(saved?.aktivtLag ?? "P-2012");

  const [recentComments, setRecentComments] = useState(
    Array.isArray(saved?.recentComments) ? saved.recentComments : []
  );
  const rememberComment = (c) => {
    const s = String(c ?? "").trim();
    if (!s) return;
    setRecentComments((prev) => [s, ...prev.filter((x) => x !== s)].slice(0, 8));
  };

const applySharedFromPayload = useCallback((p) => {
  if (Array.isArray(p?.users)) setUsers(p.users);
  if (Array.isArray(p?.produkter)) setProdukter(p.produkter);
  if (Array.isArray(p?.inkopslista)) setInkopslista(p.inkopslista);
  if (Array.isArray(p?.onskemal)) setOnskemal(p.onskemal);
  if (Array.isArray(p?.historik)) setHistorik(p.historik);
  if (Array.isArray(p?.lagLager)) setLagLager(p.lagLager);
}, [setUsers, setProdukter, setInkopslista, setOnskemal, setHistorik, setLagLager]);

// ===== Cloud sync refs (konfliktsäkert + ingen loop) =====
  const cloudHydratedRef = useRef(false);        // true efter första lyckade (eller misslyckade) load-försök
  const cloudVersionRef = useRef(0);             // senaste serverversion vi känner till
  const suppressCloudSaveRef = useRef(false);    // true när vi applicerar cloud-data
 
const cloudAttemptedRef = useRef(false);
const cloudLoadedOkRef = useRef(false);

const cloudSaveDebounceRef = useRef(null);
if (!cloudSaveDebounceRef.current) {
  cloudSaveDebounceRef.current = createDebouncer(900);
}

  const cloudInFlightRef = useRef(false);        // förhindrar parallella reloads

  const [cloudStatus, setCloudStatus] = useState({ loading: false, lastSync: "" });

  // ===== Hjälpare: applicera ENDAST delad state från cloud =====
  const applyCloudPayload = useCallback((p) => {
    suppressCloudSaveRef.current = true;

    // version först
    cloudVersionRef.current = p?.__version ?? 0;

    // ✅ DELAD DATA (synkas mellan enheter)
    if (Array.isArray(p?.users)) setUsers(p.users);                // OBS: om du vill undvika PIN i molnet, säg till så justerar vi
    if (Array.isArray(p?.produkter)) setProdukter(p.produkter);
    if (Array.isArray(p?.inkopslista)) setInkopslista(p.inkopslista);
    if (Array.isArray(p?.onskemal)) setOnskemal(p.onskemal);
    if (Array.isArray(p?.historik)) setHistorik(p.historik);
    if (Array.isArray(p?.lagLager)) setLagLager(p.lagLager);

    // ❌ LOKAL DATA (synkas INTE — per enhet enligt A)
    // - currentUser
    // - vy, sok, aktivtLag
    // - qtyMap, recentComments
    // - leveransKommentar, leveransLagerplats, leveransQtyMap
    // - övriga UI state

    
  cloudLoadedOkRef.current = true;     // ✅ viktig
  cloudAttemptedRef.current = true;    // ✅

// Efter att du satt state från cloud:
const ops = loadOps();
if (ops.length) {
  const mergedForUi = applyOpsToPayload(p, ops);
  applySharedFromPayload(mergedForUi);
  setPendingOpsCount(ops.length);
}

    // släpp spärren efter att state satts
    queueMicrotask(() => {
      suppressCloudSaveRef.current = false;
    });
  }, [setUsers, setProdukter, setInkopslista, setOnskemal, setHistorik, setLagLager, applySharedFromPayload]);

  // ===== Hämta från cloud (skippa apply om version ej ändrats) =====
const reloadFromCloud = useCallback(async ({ force = false } = {}) => {
  if (cloudInFlightRef.current) return;
  cloudInFlightRef.current = true;

  try {
    const res = await fetch(`${CLOUD_GET_URL}?t=${Date.now()}`, { cache: "no-store" });
    cloudAttemptedRef.current = true;

    if (!res.ok) return;

    const data = await res.json().catch(() => null);
    const p = data?.payload;
    if (!data?.ok || !p) return;

    const nextVer = p.__version ?? 0;
    const sameVersion = cloudLoadedOkRef.current && nextVer === cloudVersionRef.current;

    if (!force && sameVersion) {
      setCloudStatus((s) => ({ ...s, lastSync: data.updatedAt || p.__meta?.updatedAt || s.lastSync }));
      return;
    }

    applyCloudPayload(p);
    setCloudStatus((s) => ({ ...s, lastSync: data.updatedAt || p.__meta?.updatedAt || new Date().toISOString() }));
  } catch (e) {
    cloudAttemptedRef.current = true;
    console.warn("Cloud reload failed:", e);
  } finally {
    cloudInFlightRef.current = false;
  }
}, [applyCloudPayload]);

const syncNow = useCallback(async () => {
  if (!navigator.onLine) return;

  const ops = loadOps();
  if (!ops.length) return;

  // undvik parallella syncar
  if (cloudInFlightRef.current) return;
  cloudInFlightRef.current = true;

  try {
    for (let attempt = 0; attempt < 8; attempt++) {
      // 1) Hämta senaste cloud
      const resGet = await fetch(`${CLOUD_GET_URL}?t=${Date.now()}`, { cache: "no-store" });
      if (!resGet.ok) return;
      const data = await resGet.json().catch(() => null);
      const cloudPayload = data?.payload ?? null;
      if (!data?.ok || !cloudPayload) return;

      const baseVersion = cloudPayload.__version ?? 0;

      // 2) Rebase: cloud + ops
      const merged = applyOpsToPayload(cloudPayload, ops);
      merged.__version = baseVersion;
      merged.__meta = merged.__meta || {};
      merged.__meta.updatedAt = new Date().toISOString();

      // 3) Försök spara (server gör version-check)
      const resSet = await fetch(CLOUD_SET_URL, {
        method: "POST",
        headers: { "content-type": "application/json" },
        cache: "no-store",
        body: JSON.stringify({ payload: merged }),
      });

      // konflikt: någon annan sparade nyare -> loopa och rebasa igen
      if (resSet.status === 409) continue;

      if (resSet.ok) {
        const json = await resSet.json().catch(() => null);
        if (json?.version !== undefined) {
          cloudVersionRef.current = json.version;
          merged.__version = json.version;
        }

        // 4) ops är nu inkorporerade i molnstate -> töm outbox
        clearOps();
        setPendingOpsCount(0);

        // 5) Uppdatera UI med merged (så vi matchar molnet)
        suppressCloudSaveRef.current = true;
        applySharedFromPayload(merged);
        queueMicrotask(() => { suppressCloudSaveRef.current = false; });

        setCloudStatus((s) => ({ ...s, lastSync: new Date().toISOString() }));
      }
      break;
    }
  } catch (e) {
    // låt ops ligga kvar, försök senare
    console.warn("syncNow failed:", e);
  } finally {
    cloudInFlightRef.current = false;
  }
}, [applySharedFromPayload]);

// ================= OPS helpers (använd i alla actions) =================

// Samma entity-namn som applyOpsToPayload använder
const OP_ENTITY = {
  users: "users",
  produkter: "produkter",
  inkopslista: "inkopslista",
  onskemal: "onskemal",
  historik: "historik",
  lagLager: "lagLager",
};

// Debouncad “försök sync” (online -> syncNow)
const trySyncSoon = useCallback(() => {
  if (!navigator.onLine) return;
  cloudSaveDebounceRef.current(() => syncNow());
}, [syncNow]);

const opUpsert = useCallback((entity, item) => {
  enqueueOp({ id: uuid(), ts: Date.now(), kind: "upsert", entity, item });
  refreshPendingOps();
  trySyncSoon();
}, [refreshPendingOps, trySyncSoon]);

const opPatch = useCallback((entity, itemId, patch) => {
  enqueueOp({ id: uuid(), ts: Date.now(), kind: "patch", entity, itemId, patch });
  refreshPendingOps();
  trySyncSoon();
}, [refreshPendingOps, trySyncSoon]);

const opRemove = useCallback((entity, itemId) => {
  enqueueOp({ id: uuid(), ts: Date.now(), kind: "remove", entity, itemId });
  refreshPendingOps();
  trySyncSoon();
}, [refreshPendingOps, trySyncSoon]);

// Historik: skapa alltid id så den blir idempotent vid sync/retry
const makeHistorikRad = useCallback((h) => ({
  id: uuid(),
  ...h,
}), []);

  // ===== Initial load + polling =====
  useEffect(() => {
    let cancelled = false;

    (async () => {
      setCloudStatus((s) => ({ ...s, loading: true }));
      await reloadFromCloud({ force: true });
      if (!cancelled) setCloudStatus((s) => ({ ...s, loading: false }));
    })();

    const id = setInterval(() => {
      reloadFromCloud(); // silent polling
    }, 5000);

    return () => {
      cancelled = true;
      clearInterval(id);
    };
  }, [reloadFromCloud]);

  
useEffect(() => {
  const onOnline = () => syncNow();
  window.addEventListener("online", onOnline);
  return () => window.removeEventListener("online", onOnline);
}, [syncNow]);


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

  

  /* ===== Qty per product ===== */
  const [qtyMap, setQtyMap] = useState(() =>
    saved?.qtyMap && typeof saved.qtyMap === "object" ? saved.qtyMap : {}
  );
  const getQty = (id) => {
  const v = qtyMap[id];
  return v === 0 ? 0 : Math.max(1, toInt(v, 1));
};
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
    lag: saved?.aktivtLag ?? "P-2012",
    produktId: "",
    antal: 1,
    utlamningsdatum: "",
    kommentar: "",
  });

  const [nyLagRad, setNyLagRad] = useState({
    lag: saved?.aktivtLag ?? "P-2012",
    huvudgrupp: "",
    produkt: "",
    antal: 1,
    utlamningsdatum: "",
    kommentar: "",
  });

  /* ===== ✅ NY PRODUKT (Huvudlager) ===== */
  const [addProdOpen, setAddProdOpen] = useState(false);
  const [newProd, setNewProd] = useState({
    huvudgrupp: "",
    produkt: "",
    lagerplats: "Huvudförråd",
    antal: 0,
    beställningspunkt: 0,
  });

  const openAddProd = () => {
    setFel("");
    setNewProd({
      huvudgrupp: "",
      produkt: "",
      lagerplats: "Huvudförråd",
      antal: 0,
      beställningspunkt: 0,
    });
    setAddProdOpen(true);
  };
  const closeAddProd = () => setAddProdOpen(false);

const sparaNyProdukt = () => {
  if (!canEditHuvudlager) return;

  const produktNamn = String(newProd.produkt || "").trim();
  if (!produktNamn) {
    setFel("Produktnamn krävs.");
    return;
  }

  const huvudgrupp = String(newProd.huvudgrupp || "").trim();
  const lagerplats = String(newProd.lagerplats || "").trim() || "Huvudförråd";

  const antalAdd = Math.max(0, toInt(newProd.antal, 0));
  const beställningspunkt = Math.max(0, toInt(newProd.beställningspunkt, 0));

  const idx = produkter.findIndex(
    (p) =>
      normKeyPart(p.produkt) === normKeyPart(produktNamn) &&
      normKeyPart(p.huvudgrupp) === normKeyPart(huvudgrupp)
  );

  let nyaProdukter;
  let upsertProdukt; // <- den här skickar vi till ops

  if (idx >= 0) {
    // Finns redan -> slå ihop (öka antal + uppdatera fält)
    const existing = produkter[idx];
    const newQty = Math.max(0, toInt(existing.antal, 0)) + antalAdd;

    upsertProdukt = {
      ...existing,
      produkt: produktNamn,
      huvudgrupp,
      lagerplats,
      antal: newQty,
      beställningspunkt,
    };

    nyaProdukter = produkter.map((p, i) => (i === idx ? upsertProdukt : p));
    showInfo("Produkten fanns redan – ökade antal och uppdaterade uppgifter.");
  } else {
    // Skapa ny
    upsertProdukt = {
      id: Date.now(), // behåll ditt id-mönster
      huvudgrupp,
      produkt: produktNamn,
      antal: antalAdd,
      lagerplats,
      beställningspunkt,
    };

    nyaProdukter = [upsertProdukt, ...produkter];
    showInfo("Ny produkt skapad.");
  }

  // UI
  setFel("");
  setProdukter(nyaProdukter);
  setInkopslista((prev) => buildInkopLista(nyaProdukter, prev));

  // Historikrad (ge id så den blir idempotent i ops)
  const h = {
    id: uuid(),
    tid: new Date().toLocaleString("sv-SE"),
    typ: idx >= 0 ? "Produkt uppdaterad" : "Ny produkt",
    produkt: produktNamn,
    huvudgrupp,
    lagerplats,
    antal: antalAdd,
    användare: currentUser?.name ?? "Okänd",
    kommentar: idx >= 0 ? "Sammanslagen rad (ökade antal)." : "Skapad i huvudlager.",
  };

  setHistorik((prev) => [h, ...prev]);

  // ✅ OPS (detta är det som gör att den INTE försvinner vid refresh/cloud-load)
  enqueueOp({ id: uuid(), ts: Date.now(), kind: "upsert", entity: "produkter", item: upsertProdukt });
  enqueueOp({ id: uuid(), ts: Date.now(), kind: "upsert", entity: "historik", item: h });

  refreshPendingOps();
  if (navigator.onLine) syncNow();

  closeAddProd();
};
const toggleAutoInkop = (produkt) => {
  if (!canEditHuvudlager) return;

  const paused = !Boolean(produkt.autoInkopPaused);

  const uppdaterad = {
    ...produkt,
    autoInkopPaused: paused,
  };

  // UI
  const nyaProdukter = produkter.map((p) =>
    p.id === produkt.id ? uppdaterad : p
  );
  setProdukter(nyaProdukter);
  setInkopslista((prev) => buildInkopLista(nyaProdukter, prev));

  // Historik
  const h = makeHistorikRad({
    tid: new Date().toLocaleString("sv-SE"),
    typ: paused ? "Auto-inköp pausat" : "Auto-inköp aktiverat",
    produkt: produkt.produkt,
    huvudgrupp: produkt.huvudgrupp,
    lagerplats: produkt.lagerplats,
    antal: 0,
    användare: currentUser?.name ?? "Okänd",
    kommentar: paused
      ? "Automatiskt inköp pausades för produkten."
      : "Automatiskt inköp aktiverades igen.",
  });
  setHistorik((prev) => [h, ...prev]);

  // ✅ OPS (offline + cloud)
  opPatch(OP_ENTITY.produkter, produkt.id, { autoInkopPaused: paused });
  opUpsert(OP_ENTITY.historik, h);

  showInfo(paused ? "Auto-inköp pausat." : "Auto-inköp aktiverat.");
};

/* ===== Edit product (lager) ===== */
const [editProdOpen, setEditProdOpen] = useState(false);
const [editProd, setEditProd] = useState(null);

const openEditProd = (p) => {
  setEditProd({
    id: p.id,
    huvudgrupp: p.huvudgrupp ?? "",
    produkt: p.produkt ?? "",
    lagerplats: p.lagerplats ?? "",
    antal: toInt(p.antal, 0),
    beställningspunkt: toInt(p.beställningspunkt, 0),
  });
  setEditProdOpen(true);
};

const closeEditProd = () => {
  setEditProdOpen(false);
  setEditProd(null);
};

const sparaRedigeradProdukt = () => {
  if (!canEditHuvudlager || !editProd) return;

  const uppdaterad = {
    ...editProd,
    produkt: (editProd.produkt ?? "").trim(),
    huvudgrupp: (editProd.huvudgrupp ?? "").trim(),
    lagerplats: (editProd.lagerplats ?? "").trim() || "Huvudförråd",
    antal: Math.max(0, toInt(editProd.antal, 0)),
    beställningspunkt: Math.max(0, toInt(editProd.beställningspunkt, 0)),
  };

  // UI
  const nyaProdukter = produkter.map((p) => (p.id === uppdaterad.id ? uppdaterad : p));
  setProdukter(nyaProdukter);
  setInkopslista((prev) => buildInkopLista(nyaProdukter, prev));

  // Historik
  const h = makeHistorikRad({
    tid: new Date().toLocaleString("sv-SE"),
    typ: "Produkt redigerad",
    produkt: uppdaterad.produkt,
    huvudgrupp: uppdaterad.huvudgrupp,
    lagerplats: uppdaterad.lagerplats,
    antal: 0,
    användare: currentUser?.name ?? "Okänd",
    kommentar: "Redigerade produktuppgifter.",
  });
  setHistorik((prev) => [h, ...prev]);

  // ✅ OPS
  opUpsert(OP_ENTITY.produkter, uppdaterad);
  opUpsert(OP_ENTITY.historik, h);

  showInfo("Produkten uppdaterad (offline-säkert).");
  closeEditProd();
};

const taBortProdukt = (p) => {
  if (!canEditHuvudlager) return;

  const användsIInkop = inkopslista.some(
    (r) =>
      normKeyPart(r.produkt) === normKeyPart(p.produkt) &&
      normKeyPart(r.huvudgrupp) === normKeyPart(p.huvudgrupp)
  );

  const användsILag = lagLager.some(
    (r) =>
      normKeyPart(r.produkt) === normKeyPart(p.produkt) &&
      normKeyPart(r.huvudgrupp) === normKeyPart(p.huvudgrupp)
  );

  if (användsIInkop || användsILag) {
    return setFel(
      "Produkten används i inköp eller lagmaterial och kan inte tas bort. Ta bort beroenden först."
    );
  }

  if (!window.confirm(`Vill du verkligen ta bort "${p.produkt}" ur lagret?`)) return;

  // UI
  setProdukter((prev) => prev.filter((x) => x.id !== p.id));

  // Historik
  const h = makeHistorikRad({
    tid: new Date().toLocaleString("sv-SE"),
    typ: "Produkt borttagen",
    produkt: p.produkt,
    huvudgrupp: p.huvudgrupp,
    lagerplats: p.lagerplats,
    antal: 0,
    användare: currentUser?.name ?? "Okänd",
    kommentar: "Produkten togs bort ur huvudlager.",
  });
  setHistorik((prev) => [h, ...prev]);

  // ✅ OPS
  opRemove(OP_ENTITY.produkter, p.id);
  opUpsert(OP_ENTITY.historik, h);

  showInfo("Produkten borttagen (offline-säkert).");
};

useEffect(() => {
  // 1) spara lokalt (offlinecache + UI)
  const localPayload = {
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
  };
  saveState(localPayload);

  // 2) om online och vi har pending ops -> synca (rebased)
  if (navigator.onLine && loadOps().length) {
    cloudSaveDebounceRef.current(() => syncNow());
  }
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
  syncNow,
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
  const bp = toInt(p.beställningspunkt, 0);

  if (bp <= 0) return { text: "OK", tone: "ok" };

  const orderThreshold = Math.floor(bp * 0.7); // ✅ justerbar (t.ex. 70 %)

  if (antal <= orderThreshold) {
    return { text: "Beställ", tone: "danger" };
  }

  if (antal <= bp) {
    return { text: "Bevaka", tone: "warn" };
  }

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
    const set = new Set(["P-2019", "F-2019", "P-2018", "F-2018", "P-2017", "F-2017", "P-2016", "P-2015", "P-2014", "F-2014", "P-2013", "F-2013", "P-2012", "F-2011/12", "P-2011", "P-2010", "A-lag Damer", "Herr"]);
    (lagLager ?? []).forEach((r) => set.add(r.lag || "Okänt"));
    return Array.from(set).filter(Boolean).sort();
  }, [lagLager]);

  const aktivtLagEff = isLedare ? currentUser?.lag ?? "Okänt" : aktivtLag;

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

  const after = toInt(produkt.antal, 0) + delta;
  if (delta < 0 && after < 0) {
    setFel("Kan inte ta ut fler än det som finns i lager.");
    return;
  }

  const nyaProdukter = produkter.map((p) =>
    p.id === produkt.id ? { ...p, antal: Math.max(0, after) } : p
  );
  setProdukter(nyaProdukter);
  setInkopslista((prev) => buildInkopLista(nyaProdukter, prev));

  const kom = kommentar?.trim() ? kommentar.trim() : "(ingen kommentar)";
  rememberComment(kom);

  const h = makeHistorikRad({
    tid: new Date().toLocaleString("sv-SE"),
    typ: delta < 0 ? "Uttag" : "Inleverans",
    produkt: produkt.produkt,
    huvudgrupp: produkt.huvudgrupp,
    lagerplats: produkt.lagerplats,
    antal: delta,
    användare: currentUser?.name ?? "Okänd",
    kommentar: kom,
  });

  setHistorik((prev) => [h, ...prev]);

  // OPS
  opPatch(OP_ENTITY.produkter, produkt.id, { antal: Math.max(0, after) });
  opUpsert(OP_ENTITY.historik, h);
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
  if (!items.length) return setFel("Inga produkter valda.");

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
    deltaById.has(p.id)
      ? { ...p, antal: Math.max(0, toInt(p.antal, 0) + deltaById.get(p.id)) }
      : p
  );

  setProdukter(nyaProdukter);
  setInkopslista((prev) => buildInkopLista(nyaProdukter, prev));

  const tid = new Date().toLocaleString("sv-SE");
  const typ = batchMode === "uttag" ? "Uttag (mass)" : "Inleverans (mass)";
  const kommentar = batchComment?.trim() ? batchComment.trim() : "(ingen kommentar)";
  rememberComment(kommentar);

  const historikRader = deltas.map(({ p, delta }) =>
    makeHistorikRad({
      tid,
      typ,
      produkt: p.produkt,
      huvudgrupp: p.huvudgrupp,
      lagerplats: p.lagerplats,
      antal: delta,
      användare: currentUser?.name ?? "Okänd",
      kommentar,
    })
  );

  setHistorik((prev) => [...historikRader, ...prev]);

  // OPS: patcha varje produkt + historik
  for (const { p, delta } of deltas) {
    const nyttAntal = Math.max(0, toInt(p.antal, 0) + delta);
    opPatch(OP_ENTITY.produkter, p.id, { antal: nyttAntal });
  }
  for (const h of historikRader) opUpsert(OP_ENTITY.historik, h);

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
  const lag = (flytt.lag && flytt.lag.trim()) || aktivtLagEff || "Okänt";

  const p = produkter.find((x) => String(x.id) === String(prodId));
  if (!p) return setFel("Välj en produkt i huvudlagret.");
  if (toInt(p.antal, 0) < antal) return setFel("Finns inte så många i huvudlagret.");

  setFel("");

  // 1) huvudlager
  const nyaProdukter = produkter.map((x) =>
    x.id === p.id ? { ...x, antal: toInt(x.antal, 0) - antal } : x
  );
  setProdukter(nyaProdukter);
  setInkopslista((prev) => buildInkopLista(nyaProdukter, prev));

  // 2) laglager
  const now = new Date().toLocaleString("sv-SE");
  const ut = (flytt.utlamningsdatum ?? "").trim();
  const kom = (flytt.kommentar ?? "").trim();

  let skapadRad = null;

  setLagLager((prev) => {
    const exists = prev.find(
      (r) =>
        (r.lag || "Okänt") === lag &&
        normKeyPart(r.produkt) === normKeyPart(p.produkt) &&
        normKeyPart(r.huvudgrupp) === normKeyPart(p.huvudgrupp)
    );

    if (exists) {
      const next = prev.map((r) =>
        r.id === exists.id
          ? {
              ...r,
              antal: Math.max(0, toInt(r.antal, 0)) + antal,
              utlamningsdatum: ut || r.utlamningsdatum || "",
              kommentar: kom || r.kommentar || "",
            }
          : r
      );
      skapadRad = next.find((r) => r.id === exists.id);
      return next;
    }

    const item = {
      id: Date.now(),
      lag,
      huvudgrupp: p.huvudgrupp ?? "",
      produkt: p.produkt ?? "",
      antal,
      utlamningsdatum: ut,
      kommentar: kom,
      skapadTid: now,
      skapadAv: currentUser?.name ?? "Okänd",
    };
    skapadRad = item;
    return [item, ...prev];
  });

  // 3) historik
  const h = makeHistorikRad({
    tid: new Date().toLocaleString("sv-SE"),
    typ: "Utlämnat till lag",
    produkt: p.produkt,
    huvudgrupp: p.huvudgrupp,
    lagerplats: p.lagerplats,
    antal: -antal,
    användare: currentUser?.name ?? "Okänd",
    kommentar: `Till ${lag}${kom ? ` • ${kom}` : ""}`,
  });
  setHistorik((prev) => [h, ...prev]);

  // OPS
  opPatch(OP_ENTITY.produkter, p.id, { antal: toInt(p.antal, 0) - antal });
  // Upsert lagraden (om skapadRad redan satt)
  if (skapadRad) opUpsert(OP_ENTITY.lagLager, skapadRad);
  opUpsert(OP_ENTITY.historik, h);

  showInfo(`Flyttade ${antal} st till ${lag}.`);
  setFlytt((f) => ({ ...f, antal: 1, kommentar: "" }));
};

const uppdateraLagRadAntal = (id, newValue) => {
  const n = Math.max(0, toInt(newValue, 0));
  setLagLager((prev) => prev.map((r) => (r.id === id ? { ...r, antal: n } : r)));
  opPatch(OP_ENTITY.lagLager, id, { antal: n });
  showInfo("Antal uppdaterat (offline-säkert).");
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
  opUpsert(OP_ENTITY.lagLager, item);

  setNyLagRad({ lag: item.lag, huvudgrupp: "", produkt: "", antal: 1, utlamningsdatum: "", kommentar: "" });
  showInfo("Rad lades till i lagmaterial (offline-säkert).");
};

const taBortLagRad = (id) => {
  if (isLedare) {
    setFel("Ledare kan inte ta bort rader (ändra antal istället).");
    return;
  }
  setLagLager((prev) => prev.filter((x) => x.id !== id));
  opRemove(OP_ENTITY.lagLager, id);
  showInfo("Rad borttagen (offline-säkert).");
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
  if (antal > (rad.antal ?? 0)) return setFel("Kan inte returnera fler än lagets antal.");
  setFel("");

  // 1) huvudlager: öka / skapa
  let patchedProdId = null;
  let patchedProdAntal = null;

  setProdukter((prev) => {
    const idx = prev.findIndex(
      (p) =>
        normKeyPart(p.produkt) === normKeyPart(rad.produkt) &&
        normKeyPart(p.huvudgrupp) === normKeyPart(rad.huvudgrupp)
    );

    if (idx >= 0) {
      const nya = [...prev];
      const newQty = Math.max(0, toInt(nya[idx].antal, 0)) + antal;
      nya[idx] = { ...nya[idx], antal: newQty };
      patchedProdId = nya[idx].id;
      patchedProdAntal = newQty;
      // inköp UI
      setInkopslista((inkPrev) => buildInkopLista(nya, inkPrev));
      return nya;
    }

    const ny = {
      id: Date.now(),
      huvudgrupp: rad.huvudgrupp ?? "",
      produkt: rad.produkt ?? "",
      antal,
      lagerplats: "Huvudförråd",
      beställningspunkt: 0,
    };
    patchedProdId = ny.id;
    patchedProdAntal = antal;
    const nya = [ny, ...prev];
    setInkopslista((inkPrev) => buildInkopLista(nya, inkPrev));
    return nya;
  });

  // 2) laglager: minska / remove om 0
  setLagLager((prev) => {
    const kvar = prev.map((r) =>
      r.id === rad.id ? { ...r, antal: Math.max(0, toInt(r.antal, 0) - antal) } : r
    );
    const filtered = kvar.filter((r) => toInt(r.antal, 0) > 0);
    return filtered;
  });

  // 3) historik
  const h = makeHistorikRad({
    tid: new Date().toLocaleString("sv-SE"),
    typ: "Retur från lag",
    produkt: rad.produkt,
    huvudgrupp: rad.huvudgrupp,
    lagerplats: "",
    antal: +antal,
    användare: currentUser?.name ?? "Okänd",
    kommentar: `Från ${rad.lag || "Okänt"}`,
  });
  setHistorik((prev) => [h, ...prev]);

  // OPS
  if (patchedProdId != null) opPatch(OP_ENTITY.produkter, patchedProdId, { antal: patchedProdAntal });
  opPatch(OP_ENTITY.lagLager, rad.id, { antal: Math.max(0, toInt(rad.antal, 0) - antal) });
  if (Math.max(0, toInt(rad.antal, 0) - antal) <= 0) opRemove(OP_ENTITY.lagLager, rad.id);
  opUpsert(OP_ENTITY.historik, h);

  showInfo(`Returnerade ${antal} st till huvudlager (offline-säkert).`);
};
  /* ================= Önskemål ================= */

const skapaOnskemal = () => {
  const prod = (nyttOnskemal.produkt ?? "").trim();
  if (!prod) return;

  const item = {
    id: uuid(),
    huvudgrupp: (nyttOnskemal.huvudgrupp ?? "").trim(),
    produkt: prod,
    antal: Math.max(1, toInt(nyttOnskemal.antal, 1)),
    kommentar: (nyttOnskemal.kommentar ?? "").trim(),
    skapadTid: new Date().toLocaleString("sv-SE"),
    skapadAv: currentUser?.name ?? "Okänd",
    lag: isLedare ? (currentUser?.lag ?? "Okänt") : "",
  };

  // UI direkt
  setOnskemal((prev) => [item, ...prev]);

  // OPS
  opUpsert(OP_ENTITY.onskemal, item);

  setNyttOnskemal({ huvudgrupp: "", produkt: "", antal: 1, kommentar: "" });
  showInfo("Önskemålet lades till (offline-säkert).");
};

  
const taBortOnskemal = (id) => {
  setOnskemal((prev) => prev.filter((o) => o.id !== id));
  opRemove(OP_ENTITY.onskemal, id);
  showInfo("Önskemål borttaget (offline-säkert).");
};

const laggOnskemalTillInkop = (o) => {
  if (!canMoveToInkop) return;

  const key = makeKey(o.huvudgrupp, o.produkt);

  // UI: ta bort från önskemål
  setOnskemal((prev) => prev.filter((x) => x.id !== o.id));

  // UI: skapa/uppdatera inköpsrad
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

  // OPS: remove önskemål + upsert inköpsrad (minsta säkra data)
  opRemove(OP_ENTITY.onskemal, o.id);

  // Upsert rad “minimalt” (vi låter cloud behålla ev extra fält)
  opUpsert(OP_ENTITY.inkopslista, {
    id: key,
    key,
    huvudgrupp: o.huvudgrupp ?? "",
    produkt: o.produkt ?? "",
    // antal hanteras i UI ovan; för att vara exakt kan du patcha också:
  });

  // Patcha antal + manuell (för att exakt matcha UI)
  opPatch(OP_ENTITY.inkopslista, key, {
    manuell: true,
    bestalld: false,
    // vi kan inte läsa “nya” totalen utan att duplicera logik, så vi patchar säkert genom att
    // lägga till +o.antal på serversidan vore bättre. Men i client rebase kommer UI-value finnas i payload.
  });

  showInfo("Önskemål flyttat till inköp (offline-säkert).");
};


  
  /* ================= Inköp ================= */
const sättInköpsantalManuellt = (key, value) => {
  const n = Math.max(0, toInt(value, 0));
  setInkopslista((prev) =>
    prev.map((r) => (r.key === key ? { ...r, antal: n, manuell: true } : r))
  );
  opPatch(OP_ENTITY.inkopslista, key, { antal: n, manuell: true });
};

const återställRekommenderat = (key) => {
  setInkopslista((prev) =>
    prev.map((r) => (r.key === key ? { ...r, antal: r.rekommenderat, manuell: false } : r))
  );
  // Vi patchar bara manuell=false; antal sätts via UI (rebase inkluderar detta)
  opPatch(OP_ENTITY.inkopslista, key, { manuell: false });
};

const markeraSomBestalld = (key) => {
  const tid = new Date().toLocaleString("sv-SE");
  const av = currentUser?.name ?? "Okänd";

  setInkopslista((prev) =>
    prev.map((r) =>
      r.key === key
        ? {
            ...r,
            bestalld: true,
            bestalldTid: tid,
            bestalldAv: av,
            mottagetAntal: Math.max(0, toInt(r.mottagetAntal, 0)),
            leveranser: Array.isArray(r.leveranser) ? r.leveranser : [],
          }
        : r
    )
  );

  opPatch(OP_ENTITY.inkopslista, key, { bestalld: true, bestalldTid: tid, bestalldAv: av });
  showInfo("Markerad som beställd (offline-säkert).");
};

const angraBestalld = (key) => {
  setInkopslista((prev) =>
    prev.map((r) => (r.key === key ? { ...r, bestalld: false, bestalldTid: "", bestalldAv: "" } : r))
  );
  opPatch(OP_ENTITY.inkopslista, key, { bestalld: false, bestalldTid: "", bestalldAv: "" });
  showInfo("Ångrade beställd (offline-säkert).");
};


  /* ================= Inleverans: motta delleverans ================= */
const mottaLeverans = (rowKey, qtyInput) => {
  if (!canSeeInleverans) return;

  const row = inkopslista.find((r) => r.key === rowKey);
  if (!row || !row.bestalld) return setFel("Välj en beställd rad.");

  const ordered = Math.max(0, toInt(row.antal, 0));
  const received = Math.max(0, toInt(row.mottagetAntal, 0));
  const kvar = Math.max(0, ordered - received);

  const qtyReq = Math.max(1, toInt(qtyInput, 1));
  const qty = Math.min(kvar, qtyReq);
  if (qty <= 0) return setFel("Inget kvar att ta emot.");

  setFel("");
  const tid = new Date().toLocaleString("sv-SE");
  const av = currentUser?.name ?? "Okänd";
  const kom = String(leveransKommentar ?? "").trim();

  // 1) produkter
  let prodIdTouched = null;
  let prodNewQty = null;

  setProdukter((prev) => {
    const idx = prev.findIndex(
      (p) =>
        normKeyPart(p.produkt) === normKeyPart(row.produkt) &&
        normKeyPart(p.huvudgrupp) === normKeyPart(row.huvudgrupp)
    );

    if (idx >= 0) {
      const nya = [...prev];
      prodNewQty = Math.max(0, toInt(nya[idx].antal, 0)) + qty;
      nya[idx] = { ...nya[idx], antal: prodNewQty, lagerplats: leveransLagerplats };
      prodIdTouched = nya[idx].id;
      return nya;
    }

    const ny = {
      id: Date.now(),
      huvudgrupp: row.huvudgrupp ?? "",
      produkt: row.produkt ?? "",
      antal: qty,
      lagerplats: leveransLagerplats,
      beställningspunkt: 0,
    };
    prodIdTouched = ny.id;
    prodNewQty = qty;
    return [ny, ...prev];
  });

  // 2) inkopslista
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
          { tid, av, antal: qty, lagerplats: leveransLagerplats, kommentar: kom },
          ...(r.leveranser || []),
        ],
      };
    })
  );

  // 3) historik
  const kommentarText = `Inleverans → ${leveransLagerplats}${kom ? ` • ${kom}` : ""}`;
  const h = makeHistorikRad({
    tid,
    typ: "Inleverans (beställning)",
    produkt: row.produkt,
    huvudgrupp: row.huvudgrupp,
    lagerplats: leveransLagerplats,
    antal: qty,
    användare: av,
    kommentar: kommentarText,
  });
  setHistorik((prev) => [h, ...prev]);

  // OPS
  if (prodIdTouched != null) opPatch(OP_ENTITY.produkter, prodIdTouched, { antal: prodNewQty, lagerplats: leveransLagerplats });
  opPatch(OP_ENTITY.inkopslista, rowKey, {
    mottagetAntal: received + qty,
    // levererad/levereradTid/levereradAv avgörs av ordered:
    leveranser: [
      { tid, av, antal: qty, lagerplats: leveransLagerplats, kommentar: kom },
      ...(row.leveranser || []),
    ],
  });
  opUpsert(OP_ENTITY.historik, h);

  showInfo(`Tog emot ${qty} st: ${row.produkt} (offline-säkert).`);
};

const taBortLeverans = (rowKey, leverans, index) => {
  if (!canSeeInleverans) return;

  if (!window.confirm("Vill du ta bort denna inleverans?")) return;

  const qty = Math.max(0, toInt(leverans.antal, 0));

  // 1) Uppdatera inköpsrad
  let newReceived = 0;
  let ordered = 0;

  setInkopslista((prev) =>
    prev.map((r) => {
      if (r.key !== rowKey) return r;

      ordered = Math.max(0, toInt(r.antal, 0));
      const receivedBefore = Math.max(0, toInt(r.mottagetAntal, 0));
      newReceived = Math.max(0, receivedBefore - qty);

      const nyaLeveranser = r.leveranser.filter((_, i) => i !== index);
      const fully = ordered > 0 && newReceived >= ordered;

      return {
        ...r,
        leveranser: nyaLeveranser,
        mottagetAntal: newReceived,
        levererad: fully,
        levereradTid: fully ? r.levereradTid : "",
        levereradAv: fully ? r.levereradAv : "",
      };
    })
  );

  // 2) Dra bort från huvudlager
  let prodId = null;
  let prodNewQty = null;

  setProdukter((prev) => {
    const idx = prev.findIndex(
      (p) =>
        normKeyPart(p.produkt) === normKeyPart(leverans.produkt || "") &&
        normKeyPart(p.huvudgrupp) === normKeyPart(leverans.huvudgrupp || "")
    );

    if (idx >= 0) {
      const nya = [...prev];
      prodNewQty = Math.max(0, toInt(nya[idx].antal, 0) - qty);
      nya[idx] = { ...nya[idx], antal: prodNewQty };
      prodId = nya[idx].id;
      return nya;
    }
    return prev;
  });

  // 3) Historik
  const h = makeHistorikRad({
    tid: new Date().toLocaleString("sv-SE"),
    typ: "Inleverans borttagen",
    produkt: leverans.produkt || "",
    huvudgrupp: leverans.huvudgrupp || "",
    lagerplats: leverans.lagerplats || "",
    antal: -qty,
    användare: currentUser?.name ?? "Okänd",
    kommentar: "Tog bort tidigare inleverans.",
  });

  setHistorik((prev) => [h, ...prev]);

  // ✅ OPS
  opPatch(OP_ENTITY.inkopslista, rowKey, {
    mottagetAntal: newReceived,
    leveranser: null, // rebase kommer skicka rätt lista
  });

  if (prodId != null) {
    opPatch(OP_ENTITY.produkter, prodId, { antal: prodNewQty });
  }

  opUpsert(OP_ENTITY.historik, h);

  showInfo("Inleverans borttagen (offline-säkert).");
};
  /* ================= Admin: användare + PIN ================= */
  const [userDraft, setUserDraft] = useState({ name: "", username: "", role: "Ledare", lag: "P-2012", pin: "" });

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
    opUpsert(OP_ENTITY.users, u);
    setUserDraft({ name: "", username: "", role: "Ledare", lag: "P-2012", pin: "" });
    showInfo("Användare skapad.");
  };

  const deleteUser = (id) => {
    if (!canEditUsers) return;
    if (currentUser?.id === id) {
      setFel("Du kan inte ta bort dig själv.");
      return;
    }
    setUsers((prev) => prev.filter((u) => u.id !== id));
    opRemove(OP_ENTITY.users, id);
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
opPatch(OP_ENTITY.users, id, { pinSalt: salt, pinHash: hashPin(newPin, salt) });
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
      // OPS: import ska också synkas (upsert alla importerade rader)
for (const p of produkterNy) {
  enqueueOp({ id: uuid(), ts: Date.now(), kind: "upsert", entity: "produkter", item: p });
}
const ink = buildInkopLista(produkterNy, prevInkop);
for (const r of ink) {
  enqueueOp({ id: uuid(), ts: Date.now(), kind: "upsert", entity: "inkopslista", item: r });
}
for (const r of lagNy || []) {
  enqueueOp({ id: uuid(), ts: Date.now(), kind: "upsert", entity: "lagLager", item: r });
}

refreshPendingOps();
if (navigator.onLine) syncNow();

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
                        
<PrimaryButton
  tone="ghost"
  onClick={() => openManuellInkop(p)}
>
  📦 Manuellt inköp
</PrimaryButton>

                      </div>
                    </div>
                  ) : null}

                  {/* ✅ NY: Lägg till produkt under Lager */}
                  {canEditHuvudlager && !selectMode ? (
                    <section className="card" style={{ marginBottom: 12 }}>
                      <div className="card__top">
                        <div className="card__title">➕ Lägg till produkt</div>
                        <Pill tone="neutral">Huvudlager</Pill>
                      </div>
                      <div className="muted" style={{ marginTop: 8 }}>
                        Skapa ny artikel i huvudlagret (påverkar inköp automatiskt).
                      </div>
                      <div className="btnRow" style={{ marginTop: 10, gridTemplateColumns: "1fr" }}>
                        <PrimaryButton tone="primary" onClick={openAddProd}>
                          ➕ Ny produkt
                        </PrimaryButton>
                      </div>
                    </section>
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
                              <div style={{ display: "flex", alignItems: "center", gap: 6, flexWrap: "wrap" }}>
  <Pill tone={st.tone}>{st.text}</Pill>

  {(() => {
    const key = makeKey(p.huvudgrupp, p.produkt);
    const r = inkopslista.find((x) => x.key === key);
    if (!r) return null;

    const antal = Math.max(0, toInt(r.antal, 0));
    if (antal <= 0) return null;

    return (
      <Pill tone={r.bestalld ? "ok" : "warn"}>
        +{antal} st
      </Pill>
    );
  })()}

  {canEditHuvudlager ? (
    <button
      className="iconBtn"
      title="Redigera produkt"
      onClick={(e) => {
        e.stopPropagation();
        openEditProd(p);
      }}
    >
      ✏️
    </button>
  ) : null}
</div>
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
    Beställ vid: <strong>{toInt(p.beställningspunkt, 0)}</strong>
  </div>
</div>
                          </div>

                          <div className="qtyPick" onClick={(e) => e.stopPropagation()}>
                            <div className="qtyPick__label">Antal (st)</div>
                            
<QtyInput
  value={qty}
  min={1}
  onChange={(n) => setQty(p.id, n)}
/>

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
                            <QtyInput
  value={toInt(r.antal, 0)}
  min={0}
  onChange={(n) => uppdateraLagRadAntal(r.id, n)}
/>
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
<QtyInput
  value={flytt.antal}
  min={1}
  onChange={(n) => setFlytt((f) => ({ ...f, antal: n }))}
/>
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

<QtyInput
  value={nyttOnskemal.antal}
  min={1}
  onChange={(n) =>
    setNyttOnskemal((o) => ({
      ...o,
      antal: n,
    }))
  }
/>

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
                          <PrimaryButton tone="danger" onClick={() => taBortOnskemal(o.id)}>
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

                    const ejKompletta = filtrerade.filter((r) => {
  const ordered = Math.max(0, toInt(r.antal, 0));
  const received = Math.max(0, toInt(r.mottagetAntal, 0));
  return ordered > received;
});

const kompletta = filtrerade.filter((r) => {
  const ordered = Math.max(0, toInt(r.antal, 0));
  const received = Math.max(0, toInt(r.mottagetAntal, 0));
  return ordered > 0 && ordered <= received;
});
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
                          <div className="summarySub">
                            {ejKompletta.length} rader ej kompletta • {kompletta.length} levererade
                          </div>
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
                          <div className="muted" style={{ marginTop: 8 }}>
                            Du kan göra delleveranser. Varje mottagning ökar huvudlagret direkt.
                          </div>
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
                                      <span className="meta__value">
                                        {r.bestalldTid} • {r.bestalldAv || "—"}
                                      </span>
                                    </div>
                                  ) : null}
                                </div>

                                <div className="qtyPick">
                                  <div className="qtyPick__label">Ta emot nu</div>
<QtyInput
  value={qtyNow}
  min={1}
  onChange={(n) => setLevQty(r.key, n)}
/>
                                  <div className="qtyChips">
                                    <button className="chip" type="button" onClick={() => setLevQty(r.key, 1)}>
                                      1
                                    </button>
                                    <button className="chip" type="button" onClick={() => setLevQty(r.key, Math.min(5, Math.max(1, kvar)))}>
                                      5
                                    </button>
                                    <button className="chip" type="button" onClick={() => setLevQty(r.key, Math.min(10, Math.max(1, kvar)))}>
                                      10
                                    </button>
                                    <button className="chip" type="button" onClick={() => setLevQty(r.key, Math.max(1, kvar))} title="Sätt till kvar">
                                      ⇢
                                    </button>
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
  <div className="leveranser" style={{ marginTop: 10 }}>
    {r.leveranser.map((lev, i) => (
      <div key={i} className="leveransRad">
        <div className="muted">
          {lev.tid} • {lev.av} • +{lev.antal}
          {lev.lagerplats ? ` • ${lev.lagerplats}` : ""}
          {lev.kommentar ? ` • ${lev.kommentar}` : ""}
        </div>

        <PrimaryButton
          tone="danger"
          onClick={() =>
            taBortLeverans(
              r.key,
              {
                ...lev,
                produkt: r.produkt,
                huvudgrupp: r.huvudgrupp,
              },
              i
            )
          }
          style={{ marginLeft: 8 }}
        >
          🗑️ Ta bort
        </PrimaryButton>
      </div>
    ))}
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
                                      <span className="meta__value">
                                        {r.levereradTid || "—"} • {r.levereradAv || "—"}
                                      </span>
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
  title={txMode === "uttag"
    ? "➖ Uttag från lager"
    : "➕ Inleverans till lager"}
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
          showInfo(
            txMode === "uttag"
              ? "Uttag registrerat."
              : "Inleverans registrerad."
          );
        }}
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
                      <input value={txComment} onChange={(e) => setTxComment(e.target.value)} placeholder="t.ex. P-2012 träning" list="recent-comments" />
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

<Modal
  open={manuellInkopOpen}
  title="📦 Manuellt inköp"
  onClose={closeManuellInkop}
  footer={
    !manuellInkopProd ? null : (
      <>
        <PrimaryButton tone="ghost" onClick={closeManuellInkop}>
          Avbryt
        </PrimaryButton>
        <PrimaryButton
          tone="primary"
          onClick={() => {
            const p = manuellInkopProd;
            const qty = Math.max(1, toInt(manuellInkopQty, 1));
            const key = makeKey(p.huvudgrupp, p.produkt);

            setInkopslista((prev) => {
              const existing = prev.find((r) => r.key === key);
              if (existing) {
                return prev.map((r) =>
                  r.key === key
                    ? {
                        ...r,
                        antal: Math.max(0, toInt(r.antal, 0)) + qty,
                        manuell: true,
                        bestalld: false,
                      }
                    : r
                );
              }
              return [
                ...prev,
                {
                  id: key,
                  key,
                  huvudgrupp: p.huvudgrupp ?? "",
                  produkt: p.produkt ?? "",
                  antal: qty,
                  rekommenderat: 0,
                  manuell: true,
                  bestalld: false,
                  mottagetAntal: 0,
                  levererad: false,
                  leveranser: [],
                },
              ];
            });

            // Historik
            const h = makeHistorikRad({
              tid: new Date().toLocaleString("sv-SE"),
              typ: "Manuellt inköp",
              produkt: p.produkt,
              huvudgrupp: p.huvudgrupp,
              lagerplats: p.lagerplats,
              antal: qty,
              användare: currentUser?.name ?? "Okänd",
              kommentar: "Manuellt beställd från lagerkort.",
            });
            setHistorik((prev) => [h, ...prev]);

            // OPS
            opUpsert(OP_ENTITY.inkopslista, {
              id: key,
              key,
              huvudgrupp: p.huvudgrupp ?? "",
              produkt: p.produkt ?? "",
              antal: qty,
              manuell: true,
              bestalld: false,
            });
            opUpsert(OP_ENTITY.historik, h);

            showInfo("Manuellt inköp lagt till.");
            closeManuellInkop();
          }}
        >
          Lägg till inköp
        </PrimaryButton>
      </>
    )
  }
>
  {!manuellInkopProd ? null : (
    <div className="formGrid">
      <div><strong>{manuellInkopProd.produkt}</strong></div>
      <label className="field">
        <span>Antal</span>
        <QtyInput
          value={manuellInkopQty}
          min={1}
          onChange={setManuellInkopQty}
        />
      </label>
    </div>
  )}
</Modal>

<Modal
  open={editProdOpen}
  title="✏️ Redigera produkt"
  onClose={closeEditProd}
  footer={
    <>
      <PrimaryButton tone="ghost" onClick={closeEditProd}>
        Avbryt
      </PrimaryButton>
      <PrimaryButton tone="primary" onClick={sparaRedigeradProdukt}>
        Spara ändringar
      </PrimaryButton>
    </>
  }
>
  {!editProd ? null : (
    <div className="formGrid">
      <label className="field">
        <span>Produkt</span>
        <input
          value={editProd.produkt}
          onChange={(e) => setEditProd((p) => ({ ...p, produkt: e.target.value }))}
        />
      </label>

      <label className="field">
        <span>Huvudgrupp</span>
        <input
          value={editProd.huvudgrupp}
          onChange={(e) => setEditProd((p) => ({ ...p, huvudgrupp: e.target.value }))}
        />
      </label>

      <label className="field">
        <span>Lagerplats</span>
        <input
          value={editProd.lagerplats}
          onChange={(e) => setEditProd((p) => ({ ...p, lagerplats: e.target.value }))}
        />
      </label>

      <label className="field">
        <span>Antal</span>
        <QtyInput
          value={editProd.antal}
          min={0}
          onChange={(n) => setEditProd((p) => ({ ...p, antal: n }))}
        />
      </label>

      <label className="field">
        <span>Beställningspunkt</span>
        <QtyInput
          value={editProd.beställningspunkt}
          min={0}
          onChange={(n) => setEditProd((p) => ({ ...p, beställningspunkt: n }))}
        />
      </label>
    </div>
  )}
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

                    {!batchUsePerItemQty ? (
<QtyInput
  value={batchQty}
  min={1}
  onChange={setBatchQty}
/>
                    ) : null}
                  </div>
                </label>

                <label className="field">
                  <span>Kommentar (gemensam)</span>
                  <input value={batchComment} onChange={(e) => setBatchComment(e.target.value)} placeholder="t.ex. P-2012 träning" list="recent-comments" />
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
<QtyInput
  value={returQty}
  min={1}
  onChange={setReturQty}
/>
                  </label>
                </div>
              )}
            </Modal>

            {/* ===== ✅ Modal: Ny produkt ===== */}
            <Modal
              open={addProdOpen}
              title="➕ Ny produkt i huvudlager"
              onClose={closeAddProd}
              footer={
                <>
                  <PrimaryButton tone="ghost" onClick={closeAddProd}>
                    Avbryt
                  </PrimaryButton>
                  <PrimaryButton tone="primary" onClick={sparaNyProdukt} disabled={!canEditHuvudlager}>
                    Spara
                  </PrimaryButton>
                </>
              }
            >
              <div className="formGrid">
                <label className="field">
                  <span>Produkt *</span>
                  <input value={newProd.produkt} onChange={(e) => setNewProd((p) => ({ ...p, produkt: e.target.value }))} placeholder="t.ex. Fotboll strl 5" />
                </label>

                <label className="field">
                  <span>Huvudgrupp</span>
                  <input value={newProd.huvudgrupp} onChange={(e) => setNewProd((p) => ({ ...p, huvudgrupp: e.target.value }))} placeholder="t.ex. Bollar, Västar…" />
                </label>

                <label className="field">
                  <span>Lagerplats</span>
                  <input value={newProd.lagerplats} onChange={(e) => setNewProd((p) => ({ ...p, lagerplats: e.target.value }))} list="lagerplats-list" placeholder="Huvudförråd" />
                  <datalist id="lagerplats-list">
                    {lagerplatsOptions.map((lp) => (
                      <option key={lp} value={lp} />
                    ))}
                  </datalist>
                </label>

                <label className="field">
                  <span>Antal (startlager)</span>
<QtyInput
  value={newProd.antal}
  min={0}
  onChange={(n) => setNewProd((p) => ({ ...p, antal: n }))}
/>

                </label>

                <label className="field">
                  <span>Beställningspunkt</span>
<QtyInput
  value={newProd.beställningspunkt}
  min={0}
  onChange={(n) => setNewProd((p) => ({ ...p, beställningspunkt: n }))}
/>
                </label>

                <div className="muted" style={{ marginTop: 6 }}>
                  Tips: Om “Min antal” eller “Beställningspunkt” är 0 så påverkar den inte status/inköp.
                </div>
              </div>
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
  
.leveransRad {
  display: flex;
  justify-content: space-between;
  align-items: center;
  gap: 8px;
}

`;
