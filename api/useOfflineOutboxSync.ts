// useOfflineOutboxSync.ts
import { useCallback, useEffect, useRef, useState } from "react";
import { applyOpsToPayload } from "./applyOps";
import { clearOps, loadOps } from "./outbox";
import { fetchCloudState, saveCloudState } from "./cloudApi";

type ApplyShared = (p: any) => void;

export function useOfflineOutboxSync(applyShared: ApplyShared) {
  const inFlight = useRef(false);
  const cloudVersionRef = useRef(0);
  const hydratedRef = useRef(false);

  const [status, setStatus] = useState<{ loading: boolean; lastSync: string; pendingOps: number }>({
    loading: false,
    lastSync: "",
    pendingOps: loadOps().length,
  });

  const refreshPendingCount = () => {
    setStatus(s => ({ ...s, pendingOps: loadOps().length }));
  };

  // 1) Hydratera: hämta cloud -> apply -> replay ops -> apply igen
  const hydrateFromCloud = useCallback(async () => {
    if (inFlight.current) return;
    inFlight.current = true;
    setStatus(s => ({ ...s, loading: true }));

    try {
      const data = await fetchCloudState();
      const cloudPayload = data.payload ?? { __version: 0 };
      cloudVersionRef.current = cloudPayload.__version ?? 0;

      // Applicera cloud först
      applyShared(cloudPayload);

      // Replay:a ops (så cloud inte blåser bort lokala önskemål)
      const ops = loadOps();
      if (ops.length) {
        const mergedForUi = applyOpsToPayload(cloudPayload, ops);
        applyShared(mergedForUi);
      }

      hydratedRef.current = true;
      setStatus(s => ({ ...s, lastSync: data.updatedAt || cloudPayload.__updatedAt || s.lastSync }));
      refreshPendingCount();
    } catch (e) {
      hydratedRef.current = true; // vi tillåter offline-läge
      // lämna UI på local snapshot
    } finally {
      setStatus(s => ({ ...s, loading: false }));
      inFlight.current = false;
    }
  }, [applyShared]);

  // 2) Synka outbox när online: rebase -> save -> retry på 409
  const syncNow = useCallback(async () => {
    if (!navigator.onLine) return;
    const ops = loadOps();
    if (!ops.length) return;
    if (inFlight.current) return;
    inFlight.current = true;

    try {
      for (let attempt = 0; attempt < 8; attempt++) {
        const data = await fetchCloudState();
        const cloudPayload = data.payload ?? { __version: 0 };
        const baseVersion = cloudPayload.__version ?? 0;

        // Rebase: cloud + ops
        const merged = applyOpsToPayload(cloudPayload, ops);
        merged.__version = baseVersion;

        const res = await saveCloudState(merged);

        if (res.status === 409) {
          // Någon annan skrev nyare; loopar och rebasa igen
          continue;
        }

        if (res.ok) {
          const json = await res.json().catch(() => null);
          if (json?.version !== undefined) cloudVersionRef.current = json.version;

          // Nu är ops inkorporerade i molnstate => töm outbox
          clearOps();
          refreshPendingCount();

          // Uppdatera UI med merged (så vi speglar exakt vad vi skickade)
          applyShared({ ...merged, __version: json?.version ?? merged.__version });
          setStatus(s => ({ ...s, lastSync: new Date().toISOString() }));
        }
        break;
      }
    } catch {
      // offline/glapp -> ops ligger kvar, försök igen senare
    } finally {
      inFlight.current = false;
    }
  }, [applyShared]);

  // Initial hydrering + polling + online-event
  useEffect(() => {
    hydrateFromCloud();
    const onOnline = () => syncNow();
    window.addEventListener("online", onOnline);

    const poll = window.setInterval(() => {
      // om vi är online och det finns ops, försök sync
      if (navigator.onLine && loadOps().length) syncNow();
    }, 5000);

    return () => {
      window.removeEventListener("online", onOnline);
      window.clearInterval(poll);
    };
  }, [hydrateFromCloud, syncNow]);

  return {
    status,
    hydrateFromCloud,
    syncNow,
    hydratedRef,        // om du vill veta när vi försökt hydrera
    cloudVersionRef,    // senaste kända version
    refreshPendingCount,
  };
}
