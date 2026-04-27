// outbox.ts
export type Entity = "onskemal" | "produkter" | "inkopslista" | "historik" | "lagLager" | "users";

export type Op =
  | { id: string; ts: number; kind: "add"; entity: Entity; item: any }
  | { id: string; ts: number; kind: "upsert"; entity: Entity; item: any } // idempotent add/update
  | { id: string; ts: number; kind: "patch"; entity: Entity; itemId: any; patch: any }
  | { id: string; ts: number; kind: "remove"; entity: Entity; itemId: any };

const OPS_KEY = "stromsbro:pendingOps:v1";

export function loadOps(): Op[] {
  try {
    const raw = localStorage.getItem(OPS_KEY);
    const ops = raw ? JSON.parse(raw) : [];
    return Array.isArray(ops) ? ops : [];
  } catch {
    return [];
  }
}

export function saveOps(ops: Op[]) {
  try {
    localStorage.setItem(OPS_KEY, JSON.stringify(ops));
  } catch {}
}

export function enqueueOp(op: Op) {
  const ops = loadOps();
  ops.push(op);
  saveOps(ops);
}

export function clearOps() {
  saveOps([]);
}

export function removeOpsByIds(ids: string[]) {
  const set = new Set(ids);
  const ops = loadOps().filter(o => !set.has(o.id));
  saveOps(ops);
}

export function uuid() {
  return globalThis.crypto?.randomUUID
    ? globalThis.crypto.randomUUID()
    : `${Date.now()}-${Math.random().toString(16).slice(2)}`;
}
