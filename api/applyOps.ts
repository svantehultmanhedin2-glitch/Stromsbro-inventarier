// applyOps.ts
import type { Op } from "./outbox";

function ensureArray(obj: any, k: string) {
  if (!Array.isArray(obj[k])) obj[k] = [];
}

function getId(x: any) {
  // Anpassa om du har andra id-fält.
  return x?.id;
}

export function applyOpsToPayload(payload: any, ops: Op[]) {
  const next = structuredClone(payload ?? {});
  // Se till att listor finns
  ensureArray(next, "onskemal");
  ensureArray(next, "produkter");
  ensureArray(next, "inkopslista");
  ensureArray(next, "historik");
  ensureArray(next, "lagLager");
  ensureArray(next, "users");

  const arrByEntity: Record<string, any[]> = {
    onskemal: next.onskemal,
    produkter: next.produkter,
    inkopslista: next.inkopslista,
    historik: next.historik,
    lagLager: next.lagLager,
    users: next.users,
  };

  for (const op of ops) {
    const arr = arrByEntity[op.entity];
    if (!arr) continue;

    if (op.kind === "add") {
      const id = getId(op.item);
      if (id == null) {
        arr.unshift(op.item);
      } else if (!arr.some(x => getId(x) === id)) {
        arr.unshift(op.item);
      }
      continue;
    }

    if (op.kind === "upsert") {
      const id = getId(op.item);
      if (id == null) {
        arr.unshift(op.item);
      } else {
        const idx = arr.findIndex(x => getId(x) === id);
        if (idx >= 0) arr[idx] = { ...arr[idx], ...op.item };
        else arr.unshift(op.item);
      }
      continue;
    }

    if (op.kind === "patch") {
      const idx = arr.findIndex(x => getId(x) === op.itemId);
      if (idx >= 0) arr[idx] = { ...arr[idx], ...op.patch };
      continue;
    }

    if (op.kind === "remove") {
      const idx = arr.findIndex(x => getId(x) === op.itemId);
      if (idx >= 0) arr.splice(idx, 1);
      continue;
    }
  }

  return next;
}