// cloudApi.ts
export const CLOUD_GET_URL = "/api/state/get";
export const CLOUD_SET_URL = "/api/state/set";

export async function fetchCloudState() {
  const res = await fetch(`${CLOUD_GET_URL}?t=${Date.now()}`, { cache: "no-store" });
  if (!res.ok) throw new Error(`Cloud GET failed: ${res.status}`);
  const data = await res.json();
  if (!data?.ok) throw new Error(`Cloud GET not ok`);
  return data; // { ok, payload, version, updatedAt }
}

export async function saveCloudState(payload: any) {
  const res = await fetch(CLOUD_SET_URL, {
    method: "POST",
    headers: { "content-type": "application/json" },
    cache: "no-store",
    body: JSON.stringify({ payload }),
  });
  return res; // 200 eller 409 eller annat
}