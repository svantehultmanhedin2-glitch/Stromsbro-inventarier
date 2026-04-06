import { createClient } from "@vercel/kv";

const kv = createClient({
  url: process.env.KV_REST_API_URL || process.env.UPSTASH_REDIS_REST_URL!,
  token: process.env.KV_REST_API_TOKEN || process.env.UPSTASH_REDIS_REST_TOKEN!,
});

const KEY = "stromsbro:state:v1";

export default {
  async fetch(request: Request) {
    if (request.method !== "POST") {
      return new Response(JSON.stringify({ ok: false, error: "Only POST allowed" }), {
        status: 405,
        headers: { "content-type": "application/json" },
      });
    }

    try {
      const body = await request.json().catch(() => ({}));
      const incoming = body?.payload;

      if (!incoming || typeof incoming !== "object") {
        return new Response(JSON.stringify({ ok: false, error: "Missing payload" }), {
          status: 400,
          headers: { "content-type": "application/json" },
        });
      }

      const current: any = await kv.get(KEY);
      const currentVersion = current?.__version ?? 0;
      const incomingVersion = incoming.__version ?? 0;

      // ❌ Stoppa gammal enhet från att skriva över nyare data
      if (incomingVersion !== currentVersion) {
        return new Response(
          JSON.stringify({
            ok: false,
            conflict: true,
            currentVersion,
            currentUpdatedAt: current?.__updatedAt ?? "",
          }),
          {
            status: 409,
            headers: {
              "content-type": "application/json",
              "cache-control": "no-store",
            },
          }
        );
      }

      const nextVersion = currentVersion + 1;

      await kv.set(KEY, {
        ...incoming,
        __version: nextVersion,
        __updatedAt: new Date().toISOString(),
      });

      return new Response(JSON.stringify({ ok: true, version: nextVersion }), {
        status: 200,
        headers: {
          "content-type": "application/json",
          "cache-control": "no-store",
        },
      });
    } catch (e) {
      return new Response(JSON.stringify({ ok: false, error: String(e) }), {
        status: 500,
        headers: { "content-type": "application/json" },
      });
    }
  },
};