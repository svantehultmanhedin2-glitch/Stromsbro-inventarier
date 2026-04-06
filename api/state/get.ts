import { createClient } from "@vercel/kv";

const kv = createClient({
  url: process.env.KV_REST_API_URL || process.env.UPSTASH_REDIS_REST_URL!,
  token: process.env.KV_REST_API_TOKEN || process.env.UPSTASH_REDIS_REST_TOKEN!,
});

const KEY = "stromsbro:state:v1";

export default {
  async fetch(request: Request) {
    if (request.method !== "GET") {
      return new Response(JSON.stringify({ ok: false, error: "Only GET allowed" }), {
        status: 405,
        headers: { "content-type": "application/json" },
      });
    }

    try {
      const payload: any = await kv.get(KEY);

      // Anti-cache: se till att browsers/proxies inte återanvänder gammalt svar
      return new Response(
        JSON.stringify({
          ok: true,
          payload: payload ?? null,
          version: payload?.__version ?? 0,
          updatedAt: payload?.__updatedAt ?? "",
        }),
        {
          status: 200,
          headers: {
            "content-type": "application/json",
            "cache-control": "no-store, no-cache, must-revalidate, proxy-revalidate",
            pragma: "no-cache",
            expires: "0",
          },
        }
      );
    } catch (e) {
      return new Response(JSON.stringify({ ok: false, error: String(e) }), {
        status: 500,
        headers: { "content-type": "application/json" },
      });
    }
  },
};