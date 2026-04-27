import { createClient } from "@vercel/kv";

const kv = createClient({
  url: process.env.KV_REST_API_URL || process.env.UPSTASH_REDIS_REST_URL!,
  token: process.env.KV_REST_API_TOKEN || process.env.UPSTASH_REDIS_REST_TOKEN!,
});

const KEY = "stromsbro:state:v1";

export default {
  async fetch(request) {
    if (request.method !== "GET") {
      return new Response(JSON.stringify({ ok: false, error: "Only GET allowed" }), {
        status: 405,
        headers: { "content-type": "application/json" },
      });
    }

    try {
      let payload = await kv.get(KEY);

      // ✅ EXTREMT VIKTIGT: normalisera payload
      if (!payload || typeof payload !== "object") {
        payload = {
          users: [],
          produkter: [],
          inkopslista: [],
          onskemal: [],
          historik: [],
          lagLager: [],
          __version: 0,
          __updatedAt: "",
        };
      }

      // ✅ skydda mot trasiga versioner
      if (typeof payload.__version !== "number") payload.__version = 0;
      if (typeof payload.__updatedAt !== "string") payload.__updatedAt = "";

      return new Response(
        JSON.stringify({
          ok: true,
          payload,
          version: payload.__version,
          updatedAt: payload.__updatedAt,
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
      // ✅ KRITISKT: skicka tillbaka FELET så vi kan se det
      return new Response(
        JSON.stringify({
          ok: false,
          error: String(e),
        }),
        {
          status: 500,
          headers: { "content-type": "application/json" },
        }
      );
    }
  },
};