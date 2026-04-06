
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
      const payload = await kv.get(KEY);
      return new Response(JSON.stringify({ ok: true, payload: payload ?? null }), {
        status: 200,
        headers: { "content-type": "application/json" },
      });
    } catch (e) {
      return new Response(JSON.stringify({ ok: false, error: String(e) }), {
        status: 500,
        headers: { "content-type": "application/json" },
      });
    }
  },
};