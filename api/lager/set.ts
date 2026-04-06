
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
      const payload = body?.payload;

      if (!payload || typeof payload !== "object") {
        return new Response(JSON.stringify({ ok: false, error: "Missing payload object" }), {
          status: 400,
          headers: { "content-type": "application/json" },
        });
      }

      const wrapped = {
        ...payload,
        __meta: { updatedAt: new Date().toISOString() },
      };

      await kv.set(KEY, wrapped);

      return new Response(JSON.stringify({ ok: true }), {
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