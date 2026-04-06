import { kv } from "@vercel/kv";

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