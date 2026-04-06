import { kv } from '@vercel/kv';

export default async function handler(req, res) {
  if (req.method !== 'GET') {
    return res.status(405).json({ error: 'Only GET allowed' });
  }

  const items = await kv.lrange('lager:logg', 0, 1000);

  return res.status(200).json({ items });
}
