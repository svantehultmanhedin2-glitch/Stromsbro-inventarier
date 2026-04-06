import type { NextApiRequest, NextApiResponse } from 'next';
import { kv } from '@vercel/kv';

// Typ för inkommande payload
type SaveBody = {
  typ: 'Inleverans' | 'Uttag' | 'Justering';
  produkt: string;
  antal: number;
  lagerplats?: string;
  kommentar?: string;
};

// Typ för sparad post
type LagerPost = SaveBody & {
  id: string;
  tid: string; // ISO‑tid
};

export default async function handler(
  req: NextApiRequest,
  res: NextApiResponse
) {
  // Tillåt endast POST
  if (req.method !== 'POST') {
    return res.status(405).json({ error: 'Only POST allowed' });
  }

  const { typ, produkt, antal, lagerplats, kommentar } = req.body as SaveBody;

  // Enkel validering
  if (!typ || !produkt || typeof antal !== 'number') {
    return res.status(400).json({
      error: 'Invalid payload. Required: typ, produkt, antal',
    });
  }

  // Skapa posten som ska sparas
  const post: LagerPost = {
    id: crypto.randomUUID(),
    tid: new Date().toISOString(),
    typ,
    produkt,
    antal,
    lagerplats,
    kommentar,
  };

  try {
    /**
     * Vi använder en Redis LIST:
     * - Key: "lager:logg"
     * - LPUSH → snabb, atomär, rätt för loggar
     */
    await kv.lpush('lager:logg', post);

    return res.status(200).json({
      ok: true,
      post,
    });
  } catch (err) {
    console.error('KV save error:', err);
    return res.status(500).json({
      error: 'Failed to save data',
    });
  }
}