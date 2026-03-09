/**
 * Vercel serverless API for persisting dashboard data in Upstash Redis.
 * GET: Returns cached data
 * POST: Saves full dashboard state (replaces existing)
 * DELETE: Clears cached data
 *
 * Environment variables (set in Vercel):
 *   UPSTASH_REDIS_REST_URL or KV_REST_API_URL
 *   UPSTASH_REDIS_REST_TOKEN or KV_REST_API_TOKEN
 */
import { Redis } from '@upstash/redis';

const REDIS_KEY = 'abm:vehicle-services:state';

function getRedis() {
  const url = process.env.UPSTASH_REDIS_REST_URL || process.env.KV_REST_API_URL;
  const token = process.env.UPSTASH_REDIS_REST_TOKEN || process.env.KV_REST_API_TOKEN;
  if (!url || !token) {
    throw new Error('Missing Redis credentials. Set UPSTASH_REDIS_REST_URL and UPSTASH_REDIS_REST_TOKEN in Vercel.');
  }
  return new Redis({ url, token });
}

export default async function handler(req, res) {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET, POST, DELETE, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');

  if (req.method === 'OPTIONS') {
    return res.status(200).end();
  }

  if (req.method !== 'GET' && req.method !== 'POST' && req.method !== 'DELETE') {
    return res.status(405).json({ error: 'Method not allowed' });
  }

  try {
    const redis = getRedis();

    if (req.method === 'GET') {
      const data = await redis.get(REDIS_KEY);
      return res.status(200).json(data || {});
    }

    if (req.method === 'DELETE') {
      await redis.del(REDIS_KEY);
      return res.status(200).json({ ok: true });
    }

    if (req.method === 'POST') {
      const incoming = req.body || {};
      await redis.set(REDIS_KEY, incoming);
      return res.status(200).json(incoming);
    }
  } catch (err) {
    console.error('API error:', err.message);
    return res.status(500).json({ error: err.message || 'Server error' });
  }
}
