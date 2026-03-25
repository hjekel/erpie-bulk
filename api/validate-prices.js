'use strict';

const { searchBrave, searchBraveB2C } = require('../lib/web-search');

/**
 * Calculate confidence based on delta between calculated and live price
 */
function calculateConfidence(calculatedERP, livePrice) {
  if (!livePrice || !calculatedERP) return 'NONE';
  const delta = Math.abs(calculatedERP - livePrice);
  const pct = delta / calculatedERP;
  if (pct <= 0.20) return 'HIGH';
  if (pct <= 0.40) return 'MEDIUM';
  return 'LOW';
}

/**
 * Get median value from array of numbers
 */
function median(arr) {
  if (!arr.length) return null;
  const sorted = [...arr].sort((a, b) => a - b);
  const mid = Math.floor(sorted.length / 2);
  return sorted.length % 2 ? sorted[mid] : Math.round((sorted[mid - 1] + sorted[mid]) / 2);
}

module.exports = async function handler(req, res) {
  // CORS
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
  if (req.method === 'OPTIONS') return res.status(200).end();
  if (req.method !== 'POST') return res.status(405).json({ ok: false, error: 'Method not allowed' });

  try {
    const { devices = [] } = req.body || {};

    if (!devices.length) {
      return res.status(400).json({ ok: false, error: 'No devices to validate' });
    }

    if (!process.env.BRAVE_API_KEY) {
      return res.status(503).json({
        ok: false,
        error: 'Live price validation not configured — BRAVE_API_KEY missing',
      });
    }

    // Deduplicate by model name
    const uniqueModels = new Map();
    for (const d of devices) {
      const key = (d.model || '').toLowerCase().trim();
      if (key && !uniqueModels.has(key)) {
        uniqueModels.set(key, d);
      }
    }

    const validations = [];

    for (const [key, device] of uniqueModels) {
      const brand = device.brand || '';
      const model = device.model || '';
      const ram = device.ram ? `${device.ram}GB` : '';
      const storage = device.storage ? `${device.storage}GB` : '';

      const b2bQuery = `${brand} ${model} ${ram} ${storage} refurbished site:itgigant.nl OR site:thebrokersite.com`;
      const b2cQuery = `${brand} ${model} ${ram} ${storage} refurbished kopen prijs`;

      const [b2bResult, b2cResult] = await Promise.all([
        searchBrave(b2bQuery),
        searchBraveB2C(b2cQuery),
      ]);

      const livePriceB2B = b2bResult ? median(b2bResult.prices) : null;
      const livePriceB2C = b2cResult ? median(b2cResult.prices) : null;
      const calculatedERP = device.advisedPrice || 0;

      // Use B2B price for confidence calculation
      const livePrice = livePriceB2B || livePriceB2C;
      const delta = livePrice ? livePrice - calculatedERP : null;
      const deltaPercent = (livePrice && calculatedERP)
        ? Math.round(((livePrice - calculatedERP) / calculatedERP) * 100)
        : null;

      validations.push({
        model,
        calculatedERP,
        livePriceB2B,
        livePriceB2C,
        sourceB2B: b2bResult?.source || null,
        sourceB2C: b2cResult?.source || null,
        delta,
        deltaPercent,
        confidence: calculateConfidence(calculatedERP, livePrice),
      });
    }

    // Map validations back to all devices (including duplicates)
    const validationMap = new Map();
    for (const v of validations) {
      validationMap.set(v.model.toLowerCase().trim(), v);
    }

    const fullValidations = devices.map(d => {
      const key = (d.model || '').toLowerCase().trim();
      return validationMap.get(key) || {
        model: d.model,
        calculatedERP: d.advisedPrice || 0,
        livePriceB2B: null,
        livePriceB2C: null,
        sourceB2B: null,
        sourceB2C: null,
        delta: null,
        deltaPercent: null,
        confidence: 'NONE',
      };
    });

    res.status(200).json({ ok: true, validations: fullValidations });
  } catch (err) {
    console.error('[validate-prices] Error:', err);
    res.status(500).json({ ok: false, error: err.message });
  }
};
