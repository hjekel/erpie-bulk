'use strict';

const { parseText } = require('../lib/parser');
const { analyzeDevices } = require('../lib/pricing-engine');
const { priceViaOpenClaw, mapOpenClawResults, groupDevicesForPricing } = require('../lib/openclaw-pricing');

module.exports = async function handler(req, res) {
  // CORS
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
  if (req.method === 'OPTIONS') return res.status(200).end();
  if (req.method !== 'POST') return res.status(405).json({ ok: false, error: 'Method not allowed' });

  try {
    const { text = '', dealName = 'Unnamed Deal', region = 'EU', liveValidation = false } = req.body || {};

    if (!text.trim()) {
      return res.status(400).json({ ok: false, error: 'No text provided' });
    }

    const t0 = Date.now();
    const { devices, format } = parseText(text);

    if (!devices || !devices.length) {
      return res.status(400).json({
        ok: false,
        error: 'Geen apparaten herkend in de tekst. Probeer bijv: "60x Dell Latitude 7430 i7 32GB 512GB"',
      });
    }

    // Try OpenClaw first, fallback to local
    let analysis;
    let pricingSource = 'local';
    const groups = groupDevicesForPricing(devices);
    const openclawResult = await priceViaOpenClaw(groups, region, dealName);

    if (openclawResult && openclawResult.results.length > 0) {
      analysis = mapOpenClawResults(devices, openclawResult);
      pricingSource = 'openclaw';
    } else {
      analysis = analyzeDevices(devices, region);
    }

    const processingTimeMs = Date.now() - t0;

    res.status(200).json({
      ok: true,
      dealName,
      format,
      liveValidation,
      processingTimeMs,
      pricingSource,
      ...analysis,
    });
  } catch (err) {
    console.error('[analyze-text] Error:', err);
    res.status(500).json({ ok: false, error: err.message });
  }
};
