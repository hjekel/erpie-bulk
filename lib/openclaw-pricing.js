'use strict';

const { execFile } = require('child_process');
const { promisify } = require('util');
const execFileAsync = promisify(execFile);

// Resolve OpenClaw binary
const OPENCLAW_BIN = process.env.OPENCLAW_PATH
  || '/home/openclaw/.npm-global/bin/openclaw';

const TIMEOUT_MS = 90_000; // 90s for large lots

/**
 * Call OpenClaw to price a batch of device groups.
 * Sends one consolidated prompt with all unique model groups.
 * Returns parsed pricing results.
 *
 * @param {Array} deviceGroups - Array of { brand, model, cpu, ram, ssd, grade, qty }
 * @param {string} region - EU/UK/INTL
 * @param {string} dealName - Deal name for context
 * @returns {Promise<{ results: Array, raw: string } | null>} Parsed results or null on failure
 */
async function priceViaOpenClaw(deviceGroups, region, dealName) {
  // Build the lot description (one line per unique model group)
  const lines = deviceGroups.map(g => {
    const parts = [
      g.brand || '',
      g.model || '',
      g.cpu || '',
      g.ram ? (typeof g.ram === 'number' ? g.ram + 'GB' : g.ram) : '',
      g.ssd ? (typeof g.ssd === 'number' ? g.ssd + 'GB' : g.ssd) : '',
      g.grade ? 'Grade ' + g.grade : '',
      (g.qty || 1) + ' stuks',
    ].filter(Boolean).join(' · ');
    return parts;
  });

  const totalQty = deviceGroups.reduce((s, g) => s + (g.qty || 1), 0);

  const prompt = `Prijs dit lot voor PlanBit (${dealName}, ${region}, ${totalQty} stuks totaal):

${lines.join('\n')}

Geef per model: ERP per stuk en totaal.
Antwoord ALLEEN in dit exacte JSON formaat, geen andere tekst:
{
  "prices": [
    {"model": "...", "erp_per_unit": 123.45, "qty": 10, "total": 1234.50, "status": "GO|WATCH|NO-GO", "gen": "Gen10"},
    ...
  ],
  "total_value": 12345.67,
  "lot_factor": 0.92
}`;

  const args = [
    'agent', '--agent', 'main',
    '-m', prompt,
    '--json',
  ];

  console.log(`[openclaw-pricing] Sending ${deviceGroups.length} groups (${totalQty} units) to OpenClaw...`);

  try {
    const { stdout, stderr } = await execFileAsync(OPENCLAW_BIN, args, {
      timeout: TIMEOUT_MS,
      env: {
        ...process.env,
        PATH: `${process.env.PATH || ''}:/home/openclaw/.npm-global/bin:/usr/local/bin`,
        ANTHROPIC_API_KEY: process.env.ANTHROPIC_API_KEY || '',
      },
    });

    if (stderr) console.warn('[openclaw-pricing] stderr:', stderr.slice(0, 200));

    // Parse OpenClaw JSON envelope
    let parsed;
    try {
      parsed = JSON.parse(stdout);
    } catch (e) {
      console.error('[openclaw-pricing] JSON parse of OpenClaw output failed:', e.message);
      console.error('[openclaw-pricing] stdout:', stdout.slice(0, 500));
      return null;
    }

    // Extract text from OpenClaw response
    const text = parsed.result?.payloads?.[0]?.text
              || parsed.result?.text
              || (typeof parsed.result === 'string' ? parsed.result : null)
              || parsed.text
              || parsed.output;

    if (!text) {
      console.error('[openclaw-pricing] No text in OpenClaw response. Keys:', Object.keys(parsed));
      return null;
    }

    console.log('[openclaw-pricing] Got response:', text.length, 'chars');

    // Extract JSON from the response text (may be wrapped in markdown code blocks)
    const jsonMatch = text.match(/```(?:json)?\s*([\s\S]*?)```/) || text.match(/(\{[\s\S]*"prices"[\s\S]*\})/);
    const jsonStr = jsonMatch ? jsonMatch[1].trim() : text.trim();

    let priceData;
    try {
      priceData = JSON.parse(jsonStr);
    } catch (e) {
      console.error('[openclaw-pricing] Failed to parse pricing JSON from response:', e.message);
      console.error('[openclaw-pricing] Text was:', jsonStr.slice(0, 500));
      return null;
    }

    if (!priceData.prices || !Array.isArray(priceData.prices)) {
      console.error('[openclaw-pricing] No prices array in response');
      return null;
    }

    console.log(`[openclaw-pricing] Parsed ${priceData.prices.length} price entries, total: €${priceData.total_value || '?'}`);

    return {
      results: priceData.prices,
      totalValue: priceData.total_value || 0,
      lotFactor: priceData.lot_factor || 1.0,
      raw: text,
    };
  } catch (err) {
    if (err.killed) {
      console.error('[openclaw-pricing] OpenClaw timed out after', TIMEOUT_MS, 'ms');
    } else if (err.code === 'ENOENT') {
      console.error('[openclaw-pricing] OpenClaw binary not found at:', OPENCLAW_BIN);
    } else {
      console.error('[openclaw-pricing] Error:', err.message);
    }
    return null;
  }
}

/**
 * Map OpenClaw results back to the device array.
 * Matches by model name (fuzzy) and enriches each device with the OpenClaw price.
 *
 * @param {Array} devices - Original parsed devices (one per asset)
 * @param {Object} openclawResult - Result from priceViaOpenClaw()
 * @returns {Object} { results, summary } in the same shape as analyzeDevices()
 */
function mapOpenClawResults(devices, openclawResult) {
  const priceMap = new Map();
  for (const p of openclawResult.results) {
    const key = (p.model || '').toLowerCase().replace(/\s+/g, ' ').trim();
    priceMap.set(key, p);
  }

  // Group devices by model for matching
  const groups = new Map();
  for (const d of devices) {
    const key = (d.model || '').toLowerCase().replace(/\s+/g, ' ').trim();
    if (!groups.has(key)) {
      groups.set(key, { device: d, qty: 0 });
    }
    groups.get(key).qty += (d.qty || 1);
  }

  const results = [];
  let totalValue = 0;
  let goCount = 0, watchCount = 0, nogoCount = 0;

  for (const [key, group] of groups) {
    const d = group.device;
    const qty = group.qty;

    // Find matching OpenClaw price (try exact, then fuzzy substring)
    let price = priceMap.get(key);
    if (!price) {
      for (const [pk, pv] of priceMap) {
        if (key.includes(pk) || pk.includes(key)) {
          price = pv;
          break;
        }
      }
    }

    const erp = price?.erp_per_unit || 0;
    const total = erp * qty;
    const status = price?.status || (erp > 50 ? 'GO' : erp > 10 ? 'WATCH' : 'NO-GO');
    const gen = price?.gen || d.gen || '';

    totalValue += total;
    if (status === 'GO') goCount += qty;
    else if (status === 'WATCH') watchCount += qty;
    else nogoCount += qty;

    const ramGb = typeof d.ram === 'number' ? d.ram :
      (typeof d.ram === 'string' ? parseInt(d.ram) || 8 : 8);
    const ssdGb = typeof d.ssd === 'number' ? d.ssd :
      (typeof d.ssd === 'string' ? parseInt(d.ssd) || 256 : 256);

    results.push({
      model: d.model,
      brand: d.brand || '',
      gen,
      status,
      advisedPrice: Math.round(erp),
      priceLow: Math.round(erp * 0.85),
      priceHigh: Math.round(erp * 1.15),
      ramGb: ramGb > 0 ? ramGb : 8,
      ssdGb: ssdGb > 0 ? ssdGb : 256,
      grade: d.grade || 'B',
      region: d.region || 'EU',
      qty,
      cpu: d.cpu || '',
      reasoning: price ? ['Priced via OpenClaw ERPIE agent'] : ['No OpenClaw match — price = 0'],
      pricingSource: 'openclaw',
    });
  }

  const totalQty = results.reduce((s, r) => s + (r.qty || 1), 0);
  const avgValue = totalQty > 0 ? Math.round(totalValue / totalQty) : 0;
  const goPct = totalQty > 0 ? goCount / totalQty : 0;

  const summary = {
    total: totalQty,
    totalGroups: results.length,
    goCount,
    watchCount,
    nogoCount,
    totalValue: Math.round(totalValue),
    avgValue,
    bidLow: Math.round(totalValue * 0.70),
    bidHigh: Math.round(totalValue * 0.85),
    recommendation: goPct >= 0.7
      ? 'Strong portfolio — majority of devices have good resale value.'
      : goPct >= 0.4
        ? 'Mixed portfolio — some value devices, consider cherry-picking.'
        : 'Weak portfolio — limited resale value, consider recycling older models.',
    pricingSource: 'openclaw',
  };

  return { results, summary };
}

/**
 * Group devices by brand+model for sending to OpenClaw
 */
function groupDevicesForPricing(devices) {
  const groups = new Map();
  for (const d of devices) {
    const key = `${(d.brand || '').toLowerCase()}|${(d.model || '').toLowerCase()}`;
    if (!groups.has(key)) {
      groups.set(key, {
        brand: d.brand || '',
        model: d.model || '',
        cpu: d.cpu || '',
        ram: d.ram || '',
        ssd: d.ssd || '',
        grade: d.grade || 'B',
        qty: 0,
      });
    }
    const g = groups.get(key);
    g.qty += (d.qty || 1);
    // Keep the best specs (non-empty wins)
    if (!g.cpu && d.cpu) g.cpu = d.cpu;
    if (!g.ram && d.ram) g.ram = d.ram;
    if (!g.ssd && d.ssd) g.ssd = d.ssd;
  }
  return [...groups.values()];
}

module.exports = { priceViaOpenClaw, mapOpenClawResults, groupDevicesForPricing };
