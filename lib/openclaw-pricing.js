'use strict';

const { execFile } = require('child_process');
const { promisify } = require('util');
const execFileAsync = promisify(execFile);

// Resolve OpenClaw binary
const OPENCLAW_BIN = process.env.OPENCLAW_PATH
  || '/home/openclaw/.npm-global/bin/openclaw';

const TIMEOUT_MS = 120_000; // 120s for large lots

/**
 * Attempt to parse JSON from an LLM response.
 * Handles: markdown code blocks, trailing commas, single-line comments,
 * unquoted keys, and other common LLM JSON quirks.
 */
function robustJsonParse(text) {
  // 1. Extract from markdown code blocks
  const codeBlock = text.match(/```(?:json)?\s*([\s\S]*?)```/);
  let jsonStr = codeBlock ? codeBlock[1].trim() : text.trim();

  // 2. Find the JSON object/array boundaries
  const firstBrace = jsonStr.indexOf('{');
  const firstBracket = jsonStr.indexOf('[');
  let startChar, endChar;

  if (firstBracket >= 0 && (firstBrace < 0 || firstBracket < firstBrace)) {
    startChar = '['; endChar = ']';
    jsonStr = jsonStr.slice(firstBracket);
  } else if (firstBrace >= 0) {
    startChar = '{'; endChar = '}';
    jsonStr = jsonStr.slice(firstBrace);
  } else {
    throw new Error('No JSON object/array found in response');
  }

  // Find matching closing bracket/brace
  let depth = 0;
  let inString = false;
  let escape = false;
  let endIdx = -1;
  for (let i = 0; i < jsonStr.length; i++) {
    const ch = jsonStr[i];
    if (escape) { escape = false; continue; }
    if (ch === '\\' && inString) { escape = true; continue; }
    if (ch === '"' && !escape) { inString = !inString; continue; }
    if (inString) continue;
    if (ch === startChar) depth++;
    if (ch === endChar) { depth--; if (depth === 0) { endIdx = i; break; } }
  }
  if (endIdx > 0) jsonStr = jsonStr.slice(0, endIdx + 1);

  // 3. Try parsing as-is first
  try { return JSON.parse(jsonStr); } catch (e) { /* continue cleanup */ }

  // 4. Clean up common LLM JSON issues
  let cleaned = jsonStr
    .replace(/\/\/[^\n]*/g, '')                  // remove single-line comments
    .replace(/,\s*([\]}])/g, '$1')               // remove trailing commas
    .replace(/([{,]\s*)(\w+)\s*:/g, '$1"$2":')   // unquoted keys → "key":
    .replace(/:\s*'([^']*)'/g, ': "$1"')          // single-quoted values → double-quoted
    .replace(/\n/g, ' ')                          // collapse newlines
    .trim();

  try { return JSON.parse(cleaned); } catch (e) {
    throw new Error(`JSON parse failed after cleanup: ${e.message}\nFirst 300 chars: ${cleaned.slice(0, 300)}`);
  }
}

/**
 * Call OpenClaw to price a batch of device groups.
 * Uses the exact SOUL.md Eureka Rule Book format.
 *
 * @param {Array} deviceGroups - Array of { brand, model, cpu, ram, ssd, grade, qty }
 * @param {string} region - EU/UK/INTL
 * @param {string} dealName - Deal name for context
 * @returns {Promise<{ results: Array, raw: string } | null>} Parsed results or null on failure
 */
async function priceViaOpenClaw(deviceGroups, region, dealName) {
  const totalQty = deviceGroups.reduce((s, g) => s + (g.qty || 1), 0);

  // Build the lot lines (one per unique model group)
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

  // Prompt matches SOUL.md Eureka Rule Book format exactly
  const prompt = `Je bent ERPIE, de PlanBit pricing expert. Gebruik de Eureka Rule Book regels uit je SOUL.md en TOOLS.md.
Bereken Expected Resale Price (ERP) per stuk voor dit ${region} lot.
Dit zijn B2B verwachte verkoopprijzen zoals PlanBit die hanteert — GEEN inkoopprijzen, GEEN B2C consumentenprijzen.

Regels:
- Gebruik TOOLS.md basisprijzen per CPU generatie
- 0GB SSD of 0GB RAM = specs niet gedetecteerd door Blancco = gebruik standaard (256GB SSD / 8GB RAM)
- Grade multipliers: A=1.05, B=1.00, C=0.40, D=0.15
- Lot grootte ${totalQty} stuks = pas lot-korting toe: 100-200=×0.92, 200-500=×0.85, 500+=×0.80
- Gen8 lot caps: T470=€45, T480=€65, T490=€65, X280=€65
- Apple Silicon: gebruik TOOLS.md Apple Silicon tabel, minimum €650 voor M1
- Gen6 en ouder: max €10/unit
- GEEN notebook factor toepassen (base prices zijn al B2B)

LOT (${region} regio, ${dealName}, ${totalQty} stuks totaal):
${lines.join('\n')}

Antwoord ALLEEN als JSON array, geen andere tekst:
[{"model":"ThinkPad T14 Gen 1","brand":"Lenovo","erpPerUnit":140.22,"qty":98,"total":13741.56,"status":"GO","gen":"Gen10"}]`;

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

    if (stderr) console.warn('[openclaw-pricing] stderr:', stderr.slice(0, 300));

    // Parse OpenClaw JSON envelope
    let parsed;
    try {
      parsed = JSON.parse(stdout);
    } catch (e) {
      console.error('[openclaw-pricing] OpenClaw envelope parse failed:', e.message);
      console.error('[openclaw-pricing] stdout first 500:', stdout.slice(0, 500));
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

    // Parse pricing data with robust JSON parser
    let priceData;
    try {
      priceData = robustJsonParse(text);
    } catch (e) {
      console.error('[openclaw-pricing] JSON extraction failed:', e.message);
      return null;
    }

    // Handle both array format and object-with-prices format
    let prices;
    if (Array.isArray(priceData)) {
      prices = priceData;
    } else if (priceData.prices && Array.isArray(priceData.prices)) {
      prices = priceData.prices;
    } else {
      console.error('[openclaw-pricing] Response is not an array and has no prices field');
      return null;
    }

    // Normalize field names (LLM may use erp_per_unit, erpPerUnit, erp, price, etc.)
    const normalized = prices.map(p => ({
      model: p.model || '',
      brand: p.brand || '',
      erp_per_unit: p.erp_per_unit || p.erpPerUnit || p.erp || p.price || p.unitPrice || 0,
      qty: p.qty || p.quantity || 1,
      total: p.total || p.totalValue || 0,
      status: p.status || 'GO',
      gen: p.gen || p.generation || '',
    }));

    const totalValue = priceData.total_value || priceData.totalValue
      || normalized.reduce((s, p) => s + (p.erp_per_unit * p.qty), 0);

    console.log(`[openclaw-pricing] Parsed ${normalized.length} price entries, total: €${Math.round(totalValue)}`);

    return {
      results: normalized,
      totalValue,
      lotFactor: priceData.lot_factor || priceData.lotFactor || 1.0,
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
      advisedPrice: Math.round(erp * 100) / 100,
      priceLow: Math.round(erp * 0.85 * 100) / 100,
      priceHigh: Math.round(erp * 1.15 * 100) / 100,
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
    totalValue: Math.round(totalValue * 100) / 100,
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
