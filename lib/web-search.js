'use strict';

// ═══════════════════════════════════════════════════════════════════════════
// BRAVE SEARCH WRAPPER — Rate-limited, cached, price-extracting
// ═══════════════════════════════════════════════════════════════════════════

const BRAVE_ENDPOINT = 'https://api.search.brave.com/res/v1/web/search';
const CACHE_TTL_MS = 60 * 60 * 1000; // 1 hour
const MIN_REQUEST_GAP_MS = 1000;      // 1 request per second

const cache = new Map();
let lastRequestTime = 0;

/**
 * Sleep for a given number of milliseconds
 */
function sleep(ms) {
  return new Promise(resolve => setTimeout(resolve, ms));
}

/**
 * Extract prices from search result snippets
 * Looks for patterns like: €199, € 249,00, 319 euro, €149,-
 * @param {string} text
 * @returns {number[]}
 */
function extractPrices(text) {
  const prices = [];
  const patterns = [
    /€\s*(\d{1,5})[.,](\d{2})/g,           // €199,00 or €199.00
    /€\s*(\d{1,5})(?:[,\-]{1,2})?(?=\s|$)/g, // €199 or €199,- or €199,-
    /(\d{1,5})[.,](\d{2})\s*euro/gi,        // 199,00 euro
    /(\d{1,5})\s*euro/gi,                    // 199 euro
  ];

  for (const pattern of patterns) {
    let match;
    while ((match = pattern.exec(text)) !== null) {
      const whole = parseInt(match[1]);
      const cents = match[2] ? parseInt(match[2]) : 0;
      const price = whole + cents / 100;
      if (price >= 10 && price <= 5000) {
        prices.push(Math.round(price));
      }
    }
  }

  // Deduplicate
  return [...new Set(prices)];
}

/**
 * Search Brave for refurbished prices of a device model
 * @param {string} query - Search query
 * @returns {Promise<{prices: number[], source: string, snippet: string} | null>}
 */
async function searchBrave(query) {
  const apiKey = process.env.BRAVE_API_KEY;
  if (!apiKey) {
    console.warn('[web-search] BRAVE_API_KEY not set — skipping live search');
    return null;
  }

  // Check cache
  const cached = cache.get(query);
  if (cached && Date.now() - cached.timestamp < CACHE_TTL_MS) {
    return cached.result;
  }

  // Rate limit
  const now = Date.now();
  const elapsed = now - lastRequestTime;
  if (elapsed < MIN_REQUEST_GAP_MS) {
    await sleep(MIN_REQUEST_GAP_MS - elapsed);
  }

  try {
    lastRequestTime = Date.now();

    const url = new URL(BRAVE_ENDPOINT);
    url.searchParams.set('q', query);
    url.searchParams.set('count', '10');
    url.searchParams.set('search_lang', 'nl');

    const response = await fetch(url.toString(), {
      headers: {
        'Accept': 'application/json',
        'Accept-Encoding': 'gzip',
        'X-Subscription-Token': apiKey,
      },
    });

    if (!response.ok) {
      console.error(`[web-search] Brave API ${response.status}: ${response.statusText}`);
      return null;
    }

    const data = await response.json();
    const webResults = data.web?.results || [];

    // Collect prices and sources from all results
    const allPrices = [];
    let bestSource = '';
    let bestSnippet = '';

    for (const result of webResults) {
      const text = `${result.title || ''} ${result.description || ''}`;
      const prices = extractPrices(text);

      if (prices.length > 0) {
        allPrices.push(...prices);
        if (!bestSource) {
          bestSource = result.url || '';
          bestSnippet = (result.description || '').slice(0, 200);
        }
      }
    }

    const resultObj = allPrices.length > 0
      ? { prices: [...new Set(allPrices)], source: bestSource, snippet: bestSnippet }
      : null;

    // Cache result (even null to avoid repeated failed searches)
    cache.set(query, { result: resultObj, timestamp: Date.now() });

    return resultObj;
  } catch (err) {
    console.error('[web-search] Error:', err.message);
    return null;
  }
}

/**
 * Clear expired cache entries
 */
function cleanCache() {
  const now = Date.now();
  for (const [key, val] of cache) {
    if (now - val.timestamp > CACHE_TTL_MS) {
      cache.delete(key);
    }
  }
}

// Clean cache every 10 minutes
setInterval(cleanCache, 10 * 60 * 1000).unref();

module.exports = { searchBrave, extractPrices };
