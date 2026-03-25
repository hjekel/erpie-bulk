'use strict';

const XLSX = require('xlsx');

// ═══════════════════════════════════════════════════════════════════════════
// UNIVERSAL DEVICE PARSER v2.0
// Handles: ARS Excel, PWC/Vendor Quote, Zones Inventory, Generic Headers,
//          Generic Headerless, Text Input
// ═══════════════════════════════════════════════════════════════════════════

const BRANDS = ['apple','dell','hp','lenovo','microsoft','asus','acer',
                'fujitsu','toshiba','samsung','sony','lg','panasonic',
                'huawei','xiaomi','google','razer','msi','gigabyte'];

const BRAND_MODELS = {
  apple:     ['macbook','imac','mac mini','mac pro','mac studio','iphone','ipad','ipod'],
  dell:      ['latitude','inspiron','xps','precision','vostro','optiplex','alienware'],
  hp:        ['elitebook','probook','zbook','pavilion','envy','spectre','omen',
               'elitedesk','prodesk','z-workstation','chromebook','folio','revolve'],
  lenovo:    ['thinkpad','thinkcentre','ideapad','legion','yoga','tab'],
  microsoft: ['surface'],
};

const PRODUCT_TYPES = new Set([
  'notebook','laptop','desktop','allinone','all-in-one','workstation',
  'mobilephone','mobile','tablet','server','monitor',
  'mobilephone','allinone',
]);

function cellStr(v) { return String(v ?? '').trim(); }

function normalizeProductType(s) {
  return s.toLowerCase().replace(/[\s\-\/]/g,'').replace('allone','allinone');
}

function isProductType(v) {
  return PRODUCT_TYPES.has(normalizeProductType(cellStr(v)));
}

function looksLikeBrand(v) {
  const s = cellStr(v).toLowerCase();
  return BRANDS.some(b => s === b || s.startsWith(b));
}

function looksLikeModel(v) {
  const s = cellStr(v).toLowerCase();
  if (s.length < 3) return false;
  for (const brand of BRANDS) {
    if (s.includes(brand)) return true;
  }
  const modelWords = ['latitude','elitebook','probook','thinkpad','thinkcentre',
    'macbook','surface','optiplex','inspiron','precision','vostro','zbook',
    'yoga','ideapad','legion','spectre','envy','pavilion','folio','revolve'];
  return modelWords.some(m => s.includes(m));
}

// ─── SPEC EXTRACTION ────────────────────────────────────────────────────────

function extractRAM(s) {
  const m = s.match(/\b(4|8|16|32|64|128)\s*GB(?!\s*SSD)/i);
  return m ? m[1] + 'GB' : '';
}

function extractSSD(s) {
  const KNOWN_SSD_SIZES = new Set([120,128,240,256,480,512,960,1024,2048]);
  const ssdM = s.match(/(?:SSD\s*)?(\d+)\s*(GB|TB)\s*(?:SSD|M\.?2|NVMe|SATA)?/gi);
  if (!ssdM) return '';
  for (const match of ssdM) {
    const numM = match.match(/(\d+)\s*(GB|TB)/i);
    if (!numM) continue;
    const num = parseInt(numM[1]);
    const unit = numM[2].toUpperCase();
    const gb = unit === 'TB' ? num * 1024 : num;
    if (KNOWN_SSD_SIZES.has(gb)) return gb + 'GB';
    if (gb >= 600 && gb < 1000) {
      if (/SSD|NVMe|M\.?2|SATA/i.test(match)) return gb + 'GB';
      continue;
    }
    if (gb >= 100) return gb + 'GB';
  }
  return '';
}

function extractCPU(s) {
  const patterns = [
    /\b(Core\s*)?[iI][3579][-\s]?\d{4,5}[A-Z0-9]*/,
    /\bRyzen\s*[3579]\s*\d{4}[A-Z0-9]*/i,
    /\bM[123]\s*(Pro|Max|Ultra)?/,
    /\bCeleron\s*[A-Z0-9]+/i,
    /\bPentium\s*[A-Z0-9]+/i,
    /\bXeon\s*[A-Z0-9\-]+/i,
  ];
  for (const re of patterns) {
    const m = s.match(re);
    if (m) return m[0].trim();
  }
  const genHint = s.match(/[iI]([3579])[-\s]?(\d+)(th|rd|nd|st)/i);
  if (genHint) return `i${genHint[1]}-${genHint[2]}th`;
  return '';
}

function extractGrade(s) {
  const m = s.match(/\b(A1|A2|A3|A|B1|B2|B3|B4|B|C1|C2|C|D)\b/);
  return m ? m[1] : '';
}

// ─── UNIFIED COLUMN RESOLVER ─────────────────────────────────────────────────
const COLUMN_ALIASES = {
  brand:    ['brand','manufacturer','make','merk','vendor','oem','fabrikant'],
  model:    ['model','device','modelname','computername','description','productname','name','item','assetname','partnumber','sku'],
  cpu:      ['cpu','processor','proc','cputype','cpumodel','chipset','specs'],
  ram:      ['ram','memory','mem','ramsizeoutgoing','systemmemory','installedmemory','ramsize'],
  ssd:      ['ssd','storage','hdd','disk','drive','hddsizeoutgoing','capacity','disksize','hddsize','storagesize'],
  grade:    ['grade','condition','quality','cosmeticcategory','functionalgrade','cosmeticgrade','gradein'],
  serial:   ['serial','serialnumber','sn','assettag','asset','devicetag','equipmentid','uid','serialnumberasset'],
  qty:      ['qty','quantity','count','units','aantal'],
  type:     ['producttype','producttyp','productype'],
  category: ['modelcategory','category','devicetype','assettype'],
  keyboard: ['keyboard','kb','toetsenbord','layout'],
  location: ['location','country','region','site','locatie','locationdisplay'],
};

function resolveColumns(headerRow) {
  const cells = headerRow.map(v => cellStr(v).toLowerCase().replace(/[\s_\-\.]/g, ''));
  const map = {};
  for (const [field, aliases] of Object.entries(COLUMN_ALIASES)) {
    const idx = cells.findIndex(c => aliases.includes(c) || aliases.some(a => c === a));
    map[field] = idx;
  }
  return map;
}

function hasColumns(colMap, ...fields) {
  return fields.every(f => colMap[f] >= 0);
}

// ─── ADAPTIVE ROW FILTER ─────────────────────────────────────────────────────
const DEVICE_TYPES = new Set(['notebook','laptop','desktop','workstation','server','tablet','allinone','computer']);
const SKIP_TYPES = new Set(['monitor','flatpanel','mobilephone','phone','ipphone','printer','network','docking','cable','peripheral','accessory','other']);

function isDeviceRow(row, colMap) {
  if (colMap.type >= 0) {
    const t = cellStr(row[colMap.type]).toLowerCase().replace(/[\s\-\/]/g, '');
    if (DEVICE_TYPES.has(t)) return true;
    if (SKIP_TYPES.has(t)) return false;
  }
  if (colMap.category >= 0) {
    const c = cellStr(row[colMap.category]).toLowerCase();
    if (c === 'computer' || c === 'laptop' || c === 'notebook') return true;
    if (c === 'monitor' || c === 'ip phone' || c === 'phone' || c === 'printer') return false;
  }
  if (colMap.type >= 0) {
    const t = cellStr(row[colMap.type]);
    if (/EOL/i.test(t) && !/monitor|phone/i.test(t)) {
      if (colMap.category >= 0 && /computer/i.test(cellStr(row[colMap.category]))) return true;
      if (colMap.category < 0) return true;
    }
  }
  if (colMap.type < 0 && colMap.category < 0) return true;
  return false;
}

// ─── MODEL GEN INFERENCE ────────────────────────────────────────────────────
const MODEL_GEN_MAP = {
  '10': 'Gen10', '11': 'Gen10',
  '20': 'Gen11', '21': 'Gen11',
  '30': 'Gen12', '31': 'Gen12',
  '40': 'Gen13', '41': 'Gen13',
  '50': 'Gen14', '51': 'Gen14',
};

function inferGenFromModel(brand, model) {
  const m = model.toUpperCase();

  if (/7400.*2-IN-1/i.test(m)) return 'Gen8';
  if (/E5470/i.test(m)) return 'Gen6';
  if (/7480/i.test(m)) return 'Gen7';
  if (/3380/i.test(m)) return 'Gen7';
  if (/5490/i.test(m)) return 'Gen8';
  if (/7490/i.test(m)) return 'Gen8';
  if (/3590/i.test(m)) return 'Gen8';

  const latMatch = m.match(/(?:LATITUDE|PRECISION)\s*(\d)(\d)(\d{2})\b/i);
  if (latMatch) {
    const suffix = latMatch[3];
    if (MODEL_GEN_MAP[suffix]) return MODEL_GEN_MAP[suffix];
    if (suffix === '00') return 'Gen10';
    if (suffix === '90') return 'Gen8';
    if (suffix === '80') return 'Gen7';
  }

  if (/OPTIPLEX\s*70(\d)0/i.test(m)) {
    const d = parseInt(RegExp.$1);
    return { 5:'Gen7', 6:'Gen8', 7:'Gen9', 8:'Gen10', 9:'Gen11' }[d] || null;
  }

  if (/250\s*G(\d+)/i.test(m)) {
    const g = parseInt(RegExp.$1);
    return { 2:'Gen4', 3:'Gen5', 4:'Gen6', 5:'Gen7', 6:'Gen7', 7:'Gen8', 8:'Gen11', 9:'Gen12', 10:'Gen13' }[g] || null;
  }
  if (/PRO\s*3120/i.test(m)) return 'Gen4';
  if (/PROBOOK\s*6570/i.test(m)) return 'Gen3';
  if (/PAVILION/i.test(m)) return null;
  if (/ELITEBOOK\s*8[345]\d\s*G(\d+)/i.test(m)) {
    const g = parseInt(RegExp.$1);
    return { 3:'Gen5', 4:'Gen6', 5:'Gen7', 6:'Gen8', 7:'Gen10', 8:'Gen11', 9:'Gen12', 10:'Gen13' }[g] || null;
  }
  if (/(?:PROBOOK|ELITEBOOK)\s*\d{3}\s*G(\d+)/i.test(m)) {
    const g = parseInt(RegExp.$1);
    return { 3:'Gen5', 4:'Gen6', 5:'Gen7', 6:'Gen8', 7:'Gen10', 8:'Gen11', 9:'Gen12', 10:'Gen13' }[g] || null;
  }
  if (/THINKPAD.*Gen\s*(\d+)/i.test(m)) {
    const g = parseInt(RegExp.$1);
    return { 1:'Gen10', 2:'Gen11', 3:'Gen12', 4:'Gen13', 5:'Gen14' }[g] || null;
  }
  if (/MACBOOKPRO16[,.]1|macbook.*16.*pro/i.test(m)) return 'Gen9';
  if (/MACBOOKPRO15[,.]1|macbook.*15.*pro.*2018/i.test(m)) return 'Gen8';
  if (/MACBOOKPRO14[,.]1/i.test(m)) return 'Gen7';
  if (/CHROMEBOOK/i.test(m)) return 'Gen8';
  return null;
}

// ─── SMART FORMAT SCORING ────────────────────────────────────────────────────

function detectFormat(wb) {
  const scores = { VENDOR_QUOTE: 0, ARS: 0, ZONES_INVENTORY: 0, GENERIC_HEADERS: 0, ARS_SIMPLE: 0 };

  const hasArsInventory = wb.SheetNames.some(n => {
    const l = n.toLowerCase().trim();
    return l.includes('customer inventory') && !l.includes('and details');
  });
  const hasRitmSheets = wb.SheetNames.some(n => /^RITM/i.test(n.trim()));
  if (hasArsInventory) scores.ARS += 10;
  if (hasRitmSheets) scores.ARS += 5;

  for (const name of wb.SheetNames) {
    const rows = XLSX.utils.sheet_to_json(wb.Sheets[name], { header: 1, defval: '' });
    const nonEmpty = rows.filter(r => r.some(v => cellStr(v).length > 0));
    if (!nonEmpty.length) continue;

    for (let i = 0; i < Math.min(5, nonEmpty.length); i++) {
      const colMap = resolveColumns(nonEmpty[i]);
      const cells = nonEmpty[i].map(v => cellStr(v).toLowerCase());

      const hasProduct = cells.some(c => c === 'product' || c === 'configuration');
      const hasProc = cells.some(c => c.includes('processor') || c === 'p/n' || c.includes('assumed specifications'));
      if (hasProduct && hasColumns(colMap, 'brand', 'model') && hasProc) {
        scores.VENDOR_QUOTE += 12;
      }

      if (hasColumns(colMap, 'type', 'brand', 'model', 'serial') && colMap.cpu < 0 && colMap.ram < 0) {
        const typeHeader = cellStr(nonEmpty[i][colMap.type]).toLowerCase().replace(/[\s_\-]/g, '');
        if (['producttype','producttyp','pat'].includes(typeHeader)) {
          scores.ZONES_INVENTORY += 10;
        }
      }

      if (colMap.model >= 0) scores.GENERIC_HEADERS += 3;
      if (colMap.serial >= 0) scores.GENERIC_HEADERS += 2;
      if (colMap.brand >= 0) scores.GENERIC_HEADERS += 2;
      if (colMap.cpu >= 0) scores.GENERIC_HEADERS += 3;
      if (colMap.ram >= 0) scores.GENERIC_HEADERS += 1;
    }

    const dataRows = nonEmpty.filter(r => isProductType(r[0]));
    if (dataRows.length >= 2) scores.ARS_SIMPLE += 5;
  }

  const best = Object.entries(scores).sort((a, b) => b[1] - a[1])[0];
  const format = best[1] > 0 ? best[0] : 'GENERIC_HEADERLESS';
  return format;
}

// ─── PARSERS ─────────────────────────────────────────────────────────────────

function parseARS(wb) {
  let devices = [];

  const invSheet = wb.SheetNames.find(n => n.toLowerCase().includes('inventory'));
  if (invSheet) {
    const rows = XLSX.utils.sheet_to_json(wb.Sheets[invSheet], { header: 1, defval: '' });
    devices = parseARSRows(rows);
  }

  const quoteSheet = wb.SheetNames.find(n => /quote\s*request/i.test(n));
  if (quoteSheet) {
    const qRows = XLSX.utils.sheet_to_json(wb.Sheets[quoteSheet], { header: 1, defval: '' });
    const summaryPrices = extractSummaryPrices(qRows);
    if (Object.keys(summaryPrices).length > 0) {
      for (const d of devices) {
        const modelKey = d.model.toLowerCase().replace(/\s+/g, ' ').trim();
        for (const [pattern, price] of Object.entries(summaryPrices)) {
          if (modelKey.includes(pattern) || pattern.includes(modelKey.split(' ').slice(-2).join(' '))) {
            d.planbitPrice = price;
            break;
          }
        }
      }
    }
  }

  const ritmSheets = wb.SheetNames.filter(n => /^RITM/i.test(n.trim()));
  if (ritmSheets.length > 0) {
    const ritmWb = { SheetNames: ritmSheets, Sheets: {} };
    for (const n of ritmSheets) ritmWb.Sheets[n] = wb.Sheets[n];
    const ritmDevices = parseGenericHeaders(ritmWb);
    devices = devices.concat(ritmDevices);
  }

  return devices;
}

function parseARSSimple(wb) {
  for (const name of wb.SheetNames) {
    const rows = XLSX.utils.sheet_to_json(wb.Sheets[name], { header: 1, defval: '' });
    const devices = parseARSRows(rows);
    if (devices.length) return devices;
  }
  return [];
}

function parseARSRows(rows) {
  const summaryDevices = [];
  const detailDevices  = [];

  for (const row of rows) {
    const col0 = cellStr(row[0]);
    const col1 = cellStr(row[1]);
    const col2 = cellStr(row[2]);
    const col3 = cellStr(row[3]);
    const col4 = row[4];
    const col5 = row[5];

    if (isProductType(col0) && col1.length > 0 && col2.length > 0 &&
        typeof col4 === 'number' && col4 >= 0) {
      const procStr = col3 + ' ' + col2;
      const qty = Math.max(parseInt(col4) || 1, 1);
      const joepPrice = typeof col5 === 'number' ? col5 : null;

      for (let i = 0; i < qty; i++) {
        summaryDevices.push({
          model:     (col1 + ' ' + col2).trim(),
          cpu:       extractCPU(col3) || extractCPU(col2),
          ram:       extractRAM(procStr) || extractRAM(col2),
          ssd:       extractSSD(col3) || extractSSD(col2),
          joepPrice,
        });
      }
    }

    if (isProductType(col0) && col1 === '' && looksLikeModel(col2) && col3.length > 0) {
      detailDevices.push({ model: col2, serial: col3 });
    }
  }

  if (summaryDevices.length > 0 && detailDevices.length > 0) {
    detailDevices.forEach((d, i) => {
      if (summaryDevices[i]) {
        summaryDevices[i].serial = d.serial;
        if (d.model.length > summaryDevices[i].model.length) {
          summaryDevices[i].model = d.model;
        }
      }
    });
  }

  return summaryDevices.length ? summaryDevices : detailDevices;
}

function extractSummaryPrices(rows) {
  const prices = {};
  for (const row of rows) {
    const col0 = cellStr(row[0]);
    const col1 = cellStr(row[1]);
    const col2 = cellStr(row[2]);
    const col4 = row[4];
    const col5 = row[5];
    if (isProductType(col0) && col1.length > 0 && col2.length > 0 &&
        typeof col4 === 'number' && typeof col5 === 'number' && col5 > 0) {
      const modelKey = (col1 + ' ' + col2).toLowerCase().replace(/\s+/g, ' ').trim();
      prices[modelKey] = col5;
    }
  }
  return prices;
}

function parseVendorQuote(wb) {
  for (const name of wb.SheetNames) {
    const rows = XLSX.utils.sheet_to_json(wb.Sheets[name], { header: 1, defval: '' });

    let headerRowIdx = -1;
    for (let i = 0; i < Math.min(5, rows.length); i++) {
      const cells = rows[i].map(v => cellStr(v).toLowerCase());
      if (cells.some(c => c === 'product') && cells.some(c => c === 'brand') && cells.some(c => c === 'model')) {
        headerRowIdx = i;
        break;
      }
    }
    if (headerRowIdx === -1) continue;

    const header = rows[headerRowIdx].map(v => cellStr(v).toLowerCase());
    const col = (...names) => {
      for (const n of names) {
        const idx = header.findIndex(h => h.includes(n));
        if (idx >= 0) return idx;
      }
      return -1;
    };

    const productCol   = col('product');
    const brandCol     = col('brand');
    const modelCol     = col('model');
    const processorCol = col('processor');
    const hddCol       = col('hdd', 'storage', 'disk');
    const memCol       = col('mem', 'ram', 'memory');
    const qtyCol       = col('qty', 'quantity', 'units');

    const devices = [];
    for (const row of rows.slice(headerRowIdx + 1)) {
      if (!row.some(v => cellStr(v).length > 0)) continue;
      const product = cellStr(row[productCol] ?? '');
      if (!isProductType(product)) continue;

      const brand = cellStr(row[brandCol] ?? '');
      const model = cellStr(row[modelCol] ?? '');
      if (!brand && !model) continue;

      const procRaw = cellStr(row[processorCol] ?? '');
      const hddRaw  = cellStr(row[hddCol] ?? '');
      const memRaw  = cellStr(row[memCol] ?? '');

      const qtyRaw = qtyCol >= 0 ? String(row[qtyCol] ?? '').replace(/,/g, '') : '1';
      const qty = Math.min(parseInt(qtyRaw) || 1, 500);
      const expand = qty <= 20 ? qty : 1;

      const cpu = extractCPU(procRaw) || extractCPU(model);
      const ram = extractRAM(memRaw) || extractRAM(procRaw);
      const ssd = extractSSD(hddRaw) || extractSSD(procRaw);

      for (let i = 0; i < expand; i++) {
        devices.push({ model: (brand + ' ' + model).trim(), cpu, ram, ssd });
      }
    }
    if (devices.length) return devices;
  }
  return [];
}

function parseGenericHeaders(wb) {
  const results = [];

  const sortedSheets = [...wb.SheetNames].sort((a, b) => {
    const rowsA = XLSX.utils.sheet_to_json(wb.Sheets[a], { header: 1, defval: '' }).length;
    const rowsB = XLSX.utils.sheet_to_json(wb.Sheets[b], { header: 1, defval: '' }).length;
    return rowsB - rowsA;
  });

  for (const name of sortedSheets) {
    const rows = XLSX.utils.sheet_to_json(wb.Sheets[name], { header: 1, defval: '' });
    if (!rows.length) continue;

    let headerRowIdx = -1;
    const allAliases = Object.values(COLUMN_ALIASES).flat();

    for (let i = 0; i < Math.min(10, rows.length); i++) {
      const cells = rows[i].map(v => cellStr(v).toLowerCase().replace(/[\s_\-\.]/g,''));
      if (cells.filter(c => allAliases.includes(c)).length >= 1) {
        headerRowIdx = i;
        break;
      }
    }
    if (headerRowIdx === -1) continue;

    const colMap = resolveColumns(rows[headerRowIdx]);
    const mappedIdxs = new Set(Object.values(colMap).filter(i => i >= 0));

    let lastModelIdx = -1;

    for (const row of rows.slice(headerRowIdx + 1)) {
      if (!row.some(v => cellStr(v).length > 0)) continue;
      const model = colMap.model >= 0 ? cellStr(row[colMap.model]) : '';
      const hasModel = model && looksLikeModel(model);

      const cpu   = colMap.cpu   >= 0 ? cellStr(row[colMap.cpu])   : '';
      const ram   = colMap.ram   >= 0 ? cellStr(row[colMap.ram])   : '';
      const ssd   = colMap.ssd   >= 0 ? cellStr(row[colMap.ssd])   : '';
      const grade = colMap.grade >= 0 ? cellStr(row[colMap.grade]) : '';
      const hasSpecs = cpu || ram || ssd || grade;

      if (!hasModel && !hasSpecs) continue;

      if (!hasModel && hasSpecs && lastModelIdx >= 0) {
        const prev = results[lastModelIdx];
        if (cpu   && !prev.cpu)   prev.cpu   = cpu;
        if (ram   && !prev.ram)   prev.ram   = ram;
        if (ssd   && !prev.ssd)   prev.ssd   = ssd;
        if (grade && !prev.grade) prev.grade = grade;
        if (!prev.serial) {
          row.forEach((v, i) => {
            if (!mappedIdxs.has(i) && !prev.serial && /^[A-Z0-9]{6,}$/i.test(cellStr(v))) {
              prev.serial = cellStr(v);
            }
          });
        }
        continue;
      }

      if (!hasModel) continue;

      const brandVal = colMap.brand >= 0 ? cellStr(row[colMap.brand]) : '';
      let fullModel = model;
      if (brandVal && !model.toLowerCase().startsWith(brandVal.toLowerCase())) {
        fullModel = brandVal + ' ' + model;
      }

      const d = { model: fullModel };
      if (cpu)   d.cpu   = cpu;
      if (ram)   d.ram   = ram;
      if (ssd)   d.ssd   = ssd;
      if (grade) d.grade = grade;

      if (!d.serial) {
        row.forEach((v, i) => {
          if (!mappedIdxs.has(i) && !d.serial && /^[A-Z0-9]{6,}$/i.test(cellStr(v))) {
            d.serial = cellStr(v);
          }
        });
      }

      results.push(d);
      lastModelIdx = results.length - 1;
    }
  }
  return results;
}

function parseGenericHeaderless(wb) {
  const results = [];

  for (const name of wb.SheetNames) {
    const rows = XLSX.utils.sheet_to_json(wb.Sheets[name], { header: 1, defval: '' });
    const nonEmpty = rows.filter(r => r.some(v => cellStr(v).length > 0));
    if (nonEmpty.length < 2) continue;

    const numCols = Math.max(...nonEmpty.map(r => r.length), 0);

    const scores = {};
    for (let c = 0; c < numCols; c++) {
      const vals = nonEmpty.map(r => cellStr(r[c])).filter(v => v.length > 0);
      if (!vals.length) continue;
      const r = n => n / vals.length;
      scores[c] = {
        model:  r(vals.filter(looksLikeModel).length),
        serial: r(vals.filter(v => /^[A-Z0-9]{6,}$/i.test(v)).length),
        ram:    r(vals.filter(v => /^(4|8|16|32|64|128)\s*GB?$/i.test(v)).length),
        ssd:    r(vals.filter(v => /^(128|240|256|480|512|960|1024|2048)\s*GB?$/i.test(v)).length),
        grade:  r(vals.filter(v => /^[A-D]\d?$/i.test(v)).length),
      };
    }

    const pick = (field, exclude = []) => {
      let best = -1, bestScore = 0.12;
      for (let c = 0; c < numCols; c++) {
        if (exclude.includes(c) || !scores[c]) continue;
        if ((scores[c][field] || 0) > bestScore) { best = c; bestScore = scores[c][field]; }
      }
      return best;
    };

    const modelCol  = pick('model');
    if (modelCol === -1) continue;
    const serialCol = pick('serial', [modelCol]);
    const ramCol    = pick('ram',    [modelCol, serialCol]);
    const ssdCol    = pick('ssd',    [modelCol, serialCol, ramCol]);
    const gradeCol  = pick('grade',  [modelCol, serialCol, ramCol, ssdCol]);

    const sheetResults = nonEmpty
      .filter(row => looksLikeModel(row[modelCol]))
      .map(row => {
        const d = { model: cellStr(row[modelCol]) };
        if (serialCol >= 0) d.serial = cellStr(row[serialCol]);
        if (ramCol    >= 0) d.ram    = cellStr(row[ramCol]);
        if (ssdCol    >= 0) d.ssd    = cellStr(row[ssdCol]);
        if (gradeCol  >= 0) d.grade  = cellStr(row[gradeCol]);
        return d;
      });

    results.push(...sheetResults);
  }
  return results;
}

// ─── ZONES INVENTORY PARSER ──────────────────────────────────────────────────

function parseZonesInventory(wb) {
  const devices = [];

  for (const name of wb.SheetNames) {
    const rows = XLSX.utils.sheet_to_json(wb.Sheets[name], { header: 1, defval: '' });
    if (rows.length < 2) continue;

    let headerIdx = -1;
    let colMap = {};
    for (let i = 0; i < Math.min(5, rows.length); i++) {
      const cm = resolveColumns(rows[i]);
      if (cm.brand >= 0 && cm.model >= 0) {
        headerIdx = i;
        colMap = cm;
        break;
      }
    }
    if (headerIdx < 0) continue;

    const groups = {};
    for (let i = headerIdx + 1; i < rows.length; i++) {
      const row = rows[i];
      if (!isDeviceRow(row, colMap)) continue;

      const brand = cellStr(row[colMap.brand]).toUpperCase();
      const model = cellStr(row[colMap.model]).trim();
      if (!model) continue;

      const key = `${brand}|${model}`;
      if (!groups[key]) {
        const grade = colMap.grade >= 0 ? cellStr(row[colMap.grade]) : '';
        groups[key] = { brand, model, grade: grade || '', qty: 0, serials: [] };
      }
      groups[key].qty++;
      if (colMap.serial >= 0) {
        const s = cellStr(row[colMap.serial]);
        if (s && s !== 'N/A') groups[key].serials.push(s);
      }
    }

    for (const [key, g] of Object.entries(groups)) {
      const gen = inferGenFromModel(g.brand, g.model);
      devices.push({
        brand: g.brand,
        model: `${g.brand} ${g.model}`,
        cpu: gen ? `Inferred ${gen} from model` : 'Unknown',
        gen: gen || 'Unknown',
        ram: 8,
        ssd: 256,
        grade: g.grade || 'B',
        qty: g.qty,
        quantity: g.qty,
      });
    }
  }

  return devices;
}

// ─── TEXT INPUT PARSER ────────────────────────────────────────────────────────
function parseTextInput(text) {
  const EMAIL_JUNK = /^(summarize this email|inbox|re:|fwd:|from:|to me|to:|subject:|sent:|cc:|planbit|caas\s*[-–]?\s*$|feb |mar |jan |apr |may |jun |jul |aug |sep |oct |nov |dec |http|analyseer|target|geef|caps|qwertzu|lot-factor|totaal:)/i;
  const HAS_DEVICE = /\b(dell|hp|hewlett|lenovo|apple|macbook|thinkpad|latitude|elitebook|surface|asus|acer|fujitsu|toshiba|samsung|microsoft|iphone|ipad|galaxy|yoga|optiplex|probook|zbook|spectre|envy|pavilion|chromebook)\b/i;
  const SECTION_HEADER = /^(APPLE|DELL|HP|LENOVO|TOSHIBA|ACER|ASUS|MICROSOFT)\s*[—–-]\s*\d+\s*stuks?/i;

  const lines = text
    .split(/[\n;]+/)
    .map(l => l.trim())
    .filter(l => l.length > 4)
    .filter(l => !EMAIL_JUNK.test(l) || HAS_DEVICE.test(l))
    .filter(l => !SECTION_HEADER.test(l));

  const devices = [];

  for (const line of lines) {
    if (!HAS_DEVICE.test(line)) continue;

    let cleaned = line.replace(/^\d{1,3}[\.\)]\s*/, '');

    let qty = 1;
    const qtyEndMatch = cleaned.match(/[·\-]\s*(\d+)\s*[xX×]\s*$/);
    const qtyStuksMatch = cleaned.match(/[·\-]\s*(\d+)\s*stuks?\s*$/i);
    const qtyStartMatch = cleaned.match(/^(\d+)\s*[xX×]\s+/);
    if (qtyEndMatch) {
      qty = Math.min(parseInt(qtyEndMatch[1]), 9999);
      cleaned = cleaned.slice(0, cleaned.indexOf(qtyEndMatch[0])).trim();
    } else if (qtyStuksMatch) {
      qty = Math.min(parseInt(qtyStuksMatch[1]), 9999);
      cleaned = cleaned.slice(0, cleaned.indexOf(qtyStuksMatch[0])).trim();
    } else if (qtyStartMatch) {
      qty = Math.min(parseInt(qtyStartMatch[1]), 9999);
      cleaned = cleaned.slice(qtyStartMatch[0].length).trim();
    }

    const parts = cleaned.split(/\s*·\s*/);

    let modelRaw, specsPart;
    if (parts.length >= 2) {
      modelRaw = parts[0].trim();
      specsPart = parts.slice(1).join(' ');
    } else {
      const modelStop = /\b(i[3579][-\s]\d|Ryzen|M[123]\s|FHD|UHD)/i;
      const stopMatch = cleaned.match(modelStop);
      if (stopMatch) {
        modelRaw = cleaned.slice(0, stopMatch.index).trim();
        specsPart = cleaned.slice(stopMatch.index);
      } else {
        modelRaw = cleaned;
        specsPart = '';
      }
    }

    modelRaw = modelRaw.replace(/^[-\*\•\·\s]+/, '').replace(/\s+/g, ' ').trim();
    if (!HAS_DEVICE.test(modelRaw) || modelRaw.length < 4) continue;

    const ramSsdMatch = (specsPart || '').match(/(\d+)\s*(?:GB)?\s*\/\s*(\d+)\s*(?:GB)?/);
    let ram = null, ssd = null;
    if (ramSsdMatch) {
      const v1 = parseInt(ramSsdMatch[1]);
      const v2 = parseInt(ramSsdMatch[2]);
      if (v1 <= 64 && v2 >= 64) { ram = v1; ssd = v2; }
      else if (v1 <= 64 && v2 <= 64) { ram = v1; ssd = v2; }
      else { ram = v1; ssd = v2; }
    }
    if (!ram) ram = extractRAM(specsPart || cleaned);
    if (!ssd) ssd = extractSSD(specsPart || cleaned);

    const cpu = extractCPU(specsPart || cleaned);
    const grade = extractGrade(specsPart || cleaned);
    // Try model first, then full line (catches "T14 Gen 2" where Gen 2 is in specsPart)
    const gen = inferGenFromModel('', modelRaw) || inferGenFromModel('', cleaned);

    devices.push({ model: modelRaw, cpu, ram, ssd, grade, qty, quantity: qty, gen });
  }

  // Deduplicate by model name
  const merged = [];
  for (const d of devices) {
    const existing = merged.find(m => m.model.toLowerCase() === d.model.toLowerCase());
    if (existing) {
      existing.qty = (existing.qty || 1) + (d.qty || 1);
      existing.quantity = existing.qty;
      if (!existing.ram && d.ram) existing.ram = d.ram;
      if (!existing.ssd && d.ssd) existing.ssd = d.ssd;
      if (!existing.cpu && d.cpu) existing.cpu = d.cpu;
      if (!existing.grade && d.grade) existing.grade = d.grade;
      if (!existing.gen && d.gen) existing.gen = d.gen;
    } else {
      merged.push(d);
    }
  }
  return merged;
}

// ─── MAIN ENTRY POINTS ──────────────────────────────────────────────────────

function parseExcel(buffer) {
  const wb = XLSX.read(buffer, { type: 'buffer' });
  const format = detectFormat(wb);

  switch (format) {
    case 'ARS':              return { devices: parseARS(wb), format };
    case 'ARS_SIMPLE':       return { devices: parseARSSimple(wb), format };
    case 'VENDOR_QUOTE':     return { devices: parseVendorQuote(wb), format };
    case 'ZONES_INVENTORY':  return { devices: parseZonesInventory(wb), format };
    case 'GENERIC_HEADERS':  return { devices: parseGenericHeaders(wb), format };
    default:                 return { devices: parseGenericHeaderless(wb), format: 'GENERIC_HEADERLESS' };
  }
}

function parseFile(buffer, filename) {
  const ext = (filename || '').toLowerCase().split('.').pop();
  if (['xlsx', 'xls', 'csv'].includes(ext)) {
    return parseExcel(buffer);
  }
  // Try text parsing for .txt or unknown
  const text = buffer.toString('utf-8');
  return { devices: parseTextInput(text), format: 'TEXT' };
}

function parseText(text) {
  return { devices: parseTextInput(text), format: 'TEXT' };
}

module.exports = {
  parseFile,
  parseText,
  detectFormat,
  parseExcel,
  parseTextInput,
  // Also export helpers for testing
  extractRAM,
  extractSSD,
  extractCPU,
  extractGrade,
  inferGenFromModel,
};
