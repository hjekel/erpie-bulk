'use strict';

const XLSX = require('xlsx');

// ═══════════════════════════════════════════════════════════════════════════
// UNIVERSAL DEVICE PARSER v3.0
// Handles: ARS Excel, PWC/Vendor Quote, Zones Inventory, Generic Headers,
//          Generic Headerless, Text Input, Free-form text in any language
// ═══════════════════════════════════════════════════════════════════════════

const BRANDS = ['apple','dell','hp','lenovo','microsoft','asus','acer',
                'fujitsu','toshiba','samsung','sony','lg','panasonic',
                'huawei','xiaomi','google','razer','msi','gigabyte'];

const BRAND_MODELS = {
  apple:     ['macbook','imac','mac mini','mac pro','mac studio','iphone','ipad','ipod'],
  dell:      ['latitude','inspiron','xps','precision','vostro','optiplex','alienware','chromebook 3'],
  hp:        ['elitebook','probook','zbook','pavilion','envy','spectre','omen',
               'elitedesk','prodesk','z-workstation','chromebook','folio','revolve','elitedragonfly','dragonfly'],
  lenovo:    ['thinkpad','thinkcentre','thinkstation','ideapad','ideacentre','legion','yoga','tab'],
  microsoft: ['surface'],
  asus:      ['zenbook','vivobook','rog','tuf'],
  acer:      ['aspire','predator','swift','spin','travelmate','chromebook'],
  fujitsu:   ['lifebook','esprimo','celsius'],
  toshiba:   ['portege','tecra','satellite','dynabook'],
  samsung:   ['galaxy book','galaxy tab','galaxy'],
};

// Map brand keys to display names (handles special cases like HP, LG)
const BRAND_DISPLAY = { hp: 'HP', lg: 'LG', msi: 'MSI', asus: 'ASUS' };
function brandDisplay(key) {
  return BRAND_DISPLAY[key] || key.charAt(0).toUpperCase() + key.slice(1);
}

function detectBrand(model) {
  if (!model) return '';
  const lower = model.toLowerCase();
  // HP special cases first
  if (/^hewlett[\s-]?packard/i.test(lower)) return 'HP';
  if (/^hp\b/i.test(lower)) return 'HP';
  // Direct brand name match
  for (const brand of BRANDS) {
    if (lower.startsWith(brand + ' ') || lower.startsWith(brand + '-') || lower === brand) {
      return brandDisplay(brand);
    }
  }
  // Model-name based detection
  for (const [brand, models] of Object.entries(BRAND_MODELS)) {
    for (const m of models) {
      if (lower.includes(m)) return brandDisplay(brand);
    }
  }
  return '';
}

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

// Grade normalization map: natural-language conditions → standard grades
const GRADE_WORD_MAP = {
  // Grade A (like new / excellent)
  'likenew': 'A', 'asnew': 'A', 'new': 'A', 'mint': 'A',
  'excellent': 'A', 'pristine': 'A', 'perfect': 'A',
  'neuwertig': 'A', 'wieneu': 'A', 'neuf': 'A', 'comenuovo': 'A',
  'nuevo': 'A', 'uitstekend': 'A', 'nieuwstaat': 'A',
  'ausgezeichnet': 'A', 'hervorragend': 'A',
  // Grade B (good / minor wear)
  'good': 'B', 'verygood': 'B', 'minor': 'B', 'minorwear': 'B',
  'gut': 'B', 'sehrgut': 'B', 'bon': 'B', 'tresbon': 'B',
  'buono': 'B', 'bueno': 'B', 'muybueno': 'B', 'goed': 'B',
  'zeergoed': 'B', 'refurbished': 'B', 'refurb': 'B',
  'functional': 'B', 'working': 'B', 'functioneel': 'B', 'werkend': 'B',
  // Grade C (fair / visible wear)
  'fair': 'C', 'acceptable': 'C', 'moderate': 'C', 'used': 'C',
  'heavywear': 'C', 'scratched': 'C', 'damaged': 'C',
  'befriedigend': 'C', 'ausreichend': 'C', 'passable': 'C',
  'moyen': 'C', 'acceptabel': 'C', 'redelijk': 'C', 'matig': 'C',
  'gebraucht': 'C', 'gebruikt': 'C',
  // Grade D (poor / broken cosmetics)
  'poor': 'D', 'bad': 'D', 'broken': 'D', 'defect': 'D',
  'schlecht': 'D', 'mangelhaft': 'D', 'mauvais': 'D',
  'slecht': 'D', 'kapot': 'D', 'defekt': 'D',
};

function extractGrade(s) {
  if (!s) return '';
  const str = String(s).trim();

  // Exact standard grades: A1, A2, A3, A, B1-B4, B, C1, C2, C6, C, D, P7, X9
  const exact = str.match(/\b(A1|A2|A3|A|B1|B2|B3|B4|B|C1|C2|C6|C|D|P7|X9)\b/i);
  if (exact) return exact[1].toUpperCase();

  // "X-grade" / "X grade" pattern: "b-grade", "A grade"
  const xGrade = str.match(/\b([A-D])\s*[-–]\s*grade\b/i);
  if (xGrade) return xGrade[1].toUpperCase();

  // "Grade X" / "Klasse X" / "Condition X" / "Kwaliteit X" prefixed
  const prefixed = str.match(/(?:grade|klasse|condition|conditie|zustand|état|etat|grado|nota|kwaliteit|qualiteit)\s*[:=]?\s*([A-D]\d?)\b/i);
  if (prefixed) return prefixed[1].toUpperCase();

  // Natural-language condition words
  const normalized = str.toLowerCase().replace(/[\s\-_]/g, '');
  for (const [word, grade] of Object.entries(GRADE_WORD_MAP)) {
    if (normalized === word || normalized.includes(word)) return grade;
  }

  // Numeric grades: 1-10 scale (common in DE/NL)
  const numGrade = str.match(/\b([1-9]|10)\s*(?:\/\s*10)?\b/);
  if (numGrade) {
    const n = parseInt(numGrade[1]);
    if (n >= 9) return 'A';
    if (n >= 7) return 'B';
    if (n >= 5) return 'C';
    if (n >= 1) return 'D';
  }

  return '';
}

// ─── UNIFIED COLUMN RESOLVER ─────────────────────────────────────────────────
// Multilingual aliases: EN, NL, DE, FR, ES, IT, PL, PT + common abbreviations
const COLUMN_ALIASES = {
  brand:    ['brand','manufacturer','make','merk','vendor','oem','fabrikant',
             'hersteller','marke','fabricant','marque','marca','fabricante',
             'producent','marka','costruttore'],
  model:    ['model','device','modelname','computername','description','productname',
             'name','item','assetname','partnumber','sku','modell','modele','modelo',
             'apparaat','toestel','geraet','geräte','appareil','dispositivo','artikel',
             'artikelnummer','devicename','productdescription','omschrijving','beschreibung'],
  cpu:      ['cpu','processor','proc','cputype','cpumodel','chipset','specs',
             'prozessor','processeur','procesador','processore','specificaties',
             'spezifikationen','specifications'],
  ram:      ['ram','memory','mem','ramsizeoutgoing','systemmemory','installedmemory',
             'ramsize','geheugen','arbeitsspeicher','speicher','memoire','mémoire',
             'memoria','werkgeheugen','ramgb'],
  ssd:      ['ssd','storage','hdd','disk','drive','hddsizeoutgoing','capacity',
             'disksize','hddsize','storagesize','opslag','festplatte','speicherplatz',
             'stockage','discoduro','almacenamiento','archiviazione','opslagcapaciteit',
             'harddisk','hardeschijf'],
  grade:    ['grade','condition','quality','cosmeticcategory','functionalgrade',
             'cosmeticgrade','gradein','conditie','staat','zustand','klasse','note',
             'etat','état','condicion','condición','condizione','grado','kwaliteit',
             'toestand','bewertung'],
  serial:   ['serial','serialnumber','sn','assettag','asset','devicetag','equipmentid',
             'uid','serialnumberasset','serienummer','seriennummer','numérosérie',
             'numeroserie','númeroserie','numerodiserie'],
  qty:      ['qty','quantity','count','units','aantal','anzahl','menge','stück','stueck',
             'quantité','quantite','cantidad','quantità','quantita','pcs','pieces',
             'stuks','exemplaren','unidades','unità','ilość','sztuk','adet'],
  type:     ['producttype','producttyp','productype','produkttyp','typeproduit',
             'tipoprodotto','tipoproducto'],
  category: ['modelcategory','category','devicetype','assettype','categorie','kategorie',
             'catégorie','categoría','categoria'],
  keyboard: ['keyboard','kb','toetsenbord','layout','tastatur','clavier','teclado','tastiera'],
  location: ['location','country','region','site','locatie','locationdisplay','standort',
             'emplacement','ubicación','ubicacion','posizione','lokalizacja'],
};

// Strip diacritics/accents: é→e, ü→u, ö→o, etc.
function stripAccents(s) {
  return s.normalize('NFD').replace(/[\u0300-\u036f]/g, '');
}

// Fuzzy column resolver: also partial-match long headers
function fuzzyMatchAlias(header, aliases) {
  // Try with and without accents
  const variants = [header, stripAccents(header)];
  for (const h of variants) {
    if (aliases.includes(h)) return true;
    for (const a of aliases) {
      if (a.length >= 3 && h.startsWith(a)) return true;
      if (a.length >= 4 && h.includes(a)) return true;
    }
  }
  return false;
}

function resolveColumns(headerRow) {
  const cells = headerRow.map(v => cellStr(v).toLowerCase().replace(/[\s_\-\.]/g, ''));
  const map = {};
  for (const [field, aliases] of Object.entries(COLUMN_ALIASES)) {
    const idx = cells.findIndex(c => fuzzyMatchAlias(c, aliases));
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

// ─── LENOVO SHORT MODEL → FULL THINKPAD NAME ────────────────────────────────
// Maps short Lenovo model names (from Blancco "Model" column) to full names
const LENOVO_SHORT_MAP = {
  'T14 G1':      'ThinkPad T14 Gen 1',
  'T14 G2':      'ThinkPad T14 Gen 2',
  'T14 G3':      'ThinkPad T14 Gen 3',
  'T14 G4':      'ThinkPad T14 Gen 4',
  'T14S G1':     'ThinkPad T14s Gen 1',
  'T14S G2':     'ThinkPad T14s Gen 2',
  'T14S G3':     'ThinkPad T14s Gen 3',
  'T450':        'ThinkPad T450',
  'T460':        'ThinkPad T460',
  'T470':        'ThinkPad T470',
  'T480':        'ThinkPad T480',
  'T490':        'ThinkPad T490',
  'T495':        'ThinkPad T495',
  'X270':        'ThinkPad X270',
  'X280':        'ThinkPad X280',
  'X380 YOGA':   'ThinkPad X380 Yoga',
  'X13 G1':      'ThinkPad X13 Gen 1',
  'X13 G2':      'ThinkPad X13 Gen 2',
  'X13 G3':      'ThinkPad X13 Gen 3',
  'X1 CARBON G5':'ThinkPad X1 Carbon Gen 5',
  'X1 CARBON G6':'ThinkPad X1 Carbon Gen 6',
  'X1 CARBON G7':'ThinkPad X1 Carbon Gen 7',
  'X1 CARBON G8':'ThinkPad X1 Carbon Gen 8',
  'X1 CARBON G9':'ThinkPad X1 Carbon Gen 9',
  'X1 CARBON G10':'ThinkPad X1 Carbon Gen 10',
  'L14 G1':      'ThinkPad L14 Gen 1',
  'L14 G2':      'ThinkPad L14 Gen 2',
  'L14 G3':      'ThinkPad L14 Gen 3',
  'L380':        'ThinkPad L380',
  'L480':        'ThinkPad L480',
  'L490':        'ThinkPad L490',
  'E14 G2':      'ThinkPad E14 Gen 2',
  'E14 G3':      'ThinkPad E14 Gen 3',
  'P14S G1':     'ThinkPad P14s Gen 1',
  'P14S G2':     'ThinkPad P14s Gen 2',
  'P14S G3':     'ThinkPad P14s Gen 3',
};

// Lenovo Blancco part number prefix → ThinkPad model
const LENOVO_PART_PREFIX_MAP = {
  '20S1': 'ThinkPad T14 Gen 1', '20T3': 'ThinkPad T14 Gen 1',
  '20W1': 'ThinkPad T14 Gen 2', '20XK': 'ThinkPad T14 Gen 2',
  '21AJ': 'ThinkPad T14 Gen 3', '21AH': 'ThinkPad T14 Gen 3',
  '21HE': 'ThinkPad T14 Gen 4', '21HF': 'ThinkPad T14 Gen 4',
  '20T0': 'ThinkPad T14s Gen 1', '20T1': 'ThinkPad T14s Gen 1',
  '20L6': 'ThinkPad T480', '20L5': 'ThinkPad T480',
  '20N3': 'ThinkPad T490', '20N2': 'ThinkPad T490',
  '20K5': 'ThinkPad X270', '20K4': 'ThinkPad X270',
  '20KE': 'ThinkPad X280', '20KF': 'ThinkPad X280',
  '20LH': 'ThinkPad X380 Yoga', '20LJ': 'ThinkPad X380 Yoga',
  '20UB': 'ThinkPad X13 Gen 1', '20UA': 'ThinkPad X13 Gen 1',
  '20WK': 'ThinkPad X13 Gen 2', '20WL': 'ThinkPad X13 Gen 2',
  '20BU': 'ThinkPad T450', '20BT': 'ThinkPad T450',
  '20F1': 'ThinkPad T460', '20F2': 'ThinkPad T460',
  '20HE': 'ThinkPad T470', '20HF': 'ThinkPad T470',
  '20JN': 'ThinkPad T470', '20FM': 'ThinkPad T460',
};

// Apple Blancco model identifiers → human-readable names
const APPLE_BLANCCO_MAP = {
  'MACBOOKAIR10,1':  'MacBook Air M1',
  'MACBOOKAIR10,2':  'MacBook Air M2',
  'MACBOOKAIR7,1':   'MacBook Air 2015',
  'MACBOOKAIR7,2':   'MacBook Air 2015',
  'MACBOOKAIR8,1':   'MacBook Air 2018',
  'MACBOOKAIR8,2':   'MacBook Air 2019',
  'MACBOOKAIR9,1':   'MacBook Air 2020',
  'MACBOOKPRO16,1':  'MacBook Pro 16" 2019',
  'MACBOOKPRO16,2':  'MacBook Pro 13" 2020',
  'MACBOOKPRO16,3':  'MacBook Pro 13" 2020',
  'MACBOOKPRO16,4':  'MacBook Pro 16" 2019',
  'MACBOOKPRO15,1':  'MacBook Pro 15" 2018',
  'MACBOOKPRO15,2':  'MacBook Pro 13" 2018',
  'MACBOOKPRO15,3':  'MacBook Pro 15" 2019',
  'MACBOOKPRO15,4':  'MacBook Pro 13" 2019',
  'MACBOOKPRO14,1':  'MacBook Pro 13" 2017',
  'MACBOOKPRO14,2':  'MacBook Pro 13" 2017',
  'MACBOOKPRO14,3':  'MacBook Pro 15" 2017',
  'MACBOOKPRO18,1':  'MacBook Pro 16" 2021',
  'MACBOOKPRO18,2':  'MacBook Pro 16" 2021',
  'MACBOOKPRO18,3':  'MacBook Pro 14" 2021',
  'MACBOOKPRO18,4':  'MacBook Pro 14" 2021',
  'MACBOOKPRO17,1':  'MacBook Pro 13" M1',
};

// Apple A-code → model name (for when Model column has Apple part numbers)
const APPLE_ACODE_MAP = {
  'A2141': 'MacBook Pro 16" 2019',
  'A2337': 'MacBook Air M1',
  'A2338': 'MacBook Pro 13" M1',
  'A1708': 'MacBook Pro 13" 2017',
  'A1706': 'MacBook Pro 13" 2017',
  'A1707': 'MacBook Pro 15" 2017',
  'A1990': 'MacBook Pro 15" 2018',
  'A1989': 'MacBook Pro 13" 2018',
  'A2159': 'MacBook Pro 13" 2019',
  'A2289': 'MacBook Pro 13" 2020',
  'A2251': 'MacBook Pro 13" 2020',
  'A1932': 'MacBook Air 2018',
  'A2179': 'MacBook Air 2020',
  'A1466': 'MacBook Air 2015',
};

/**
 * Normalize a Blancco/ARS model name to a human-readable name
 * @param {string} brand - Brand name (uppercase)
 * @param {string} model - Model from "Model" column (may be short name or A-code)
 * @param {string} blancco - Model from "Model (Blancco)" column
 * @param {string} cpu - CPU string for disambiguation
 * @returns {string} Human-readable model name
 */
function normalizeBlanccoModel(brand, model, blancco, cpu) {
  const brandUp = (brand || '').toUpperCase();
  const modelUp = (model || '').toUpperCase().trim();
  const blanccoUp = (blancco || '').toUpperCase().trim();

  if (brandUp === 'LENOVO') {
    // 1) Try short model map: "T14 G1" → "ThinkPad T14 Gen 1"
    const shortKey = modelUp.replace(/\s+/g, ' ');
    if (LENOVO_SHORT_MAP[shortKey]) return LENOVO_SHORT_MAP[shortKey];

    // 2) Try Blancco part number prefix: "20S1S4RB2E" → first 4 chars "20S1"
    if (/^(20|21)[A-Z0-9]/i.test(blanccoUp)) {
      const prefix = blanccoUp.slice(0, 4).toUpperCase();
      if (LENOVO_PART_PREFIX_MAP[prefix]) return LENOVO_PART_PREFIX_MAP[prefix];
    }

    // 3) If model already says "ThinkPad", use as-is
    if (/thinkpad/i.test(modelUp)) {
      return model.trim();
    }

    // 4) Fallback: prepend ThinkPad if it looks like a T/X/L/E/P series
    if (/^[TXLEP]\d/i.test(modelUp)) {
      return 'ThinkPad ' + model.trim();
    }

    return model.trim();
  }

  if (brandUp === 'APPLE') {
    // 1) Try A-code map: "A2141" → "MacBook Pro 16\" 2019"
    if (APPLE_ACODE_MAP[modelUp]) return APPLE_ACODE_MAP[modelUp];

    // 2) Try Blancco identifier: "MACBOOKPRO16,1" → "MacBook Pro 16\" 2019"
    const blanccoClean = blanccoUp.replace(/\s+/g, '');
    if (APPLE_BLANCCO_MAP[blanccoClean]) return APPLE_BLANCCO_MAP[blanccoClean];

    // 3) Try model field if it's a Blancco-style ID
    const modelClean = modelUp.replace(/\s+/g, '');
    if (APPLE_BLANCCO_MAP[modelClean]) return APPLE_BLANCCO_MAP[modelClean];

    return model.trim();
  }

  if (brandUp === 'HP') {
    // If model is "NA" or empty, use Blancco model (strip "HP " prefix and " NOTEBOOK PC" suffix)
    if (!modelUp || modelUp === 'NA') {
      let name = (blancco || '').replace(/^HP\s+/i, '').replace(/\s+NOTEBOOK\s+PC$/i, '').trim();
      return name || model.trim();
    }
    // Strip "HP " prefix from model if present in Blancco but model is clean
    return model.trim();
  }

  if (brandUp === 'MICROSOFT') {
    // If model is "NA", use Blancco
    if (!modelUp || modelUp === 'NA') {
      return (blancco || '').trim();
    }
    return model.trim();
  }

  if (brandUp === 'ACER') {
    // Use Blancco if it's more descriptive (e.g., "TRAVELMATE P215-53" vs "P215-53")
    if (blanccoUp.length > modelUp.length && /[A-Z]/i.test(blanccoUp)) {
      return blancco.trim();
    }
    return model.trim();
  }

  // Default: use Blancco if model is empty/NA, otherwise model
  if (!modelUp || modelUp === 'NA') return (blancco || '').trim();
  return model.trim();
}

// ─── BLANCCO DETAIL FORMAT PARSER ────────────────────────────────────────────
// Detects PlanBit ARS format with "Original Input" detail section
// Columns: Facility | Category | Asset ID | Part Number | Mfg | Model |
//          Model (Blancco) | Serial Number | CPU | ... | Memory (GB) | HDD (GB) | ... | Grade
function parseBlanccoDetail(wb) {
  // Search ALL sheets for Blancco detail headers (not just "Customer Inventory")
  let sheetName = null;
  let headerIdx = -1;

  for (const name of wb.SheetNames) {
    const rows = XLSX.utils.sheet_to_json(wb.Sheets[name], { header: 1, defval: '' });
    for (let i = 0; i < Math.min(rows.length, 60); i++) {
      const cells = rows[i].map(v => cellStr(v).toLowerCase());
      if ((cells.some(c => c === 'mfg' || c === 'manufacturer')) &&
          (cells.some(c => c.includes('blancco') || c.includes('serial')))) {
        sheetName = name;
        headerIdx = i;
        break;
      }
    }
    if (sheetName) break;
  }
  if (!sheetName || headerIdx < 0) return null;

  const rows = XLSX.utils.sheet_to_json(wb.Sheets[sheetName], { header: 1, defval: '' });

  // Map columns
  const header = rows[headerIdx].map(v => cellStr(v).toLowerCase().replace(/[\r\n]/g, ' ').replace(/\s+/g, ' ').trim());
  const col = (name) => header.findIndex(h => h.includes(name));

  const mfgCol = col('mfg');
  const modelCol = header.findIndex(h => h === 'model');
  const blanccoCol = col('blancco');
  const cpuCol = col('cpu');
  const ramCol = header.findIndex(h => h.includes('memory') && h.includes('gb'));
  const ssdCol = header.findIndex(h => (h.includes('hdd') || h.includes('storage')) && h.includes('gb'));
  const gradeCol = header.findIndex(h => h === 'grade');
  const serialCol = col('serial');
  const kbCol = col('keyboard');

  if (mfgCol < 0 && modelCol < 0) return null;

  const devices = [];
  for (let i = headerIdx + 1; i < rows.length; i++) {
    const row = rows[i];
    const brand = cellStr(row[mfgCol] ?? '').toUpperCase();
    if (!brand || brand.length < 2) continue;

    const modelRaw = cellStr(row[modelCol] ?? '');
    const blanccoRaw = blanccoCol >= 0 ? cellStr(row[blanccoCol] ?? '') : '';
    const cpuRaw = cpuCol >= 0 ? cellStr(row[cpuCol] ?? '') : '';
    const ramRaw = ramCol >= 0 ? row[ramCol] : '';
    const ssdRaw = ssdCol >= 0 ? row[ssdCol] : '';
    const gradeRaw = gradeCol >= 0 ? cellStr(row[gradeCol] ?? '') : 'B';
    const serialRaw = serialCol >= 0 ? cellStr(row[serialCol] ?? '') : '';
    const kbRaw = kbCol >= 0 ? cellStr(row[kbCol] ?? '') : '';

    // Normalize model name
    const modelName = normalizeBlanccoModel(brand, modelRaw, blanccoRaw, cpuRaw);
    if (!modelName || modelName === 'NA') continue;

    // Parse specs
    const ram = typeof ramRaw === 'number' ? ramRaw : (parseInt(String(ramRaw)) || 0);
    const ssd = typeof ssdRaw === 'number' ? ssdRaw : (parseInt(String(ssdRaw)) || 0);
    const cpu = extractCPU(cpuRaw);
    const grade = extractGrade(gradeRaw) || 'B';

    devices.push({
      brand: brandDisplay(brand.toLowerCase()),
      model: modelName,
      cpu,
      ram: ram > 0 ? ram + 'GB' : '',
      ssd: ssd > 0 ? ssd + 'GB' : '',
      grade,
      serial: serialRaw,
      keyboard: kbRaw,
      qty: 1,
      quantity: 1,
    });
  }

  return devices.length > 0 ? devices : null;
}

// ─── SMART FORMAT SCORING ────────────────────────────────────────────────────

function detectFormat(wb) {
  const scores = { BLANCCO_DETAIL: 0, VENDOR_QUOTE: 0, ARS: 0, ZONES_INVENTORY: 0, GENERIC_HEADERS: 0, ARS_SIMPLE: 0 };

  // Check for PlanBit ARS Blancco detail format — scan ALL sheets for Blancco headers
  for (const sheetName of wb.SheetNames) {
    const rows = XLSX.utils.sheet_to_json(wb.Sheets[sheetName], { header: 1, defval: '' });
    for (let i = 0; i < Math.min(rows.length, 60); i++) {
      const cells = rows[i].map(v => cellStr(v).toLowerCase());
      if ((cells.some(c => c === 'mfg' || c === 'manufacturer')) &&
          (cells.some(c => c.includes('blancco') || c === 'serial number' || c === 'serial'))) {
        scores.BLANCCO_DETAIL = 20; // highest priority
        break;
      }
    }
    if (scores.BLANCCO_DETAIL > 0) break;
  }

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

// ─── QUANTITY EXTRACTION ─────────────────────────────────────────────────────
// Handles: "60x", "60 x", "60×", "60 pcs", "60 stuks", "60 stück", "qty: 60",
// "(60)", "60 units", "60 st", "Anzahl: 60", "quantité: 60", "aantal: 60"
function extractQuantity(s) {
  if (!s) return { qty: 1, rest: s };
  let str = String(s).trim();

  // Quantity patterns, ordered by specificity
  const QTY_UNIT_RE = /stuks?|st[üu]ck|pcs|pieces?|units?|exemplaren|unidades|unit[àa]|sztuk|adet/;
  const patterns = [
    // "qty: 60" / "qty 60" / "quantity: 60" / "aantal: 60" / "anzahl: 60" (colon optional)
    { re: /(?:qty|quantity|aantal|anzahl|quantit[éeà]|cantidad|menge)\s*[:=]?\s*(\d+)/i, pos: 'remove' },
    // "60x " / "60 x " / "60× " at start
    { re: /^(\d+)\s*[xX×]\s+/, pos: 'start' },
    // "- 60x" / "· 60x" at end
    { re: /[·\-–—]\s*(\d+)\s*[xX×]\s*$/, pos: 'end' },
    // "60 stuks" etc anywhere (not just end) — most common Dutch format
    { re: new RegExp('[·,\\-–—]?\\s*(\\d+)\\s*(?:' + QTY_UNIT_RE.source + ')(?:\\s*[·,\\-–—]|\\s*$)', 'i'), pos: 'remove' },
    // "60 stuks/pcs/units" at start
    { re: new RegExp('^(\\d+)\\s*(?:' + QTY_UNIT_RE.source + ')\\s+', 'i'), pos: 'start' },
    // "(60)" anywhere — quantity in parens
    { re: /\((\d+)\)/, pos: 'remove' },
    // "× 60" / "x60" / "x 60" at end or before grade/conditie
    { re: /\s*[xX×]\s*(\d+)\s*(?:$|\s+(?:grade|conditie|kwaliteit|condition))/i, pos: 'remove' },
    // "60x" at start (no space needed)
    { re: /^(\d+)[xX×](?=[A-Za-z])/, pos: 'start' },
  ];

  for (const { re, pos } of patterns) {
    const m = str.match(re);
    if (m && parseInt(m[1]) > 0 && parseInt(m[1]) <= 9999) {
      const qty = parseInt(m[1]);
      if (pos === 'start') {
        str = str.slice(m[0].length).trim();
      } else if (pos === 'end') {
        str = str.slice(0, m.index).trim();
      } else {
        str = str.replace(m[0], ' ').replace(/\s+/g, ' ').trim();
      }
      return { qty, rest: str };
    }
  }
  return { qty: 1, rest: str };
}

// ─── CSV / TSV STRUCTURED PARSER ─────────────────────────────────────────────
// Detects tab-separated or comma-separated data with a header row
function parseStructuredText(text) {
  const lines = text.split(/\n/).map(l => l.trim()).filter(l => l.length > 0);
  if (lines.length < 2) return null;

  // Detect delimiter: tabs or commas (but not commas inside free-text)
  const firstLine = lines[0];
  const hasTabs = firstLine.includes('\t');
  const commaFields = firstLine.split(',').length;
  const delimiter = hasTabs ? '\t' : (commaFields >= 3 ? ',' : null);
  if (!delimiter) return null;

  // Parse header row
  const headerCells = firstLine.split(delimiter).map(c => c.trim());
  const colMap = resolveColumns(headerCells);

  // Must have at least a model column to be structured
  if (colMap.model < 0 && colMap.brand < 0) return null;

  // Verify at least 1 data row has device-like content
  const dataLines = lines.slice(1).filter(l => l.split(delimiter).length >= 2);
  if (!dataLines.length) return null;

  const HAS_DEVICE = /\b(dell|hp|hewlett|lenovo|apple|macbook|thinkpad|latitude|elitebook|surface|asus|acer|fujitsu|toshiba|samsung|microsoft|optiplex|probook|zbook|precision|vostro|inspiron|xps|ideapad|zenbook|vivobook|lifebook|portege|tecra|dynabook|travelmate|dragonfly|chromebook)\b/i;

  const devices = [];
  for (const line of dataLines) {
    const cells = line.split(delimiter).map(c => c.trim());
    const brandRaw = colMap.brand >= 0 ? cells[colMap.brand] || '' : '';
    const modelRaw = colMap.model >= 0 ? cells[colMap.model] || '' : '';
    const combined = (brandRaw + ' ' + modelRaw).trim();
    if (!combined || !HAS_DEVICE.test(combined)) continue;

    const cpuRaw   = colMap.cpu   >= 0 ? cells[colMap.cpu]   || '' : '';
    const ramRaw   = colMap.ram   >= 0 ? cells[colMap.ram]   || '' : '';
    const ssdRaw   = colMap.ssd   >= 0 ? cells[colMap.ssd]   || '' : '';
    const gradeRaw = colMap.grade >= 0 ? cells[colMap.grade] || '' : '';
    const qtyRaw   = colMap.qty   >= 0 ? cells[colMap.qty]   || '' : '';

    // Build model: prepend brand if not already in model
    let fullModel = modelRaw;
    if (brandRaw && !modelRaw.toLowerCase().startsWith(brandRaw.toLowerCase())) {
      fullModel = brandRaw + ' ' + modelRaw;
    }
    fullModel = fullModel.trim();

    // Parse specs
    const ram = extractRAM(ramRaw + 'GB') || extractRAM(ramRaw) || extractRAM(combined);
    const ssd = extractSSD(ssdRaw + 'GB') || extractSSD(ssdRaw) || extractSSD(combined);
    const cpu = extractCPU(cpuRaw) || extractCPU(combined);
    const grade = extractGrade(gradeRaw) || 'B';
    const qty = Math.max(parseInt(qtyRaw) || 1, 1);
    const brand = detectBrand(fullModel);
    const gen = inferGenFromModel(brand, fullModel);

    devices.push({ model: fullModel, brand, cpu, ram, ssd, grade, qty, quantity: qty, gen });
  }

  return devices.length > 0 ? devices : null;
}

// ─── CONCATENATED TABLE PARSER ──────────────────────────────────────────────
// Handles copy-pasted tables where all rows merge into one string:
// "#AssetQtySpec1Dell Latitude 542028i5-1145G7, 8GB, 256GB SSD..."
// Pattern: #[RowNum][Brand Model][Qty][Specs] repeating
function parseConcatenatedTable(text) {
  // Strip any preamble text before the actual data (e.g. "Test with this exact input:")
  const headerIdx = text.search(/#\s*Asset\s*Qty\s*Spec/i);
  const hashRowIdx = text.search(/#\d+[A-Z]/i);
  if (headerIdx >= 0) {
    text = text.slice(headerIdx);
  } else if (hashRowIdx >= 0) {
    text = text.slice(hashRowIdx);
  }

  // Detect: "#AssetQtySpec1Dell..." or "#1Dell...#2HP..." or bare "1Dell...2HP..."
  const hasHeader = /^#?\s*Asset\s*Qty\s*Spec/i.test(text);
  const hasHashRows = /#\d+\s*[A-Z]/i.test(text);
  if (!hasHeader && !hasHashRows) return null;

  // Strip header prefix
  let cleaned = text.replace(/^#?\s*Asset\s*Qty\s*Spec\s*/i, '');

  // Determine segment split pattern:
  // If segments start with #N, split on #N
  // If bare numbers (after header strip), split on number before uppercase brand
  let segments;
  if (/#\d+\s*[A-Z]/.test(cleaned)) {
    segments = cleaned.split(/(?=#\d+\s*[A-Z])/i).filter(s => s.trim().length > 0);
  } else {
    // Bare row numbers: "1Dell...2Dell..." — split before a digit that precedes a brand
    segments = cleaned.split(/(?=\d{1,3}(?:Dell|HP|Hewlett|Lenovo|Apple|Microsoft|Asus|Acer|Fujitsu|Toshiba|Samsung)\b)/i).filter(s => s.trim().length > 0);
  }
  if (segments.length < 2) return null;

  const HAS_DEVICE = /\b(dell|hp|hewlett|lenovo|apple|macbook|thinkpad|latitude|elitebook|surface|asus|acer|fujitsu|toshiba|optiplex|probook|zbook|precision|thinkcentre|ideapad|chromebook)\b/i;

  // Known brand+model patterns — match the full model name including trailing numbers
  // IMPORTANT: G-numbers use G\d (single digit) to avoid consuming qty digits in concatenated text
  const MODEL_PATTERNS = [
    // Dell: Latitude/Precision/Inspiron/XPS/Vostro XXXX [2-in-1], OptiPlex XXXX [SFF|MT]
    /^((?:Dell\s+)?(?:Latitude|Precision|Inspiron|XPS|Vostro)\s+\d{4}(?:\s+\d-in-\d)?)/i,
    /^((?:Dell\s+)?OptiPlex\s+\d{4}(?:\s+(?:SFF|MT|DT|Micro|Tower))?)/i,
    // HP: EliteBook/ProBook/ZBook XXX GX, EliteDesk/ProDesk XXX GX [SFF]
    /^((?:HP\s+)?(?:EliteBook|ProBook|ZBook|EliteDragonfly|Dragonfly|Folio)\s+\d{3}\s+G\d)/i,
    /^((?:HP\s+)?(?:EliteDesk|ProDesk)\s+\d{3}\s+G\d(?:\s+(?:SFF|MT|DM|Mini|Tower))?)/i,
    // Lenovo: ThinkPad TXX[s] GX / Gen X, ThinkCentre MXXx GX
    /^((?:Lenovo\s+)?ThinkPad\s+[A-Z]\d+[a-z]?\s+G\d)/i,
    /^((?:Lenovo\s+)?ThinkPad\s+[A-Z]\d+[a-z]?\s+Gen\s*\d)/i,
    /^((?:Lenovo\s+)?ThinkCentre\s+\w+\s+G\d)/i,
    /^((?:Lenovo\s+)?ThinkPad\s+[A-Z]\d+[a-z]?)/i,
    // Apple: MacBook Pro/Air XX M1/M2/M3 [Pro|Max|Ultra]
    /^((?:Apple\s+)?MacBook\s+(?:Pro|Air)\s+\d+\s+M\d(?:\s+(?:Pro|Max|Ultra))?)/i,
    /^((?:Apple\s+)?MacBook\s+(?:Pro|Air)\s+\d+)/i,
    /^((?:Apple\s+)?MacBook\s+(?:Pro|Air)\s+M\d)/i,
  ];

  const devices = [];
  for (const seg of segments) {
    // Strip leading #number or bare number FIRST (before HAS_DEVICE check)
    // because "7HP" has no word boundary between digit and brand
    const body = seg.replace(/^#?\d{1,3}/, '').trim();
    if (!HAS_DEVICE.test(body)) continue;

    // Try each brand+model pattern to find model boundary
    let modelRaw = null, remainder = null;
    for (const pat of MODEL_PATTERNS) {
      const m = body.match(pat);
      if (m) {
        modelRaw = m[1].trim();
        remainder = body.slice(m[0].length);
        break;
      }
    }

    if (!modelRaw) continue;

    // Remainder starts with qty (digits) then specs
    // e.g. "28i5-1145G7, 8GB, 256GB SSD" or "19i5-1135G7, 8GB..."
    // Or for Apple: "9M1 Pro 10-core, 16GB, 512GB SSD, 2021"
    const qtyMatch = remainder.match(/^(\d{1,4})/);
    const qty = qtyMatch ? parseInt(qtyMatch[1]) : 1;
    if (qty > 5000 || qty < 1) continue;

    const specsRaw = qtyMatch ? remainder.slice(qtyMatch[0].length).trim() : remainder.trim();

    const ram = extractRAM(specsRaw);
    const ssd = extractSSD(specsRaw);
    const cpu = extractCPU(specsRaw);
    const grade = extractGrade(specsRaw) || 'B';
    const brand = detectBrand(modelRaw);
    const gen = inferGenFromModel(brand, modelRaw);

    devices.push({ model: modelRaw, brand, cpu, ram, ssd, grade, qty, quantity: qty, gen });
  }

  return devices.length >= 2 ? devices : null;
}

// ─── TEXT INPUT PARSER ────────────────────────────────────────────────────────
function stripHtml(text) {
  if (!/<[a-z][\s\S]*>/i.test(text)) return text;
  return text
    .replace(/<br\s*\/?>/gi, '\n')
    .replace(/<\/tr>/gi, '\n')
    .replace(/<\/td>/gi, '\t')
    .replace(/<\/th>/gi, '\t')
    .replace(/<\/p>/gi, '\n')
    .replace(/<\/div>/gi, '\n')
    .replace(/<\/li>/gi, '\n')
    .replace(/<[^>]+>/g, '')
    .replace(/&nbsp;/gi, ' ')
    .replace(/&amp;/gi, '&')
    .replace(/&lt;/gi, '<')
    .replace(/&gt;/gi, '>')
    .replace(/&euro;/gi, '\u20ac')
    .replace(/&#\d+;/g, '')
    .replace(/\r\n/g, '\n');
}

function parseTextInput(text) {
  // ── Step -1: Strip HTML (email paste, HTML tables) ──
  text = stripHtml(text);

  // ── Step 0: Try structured CSV/TSV first ──
  const structured = parseStructuredText(text);
  if (structured) return structured;

  // ── Step 0b: Try concatenated table format (#1Brand Model...#2Brand Model...) ──
  const concatenated = parseConcatenatedTable(text);
  if (concatenated) return concatenated;

  // ── Step 0c: Extract total-units validation hint ──
  const totalUnitsMatch = text.match(/total\s*units?\s*[:=]?\s*(\d+)/i);
  const expectedTotal = totalUnitsMatch ? parseInt(totalUnitsMatch[1]) : null;

  // ── Step 1: Split into lines and filter junk ──
  const EMAIL_JUNK = /^(summarize this email|inbox|re:|fwd:|from:|to me|to:|subject:|sent:|cc:|planbit|caas\s*[-–]?\s*$|feb |mar |jan |apr |may |jun |jul |aug |sep |oct |nov |dec |http|analyseer|target|geef|caps|qwertzu|lot-factor|totaal:|total:|zusammenfassung|résumé|beste|hierbij|met vriendelijke|mvg|groet|graag\b|geachte\b|hi |dear |kind regards|with regards|hartelijke|hoogachtend|dhr\.|mevr\.|bedankt|thanks|cheers|sincerely|cordialement|mit freundlichen)/i;
  const EMAIL_SIG = /^[\w\s.,-]*@[\w.-]+\.\w{2,}|^\+?\d[\d\s()-]{7,}$|^(?:BV|Inc|Ltd|GmbH|B\.V\.|NV|AG|SAS|SARL|LLC|Corp)\s*\.?$/i;
  const HAS_DEVICE = /\b(dell|hp|hewlett|lenovo|apple|macbook|thinkpad|latitude|elitebook|surface|asus|acer|fujitsu|toshiba|samsung|microsoft|iphone|ipad|galaxy|yoga|optiplex|probook|zbook|spectre|envy|pavilion|chromebook|precision|vostro|inspiron|xps|ideapad|zenbook|vivobook|lifebook|portege|tecra|dynabook|travelmate|dragonfly)\b/i;
  const SECTION_HEADER = /^(APPLE|DELL|HP|LENOVO|TOSHIBA|ACER|ASUS|MICROSOFT)\s*[—–-]\s*\d+\s*(?:stuks?|st[üu]ck|pcs|units?|pieces?)/i;

  const lines = text
    .split(/[\n;]+/)
    .map(l => l.trim())
    .filter(l => l.length > 3)
    .filter(l => !EMAIL_JUNK.test(l) || HAS_DEVICE.test(l))
    .filter(l => !EMAIL_SIG.test(l) || HAS_DEVICE.test(l))
    .filter(l => !SECTION_HEADER.test(l))
    .filter(l => !/^total\s*units?\s*[:=]?\s*\d+/i.test(l));

  const devices = [];

  for (const line of lines) {
    if (!HAS_DEVICE.test(line)) continue;

    // Strip numbered-list prefix: "1. ", "2) "
    let cleaned = line.replace(/^\d{1,3}[\.\)]\s*/, '');
    // Strip bullet/dash prefix: "- ", "• ", "· ", "* "
    cleaned = cleaned.replace(/^[-\*\•\·–—]\s*/, '');

    // ── Extract quantity (from anywhere in line) ──
    const { qty, rest } = extractQuantity(cleaned);
    cleaned = rest;

    // ── Split into model and specs ──
    // Try middot separator first, then comma-separated parts
    let parts;
    if (cleaned.includes('·')) {
      parts = cleaned.split(/\s*·\s*/);
    } else if (cleaned.includes(',')) {
      parts = cleaned.split(/\s*,\s*/);
    } else {
      parts = [cleaned];
    }

    let modelRaw, specsPart;
    if (parts.length >= 2) {
      // First part is likely the model; rest are specs
      // But sometimes first part is brand+model+specs without separator
      // Find where model ends: first part that looks like specs
      let splitAt = 1;
      for (let i = 1; i < parts.length; i++) {
        const p = parts[i].toLowerCase();
        if (/\b(i[3579]|ryzen|m[123]\b|celeron|\d+\s*gb|grade|conditie|kwaliteit|qty|ssd|ram)/i.test(p)) {
          splitAt = i;
          break;
        }
      }
      modelRaw = parts.slice(0, splitAt).join(' ').trim();
      specsPart = parts.slice(splitAt).join(' ');
    } else {
      // No separator — split on first spec boundary
      const modelStop = /\b(i[3579][-\s]\d|Ryzen|M[123]\s|FHD|UHD|\d+\s*GB\b)/i;
      const stopMatch = cleaned.match(modelStop);
      if (stopMatch) {
        modelRaw = cleaned.slice(0, stopMatch.index).trim();
        specsPart = cleaned.slice(stopMatch.index);
      } else {
        modelRaw = cleaned;
        specsPart = '';
      }
    }

    // ── Extract parenthesized content as specs before cleaning ──
    const parenMatch = modelRaw.match(/\(([^)]+)\)/);
    if (parenMatch) {
      specsPart = (specsPart + ' ' + parenMatch[1]).trim();
    }

    // ── Clean model name ──
    modelRaw = modelRaw
      .replace(/^[-\*\•\·\s]+/, '')         // leading bullets
      .replace(/[,;]+$/, '')                  // trailing commas
      .replace(/\s*\([^)]*\)\s*/g, ' ')       // remove parenthesized content
      .replace(/\s*\(\s*$/, '')               // trailing orphan "("
      .replace(/\s+/g, ' ')
      .trim();

    if (!HAS_DEVICE.test(modelRaw)) {
      // modelRaw lost the device keyword due to aggressive cleaning — try full line
      if (!HAS_DEVICE.test(cleaned)) continue;
      modelRaw = cleaned.replace(/\s+/g, ' ').trim();
    }
    if (modelRaw.length < 3) continue;

    // ── Parse specs from specsPart + full line ──
    const specsStr = specsPart || cleaned;
    const fullStr = cleaned;

    // RAM/SSD from "8/256" slash notation (with or without GB)
    const ramSsdMatch = (specsStr + ' ' + fullStr).match(/\b(\d+)\s*(?:GB)?\s*\/\s*(\d+)\s*(?:GB)?/);
    let ram = null, ssd = null;
    if (ramSsdMatch) {
      const v1 = parseInt(ramSsdMatch[1]);
      const v2 = parseInt(ramSsdMatch[2]);
      if (v1 <= 128 && v2 >= 64) { ram = v1; ssd = v2; }
      else { ram = v1; ssd = v2; }
    }
    if (!ram) ram = extractRAM(specsStr) || extractRAM(fullStr);
    if (!ssd) ssd = extractSSD(specsStr) || extractSSD(fullStr);

    const cpu = extractCPU(specsStr) || extractCPU(fullStr);
    const grade = extractGrade(specsStr) || extractGrade(fullStr);
    const gen = inferGenFromModel('', modelRaw) || inferGenFromModel('', fullStr);
    const brand = detectBrand(modelRaw);

    // ── Clean up model: remove specs that leaked into model name ──
    let modelClean = modelRaw
      .replace(/\b\d+\s*GB\b/gi, '')             // remove "8GB", "256GB"
      .replace(/\b\d+\s*\/\s*\d+\b/g, '')        // remove "8/256"
      .replace(/\b[A-D]\s*-?\s*grade\b/gi, '')   // remove "b-grade"
      .replace(/\bgrade\s*[A-D]\d?\b/gi, '')      // remove "Grade B"
      .replace(/\bconditie\s*[A-D]\d?\b/gi, '')   // remove "conditie A"
      .replace(/\bkwaliteit\s*[A-D]\d?\b/gi, '')  // remove "kwaliteit A"
      .replace(/\s+/g, ' ')
      .trim();
    if (modelClean.length >= 3 && HAS_DEVICE.test(modelClean)) {
      modelRaw = modelClean;
    }

    devices.push({ model: modelRaw, brand, cpu, ram, ssd, grade, qty, quantity: qty, gen });
  }

  // ── Deduplicate by model name ──
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
      if (!existing.brand && d.brand) existing.brand = d.brand;
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
    case 'BLANCCO_DETAIL': {
      const devices = parseBlanccoDetail(wb);
      if (devices && devices.length) return { devices, format };
      // Fallback to ARS if Blancco parse fails
      return { devices: parseARS(wb), format: 'ARS' };
    }
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
  extractQuantity,
  detectBrand,
  inferGenFromModel,
  resolveColumns,
  normalizeBlanccoModel,
  parseBlanccoDetail,
  parseConcatenatedTable,
};
