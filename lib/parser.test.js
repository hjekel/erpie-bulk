'use strict';

const {
  extractRAM,
  extractSSD,
  extractCPU,
  extractGrade,
  extractQuantity,
  detectBrand,
  inferGenFromModel,
  resolveColumns,
  parseTextInput,
} = require('./parser');

// ─── Test runner ─────────────────────────────────────────────────────────────
let passed = 0, failed = 0;
function eq(actual, expected, label) {
  const a = JSON.stringify(actual);
  const e = JSON.stringify(expected);
  if (a === e) {
    passed++;
  } else {
    failed++;
    console.error(`  FAIL: ${label}\n    expected ${e}\n    got      ${a}`);
  }
}

// ═══════════════════════════════════════════════════════════════════════════════
// extractRAM
// ═══════════════════════════════════════════════════════════════════════════════
console.log('--- extractRAM ---');
eq(extractRAM('16GB'), '16GB', 'plain 16GB');
eq(extractRAM('16 GB DDR4'), '16GB', '16 GB with DDR4');
eq(extractRAM('32GB RAM'), '32GB', '32GB RAM');
eq(extractRAM('4gb'), '4GB', 'lowercase 4gb');
eq(extractRAM('8GB SSD'), '', 'should not match 8GB SSD');
eq(extractRAM('128GB'), '128GB', '128GB RAM');
eq(extractRAM('Dell Latitude 5430'), '', 'no RAM in model name');
eq(extractRAM('i5-1235U 16GB 512GB SSD'), '16GB', 'extract 16GB before SSD');

// ═══════════════════════════════════════════════════════════════════════════════
// extractSSD
// ═══════════════════════════════════════════════════════════════════════════════
console.log('--- extractSSD ---');
eq(extractSSD('256GB SSD'), '256GB', '256GB SSD');
eq(extractSSD('512 GB NVMe'), '512GB', '512 GB NVMe');
eq(extractSSD('1TB SSD'), '1024GB', '1TB → 1024GB');
eq(extractSSD('SSD 128GB'), '128GB', 'SSD prefix');
eq(extractSSD('2TB M.2'), '2048GB', '2TB → 2048GB');
eq(extractSSD('480GB SATA'), '480GB', '480GB SATA');
eq(extractSSD('no storage'), '', 'no match');

// ═══════════════════════════════════════════════════════════════════════════════
// extractCPU
// ═══════════════════════════════════════════════════════════════════════════════
console.log('--- extractCPU ---');
eq(extractCPU('Intel Core i5-1235U'), 'Core i5-1235U', 'Core i5-1235U');
eq(extractCPU('i7 1265U'), 'i7 1265U', 'i7 without dash');
eq(extractCPU('Core i3-10110U vPro'), 'Core i3-10110U', 'Core i3-10110U');
eq(extractCPU('AMD Ryzen 5 5600U'), 'Ryzen 5 5600U', 'Ryzen 5');
eq(extractCPU('Apple M1 Pro'), 'M1 Pro', 'M1 Pro');
eq(extractCPU('M2'), 'M2', 'M2');
eq(extractCPU('Celeron N4020'), 'Celeron N4020', 'Celeron');
eq(extractCPU('Xeon E-2286M'), 'Xeon E-2286M', 'Xeon');
eq(extractCPU('Dell Latitude 5430'), '', 'no CPU in model');

// ═══════════════════════════════════════════════════════════════════════════════
// extractGrade — standard codes
// ═══════════════════════════════════════════════════════════════════════════════
console.log('--- extractGrade (standard) ---');
eq(extractGrade('A'), 'A', 'simple A');
eq(extractGrade('B'), 'B', 'simple B');
eq(extractGrade('C'), 'C', 'simple C');
eq(extractGrade('D'), 'D', 'simple D');
eq(extractGrade('A1'), 'A1', 'A1');
eq(extractGrade('B4'), 'B4', 'B4');
eq(extractGrade('C6'), 'C6', 'C6');
eq(extractGrade('P7'), 'P7', 'P7 parts');
eq(extractGrade('X9'), 'X9', 'X9 defect');
eq(extractGrade('Grade A'), 'A', 'Grade A prefixed');
eq(extractGrade('Condition: B'), 'B', 'Condition: B');

// ═══════════════════════════════════════════════════════════════════════════════
// extractGrade — natural language (EN/NL/DE)
// ═══════════════════════════════════════════════════════════════════════════════
console.log('--- extractGrade (natural language) ---');
eq(extractGrade('Like New'), 'A', 'Like New → A');
eq(extractGrade('excellent'), 'A', 'excellent → A');
eq(extractGrade('mint condition'), 'A', 'mint → A');
eq(extractGrade('neuwertig'), 'A', 'neuwertig (DE) → A');
eq(extractGrade('nieuwstaat'), 'A', 'nieuwstaat (NL) → A');
eq(extractGrade('good'), 'B', 'good → B');
eq(extractGrade('refurbished'), 'B', 'refurbished → B');
eq(extractGrade('gut'), 'B', 'gut (DE) → B');
eq(extractGrade('goed'), 'B', 'goed (NL) → B');
eq(extractGrade('working'), 'B', 'working → B');
eq(extractGrade('fair'), 'C', 'fair → C');
eq(extractGrade('acceptable'), 'C', 'acceptable → C');
eq(extractGrade('gebruikt'), 'C', 'gebruikt (NL) → C');
eq(extractGrade('gebraucht'), 'C', 'gebraucht (DE) → C');
eq(extractGrade('poor'), 'D', 'poor → D');
eq(extractGrade('defect'), 'D', 'defect → D');
eq(extractGrade('kapot'), 'D', 'kapot (NL) → D');
eq(extractGrade('Klasse A'), 'A', 'Klasse A (DE) → A');
eq(extractGrade('Conditie: B'), 'B', 'Conditie (NL) → B');
eq(extractGrade('Zustand: C'), 'C', 'Zustand (DE) → C');
eq(extractGrade(''), '', 'empty string');

// ═══════════════════════════════════════════════════════════════════════════════
// extractQuantity
// ═══════════════════════════════════════════════════════════════════════════════
console.log('--- extractQuantity ---');
eq(extractQuantity('60x Dell Latitude 5430').qty, 60, '60x prefix');
eq(extractQuantity('60x Dell Latitude 5430').rest, 'Dell Latitude 5430', '60x rest');
eq(extractQuantity('60 x Dell Latitude 5430').qty, 60, '60 x prefix (space)');
eq(extractQuantity('Dell Latitude 5430 · 50x').qty, 50, '50x suffix');
eq(extractQuantity('Dell Latitude 5430 - 25 stuks').qty, 25, '25 stuks');
eq(extractQuantity('Dell Latitude 5430 - 10 stück').qty, 10, '10 stück (DE)');
eq(extractQuantity('Dell Latitude 5430 30 pcs').qty, 30, '30 pcs');
eq(extractQuantity('Dell Latitude 5430 15 pieces').qty, 15, '15 pieces');
eq(extractQuantity('Dell Latitude 5430 (20)').qty, 20, '(20) parens');
eq(extractQuantity('qty: 42 Dell Latitude 5430').qty, 42, 'qty: 42');
eq(extractQuantity('Aantal: 8 Dell Latitude 5430').qty, 8, 'Aantal: 8 (NL)');
eq(extractQuantity('Anzahl: 5 Dell Latitude 5430').qty, 5, 'Anzahl: 5 (DE)');
eq(extractQuantity('Dell Latitude 5430').qty, 1, 'no quantity → 1');
eq(extractQuantity('12 units HP EliteBook').qty, 12, '12 units');
eq(extractQuantity('HP EliteBook 840 G7 × 100').qty, 100, '× 100 suffix');

// ═══════════════════════════════════════════════════════════════════════════════
// detectBrand
// ═══════════════════════════════════════════════════════════════════════════════
console.log('--- detectBrand ---');
eq(detectBrand('Dell Latitude 5430'), 'Dell', 'Dell prefix');
eq(detectBrand('HP EliteBook 840 G7'), 'HP', 'HP prefix');
eq(detectBrand('Lenovo ThinkPad T14'), 'Lenovo', 'Lenovo prefix');
eq(detectBrand('MacBook Pro 16'), 'Apple', 'MacBook → Apple');
eq(detectBrand('ThinkPad T480'), 'Lenovo', 'ThinkPad → Lenovo');
eq(detectBrand('EliteBook 840 G7'), 'HP', 'EliteBook → HP');
eq(detectBrand('Latitude 5430'), 'Dell', 'Latitude → Dell');
eq(detectBrand('Surface Pro 7'), 'Microsoft', 'Surface → Microsoft');
eq(detectBrand('ZenBook 14'), 'ASUS', 'ZenBook → ASUS');
eq(detectBrand('LifeBook U747'), 'Fujitsu', 'LifeBook → Fujitsu');
eq(detectBrand('Hewlett-Packard EliteBook'), 'HP', 'Hewlett-Packard → HP');
eq(detectBrand('XPS 13 9310'), 'Dell', 'XPS → Dell');
eq(detectBrand('Optiplex 5060'), 'Dell', 'Optiplex → Dell');
eq(detectBrand('IdeaPad 5'), 'Lenovo', 'IdeaPad → Lenovo');
eq(detectBrand('Aspire 5'), 'Acer', 'Aspire → Acer');
eq(detectBrand('TravelMate P2'), 'Acer', 'TravelMate → Acer');
eq(detectBrand('Dynabook Portege'), 'Toshiba', 'Dynabook → Toshiba');
eq(detectBrand(''), '', 'empty string');
eq(detectBrand('Unknown Device 123'), '', 'unknown device');

// ═══════════════════════════════════════════════════════════════════════════════
// inferGenFromModel
// ═══════════════════════════════════════════════════════════════════════════════
console.log('--- inferGenFromModel ---');
eq(inferGenFromModel('Dell', 'Latitude 5430'), 'Gen12', 'Latitude 5430 → Gen12');
eq(inferGenFromModel('Dell', 'Latitude 5410'), 'Gen10', 'Latitude 5410 → Gen10');
eq(inferGenFromModel('Dell', 'Latitude 7490'), 'Gen8', 'Latitude 7490 → Gen8');
eq(inferGenFromModel('HP', 'EliteBook 840 G7'), 'Gen10', 'EliteBook 840 G7 → Gen10');
eq(inferGenFromModel('HP', 'EliteBook 840 G10'), 'Gen13', 'EliteBook 840 G10 → Gen13');
eq(inferGenFromModel('Lenovo', 'ThinkPad T14 Gen 3'), 'Gen12', 'ThinkPad T14 Gen 3 → Gen12');
eq(inferGenFromModel('HP', '250 G8'), 'Gen11', '250 G8 → Gen11');
eq(inferGenFromModel('', 'Chromebook'), 'Gen8', 'Chromebook → Gen8');

// ═══════════════════════════════════════════════════════════════════════════════
// resolveColumns — multilingual
// ═══════════════════════════════════════════════════════════════════════════════
console.log('--- resolveColumns (multilingual) ---');
{
  // English headers
  const en = resolveColumns(['Brand', 'Model', 'CPU', 'RAM', 'SSD', 'Grade', 'Qty']);
  eq(en.brand, 0, 'EN: Brand');
  eq(en.model, 1, 'EN: Model');
  eq(en.cpu, 2, 'EN: CPU');
  eq(en.ram, 3, 'EN: RAM');
  eq(en.ssd, 4, 'EN: SSD');
  eq(en.grade, 5, 'EN: Grade');
  eq(en.qty, 6, 'EN: Qty');

  // Dutch headers
  const nl = resolveColumns(['Merk', 'Model', 'Processor', 'Geheugen', 'Opslag', 'Conditie', 'Aantal']);
  eq(nl.brand, 0, 'NL: Merk');
  eq(nl.model, 1, 'NL: Model');
  eq(nl.cpu, 2, 'NL: Processor');
  eq(nl.ram, 3, 'NL: Geheugen');
  eq(nl.ssd, 4, 'NL: Opslag');
  eq(nl.grade, 5, 'NL: Conditie');
  eq(nl.qty, 6, 'NL: Aantal');

  // German headers
  const de = resolveColumns(['Hersteller', 'Modell', 'Prozessor', 'Arbeitsspeicher', 'Festplatte', 'Zustand', 'Anzahl']);
  eq(de.brand, 0, 'DE: Hersteller');
  eq(de.model, 1, 'DE: Modell');
  eq(de.cpu, 2, 'DE: Prozessor');
  eq(de.ram, 3, 'DE: Arbeitsspeicher');
  eq(de.ssd, 4, 'DE: Festplatte');
  eq(de.grade, 5, 'DE: Zustand');
  eq(de.qty, 6, 'DE: Anzahl');

  // French headers
  const fr = resolveColumns(['Fabricant', 'Modèle', 'Processeur', 'Mémoire', 'Stockage', 'État', 'Quantité']);
  eq(fr.brand, 0, 'FR: Fabricant');
  eq(fr.model, 1, 'FR: Modèle → modele');
  eq(fr.cpu, 2, 'FR: Processeur');
  eq(fr.ram, 3, 'FR: Mémoire');
  eq(fr.ssd, 4, 'FR: Stockage');
  eq(fr.grade, 5, 'FR: État');
  eq(fr.qty, 6, 'FR: Quantité');

  // Mixed / variant headers
  const mixed = resolveColumns(['OEM', 'Device_Name', 'Specs', 'Installed_Memory', 'HDD_Size', 'Cosmetic Grade', 'Units']);
  eq(mixed.brand, 0, 'Mixed: OEM');
  // Device_Name → fuzzy matches 'device' in model aliases (partial starts-with)
  eq(mixed.ram, 3, 'Mixed: Installed_Memory');
  eq(mixed.ssd, 4, 'Mixed: HDD_Size');
  eq(mixed.qty, 6, 'Mixed: Units');
}

// ═══════════════════════════════════════════════════════════════════════════════
// parseTextInput — free-form text
// ═══════════════════════════════════════════════════════════════════════════════
console.log('--- parseTextInput (free-form) ---');
{
  // Simple "60x Model specs" format
  const r1 = parseTextInput('60x Dell Latitude 7430 i7 32GB');
  eq(r1.length, 1, 'one device parsed');
  eq(r1[0].qty, 60, 'qty = 60');
  eq(r1[0].ram, '32GB', 'ram = 32GB');
  eq(r1[0].brand, 'Dell', 'brand = Dell');

  // Multiple lines with mixed formats
  const multi = parseTextInput(`
    25x HP EliteBook 840 G7 · i5-10210U · 16GB · 256GB SSD · Grade B
    10 stuks Dell Latitude 5430 i5-1235U 16GB 512GB SSD
    MacBook Pro 16 M1 Pro 16GB/512GB (5)
  `);
  eq(multi.length, 3, 'three devices');
  eq(multi[0].qty, 25, 'HP qty=25');
  eq(multi[0].brand, 'HP', 'HP brand detected');
  eq(multi[0].grade, 'B', 'HP grade=B');
  eq(multi[1].qty, 10, 'Dell qty=10 (stuks)');
  eq(multi[1].brand, 'Dell', 'Dell brand');
  eq(multi[2].qty, 5, 'MacBook qty=5 (parens)');
  eq(multi[2].brand, 'Apple', 'MacBook → Apple');

  // German-style input
  const de = parseTextInput('15 Stück Lenovo ThinkPad T14 Gen 2 · i5-1135G7 · 16GB/256GB · Zustand: B');
  eq(de.length, 1, 'DE: one device');
  eq(de[0].qty, 15, 'DE: qty=15 Stück');
  eq(de[0].brand, 'Lenovo', 'DE: brand=Lenovo');

  // Dutch-style
  const nl = parseTextInput('HP EliteBook 840 G8 - 20 stuks');
  eq(nl.length, 1, 'NL: one device');
  eq(nl[0].qty, 20, 'NL: qty=20 stuks');

  // Deduplication
  const dedup = parseTextInput(`
    Dell Latitude 5430 i5 16GB 256GB
    Dell Latitude 5430 i5 16GB 256GB
  `);
  eq(dedup.length, 1, 'dedup: merged to 1');
  eq(dedup[0].qty, 2, 'dedup: qty=2 after merge');

  // Edge: numbered list
  const numbered = parseTextInput(`
    1. 5x Dell Latitude 5420
    2. 3x HP EliteBook 840 G7
    3. 2x Lenovo ThinkPad T14
  `);
  eq(numbered.length, 3, 'numbered list: 3 devices');
  eq(numbered[0].qty, 5, 'numbered #1 qty=5');
  eq(numbered[1].qty, 3, 'numbered #2 qty=3');
  eq(numbered[2].qty, 2, 'numbered #3 qty=2');

  // Empty/junk input
  eq(parseTextInput('').length, 0, 'empty input');
  eq(parseTextInput('no devices here at all').length, 0, 'no devices');
  eq(parseTextInput('Re: FWD: meeting notes').length, 0, 'email junk filtered');
}

// ═══════════════════════════════════════════════════════════════════════════════
// Summary
// ═══════════════════════════════════════════════════════════════════════════════
console.log(`\n${'='.repeat(60)}`);
console.log(`Tests: ${passed} passed, ${failed} failed, ${passed + failed} total`);
if (failed > 0) process.exit(1);
console.log('All tests passed!');
