'use strict';

const ExcelJS = require('exceljs');
const { detectBrand } = require('./parser');

// ═══════════════════════════════════════════════════════════════════════════════
// PlanBit ARS Excel Export — matches exact format of ARS_Quotation template
// 3 tabs: Quote Request | Customer Inventory and Details | PlanBit conditions
// ═══════════════════════════════════════════════════════════════════════════════

const ORANGE_FILL = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFF9900' } };
const ARIAL_10 = { name: 'Arial', size: 10 };
const ARIAL_10_BOLD = { name: 'Arial', size: 10, bold: true };

function euroFmt(cell) { cell.numFmt = '#,##0.00'; }

/**
 * Generate PlanBit ARS format Excel
 * @param {Array} results - Analyzed device results from pricing engine
 * @param {Object} summary - Summary from analyzeDevices
 * @param {string} dealName - Deal/customer name
 * @returns {Promise<Buffer>} Excel buffer
 */
async function generateExcel(results, summary, dealName) {
  const workbook = new ExcelJS.Workbook();
  workbook.creator = 'ERPIE PriceFinder';
  workbook.created = new Date();

  const dateStr = new Date().toLocaleDateString('en-US', { month: 'long', year: 'numeric' });
  const totalQty = summary.total || results.reduce((s, r) => s + (r.qty || 1), 0);
  const totalValue = summary.totalValue || 0;

  // ═══════════════════════════════════════════════════════════════════════════
  // TAB 1: Quote Request
  // ═══════════════════════════════════════════════════════════════════════════
  const qr = workbook.addWorksheet('Quote Request');

  // Column widths matching reference
  qr.getColumn(1).width = 43;
  qr.getColumn(2).width = 30;
  qr.getColumn(3).width = 25;
  qr.getColumn(4).width = 18;
  qr.getColumn(5).width = 37;

  // Row 1: empty (merged header area in reference)
  // Row 2: Title bar
  const titleCell = qr.getCell('A2');
  titleCell.value = 'Request for Asset Recovery Quotation';
  titleCell.font = ARIAL_10;
  titleCell.fill = ORANGE_FILL;
  titleCell.alignment = { horizontal: 'left' };
  for (let c = 2; c <= 5; c++) {
    const cell = qr.getCell(2, c);
    cell.fill = ORANGE_FILL;
  }

  // Metadata rows
  const meta = [
    ['Date', dateStr],
    ['Customer', dealName],
    ['Contact', ''],
    ['Client name reference & location', ''],
    ['Country', 'NETHERLANDS'],
    ['Contact', 'joep@planbit.nl'],
    ['Tel', '+31 (0) 23 303 0085'],
    ['E-mail', 'joep@planbit.nl'],
    ['Web', 'www.PlanBit.nl'],
  ];
  meta.forEach(([label, value], i) => {
    const row = i + 4; // start at row 4
    const lCell = qr.getCell(row, 1);
    lCell.value = label;
    lCell.font = ARIAL_10_BOLD;
    lCell.alignment = { horizontal: 'left' };
    const vCell = qr.getCell(row, 2);
    vCell.value = value;
    vCell.font = label === 'Customer' ? ARIAL_10_BOLD : ARIAL_10;
  });

  // Service Cost section
  qr.getCell('A15').value = 'Service Cost';
  qr.getCell('A15').font = ARIAL_10_BOLD;

  // Service cost headers
  const scHeaders = ['Asset Type', 'Quantity', 'Service Charge per item', 'Total cost to'];
  scHeaders.forEach((h, i) => {
    const cell = qr.getCell(16, i + 1);
    cell.value = h;
    cell.font = ARIAL_10_BOLD;
  });

  // Count asset types
  const assetTypes = {};
  for (const r of results) {
    const type = 'NOTEBOOK'; // default; could be extended
    assetTypes[type] = (assetTypes[type] || 0) + (r.qty || 1);
  }

  let serviceRow = 17;
  const serviceCost = 10; // €10 per unit
  for (const [type, qty] of Object.entries(assetTypes)) {
    qr.getCell(serviceRow, 1).value = type;
    qr.getCell(serviceRow, 1).font = ARIAL_10;
    qr.getCell(serviceRow, 2).value = qty;
    qr.getCell(serviceRow, 2).font = ARIAL_10;
    qr.getCell(serviceRow, 3).value = serviceCost;
    qr.getCell(serviceRow, 3).font = ARIAL_10;
    qr.getCell(serviceRow, 4).value = qty * serviceCost;
    qr.getCell(serviceRow, 4).font = ARIAL_10;
    serviceRow++;
  }

  // Total row
  serviceRow++;
  qr.getCell(serviceRow, 1).value = 'Total';
  qr.getCell(serviceRow, 1).font = ARIAL_10_BOLD;
  qr.getCell(serviceRow, 2).value = totalQty;
  qr.getCell(serviceRow, 2).font = ARIAL_10_BOLD;
  qr.getCell(serviceRow, 4).value = totalQty * serviceCost;
  qr.getCell(serviceRow, 4).font = ARIAL_10_BOLD;

  // Logistics
  serviceRow += 2;
  qr.getCell(serviceRow, 1).value = 'Estimated Logistics';
  qr.getCell(serviceRow, 1).font = ARIAL_10;
  qr.getCell(serviceRow, 2).value = totalQty;
  qr.getCell(serviceRow, 2).font = ARIAL_10;
  qr.getCell(serviceRow, 3).value = 'Pick, Pack, Ship';
  qr.getCell(serviceRow, 3).font = ARIAL_10;

  // Estimated sales price
  serviceRow += 2;
  qr.getCell(serviceRow, 1).value = 'Estimated Sales price (see 2nd Tab) *';
  qr.getCell(serviceRow, 1).font = ARIAL_10_BOLD;
  qr.getCell(serviceRow, 3).value = dateStr;
  qr.getCell(serviceRow, 3).font = ARIAL_10;
  qr.getCell(serviceRow, 4).value = totalValue;
  qr.getCell(serviceRow, 4).font = ARIAL_10_BOLD;
  euroFmt(qr.getCell(serviceRow, 4));

  // Estimated nett return (70%)
  serviceRow += 2;
  qr.getCell(serviceRow, 1).value = 'Estimated Nett return';
  qr.getCell(serviceRow, 1).font = ARIAL_10_BOLD;
  qr.getCell(serviceRow, 2).value = '70%';
  qr.getCell(serviceRow, 2).font = ARIAL_10;
  qr.getCell(serviceRow, 4).value = totalValue * 0.7;
  qr.getCell(serviceRow, 4).font = ARIAL_10_BOLD;
  euroFmt(qr.getCell(serviceRow, 4));

  // Footer disclaimers
  serviceRow += 2;
  const disclaimers = [
    '* All devices need to be removed from MMS, Splash screens if present have been removed from BIOS, Apple devices have been removed from iCloud & MMS',
    '* PlanBit reserves the right to charge \u20ac 7.00,- (additional service fee) per asset if BIOS or MDM locks are present!',
    '* If notebooks are without correct and working adapter a deduction of \u20ac 15 will be applicable',
  ];
  for (const text of disclaimers) {
    qr.getCell(serviceRow, 1).value = text;
    qr.getCell(serviceRow, 1).font = { name: 'Arial', size: 9, italic: true };
    serviceRow++;
  }

  // Generated-by footer
  serviceRow += 1;
  qr.getCell(serviceRow, 1).value = `Generated by ERPIE PriceFinder \u00b7 ${new Date().toISOString().slice(0, 10)}`;
  qr.getCell(serviceRow, 1).font = { name: 'Arial', size: 8, italic: true, color: { argb: 'FF999999' } };

  // ═══════════════════════════════════════════════════════════════════════════
  // TAB 2: Customer Inventory and Details
  // ═══════════════════════════════════════════════════════════════════════════
  const inv = workbook.addWorksheet('Customer Inventory and Details');

  // Column widths matching reference exactly
  inv.getColumn(1).width = 13.5;  // Product
  inv.getColumn(2).width = 11.5;  // Brand
  inv.getColumn(3).width = 20.5;  // Model
  inv.getColumn(4).width = 63;    // Specifications Processor
  inv.getColumn(5).width = 9;     // Remarks
  inv.getColumn(6).width = 6;     // Qty
  inv.getColumn(7).width = 15.3;  // Sales Price Indication
  inv.getColumn(8).width = 15;    // (empty spacer)
  inv.getColumn(9).width = 16;    // Sum

  // Row 1: empty
  // Row 2: "Configuration" header
  const confCell = inv.getCell('A2');
  confCell.value = 'Configuration';
  confCell.font = ARIAL_10_BOLD;
  confCell.fill = ORANGE_FILL;

  // Row 3: Column headers
  const invHeaders = [
    'Product', 'Brand', 'Model', 'Specifications Processor',
    'Remarks', 'Qty', `Sales Price Indication\n${dateStr}`, '', 'Sum',
  ];
  invHeaders.forEach((h, i) => {
    const cell = inv.getCell(3, i + 1);
    cell.value = h;
    cell.font = ARIAL_10_BOLD;
    cell.fill = ORANGE_FILL;
    cell.alignment = { horizontal: 'left', vertical: 'bottom', wrapText: true };
  });

  // ─── Summary rows (unique model groups) ────────────────────────────────
  // Group results by brand+model
  const groups = new Map();
  for (const d of results) {
    const brand = d.brand || detectBrand(d.model || '') || '';
    const brandUpper = brand.toUpperCase();
    // Clean model: remove brand prefix if present
    let modelClean = (d.model || '').trim();
    if (modelClean.toLowerCase().startsWith(brand.toLowerCase())) {
      modelClean = modelClean.slice(brand.length).trim();
    }
    const key = `${brandUpper}|${modelClean}`;

    if (!groups.has(key)) {
      // Build specifications string like the reference: "CPU - SSDGB SSD - RAMGB"
      const cpuStr = d.cpu || (d.gen ? `Inferred ${d.gen}` : '');
      const ssdStr = d.ssdGb ? `${d.ssdGb}GB SSD` : '0GB SSD';
      const ramStr = d.ramGb ? `${d.ramGb}GB` : '0GB';
      const specs = [cpuStr, ssdStr, ramStr].filter(Boolean).join(' - ');

      groups.set(key, {
        product: 'NOTEBOOK',
        brand: brandUpper,
        model: modelClean,
        specs,
        grade: d.grade || 'B',
        gen: d.gen || '',
        status: d.status || '',
        unitPrice: d.advisedPrice || 0,
        qty: 0,
      });
    }
    groups.get(key).qty += (d.qty || 1);
  }

  // Sort groups: by brand, then model
  const sortedGroups = [...groups.values()].sort((a, b) => {
    const brandCmp = a.brand.localeCompare(b.brand);
    if (brandCmp !== 0) return brandCmp;
    return a.model.localeCompare(b.model);
  });

  let dataRow = 4;
  for (const g of sortedGroups) {
    const row = inv.getRow(dataRow);
    row.getCell(1).value = g.product;
    row.getCell(2).value = g.brand;
    row.getCell(3).value = g.model;
    row.getCell(4).value = g.specs;
    row.getCell(5).value = ''; // Remarks
    row.getCell(6).value = g.qty;
    row.getCell(7).value = g.unitPrice;
    euroFmt(row.getCell(7));
    // Sum = Qty * Unit Price (formula)
    row.getCell(9).value = { formula: `F${dataRow}*G${dataRow}` };
    euroFmt(row.getCell(9));

    // Style all cells
    for (let c = 1; c <= 9; c++) {
      row.getCell(c).font = ARIAL_10;
      row.getCell(c).alignment = { horizontal: 'left' };
    }
    dataRow++;
  }

  // ─── Totals row ────────────────────────────────────────────────────────
  dataRow++; // blank row
  const totalRow = inv.getRow(dataRow);
  totalRow.getCell(6).value = totalQty;
  totalRow.getCell(6).font = ARIAL_10_BOLD;
  totalRow.getCell(9).value = { formula: `SUM(I4:I${dataRow - 2})` };
  totalRow.getCell(9).font = ARIAL_10_BOLD;
  euroFmt(totalRow.getCell(9));

  // ─── Blank rows + "Original Input" label ───────────────────────────────
  dataRow += 4;
  inv.getCell(dataRow, 1).value = 'Original Input';
  inv.getCell(dataRow, 1).font = ARIAL_10_BOLD;
  inv.getCell(dataRow, 1).fill = ORANGE_FILL;
  dataRow++;

  // Detail column headers
  const detailHeaders = [
    'Facility', 'Category', 'Asset ID', 'Part Number',
    'Mfg', 'Model', 'Model (Blancco)', 'Serial Number',
    'CPU', 'Number of Processors', 'CPU Speed\n(MHz)', 'Memory\n(GB)',
    'HDD\n(GB)', 'HDD Count', 'Monitor Size', 'Test Status',
    'Cosmetic Notes', 'Missing Notes', 'Grade', 'Battery Present',
    'AC Adapter', 'Curr. mAh/OEM (%)', 'Keyboard Language', 'Graphic Card',
  ];

  // Expand columns for detail section (cols 10+ need width)
  for (let c = 10; c <= detailHeaders.length; c++) {
    inv.getColumn(c).width = 16;
  }

  detailHeaders.forEach((h, i) => {
    const cell = inv.getCell(dataRow, i + 1);
    cell.value = h;
    cell.font = ARIAL_10_BOLD;
    cell.fill = ORANGE_FILL;
    cell.alignment = { horizontal: 'left', wrapText: true };
  });
  dataRow++;

  // ─── Detail rows (all individual assets) ───────────────────────────────
  for (const d of results) {
    const brand = d.brand || detectBrand(d.model || '') || '';
    const brandUpper = brand.toUpperCase();
    let modelClean = (d.model || '').trim();
    if (modelClean.toLowerCase().startsWith(brand.toLowerCase())) {
      modelClean = modelClean.slice(brand.length).trim();
    }
    const cpu = d.cpu || '';
    const ramGb = d.ramGb || 0;
    const ssdGb = d.ssdGb || 0;
    const qty = d.qty || 1;
    const grade = d.grade || 'B';

    // Write one row per individual unit (expand qty)
    const expand = Math.min(qty, 500);
    for (let u = 0; u < expand; u++) {
      const row = inv.getRow(dataRow);
      row.getCell(1).value = dealName; // Facility
      row.getCell(2).value = 'NOTEBOOKS'; // Category
      row.getCell(3).value = d.serial || ''; // Asset ID
      row.getCell(4).value = ''; // Part Number
      row.getCell(5).value = brandUpper; // Mfg
      row.getCell(6).value = modelClean; // Model
      row.getCell(7).value = `${brandUpper} ${modelClean}`; // Model (Blancco)
      row.getCell(8).value = d.serial || ''; // Serial Number
      row.getCell(9).value = cpu; // CPU
      row.getCell(10).value = 1; // Number of Processors
      row.getCell(11).value = ''; // CPU Speed
      row.getCell(12).value = ramGb; // Memory (GB)
      row.getCell(13).value = ssdGb; // HDD (GB)
      row.getCell(14).value = ssdGb > 0 ? 1 : 0; // HDD Count
      row.getCell(15).value = ''; // Monitor Size
      row.getCell(16).value = 'Functional'; // Test Status
      row.getCell(17).value = ''; // Cosmetic Notes
      row.getCell(18).value = ''; // Missing Notes
      row.getCell(19).value = grade; // Grade
      row.getCell(20).value = ''; // Battery Present
      row.getCell(21).value = ''; // AC Adapter
      row.getCell(22).value = ''; // Curr. mAh/OEM
      row.getCell(23).value = ''; // Keyboard Language
      row.getCell(24).value = ''; // Graphic Card

      // Style
      for (let c = 1; c <= 24; c++) {
        row.getCell(c).font = ARIAL_10;
        row.getCell(c).alignment = { horizontal: 'left' };
      }
      dataRow++;
    }
  }

  // ═══════════════════════════════════════════════════════════════════════════
  // TAB 3: PlanBit conditions
  // ═══════════════════════════════════════════════════════════════════════════
  const cond = workbook.addWorksheet('PlanBit conditions');
  cond.getColumn(1).width = 160;

  // Title
  cond.getCell('A2').value = 'Maximising and understanding value back';
  cond.getCell('A2').font = ARIAL_10_BOLD;
  cond.getCell('A2').fill = ORANGE_FILL;
  cond.getCell('A2').alignment = { horizontal: 'left' };

  const conditions = [
    '', // row 3 empty
    'Value is driven by three key areas, condition of equipment, components within equipment and timing of receipt of equipment.',
    'Age of Equipment: Typically computer product older than 4 years has little or no value back potential.',
    'Pricing is subject to change at the beginning of each calendar month, in accordance with changes in the market value of used equipment of the same manufacture and of similar configuration, capacity, and condition.',
    'Prices listed are only valid for equipment received and processed by PlanBit in the month specified.',
    'All prices listed are for equipment received complete, functional, and in good cosmetic condition.',
    'Complete notebooks and desktops include, but are not limited to, the system processor, system memory, system chassis with power supply, AC adapter, battery, one fixed disk drive, and an optical drive.',
    'Processor Type: Prices shown are based on Intel Pentium Processor. Non Intel Pentium Processors will achieve reduced values (30 to 50% less).',
    'Passwords and locking mechanisms will cause processing delays and may impact value. All Apple devices (serial number) are removed from the Apple Device Enrolment Program and or are not in Managed Mode Software enrolled.',
    'All Notebook devices are removed from MMS, Splash screens if present have been removed from BIOS - All Apple devices have been removed from I-Cloud and MMS',
    'With Apple products with iOS 7 need "Find my iPhone" to be de-activated. In case the Activation Lock is still turned on, the product will have no recovery value.',
    'Equipment needing repair, incomplete, or non-functional will be subject to reduced or no recovery value.',
    'Equipment that is no longer marketable, is missing major components and/or is cosmetically damaged so that it cannot be resold will be ethically recycled in accordance with local and EU guidelines.',
    'All value recovery quotations are estimated. Final values are applied following physical assessment at PlanBit\'s location.',
    'Customers receive an inventory and settlement report within 35 business days after collection date, with detailed instructions regarding invoicing to recover residual value.',
    '', '', // empty rows
    'Conditions and assumptions',
    '',
    'Service is based on assessment of product for re-sale.',
    'Prices are based on the numbers and models provided by the customer, changes could influence the price.',
    'Equipment to be ready for transport, in a centralized (preferable ground floor location) at customer site, with pallet access and close to an exit (entrance 1,2 meter wide)',
    'Pricing does not include waiting time. Product must be consolidated and prepared for collection on the agreed collection date and time otherwise subject to extra cost.',
    'Products need to be packed to conform with PlanBit standard packing instruction.',
    'PlanBit will not be liable for any damage to the equipment prior to inventory/collection.',
    'Site access issues for vehicles must be declared in advance. Access limitations for appropriate transportation may attract additional costs.',
    'Capability indicated is dependant upon site readiness and accessibility at time of services implementation',
    'All work is carried out during local normal office hours (09.00-17.00) on local normal business days.',
    'Prices are based on the availability to the equipment in the Netherlands',
    'Collection and recycling of packaging is not included in this quotation.',
    'Pricing (unless stated otherwise) is in \u20ac (Euro) excluding 21% VAT.',
    'Prices are including collection and transport. PlanBit will contact the customer to set an appointment',
    'All equipment will be audited and tested.',
    'All data on hard drives will be overwritten with certified software (Blancco) we guarantees the highest security level - level 2/NIST SP800-88 Purge "Enhanced Secure Erase" or otherwise physically destroyed where disk is inoperable',
    'Detailed product reports, data removal and recycling-certificates will be provided to the customer.',
    'Quotation is valid for 14 days',
  ];

  conditions.forEach((text, i) => {
    const row = i + 4; // start at row 4
    if (!text) return;
    const cell = cond.getCell(row, 1);
    cell.value = text;
    if (text === 'Conditions and assumptions') {
      cell.font = ARIAL_10_BOLD;
      cell.fill = ORANGE_FILL;
      cell.alignment = { horizontal: 'left' };
    } else {
      cell.font = { name: 'Arial', size: 10, color: { argb: 'FF000000' } };
      cell.alignment = { horizontal: 'left', wrapText: true };
    }
  });

  // ─── Write buffer ──────────────────────────────────────────────────────
  const buf = await workbook.xlsx.writeBuffer();
  return Buffer.from(buf);
}

module.exports = { generateExcel };
