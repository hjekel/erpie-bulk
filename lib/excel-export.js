'use strict';

const ExcelJS = require('exceljs');

/**
 * Generate PlanBit ARS Excel export
 * @param {Array} results - Analyzed device results from pricing engine
 * @param {Object} summary - Summary object from analyzeDevices
 * @param {string} dealName - Name of the deal
 * @returns {Promise<Buffer>} Excel file buffer
 */
async function generateExcel(results, summary, dealName) {
  const workbook = new ExcelJS.Workbook();
  workbook.creator = 'ERPIE Bulk Pricing';
  workbook.created = new Date();

  const sheet = workbook.addWorksheet('ARS Pricing', {
    properties: { defaultRowHeight: 20 },
  });

  // ─── Header section ──────────────────────────────────────────────────────
  const date = new Date().toLocaleDateString('nl-NL', {
    day: '2-digit', month: 'long', year: 'numeric',
  });

  // Title row
  sheet.mergeCells('A1:K1');
  const titleCell = sheet.getCell('A1');
  titleCell.value = `PlanBit ERPIE Bulk Pricing — ${dealName}`;
  titleCell.font = { name: 'Segoe UI', size: 16, bold: true, color: { argb: 'FF1A1A3E' } };
  titleCell.alignment = { vertical: 'middle' };
  sheet.getRow(1).height = 36;

  // Date row
  sheet.mergeCells('A2:K2');
  const dateCell = sheet.getCell('A2');
  dateCell.value = `Generated: ${date}`;
  dateCell.font = { name: 'Segoe UI', size: 10, italic: true, color: { argb: 'FF718096' } };
  sheet.getRow(2).height = 20;

  // Empty row
  sheet.getRow(3).height = 8;

  // ─── Column headers ──────────────────────────────────────────────────────
  const headers = ['#', 'Brand', 'Model', 'CPU Gen', 'RAM', 'SSD', 'Grade', 'Status', 'ERP per unit', 'Quantity', 'Total ERP'];

  const headerRow = sheet.getRow(4);
  headers.forEach((h, i) => {
    const cell = headerRow.getCell(i + 1);
    cell.value = h;
    cell.font = { name: 'Segoe UI', size: 10, bold: true, color: { argb: 'FFFFFFFF' } };
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF2D3748' } };
    cell.alignment = { vertical: 'middle', horizontal: 'center' };
    cell.border = {
      bottom: { style: 'thin', color: { argb: 'FF4A5568' } },
    };
  });
  headerRow.height = 28;

  // Column widths
  sheet.getColumn(1).width = 6;   // #
  sheet.getColumn(2).width = 14;  // Brand
  sheet.getColumn(3).width = 32;  // Model
  sheet.getColumn(4).width = 10;  // CPU Gen
  sheet.getColumn(5).width = 8;   // RAM
  sheet.getColumn(6).width = 8;   // SSD
  sheet.getColumn(7).width = 8;   // Grade
  sheet.getColumn(8).width = 10;  // Status
  sheet.getColumn(9).width = 14;  // ERP per unit
  sheet.getColumn(10).width = 10; // Quantity
  sheet.getColumn(11).width = 14; // Total ERP

  // ─── Data rows ────────────────────────────────────────────────────────────
  const statusColors = {
    'GO':     { font: 'FF00C853', bg: 'FFE8F5E9' },
    'WATCH':  { font: 'FFFF8F00', bg: 'FFFFF3E0' },
    'NO-GO':  { font: 'FFD50000', bg: 'FFFFEBEE' },
    'ERROR':  { font: 'FF9E9E9E', bg: 'FFFAFAFA' },
  };

  results.forEach((d, idx) => {
    const rowNum = idx + 5;
    const row = sheet.getRow(rowNum);
    const qty = d.qty || 1;
    const totalErp = (d.advisedPrice || 0) * qty;

    // Extract brand from model
    const modelLower = (d.model || '').toLowerCase();
    let brand = '';
    if (modelLower.includes('dell')) brand = 'Dell';
    else if (modelLower.includes('hp') || modelLower.includes('hewlett')) brand = 'HP';
    else if (modelLower.includes('lenovo') || modelLower.includes('thinkpad')) brand = 'Lenovo';
    else if (modelLower.includes('apple') || modelLower.includes('macbook')) brand = 'Apple';
    else if (modelLower.includes('microsoft') || modelLower.includes('surface')) brand = 'Microsoft';
    else if (modelLower.includes('asus')) brand = 'Asus';
    else if (modelLower.includes('acer')) brand = 'Acer';
    else if (modelLower.includes('fujitsu')) brand = 'Fujitsu';
    else if (modelLower.includes('toshiba')) brand = 'Toshiba';
    else brand = d.brand || '';

    const values = [
      idx + 1,
      brand,
      d.model || '',
      d.gen || '',
      d.ramGb ? `${d.ramGb}GB` : '',
      d.ssdGb ? `${d.ssdGb}GB` : '',
      d.grade || '',
      d.status || '',
      d.advisedPrice || 0,
      qty,
      totalErp,
    ];

    values.forEach((v, i) => {
      const cell = row.getCell(i + 1);
      cell.value = v;
      cell.font = { name: 'Segoe UI', size: 10 };
      cell.alignment = { vertical: 'middle', horizontal: i >= 8 ? 'right' : (i === 0 ? 'center' : 'left') };

      // Alternate row shading
      if (idx % 2 === 1) {
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF7FAFC' } };
      }

      cell.border = {
        bottom: { style: 'hair', color: { argb: 'FFE2E8F0' } },
      };
    });

    // Status cell color coding
    const statusCell = row.getCell(8);
    const colors = statusColors[d.status] || statusColors['ERROR'];
    statusCell.font = { name: 'Segoe UI', size: 10, bold: true, color: { argb: colors.font } };
    statusCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: colors.bg } };
    statusCell.alignment = { vertical: 'middle', horizontal: 'center' };

    // Currency formatting for ERP columns
    row.getCell(9).numFmt = '€#,##0';
    row.getCell(11).numFmt = '€#,##0';

    row.height = 22;
  });

  // ─── Summary row ──────────────────────────────────────────────────────────
  const summaryRowNum = results.length + 5;
  const sumRow = sheet.getRow(summaryRowNum);

  sheet.mergeCells(`A${summaryRowNum}:H${summaryRowNum}`);
  const totalLabelCell = sumRow.getCell(1);
  totalLabelCell.value = `TOTAAL (${summary.total || results.length} devices)`;
  totalLabelCell.font = { name: 'Segoe UI', size: 11, bold: true, color: { argb: 'FF1A1A3E' } };
  totalLabelCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFEDF2F7' } };
  totalLabelCell.alignment = { vertical: 'middle' };

  // ERP per unit (average)
  const avgCell = sumRow.getCell(9);
  avgCell.value = summary.avgValue || 0;
  avgCell.font = { name: 'Segoe UI', size: 11, bold: true };
  avgCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFEDF2F7' } };
  avgCell.numFmt = '€#,##0';
  avgCell.alignment = { vertical: 'middle', horizontal: 'right' };

  // Total qty
  const qtyCell = sumRow.getCell(10);
  qtyCell.value = summary.total || results.reduce((s, r) => s + (r.qty || 1), 0);
  qtyCell.font = { name: 'Segoe UI', size: 11, bold: true };
  qtyCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFEDF2F7' } };
  qtyCell.alignment = { vertical: 'middle', horizontal: 'center' };

  // Total ERP
  const totalCell = sumRow.getCell(11);
  totalCell.value = summary.totalValue || 0;
  totalCell.font = { name: 'Segoe UI', size: 11, bold: true };
  totalCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFEDF2F7' } };
  totalCell.numFmt = '€#,##0';
  totalCell.alignment = { vertical: 'middle', horizontal: 'right' };

  sumRow.height = 28;

  // Bid range row
  const bidRowNum = summaryRowNum + 1;
  const bidRow = sheet.getRow(bidRowNum);
  sheet.mergeCells(`A${bidRowNum}:H${bidRowNum}`);
  const bidLabel = bidRow.getCell(1);
  bidLabel.value = 'Bid Range (suggested offer)';
  bidLabel.font = { name: 'Segoe UI', size: 10, italic: true, color: { argb: 'FF718096' } };

  sheet.mergeCells(`I${bidRowNum}:K${bidRowNum}`);
  const bidVal = bidRow.getCell(9);
  bidVal.value = `€${(summary.bidLow || 0).toLocaleString('nl-NL')} – €${(summary.bidHigh || 0).toLocaleString('nl-NL')}`;
  bidVal.font = { name: 'Segoe UI', size: 10, italic: true, color: { argb: 'FF718096' } };
  bidVal.alignment = { horizontal: 'right' };

  // Status breakdown row
  const breakdownRowNum = bidRowNum + 1;
  const breakdownRow = sheet.getRow(breakdownRowNum);
  sheet.mergeCells(`A${breakdownRowNum}:K${breakdownRowNum}`);
  const breakdownCell = breakdownRow.getCell(1);
  breakdownCell.value = `GO: ${summary.goCount || 0}  |  WATCH: ${summary.watchCount || 0}  |  NO-GO: ${summary.nogoCount || 0}`;
  breakdownCell.font = { name: 'Segoe UI', size: 10, color: { argb: 'FF718096' } };

  // Footer
  const footerRowNum = breakdownRowNum + 2;
  const footerRow = sheet.getRow(footerRowNum);
  sheet.mergeCells(`A${footerRowNum}:K${footerRowNum}`);
  const footerCell = footerRow.getCell(1);
  footerCell.value = 'ERPIE PriceFinder · PlanBit ITAD · Prijzen zijn indicatief op basis van marktdata.';
  footerCell.font = { name: 'Segoe UI', size: 9, italic: true, color: { argb: 'FFA0AEC0' } };

  const buf = await workbook.xlsx.writeBuffer();
  return Buffer.from(buf);
}

module.exports = { generateExcel };
