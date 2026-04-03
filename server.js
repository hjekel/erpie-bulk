'use strict';

const express = require('express');
const path = require('path');
const app = express();
const PORT = 3004;

app.use(express.json({ limit: '50mb' }));
app.use(express.static(path.join(__dirname, 'public')));

// Import handlers
const { parseFile, parseText } = require('./lib/parser');
const { calculatePrice, analyzeDevices } = require('./lib/pricing-engine');
const { generateExcel } = require('./lib/excel-export');
const { generateHTMLReport } = require('./lib/html-report');

// PDF and DOCX support
let pdfParse, mammoth;
try { pdfParse = require('pdf-parse'); } catch (e) { console.warn('[startup] pdf-parse not installed — PDF upload disabled'); }
try { mammoth = require('mammoth'); } catch (e) { console.warn('[startup] mammoth not installed — DOCX upload disabled'); }
const { priceViaOpenClaw, mapOpenClawResults, groupDevicesForPricing, normalizeViaOpenClaw } = require('./lib/openclaw-pricing');

// Multer for file upload
const multer = require('multer');
const upload = multer({ storage: multer.memoryStorage(), limits: { fileSize: 50 * 1024 * 1024 } });

// ─── Shared pricing logic: OpenClaw first, local fallback ────────────────────
async function priceDevices(devices, region, dealName) {
  const t0 = Date.now();

  // Try OpenClaw first
  const groups = groupDevicesForPricing(devices);
  console.log(`[pricing] ${groups.length} groups, trying OpenClaw...`);

  const openclawResult = await priceViaOpenClaw(groups, region, dealName);
  if (openclawResult && openclawResult.results.length > 0) {
    const analysis = mapOpenClawResults(devices, openclawResult);
    const processingTimeMs = Date.now() - t0;
    console.log(`[pricing] OpenClaw success: ${analysis.results.length} groups, €${analysis.summary.totalValue} in ${processingTimeMs}ms`);
    return { ...analysis, processingTimeMs, pricingSource: 'openclaw' };
  }

  // Fallback to local pricing engine
  console.log('[pricing] OpenClaw unavailable, falling back to local engine');
  const analysis = analyzeDevices(devices, region);
  const processingTimeMs = Date.now() - t0;
  return { ...analysis, processingTimeMs, pricingSource: 'local' };
}

// POST /api/analyze — file upload
app.post('/api/analyze', upload.single('file'), async (req, res) => {
  try {
    const dealName = req.body.dealName || 'Unnamed Deal';
    const region = req.body.region || 'EU';
    const liveValidation = req.body.liveValidation === 'true';

    if (!req.file) return res.status(400).json({ error: 'No file uploaded' });

    const ext = (req.file.originalname || '').toLowerCase().split('.').pop();
    const mime = (req.file.mimetype || '').toLowerCase();
    let parsed;

    // PDF → extract text → text parser
    if (ext === 'pdf' || mime === 'application/pdf') {
      if (!pdfParse) return res.status(400).json({ error: 'PDF support not installed (npm install pdf-parse)' });
      const pdfData = await pdfParse(req.file.buffer);
      console.log(`[upload] PDF: ${pdfData.numpages} pages, ${pdfData.text.length} chars extracted`);
      parsed = parseText(pdfData.text);
      parsed.format = 'PDF';
    }
    // DOCX → extract text → text parser
    else if (ext === 'docx' || mime === 'application/vnd.openxmlformats-officedocument.wordprocessingml.document') {
      if (!mammoth) return res.status(400).json({ error: 'DOCX support not installed (npm install mammoth)' });
      const result = await mammoth.extractRawText({ buffer: req.file.buffer });
      console.log(`[upload] DOCX: ${result.value.length} chars extracted`);
      parsed = parseText(result.value);
      parsed.format = 'DOCX';
    }
    // Excel / CSV / TXT → existing parser
    else {
      parsed = parseFile(req.file.buffer, req.file.originalname);
    }

    const devices = parsed.devices || parsed;

    if (!devices || !devices.length) {
      return res.status(400).json({ error: 'No valid devices found in file. Supported: .xlsx, .xls, .csv, .pdf, .docx, .txt' });
    }

    const analysis = await priceDevices(devices, region, dealName);

    res.json({ dealName, format: parsed.format || 'file', liveValidation, ...analysis });
  } catch (e) {
    console.error('[analyze] Error:', e);
    res.status(500).json({ error: e.message });
  }
});

// POST /api/analyze-text — text paste
app.post('/api/analyze-text', async (req, res) => {
  try {
    const { text, dealName = 'Unnamed Deal', region = 'EU', liveValidation = false } = req.body;

    if (!text) return res.status(400).json({ error: 'No text provided' });

    const parsed = parseText(text);
    const devices = parsed.devices || parsed;

    if (!devices || !devices.length) {
      return res.status(400).json({ error: 'Geen apparaten herkend in de tekst.' });
    }

    const analysis = await priceDevices(devices, region, dealName);

    res.json({ dealName, format: 'text', liveValidation, ...analysis });
  } catch (e) {
    console.error('[analyze-text] Error:', e);
    res.status(500).json({ error: e.message });
  }
});

// POST /api/normalize — AI normalisation (Stage 1 of two-stage rocket)
app.post('/api/normalize', async (req, res) => {
  try {
    const { text } = req.body;
    if (!text) return res.status(400).json({ error: 'No text provided' });

    // Try OpenClaw AI normalisation first
    const aiResult = await normalizeViaOpenClaw(text);
    if (aiResult) {
      return res.json(aiResult);
    }

    // Fallback: use local parser and format output as standard lines
    console.log('[normalize] OpenClaw unavailable, falling back to local parser');
    const parsed = parseText(text);
    const devices = parsed.devices || [];

    if (!devices.length) {
      return res.status(400).json({ error: 'Geen apparaten herkend in de tekst.' });
    }

    const lines = devices.map(d => {
      const parts = [d.model];
      if (d.cpu) parts.push(d.cpu);
      if (d.ram) parts.push(typeof d.ram === 'number' ? d.ram + 'GB' : d.ram);
      if (d.ssd) parts.push(typeof d.ssd === 'number' ? d.ssd + 'GB' : d.ssd);
      parts.push('Grade ' + (d.grade || 'B'));
      return `${d.qty || 1}x ${parts.join(' \u00b7 ')}`;
    });

    const assetCount = devices.reduce((s, d) => s + (d.qty || 1), 0);
    res.json({
      normalised: lines.join('\n'),
      assetCount,
      modelCount: devices.length,
      source: 'local',
    });
  } catch (e) {
    console.error('[normalize] Error:', e);
    res.status(500).json({ error: e.message });
  }
});

// POST /api/export-excel — generate ARS Excel
app.post('/api/export-excel', async (req, res) => {
  try {
    const { results, summary, dealName = 'Deal' } = req.body;
    const buffer = await generateExcel(results, summary, dealName);

    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', `attachment; filename="ERPIE_ARS_${dealName.replace(/\s+/g,'_')}.xlsx"`);
    res.send(buffer);
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

// POST /api/validate-prices — live price validation
const validatePricesHandler = require('./api/validate-prices');
app.post('/api/validate-prices', validatePricesHandler);

// POST /api/export-html — generate HTML one-pager
app.post('/api/export-html', (req, res) => {
  try {
    const { results, summary, dealName = 'Deal' } = req.body;
    const html = generateHTMLReport(results, summary, dealName);

    res.setHeader('Content-Type', 'text/html');
    res.send(html);
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

app.listen(PORT, () => {
  console.log(`🦞 ERPIE Bulk running at http://localhost:${PORT}`);
});
