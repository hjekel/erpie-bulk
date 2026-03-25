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

// Multer for file upload
const multer = require('multer');
const upload = multer({ storage: multer.memoryStorage(), limits: { fileSize: 50 * 1024 * 1024 } });

// POST /api/analyze — file upload
app.post('/api/analyze', upload.single('file'), (req, res) => {
  try {
    const dealName = req.body.dealName || 'Unnamed Deal';
    const region = req.body.region || 'EU';
    const liveValidation = req.body.liveValidation === 'true';

    if (!req.file) return res.status(400).json({ error: 'No file uploaded' });

    const t0 = Date.now();
    const parsed = parseFile(req.file.buffer, req.file.originalname);
    const devices = parsed.devices || parsed;
    const analysis = analyzeDevices(devices, region);
    const processingTimeMs = Date.now() - t0;

    res.json({ dealName, format: 'file', liveValidation, processingTimeMs, ...analysis });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

// POST /api/analyze-text — text paste
app.post('/api/analyze-text', (req, res) => {
  try {
    const { text, dealName = 'Unnamed Deal', region = 'EU', liveValidation = false } = req.body;

    if (!text) return res.status(400).json({ error: 'No text provided' });

    const t0 = Date.now();
    const parsed = parseText(text);
    const devices = parsed.devices || parsed;
    const analysis = analyzeDevices(devices, region);
    const processingTimeMs = Date.now() - t0;

    res.json({ dealName, format: 'text', liveValidation, processingTimeMs, ...analysis });
  } catch (e) {
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
