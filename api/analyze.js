'use strict';

const formidable = require('formidable');
const { parseFile } = require('../lib/parser');
const { analyzeDevices } = require('../lib/pricing-engine');

module.exports = async function handler(req, res) {
  // CORS
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
  if (req.method === 'OPTIONS') return res.status(200).end();
  if (req.method !== 'POST') return res.status(405).json({ ok: false, error: 'Method not allowed' });

  try {
    const form = formidable({
      maxFileSize: 20 * 1024 * 1024, // 20MB
      keepExtensions: true,
    });

    const [fields, files] = await form.parse(req);

    const file = files.file?.[0];
    if (!file) {
      return res.status(400).json({ ok: false, error: 'No file uploaded' });
    }

    const dealName = (fields.dealName?.[0]) || 'Unnamed Deal';
    const region = (fields.region?.[0]) || 'EU';
    const liveValidation = (fields.liveValidation?.[0]) === 'true';

    // Read file into buffer
    const fs = require('fs');
    const buffer = fs.readFileSync(file.filepath);
    const filename = file.originalFilename || 'upload.xlsx';

    // Parse
    const { devices, format } = parseFile(buffer, filename);

    if (!devices || !devices.length) {
      return res.status(400).json({ ok: false, error: 'No valid devices found in file' });
    }

    // Price
    const { results, summary } = analyzeDevices(devices, region);

    // Cleanup temp file
    try { fs.unlinkSync(file.filepath); } catch (e) { /* ignore */ }

    res.status(200).json({
      ok: true,
      dealName,
      format,
      liveValidation,
      results,
      summary,
    });
  } catch (err) {
    console.error('[analyze] Error:', err);
    res.status(500).json({ ok: false, error: err.message });
  }
};

// Disable body parsing — formidable handles it
module.exports.config = {
  api: { bodyParser: false },
};
