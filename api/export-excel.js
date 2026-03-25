'use strict';

const { generateExcel } = require('../lib/excel-export');

module.exports = async function handler(req, res) {
  // CORS
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
  if (req.method === 'OPTIONS') return res.status(200).end();
  if (req.method !== 'POST') return res.status(405).json({ ok: false, error: 'Method not allowed' });

  try {
    const { results = [], summary = {}, dealName = 'ERPIE Deal' } = req.body || {};

    if (!results.length) {
      return res.status(400).json({ ok: false, error: 'No results to export' });
    }

    const buffer = await generateExcel(results, summary, dealName);

    const safeFilename = dealName.replace(/[^a-zA-Z0-9_\-]/g, '_').slice(0, 50);
    const filename = `ERPIE_ARS_${safeFilename}_${new Date().toISOString().slice(0, 10)}.xlsx`;

    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', `attachment; filename="${filename}"`);
    res.status(200).send(buffer);
  } catch (err) {
    console.error('[export-excel] Error:', err);
    res.status(500).json({ ok: false, error: err.message });
  }
};
