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

// ═══ PTOLEMAEUS — Knowledge Library API ═══════════════════════════════════════
const PTOLEMAEUS_DIR = '/home/openclaw/.openclaw/workspace/PTOLEMAEUS';
const LEADERS_FILE = path.join(PTOLEMAEUS_DIR, 'THOUGHT_LEADERS.md');
const ACTIONS_FILE = path.join(PTOLEMAEUS_DIR, 'APPROVED_ACTIONS.md');

// Helper: parse card metadata from markdown
function parseCardMeta(filename, content) {
  const title = content.match(/^# (.+)$/m)?.[1] || filename;
  const date = content.match(/\*\*Datum:\*\* (.+)$/m)?.[1] || filename.slice(0, 10);
  const type = content.match(/\*\*Type:\*\* (.+)$/m)?.[1] || 'Tekst';
  const tldr = content.match(/## TL;DR\n([\s\S]*?)(?=\n##)/)?.[1]?.trim() || '';
  const proposals = (content.match(/\*\*Status:\*\* PENDING/g) || []).length;
  return { filename, title, date, type, tldr, proposals };
}

// GET /api/ptolemaeus/cards
app.get('/api/ptolemaeus/cards', (req, res) => {
  try {
    const fs = require('fs');
    if (!fs.existsSync(PTOLEMAEUS_DIR)) return res.json({ cards: [] });
    const files = fs.readdirSync(PTOLEMAEUS_DIR)
      .filter(f => f.endsWith('.md') && f !== 'APPROVED_ACTIONS.md' && f !== 'THOUGHT_LEADERS.md')
      .sort().reverse();
    const cards = files.map(f => {
      const content = fs.readFileSync(path.join(PTOLEMAEUS_DIR, f), 'utf8');
      return parseCardMeta(f, content);
    });
    res.json({ cards });
  } catch(e) { res.status(500).json({ error: e.message }); }
});

// GET /api/ptolemaeus/card/:filename
app.get('/api/ptolemaeus/card/:filename', (req, res) => {
  try {
    const fs = require('fs');
    const filepath = path.join(PTOLEMAEUS_DIR, req.params.filename);
    if (!fs.existsSync(filepath)) return res.status(404).json({ error: 'Not found' });
    const content = fs.readFileSync(filepath, 'utf8');
    res.json({ content, ...parseCardMeta(req.params.filename, content) });
  } catch(e) { res.status(500).json({ error: e.message }); }
});

// POST /api/ptolemaeus/ingest
app.post('/api/ptolemaeus/ingest', async (req, res) => {
  try {
    const { input } = req.body;
    if (!input) return res.status(400).json({ error: 'No input provided' });

    const { execFile } = require('child_process');
    const { promisify } = require('util');
    const execFileAsync = promisify(execFile);

    console.log(`[ptolemaeus] Ingesting: ${input.slice(0, 80)}...`);
    let stdout, stderr;
    try {
      const result = await execFileAsync('node', ['/home/openclaw/ptolemaeus/ingest.js', input], {
        timeout: 90_000,
        env: { ...process.env, PATH: `${process.env.PATH}:/home/openclaw/.npm-global/bin:/home/openclaw/.local/bin:/usr/local/bin` },
      });
      stdout = result.stdout;
      stderr = result.stderr;
    } catch(childErr) {
      // ingest.js exits with code 1 on YouTube/fetch errors — check stderr for message
      const msg = (childErr.stderr || childErr.message || '').trim();
      if (msg.includes('Shorts') || msg.includes('transcript') || msg.includes('opgehaald')) {
        // Structured error from ingest.js
        const lines = msg.split('\n').filter(l => l.length > 0);
        return res.status(400).json({
          error: lines[0]?.replace(/^\u274C\s*/, '') || 'Ingest mislukt',
          tip: lines[1] || '',
        });
      }
      throw childErr;
    }
    if (stderr) console.warn('[ptolemaeus] stderr:', stderr.slice(0, 200));

    // Re-read cards to get the new one
    const fs = require('fs');
    const files = fs.readdirSync(PTOLEMAEUS_DIR).filter(f => f.endsWith('.md') && f !== 'APPROVED_ACTIONS.md' && f !== 'THOUGHT_LEADERS.md').sort().reverse();
    const latest = files[0];
    if (latest) {
      const content = fs.readFileSync(path.join(PTOLEMAEUS_DIR, latest), 'utf8');
      return res.json({ success: true, ...parseCardMeta(latest, content) });
    }
    res.json({ success: true, message: 'Ingested' });
  } catch(e) {
    console.error('[ptolemaeus] Ingest error:', e.message);
    res.status(500).json({ error: e.message });
  }
});

// GET /api/ptolemaeus/proposals
app.get('/api/ptolemaeus/proposals', (req, res) => {
  try {
    const fs = require('fs');
    if (!fs.existsSync(PTOLEMAEUS_DIR)) return res.json({ pending: [], approved: [] });
    const files = fs.readdirSync(PTOLEMAEUS_DIR).filter(f => f.endsWith('.md') && f !== 'APPROVED_ACTIONS.md' && f !== 'THOUGHT_LEADERS.md');
    const pending = [];
    for (const file of files) {
      const content = fs.readFileSync(path.join(PTOLEMAEUS_DIR, file), 'utf8');
      const matches = content.matchAll(/### (Voorstel \d+: .+)\n\*\*Wat:\*\* (.+)\n\*\*Waarom:\*\* (.+)\n\*\*Effort:\*\* (.+)\n\*\*Status:\*\* PENDING/g);
      for (const m of matches) {
        pending.push({ file, title: m[1], wat: m[2], waarom: m[3], effort: m[4] });
      }
    }
    // Parse approved actions
    const approved = [];
    if (fs.existsSync(ACTIONS_FILE)) {
      const actions = fs.readFileSync(ACTIONS_FILE, 'utf8');
      const matches = actions.matchAll(/## .+ — (.+)\n\*\*Bron:\*\* .+\n\*\*Actie:\*\* (.+)\n[\s\S]*?\*\*Status:\*\* (.+)/g);
      for (const m of matches) {
        approved.push({ title: m[1], action: m[2], status: m[3] });
      }
    }
    res.json({ pending, approved });
  } catch(e) { res.status(500).json({ error: e.message }); }
});

// POST /api/ptolemaeus/vote
app.post('/api/ptolemaeus/vote', (req, res) => {
  try {
    const fs = require('fs');
    const { file, title, vote } = req.body;
    const filepath = path.join(PTOLEMAEUS_DIR, file);
    if (!fs.existsSync(filepath)) return res.status(404).json({ error: 'File not found' });

    let content = fs.readFileSync(filepath, 'utf8');
    const escaped = title.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    const statusMap = { approve: 'APPROVED', reject: 'REJECTED', defer: 'DEFERRED' };
    const newStatus = statusMap[vote] || 'DEFERRED';

    content = content.replace(
      new RegExp(`(### ${escaped}[\\s\\S]*?\\*\\*Status:\\*\\*) PENDING`),
      `$1 ${newStatus}`
    );
    fs.writeFileSync(filepath, content, 'utf8');

    // If approved, append to actions file
    if (vote === 'approve') {
      const watMatch = content.match(new RegExp(`### ${escaped}[\\s\\S]*?\\*\\*Wat:\\*\\* (.+)`));
      const waaromMatch = content.match(new RegExp(`### ${escaped}[\\s\\S]*?\\*\\*Waarom:\\*\\* (.+)`));
      const effortMatch = content.match(new RegExp(`### ${escaped}[\\s\\S]*?\\*\\*Effort:\\*\\* (.+)`));
      const today = new Date().toISOString().slice(0, 10);
      const entry = `\n## ${today} — ${title}\n**Bron:** ${file}\n**Actie:** ${watMatch?.[1] || ''}\n**Motivatie:** ${waaromMatch?.[1] || ''}\n**Effort:** ${effortMatch?.[1] || ''}\n**Assignee:** Claude Code\n**Status:** TODO\n`;
      if (!fs.existsSync(ACTIONS_FILE)) fs.writeFileSync(ACTIONS_FILE, '# Goedgekeurde Acties\n');
      fs.appendFileSync(ACTIONS_FILE, entry);
    }

    res.json({ success: true, newStatus });
  } catch(e) { res.status(500).json({ error: e.message }); }
});

// GET /api/ptolemaeus/leaders
app.get('/api/ptolemaeus/leaders', (req, res) => {
  try {
    const fs = require('fs');
    if (!fs.existsSync(LEADERS_FILE)) return res.json({ leaders: [] });
    const content = fs.readFileSync(LEADERS_FILE, 'utf8');
    const leaders = [];
    const matches = content.matchAll(/^- (.+?) \(search: "(.+?)"\)$/gm);
    for (const m of matches) {
      leaders.push({ name: m[1], search: m[2], lastScan: null });
    }
    // Check radar state for last scan dates
    const stateFile = path.join(PTOLEMAEUS_DIR, 'RADAR_STATE.json');
    if (fs.existsSync(stateFile)) {
      try {
        const state = JSON.parse(fs.readFileSync(stateFile, 'utf8'));
        for (const l of leaders) {
          if (state[l.name]) l.lastScan = state[l.name].lastScan;
        }
      } catch(e) {}
    }
    res.json({ leaders });
  } catch(e) { res.status(500).json({ error: e.message }); }
});

// POST /api/ptolemaeus/leaders — add a thought leader
app.post('/api/ptolemaeus/leaders', (req, res) => {
  try {
    const fs = require('fs');
    const { name, search } = req.body;
    if (!name || !search) return res.status(400).json({ error: 'Name and search required' });
    if (!fs.existsSync(PTOLEMAEUS_DIR)) fs.mkdirSync(PTOLEMAEUS_DIR, { recursive: true });
    const line = `- ${name} (search: "${search}")\n`;
    if (!fs.existsSync(LEADERS_FILE)) {
      fs.writeFileSync(LEADERS_FILE, '## Thought Leaders\n' + line);
    } else {
      fs.appendFileSync(LEADERS_FILE, line);
    }
    res.json({ success: true });
  } catch(e) { res.status(500).json({ error: e.message }); }
});

// GET /ptolemaeus — serve the library page
app.get('/ptolemaeus', (req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'ptolemaeus.html'));
});

app.listen(PORT, () => {
  console.log(`🦞 ERPIE Bulk running at http://localhost:${PORT}`);
  console.log(`📚 Ptolemaeus Library at http://localhost:${PORT}/ptolemaeus`);
});
