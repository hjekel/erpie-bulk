'use strict';

function escHtml(str) {
  return String(str).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
}

function statusColor(status) {
  if (status === 'GO')    return '#00c853';
  if (status === 'WATCH') return '#ffab00';
  return '#d50000';
}

function fmtCurrency(n) {
  return new Intl.NumberFormat('nl-NL', { style: 'decimal', minimumFractionDigits: 0 }).format(n || 0);
}

/**
 * Generate a self-contained HTML one-pager report
 * @param {Array} results - Analyzed device results
 * @param {Object} summary - Summary from analyzeDevices
 * @param {string} dealName - Name of the deal
 * @returns {string} Complete HTML string
 */
function generateHTMLReport(results, summary, dealName) {
  const date = new Date().toLocaleDateString('nl-NL', {
    day: '2-digit', month: 'long', year: 'numeric',
  });

  const rows = results.map((d, i) => {
    const qty = d.qty || 1;
    const totalErp = (d.advisedPrice || 0) * qty;
    return `
    <tr>
      <td class="num">${i + 1}</td>
      <td>${escHtml(d.model || '')}</td>
      <td class="center">${escHtml(d.gen || '')}</td>
      <td class="center">${d.ramGb || '—'}GB</td>
      <td class="center">${d.ssdGb || '—'}GB</td>
      <td class="center">${escHtml(d.grade || '—')}</td>
      <td class="center"><span class="pill" style="background:${statusColor(d.status)}20;color:${statusColor(d.status)};border:1px solid ${statusColor(d.status)}">${d.status}</span></td>
      <td class="money">&euro;${fmtCurrency(d.advisedPrice)}</td>
      <td class="center">${qty}</td>
      <td class="money">&euro;${fmtCurrency(totalErp)}</td>
    </tr>`;
  }).join('');

  return `<!DOCTYPE html>
<html lang="nl">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>ERPIE Report &ndash; ${escHtml(dealName)}</title>
<style>
  * { box-sizing: border-box; margin: 0; padding: 0; }
  body { font-family: 'Segoe UI', system-ui, -apple-system, sans-serif; background: #0a0a1f; color: #e2e8f0; min-height: 100vh; }

  .header {
    background: linear-gradient(135deg, #0a0a1f 0%, #1a1a3e 50%, #0d1b2a 100%);
    color: #fff; padding: 40px 48px; border-bottom: 2px solid #f59e0b33;
  }
  .header h1 { font-size: 28px; font-weight: 700; color: #f59e0b; }
  .header .sub { font-size: 13px; color: #94a3b8; margin-top: 6px; }

  .content { max-width: 1200px; margin: 32px auto; padding: 0 24px; }

  .cards { display: grid; grid-template-columns: repeat(auto-fit, minmax(220px, 1fr)); gap: 16px; margin-bottom: 28px; }
  .card {
    background: linear-gradient(135deg, #1e293b 0%, #1a1a3e 100%);
    border-radius: 12px; padding: 24px;
    border: 1px solid #334155;
    box-shadow: 0 4px 12px rgba(0,0,0,.3);
  }
  .card .label { font-size: 11px; color: #94a3b8; text-transform: uppercase; letter-spacing: .08em; font-weight: 600; }
  .card .value { font-size: 32px; font-weight: 700; color: #f1f5f9; margin-top: 8px; }
  .card .value.accent { color: #f59e0b; }
  .card .sub { font-size: 11px; color: #64748b; margin-top: 4px; }

  .status-pills { display: flex; gap: 12px; margin-bottom: 20px; flex-wrap: wrap; }
  .pill { padding: 4px 14px; border-radius: 999px; font-size: 13px; font-weight: 600; display: inline-block; }

  .rec {
    background: #1e293b; border-left: 4px solid #f59e0b;
    padding: 16px 20px; border-radius: 0 8px 8px 0;
    margin-bottom: 28px; font-size: 14px; color: #cbd5e1;
  }
  .rec strong { color: #f59e0b; }

  .table-wrap {
    background: #1e293b; border-radius: 12px; overflow: hidden;
    box-shadow: 0 4px 12px rgba(0,0,0,.3); border: 1px solid #334155;
  }

  table { width: 100%; border-collapse: collapse; }
  thead th {
    background: #0f172a; color: #94a3b8; padding: 14px 16px;
    text-align: left; font-size: 11px; text-transform: uppercase;
    letter-spacing: .06em; font-weight: 600;
    border-bottom: 2px solid #334155;
  }
  tbody tr { transition: background .15s; }
  tbody tr:hover { background: #f59e0b08; }
  tbody tr:nth-child(even) { background: #0f172a44; }
  tbody td {
    padding: 10px 16px; font-size: 13px;
    border-bottom: 1px solid #1e293b;
    color: #cbd5e1;
  }
  td.num { color: #64748b; text-align: center; width: 40px; }
  td.center { text-align: center; }
  td.money { text-align: right; font-variant-numeric: tabular-nums; font-weight: 500; }

  tfoot td {
    padding: 14px 16px; font-weight: 700; background: #0f172a;
    color: #f1f5f9; border-top: 2px solid #334155;
  }

  .footer {
    text-align: center; padding: 40px 24px;
    font-size: 11px; color: #475569;
    border-top: 1px solid #1e293b;
    margin-top: 40px;
  }

  @media print {
    body { background: #fff; color: #1a202c; }
    .header { background: #1a1a3e !important; }
    .card { background: #f8fafc !important; border-color: #e2e8f0 !important; }
    .card .value { color: #1a202c !important; }
    .table-wrap { border-color: #e2e8f0 !important; }
    thead th { background: #2d3748 !important; }
    tbody td { color: #1a202c !important; border-color: #e2e8f0 !important; }
    tfoot td { background: #edf2f7 !important; color: #1a202c !important; }
  }

  @media (max-width: 768px) {
    .header { padding: 24px; }
    .header h1 { font-size: 20px; }
    .cards { grid-template-columns: repeat(2, 1fr); }
    .card .value { font-size: 24px; }
    .content { padding: 0 12px; }
    table { font-size: 11px; }
    thead th, tbody td, tfoot td { padding: 8px 10px; }
  }
</style>
</head>
<body>
<div class="header">
  <h1>PlanBit &ndash; ERPIE Price Report</h1>
  <div class="sub">${escHtml(dealName)} &nbsp;|&nbsp; ${date} &nbsp;|&nbsp; Powered by ERPIE PriceFinder v1.0</div>
</div>

<div class="content">
  <div class="cards">
    <div class="card">
      <div class="label">Total Assets</div>
      <div class="value">${summary.total || 0}</div>
      <div class="sub">devices analysed</div>
    </div>
    <div class="card">
      <div class="label">Advised Value</div>
      <div class="value accent">&euro;${fmtCurrency(summary.totalValue)}</div>
      <div class="sub">sum of ERP prices</div>
    </div>
    <div class="card">
      <div class="label">Average ERP</div>
      <div class="value">&euro;${fmtCurrency(summary.avgValue)}</div>
      <div class="sub">per device</div>
    </div>
    <div class="card">
      <div class="label">Bid Range</div>
      <div class="value" style="font-size:22px">&euro;${fmtCurrency(summary.bidLow)} &ndash; &euro;${fmtCurrency(summary.bidHigh)}</div>
      <div class="sub">suggested offer</div>
    </div>
  </div>

  <div class="status-pills">
    <span class="pill" style="background:#00c85320;color:#00c853;border:1px solid #00c853">GO: ${summary.goCount || 0}</span>
    <span class="pill" style="background:#ffab0020;color:#ffab00;border:1px solid #ffab00">WATCH: ${summary.watchCount || 0}</span>
    <span class="pill" style="background:#d5000020;color:#d50000;border:1px solid #d50000">NO-GO: ${summary.nogoCount || 0}</span>
  </div>

  <div class="rec"><strong>Recommendation:</strong> ${escHtml(summary.recommendation || '')}</div>

  <div class="table-wrap">
    <table>
      <thead>
        <tr>
          <th>#</th><th>Model</th><th>Gen</th><th>RAM</th><th>SSD</th>
          <th>Grade</th><th>Status</th><th>ERP</th><th>Qty</th><th>Total</th>
        </tr>
      </thead>
      <tbody>${rows}</tbody>
      <tfoot>
        <tr>
          <td colspan="7"><strong>TOTAAL (${summary.total || 0} devices)</strong></td>
          <td class="money"><strong>&euro;${fmtCurrency(summary.avgValue)}</strong></td>
          <td class="center"><strong>${summary.total || 0}</strong></td>
          <td class="money"><strong>&euro;${fmtCurrency(summary.totalValue)}</strong></td>
        </tr>
      </tfoot>
    </table>
  </div>
</div>

<div class="footer">
  ERPIE PriceFinder &middot; PlanBit ITAD &middot; Prijzen zijn indicatief op basis van marktdata.<br>
  Werkelijke opbrengst kan afwijken o.b.v. conditie, vraag en logistieke kosten.
</div>
</body>
</html>`;
}

module.exports = { generateHTMLReport };
