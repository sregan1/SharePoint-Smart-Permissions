'use strict';
// Generates docs/screenshots/10_report_history.png and 11_explorer_external.png
// for the USER-GUIDE.md "Report History and Compare" and "Filtering for External
// Users" sections. Same HTML-mockup + puppeteer-core approach as
// docs/generate-screenshots.js — these are illustrative mockups, not captures
// of the live app.
// Run: node docs/generate-history-external-screenshots.js

const puppeteer = require('puppeteer-core');
const path = require('path');
const fs = require('fs');

const CHROME = fs.existsSync('C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe')
  ? 'C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe'
  : 'C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedge.exe';

const DOCS_OUT = path.join(__dirname, 'screenshots');
if (!fs.existsSync(DOCS_OUT)) fs.mkdirSync(DOCS_OUT, { recursive: true });

const font = `'Segoe UI', Arial, sans-serif`;

function appBanner() {
  return `
  <div style="background:#0078D4;display:flex;align-items:center;padding:0 14px;height:44px;gap:10px;flex-shrink:0;">
    <svg width="18" height="18" viewBox="0 0 24 24" fill="white">
      <path d="M12 1L3 5v6c0 5.55 3.84 10.74 9 12 5.16-1.26 9-6.45 9-12V5l-9-4z"/>
    </svg>
    <span style="color:#fff;font-size:14px;font-weight:600;font-family:${font};flex-shrink:0;">SharePoint Smart Permissions</span>
    <span style="color:rgba(255,255,255,0.7);font-size:12px;font-family:${font};margin-left:8px;">●</span>
    <span style="color:rgba(255,255,255,0.85);font-size:12px;font-family:${font};">https://contoso.sharepoint.com/sites/Marketing</span>
  </div>`;
}

function headerBar(title, rightHtml) {
  return `
  <div style="display:flex;align-items:center;gap:10px;padding:14px 18px 0;">
    <div style="display:inline-flex;align-items:center;gap:6px;border:none;
      font-size:13px;color:#323130;font-family:${font};cursor:pointer;">← Back</div>
    <span style="font-size:20px;font-weight:600;color:#323130;font-family:${font};flex:1;">${title}</span>
    ${rightHtml}
  </div>`;
}

function badge(label, bg, color) {
  return `<span style="display:inline-block;padding:2px 10px;border-radius:10px;font-size:11px;
    font-weight:600;background:${bg};color:${color};font-family:${font};white-space:nowrap;">${label}</span>`;
}

// ── Image: Permissions Report — History panel with Compare ──────────────────
function reportHistoryPage() {
  const reports = [
    { sel: true,  date: 'Jul 1, 2026, 9:02 AM',  site: '/sites/Marketing', scope: 'Full site', objects: 214, unique: 18 },
    { sel: true,  date: 'Jun 1, 2026, 9:00 AM',  site: '/sites/Marketing', scope: 'Full site', objects: 208, unique: 15 },
    { sel: false, date: 'May 1, 2026, 9:01 AM',  site: '/sites/Marketing', scope: 'Full site', objects: 201, unique: 14 },
    { sel: false, date: 'Apr 1, 2026, 9:03 AM',  site: '/sites/Marketing', scope: 'Full site', objects: 197, unique: 13 },
    { sel: false, date: 'Mar 3, 2026, 8:58 AM',  site: '/sites/Marketing', scope: 'Full site', objects: 190, unique: 12 },
  ];

  const rightHtml = `<div style="margin-left:auto;display:inline-flex;align-items:center;gap:6px;
      border:1px solid #0078D4;border-radius:4px;padding:6px 14px;background:#0078D4;color:#fff;
      font-size:13px;font-weight:600;font-family:${font};">Compare selected (2/2)</div>`;

  const rows = reports.map((r) => `
    <tr style="border-bottom:1px solid #F3F2F1;">
      <td style="padding:9px 6px;text-align:center;">
        <div style="width:16px;height:16px;border:1.5px solid ${r.sel ? '#0078D4' : '#8A8886'};border-radius:2px;
          background:${r.sel ? '#0078D4' : '#fff'};display:inline-flex;align-items:center;justify-content:center;
          color:#fff;font-size:11px;line-height:1;">${r.sel ? '✓' : ''}</div>
      </td>
      <td style="padding:9px 16px;font-size:13px;color:#323130;font-family:${font};white-space:nowrap;">${r.date}</td>
      <td style="padding:9px 16px;font-size:12px;color:#323130;font-family:${font};">${r.site}</td>
      <td style="padding:9px 16px;font-size:13px;color:#323130;font-family:${font};white-space:nowrap;">${r.scope}</td>
      <td style="padding:9px 16px;font-size:13px;color:#323130;font-family:${font};">${r.objects}</td>
      <td style="padding:9px 16px;font-size:13px;color:#323130;font-family:${font};">${r.unique}</td>
      <td style="padding:9px 16px;">
        <div style="display:flex;gap:6px;">
          <div style="border:none;border-radius:4px;padding:4px 10px;background:#0078D4;color:#fff;
            font-size:12px;font-weight:600;font-family:${font};">Export</div>
          <div style="border:1px solid #d1d1d1;border-radius:4px;padding:4px 8px;background:#fff;
            display:flex;align-items:center;">
            <svg width="14" height="14" viewBox="0 0 24 24" fill="#605E5C">
              <path d="M9 3v1H4v2h1v13a2 2 0 0 0 2 2h10a2 2 0 0 0 2-2V6h1V4h-5V3H9zm2 5h2v9h-2V8zm-4 0h2v9H7V8zm8 0h2v9h-2V8z"/>
            </svg>
          </div>
        </div>
      </td>
    </tr>`).join('');

  return `<!DOCTYPE html><html><head><meta charset="utf-8">
<style>* { margin:0; padding:0; box-sizing:border-box; }</style>
</head><body style="background:#fff;width:900px;font-family:${font};color:#323130;">
  ${appBanner()}
  ${headerBar('Permissions Report', rightHtml)}

  <div style="padding:14px 18px 0;font-size:13px;color:#605E5C;font-family:${font};">
    Tick two reports and select <strong>Compare selected</strong> to see what changed between scans.
  </div>

  <table style="width:100%;border-collapse:collapse;margin-top:12px;">
    <thead>
      <tr style="border-bottom:1px solid #EDEBE9;">
        <th style="padding:8px 6px;width:32px;"></th>
        <th style="padding:8px 16px;text-align:left;font-size:12px;font-weight:600;color:#323130;font-family:${font};">Date / Time ▼</th>
        <th style="padding:8px 16px;text-align:left;font-size:12px;font-weight:600;color:#323130;font-family:${font};">Site</th>
        <th style="padding:8px 16px;text-align:left;font-size:12px;font-weight:600;color:#323130;font-family:${font};">Scope</th>
        <th style="padding:8px 16px;text-align:left;font-size:12px;font-weight:600;color:#323130;font-family:${font};">Objects</th>
        <th style="padding:8px 16px;text-align:left;font-size:12px;font-weight:600;color:#323130;font-family:${font};">Unique</th>
        <th style="padding:8px 16px;"></th>
      </tr>
    </thead>
    <tbody>${rows}</tbody>
  </table>
</body></html>`;
}

// ── Image: Permissions Explorer — External users only filter ────────────────
function explorerExternalPage() {
  return `<!DOCTYPE html><html><head><meta charset="utf-8">
<style>
  * { margin:0; padding:0; box-sizing:border-box; }
  body { background:#fff; width:900px; font-family:${font}; color:#323130; }
  .tree-item { display:flex; align-items:center; gap:5px; padding:5px 6px; cursor:pointer; border-radius:3px; user-select:none; font-size:13px; }
  .tree-item.selected { background:#EFF6FC; }
  .tree-item.dim { color:#C8C6C4; }
</style>
</head><body>
  ${appBanner()}
  ${headerBar('Permissions Explorer', '')}

  <div style="padding:10px 18px 0;">
    <div style="margin-bottom:10px;">
      <div style="font-size:13px;font-weight:600;color:#323130;margin-bottom:4px;font-family:${font};">Library</div>
      <div style="display:inline-flex;align-items:center;gap:8px;border:1px solid #d1d1d1;
        border-radius:4px;padding:6px 10px;background:#fff;font-size:13px;color:#323130;min-width:220px;font-family:${font};">
        Campaign Documents
        <span style="margin-left:auto;color:#605E5C;font-size:10px;">▼</span>
      </div>
    </div>

    <div style="display:flex;gap:8px;margin-bottom:8px;">
      <div style="display:inline-flex;align-items:center;gap:6px;border:1px solid #d1d1d1;
        border-radius:4px;padding:5px 10px;background:#fff;font-size:12px;color:#323130;font-family:${font};">
        <span>≡</span> Unique permissions only
      </div>
      <div style="display:inline-flex;align-items:center;gap:6px;border:1.5px solid #D83B01;
        border-radius:4px;padding:5px 10px;background:#FDE7E9;font-size:12px;color:#A4262C;font-weight:600;font-family:${font};">
        <span style="font-size:14px;">⚠👤</span> External users only
      </div>
      <span style="align-self:center;font-size:12px;color:#605E5C;font-family:${font};">Filter active — showing 3 of 12 items</span>
    </div>

    <div style="display:flex;align-items:center;gap:12px;font-size:11px;color:#605E5C;margin-bottom:8px;flex-wrap:wrap;font-family:${font};">
      <span>Legend:</span>
      <span style="color:#A4262C;">⚠ External user access on this item</span>
      <span style="color:#A4262C;">△ Contains external user access below</span>
    </div>
  </div>

  <div style="display:flex;border-top:1px solid #EDEBE9;height:360px;">

    <!-- Left: tree, filtered to external-access items only -->
    <div style="width:320px;border-right:1px solid #EDEBE9;overflow-y:auto;padding:6px 4px;flex-shrink:0;">
      <div class="tree-item" style="padding-left:6px;">
        <span style="font-size:11px;color:#605E5C;width:12px;">▼</span>
        <span style="font-size:16px;">📁</span>
        <span style="flex:1;">Campaign Documents</span>
        <span style="font-size:13px;color:#A4262C;" title="Contains external user access below">△</span>
      </div>
      <div class="tree-item selected" style="padding-left:24px;">
        <span style="font-size:11px;color:#605E5C;width:12px;">▼</span>
        <span style="font-size:16px;">📁</span>
        <span style="flex:1;font-weight:600;">Q4 Launch</span>
        <span style="font-size:13px;color:#A4262C;">△</span>
        ${badge('Unique', '#FFF4CE', '#7D4A00')}
        <span style="font-size:14px;color:#A4262C;" title="External user access on this item">⚠</span>
      </div>
      <div class="tree-item" style="padding-left:42px;">
        <span style="width:12px;"></span>
        <span style="font-size:16px;">📗</span>
        <span style="flex:1;">Budget_Confidential.xlsx</span>
        ${badge('Unique', '#FFF4CE', '#7D4A00')}
        <span style="font-size:14px;color:#A4262C;" title="External user access on this item">⚠</span>
      </div>
      <div class="tree-item" style="padding-left:24px;">
        <span style="font-size:11px;color:#605E5C;width:12px;">▶</span>
        <span style="font-size:16px;">📁</span>
        <span style="flex:1;">Executive Reports</span>
        <span style="font-size:13px;color:#A4262C;">△</span>
      </div>
    </div>

    <!-- Right: permissions panel, filtered to external rows only -->
    <div style="flex:1;padding:14px 16px;overflow-y:auto;">
      <div style="display:flex;align-items:center;gap:10px;margin-bottom:6px;">
        <span style="font-size:15px;">📁</span>
        <span style="font-size:15px;font-weight:600;color:#323130;">Q4 Launch</span>
        ${badge('Unique permissions', '#FFF4CE', '#7D4A00')}
      </div>
      <div style="font-size:12px;color:#605E5C;margin-bottom:12px;">
        Showing external-user rows only (External users only filter is active).
      </div>

      <table style="width:100%;border-collapse:collapse;">
        <thead>
          <tr style="border-bottom:1px solid #EDEBE9;">
            <th style="text-align:left;padding:6px 8px 8px 0;font-size:12px;font-weight:600;color:#323130;">User / Group</th>
            <th style="text-align:left;padding:6px 8px 8px;font-size:12px;font-weight:600;color:#323130;width:80px;">Type</th>
            <th style="text-align:left;padding:6px 0 8px 8px;font-size:12px;font-weight:600;color:#323130;width:120px;">Permission Level</th>
          </tr>
        </thead>
        <tbody>
          ${[
            { name: 'Alex Thompson\nalex@externalpartner.com', level: 'Read', color: '#DFF6DD', tcolor: '#107C10' },
            { name: 'Priya Nair\npriya.nair@vendorco.example', level: 'Edit', color: '#FFF4CE', tcolor: '#7D4A00' },
          ].map((r, i) => `
          <tr style="border-bottom:1px solid #F3F2F1;background:#FDF7F7;">
            <td style="padding:8px 8px 8px 0;font-size:13px;color:#323130;">
              <div style="display:flex;align-items:flex-start;gap:7px;">
                <span style="font-size:15px;margin-top:1px;color:#A4262C;">⚠</span>
                <div>
                  <div style="font-size:13px;color:#323130;">${r.name.split('\n')[0]}</div>
                  <div style="font-size:11px;color:#605E5C;">${r.name.split('\n')[1]}</div>
                </div>
              </div>
            </td>
            <td style="padding:8px;font-size:13px;color:#605E5C;">User</td>
            <td style="padding:8px 0 8px 8px;">${badge(r.level, r.color, r.tcolor)}</td>
          </tr>`).join('')}
        </tbody>
      </table>
    </div>
  </div>
</body></html>`;
}

async function main() {
  const browser = await puppeteer.launch({
    executablePath: CHROME,
    headless: 'new',
    args: ['--no-sandbox', '--disable-setuid-sandbox', '--font-render-hinting=none'],
  });

  const shots = [
    ['10_report_history.png', reportHistoryPage, { width: 900, height: 430 }],
    ['11_explorer_external.png', explorerExternalPage, { width: 900, height: 620 }],
  ];

  for (const [filename, htmlFn, vp] of shots) {
    const pg = await browser.newPage();
    await pg.setViewport(vp);
    await pg.setContent(htmlFn(), { waitUntil: 'networkidle0' });
    await pg.screenshot({ path: path.join(DOCS_OUT, filename), clip: { x: 0, y: 0, width: vp.width, height: vp.height } });
    await pg.close();
    console.log('✓', filename);
  }

  await browser.close();
  console.log(`\nDocs screenshots → ${DOCS_OUT}`);
}

main().catch(e => { console.error(e); process.exit(1); });
