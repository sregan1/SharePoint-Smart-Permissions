'use strict';
// Generates custom card-header images for the HomeView.
// Output: src/webparts/smartPermissions/assets/screenshot_*.png
// Also generates docs/screenshots/01_home.png for the README.
// Run: node docs/generate-screenshots.js

const puppeteer = require('puppeteer-core');
const path = require('path');
const fs = require('fs');

const CHROME = fs.existsSync('C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe')
  ? 'C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe'
  : 'C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedge.exe';

const OUT = path.join(__dirname, '..', 'src', 'webparts', 'smartPermissions', 'assets');
if (!fs.existsSync(OUT)) fs.mkdirSync(OUT, { recursive: true });

const DOCS_OUT = path.join(__dirname, 'screenshots');
if (!fs.existsSync(DOCS_OUT)) fs.mkdirSync(DOCS_OUT, { recursive: true });

const font = `'Segoe UI', Arial, sans-serif`;

// ── Shared UI primitives ──────────────────────────────────────────────────────
function appBanner() {
  return `
  <div style="background:#0078D4;display:flex;align-items:center;padding:0 14px;height:44px;gap:10px;flex-shrink:0;">
    <svg width="18" height="18" viewBox="0 0 24 24" fill="white">
      <path d="M12 1L3 5v6c0 5.55 3.84 10.74 9 12 5.16-1.26 9-6.45 9-12V5l-9-4z"/>
    </svg>
    <span style="color:#fff;font-size:14px;font-weight:600;font-family:${font};flex-shrink:0;">SharePoint Smart Permissions</span>
    <span style="color:rgba(255,255,255,0.7);font-size:12px;font-family:${font};margin-left:8px;">●</span>
    <span style="color:rgba(255,255,255,0.85);font-size:12px;font-family:${font};">https://contoso.sharepoint.com/sites/Marketing</span>
    <span style="color:rgba(255,255,255,0.85);font-size:12px;font-family:${font};margin-left:4px;text-decoration:underline;cursor:pointer;">Change URL</span>
    <svg style="margin-left:auto;" width="16" height="16" viewBox="0 0 24 24" fill="white" opacity="0.9">
      <path d="M19.14 12.94c.04-.3.06-.61.06-.94s-.02-.64-.07-.94l2.03-1.58a.49.49 0 0 0 .12-.61l-1.92-3.32a.49.49 0 0 0-.59-.22l-2.39.96a7.02 7.02 0 0 0-1.62-.94l-.36-2.54A.484.484 0 0 0 14 2h-4a.484.484 0 0 0-.48.41l-.36 2.54a7.36 7.36 0 0 0-1.62.94l-2.39-.96a.48.48 0 0 0-.59.22L2.74 8.87a.47.47 0 0 0 .12.61l2.03 1.58c-.05.3-.09.63-.09.94s.02.64.07.94l-2.03 1.58a.49.49 0 0 0-.12.61l1.92 3.32c.12.22.37.29.59.22l2.39-.96c.5.36 1.04.67 1.62.94l.36 2.54c.05.24.27.41.49.41h4c.22 0 .44-.17.47-.41l.36-2.54a7.36 7.36 0 0 0 1.62-.94l2.39.96c.22.08.47 0 .59-.22l1.92-3.32a.47.47 0 0 0-.12-.61l-2.01-1.58zM12 15.6c-1.98 0-3.6-1.62-3.6-3.6s1.62-3.6 3.6-3.6 3.6 1.62 3.6 3.6-1.62 3.6-3.6 3.6z"/>
    </svg>
  </div>`;
}

function backButton() {
  return `<div style="display:inline-flex;align-items:center;gap:6px;border:1px solid #d1d1d1;
    border-radius:4px;padding:5px 12px;font-size:13px;color:#323130;font-family:${font};
    cursor:pointer;background:#fff;margin-bottom:16px;">← Back</div>`;
}

function typeBadge(type) {
  const map = {
    Site:    { bg: '#D0E7FF', color: '#0050A0' },
    Library: { bg: '#E6F4EA', color: '#1A6B2E' },
    Folder:  { bg: '#FFF0CC', color: '#7D4A00' },
    File:    { bg: '#F3F2F1', color: '#605E5C' },
  };
  const s = map[type] || map.File;
  return `<span style="display:inline-block;padding:1px 8px;border-radius:10px;font-size:11px;
    font-weight:600;background:${s.bg};color:${s.color};font-family:${font};white-space:nowrap;">${type}</span>`;
}

function permBadge(level) {
  if (!level) return '';
  const map = {
    'Full Control': { bg: '#FDE7E9', color: '#A4262C' },
    'Edit':         { bg: '#FFF4CE', color: '#7D4A00' },
    'Read':         { bg: '#DFF6DD', color: '#107C10' },
  };
  const s = map[level] || { bg: '#F3F2F1', color: '#323130' };
  return `<span style="display:inline-block;padding:2px 10px;border-radius:10px;font-size:11px;
    font-weight:600;background:${s.bg};color:${s.color};font-family:${font};white-space:nowrap;">${level}</span>`;
}

function uniqueBadge() {
  return `<span style="display:inline-block;padding:2px 8px;border-radius:10px;font-size:11px;
    font-weight:600;background:#FFF4CE;color:#7D4A00;font-family:${font};">Unique</span>`;
}

// ── Image 1: Excel output ─────────────────────────────────────────────────────
function reportPage() {
  const rows = [
    { type: 'Site',    name: 'Marketing',                   path: '/sites/Marketing',                                          unique: 'Yes', level: 'Read',         users: 'Marketing Owners'           },
    { type: 'Library', name: 'Campaign Documents',          path: '/sites/Marketing/Campaign Documents',                       unique: 'No',  level: '',             users: ''                           },
    { type: 'Folder',  name: 'Q4 Launch',                   path: '/sites/Marketing/Campaign Documents/Q4 Launch',             unique: 'Yes', level: 'Edit',         users: 'Marketing Members; Alex T.' },
    { type: 'File',    name: 'Budget_Confidential.xlsx',    path: '/sites/Marketing/Campaign Documents/Q4 Launch/Budget_C...', unique: 'Yes', level: 'Full Control', users: 'jane.doe@contoso.com'       },
    { type: 'Library', name: 'Marketing Assets',            path: '/sites/Marketing/Marketing Assets',                         unique: 'No',  level: '',             users: ''                           },
    { type: 'Folder',  name: 'Brand Photos',                path: '/sites/Marketing/Marketing Assets/Brand Photos',            unique: 'Yes', level: 'Edit',         users: 'Design Team'                },
    { type: 'Folder',  name: 'Q1 Planning',                 path: '/sites/Marketing/Campaign Documents/Q1 Planning',           unique: 'No',  level: '',             users: ''                           },
    { type: 'Folder',  name: 'Archive',                     path: '/sites/Marketing/Campaign Documents/Archive',               unique: 'No',  level: '',             users: ''                           },
    { type: 'File',    name: 'Campaign_Brief.docx',         path: '/sites/Marketing/Campaign Documents/Q4 Launch/Campaign_...', unique: 'No',  level: '',             users: ''                               },
    { type: 'Folder',  name: 'Executive Reports',           path: '/sites/Marketing/Campaign Documents/Executive Reports',        unique: 'Yes', level: 'Full Control', users: 'Marketing Owners; CEO-Group'    },
    { type: 'File',    name: 'Board_Deck_Oct.pptx',         path: '/sites/Marketing/Campaign Documents/Executive Reports/Boa...', unique: 'Yes', level: 'Full Control', users: 'CEO-Group'                      },
    { type: 'Library', name: 'Site Pages',                  path: '/sites/Marketing/Site Pages',                                  unique: 'No',  level: '',             users: ''                               },
    { type: 'Folder',  name: 'Press Releases',              path: '/sites/Marketing/Campaign Documents/Press Releases',           unique: 'Yes', level: 'Edit',         users: 'PR Team; Marketing Members'     },
    { type: 'File',    name: 'PR_Oct_Launch.docx',          path: '/sites/Marketing/Campaign Documents/Press Releases/PR_Oc...', unique: 'No',  level: '',             users: ''                               },
    { type: 'File',    name: 'Annual_Report_2025.pdf',      path: '/sites/Marketing/Marketing Assets/Annual_Report_2025.pdf',     unique: 'Yes', level: 'Read',         users: 'All Company'                    },
  ];

  const cols = ['Type', 'Name', 'Path', 'Has Unique Permissions', 'Permission Level', 'Users / Groups'];
  const colWidths = ['70px', '155px', '195px', '110px', '105px', 'auto'];

  return `<!DOCTYPE html><html><head><meta charset="utf-8">
<style>*{margin:0;padding:0;box-sizing:border-box;}</style>
</head><body style="background:#F0F0F0;width:800px;font-family:${font};">

  <!-- Window chrome -->
  <div style="background:#217346;display:flex;align-items:center;padding:0 12px;height:32px;gap:8px;">
    <svg width="16" height="16" viewBox="0 0 24 24"><rect width="24" height="24" rx="3" fill="#fff" opacity="0.15"/>
      <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8l-6-6z" fill="#fff" opacity="0.9"/>
      <path d="M8 12h8M8 15h5" stroke="#217346" stroke-width="1.5" stroke-linecap="round"/>
    </svg>
    <span style="color:#fff;font-size:12px;flex:1;">SharePoint Permissions - Marketing Site.xlsx - Excel</span>
    <div style="display:flex;gap:0;">
      <div style="width:46px;height:32px;display:flex;align-items:center;justify-content:center;color:#fff;font-size:18px;opacity:0.8;cursor:pointer;">─</div>
      <div style="width:46px;height:32px;display:flex;align-items:center;justify-content:center;color:#fff;font-size:14px;opacity:0.8;cursor:pointer;">□</div>
      <div style="width:46px;height:32px;display:flex;align-items:center;justify-content:center;color:#fff;font-size:14px;opacity:0.8;cursor:pointer;">✕</div>
    </div>
  </div>

  <!-- Ribbon tabs -->
  <div style="background:#217346;display:flex;align-items:flex-end;padding:0 8px;border-bottom:1px solid #1a5c38;">
    ${['File','Home','Insert','Page Layout','Formulas','Data','Review','View'].map((t, i) =>
      `<div style="padding:5px 12px;font-size:12px;color:${i === 1 ? '#217346' : 'rgba(255,255,255,0.85)'};
        background:${i === 1 ? '#fff' : 'transparent'};border-radius:3px 3px 0 0;cursor:pointer;
        font-family:${font};">${t}</div>`
    ).join('')}
  </div>

  <!-- Formula bar -->
  <div style="background:#fff;display:flex;align-items:center;border-bottom:1px solid #D4D4D4;height:26px;gap:0;">
    <div style="width:52px;border-right:1px solid #D4D4D4;height:100%;display:flex;align-items:center;
      justify-content:center;font-size:12px;color:#323130;">A1</div>
    <div style="width:24px;height:100%;display:flex;align-items:center;justify-content:center;
      border-right:1px solid #D4D4D4;font-size:13px;color:#605E5C;">ƒx</div>
    <div style="padding:0 8px;font-size:12px;color:#323130;">Type</div>
  </div>

  <!-- Spreadsheet grid -->
  <div style="background:#fff;overflow:hidden;">

    <!-- Column letter headers -->
    <div style="display:flex;border-bottom:1px solid #D4D4D4;background:#F2F2F2;">
      <div style="width:36px;border-right:1px solid #D4D4D4;height:20px;flex-shrink:0;"></div>
      ${cols.map((c, i) => `<div style="width:${colWidths[i]};${i === cols.length-1 ? 'flex:1;' : ''}border-right:1px solid #D4D4D4;
        height:20px;display:flex;align-items:center;justify-content:center;
        font-size:11px;color:#605E5C;font-family:${font};">
        ${String.fromCharCode(65+i)}
      </div>`).join('')}
    </div>

    <!-- Header row (row 1) -->
    <div style="display:flex;border-bottom:2px solid #1a5c38;background:#1F5C8B;">
      <div style="width:36px;border-right:1px solid #D4D4D4;height:24px;flex-shrink:0;
        display:flex;align-items:center;justify-content:center;font-size:11px;color:#605E5C;
        background:#F2F2F2;">1</div>
      ${cols.map((c, i) => `<div style="width:${colWidths[i]};${i === cols.length-1 ? 'flex:1;' : ''}
        border-right:1px solid #2a6e9e;padding:4px 8px;font-size:11px;font-weight:700;
        color:#fff;font-family:${font};">${c}</div>`).join('')}
    </div>

    <!-- Data rows -->
    <table style="width:100%;border-collapse:collapse;">
      <colgroup>
        <col style="width:36px;"/>
        <col style="width:70px;"/>
        <col style="width:155px;"/>
        <col style="width:195px;"/>
        <col style="width:110px;"/>
        <col style="width:105px;"/>
        <col/>
      </colgroup>
      ${rows.map((r, i) => {
        const bg = r.unique === 'Yes' ? '#EBF3FB' : (i % 2 === 0 ? '#FFFFFF' : '#F9F9F9');
        const rowNum = i + 2;
        const levelStyle = r.level
          ? { 'Full Control': { bg: '#FDE7E9', color: '#A4262C' }, 'Edit': { bg: '#FFF4CE', color: '#7D4A00' }, 'Read': { bg: '#DFF6DD', color: '#107C10' } }[r.level] || {}
          : {};
        return `<tr style="background:${bg};">
          <td style="padding:3px 0;border-right:1px solid #D4D4D4;border-bottom:1px solid #E8E8E8;
            text-align:center;font-size:11px;color:#8A8886;font-family:${font};
            background:#F2F2F2;">${rowNum}</td>
          <td style="padding:3px 8px;border-right:1px solid #D4D4D4;border-bottom:1px solid #E8E8E8;
            font-size:11px;color:#323130;font-family:${font};">${r.type}</td>
          <td style="padding:3px 8px;border-right:1px solid #D4D4D4;border-bottom:1px solid #E8E8E8;
            font-size:11px;color:#323130;font-family:${font};font-weight:${r.unique==='Yes'?'600':'400'};">${r.name}</td>
          <td style="padding:3px 8px;border-right:1px solid #D4D4D4;border-bottom:1px solid #E8E8E8;
            font-size:10px;color:#605E5C;font-family:${font};overflow:hidden;white-space:nowrap;max-width:195px;">${r.path}</td>
          <td style="padding:3px 8px;border-right:1px solid #D4D4D4;border-bottom:1px solid #E8E8E8;
            text-align:center;font-size:11px;font-family:${font};
            ${r.unique==='Yes' ? 'color:#107C10;font-weight:700;' : 'color:#A19F9D;'}">${r.unique==='Yes'?'✓ Yes':'No'}</td>
          <td style="padding:3px 8px;border-right:1px solid #D4D4D4;border-bottom:1px solid #E8E8E8;
            font-size:11px;font-family:${font};
            ${r.level ? `background:${levelStyle.bg};color:${levelStyle.color};font-weight:600;` : ''}">${r.level}</td>
          <td style="padding:3px 8px;border-bottom:1px solid #E8E8E8;
            font-size:10px;color:#323130;font-family:${font};">${r.users}</td>
        </tr>`;
      }).join('')}
    </table>
  </div>

  <!-- Sheet tab bar -->
  <div style="background:#F2F2F2;border-top:1px solid #D4D4D4;display:flex;align-items:center;
    height:26px;padding:0 8px;gap:0;">
    <div style="background:#fff;border:1px solid #D4D4D4;border-bottom:none;padding:3px 12px;
      font-size:11px;color:#323130;font-family:${font};border-radius:3px 3px 0 0;">Permissions</div>
    <div style="padding:3px 12px;font-size:11px;color:#605E5C;font-family:${font};">Summary</div>
  </div>
</body></html>`;
}

// ── Image 2: Permissions Explorer (matches real app) ──────────────────────────
function explorerPage() {
  return `<!DOCTYPE html><html><head><meta charset="utf-8">
<style>
  * { margin:0; padding:0; box-sizing:border-box; }
  body { background:#fff; width:800px; font-family:${font}; color:#323130; }
  .tree-item { display:flex; align-items:center; gap:5px; padding:5px 6px; cursor:pointer; border-radius:3px; user-select:none; font-size:13px; }
  .tree-item.selected { background:#EFF6FC; }
  .tree-item:hover { background:#F3F2F1; }
</style>
</head><body>
  ${appBanner()}

  <div style="padding:14px 18px 0;">
    <h1 style="font-size:22px;font-weight:600;color:#323130;margin-bottom:4px;">Permissions Explorer</h1>
    <div style="font-size:13px;color:#0078D4;margin-bottom:14px;">Connected — 4 libraries found</div>

    <div style="margin-bottom:10px;">
      <div style="font-size:13px;font-weight:600;color:#323130;margin-bottom:4px;">Library</div>
      <div style="display:inline-flex;align-items:center;gap:8px;border:1px solid #d1d1d1;
        border-radius:4px;padding:6px 10px;background:#fff;font-size:13px;color:#323130;min-width:200px;cursor:pointer;">
        Campaign Documents
        <span style="margin-left:auto;color:#605E5C;font-size:10px;">▼</span>
      </div>
    </div>

    <div style="display:flex;gap:8px;margin-bottom:8px;">
      <div style="display:inline-flex;align-items:center;gap:6px;border:1px solid #d1d1d1;
        border-radius:4px;padding:5px 10px;background:#fff;font-size:12px;color:#323130;cursor:pointer;">
        <span style="font-size:14px;">≡</span> Unique permissions only
      </div>
      <div style="display:inline-flex;align-items:center;gap:6px;border:1px solid #d1d1d1;
        border-radius:4px;padding:5px 10px;background:#fff;font-size:12px;color:#323130;cursor:pointer;">
        <span style="font-size:14px;">👤</span> External users only
      </div>
    </div>

    <div style="display:flex;align-items:center;gap:12px;font-size:11px;color:#605E5C;margin-bottom:10px;flex-wrap:wrap;">
      <span>Legend:</span>
      <span style="background:#FFF4CE;color:#7D4A00;padding:1px 7px;border-radius:10px;font-weight:600;font-size:11px;">Unique</span>
      <span>Item has unique permissions</span>
      <span style="color:#323130;">⊙ Contains unique permissions below</span>
      <span style="color:#323130;">△ Contains external user access below</span>
    </div>
  </div>

  <!-- Split panel -->
  <div style="display:flex;border-top:1px solid #EDEBE9;height:340px;">

    <!-- Left: tree -->
    <div style="width:295px;border-right:1px solid #EDEBE9;overflow-y:auto;padding:6px 4px;flex-shrink:0;">
      <!-- Campaign Documents -->
      <div class="tree-item" style="padding-left:6px;">
        <span style="font-size:11px;color:#605E5C;width:12px;">▼</span>
        <span style="font-size:16px;">📁</span>
        <span style="flex:1;">Campaign Documents</span>
        <span style="font-size:13px;color:#0078D4;" title="Contains unique permissions below">⊙</span>
        <span style="font-size:13px;color:#D83B01;" title="Contains external users below">△</span>
      </div>
      <!-- Q4 Launch -->
      <div class="tree-item selected" style="padding-left:24px;">
        <span style="font-size:11px;color:#605E5C;width:12px;">▼</span>
        <span style="font-size:16px;">📁</span>
        <span style="flex:1;font-weight:600;">Q4 Launch</span>
        <span style="font-size:13px;color:#0078D4;">⊙</span>
        <span style="font-size:13px;color:#D83B01;">△</span>
        ${uniqueBadge()}
        <span style="font-size:14px;color:#D83B01;" title="External user access on this item">👤</span>
      </div>
      <!-- Assets -->
      <div class="tree-item" style="padding-left:42px;">
        <span style="font-size:11px;color:#605E5C;width:12px;">▶</span>
        <span style="font-size:16px;">📁</span>
        <span style="flex:1;">Assets</span>
        <span style="font-size:13px;color:#0078D4;">⊙</span>
      </div>
      <!-- Campaign_Brief.docx -->
      <div class="tree-item" style="padding-left:42px;">
        <span style="width:12px;"></span>
        <span style="font-size:16px;">📝</span>
        <span style="flex:1;color:#605E5C;">Campaign_Brief.docx</span>
      </div>
      <!-- Budget_Confidential.xlsx -->
      <div class="tree-item" style="padding-left:42px;">
        <span style="width:12px;"></span>
        <span style="font-size:16px;">📗</span>
        <span style="flex:1;">Budget_Confidential.xlsx</span>
        ${uniqueBadge()}
        <span style="font-size:11px;color:#605E5C;">·</span>
      </div>
      <!-- Q1 Planning -->
      <div class="tree-item" style="padding-left:24px;">
        <span style="font-size:11px;color:#605E5C;width:12px;">▶</span>
        <span style="font-size:16px;">📁</span>
        <span style="flex:1;">Q1 Planning</span>
      </div>
      <!-- Archive -->
      <div class="tree-item" style="padding-left:24px;">
        <span style="font-size:11px;color:#605E5C;width:12px;">▶</span>
        <span style="font-size:16px;">📁</span>
        <span style="flex:1;">Archive</span>
      </div>
    </div>

    <!-- Right: permissions panel -->
    <div style="flex:1;padding:14px 16px;overflow-y:auto;">
      <div style="display:flex;align-items:center;gap:10px;margin-bottom:12px;">
        <span style="font-size:15px;">📁</span>
        <span style="font-size:15px;font-weight:600;color:#323130;">Q4 Launch</span>
        <span style="background:#FFF4CE;color:#7D4A00;padding:2px 10px;border-radius:10px;
          font-size:12px;font-weight:600;">Unique permissions</span>
      </div>

      <div style="display:flex;align-items:center;gap:8px;margin-bottom:12px;">
        <div style="width:16px;height:16px;border:1.5px solid #d1d1d1;border-radius:2px;background:#fff;"></div>
        <span style="font-size:13px;color:#323130;">Expand group members</span>
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
            { name: 'Marketing Owners',                          type: 'SP Group', level: 'Full Control', icon: '👥' },
            { name: 'Marketing Members',                         type: 'SP Group', level: 'Edit',         icon: '👥' },
            { name: 'Alex Thompson\nalex@externalpartner.com',   type: 'User',     level: 'Read',         icon: '👤', ext: true },
            { name: 'Marketing Visitors',                        type: 'SP Group', level: 'Read',         icon: '👥' },
          ].map((r, i) => `
          <tr style="border-bottom:1px solid #F3F2F1;${i % 2 !== 0 ? 'background:#FAFAFA;' : ''}">
            <td style="padding:8px 8px 8px 0;font-size:13px;color:#323130;">
              <div style="display:flex;align-items:flex-start;gap:7px;">
                <span style="font-size:15px;margin-top:1px;">${r.icon}</span>
                <div>
                  ${r.name.split('\n').map((line, li) =>
                    `<div style="font-size:${li===0?'13px':'11px'};color:${li===0?'#323130':'#605E5C'};">${line}</div>`
                  ).join('')}
                </div>
              </div>
            </td>
            <td style="padding:8px;font-size:13px;color:#605E5C;">${r.type}</td>
            <td style="padding:8px 0 8px 8px;">
              ${permBadge(r.level)}
              ${r.ext ? `<span style="font-size:14px;margin-left:6px;" title="External user">👤</span>` : ''}
            </td>
          </tr>`).join('')}
        </tbody>
      </table>
    </div>
  </div>
</body></html>`;
}

// ── Image 3: User Access (matches real app) ────────────────────────────────────
function userAccessPage() {
  const results = [
    { type: 'Site',    name: 'Marketing',                   path: '/sites/Marketing',                                    level: 'Read'         },
    { type: 'Library', name: 'Campaign Documents',          path: '/sites/Marketing/Campaign Documents',                 level: 'Edit'         },
    { type: 'Folder',  name: 'Q4 Launch',                   path: '/sites/Marketing/Campaign Documents/Q4 Launch',       level: 'Edit'         },
    { type: 'File',    name: 'Budget_Confidential.xlsx',    path: '/sites/Marketing/Campaign Documents/Q4 Launch/...',   level: 'Full Control' },
    { type: 'Library', name: 'Marketing Assets',            path: '/sites/Marketing/Marketing Assets',                   level: 'Read'         },
    { type: 'Folder',  name: 'Brand Photos',                path: '/sites/Marketing/Marketing Assets/Brand Photos',      level: 'Edit'         },
    { type: 'Folder',  name: 'Q1 Planning',                 path: '/sites/Marketing/Campaign Documents/Q1 Planning',      level: 'Read'         },
  ];

  return `<!DOCTYPE html><html><head><meta charset="utf-8">
<style>* { margin:0; padding:0; box-sizing:border-box; }</style>
</head><body style="background:#fff;width:800px;font-family:${font};color:#323130;">
  ${appBanner()}
  <div style="padding:14px 18px 0;">
    <div style="display:flex;align-items:center;gap:10px;margin-bottom:4px;">
      <h1 style="font-size:22px;font-weight:600;color:#323130;">User Access</h1>
      <div style="margin-left:auto;display:flex;gap:8px;">
        <div style="border:1px solid #d1d1d1;border-radius:4px;padding:5px 12px;
          font-size:13px;color:#323130;cursor:pointer;background:#fff;">New scan</div>
        <div style="border:1px solid #d1d1d1;border-radius:4px;padding:5px 12px;
          font-size:13px;color:#323130;cursor:pointer;background:#fff;">⌚ History (3)</div>
      </div>
    </div>
    <div style="font-size:13px;color:#0078D4;margin-bottom:14px;">Connected — 18 users found</div>

    <div style="margin-bottom:12px;">
      <div style="font-size:13px;font-weight:600;color:#323130;margin-bottom:4px;">Select a user</div>
      <div style="display:inline-flex;align-items:center;gap:8px;border:1px solid #d1d1d1;
        border-radius:4px;padding:6px 10px;background:#fff;font-size:13px;color:#323130;
        min-width:240px;cursor:pointer;">
        Alice Chen
        <span style="margin-left:auto;color:#605E5C;font-size:10px;">▼</span>
      </div>
    </div>

    <!-- Info banner -->
    <div style="background:#EFF6FC;border-left:3px solid #0078D4;border-radius:0 4px 4px 0;
      padding:10px 14px;margin-bottom:12px;display:flex;gap:10px;align-items:flex-start;">
      <span style="color:#0078D4;font-size:16px;flex-shrink:0;">ℹ</span>
      <span style="font-size:12.5px;color:#323130;line-height:1.5;">
        This user has site-level access. Only locations with <strong>unique permission assignments</strong> appear below — all
        other content is accessible through the site-level permission shown at the top of the list.
      </span>
    </div>

    <div style="display:flex;align-items:center;margin-bottom:10px;">
      <span style="font-size:13px;color:#323130;">6 accessible location(s) found.</span>
      <div style="margin-left:auto;display:flex;gap:8px;">
        <div style="border:1px solid #d1d1d1;border-radius:4px;padding:5px 12px;
          font-size:13px;color:#323130;cursor:pointer;background:#fff;">Export to Excel</div>
        <div style="border:1px solid #d1d1d1;border-radius:4px;padding:5px 12px;
          font-size:13px;color:#323130;cursor:pointer;background:#fff;">Export to CSV</div>
      </div>
    </div>
  </div>

  <!-- Results table -->
  <table style="width:100%;border-collapse:collapse;margin-top:2px;">
    <thead>
      <tr style="border-bottom:1px solid #EDEBE9;">
        <th style="padding:8px 18px;text-align:left;font-size:12px;font-weight:600;color:#323130;width:85px;">
          Type <span style="color:#0078D4;font-size:10px;">▲</span>
        </th>
        <th style="padding:8px 8px;text-align:left;font-size:12px;font-weight:600;color:#323130;">Name</th>
        <th style="padding:8px 8px;text-align:left;font-size:12px;font-weight:600;color:#323130;">Path</th>
        <th style="padding:8px 18px;text-align:left;font-size:12px;font-weight:600;color:#323130;width:130px;">Permission Level</th>
      </tr>
    </thead>
    <tbody>
      ${results.map((r, i) => `
      <tr style="border-bottom:1px solid #F3F2F1;${i % 2 !== 0 ? 'background:#FAFAFA;' : ''}">
        <td style="padding:9px 18px;">${typeBadge(r.type)}</td>
        <td style="padding:9px 8px;font-size:13px;color:#323130;">${r.name}</td>
        <td style="padding:9px 8px;font-size:12px;color:#605E5C;">${r.path}</td>
        <td style="padding:9px 18px;">${permBadge(r.level)}</td>
      </tr>`).join('')}
    </tbody>
  </table>
</body></html>`;
}

// ── Image 4: Permission Groups browser ───────────────────────────────────────
function groupsPage() {
  const groups = [
    {
      name: 'Marketing Members',
      desc: 'Members of the Marketing site',
      members: [
        { name: 'Alice Chen',    type: 'User',  login: 'alice.chen@contoso.com'       },
        { name: 'Bob Martinez',  type: 'User',  login: 'bob.martinez@contoso.com'     },
        { name: 'Design Team',   type: 'M365',  login: 'c:0o.c|fed...|a1b2c3d4'      },
      ],
      expanded: true,
    },
    {
      name: 'Marketing Owners',
      desc: 'Owners of the Marketing site',
      members: [
        { name: 'Jane Smith',    type: 'User',  login: 'jane.smith@contoso.com'       },
      ],
      expanded: true,
    },
    {
      name: 'Marketing Visitors',
      desc: 'Visitors of the Marketing site',
      members: [],
      expanded: false,
    },
    {
      name: 'PR Editors',
      desc: 'Press release editors — restricted folder access',
      members: [],
      expanded: false,
    },
    {
      name: 'Executive Access',
      desc: 'Board reports and confidential documents',
      members: [],
      expanded: false,
    },
  ];

  const typeColor = { User: '#605E5C', M365: '#0050A0' };
  const typeBg   = { User: '#F3F2F1', M365: '#D0E7FF' };

  function memberRows(members) {
    if (!members.length) return `<div style="padding:8px 16px;font-size:12px;color:#A19F9D;">This group has no members.</div>`;
    return members.map(m => `
      <div style="display:flex;align-items:center;gap:8px;padding:7px 16px;border-bottom:1px solid #F3F2F1;">
        <svg width="14" height="14" viewBox="0 0 24 24" fill="${typeColor[m.type] || '#605E5C'}" style="flex-shrink:0;">
          ${m.type === 'User'
            ? '<path d="M12 12a5 5 0 1 0 0-10 5 5 0 0 0 0 10zm0 2c-5.33 0-8 2.67-8 4v2h16v-2c0-1.33-2.67-4-8-4z"/>'
            : '<path d="M16 11c1.66 0 2.99-1.34 2.99-3S17.66 5 16 5c-1.66 0-3 1.34-3 3s1.34 3 3 3zm-8 0c1.66 0 2.99-1.34 2.99-3S9.66 5 8 5C6.34 5 5 6.34 5 8s1.34 3 3 3zm0 2c-2.33 0-7 1.17-7 3.5V19h14v-2.5c0-2.33-4.67-3.5-7-3.5zm8 0c-.29 0-.62.02-.97.05 1.16.84 1.97 1.97 1.97 3.45V19h6v-2.5c0-2.33-4.67-3.5-7-3.5z"/>'}
        </svg>
        <span style="flex:1;font-size:13px;color:#323130;font-family:${font};">${m.name}</span>
        <span style="padding:2px 8px;border-radius:10px;font-size:11px;font-weight:600;
          background:${typeBg[m.type]||'#F3F2F1'};color:${typeColor[m.type]||'#605E5C'};
          font-family:${font};white-space:nowrap;">${m.type === 'M365' ? 'M365 Group' : 'User'}</span>
        <span style="font-size:11px;color:#A19F9D;font-family:monospace;overflow:hidden;
          text-overflow:ellipsis;white-space:nowrap;max-width:180px;">${m.login}</span>
      </div>`).join('');
  }

  function groupBlock(g) {
    const chevron = g.expanded
      ? `<path d="M7 10l5 5 5-5z" fill="#605E5C"/>`
      : `<path d="M10 17l5-5-5-5v10z" fill="#605E5C"/>`;
    const memberCount = g.expanded ? `` : (g.members.length > 0
      ? `<span style="padding:2px 8px;border-radius:10px;border:1px solid #EDEBE9;
           font-size:11px;color:#605E5C;font-family:${font};">${g.members.length} member${g.members.length !== 1 ? 's' : ''}</span>`
      : '');
    return `
    <div style="border:1px solid #EDEBE9;border-radius:4px;overflow:hidden;margin-bottom:8px;">
      <div style="display:flex;align-items:center;gap:8px;padding:10px 14px;
        background:${g.expanded ? '#FAFAFA' : '#fff'};cursor:pointer;">
        <svg width="16" height="16" viewBox="0 0 24 24" style="flex-shrink:0;">${chevron}</svg>
        <svg width="18" height="18" viewBox="0 0 24 24" fill="#0078D4" style="flex-shrink:0;">
          <path d="M16 11c1.66 0 2.99-1.34 2.99-3S17.66 5 16 5c-1.66 0-3 1.34-3 3s1.34 3 3 3zm-8 0c1.66 0 2.99-1.34 2.99-3S9.66 5 8 5C6.34 5 5 6.34 5 8s1.34 3 3 3zm0 2c-2.33 0-7 1.17-7 3.5V19h14v-2.5c0-2.33-4.67-3.5-7-3.5zm8 0c-.29 0-.62.02-.97.05 1.16.84 1.97 1.97 1.97 3.45V19h6v-2.5c0-2.33-4.67-3.5-7-3.5z"/>
        </svg>
        <div style="flex:1;min-width:0;">
          <div style="font-size:14px;font-weight:600;color:#323130;font-family:${font};">${g.name}</div>
          ${g.desc ? `<div style="font-size:12px;color:#605E5C;font-family:${font};">${g.desc}</div>` : ''}
        </div>
        ${memberCount}
      </div>
      ${g.expanded ? `<div style="border-top:1px solid #EDEBE9;background:#F9F8F7;">${memberRows(g.members)}</div>` : ''}
    </div>`;
  }

  return `<!DOCTYPE html><html><head><meta charset="utf-8">
<style>* { margin:0; padding:0; box-sizing:border-box; }</style>
</head><body style="background:#fff;width:800px;font-family:${font};color:#323130;">
  ${appBanner()}
  <div style="padding:14px 18px 16px;">
    <div style="display:flex;align-items:center;gap:10px;margin-bottom:4px;">
      ${backButton()}
    </div>
    <h1 style="font-size:22px;font-weight:600;color:#323130;margin-bottom:6px;">Permission Groups</h1>
    <div style="font-size:13px;color:#605E5C;margin-bottom:16px;">
      5 groups — click any group to see its members.
    </div>
    ${groups.map(groupBlock).join('')}
  </div>
</body></html>`;
}

// ── Image 5: Home screen ──────────────────────────────────────────────────────
function homePage() {
  const blue = '#0078D4';

  // Embed each card page as a scaled iframe using a data URI.
  function cardThumb(htmlFn, srcW, srcH) {
    // Card column width ≈ (1000 - 56 padding - 32 gaps) / 3 ≈ 304px
    const colW = 304;
    const scale = colW / srcW;
    const scaledH = Math.ceil(180 / scale);
    const encoded = encodeURIComponent(htmlFn());
    return `<div style="width:100%;height:180px;overflow:hidden;border-bottom:1px solid #EDEBE9;flex-shrink:0;background:#F8F8F8;">
      <iframe src="data:text/html;charset=utf-8,${encoded}"
        style="width:${srcW}px;height:${srcH}px;border:none;
          transform:scale(${scale.toFixed(4)});transform-origin:top left;pointer-events:none;"
        scrolling="no"></iframe>
    </div>`;
  }

  const cards = [
    {
      title: 'Permissions Explorer',
      icon: `<svg width="20" height="20" viewBox="0 0 24 24" fill="${blue}" style="flex-shrink:0;"><path d="M20 6h-8l-2-2H4c-1.1 0-2 .9-2 2v12c0 1.1.9 2 2 2h16c1.1 0 2-.9 2-2V8c0-1.1-.9-2-2-2zm-2 10H6v-2h12v2zm0-4H6v-2h12v2z"/></svg>`,
      desc: 'Browse any folder or file and instantly see who has access — with live, real-time permission lookups.',
      buttonLabel: 'Open Permissions Explorer',
      thumb: cardThumb(explorerPage, 800, 530),
    },
    {
      title: 'Permissions Report',
      icon: `<svg width="20" height="20" viewBox="0 0 24 24" fill="${blue}" style="flex-shrink:0;"><path d="M19 3H5c-1.1 0-2 .9-2 2v14c0 1.1.9 2 2 2h14c1.1 0 2-.9 2-2V5c0-1.1-.9-2-2-2zm-7 14H7v-2h5v2zm5-4H7v-2h10v2zm0-4H7V7h10v2z"/></svg>`,
      desc: 'Generate a color-coded Excel report of every unique permission assignment across your site.',
      buttonLabel: 'Run Permissions Report',
      thumb: cardThumb(reportPage, 800, 560),
    },
    {
      title: 'User Access',
      icon: `<svg width="20" height="20" viewBox="0 0 24 24" fill="${blue}" style="flex-shrink:0;"><path d="M12 12c2.21 0 4-1.79 4-4s-1.79-4-4-4-4 1.79-4 4 1.79 4 4 4zm0 2c-2.67 0-8 1.34-8 4v2h16v-2c0-2.66-5.33-4-8-4z"/></svg>`,
      desc: 'Look up any user to see every location they can access on a site, with their exact permission level.',
      buttonLabel: 'Check User Access',
      thumb: cardThumb(userAccessPage, 800, 500),
    },
  ];

  const cardHtml = cards.map(c => `
    <div style="display:flex;flex-direction:column;overflow:hidden;border-radius:4px;
      border:1px solid #EDEBE9;box-shadow:0 1px 4px rgba(0,0,0,0.08);background:#fff;">
      ${c.thumb}
      <div style="padding:14px 16px;display:flex;flex-direction:column;gap:8px;flex:1;">
        <div style="display:flex;align-items:center;gap:8px;">
          ${c.icon}
          <span style="font-size:15px;font-weight:600;color:#323130;font-family:${font};">${c.title}</span>
        </div>
        <div style="font-size:12px;color:#605E5C;font-family:${font};flex:1;line-height:1.5;">${c.desc}</div>
        <button style="width:100%;padding:7px 12px;background:${blue};color:#fff;border:none;
          border-radius:4px;font-size:13px;font-family:${font};font-weight:600;cursor:pointer;">
          ${c.buttonLabel}
        </button>
      </div>
    </div>`).join('');

  return `<!DOCTYPE html><html><head><meta charset="utf-8">
<style>* { margin:0; padding:0; box-sizing:border-box; }</style>
</head><body style="background:#fff;width:1000px;font-family:${font};color:#323130;">

  <!-- Blue banner -->
  <div style="background:${blue};display:flex;align-items:center;padding:0 16px;height:44px;gap:10px;">
    <svg width="20" height="20" viewBox="0 0 24 24" fill="white">
      <path d="M12 1L3 5v6c0 5.55 3.84 10.74 9 12 5.16-1.26 9-6.45 9-12V5l-9-4z"/>
    </svg>
    <span style="color:#fff;font-size:15px;font-weight:600;font-family:${font};">SharePoint Smart Permissions</span>
  </div>

  <!-- Main content -->
  <div style="padding:24px 28px;">

    <!-- Subtitle -->
    <div style="font-size:13px;color:#8A8886;margin-bottom:24px;font-family:${font};">
      Audit and understand SharePoint permissions — no PowerShell required.
    </div>

    <!-- 3-column card grid -->
    <div style="display:grid;grid-template-columns:repeat(3,1fr);gap:16px;margin-bottom:24px;">
      ${cardHtml}
    </div>

    <!-- Note at bottom -->
    <div style="padding:14px 16px;background:#F3F2F1;border-radius:4px;">
      <span style="font-size:12.5px;color:#605E5C;font-family:${font};line-height:1.5;">
        <strong>Note:</strong> This web part runs as the currently signed-in user. It can only see sites and
        items that user has permission to view. For a full tenant scan, use an account with appropriate
        read access across all sites.
      </span>
    </div>
  </div>

</body></html>`;
}

// ── Runner ────────────────────────────────────────────────────────────────────
async function main() {
  const browser = await puppeteer.launch({
    executablePath: CHROME,
    headless: 'new',
    args: ['--no-sandbox', '--disable-setuid-sandbox', '--font-render-hinting=none'],
  });

  // Card header images (used by HomeView as card thumbnails)
  const assetShots = [
    ['screenshot_report.png',      reportPage,      { width: 800, height: 560 }],
    ['screenshot_explorer.png',    explorerPage,    { width: 800, height: 530 }],
    ['screenshot_user_access.png', userAccessPage,  { width: 800, height: 500 }],
  ];

  for (const [filename, htmlFn, vp] of assetShots) {
    const pg = await browser.newPage();
    await pg.setViewport(vp);
    await pg.setContent(htmlFn(), { waitUntil: 'load' });
    await pg.screenshot({ path: path.join(OUT, filename), clip: { x: 0, y: 0, width: vp.width, height: vp.height } });
    await pg.close();
    console.log('✓', filename);
  }

  // README / docs screenshots
  const docsShots = [
    ['01_home.png', homePage, { width: 1000, height: 620 }],
  ];

  for (const [filename, htmlFn, vp] of docsShots) {
    const pg = await browser.newPage();
    await pg.setViewport(vp);
    await pg.setContent(htmlFn(), { waitUntil: 'networkidle0' });
    await pg.screenshot({ path: path.join(DOCS_OUT, filename), clip: { x: 0, y: 0, width: vp.width, height: vp.height } });
    await pg.close();
    console.log('✓', filename);
  }

  await browser.close();
  console.log(`\nAsset images → ${OUT}`);
  console.log(`Docs screenshots → ${DOCS_OUT}`);
}

main().catch(e => { console.error(e); process.exit(1); });
