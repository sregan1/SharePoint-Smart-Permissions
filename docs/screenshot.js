// docs/screenshot.js
// Generates documentation screenshots using puppeteer-core + local Chrome.
// Run: node docs/screenshot.js
'use strict';

const puppeteer = require('puppeteer-core');
const path = require('path');
const fs = require('fs');

const CHROME =
  process.env.CHROME_PATH ||
  (process.platform === 'darwin'
    ? '/Applications/Google Chrome.app/Contents/MacOS/Google Chrome'
    : process.platform === 'win32'
    ? 'C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe'
    : '/usr/bin/google-chrome');
const OUT = path.join(__dirname, 'screenshots');
if (!fs.existsSync(OUT)) fs.mkdirSync(OUT, { recursive: true });

// ─── shared design tokens ────────────────────────────────────────────────────
const T = {
  brand:        '#0078d4',
  brandDark:    '#005a9e',
  brandBg2:     '#eff6fc',
  brandFg1:     '#0078d4',
  white:        '#ffffff',
  bg:           '#f5f5f5',
  bgCard:       '#ffffff',
  border:       '#d1d1d1',
  borderStrong: '#ababab',
  text:         '#242424',
  textSubtle:   '#616161',
  textDisabled: '#bdbdbd',
  green:        '#107c10',
  greenBg:      '#dff6dd',
  orange:       '#ca5010',
  orangeBg:     '#fff4ce',
  red:          '#a4262c',
  redBg:        '#fde7e9',
  blue:         '#0078d4',
  blueBg:       '#eff6fc',
  radius:       '4px',
  radiusLg:     '8px',
  shadow:       '0 2px 4px rgba(0,0,0,0.1)',
  shadowMd:     '0 4px 12px rgba(0,0,0,0.14)',
};

// ─── base HTML shell ─────────────────────────────────────────────────────────
function shell(title, body) {
  return `<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8"/>
<title>${title}</title>
<style>
  * { box-sizing: border-box; margin: 0; padding: 0; }
  body {
    font-family: 'Segoe UI', -apple-system, BlinkMacSystemFont, sans-serif;
    font-size: 14px;
    color: ${T.text};
    background: ${T.bg};
    width: 960px;
  }
  .webpart {
    background: ${T.white};
    min-height: 560px;
    position: relative;
    overflow: hidden;
  }
  /* Banner */
  .banner {
    display: grid;
    grid-template-columns: auto 1fr auto;
    align-items: center;
    padding: 8px 12px 8px 16px;
    background: ${T.brand};
    gap: 16px;
  }
  .banner-brand { display: flex; align-items: center; gap: 8px; flex-shrink: 0; }
  .banner-brand span { color: white; font-weight: 600; white-space: nowrap; font-size: 14px; }
  .banner-center { display: flex; align-items: center; justify-content: center; gap: 8px; min-width: 0; }
  .banner-dot { width: 8px; height: 8px; border-radius: 50%; background: rgba(255,255,255,0.75); flex-shrink: 0; }
  .banner-url { color: white; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; font-size: 13px; }
  .banner-right { flex-shrink: 0; }
  .btn-transparent-white {
    background: transparent; border: none; color: white; cursor: pointer;
    font-size: 12px; padding: 3px 8px; border-radius: ${T.radius};
    font-family: inherit;
  }
  .btn-transparent-white:hover { background: rgba(255,255,255,0.15); }
  .icon-btn {
    background: transparent; border: none; cursor: pointer;
    padding: 4px; border-radius: ${T.radius}; display: flex; align-items: center;
  }
  /* Content area */
  .content { padding: 24px; }
  /* Back button row */
  .back-row { margin-bottom: 20px; }
  .btn-back {
    background: transparent; border: 1px solid ${T.border}; color: ${T.text};
    padding: 5px 12px; border-radius: ${T.radius}; cursor: pointer; font-family: inherit; font-size: 13px;
    display: inline-flex; align-items: center; gap: 6px;
  }
  /* Buttons */
  .btn-primary {
    background: ${T.brand}; color: white; border: none;
    padding: 6px 16px; border-radius: ${T.radius}; cursor: pointer;
    font-family: inherit; font-size: 14px; font-weight: 500;
  }
  .btn-secondary {
    background: white; color: ${T.text}; border: 1px solid ${T.border};
    padding: 6px 16px; border-radius: ${T.radius}; cursor: pointer;
    font-family: inherit; font-size: 14px;
  }
  .btn-disabled {
    background: ${T.bg}; color: ${T.textDisabled}; border: 1px solid ${T.border};
    padding: 6px 16px; border-radius: ${T.radius}; cursor: default;
    font-family: inherit; font-size: 14px;
  }
  /* Labels & headings */
  .title { font-size: 20px; font-weight: 600; color: ${T.text}; margin-bottom: 4px; }
  .subtitle { font-size: 13px; color: ${T.textSubtle}; margin-bottom: 20px; }
  .section-label { font-size: 12px; font-weight: 600; color: ${T.textSubtle}; text-transform: uppercase; letter-spacing: 0.5px; margin-bottom: 8px; }
  /* Form row */
  .form-row { margin-bottom: 16px; }
  .form-label { font-size: 13px; font-weight: 500; margin-bottom: 6px; color: ${T.text}; display: block; }
  /* Dropdown */
  .dropdown {
    border: 1px solid ${T.border}; border-radius: ${T.radius};
    padding: 6px 10px; background: white; font-family: inherit; font-size: 13px;
    width: 100%; max-width: 320px; color: ${T.text};
    display: flex; align-items: center; justify-content: space-between; cursor: pointer;
  }
  /* Progress */
  .progress-track { height: 4px; background: ${T.border}; border-radius: 2px; overflow: hidden; }
  .progress-fill { height: 100%; background: ${T.brand}; border-radius: 2px; }
  /* Badge */
  .badge {
    display: inline-flex; align-items: center; padding: 2px 8px;
    border-radius: 10px; font-size: 11px; font-weight: 600; white-space: nowrap;
  }
  .badge-unique { background: ${T.orangeBg}; color: ${T.orange}; }
  .badge-inherited { background: ${T.bg}; color: ${T.textSubtle}; border: 1px solid ${T.border}; }
  .badge-red { background: ${T.redBg}; color: ${T.red}; }
  .badge-orange { background: ${T.orangeBg}; color: ${T.orange}; }
  .badge-green { background: ${T.greenBg}; color: ${T.green}; }
  .badge-blue { background: ${T.blueBg}; color: ${T.blue}; }
  /* Table */
  table { width: 100%; border-collapse: collapse; font-size: 13px; }
  th { text-align: left; padding: 8px 12px; font-size: 12px; font-weight: 600; color: ${T.textSubtle}; border-bottom: 2px solid ${T.border}; }
  td { padding: 8px 12px; border-bottom: 1px solid ${T.bg}; }
  tr:hover td { background: ${T.bg}; }
  /* Tree */
  .tree-item {
    display: flex; align-items: center; gap: 6px;
    padding: 5px 8px; border-radius: ${T.radius}; cursor: pointer; font-size: 13px;
  }
  .tree-item:hover { background: ${T.bg}; }
  .tree-item.selected { background: ${T.blueBg}; }
  .tree-indent { padding-left: 20px; }
  .tree-indent2 { padding-left: 40px; }
  /* Two-col layout */
  .two-col { display: grid; grid-template-columns: 280px 1fr; gap: 16px; }
  .panel { border: 1px solid ${T.border}; border-radius: ${T.radiusLg}; overflow: hidden; }
  .panel-header { padding: 10px 14px; border-bottom: 1px solid ${T.border}; font-weight: 600; font-size: 13px; background: ${T.bg}; }
  .panel-body { padding: 0; }
  /* Cards */
  .card {
    background: white; border: 1px solid ${T.border}; border-radius: ${T.radiusLg};
    padding: 20px; box-shadow: ${T.shadow};
  }
  /* Misc */
  .flex { display: flex; }
  .flex-col { display: flex; flex-direction: column; }
  .items-center { align-items: center; }
  .gap-8 { gap: 8px; }
  .gap-12 { gap: 12px; }
  .gap-16 { gap: 16px; }
  .text-subtle { color: ${T.textSubtle}; font-size: 13px; }
  .text-small { font-size: 12px; }
  .text-semibold { font-weight: 600; }
  .mb-4 { margin-bottom: 4px; }
  .mb-8 { margin-bottom: 8px; }
  .mb-12 { margin-bottom: 12px; }
  .mb-16 { margin-bottom: 16px; }
  .mb-20 { margin-bottom: 20px; }
  .mr-8 { margin-right: 8px; }
  .ml-4 { margin-left: 4px; }
</style>
</head>
<body>
<div class="webpart">
${body}
</div>
</body>
</html>`;
}

// ─── SVG icons (inline) ──────────────────────────────────────────────────────
const icons = {
  shield: `<svg width="20" height="20" viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg"><path d="M12 2L4 6v6c0 5.55 3.84 10.74 8 12 4.16-1.26 8-6.45 8-12V6l-8-4z" fill="white" opacity="0.9"/><path d="M10 14.5l-2-2-1 1 3 3 5-5-1-1-4 4z" fill="${T.brand}"/></svg>`,
  settings: `<svg width="18" height="18" viewBox="0 0 24 24" fill="white"><path d="M19.14 12.94c.04-.3.06-.61.06-.94 0-.32-.02-.64-.07-.94l2.03-1.58c.18-.14.23-.41.12-.61l-1.92-3.32c-.12-.22-.37-.29-.59-.22l-2.39.96c-.5-.38-1.03-.7-1.62-.94l-.36-2.54c-.04-.24-.24-.41-.48-.41h-3.84c-.24 0-.43.17-.47.41l-.36 2.54c-.59.24-1.13.57-1.62.94l-2.39-.96c-.22-.08-.47 0-.59.22L2.74 8.87c-.12.21-.08.47.12.61l2.03 1.58c-.05.3-.09.63-.09.94s.02.64.07.94l-2.03 1.58c-.18.14-.23.41-.12.61l1.92 3.32c.12.22.37.29.59.22l2.39-.96c.5.38 1.03.7 1.62.94l.36 2.54c.05.24.24.41.48.41h3.84c.24 0 .44-.17.47-.41l.36-2.54c.59-.24 1.13-.56 1.62-.94l2.39.96c.22.08.47 0 .59-.22l1.92-3.32c.12-.22.07-.47-.12-.61l-2.01-1.58zM12 15.6c-1.98 0-3.6-1.62-3.6-3.6s1.62-3.6 3.6-3.6 3.6 1.62 3.6 3.6-1.62 3.6-3.6 3.6z"/></svg>`,
  folder: `<svg width="16" height="16" viewBox="0 0 24 24" fill="none"><path d="M10 4H4c-1.1 0-2 .9-2 2v12c0 1.1.9 2 2 2h16c1.1 0 2-.9 2-2V8c0-1.1-.9-2-2-2h-8l-2-2z" fill="#f0a500"/></svg>`,
  folderBlue: `<svg width="16" height="16" viewBox="0 0 24 24" fill="none"><path d="M10 4H4c-1.1 0-2 .9-2 2v12c0 1.1.9 2 2 2h16c1.1 0 2-.9 2-2V8c0-1.1-.9-2-2-2h-8l-2-2z" fill="${T.brand}"/></svg>`,
  file: `<svg width="14" height="16" viewBox="0 0 24 24" fill="none"><path d="M14 2H6c-1.1 0-2 .9-2 2v16c0 1.1.9 2 2 2h12c1.1 0 2-.9 2-2V8l-6-6zm2 16H8v-2h8v2zm0-4H8v-2h8v2zm-3-5V3.5L18.5 9H13z" fill="#0078d4"/></svg>`,
  chevronDown: `<svg width="12" height="12" viewBox="0 0 24 24" fill="${T.textSubtle}"><path d="M7 10l5 5 5-5H7z"/></svg>`,
  chevronRight: `<svg width="12" height="12" viewBox="0 0 24 24" fill="${T.textSubtle}"><path d="M10 17l5-5-5-5v10z"/></svg>`,
  arrowDown: `<svg width="11" height="11" viewBox="0 0 24 24" fill="${T.orange}"><path d="M20 12l-1.41-1.41L13 16.17V4h-2v12.17l-5.58-5.59L4 12l8 8 8-8z"/></svg>`,
  check: `<svg width="14" height="14" viewBox="0 0 24 24" fill="${T.green}"><path d="M9 16.17L4.83 12l-1.42 1.41L9 19 21 7l-1.41-1.41L9 16.17z"/></svg>`,
  person: `<svg width="14" height="14" viewBox="0 0 24 24" fill="${T.textSubtle}"><path d="M12 12c2.21 0 4-1.79 4-4s-1.79-4-4-4-4 1.79-4 4 1.79 4 4 4zm0 2c-2.67 0-8 1.34-8 4v2h16v-2c0-2.66-5.33-4-8-4z"/></svg>`,
  people: `<svg width="14" height="14" viewBox="0 0 24 24" fill="${T.textSubtle}"><path d="M16 11c1.66 0 2.99-1.34 2.99-3S17.66 5 16 5c-1.66 0-3 1.34-3 3s1.34 3 3 3zm-8 0c1.66 0 2.99-1.34 2.99-3S9.66 5 8 5C6.34 5 5 6.34 5 8s1.34 3 3 3zm0 2c-2.33 0-7 1.17-7 3.5V19h14v-2.5c0-2.33-4.67-3.5-7-3.5zm8 0c-.29 0-.62.02-.97.05 1.16.84 1.97 1.97 1.97 3.45V19h6v-2.5c0-2.33-4.67-3.5-7-3.5z"/></svg>`,
  back: `<svg width="14" height="14" viewBox="0 0 24 24" fill="${T.text}"><path d="M20 11H7.83l5.59-5.59L12 4l-8 8 8 8 1.41-1.41L7.83 13H20v-2z"/></svg>`,
  info: `<svg width="14" height="14" viewBox="0 0 24 24" fill="${T.brand}"><path d="M12 2C6.48 2 2 6.48 2 12s4.48 10 10 10 10-4.48 10-10S17.52 2 12 2zm1 15h-2v-6h2v6zm0-8h-2V7h2v2z"/></svg>`,
  library: `<svg width="16" height="16" viewBox="0 0 24 24" fill="${T.brand}"><path d="M4 6H2v14c0 1.1.9 2 2 2h14v-2H4V6zm16-4H8c-1.1 0-2 .9-2 2v12c0 1.1.9 2 2 2h12c1.1 0 2-.9 2-2V4c0-1.1-.9-2-2-2zm-1 9H9V9h10v2zm-4 4H9v-2h6v2zm4-8H9V5h10v2z"/></svg>`,
  report: `<svg width="20" height="20" viewBox="0 0 24 24" fill="${T.brand}"><path d="M19 3H5c-1.1 0-2 .9-2 2v14c0 1.1.9 2 2 2h14c1.1 0 2-.9 2-2V5c0-1.1-.9-2-2-2zm-5 14H7v-2h7v2zm3-4H7v-2h10v2zm0-4H7V7h10v2z"/></svg>`,
  explore: `<svg width="20" height="20" viewBox="0 0 24 24" fill="${T.brand}"><path d="M10 4H4c-1.1 0-2 .9-2 2v12c0 1.1.9 2 2 2h16c1.1 0 2-.9 2-2V8c0-1.1-.9-2-2-2h-8l-2-2zM5 18l3-6 2.06 2.54L13 10l5 8H5z"/></svg>`,
  user: `<svg width="20" height="20" viewBox="0 0 24 24" fill="${T.brand}"><path d="M12 12c2.21 0 4-1.79 4-4s-1.79-4-4-4-4 1.79-4 4 1.79 4 4 4zm0 2c-2.67 0-8 1.34-8 4v2h16v-2c0-2.66-5.33-4-8-4z"/></svg>`,
  excel: `<svg width="16" height="16" viewBox="0 0 24 24" fill="none"><rect width="24" height="24" rx="3" fill="#1D6F42"/><path d="M7 6l3.5 5.5L7 17h2.5l2-3.2 2 3.2H16l-3.5-5.5L16 6h-2.5l-2 3.2L9.5 6H7z" fill="white"/></svg>`,
};

// ─── banner component ─────────────────────────────────────────────────────────
function banner(siteUrl = 'https://contoso.sharepoint.com/sites/Marketing') {
  return `<div class="banner">
  <div class="banner-brand">
    ${icons.shield}
    <span>SharePoint Smart Permissions</span>
  </div>
  <div class="banner-center">
    <div class="banner-dot"></div>
    <span class="banner-url">${siteUrl}</span>
    <button class="btn-transparent-white" style="font-size:11px">Change URL</button>
  </div>
  <div class="banner-right">
    <button class="icon-btn">${icons.settings}</button>
  </div>
</div>`;
}

// ─── SCREEN 1: Home ───────────────────────────────────────────────────────────
function homeHTML() {
  const cards = [
    {
      icon: icons.report,
      title: 'Permissions Report',
      desc: 'Generate a comprehensive, exportable audit report of every unique permission assignment across libraries, folders, and files.',
      features: [
        'Scan site, libraries, folders, or files',
        'Configurable folder depth (1–10 levels)',
        'Tenant-wide scan across all site collections',
        'Color-coded Excel export with auto-filter',
      ],
      btn: 'Run Permissions Report',
    },
    {
      icon: icons.explore,
      title: 'Permissions Explorer',
      desc: 'Browse a document library interactively and inspect live permissions on any folder or file in real time.',
      features: [
        'Expandable folder and file tree',
        'Unique permission badges and indicators',
        'Expand SharePoint group members inline',
        'Show parent permissions for inherited items',
      ],
      btn: 'Open Permissions Explorer',
    },
    {
      icon: icons.user,
      title: 'User Access',
      desc: "Select any user and see every location they can access, with their exact permission level at each location.",
      features: [
        'Full site access detection for owners',
        'Exact permission level at each location',
        'Live elapsed timer during scan',
        'Cancel at any time for partial results',
      ],
      btn: 'Check User Access',
    },
  ];

  return shell('Home', `
  <div style="position:absolute;top:8px;right:8px;z-index:10">
    <button class="icon-btn" style="color:${T.textSubtle}">${icons.settings}</button>
  </div>
  <div class="content" style="padding: 32px 40px;">
    <div style="text-align:center;margin-bottom:32px">
      <div style="display:inline-flex;align-items:center;gap:10px;margin-bottom:12px">
        ${icons.shield.replace('white','#0078d4').replace('fill="${T.brand}"','fill="white"')}
        <span style="font-size:22px;font-weight:700;color:${T.text}">SharePoint Smart Permissions</span>
      </div>
      <p style="color:${T.textSubtle};font-size:14px;max-width:600px;margin:0 auto">
        Audit and understand SharePoint permissions in real time, directly from your browser — no PowerShell or admin tools required.
      </p>
    </div>
    <div style="display:grid;grid-template-columns:repeat(3,1fr);gap:20px">
      ${cards.map(c => `
      <div class="card" style="display:flex;flex-direction:column;gap:16px">
        <div style="display:flex;align-items:center;gap:10px">
          ${c.icon}
          <span style="font-size:16px;font-weight:600">${c.title}</span>
        </div>
        <p style="color:${T.textSubtle};font-size:13px;line-height:1.5">${c.desc}</p>
        <div style="display:flex;flex-direction:column;gap:6px">
          ${c.features.map(f => `<div style="display:flex;align-items:flex-start;gap:6px;font-size:13px">
            <span style="flex-shrink:0;margin-top:1px">${icons.check}</span>
            <span style="color:${T.text}">${f}</span>
          </div>`).join('')}
        </div>
        <button class="btn-primary" style="margin-top:auto;width:100%;padding:8px 0">${c.btn}</button>
      </div>`).join('')}
    </div>
  </div>`);
}

// ─── SCREEN 2: Permissions Report — Config ────────────────────────────────────
function reportConfigHTML() {
  const options = [
    { val: 'site',    label: 'Site only',        desc: 'Top-level site permissions', selected: false },
    { val: 'libs',    label: 'Libraries',         desc: 'All document libraries',     selected: false },
    { val: 'folders', label: 'Folders',           desc: 'Libraries + folder tree',    selected: true  },
    { val: 'files',   label: 'Files & Folders',   desc: 'Everything including files', selected: false },
  ];
  return shell('Permissions Report - Config', `
  ${banner()}
  <div class="content">
    <div class="back-row">
      <button class="btn-back">${icons.back} Back</button>
    </div>
    <div class="title">Permissions Report</div>
    <div class="subtitle">Scan this site and export a colour-coded Excel workbook of every unique permission assignment.</div>

    <div class="form-row">
      <span class="form-label">Scan depth</span>
      <div style="display:grid;grid-template-columns:repeat(4,1fr);gap:10px;max-width:720px">
        ${options.map(o => `
        <div style="border:${o.selected ? `2px solid ${T.brand}` : `1px solid ${T.border}`};border-radius:${T.radius};padding:10px 14px;cursor:pointer;background:${o.selected ? T.brandBg2 : 'white'}">
          <div style="font-weight:600;font-size:13px;margin-bottom:3px;color:${o.selected ? T.brand : T.text}">${o.label}</div>
          <div style="font-size:12px;color:${T.textSubtle}">${o.desc}</div>
        </div>`).join('')}
      </div>
    </div>

    <div class="form-row" style="margin-top:16px">
      <span class="form-label">Folder depth limit</span>
      <div style="display:flex;align-items:center;gap:10px">
        <button class="btn-secondary" style="width:32px;padding:4px 0;text-align:center">−</button>
        <span style="width:30px;text-align:center;font-weight:600">3</span>
        <button class="btn-secondary" style="width:32px;padding:4px 0;text-align:center">+</button>
        <span class="text-subtle text-small">levels deep</span>
      </div>
    </div>

    <div class="form-row" style="max-width:420px;background:${T.bg};border-radius:${T.radius};padding:12px 14px;margin-top:4px">
      <div style="display:flex;align-items:center;gap:8px">
        <input type="checkbox" id="cb" style="width:16px;height:16px;accent-color:${T.brand}"/>
        <label for="cb" style="font-size:13px;cursor:pointer">Scan all site collections in this tenant</label>
        <span style="margin-left:4px">${icons.info}</span>
      </div>
      <p style="font-size:12px;color:${T.textSubtle};margin-top:6px;margin-left:24px">
        Only available on the root site. Requires read access across all site collections.
      </p>
    </div>

    <div style="margin-top:28px">
      <button class="btn-primary" style="padding:8px 24px;font-size:14px">Run Report</button>
    </div>
  </div>`);
}

// ─── SCREEN 3: Permissions Report — Running ────────────────────────────────────
function reportRunningHTML() {
  return shell('Permissions Report - Running', `
  ${banner()}
  <div class="content">
    <div class="back-row">
      <button class="btn-back">${icons.back} Back</button>
    </div>
    <div class="title">Permissions Report</div>
    <div class="subtitle">Scan in progress — do not navigate away.</div>

    <div style="max-width:560px;margin-top:8px">
      <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:8px">
        <span style="font-size:13px;color:${T.text}">Scanning libraries and folders…</span>
        <button class="btn-secondary" style="font-size:12px;padding:4px 12px">Cancel</button>
      </div>
      <div class="progress-track" style="margin-bottom:6px">
        <div class="progress-fill" style="width:58%"></div>
      </div>
      <div style="display:flex;justify-content:space-between;font-size:12px;color:${T.textSubtle}">
        <span>Currently scanning: Campaign Documents / Q4 Launch / Assets</span>
        <span>58%</span>
      </div>
    </div>

    <div style="margin-top:24px;display:grid;grid-template-columns:repeat(4,1fr);gap:12px;max-width:560px">
      ${[
        { label: 'Libraries scanned', val: '7' },
        { label: 'Folders scanned',   val: '43' },
        { label: 'Items with unique permissions', val: '12' },
        { label: 'Elapsed time', val: '0:32' },
      ].map(s => `<div style="background:${T.bg};border-radius:${T.radius};padding:12px;text-align:center">
        <div style="font-size:22px;font-weight:700;color:${T.brand}">${s.val}</div>
        <div style="font-size:11px;color:${T.textSubtle};margin-top:4px">${s.label}</div>
      </div>`).join('')}
    </div>

    <div style="margin-top:24px">
      <button class="btn-disabled">Export to Excel</button>
      <span style="font-size:12px;color:${T.textSubtle};margin-left:10px">Export available when scan is complete</span>
    </div>
  </div>`);
}

// ─── SCREEN 4: Permissions Report — Complete ────────────────────────────────────
function reportCompleteHTML() {
  const rows = [
    { type: 'Site',    name: 'Marketing',             path: '/sites/Marketing',                        unique: true,  user: 'Marketing Owners',  role: 'Full Control' },
    { type: 'Library', name: 'Campaign Documents',     path: '/sites/Marketing/Campaign Documents',     unique: true,  user: 'External Reviewers',role: 'Read'         },
    { type: 'Folder',  name: 'Q4 Launch',              path: '/sites/Marketing/Campaign Documents/Q4 Launch', unique: true, user: 'Carol White', role: 'Edit' },
    { type: 'Folder',  name: 'Assets',                 path: '/sites/Marketing/Campaign Documents/Q4 Launch/Assets', unique: false, user: '(inherited)', role: '' },
    { type: 'File',    name: 'Budget_Confidential.xlsx', path: '/sites/Marketing/Campaign Documents/Q4 Launch/Budget_Confidential.xlsx', unique: true, user: 'Alice Chen', role: 'Full Control' },
    { type: 'Library', name: 'Brand Guidelines',       path: '/sites/Marketing/Brand Guidelines',       unique: false, user: '(inherited)',        role: ''             },
    { type: 'Library', name: 'Marketing Assets',       path: '/sites/Marketing/Marketing Assets',       unique: true,  user: 'Design Contractors', role: 'Read' },
  ];

  const roleColor = r => r === 'Full Control' ? 'badge-red' : r === 'Edit' ? 'badge-orange' : r === 'Read' ? 'badge-green' : '';
  const typeColor = t => t === 'Site' ? 'badge-red' : t === 'Library' ? 'badge-blue' : t === 'Folder' ? 'badge-orange' : 'badge-green';

  return shell('Permissions Report - Complete', `
  ${banner()}
  <div class="content">
    <div class="back-row">
      <button class="btn-back">${icons.back} Back</button>
    </div>
    <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:20px">
      <div>
        <div class="title">Permissions Report</div>
        <div class="subtitle" style="margin-bottom:0">Scan complete — 75 objects scanned, 5 with unique permissions</div>
      </div>
      <button class="btn-primary" style="display:flex;align-items:center;gap:8px;padding:8px 18px">
        ${icons.excel} Export to Excel
      </button>
    </div>

    <div class="panel">
      <div class="panel-header">Results</div>
      <table>
        <thead>
          <tr>
            <th>Type</th><th>Name</th><th>Path</th><th>Unique Permissions</th><th>User / Group</th><th>Permission Level</th>
          </tr>
        </thead>
        <tbody>
          ${rows.map(r => `<tr style="${r.unique ? `background:#fffbf5` : ''}">
            <td><span class="badge ${typeColor(r.type)}">${r.type}</span></td>
            <td style="font-weight:${r.unique ? '500' : 'normal'}">${r.name}</td>
            <td style="color:${T.textSubtle};font-size:12px">${r.path}</td>
            <td>${r.unique ? `<span class="badge badge-unique">Yes</span>` : `<span style="color:${T.textSubtle};font-size:12px">No</span>`}</td>
            <td style="font-size:13px">${r.user}</td>
            <td>${r.role ? `<span class="badge ${roleColor(r.role)}">${r.role}</span>` : ''}</td>
          </tr>`).join('')}
        </tbody>
      </table>
    </div>
  </div>`);
}

// ─── SCREEN 5: Permissions Explorer ───────────────────────────────────────────
function explorerHTML() {
  return shell('Permissions Explorer', `
  ${banner()}
  <div class="content">
    <div class="back-row">
      <button class="btn-back">${icons.back} Back</button>
    </div>
    <div class="title">Permissions Explorer</div>
    <div class="subtitle">Select a library and click any folder or file to inspect its permissions.</div>

    <div class="form-row" style="margin-bottom:16px">
      <span class="form-label">Library</span>
      <div class="dropdown" style="max-width:280px">
        <span>Campaign Documents</span>
        ${icons.chevronDown}
      </div>
    </div>

    <div class="two-col">
      <!-- Tree panel -->
      <div class="panel">
        <div class="panel-header">Folders &amp; Files</div>
        <div class="panel-body" style="padding:8px">
          <div class="tree-item">
            ${icons.chevronDown}${icons.folder}
            <span>Campaign Documents</span>
          </div>
          <div class="tree-indent">
            <div class="tree-item">
              ${icons.chevronDown}${icons.folder}
              <span>Q4 Launch</span>
              <span class="badge badge-unique" style="margin-left:auto;font-size:10px">Unique</span>
            </div>
            <div class="tree-indent">
              <div class="tree-item selected">
                ${icons.chevronRight}${icons.folder}
                <span>Assets</span>
                <span style="margin-left:auto" title="Contains items with unique permissions">${icons.arrowDown}</span>
              </div>
              <div class="tree-item">
                ${icons.file}
                <span>Campaign_Brief.docx</span>
              </div>
              <div class="tree-item">
                ${icons.file}
                <span>Budget_Confidential.xlsx</span>
                <span class="badge badge-unique" style="margin-left:auto;font-size:10px">Unique</span>
              </div>
            </div>
            <div class="tree-item">
              ${icons.chevronRight}${icons.folder}
              <span>Q1 Planning</span>
            </div>
            <div class="tree-item">
              ${icons.chevronRight}${icons.folder}
              <span>Archive</span>
            </div>
          </div>
        </div>
      </div>

      <!-- Permission panel -->
      <div class="panel">
        <div class="panel-header" style="display:flex;align-items:center;gap:8px">
          ${icons.folder}
          <span>Assets</span>
          <span class="badge badge-inherited" style="margin-left:8px">Inherited permissions</span>
        </div>
        <div style="padding:12px 14px">
          <div style="display:flex;gap:16px;margin-bottom:16px;flex-wrap:wrap">
            <label style="display:flex;align-items:center;gap:6px;font-size:13px;cursor:pointer">
              <input type="checkbox" style="accent-color:${T.brand}"/>
              Expand SharePoint group members
            </label>
            <label style="display:flex;align-items:center;gap:6px;font-size:13px;cursor:pointer">
              <input type="checkbox" checked style="accent-color:${T.brand}"/>
              Show parent permissions (Q4 Launch)
            </label>
          </div>
          <div style="font-size:12px;color:${T.textSubtle};margin-bottom:10px">
            Showing permissions from nearest ancestor with unique permissions: <strong>Q4 Launch</strong>
          </div>
          <table>
            <thead>
              <tr><th>User / Group</th><th>Type</th><th>Permission Level</th></tr>
            </thead>
            <tbody>
              <tr>
                <td style="display:flex;align-items:center;gap:6px">${icons.people}<span>Marketing Owners</span></td>
                <td><span class="badge" style="background:${T.bg};color:${T.textSubtle};font-size:11px">SP Group</span></td>
                <td><span class="badge badge-red">Full Control</span></td>
              </tr>
              <tr>
                <td style="display:flex;align-items:center;gap:6px">${icons.people}<span>Marketing Members</span></td>
                <td><span class="badge" style="background:${T.bg};color:${T.textSubtle};font-size:11px">SP Group</span></td>
                <td><span class="badge badge-orange">Edit</span></td>
              </tr>
              <tr>
                <td style="display:flex;align-items:center;gap:6px">${icons.person}<span>Carol White</span></td>
                <td><span class="badge" style="background:${T.bg};color:${T.textSubtle};font-size:11px">User</span></td>
                <td><span class="badge badge-orange">Edit</span></td>
              </tr>
              <tr>
                <td style="display:flex;align-items:center;gap:6px">${icons.people}<span>Marketing Visitors</span></td>
                <td><span class="badge" style="background:${T.bg};color:${T.textSubtle};font-size:11px">SP Group</span></td>
                <td><span class="badge badge-green">Read</span></td>
              </tr>
              <tr>
                <td style="display:flex;align-items:center;gap:6px">${icons.people}<span>External Reviewers</span></td>
                <td><span class="badge" style="background:${T.bg};color:${T.textSubtle};font-size:11px">SP Group</span></td>
                <td><span class="badge badge-green">Read</span></td>
              </tr>
            </tbody>
          </table>
        </div>
      </div>
    </div>
  </div>`);
}

// ─── SCREEN 6: User Access — Scanning ─────────────────────────────────────────
function userAccessRunningHTML() {
  return shell('User Access - Scanning', `
  ${banner()}
  <div class="content">
    <div class="back-row">
      <button class="btn-back">${icons.back} Back</button>
    </div>
    <div class="title">User Access</div>
    <div class="subtitle">See every location a specific user can access, with their exact permission level.</div>

    <div class="form-row" style="margin-bottom:20px">
      <span class="form-label">Select a user</span>
      <div class="dropdown" style="max-width:320px">
        <div style="display:flex;align-items:center;gap:8px">${icons.person}<span>Alice Chen (alice.chen@contoso.com)</span></div>
        ${icons.chevronDown}
      </div>
    </div>

    <div style="max-width:540px;background:white;border:1px solid ${T.border};border-radius:${T.radiusLg};padding:20px">
      <div style="display:flex;align-items:center;gap:12px;margin-bottom:16px">
        <div style="width:36px;height:36px;border-radius:50%;background:${T.brandBg2};display:flex;align-items:center;justify-content:center">
          ${icons.person.replace('${T.textSubtle}','#0078d4')}
        </div>
        <div>
          <div style="font-weight:600">Alice Chen</div>
          <div style="font-size:12px;color:${T.textSubtle}">alice.chen@contoso.com</div>
        </div>
      </div>

      <div style="margin-bottom:12px">
        <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:6px">
          <span style="font-size:13px">Scanning Campaign Documents…</span>
          <span style="font-size:12px;color:${T.textSubtle};font-variant-numeric:tabular-nums">0:47</span>
        </div>
        <div class="progress-track">
          <div class="progress-fill" style="width:38%;animation:none"></div>
        </div>
        <div style="font-size:12px;color:${T.textSubtle};margin-top:4px">Library 3 of 8</div>
      </div>

      <div style="display:flex;align-items:center;gap:12px;margin-top:16px">
        <button class="btn-secondary" style="font-size:13px;padding:5px 14px">Cancel</button>
        <span style="font-size:12px;color:${T.textSubtle}">This scan may take several minutes depending on the size of the site.</span>
      </div>
    </div>
  </div>`);
}

// ─── SCREEN 7: User Access — Complete ─────────────────────────────────────────
function userAccessCompleteHTML() {
  const rows = [
    { type: 'Site',    name: 'Marketing',              path: '/sites/Marketing',                                   role: 'Read'         },
    { type: 'Library', name: 'Campaign Documents',      path: '/sites/Marketing/Campaign Documents',                role: 'Edit'         },
    { type: 'Folder',  name: 'Q4 Launch',               path: '/sites/Marketing/Campaign Documents/Q4 Launch',     role: 'Edit'         },
    { type: 'File',    name: 'Budget_Confidential.xlsx', path: '/sites/Marketing/Campaign Documents/Q4 Launch/Budget_Confidential.xlsx', role: 'Full Control' },
    { type: 'Library', name: 'Marketing Assets',        path: '/sites/Marketing/Marketing Assets',                  role: 'Read'         },
    { type: 'Folder',  name: 'Brand Photos',            path: '/sites/Marketing/Marketing Assets/Brand Photos',     role: 'Edit'         },
  ];
  const roleColor = r => r === 'Full Control' ? 'badge-red' : r === 'Edit' ? 'badge-orange' : 'badge-green';
  const typeColor = t => t === 'Site' ? 'badge-red' : t === 'Library' ? 'badge-blue' : t === 'Folder' ? 'badge-orange' : 'badge-green';

  return shell('User Access - Complete', `
  ${banner()}
  <div class="content">
    <div class="back-row">
      <button class="btn-back">${icons.back} Back</button>
    </div>
    <div style="display:flex;align-items:center;gap:16px;margin-bottom:20px">
      <div>
        <div class="title">User Access</div>
        <div class="subtitle" style="margin-bottom:0">Access report for <strong>Alice Chen</strong> — 6 locations found</div>
      </div>
    </div>

    <div class="form-row" style="margin-bottom:16px">
      <div class="dropdown" style="max-width:320px">
        <div style="display:flex;align-items:center;gap:8px">${icons.person}<span>Alice Chen (alice.chen@contoso.com)</span></div>
        ${icons.chevronDown}
      </div>
    </div>

    <div class="panel">
      <div class="panel-header">Accessible locations (6)</div>
      <table>
        <thead>
          <tr><th>Type</th><th>Name</th><th>Path</th><th>Permission Level</th></tr>
        </thead>
        <tbody>
          ${rows.map(r => `<tr>
            <td><span class="badge ${typeColor(r.type)}">${r.type}</span></td>
            <td>${r.name}</td>
            <td style="color:${T.textSubtle};font-size:12px">${r.path}</td>
            <td><span class="badge ${roleColor(r.role)}">${r.role}</span></td>
          </tr>`).join('')}
        </tbody>
      </table>
    </div>
  </div>`);
}

// ─── SCREEN 8: User Access — Full Site Access ─────────────────────────────────
function userAccessFullHTML() {
  return shell('User Access - Full Site Access', `
  ${banner()}
  <div class="content">
    <div class="back-row">
      <button class="btn-back">${icons.back} Back</button>
    </div>
    <div class="title">User Access</div>
    <div class="subtitle">See every location a specific user can access, with their exact permission level.</div>

    <div class="form-row" style="margin-bottom:20px">
      <span class="form-label">Select a user</span>
      <div class="dropdown" style="max-width:320px">
        <div style="display:flex;align-items:center;gap:8px">${icons.person}<span>Bob Martinez (bob.martinez@contoso.com)</span></div>
        ${icons.chevronDown}
      </div>
    </div>

    <div style="max-width:500px;background:${T.redBg};border:1px solid #f1707b;border-radius:${T.radiusLg};padding:24px;display:flex;flex-direction:column;align-items:center;text-align:center;gap:12px">
      <svg width="40" height="40" viewBox="0 0 24 24" fill="none"><path d="M12 2L4 6v6c0 5.55 3.84 10.74 8 12 4.16-1.26 8-6.45 8-12V6l-8-4z" fill="${T.red}" opacity="0.9"/></svg>
      <div style="font-size:17px;font-weight:700;color:${T.red}">Full Site Access</div>
      <div style="font-size:14px;color:${T.text}">
        <strong>Bob Martinez</strong> has <strong>Full Control</strong> at the site level via the <strong>Marketing Owners</strong> group.
      </div>
      <div style="font-size:13px;color:${T.textSubtle}">
        This user can access all content on the site. Individual item listing is not shown for owner-level accounts.
      </div>
    </div>
  </div>`);
}

// ─── SCREEN 9: Settings popover + URL edit mode ───────────────────────────────
function settingsHTML() {
  return shell('Settings', `
  ${banner()}
  <div class="content">
    <div class="back-row">
      <button class="btn-back">${icons.back} Back</button>
    </div>
    <div class="title">Global Settings</div>
    <div class="subtitle">Access settings from the gear icon (⚙) in the top-right corner of the banner on any screen.</div>

    <div style="display:flex;gap:40px;margin-top:8px;flex-wrap:wrap">
      <!-- Settings panel -->
      <div>
        <div class="section-label" style="margin-bottom:12px">Settings panel</div>
        <div style="background:white;border:1px solid ${T.border};border-radius:${T.radiusLg};box-shadow:${T.shadowMd};padding:16px;min-width:300px">
          <div style="font-weight:600;font-size:14px;margin-bottom:14px">Settings</div>
          <div style="display:flex;align-items:center;gap:8px">
            <input type="checkbox" checked style="width:16px;height:16px;accent-color:${T.brand}"/>
            <label style="font-size:13px">Include system and hidden libraries</label>
            <span style="margin-left:2px">${icons.info}</span>
          </div>
          <p style="font-size:12px;color:${T.textSubtle};margin-top:8px;margin-left:24px;line-height:1.5">
            When checked, includes Style Library, Form Templates,<br/>Site Assets, and other hidden libraries.
          </p>
        </div>
      </div>

      <!-- URL edit mode -->
      <div>
        <div class="section-label" style="margin-bottom:12px">Changing the target site</div>
        <div style="background:${T.brand};border-radius:${T.radiusLg};overflow:hidden;min-width:400px">
          <div style="display:grid;grid-template-columns:auto 1fr auto;align-items:center;padding:8px 12px 8px 16px;gap:16px">
            <div style="display:flex;align-items:center;gap:8px;flex-shrink:0">
              ${icons.shield}
              <span style="color:white;font-weight:600;font-size:14px;white-space:nowrap">SharePoint Smart Permissions</span>
            </div>
            <div style="display:flex;align-items:center;gap:8px;min-width:0">
              <input
                value="https://contoso.sharepoint.com/sites/HR"
                style="flex:1;border:1px solid rgba(255,255,255,0.5);border-radius:${T.radius};padding:5px 10px;background:white;font-size:13px;font-family:inherit;color:${T.text};min-width:220px"
                readonly
              />
              <button style="background:white;border:none;border-radius:${T.radius};padding:5px 12px;font-size:13px;font-family:inherit;cursor:pointer;font-weight:500;white-space:nowrap">Connect</button>
              <button class="btn-transparent-white" style="white-space:nowrap">Cancel</button>
            </div>
            <div><button class="icon-btn">${icons.settings}</button></div>
          </div>
        </div>
        <p style="font-size:12px;color:${T.textSubtle};margin-top:10px">
          Click <strong>Change URL</strong> in the banner, enter the target site URL, then click <strong>Connect</strong>.
        </p>
      </div>
    </div>
  </div>`);
}

// ─── main ────────────────────────────────────────────────────────────────────
const screens = [
  { name: '01_home',                   html: homeHTML() },
  { name: '02_report_config',          html: reportConfigHTML() },
  { name: '03_report_running',         html: reportRunningHTML() },
  { name: '04_report_complete',        html: reportCompleteHTML() },
  { name: '05_explorer',               html: explorerHTML() },
  { name: '06_user_access_scanning',   html: userAccessRunningHTML() },
  { name: '07_user_access_complete',   html: userAccessCompleteHTML() },
  { name: '08_user_access_full_site',  html: userAccessFullHTML() },
  { name: '09_settings',              html: settingsHTML() },
];

(async () => {
  const browser = await puppeteer.launch({
    executablePath: CHROME,
    headless: 'new',
    args: ['--no-sandbox', '--disable-setuid-sandbox'],
  });

  for (const { name, html } of screens) {
    const page = await browser.newPage();
    await page.setViewport({ width: 960, height: 1, deviceScaleFactor: 2 });
    await page.setContent(html, { waitUntil: 'domcontentloaded' });

    // Auto-size height
    const bodyHeight = await page.evaluate(() => document.body.scrollHeight);
    await page.setViewport({ width: 960, height: bodyHeight, deviceScaleFactor: 2 });
    await page.setContent(html, { waitUntil: 'domcontentloaded' });

    const outPath = path.join(OUT, `${name}.png`);
    await page.screenshot({ path: outPath, fullPage: true });
    console.log(`  ✓ ${name}.png`);
    await page.close();
  }

  await browser.close();
  console.log(`\nScreenshots written to docs/screenshots/`);
})();
