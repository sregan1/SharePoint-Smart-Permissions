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
  orange:       '#835b00',
  orangeBg:     '#fef7b2',
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
  /* Banner — full-width across all screens */
  .banner {
    display: grid;
    grid-template-columns: auto 1fr auto;
    align-items: center;
    padding: 8px 12px 8px 16px;
    background: ${T.brand};
    gap: 16px;
  }
  .banner-simple {
    display: flex;
    align-items: center;
    gap: 8px;
    padding: 8px 16px;
    background: ${T.brand};
    position: relative;
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
    display: inline-flex; align-items: center; gap: 6px;
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
  /* Radio boxes — horizontal layout matching Fluent UI RadioGroup */
  .radio-group { display: flex; gap: 10px; flex-wrap: wrap; margin-top: 6px; }
  .radio-box {
    border: 1px solid ${T.border}; border-radius: ${T.radius};
    padding: 8px 14px; cursor: pointer; display: flex; align-items: center; gap: 8px;
    font-size: 13px;
  }
  .radio-box.selected {
    border: 2px solid ${T.brand}; background: ${T.blueBg};
  }
  .radio-dot {
    width: 16px; height: 16px; border-radius: 50%; border: 2px solid ${T.border};
    display: inline-flex; align-items: center; justify-content: center; flex-shrink: 0;
  }
  .radio-dot.checked { border-color: ${T.brand}; background: ${T.brand}; }
  .radio-dot.checked::after { content:''; width: 6px; height: 6px; border-radius: 50%; background: white; }
  /* Checkboxes */
  .checkbox-row { display: flex; align-items: center; gap: 8px; font-size: 13px; margin-bottom: 8px; }
  .checkbox-box {
    width: 16px; height: 16px; border: 2px solid ${T.border}; border-radius: 2px;
    display: inline-flex; align-items: center; justify-content: center; flex-shrink: 0;
  }
  .checkbox-box.checked { background: ${T.brand}; border-color: ${T.brand}; }
  .checkbox-box.checked::after { content:'✓'; color: white; font-size: 11px; line-height: 1; }
  /* Divider */
  .divider { border: none; border-top: 1px solid ${T.border}; margin: 12px 0; }
  /* SpinButton */
  .spinbutton {
    display: inline-flex; align-items: center; border: 1px solid ${T.border}; border-radius: ${T.radius}; overflow: hidden;
  }
  .spinbutton input {
    border: none; width: 50px; text-align: center; padding: 4px 8px; font-family: inherit; font-size: 13px;
  }
  .spinbutton button {
    background: ${T.bg}; border: none; border-left: 1px solid ${T.border}; padding: 4px 8px; cursor: pointer; font-size: 14px;
  }
  /* Progress */
  .progress-track { height: 4px; background: ${T.border}; border-radius: 2px; overflow: hidden; }
  .progress-fill { height: 100%; background: ${T.brand}; border-radius: 2px; }
  .progress-area {
    background: ${T.bg}; border-radius: ${T.radiusLg}; padding: 12px 16px;
    display: flex; flex-direction: column; gap: 8px;
  }
  .result-area {
    background: ${T.greenBg}; border-radius: ${T.radiusLg}; padding: 12px 16px;
    display: flex; flex-direction: column; gap: 8px;
  }
  /* Badge */
  .badge {
    display: inline-flex; align-items: center; padding: 2px 8px;
    border-radius: 10px; font-size: 11px; font-weight: 600; white-space: nowrap;
  }
  .badge-danger    { background: ${T.redBg};    color: ${T.red}; }
  .badge-warning   { background: ${T.orangeBg}; color: ${T.orange}; }
  .badge-success   { background: ${T.greenBg};  color: ${T.green}; }
  .badge-info      { background: ${T.blueBg};   color: ${T.blue}; }
  .badge-brand     { background: ${T.brand};    color: white; }
  .badge-outline   { background: white; color: ${T.textSubtle}; border: 1px solid ${T.border}; }
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
  .tree-indent  { padding-left: 20px; }
  .tree-indent2 { padding-left: 40px; }
  /* Two-col layout */
  .two-col { display: grid; grid-template-columns: 280px 1fr; gap: 0; }
  .tree-panel { border-right: 1px solid ${T.border}; padding-right: 12px; }
  .perm-panel { padding-left: 16px; }
  /* Cards */
  .card {
    background: white; border: 1px solid ${T.border}; border-radius: ${T.radiusLg};
    padding: 20px; box-shadow: ${T.shadow};
    display: flex; flex-direction: column; gap: 12px;
  }
  /* Inherited banner */
  .inherited-banner {
    display: flex; align-items: center; gap: 8px;
    padding: 8px 12px;
    background: ${T.blueBg};
    border-left: 3px solid ${T.brand};
    border-radius: ${T.radius};
    margin-bottom: 12px;
    font-size: 13px;
  }
  /* Options bar */
  .options-bar { display: flex; flex-direction: column; gap: 6px; margin-bottom: 12px; }
  /* MessageBar */
  .msgbar {
    display: flex; align-items: flex-start; gap: 10px;
    padding: 10px 14px; border-radius: ${T.radius}; margin-bottom: 12px; font-size: 13px;
  }
  .msgbar-success { background: ${T.greenBg}; color: #0e5c10; border: 1px solid #9fd89f; }
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
  shield:      `<svg width="20" height="20" viewBox="0 0 24 24" fill="none"><path d="M12 2L4 6v6c0 5.55 3.84 10.74 8 12 4.16-1.26 8-6.45 8-12V6l-8-4z" fill="white" opacity="0.9"/><path d="M10 14.5l-2-2-1 1 3 3 5-5-1-1-4 4z" fill="${T.brand}"/></svg>`,
  settings:    `<svg width="18" height="18" viewBox="0 0 24 24" fill="white"><path d="M19.14 12.94c.04-.3.06-.61.06-.94 0-.32-.02-.64-.07-.94l2.03-1.58c.18-.14.23-.41.12-.61l-1.92-3.32c-.12-.22-.37-.29-.59-.22l-2.39.96c-.5-.38-1.03-.7-1.62-.94l-.36-2.54c-.04-.24-.24-.41-.48-.41h-3.84c-.24 0-.43.17-.47.41l-.36 2.54c-.59.24-1.13.57-1.62.94l-2.39-.96c-.22-.08-.47 0-.59.22L2.74 8.87c-.12.21-.08.47.12.61l2.03 1.58c-.05.3-.09.63-.09.94s.02.64.07.94l-2.03 1.58c-.18.14-.23.41-.12.61l1.92 3.32c.12.22.37.29.59.22l2.39-.96c.5.38 1.03.7 1.62.94l.36 2.54c.05.24.24.41.48.41h3.84c.24 0 .44-.17.47-.41l.36-2.54c.59-.24 1.13-.56 1.62-.94l2.39.96c.22.08.47 0 .59-.22l1.92-3.32c.12-.22.07-.47-.12-.61l-2.01-1.58zM12 15.6c-1.98 0-3.6-1.62-3.6-3.6s1.62-3.6 3.6-3.6 3.6 1.62 3.6 3.6-1.62 3.6-3.6 3.6z"/></svg>`,
  folder:      `<svg width="16" height="16" viewBox="0 0 24 24" fill="none"><path d="M10 4H4c-1.1 0-2 .9-2 2v12c0 1.1.9 2 2 2h16c1.1 0 2-.9 2-2V8c0-1.1-.9-2-2-2h-8l-2-2z" fill="#f0a500"/></svg>`,
  file:        `<svg width="14" height="16" viewBox="0 0 24 24" fill="none"><path d="M14 2H6c-1.1 0-2 .9-2 2v16c0 1.1.9 2 2 2h12c1.1 0 2-.9 2-2V8l-6-6zm2 16H8v-2h8v2zm0-4H8v-2h8v2zm-3-5V3.5L18.5 9H13z" fill="#0078d4"/></svg>`,
  chevronDown: `<svg width="12" height="12" viewBox="0 0 24 24" fill="${T.textSubtle}"><path d="M7 10l5 5 5-5H7z"/></svg>`,
  chevronRight:`<svg width="12" height="12" viewBox="0 0 24 24" fill="${T.textSubtle}"><path d="M10 17l5-5-5-5v10z"/></svg>`,
  arrowCircleDown:`<svg width="14" height="14" viewBox="0 0 24 24" fill="none"><circle cx="12" cy="12" r="9" stroke="${T.textSubtle}" stroke-width="2"/><path d="M12 8v8M8.5 13.5L12 17l3.5-3.5" stroke="${T.textSubtle}" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"/></svg>`,
  check:       `<svg width="14" height="14" viewBox="0 0 24 24" fill="${T.brand}"><path d="M9 16.17L4.83 12l-1.42 1.41L9 19 21 7l-1.41-1.41L9 16.17z"/></svg>`,
  person:      `<svg width="14" height="14" viewBox="0 0 24 24" fill="${T.textSubtle}"><path d="M12 12c2.21 0 4-1.79 4-4s-1.79-4-4-4-4 1.79-4 4 1.79 4 4 4zm0 2c-2.67 0-8 1.34-8 4v2h16v-2c0-2.66-5.33-4-8-4z"/></svg>`,
  people:      `<svg width="14" height="14" viewBox="0 0 24 24" fill="${T.textSubtle}"><path d="M16 11c1.66 0 2.99-1.34 2.99-3S17.66 5 16 5c-1.66 0-3 1.34-3 3s1.34 3 3 3zm-8 0c1.66 0 2.99-1.34 2.99-3S9.66 5 8 5C6.34 5 5 6.34 5 8s1.34 3 3 3zm0 2c-2.33 0-7 1.17-7 3.5V19h14v-2.5c0-2.33-4.67-3.5-7-3.5zm8 0c-.29 0-.62.02-.97.05 1.16.84 1.97 1.97 1.97 3.45V19h6v-2.5c0-2.33-4.67-3.5-7-3.5z"/></svg>`,
  back:        `<svg width="14" height="14" viewBox="0 0 24 24" fill="${T.text}"><path d="M20 11H7.83l5.59-5.59L12 4l-8 8 8 8 1.41-1.41L7.83 13H20v-2z"/></svg>`,
  info:        `<svg width="14" height="14" viewBox="0 0 24 24" fill="${T.brand}"><path d="M12 2C6.48 2 2 6.48 2 12s4.48 10 10 10 10-4.48 10-10S17.52 2 12 2zm1 15h-2v-6h2v6zm0-8h-2V7h2v2z"/></svg>`,
  excel:       `<svg width="16" height="16" viewBox="0 0 24 24" fill="none"><rect width="24" height="24" rx="3" fill="#1D6F42"/><path d="M7 6l3.5 5.5L7 17h2.5l2-3.2 2 3.2H16l-3.5-5.5L16 6h-2.5l-2 3.2L9.5 6H7z" fill="white"/></svg>`,
  docArrow:    `<svg width="16" height="16" viewBox="0 0 24 24" fill="white"><path d="M14 2H6c-1.1 0-2 .9-2 2v16c0 1.1.9 2 2 2h12c1.1 0 2-.9 2-2V8l-6-6zm-1 14l-3-3h2v-4h2v4h2l-3 3zm2-9V3.5L18.5 7H15z"/></svg>`,
  spinner:     `<svg width="16" height="16" viewBox="0 0 24 24" fill="none"><circle cx="12" cy="12" r="9" stroke="${T.border}" stroke-width="3"/><path d="M12 3a9 9 0 0 1 9 9" stroke="${T.brand}" stroke-width="3" stroke-linecap="round"/></svg>`,
  link:        `<svg width="14" height="14" viewBox="0 0 24 24" fill="${T.brand}"><path d="M3.9 12c0-1.71 1.39-3.1 3.1-3.1h4V7H7c-2.76 0-5 2.24-5 5s2.24 5 5 5h4v-1.9H7c-1.71 0-3.1-1.39-3.1-3.1zM8 13h8v-2H8v2zm9-6h-4v1.9h4c1.71 0 3.1 1.39 3.1 3.1 0 1.71-1.39 3.1-3.1 3.1h-4V17h4c2.76 0 5-2.24 5-5s-2.24-5-5-5z"/></svg>`,
  globe:       `<svg width="14" height="14" viewBox="0 0 24 24" fill="${T.textSubtle}"><path d="M12 2C6.48 2 2 6.48 2 12s4.48 10 10 10 10-4.48 10-10S17.52 2 12 2zm-1 17.93c-3.95-.49-7-3.85-7-7.93 0-.62.08-1.21.21-1.79L9 15v1c0 1.1.9 2 2 2v1.93zm6.9-2.54c-.26-.81-1-1.39-1.9-1.39h-1v-3c0-.55-.45-1-1-1H8v-2h2c.55 0 1-.45 1-1V7h2c1.1 0 2-.9 2-2v-.41c2.93 1.19 5 4.06 5 7.41 0 2.08-.8 3.97-2.1 5.39z"/></svg>`,
  bookDb:      `<svg width="14" height="14" viewBox="0 0 24 24" fill="${T.textSubtle}"><path d="M4 6H2v14c0 1.1.9 2 2 2h14v-2H4V6zm16-4H8c-1.1 0-2 .9-2 2v12c0 1.1.9 2 2 2h12c1.1 0 2-.9 2-2V4c0-1.1-.9-2-2-2zm-1 9H9V9h10v2zm-4 4H9v-2h6v2zm4-8H9V5h10v2z"/></svg>`,
  folderIcon:  `<svg width="14" height="14" viewBox="0 0 24 24" fill="${T.textSubtle}"><path d="M10 4H4c-1.1 0-2 .9-2 2v12c0 1.1.9 2 2 2h16c1.1 0 2-.9 2-2V8c0-1.1-.9-2-2-2h-8l-2-2z"/></svg>`,
  folderOpen:  `<svg width="14" height="14" viewBox="0 0 24 24" fill="${T.textSubtle}"><path d="M20 6h-8l-2-2H4c-1.1 0-2 .9-2 2v12c0 1.1.9 2 2 2h16c1.1 0 2-.9 2-2V8c0-1.1-.9-2-2-2zm0 12H4V8h16v10z"/></svg>`,
};

// ─── banner component — full-width top bar on all non-home screens ────────────
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

// ─── home banner — simple left-aligned brand strip ────────────────────────────
function homeBanner() {
  return `<div class="banner-simple">
  ${icons.shield}
  <span style="color:white;font-weight:600;white-space:nowrap;font-size:14px">SharePoint Smart Permissions</span>
  <button class="icon-btn" style="position:absolute;right:8px;top:4px">${icons.settings}</button>
</div>`;
}

// ─── SCREEN 1: Home ───────────────────────────────────────────────────────────
function homeHTML() {
  const cards = [
    {
      icon: `<svg width="20" height="20" viewBox="0 0 24 24" fill="${T.brand}"><path d="M19 3H5c-1.1 0-2 .9-2 2v14c0 1.1.9 2 2 2h14c1.1 0 2-.9 2-2V5c0-1.1-.9-2-2-2zm-5 14H7v-2h7v2zm3-4H7v-2h10v2zm0-4H7V7h10v2z"/></svg>`,
      title: 'Permissions Report',
      desc: 'Generate a comprehensive Excel report showing every unique permission assignment across your site. Only highlights differences from inherited permissions.',
      features: [
        'Site, Library, Folder, or Item level',
        'Configurable folder depth',
        'Color-coded Excel export',
        'Scan all sites or a single site',
      ],
      btn: 'Run Permissions Report',
    },
    {
      icon: `<svg width="20" height="20" viewBox="0 0 24 24" fill="${T.brand}"><path d="M10 4H4c-1.1 0-2 .9-2 2v12c0 1.1.9 2 2 2h16c1.1 0 2-.9 2-2V8c0-1.1-.9-2-2-2h-8l-2-2zM5 18l3-6 2.06 2.54L13 10l5 8H5z"/></svg>`,
      title: 'Permissions Explorer',
      desc: 'Browse permissions interactively in real time. Select any folder or file to instantly see who has access and what permission levels are assigned.',
      features: [
        'Interactive folder/file tree',
        'Instant permission lookup',
        'Unique vs. inherited permissions',
        'Expand SharePoint group members',
      ],
      btn: 'Open Permissions Explorer',
    },
    {
      icon: `<svg width="20" height="20" viewBox="0 0 24 24" fill="${T.brand}"><path d="M12 12c2.21 0 4-1.79 4-4s-1.79-4-4-4-4 1.79-4 4 1.79 4 4 4zm0 2c-2.67 0-8 1.34-8 4v2h16v-2c0-2.66-5.33-4-8-4z"/></svg>`,
      title: 'User Access',
      desc: 'Look up a specific user to see every location they can access on a site. Quickly identify over-privileged accounts or verify that access is correctly scoped.',
      features: [
        'Per-user access analysis',
        'Full Site Access detection',
        'Shows path and permission level',
        'Search users by name',
      ],
      btn: 'Check User Access',
    },
  ];

  return shell('Home', `
  ${homeBanner()}
  <div style="padding: 24px 40px 32px; max-width:1100px; margin:0 auto">
    <div style="margin-bottom:24px">
      <p style="color:${T.textSubtle};font-size:13px;max-width:700px">
        Audit and understand SharePoint permissions in real time, directly from your browser — no PowerShell or admin tools required.
      </p>
    </div>
    <div style="display:grid;grid-template-columns:repeat(3,1fr);gap:20px;margin-bottom:24px">
      ${cards.map(c => `
      <div class="card">
        <div style="display:flex;align-items:center;gap:10px">
          ${c.icon}
          <span style="font-size:15px;font-weight:600">${c.title}</span>
        </div>
        <p style="color:${T.textSubtle};font-size:12px;line-height:1.5">${c.desc}</p>
        <div style="display:flex;flex-direction:column;gap:5px;flex-grow:1">
          ${c.features.map(f => `<div style="display:flex;align-items:flex-start;gap:6px;font-size:12px">
            <span style="flex-shrink:0;margin-top:1px">${icons.check}</span>
            <span>${f}</span>
          </div>`).join('')}
        </div>
        <button class="btn-primary" style="width:100%;justify-content:center;padding:9px 0;margin-top:4px">${c.btn}</button>
      </div>`).join('')}
    </div>
    <div style="padding:12px 14px;background:${T.bg};border-radius:${T.radius}">
      <p style="font-size:12px;color:${T.textSubtle};line-height:1.5">
        <strong>Note:</strong> This web part runs as the currently signed-in user. It can only see sites and items that user has permission to view. For a full tenant scan, use an account with appropriate read access across all sites.
      </p>
    </div>
  </div>`);
}

// ─── SCREEN 2: Permissions Report — Config ────────────────────────────────────
function reportConfigHTML() {
  const scopeOptions = [
    { value: 'Site',    icon: icons.globe,      label: 'Site only',       selected: true  },
    { value: 'Library', icon: icons.bookDb,     label: 'Libraries',       selected: false },
    { value: 'Folder',  icon: icons.folderIcon, label: 'Folders',         selected: false },
    { value: 'Item',    icon: icons.folderOpen,  label: 'Files & Folders', selected: false },
  ];
  return shell('Permissions Report - Config', `
  ${banner()}
  <div class="content">
    <div class="back-row">
      <button class="btn-back">${icons.back} Back</button>
    </div>
    <div class="title">Permissions Report</div>
    <div style="margin-bottom:24px"></div>
    <div style="display:flex;flex-direction:column;gap:14px;max-width:680px">
      <!-- All-sites checkbox -->
      <div class="checkbox-row" style="color:${T.textDisabled}">
        <span class="checkbox-box" style="border-color:${T.textDisabled}"></span>
        <span style="color:${T.textDisabled}">Scan all site collections in this tenant (only available in root site)</span>
      </div>

      <hr class="divider"/>

      <!-- Scan depth -->
      <div>
        <span class="form-label">Scan depth</span>
        <div class="radio-group">
          ${scopeOptions.map(o => `
          <div class="radio-box ${o.selected ? 'selected' : ''}">
            <span class="radio-dot ${o.selected ? 'checked' : ''}"></span>
            <span style="display:flex;align-items:center;gap:5px;font-size:13px;color:${o.selected ? T.brand : T.text}">
              ${o.icon}${o.label}
            </span>
          </div>`).join('')}
        </div>
      </div>

      <!-- Expand groups checkbox (checked by default) -->
      <div class="checkbox-row">
        <span class="checkbox-box checked"></span>
        <span>Expand group members in report (SharePoint groups, Security groups, and M365 groups)</span>
      </div>

      <hr class="divider"/>

      <div>
        <button class="btn-primary">Run Report</button>
      </div>
    </div>
  </div>`);
}

// ─── SCREEN 3: Permissions Report — Running ───────────────────────────────────
function reportRunningHTML() {
  const scopeOptions = [
    { icon: icons.globe,      label: 'Site only',       selected: false },
    { icon: icons.bookDb,     label: 'Libraries',        selected: false },
    { icon: icons.folderIcon, label: 'Folders',          selected: true  },
    { icon: icons.folderOpen, label: 'Files & Folders',  selected: false },
  ];
  return shell('Permissions Report - Running', `
  ${banner()}
  <div class="content">
    <div class="back-row">
      <button class="btn-back" style="color:${T.textDisabled};border-color:${T.border}">${icons.back} Back</button>
    </div>
    <div class="title">Permissions Report</div>
    <div style="margin-bottom:24px"></div>
    <div style="display:flex;flex-direction:column;gap:14px;max-width:680px">

      <div class="checkbox-row" style="color:${T.textDisabled}">
        <span class="checkbox-box" style="border-color:${T.textDisabled}"></span>
        <span style="color:${T.textDisabled}">Scan all site collections in this tenant (only available in root site)</span>
      </div>

      <hr class="divider"/>

      <div>
        <span class="form-label" style="color:${T.textSubtle}">Scan depth</span>
        <div class="radio-group" style="opacity:0.6;pointer-events:none">
          ${scopeOptions.map(o => `
          <div class="radio-box ${o.selected ? 'selected' : ''}">
            <span class="radio-dot ${o.selected ? 'checked' : ''}"></span>
            <span style="display:flex;align-items:center;gap:5px;font-size:13px">${o.icon}${o.label}</span>
          </div>`).join('')}
        </div>
      </div>

      <!-- Folder depth (visible, disabled while running) -->
      <div style="display:flex;align-items:center;gap:10px;opacity:0.6">
        <span style="font-size:13px;font-weight:500">Folder depth limit:</span>
        <div class="spinbutton">
          <input value="3" readonly/>
          <button>▲</button>
          <button style="border-left:1px solid ${T.border}">▼</button>
        </div>
        <span style="font-size:12px;color:${T.textSubtle}">levels deep</span>
      </div>

      <!-- Expand groups checkbox (disabled while running) -->
      <div class="checkbox-row" style="opacity:0.6">
        <span class="checkbox-box checked"></span>
        <span>Expand group members in report (SharePoint groups, Security groups, and M365 groups)</span>
      </div>

      <hr class="divider"/>

      <div style="display:flex;align-items:center;gap:10px">
        <button class="btn-disabled">Run Report</button>
        <button class="btn-secondary">Cancel</button>
      </div>

      <!-- Progress area -->
      <div class="progress-area">
        <div class="progress-track">
          <div class="progress-fill" style="width:58%"></div>
        </div>
        <div style="display:flex;justify-content:space-between;align-items:center">
          <span style="font-size:13px">Campaign Documents / Q4 Launch / Assets</span>
          <span style="font-size:13px;color:${T.textSubtle};white-space:nowrap">0:32</span>
        </div>
        <span style="font-size:13px;color:${T.textSubtle}">43 items scanned · Library 7 of 12</span>
      </div>
    </div>
  </div>`);
}

// ─── SCREEN 4: Permissions Report — Complete ─────────────────────────────────
function reportCompleteHTML() {
  const scopeOptions = [
    { icon: icons.globe,      label: 'Site only',       selected: false },
    { icon: icons.bookDb,     label: 'Libraries',        selected: false },
    { icon: icons.folderIcon, label: 'Folders',          selected: true  },
    { icon: icons.folderOpen, label: 'Files & Folders',  selected: false },
  ];
  return shell('Permissions Report - Complete', `
  ${banner()}
  <div class="content">
    <div class="back-row">
      <button class="btn-back">${icons.back} Back</button>
    </div>
    <div class="title">Permissions Report</div>
    <div style="margin-bottom:24px"></div>
    <div style="display:flex;flex-direction:column;gap:14px;max-width:680px">

      <div class="checkbox-row" style="color:${T.textDisabled}">
        <span class="checkbox-box" style="border-color:${T.textDisabled}"></span>
        <span style="color:${T.textDisabled}">Scan all site collections in this tenant (only available in root site)</span>
      </div>

      <hr class="divider"/>

      <div>
        <span class="form-label">Scan depth</span>
        <div class="radio-group">
          ${scopeOptions.map(o => `
          <div class="radio-box ${o.selected ? 'selected' : ''}">
            <span class="radio-dot ${o.selected ? 'checked' : ''}"></span>
            <span style="display:flex;align-items:center;gap:5px;font-size:13px;color:${o.selected ? T.brand : T.text}">${o.icon}${o.label}</span>
          </div>`).join('')}
        </div>
      </div>

      <!-- Folder depth (visible since Folders selected) -->
      <div style="display:flex;align-items:center;gap:10px">
        <span style="font-size:13px;font-weight:500">Folder depth limit:</span>
        <div class="spinbutton">
          <input value="3" readonly/>
          <button>▲</button>
          <button style="border-left:1px solid ${T.border}">▼</button>
        </div>
        <span style="font-size:12px;color:${T.textSubtle}">levels deep</span>
      </div>

      <!-- Expand groups checkbox (checked) -->
      <div class="checkbox-row">
        <span class="checkbox-box checked"></span>
        <span>Expand group members in report (SharePoint groups, Security groups, and M365 groups)</span>
      </div>

      <hr class="divider"/>

      <div>
        <button class="btn-primary">Run Report</button>
      </div>

      <!-- Completion message (gray area — final status text) -->
      <div class="progress-area">
        <span style="font-size:13px;color:${T.textSubtle}">Scan complete — 75 object(s) found, 5 with unique permissions.</span>
      </div>

      <!-- Result area (green) -->
      <div class="result-area">
        <div style="display:flex;align-items:center;gap:8px;flex-wrap:wrap">
          <span style="font-weight:600;font-size:13px">Scan complete</span>
          <span class="badge badge-success">75 objects</span>
          <span class="badge badge-warning">5 unique</span>
          <span class="badge badge-outline">70 inherited</span>
        </div>
        <span style="font-size:13px;color:${T.textSubtle}">Scan complete — 75 object(s) found, 5 with unique permissions.</span>
        <div>
          <button class="btn-primary">${icons.docArrow} Export to Excel</button>
        </div>
      </div>
    </div>
  </div>`);
}

// ─── SCREEN 4b: Permissions Report — Export ready ────────────────────────────
function reportExportHTML() {
  const rows = [
    { type: 'Site',    path: '/sites/Marketing',                                         unique: true  },
    { type: 'Library', path: '/sites/Marketing/Campaign Documents',                      unique: true  },
    { type: 'Folder',  path: '/sites/Marketing/Campaign Documents/Q4 Launch',            unique: true  },
    { type: 'File',    path: '/sites/Marketing/Campaign Documents/Q4 Launch/Budget.xlsx', unique: true  },
    { type: 'Folder',  path: '/sites/Marketing/Marketing Assets/Brand Photos',           unique: true  },
  ];
  return shell('Permissions Report - Export Ready', `
  ${banner()}
  <div class="content">
    <div class="back-row">
      <button class="btn-back">${icons.back} Back</button>
    </div>
    <div class="title">Permissions Report</div>
    <div style="margin-bottom:16px"></div>

    <!-- Compact summary of chosen settings -->
    <div style="display:flex;gap:8px;flex-wrap:wrap;margin-bottom:16px">
      <span class="badge badge-outline" style="font-size:12px">${icons.folderIcon}&nbsp; Folders</span>
      <span class="badge badge-outline" style="font-size:12px">Depth: 3</span>
      <span class="badge badge-outline" style="font-size:12px">Expand groups: on</span>
    </div>

    <!-- Grey status area -->
    <div class="progress-area" style="max-width:680px;margin-bottom:12px">
      <span style="font-size:13px;color:${T.textSubtle}">Scan complete — 147 items scanned across 4 libraries in 0:38.</span>
    </div>

    <!-- Green result area -->
    <div class="result-area" style="max-width:680px;margin-bottom:20px">
      <div style="display:flex;align-items:center;gap:8px;flex-wrap:wrap">
        <span style="font-weight:600;font-size:14px;color:${T.green}">Scan complete</span>
        <span class="badge badge-success">147 objects</span>
        <span class="badge badge-warning">5 unique permissions</span>
        <span class="badge badge-outline">142 inherited</span>
      </div>
      <p style="font-size:13px;color:${T.textSubtle};margin:4px 0 12px">
        5 location(s) have unique permission assignments. The Excel export includes one row per
        user or group assignment, with an <strong style="color:${T.text}">Access Via</strong> column
        showing group membership.
      </p>
      <button class="btn-primary" style="font-size:14px;padding:8px 20px;gap:8px">
        ${icons.docArrow} Export to Excel
      </button>
    </div>

    <!-- Preview table of unique items -->
    <div style="max-width:680px">
      <p style="font-size:12px;font-weight:600;color:${T.textSubtle};text-transform:uppercase;letter-spacing:0.5px;margin-bottom:8px">Items with unique permissions</p>
      <table>
        <thead>
          <tr><th>Type</th><th>Path</th></tr>
        </thead>
        <tbody>
          ${rows.map(r => `<tr>
            <td><span class="badge badge-${r.type === 'Site' ? 'brand' : r.type === 'Library' ? 'info' : r.type === 'Folder' ? 'warning' : 'outline'}">${r.type}</span></td>
            <td style="font-size:12px;color:${T.textSubtle}">${r.path}</td>
          </tr>`).join('')}
        </tbody>
      </table>
    </div>
  </div>`);
}

// ─── SCREEN 5: Permissions Explorer ──────────────────────────────────────────
function explorerHTML() {
  return shell('Permissions Explorer', `
  ${banner()}
  <div class="content">
    <div class="back-row">
      <button class="btn-back">${icons.back} Back</button>
    </div>
    <div class="title">Permissions Explorer</div>
    <div style="margin-bottom:4px">
      <span style="font-size:12px;color:${T.brand}">Connected — 4 libraries found</span>
    </div>

    <div class="form-row" style="margin-bottom:16px;margin-top:12px">
      <span class="form-label">Library</span>
      <select style="border:1px solid ${T.border};border-radius:${T.radius};padding:5px 10px;font-family:inherit;font-size:13px;max-width:320px;color:${T.text}">
        <option>Campaign Documents</option>
      </select>
    </div>

    <div class="two-col">
      <!-- Tree panel -->
      <div class="tree-panel">
        <div class="tree-item">
          ${icons.chevronDown}${icons.folder}
          <span>Campaign Documents</span>
        </div>
        <div class="tree-indent">
          <div class="tree-item">
            ${icons.chevronDown}${icons.folder}
            <span style="flex:1">Q4 Launch</span>
            <span class="badge badge-warning" style="font-size:10px">Unique</span>
          </div>
          <div class="tree-indent">
            <div class="tree-item selected">
              ${icons.chevronRight}${icons.folder}
              <span style="flex:1">Assets</span>
              <span title="Contains items with unique permissions">${icons.arrowCircleDown}</span>
            </div>
            <div class="tree-item">
              ${icons.file}
              <span>Campaign_Brief.docx</span>
            </div>
            <div class="tree-item">
              ${icons.file}
              <span style="flex:1">Budget_Confidential.xlsx</span>
              <span class="badge badge-warning" style="font-size:10px">Unique</span>
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

      <!-- Permissions panel -->
      <div class="perm-panel">
        <!-- Item name + badge -->
        <div style="display:flex;align-items:center;gap:8px;margin-bottom:12px;flex-wrap:wrap">
          <span style="font-weight:600;font-size:14px">Assets</span>
          <span class="badge badge-info">Inherited</span>
        </div>

        <!-- Options bar — always at top -->
        <div class="options-bar">
          <div class="checkbox-row">
            <span class="checkbox-box"></span>
            <span>Expand group members</span>
          </div>
          <div class="checkbox-row">
            <span class="checkbox-box checked"></span>
            <span>Show parent permissions</span>
          </div>
        </div>

        <!-- Inherited banner -->
        <div class="inherited-banner">
          ${icons.link}
          <span>This item inherits permissions from its parent.</span>
        </div>

        <!-- Parent perms label -->
        <p style="font-size:12px;color:${T.textSubtle};margin-bottom:8px">Inherited from: <strong>Q4 Launch</strong></p>

        <!-- Parent permissions table -->
        <table>
          <thead>
            <tr><th>User / Group</th><th>Type</th><th>Permission Level</th></tr>
          </thead>
          <tbody>
            <tr>
              <td style="display:flex;align-items:center;gap:6px">${icons.people}<span>Marketing Owners</span></td>
              <td><span style="font-size:12px;color:${T.textSubtle}">SP Group</span></td>
              <td><span class="badge badge-danger">Full Control</span></td>
            </tr>
            <tr>
              <td style="display:flex;align-items:center;gap:6px">${icons.people}<span>Marketing Members</span></td>
              <td><span style="font-size:12px;color:${T.textSubtle}">SP Group</span></td>
              <td><span class="badge badge-warning">Edit</span></td>
            </tr>
            <tr>
              <td style="display:flex;align-items:center;gap:6px">${icons.person}<span>Carol White</span></td>
              <td><span style="font-size:12px;color:${T.textSubtle}">User</span></td>
              <td><span class="badge badge-warning">Edit</span></td>
            </tr>
            <tr>
              <td style="display:flex;align-items:center;gap:6px">${icons.people}<span>Marketing Visitors</span></td>
              <td><span style="font-size:12px;color:${T.textSubtle}">SP Group</span></td>
              <td><span class="badge badge-success">Read</span></td>
            </tr>
            <tr>
              <td style="display:flex;align-items:center;gap:6px">${icons.people}<span>External Reviewers</span></td>
              <td><span style="font-size:12px;color:${T.textSubtle}">SP Group</span></td>
              <td><span class="badge badge-success">Read</span></td>
            </tr>
          </tbody>
        </table>
      </div>
    </div>
  </div>`);
}

// ─── SCREEN 6: User Access — Scanning ────────────────────────────────────────
function userAccessRunningHTML() {
  return shell('User Access - Scanning', `
  ${banner()}
  <div class="content">
    <div class="back-row">
      <button class="btn-back" style="color:${T.textDisabled};border-color:${T.border}">${icons.back} Back</button>
    </div>
    <div class="title">User Access</div>
    <div style="margin-bottom:4px">
      <span style="font-size:12px;color:${T.brand}">Connected — 18 users found</span>
    </div>

    <div class="form-row" style="margin-top:16px;margin-bottom:16px">
      <span class="form-label">Select a user</span>
      <select disabled style="border:1px solid ${T.border};border-radius:${T.radius};padding:5px 10px;font-family:inherit;font-size:13px;max-width:360px;color:${T.text};opacity:0.7">
        <option>Alice Chen</option>
      </select>
    </div>

    <!-- Scan area -->
    <div class="progress-area" style="max-width:560px">
      <!-- Indeterminate progress bar -->
      <div class="progress-track">
        <div class="progress-fill" style="width:38%"></div>
      </div>
      <div style="display:flex;align-items:center;gap:10px;justify-content:space-between">
        <div style="display:flex;align-items:center;gap:8px">
          ${icons.spinner}
          <span style="font-size:13px">Scanning Campaign Documents…</span>
        </div>
        <span style="font-size:13px;color:${T.textSubtle};white-space:nowrap">0:47</span>
      </div>
      <div style="display:flex;align-items:center;gap:12px;flex-wrap:wrap">
        <button class="btn-secondary" style="font-size:12px;padding:4px 12px">Cancel</button>
        <span style="font-size:12px;color:${T.textSubtle}">This scan may take several minutes depending on the size of the site.</span>
      </div>
    </div>
  </div>`);
}

// ─── SCREEN 7: User Access — Complete ─────────────────────────────────────────
function userAccessCompleteHTML() {
  // objectType badge colors from UserAccessView.tsx:
  // Site → brand (blue), Library → informative (blue-info), Folder → warning (orange), File → no color
  const rows = [
    { type: 'Site',    typeCls: 'badge-brand',    name: 'Marketing',               path: '/sites/Marketing',                                                    role: 'Read',         roleCls: 'badge-success', depth: 0 },
    { type: 'Library', typeCls: 'badge-info',     name: 'Campaign Documents',       path: '/sites/Marketing/Campaign Documents',                                 role: 'Edit',         roleCls: 'badge-warning', depth: 0 },
    { type: 'Folder',  typeCls: 'badge-warning',  name: 'Q4 Launch',                path: '/sites/Marketing/Campaign Documents/Q4 Launch',                       role: 'Edit',         roleCls: 'badge-warning', depth: 1 },
    { type: 'File',    typeCls: '',               name: 'Budget_Confidential.xlsx', path: '/sites/Marketing/Campaign Documents/Q4 Launch/Budget_Confidential.xlsx', role: 'Full Control', roleCls: 'badge-danger',  depth: 2 },
    { type: 'Library', typeCls: 'badge-info',     name: 'Marketing Assets',         path: '/sites/Marketing/Marketing Assets',                                   role: 'Read',         roleCls: 'badge-success', depth: 0 },
    { type: 'Folder',  typeCls: 'badge-warning',  name: 'Brand Photos',             path: '/sites/Marketing/Marketing Assets/Brand Photos',                      role: 'Edit',         roleCls: 'badge-warning', depth: 1 },
  ];

  return shell('User Access - Complete', `
  ${banner()}
  <div class="content">
    <div class="back-row">
      <button class="btn-back">${icons.back} Back</button>
    </div>
    <div class="title">User Access</div>
    <div style="margin-bottom:4px">
      <span style="font-size:12px;color:${T.brand}">Connected — 18 users found</span>
    </div>

    <div class="form-row" style="margin-top:16px;margin-bottom:16px">
      <span class="form-label">Select a user</span>
      <select style="border:1px solid ${T.border};border-radius:${T.radius};padding:5px 10px;font-family:inherit;font-size:13px;max-width:360px;color:${T.text}">
        <option>Alice Chen</option>
      </select>
    </div>

    <p style="font-size:13px;color:${T.textSubtle};margin-bottom:14px">6 accessible location(s) found.</p>

    <table>
      <thead>
        <tr><th>Type</th><th>Name</th><th>Path</th><th>Permission Level</th></tr>
      </thead>
      <tbody>
        ${rows.map(r => `<tr>
          <td>${r.typeCls ? `<span class="badge ${r.typeCls}">${r.type}</span>` : `<span style="font-size:12px;color:${T.textSubtle}">${r.type}</span>`}</td>
          <td style="padding-left:${8 + r.depth * 12}px">${r.name}</td>
          <td style="font-size:11px;color:${T.textSubtle}">${r.path}</td>
          <td><span class="badge ${r.roleCls}">${r.role}</span></td>
        </tr>`).join('')}
      </tbody>
    </table>
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
    <div style="margin-bottom:4px">
      <span style="font-size:12px;color:${T.brand}">Connected — 18 users found</span>
    </div>

    <div class="form-row" style="margin-top:16px;margin-bottom:16px">
      <span class="form-label">Select a user</span>
      <select style="border:1px solid ${T.border};border-radius:${T.radius};padding:5px 10px;font-family:inherit;font-size:13px;max-width:360px;color:${T.text}">
        <option>Bob Martinez</option>
      </select>
    </div>

    <!-- Success MessageBar (matches the actual MessageBar intent="success" in UserAccessView) -->
    <div class="msgbar msgbar-success" style="max-width:600px">
      <svg width="16" height="16" viewBox="0 0 24 24" fill="#107c10" style="flex-shrink:0;margin-top:1px"><path d="M12 2C6.48 2 2 6.48 2 12s4.48 10 10 10 10-4.48 10-10S17.52 2 12 2zm-2 15l-5-5 1.41-1.41L10 14.17l7.59-7.59L19 8l-9 9z"/></svg>
      <span style="font-size:13px;line-height:1.5">
        This user has <strong>Full Control</strong> or Owner-level access to the entire site — all libraries and folders are accessible.
      </span>
    </div>
  </div>`);
}

// ─── SCREEN 9: Settings full page ────────────────────────────────────────────
function settingsHTML() {
  return shell('Settings', `
  ${banner()}
  <div class="content" style="max-width:540px">
    <div class="back-row">
      <button class="btn-back">${icons.back} Back</button>
    </div>
    <div class="title" style="margin-bottom:24px">Settings</div>

    <!-- Libraries section -->
    <div style="margin-bottom:24px">
      <div style="font-weight:600;font-size:14px;margin-bottom:10px">Libraries</div>
      <div class="checkbox-row" style="margin-bottom:6px">
        <span class="checkbox-box checked"></span>
        <label style="font-size:13px">Include system and hidden libraries</label>
        <span style="margin-left:2px">${icons.info}</span>
      </div>
      <p style="font-size:12px;color:${T.textSubtle};margin-left:24px;line-height:1.5">
        When checked, includes Style Library, Form Templates, Site Assets, and other libraries hidden
        from default views. Applies to Permissions Explorer and User Access.
      </p>
    </div>

    <hr class="divider" style="margin:20px 0"/>

    <!-- Default view instructions section -->
    <div>
      <div style="font-weight:600;font-size:14px;margin-bottom:10px">Default view on load</div>
      <p style="font-size:13px;color:${T.textSubtle};margin-bottom:10px;line-height:1.5">
        To change which screen opens when the web part first loads, edit the web part properties:
      </p>
      <ol style="margin:0;padding-left:20px;line-height:2;font-size:13px;color:${T.textSubtle}">
        <li>Put the SharePoint page into <strong style="color:${T.text}">Edit</strong> mode.</li>
        <li>Click the <strong style="color:${T.text}">pencil (edit)</strong> icon on the Smart Permissions web part.</li>
        <li>In the property panel, choose a view from the <strong style="color:${T.text}">Default view on open</strong> dropdown.</li>
        <li><strong style="color:${T.text}">Republish</strong> the page to save the change.</li>
      </ol>
    </div>
  </div>`);
}

// ─── main ────────────────────────────────────────────────────────────────────
const screens = [
  { name: '01_home',                   html: homeHTML() },
  { name: '02_report_config',          html: reportConfigHTML() },
  { name: '03_report_running',         html: reportRunningHTML() },
  { name: '04_report_complete',        html: reportCompleteHTML() },
  { name: '04b_report_export',         html: reportExportHTML() },
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
