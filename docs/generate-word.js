// docs/generate-word.js
// Generates docs/UserGuide.docx from content mirroring UserGuide.md.
// Run: node docs/generate-word.js
'use strict';

const {
  Document, Packer, Paragraph, TextRun, HeadingLevel,
  Table, TableRow, TableCell, ImageRun,
  AlignmentType, WidthType, TableBorders, BorderStyle,
  ShadingType, convertInchesToTwip, PageBreak,
  ExternalHyperlink, UnderlineType, LevelFormat,
  NumberFormat, Header, Footer, PageNumber,
} = require('docx');
const fs = require('fs');
const path = require('path');
const sizeOf = require('image-size');

const SCREENSHOTS = path.join(__dirname, 'screenshots');
const OUT = path.join(__dirname, 'UserGuide.docx');

// ─── colour palette ───────────────────────────────────────────────────────────
const C = {
  brand:      '0078D4',
  brandDark:  '005A9E',
  white:      'FFFFFF',
  black:      '000000',
  text:       '242424',
  subtle:     '616161',
  border:     'D1D1D1',
  headerBg:   'EFF6FC',
  rowAlt:     'F9F9F9',
  green:      '107C10',
  greenBg:    'DFF6DD',
  orange:     'CA5010',
  orangeBg:   'FFF4CE',
  red:        'A4262C',
  redBg:      'FDE7E9',
  blue:       '0078D4',
  blueBg:     'EFF6FC',
};

// ─── helpers ──────────────────────────────────────────────────────────────────
function pt(n) { return n * 2; }  // half-points for font sizes

function img(filename, maxWidthPx = 600) {
  const fullPath = path.join(SCREENSHOTS, filename);
  if (!fs.existsSync(fullPath)) {
    console.warn(`  ⚠ Missing screenshot: ${filename}`);
    return new Paragraph({ text: `[Screenshot: ${filename}]`, style: 'Normal' });
  }
  const data = fs.readFileSync(fullPath);
  const sizeFn = sizeOf.default || sizeOf;
  const dims = sizeFn(data);
  const aspect = dims.height / dims.width;
  const w = maxWidthPx;
  const h = Math.round(w * aspect);
  return new Paragraph({
    children: [new ImageRun({ data, transformation: { width: w, height: h }, type: 'png' })],
    spacing: { before: 80, after: 120 },
  });
}

function h1(text) {
  return new Paragraph({
    text,
    heading: HeadingLevel.HEADING_1,
    spacing: { before: 360, after: 120 },
    pageBreakBefore: true,
  });
}

function h2(text) {
  return new Paragraph({
    text,
    heading: HeadingLevel.HEADING_2,
    spacing: { before: 240, after: 80 },
  });
}

function h3(text) {
  return new Paragraph({
    text,
    heading: HeadingLevel.HEADING_3,
    spacing: { before: 160, after: 60 },
  });
}

function para(runs, opts = {}) {
  const children = typeof runs === 'string'
    ? [new TextRun({ text: runs, size: pt(11), color: C.text })]
    : runs;
  return new Paragraph({
    children,
    spacing: { before: 60, after: 100, line: 276 },
    ...opts,
  });
}

function bold(text) {
  return new TextRun({ text, bold: true, size: pt(11), color: C.text });
}

function run(text, opts = {}) {
  return new TextRun({ text, size: pt(11), color: C.text, ...opts });
}

function italic(text) {
  return new TextRun({ text, italics: true, size: pt(11), color: C.text });
}

function bullet(text, level = 0) {
  const indent = { left: convertInchesToTwip(0.25 + level * 0.25), hanging: convertInchesToTwip(0.25) };
  return new Paragraph({
    children: typeof text === 'string' ? [run(text)] : text,
    bullet: { level },
    indent,
    spacing: { before: 40, after: 40 },
  });
}

function note(text) {
  return new Paragraph({
    children: [
      new TextRun({ text: 'Note: ', bold: true, size: pt(10.5), color: C.subtle }),
      new TextRun({ text, size: pt(10.5), color: C.subtle }),
    ],
    shading: { type: ShadingType.SOLID, color: C.headerBg, fill: C.headerBg },
    border: {
      left: { style: BorderStyle.SINGLE, size: 12, color: C.brand },
    },
    spacing: { before: 80, after: 80 },
    indent: { left: convertInchesToTwip(0.15) },
  });
}

function spacer() {
  return new Paragraph({ text: '', spacing: { before: 0, after: 80 } });
}

// ─── table builder ────────────────────────────────────────────────────────────
function buildTable(headers, rows, colWidths) {
  const totalWidth = 9360; // twips (6.5 inches)
  const widths = colWidths
    ? colWidths.map(w => Math.round(w * totalWidth))
    : headers.map(() => Math.round(totalWidth / headers.length));

  const makeCell = (content, isHeader = false, width) => {
    const children = Array.isArray(content) ? content : [
      new Paragraph({
        children: [new TextRun({
          text: String(content),
          size: pt(10),
          bold: isHeader,
          color: isHeader ? C.white : C.text,
        })],
        spacing: { before: 40, after: 40 },
      }),
    ];
    return new TableCell({
      children,
      width: width ? { size: width, type: WidthType.DXA } : undefined,
      shading: isHeader ? { type: ShadingType.SOLID, color: C.brand, fill: C.brand } : undefined,
      margins: { top: 60, bottom: 60, left: 80, right: 80 },
    });
  };

  const headerRow = new TableRow({
    children: headers.map((h, i) => makeCell(h, true, widths[i])),
    tableHeader: true,
  });

  const dataRows = rows.map((row, ri) =>
    new TableRow({
      children: row.map((cell, i) => {
        const cellEl = makeCell(cell, false, widths[i]);
        if (ri % 2 === 1) {
          cellEl.CellProperties = { shading: { type: ShadingType.SOLID, color: C.rowAlt, fill: C.rowAlt } };
        }
        return cellEl;
      }),
    })
  );

  return new Table({
    rows: [headerRow, ...dataRows],
    width: { size: 100, type: WidthType.PERCENTAGE },
    borders: {
      top:           { style: BorderStyle.SINGLE, size: 1, color: C.border },
      bottom:        { style: BorderStyle.SINGLE, size: 1, color: C.border },
      left:          { style: BorderStyle.SINGLE, size: 1, color: C.border },
      right:         { style: BorderStyle.SINGLE, size: 1, color: C.border },
      insideH:       { style: BorderStyle.SINGLE, size: 1, color: C.border },
      insideV:       { style: BorderStyle.SINGLE, size: 1, color: C.border },
    },
    margins: { top: convertInchesToTwip(0.05), bottom: convertInchesToTwip(0.05) },
  });
}

// ─── title page ───────────────────────────────────────────────────────────────
function titlePage() {
  return [
    new Paragraph({
      children: [new TextRun({ text: 'SharePoint Smart Permissions', size: pt(28), bold: true, color: C.brand })],
      alignment: AlignmentType.CENTER,
      spacing: { before: convertInchesToTwip(1.5), after: 160 },
    }),
    new Paragraph({
      children: [new TextRun({ text: 'User Guide', size: pt(20), color: C.text })],
      alignment: AlignmentType.CENTER,
      spacing: { before: 0, after: 80 },
    }),
    new Paragraph({
      children: [new TextRun({ text: 'Version 1.0', size: pt(12), color: C.subtle })],
      alignment: AlignmentType.CENTER,
      spacing: { before: 0, after: 40 },
    }),
    new Paragraph({
      children: [new TextRun({ text: 'Applies to: SharePoint Online', size: pt(12), color: C.subtle })],
      alignment: AlignmentType.CENTER,
      spacing: { before: 0, after: convertInchesToTwip(2) },
    }),
    img('01_home.png', 580),
    new Paragraph({ children: [new PageBreak()] }),
  ];
}

// ─── document body ────────────────────────────────────────────────────────────
function body() {
  return [

    // ── 1. Overview ──────────────────────────────────────────────────────────
    h1('Overview'),
    para([
      run('SharePoint Smart Permissions is a browser-based auditing tool built directly into SharePoint Online as a web part. It gives site owners, administrators, and compliance teams a clear, real-time view of '),
      bold('who has access to what'),
      run(' — without requiring PowerShell, third-party software, or IT assistance.'),
    ]),
    para("SharePoint's default interface makes it difficult to understand the full picture of permissions across a site. Unique permission breaks are hidden deep in menus, group memberships are opaque, and there is no built-in way to ask 'what can this specific user actually see?' SharePoint Smart Permissions solves all three problems from a single, easy-to-use interface."),
    spacer(),
    h2('What You Can Do'),
    buildTable(
      ['Tool', 'Purpose'],
      [
        ['Permissions Report',    'Generate a full Excel report of every unique permission assignment across a site or the entire tenant'],
        ['Permissions Explorer',  'Interactively browse a document library and inspect permissions on any folder or file in real time'],
        ['User Access',           'Look up any user and see every location they can access, with their exact permission level at each location'],
      ],
      [0.25, 0.75]
    ),

    // ── 2. Who Is This For? ───────────────────────────────────────────────────
    h1('Who Is This For?'),
    bullet([bold('Site Owners'), run(' who want to understand and clean up permissions on their sites')]),
    bullet([bold('IT Administrators'), run(' performing periodic access reviews or compliance audits')]),
    bullet([bold('Compliance Officers'), run(' who need documentation of who has access to sensitive content')]),
    bullet([bold('Help Desk Staff'), run(' diagnosing "why can\'t this user see this file?" questions')]),
    bullet([bold('Security Teams'), run(' identifying over-privileged accounts or verifying least-privilege access')]),
    spacer(),
    note('The web part runs as the currently signed-in user. You can only see sites and content that your account has permission to view. To audit an entire tenant, use an account with read access across all sites.'),

    // ── 3. Getting Started ────────────────────────────────────────────────────
    h1('Getting Started'),
    h2('Prerequisites'),
    bullet([run('You must have at least '), bold('Read'), run(' access to the site you want to audit')]),
    bullet([run('To scan all libraries, you should have '), bold('Site Owner'), run(' or '), bold('Site Collection Administrator'), run(' access')]),
    bullet([run('To run a full tenant scan, you need read access across all site collections (typically a Global Admin or SharePoint Admin account)')]),
    spacer(),
    h2('Accessing the Web Part'),
    para('The SharePoint Smart Permissions web part is added to a SharePoint page by a site administrator. Once added, simply navigate to the page where it has been placed.'),
    para('When you first open the web part, you will see the Home Screen with three options. The web part automatically connects to the current SharePoint site — no configuration is needed to get started.'),
    img('01_home.png'),

    // ── 4. The Home Screen ────────────────────────────────────────────────────
    h1('The Home Screen'),
    para([
      run('The home screen is your starting point. It presents the three tools as cards, each with a brief description of its purpose and a button to launch it.'),
    ]),
    para([
      run('Click any button to enter that tool. You can always return to the home screen using the '),
      bold('Back'),
      run(' button at the top left of any tool screen.'),
    ]),

    // ── 5. Permissions Report ─────────────────────────────────────────────────
    h1('Permissions Report'),
    h2('What It Does'),
    para([
      run('The Permissions Report scans a site\'s libraries, folders, and files and produces a comprehensive view of every location where permissions differ from the site default. It focuses on '),
      bold('unique permission breaks'),
      run(' — places where someone has explicitly changed who can access a specific item.'),
    ]),
    para([
      run('Once the scan is complete, you can export the results as a '),
      bold('colour-coded Excel workbook'),
      run(' that is suitable for sharing with stakeholders or retaining for compliance records.'),
    ]),
    img('02_report_config.png'),
    h2('How to Use It'),
    para([bold('Step 1: '), run('Click '), bold('Run Permissions Report'), run(' from the home screen.')]),
    para([bold('Step 2: '), run('Choose your '), bold('Scan Depth:')]),
    buildTable(
      ['Option', 'What Is Scanned'],
      [
        ['Site only',       'The top-level site permissions only'],
        ['Libraries',       'All document libraries on the site'],
        ['Folders',         'All libraries and folders (configurable depth)'],
        ['Files & Folders', 'Everything — libraries, folders, and individual files'],
      ],
      [0.25, 0.75]
    ),
    spacer(),
    para([bold('Step 3: '), run('If you select '), bold('Folders'), run(', set the '), bold('Folder depth limit'), run(' (1–10 levels deep).')]),
    para([bold('Step 4: '), run('If you are on the root site and have tenant-wide access, you can enable '), bold('Scan all site collections in this tenant'), run(' to audit the entire organisation.')]),
    para([bold('Step 5: '), run('Click '), bold('Run Report'), run('.')]),
    img('03_report_running.png'),
    para([bold('Step 6: '), run('Wait for the scan to complete. A summary shows the number of objects found and how many have unique permissions.')]),
    para([bold('Step 7: '), run('Click '), bold('Export to Excel'), run(' to download the results.')]),
    img('04_report_complete.png'),
    h2('Understanding the Results'),
    para('The Excel export contains one row per scanned object (site, library, folder, or file). Key columns include:'),
    bullet([bold('Type'), run(' — Site, Library, Folder, or File')]),
    bullet([bold('Name'), run(' — The display name of the object')]),
    bullet([bold('Path'), run(' — The server-relative URL')]),
    bullet([bold('Has Unique Permissions'), run(' — Yes/No indicator')]),
    bullet([bold('Users / Groups'), run(' — Everyone who has been explicitly granted access')]),
    bullet([bold('Permission Level'), run(' — The role assigned (Full Control, Edit, Read, etc.)')]),
    spacer(),
    para('Rows for items with unique permissions are highlighted in the Excel workbook so they stand out immediately.'),
    h2('Cancelling a Long Scan'),
    para([
      run('If a scan is taking too long, click the '),
      bold('Cancel'),
      run(' button that appears next to the progress bar. The scan will stop after the current item completes.'),
    ]),

    // ── 6. Permissions Explorer ───────────────────────────────────────────────
    h1('Permissions Explorer'),
    h2('What It Does'),
    para([
      run('The Permissions Explorer lets you browse a document library interactively — folder by folder, file by file — and see the live permissions on any item instantly. It is ideal for investigating a specific area of a site rather than producing a full report.'),
    ]),
    img('05_explorer.png'),
    h2('How to Use It'),
    bullet([bold('Step 1: '), run('Click '), bold('Open Permissions Explorer'), run(' from the home screen.')]),
    bullet([bold('Step 2: '), run('The web part automatically connects to the site and loads the available document libraries.')]),
    bullet([bold('Step 3: '), run('Use the '), bold('Library'), run(' dropdown to select the library you want to browse.')]),
    bullet([bold('Step 4: '), run('The '), bold('left panel'), run(' shows the folder and file tree. Click any item to select it.')]),
    bullet([bold('Step 5: '), run('The '), bold('right panel'), run(' shows the permissions for the selected item.')]),
    spacer(),
    h2('Understanding the Permission Panel'),
    para('When you select an item, the right panel shows:'),
    bullet([run('A '), bold('Unique permissions'), run(' badge (orange) if the item has its own permission assignment, or an '), bold('Inherited permissions'), run(' badge (grey) if it inherits from a parent')]),
    bullet('A table of every user and group that has access, their type (User, SP Group, Security Group), and their permission level'),
    bullet([run('Colour-coded permission badges: '), bold('red'), run(' for Full Control, '), bold('orange'), run(' for Edit/Contribute, '), bold('green'), run(' for Read/View')]),
    spacer(),
    h2('Expand Group Members'),
    para([
      run('Check '),
      bold('Expand SharePoint group members'),
      run(' to expand each SharePoint group in the permissions table and show the individual users inside it. This is useful when you need to see exactly which people are covered by a group assignment.'),
    ]),
    h2('Show Parent Permissions'),
    para([
      run('For items that inherit permissions, check '),
      bold('Show parent permissions'),
      run(' to immediately see where those permissions come from. The panel will display the permissions of the nearest ancestor that has unique permissions.'),
    ]),
    h2('Finding Unique Permissions Quickly'),
    para([
      run('Folders that contain items with unique permissions deeper in their tree are marked with a '),
      bold('down-arrow indicator (↓)'),
      run(' in the tree. This lets you quickly navigate to the areas of a library where permission breaks exist without having to expand every folder.'),
    ]),

    // ── 7. User Access ────────────────────────────────────────────────────────
    h1('User Access'),
    h2('What It Does'),
    para([
      italic('User Access answers the question: "What can this specific person actually see?"'),
      run(' Select any user on the site and the tool scans every library, folder, and file to find every location they have explicit access — showing their exact permission level at each one.'),
    ]),
    img('06_user_access_scanning.png'),
    h2('How to Use It'),
    bullet([bold('Step 1: '), run('Click '), bold('Check User Access'), run(' from the home screen.')]),
    bullet([bold('Step 2: '), run('The web part loads the list of users on the site.')]),
    bullet([bold('Step 3: '), run('Select a user from the '), bold('Select a user'), run(' dropdown.')]),
    bullet([bold('Step 4: '), run('The scan begins automatically. A progress bar, elapsed timer, and status message show the scan\'s progress.')]),
    bullet([bold('Step 5: '), run('Once complete, a table shows every location the user can access.')]),
    spacer(),
    img('07_user_access_complete.png'),
    h2('Cancelling a Scan'),
    para([
      run('Because User Access scans every folder and file on the site, it can take several minutes on large sites. If you need to stop, click the '),
      bold('Cancel'),
      run(' button while the scan is running. The tool returns whatever results have been found so far.'),
    ]),
    h2('Understanding the Results'),
    buildTable(
      ['Column', 'Description'],
      [
        ['Type',             'Site, Library, Folder, or File'],
        ['Name',             'Display name of the location'],
        ['Path',             'Server-relative URL'],
        ['Permission Level', 'The user\'s effective role at this location'],
      ],
      [0.25, 0.75]
    ),
    spacer(),
    para([
      run('Permission level badges use the same colour coding as the Permissions Explorer ('),
      bold('red'),
      run(' = Full Control, '),
      bold('orange'),
      run(' = Edit, '),
      bold('green'),
      run(' = Read).'),
    ]),
    h2('Full Site Access'),
    para([
      run('If a user has Full Control or Owner-level access at the site level, the tool detects this and displays a '),
      bold('Full Site Access'),
      run(' message instead of listing every individual item — because they can access everything.'),
    ]),
    img('08_user_access_full_site.png'),

    // ── 8. Settings ───────────────────────────────────────────────────────────
    h1('Settings'),
    para([
      run('The '),
      bold('Settings'),
      run(' panel is accessible from the gear icon (⚙) in the top-right corner of the banner on any screen.'),
    ]),
    img('09_settings.png'),
    h2('Include System and Hidden Libraries'),
    para([
      run('When this option is '),
      bold('unchecked'),
      run(' (the default), the Permissions Explorer and User Access tools only show standard document libraries — the ones users typically see and interact with.'),
    ]),
    para([run('When '), bold('checked'), run(', the tools also include system and hidden libraries such as:')]),
    bullet('Style Library'),
    bullet('Form Templates'),
    bullet('Site Assets'),
    bullet('Pages'),
    bullet('Other libraries hidden from default views'),
    spacer(),
    note('This setting applies to the Permissions Explorer and User Access tools. The Permissions Report has its own scan settings that control hidden library inclusion.'),

    // ── 9. Changing the Target Site ───────────────────────────────────────────
    h1('Changing the Target Site'),
    para([
      run('By default, the web part connects to the SharePoint site where it is installed. The connected site URL is always visible in the blue banner at the top of every tool screen.'),
    ]),
    para([bold('To audit a different site:')]),
    bullet([run('Click '), bold('Change URL'), run(' in the banner.')]),
    bullet([run('Type or paste the full URL of the target site (e.g., '), run('https://contoso.sharepoint.com/sites/finance', { font: 'Courier New', size: pt(10) }), run(').')]),
    bullet([run('Press '), bold('Enter'), run(' or click '), bold('Connect'), run('.')]),
    spacer(),
    note('You must have at least Read access to the target site. If your account does not have permission, the connection will fail with an error message.'),

    // ── 10. Security & Privacy ────────────────────────────────────────────────
    h1('Security & Privacy'),
    h2('How It Works'),
    para([
      run('SharePoint Smart Permissions runs entirely inside your browser as a SharePoint web part. It makes direct calls to the '),
      bold('SharePoint REST API'),
      run(' using your signed-in credentials — the same API that SharePoint itself uses.'),
    ]),
    h2('Key Security Properties'),
    buildTable(
      ['Property', 'Detail'],
      [
        ['No elevated permissions', 'The tool uses only your existing access rights. It cannot see anything you cannot already see.'],
        ['Read-only',               'The tool never creates, modifies, or deletes any SharePoint content or permissions. It only reads.'],
        ['No external services',    'All data stays within your Microsoft 365 tenant. Nothing is sent to any external server or third-party service.'],
        ['No data storage',         'Results exist only in your browser session. Closing the tab or navigating away clears everything.'],
        ['Standard authentication', 'Authentication is handled entirely by SharePoint and Microsoft 365. The web part never handles passwords or tokens directly.'],
      ],
      [0.3, 0.7]
    ),
    spacer(),
    h2('What the Tool Can See'),
    para([
      run('The tool can only access sites, libraries, folders, and files that '),
      bold('your account'),
      run(' has permission to view. If you do not have access to a library, it will not appear in results.'),
    ]),
    h2('Compliance Considerations'),
    para([
      run('Because the tool produces no audit trail of its own, access reviews performed with it should be documented separately (e.g., by retaining the exported Excel reports with a date and the name of the reviewer).'),
    ]),

    // ── 11. Frequently Asked Questions ───────────────────────────────────────
    h1('Frequently Asked Questions'),
    ...[
      {
        q: 'Do I need any special permissions to use this tool?',
        a: 'You need at least Read access to the site you want to audit. Some features (like scanning all libraries) work best with Site Owner or Site Collection Administrator access.',
      },
      {
        q: 'Will using this tool change any permissions or affect other users?',
        a: 'No. The tool is entirely read-only. It never modifies any SharePoint settings, permissions, or content.',
      },
      {
        q: 'Why does the User Access scan take several minutes?',
        a: "SharePoint's REST API does not provide a single call that returns all permissions for a user. The tool must inspect each library, folder, and file individually. On a large site with many libraries and deeply nested folders, this can take a significant amount of time. You can cancel at any time and see partial results.",
      },
      {
        q: "Why can't I see some libraries in the Permissions Explorer?",
        a: 'By default, hidden and system libraries are excluded (Style Library, Form Templates, Site Assets, etc.). Enable Include system and hidden libraries in Settings to show them.',
      },
      {
        q: 'What does "Inherited permissions" mean?',
        a: 'SharePoint permissions flow down from parent objects. When an item shows "Inherited permissions," it means it uses the same permissions as its parent folder, library, or site — no unique permission assignment has been made for that specific item.',
      },
      {
        q: 'What is a "unique permission break"?',
        a: 'A unique permission break occurs when a specific item (library, folder, or file) has had its permissions explicitly changed, making them different from the parent. This is also called "breaking inheritance." The Permissions Report and Explorer both identify and highlight these breaks.',
      },
      {
        q: 'Can I scan a different site than the one I\'m on?',
        a: 'Yes. Click Change URL in the banner and enter the URL of any site you have access to.',
      },
      {
        q: 'Can I scan the entire tenant at once?',
        a: 'Yes, but only from the root site (e.g., https://contoso.sharepoint.com) and only if your account has read access across all site collections. Enable Scan all site collections in this tenant in the Permissions Report settings.',
      },
      {
        q: 'The Excel export — can I share it with someone who doesn\'t have SharePoint access?',
        a: 'Yes. The exported Excel file is a standard .xlsx file. It contains a snapshot of permissions at the time of the scan and can be shared freely. It contains no live links to SharePoint.',
      },
      {
        q: 'Why does the tool show "Full Site Access" for some users instead of listing their locations?',
        a: 'If a user has Full Control or Owner-level access at the site level, they can access every item on the site. Rather than listing thousands of rows, the tool detects this and shows the Full Site Access message instead.',
      },
      {
        q: 'Is my data secure when I use this tool?',
        a: 'Yes. The tool communicates only with your own SharePoint environment via the standard Microsoft 365 REST API. No data is sent to any external server. See the Security & Privacy section for full details.',
      },
    ].flatMap(({ q, a }) => [
      new Paragraph({
        children: [new TextRun({ text: q, bold: true, size: pt(11), color: C.text })],
        spacing: { before: 160, after: 60 },
      }),
      new Paragraph({
        children: [new TextRun({ text: a, size: pt(11), color: C.text })],
        spacing: { before: 0, after: 80, line: 276 },
        indent: { left: convertInchesToTwip(0.15) },
      }),
    ]),

    // ── 12. Troubleshooting ───────────────────────────────────────────────────
    h1('Troubleshooting'),
    buildTable(
      ['Symptom', 'Likely Cause', 'Fix'],
      [
        ['"Connection failed" when opening a tool',       'Your account lacks Read access, or the site URL is incorrect',    'Confirm you have at least Read access. Check that the URL has no trailing slash and the domain is correct. Try refreshing the page.'],
        ['Scan completes but shows 0 results',            'Insufficient permissions to read library metadata',               'Try enabling Include system and hidden libraries in Settings.'],
        ['"Export to Excel" button is greyed out',        'Scan has not yet completed',                                      'Run the scan first — the export button activates when the scan finishes.'],
        ['The scan hangs on one library',                 'Very large library with thousands of items',                      'Click Cancel and try a narrower scan scope (e.g., Libraries instead of Files & Folders).'],
        ['"You do not have permission" in browser console', 'Account lacks access to some areas of the site',               'Results for those areas will be omitted. This is expected behaviour — the tool only reports what it can see.'],
      ],
      [0.25, 0.3, 0.45]
    ),
    spacer(),

    // ── Footer note ───────────────────────────────────────────────────────────
    new Paragraph({
      children: [
        new TextRun({
          text: 'SharePoint Smart Permissions is a browser-based utility that runs within your Microsoft 365 environment. It does not store, transmit, or log any data outside of your browser session.',
          italics: true,
          size: pt(10),
          color: C.subtle,
        }),
      ],
      border: { top: { style: BorderStyle.SINGLE, size: 6, color: C.border } },
      spacing: { before: 160, after: 80 },
    }),
  ];
}

// ─── build document ───────────────────────────────────────────────────────────
async function main() {
  const doc = new Document({
    creator: 'SharePoint Smart Permissions',
    title: 'SharePoint Smart Permissions — User Guide',
    description: 'User guide for the SharePoint Smart Permissions web part',
    styles: {
      default: {
        document: {
          run: { font: 'Calibri', size: pt(11), color: C.text },
        },
        heading1: {
          run: { font: 'Calibri', size: pt(18), bold: true, color: C.brand },
          paragraph: { spacing: { before: 360, after: 120 } },
        },
        heading2: {
          run: { font: 'Calibri', size: pt(14), bold: true, color: C.brandDark },
          paragraph: { spacing: { before: 240, after: 80 } },
        },
        heading3: {
          run: { font: 'Calibri', size: pt(12), bold: true, color: C.text },
          paragraph: { spacing: { before: 160, after: 60 } },
        },
      },
    },
    sections: [
      {
        properties: {
          page: {
            margin: {
              top:    convertInchesToTwip(1),
              bottom: convertInchesToTwip(1),
              left:   convertInchesToTwip(1),
              right:  convertInchesToTwip(1),
            },
          },
        },
        headers: {
          default: new Header({
            children: [
              new Paragraph({
                children: [
                  new TextRun({ text: 'SharePoint Smart Permissions — User Guide', size: pt(9), color: C.subtle }),
                ],
                border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: C.border } },
                spacing: { after: 80 },
              }),
            ],
          }),
        },
        footers: {
          default: new Footer({
            children: [
              new Paragraph({
                children: [
                  new TextRun({ text: 'Page ', size: pt(9), color: C.subtle }),
                  new TextRun({ children: [PageNumber.CURRENT], size: pt(9), color: C.subtle }),
                  new TextRun({ text: ' of ', size: pt(9), color: C.subtle }),
                  new TextRun({ children: [PageNumber.TOTAL_PAGES], size: pt(9), color: C.subtle }),
                ],
                alignment: AlignmentType.RIGHT,
                border: { top: { style: BorderStyle.SINGLE, size: 4, color: C.border } },
                spacing: { before: 80 },
              }),
            ],
          }),
        },
        children: [...titlePage(), ...body()],
      },
    ],
  });

  const buffer = await Packer.toBuffer(doc);
  fs.writeFileSync(OUT, buffer);
  console.log(`Written: ${OUT}  (${(buffer.length / 1024).toFixed(0)} KB)`);
}

main().catch(err => { console.error(err); process.exit(1); });
