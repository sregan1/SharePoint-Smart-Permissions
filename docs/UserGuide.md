# SharePoint Smart Permissions — User Guide

**Version 1.0**
**Applies to:** SharePoint Online

---

## Table of Contents

1. [Overview](#overview)
2. [Who Is This For?](#who-is-this-for)
3. [Getting Started](#getting-started)
4. [The Home Screen](#the-home-screen)
5. [Permissions Report](#permissions-report)
6. [Permissions Explorer](#permissions-explorer)
7. [User Access](#user-access)
8. [Settings](#settings)
9. [Changing the Target Site](#changing-the-target-site)
10. [Security & Privacy](#security--privacy)
11. [Frequently Asked Questions](#frequently-asked-questions)
12. [Troubleshooting](#troubleshooting)

---

## Overview

**SharePoint Smart Permissions** is a browser-based auditing tool built directly into SharePoint Online as a web part. It gives site owners, administrators, and compliance teams a clear, real-time view of who has access to what — without requiring PowerShell, third-party software, or IT assistance.

SharePoint's default interface makes it difficult to understand the full picture of permissions across a site. Unique permission breaks are hidden deep in menus, group memberships are opaque, and there is no built-in way to ask "what can this specific user actually see?" SharePoint Smart Permissions solves all three problems from a single, easy-to-use interface.

### What You Can Do

| Tool | Purpose |
|------|---------|
| **Permissions Report** | Generate a full Excel report of every unique permission assignment across a site or the entire tenant |
| **Permissions Explorer** | Interactively browse a document library and inspect permissions on any folder or file in real time |
| **User Access** | Look up any user and see every location they can access, with their exact permission level at each location |

---

## Who Is This For?

- **Site Owners** who want to understand and clean up permissions on their sites
- **IT Administrators** performing periodic access reviews or compliance audits
- **Compliance Officers** who need documentation of who has access to sensitive content
- **Help Desk Staff** diagnosing "why can't this user see this file?" questions
- **Security Teams** identifying over-privileged accounts or verifying least-privilege access

> **Note:** The web part runs as the currently signed-in user. You can only see sites and content that your account has permission to view. To audit an entire tenant, use an account with read access across all sites.

---

## Getting Started

### Prerequisites

- You must have at least **Read** access to the site you want to audit
- To scan all libraries, you should have **Site Owner** or **Site Collection Administrator** access
- To run a full tenant scan, you need read access across all site collections (typically a Global Admin or SharePoint Admin account)

### Accessing the Web Part

The SharePoint Smart Permissions web part is added to a SharePoint page by a site administrator. Once added, simply navigate to the page where it has been placed.

![The web part home screen showing three feature cards](screenshots/01_home.png)

When you first open the web part, you will see the **Home Screen** with three options. The web part automatically connects to the current SharePoint site — no configuration is needed to get started.

---

## The Home Screen

The home screen is your starting point. It presents the three tools as cards, each with a brief description of its purpose and a button to launch it.

![Home screen showing all three feature cards](screenshots/01_home.png)

**What you see:**
- **Permissions Report** — for generating exportable audit reports
- **Permissions Explorer** — for real-time interactive browsing
- **User Access** — for per-user access lookups

Click any button to enter that tool. You can always return to the home screen using the **Back** button at the top left of any tool screen.

---

## Permissions Report

### What It Does

The Permissions Report scans a site's libraries, folders, and files and produces a comprehensive view of every location where permissions differ from the site default. It focuses on **unique permission breaks** — places where someone has explicitly changed who can access a specific item.

Once the scan is complete, you can export the results as a **color-coded Excel workbook** that is suitable for sharing with stakeholders or retaining for compliance records.

![Permissions Report configuration screen with scan depth options](screenshots/02_report_config.png)

### How to Use It

1. Click **Run Permissions Report** from the home screen.
2. Choose your **Scan Depth**:

   | Option | What Is Scanned |
   |--------|----------------|
   | **Site only** | The top-level site permissions only |
   | **Libraries** | All document libraries on the site |
   | **Folders** | All libraries and folders (configurable depth) |
   | **Files & Folders** | Everything — libraries, folders, and individual files |

3. If you select **Folders**, set the **Folder depth limit** (1–10 levels deep).
4. If you are on the root site and have tenant-wide access, you can enable **Scan all site collections in this tenant** to audit the entire organisation.
5. Click **Run Report**.

![Report scan in progress showing progress bar and status](screenshots/03_report_running.png)

6. Wait for the scan to complete. A summary shows the number of objects found and how many have unique permissions.
7. Click **Export to Excel** to download the results.

![Completed report with results table and Export to Excel button](screenshots/04_report_complete.png)

### Understanding the Results

The Excel export contains one row per scanned object (site, library, folder, or file). Key columns include:

- **Type** — Site, Library, Folder, or File
- **Name** — The display name of the object
- **Path** — The server-relative URL
- **Has Unique Permissions** — Yes/No indicator
- **Users / Groups** — Everyone who has been explicitly granted access
- **Permission Level** — The role assigned (Full Control, Edit, Read, etc.)

Rows for items with **unique permissions** are highlighted in the Excel workbook so they stand out immediately.

### Cancelling a Long Scan

If a scan is taking too long, click the **Cancel** button that appears next to the progress bar. The scan will stop after the current item completes.

---

## Permissions Explorer

### What It Does

The Permissions Explorer lets you browse a document library interactively — folder by folder, file by file — and see the live permissions on any item instantly. It is ideal for investigating a specific area of a site rather than producing a full report.

![Permissions Explorer with folder tree on the left and permissions panel on the right](screenshots/05_explorer.png)

### How to Use It

1. Click **Open Permissions Explorer** from the home screen.
2. The web part automatically connects to the site and loads the available document libraries.
3. Use the **Library** dropdown to select the library you want to browse.
4. The **left panel** shows the folder and file tree. Click any item to select it.
5. The **right panel** shows the permissions for the selected item.

![A folder selected in the tree with its permissions table on the right](screenshots/05_explorer.png)

### Understanding the Permission Panel

When you select an item, the right panel shows:

- A **Unique permissions** badge (orange) if the item has its own permission assignment, or an **Inherited permissions** badge (grey) if it inherits from a parent
- A table of every user and group that has access, their type (User, SP Group, Security Group), and their permission level
- Color-coded permission badges: **red** for Full Control, **orange** for Edit/Contribute, **green** for Read/View

### Expand Group Members

Check **Expand SharePoint group members** to expand each SharePoint group in the permissions table and show the individual users inside it. This is useful when you need to see exactly which people are covered by a group assignment.

![Permissions table showing group members and individual users](screenshots/05_explorer.png)

### Show Parent Permissions

For items that inherit permissions, check **Show parent permissions** to immediately see where those permissions come from. The panel will display the permissions of the nearest ancestor that has unique permissions.

### Finding Unique Permissions Quickly

Folders that contain items with unique permissions deeper in their tree are marked with a **down-arrow indicator** (↓) in the tree. This lets you quickly navigate to the areas of a library where permission breaks exist without having to expand every folder.

![Folder tree showing down-arrow indicators on folders containing unique-permission items](screenshots/05_explorer.png)

---

## User Access

### What It Does

User Access answers the question: *"What can this specific person actually see?"* Select any user on the site and the tool scans every library, folder, and file to find every location they have explicit access — showing their exact permission level at each one.

![User Access screen showing user dropdown and completed access list](screenshots/07_user_access_complete.png)

### How to Use It

1. Click **Check User Access** from the home screen.
2. The web part loads the list of users on the site.
3. Select a user from the **Select a user** dropdown.
4. The scan begins automatically. A progress bar, elapsed timer, and status message show the scan's progress.

![User Access scan in progress with progress bar, elapsed time, and cancel button](screenshots/06_user_access_scanning.png)

5. Once complete, a table shows every location the user can access.

### Cancelling a Scan

Because User Access scans every folder and file on the site, it can take several minutes on large sites. If you need to stop, click the **Cancel** button while the scan is running.

### Understanding the Results

The results table shows one row per accessible location:

| Column | Description |
|--------|-------------|
| **Type** | Site, Library, Folder, or File |
| **Name** | Display name of the location |
| **Path** | Server-relative URL |
| **Permission Level** | The user's effective role at this location |

Permission level badges use the same colour coding as the Permissions Explorer (red = Full Control, orange = Edit, green = Read).

### Full Site Access

If a user has Full Control or Owner-level access at the site level, the tool detects this and displays a **Full Site Access** message instead of listing every individual item — because they can access everything.

![Full Site Access message shown for an owner-level user](screenshots/08_user_access_full_site.png)

---

## Settings

The **Settings** panel is accessible from the gear icon (⚙) in the top-right corner of the banner on any screen, or in the top-right of the home screen.

![Settings popover showing the Include system and hidden libraries checkbox](screenshots/09_settings.png)

### Include System and Hidden Libraries

When this option is **unchecked** (the default), the Permissions Explorer and User Access tools only show standard document libraries — the ones users typically see and interact with.

When **checked**, the tools also include system and hidden libraries such as:
- Style Library
- Form Templates
- Site Assets
- Pages
- Other libraries hidden from default views

This setting is useful during a thorough security audit where you need to account for all content, including system-managed locations.

> **Note:** This setting applies to the Permissions Explorer and User Access tools. The Permissions Report has its own scan settings that control hidden library inclusion.

---

## Changing the Target Site

By default, the web part connects to the SharePoint site where it is installed. The connected site URL is always visible in the blue banner at the top of every tool screen.

To audit a **different site**:

1. Click **Change URL** in the banner.
2. Type or paste the full URL of the target site (e.g., `https://contoso.sharepoint.com/sites/finance`).
3. Press **Enter** or click **Connect**.

![Banner in URL edit mode with input field and Connect button](screenshots/09_settings.png)

The web part will reconnect to the new site. All tools will now operate against that site until you change it again or navigate back to the home screen and return.

> **Important:** You must have at least Read access to the target site. If your account does not have permission, the connection will fail with an error message.

---

## Security & Privacy

### How It Works

SharePoint Smart Permissions runs entirely inside your browser as a SharePoint web part. It makes direct calls to the **SharePoint REST API** using your signed-in credentials — the same API that SharePoint itself uses.

### Key Security Properties

| Property | Detail |
|----------|--------|
| **No elevated permissions** | The tool uses only your existing access rights. It cannot see anything you cannot already see. |
| **Read-only** | The tool never creates, modifies, or deletes any SharePoint content or permissions. It only reads. |
| **No external services** | All data stays within your Microsoft 365 tenant. Nothing is sent to any external server or third-party service. |
| **No data storage** | Results exist only in your browser session. Closing the tab or navigating away clears everything. |
| **Standard authentication** | Authentication is handled entirely by SharePoint and Microsoft 365. The web part never handles passwords or tokens directly. |

### What the Tool Can See

The tool can only access sites, libraries, folders, and files that **your account** has permission to view. If you do not have access to a library, it will not appear in results.

### Compliance Considerations

Because the tool produces no audit trail of its own, access reviews performed with it should be documented separately (e.g., by retaining the exported Excel reports with a date and the name of the reviewer).

---

## Frequently Asked Questions

**Q: Do I need any special permissions to use this tool?**
A: You need at least Read access to the site you want to audit. Some features (like scanning all libraries) work best with Site Owner or Site Collection Administrator access.

---

**Q: Will using this tool change any permissions or affect other users?**
A: No. The tool is entirely read-only. It never modifies any SharePoint settings, permissions, or content.

---

**Q: Why does the User Access scan take several minutes?**
A: SharePoint's REST API does not provide a single call that returns all permissions for a user. The tool must inspect each library, folder, and file individually. On a large site with many libraries and deeply nested folders, this can take a significant amount of time. You can cancel at any time and see partial results.

---

**Q: Why can't I see some libraries in the Permissions Explorer?**
A: By default, hidden and system libraries are excluded (Style Library, Form Templates, Site Assets, etc.). Enable **Include system and hidden libraries** in Settings to show them.

---

**Q: What does "Inherited permissions" mean?**
A: SharePoint permissions flow down from parent objects. When an item shows "Inherited permissions," it means it uses the same permissions as its parent folder, library, or site — no unique permission assignment has been made for that specific item.

---

**Q: What is a "unique permission break"?**
A: A unique permission break occurs when a specific item (library, folder, or file) has had its permissions explicitly changed, making them different from the parent. This is also called "breaking inheritance." The Permissions Report and Explorer both identify and highlight these breaks.

---

**Q: Can I scan a different site than the one I'm on?**
A: Yes. Click **Change URL** in the banner and enter the URL of any site you have access to.

---

**Q: Can I scan the entire tenant at once?**
A: Yes, but only from the root site (e.g., `https://contoso.sharepoint.com`) and only if your account has read access across all site collections. Enable **Scan all site collections in this tenant** in the Permissions Report settings.

---

**Q: The Excel export — can I share it with someone who doesn't have SharePoint access?**
A: Yes. The exported Excel file is a standard `.xlsx` file. It contains a snapshot of permissions at the time of the scan and can be shared freely. It contains no live links to SharePoint.

---

**Q: Why does the tool show "Full Site Access" for some users instead of listing their locations?**
A: If a user has Full Control or Owner-level access at the site level, they can access every item on the site. Rather than listing thousands of rows, the tool detects this and shows the Full Site Access message instead.

---

**Q: Is my data secure when I use this tool?**
A: Yes. The tool communicates only with your own SharePoint environment via the standard Microsoft 365 REST API. No data is sent to any external server. See the [Security & Privacy](#security--privacy) section for full details.

---

## Troubleshooting

### "Connection failed" when opening a tool

- Confirm you have at least Read access to the site.
- Check that the site URL is correct (no trailing slash, correct domain).
- Ensure your Microsoft 365 session is still active (try refreshing the page).

### Scan completes but shows 0 results

- You may not have sufficient permissions to read library metadata.
- The site may genuinely have no libraries (unlikely but possible for new sites).
- Try enabling **Include system and hidden libraries** in Settings.

### Export to Excel button is greyed out

- The export button is only active after a scan has completed successfully. Run the scan first.

### The scan seems to hang on one library

- Large libraries with thousands of items can take time. The status message updates as each library is processed.
- If it remains stuck for more than a few minutes, click **Cancel** and try a narrower scan scope (e.g., **Libraries** instead of **Files & Folders**).

### "You do not have permission" errors in the browser console

- Your account does not have sufficient access to some areas of the site. Results for those areas will be omitted. This is expected behaviour — the tool only reports what it can see.

---

*SharePoint Smart Permissions is a browser-based utility that runs within your Microsoft 365 environment. It does not store, transmit, or log any data outside of your browser session.*
