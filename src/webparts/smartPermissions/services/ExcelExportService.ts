// exceljs is the largest dependency in the bundle and is only needed when the
// user actually exports to .xlsx — load it on demand as a separate webpack
// chunk. Only types are imported statically (erased at compile time).
import type * as ExcelJS from 'exceljs';
import { PermissionEntry, ObjectType, UserPermissionInfo } from '../models/models';
import { roleAccessTier } from './sp/spCore';

let excelModulePromise: Promise<typeof ExcelJS> | undefined;
function loadExcelJS(): Promise<typeof ExcelJS> {
  if (!excelModulePromise) {
    excelModulePromise = import(/* webpackChunkName: 'exceljs' */ 'exceljs')
      .then((m: any) => (m.default ?? m) as typeof ExcelJS);
  }
  return excelModulePromise;
}

// Colors matching ExcelExportService.cs
const COLOR = {
  siteFill: 'FF0078D4',
  libraryFill: 'FF00BCF2',
  listFill: 'FF038387',
  folderFill: 'FFFFB900',
  fileFill: 'FF8A8886',
  headerFill: 'FF0078D4',
  headerFont: 'FFFFFFFF',
  uniqueFont: 'FF107C10',
  inheritedFont: 'FF605E5C',
  titleFont: 'FF0078D4',
  roleFullControl: 'FFFDE7E9',
  roleEdit: 'FFFFF4CE',
  roleRead: 'FFDFF6DD',
  roleOther: 'FFF0F0F0',
  folderTypeFont: 'FF323130',
};

function roleColor(roles: string[], roleTypeKinds?: Record<string, number>): string {
  const tiers = roles.map((r) => roleAccessTier(r, roleTypeKinds?.[r]));
  if (tiers.indexOf('admin') !== -1) return COLOR.roleFullControl;
  if (tiers.indexOf('edit') !== -1) return COLOR.roleEdit;
  if (tiers.indexOf('read') !== -1) return COLOR.roleRead;
  return COLOR.roleOther;
}

function friendlyPrincipalType(raw: string): string {
  if (raw === 'SecurityGroup') return 'Security Group';
  if (raw === 'SharePointGroup') return 'SP Group';
  return raw;
}

function argbFill(hex: string): ExcelJS.Fill {
  return { type: 'pattern', pattern: 'solid', fgColor: { argb: hex } };
}

function typeLabel(t: ObjectType): string {
  return t.toString();
}

function typeFillArgb(t: ObjectType): string {
  switch (t) {
    case ObjectType.Site:    return COLOR.siteFill;
    case ObjectType.Library: return COLOR.libraryFill;
    case ObjectType.List:    return COLOR.listFill;
    case ObjectType.Folder:  return COLOR.folderFill;
    case ObjectType.File:    return COLOR.fileFill;
    default:                 return COLOR.fileFill;
  }
}

// Display name with the NoCrawl marker appended, used by all exports.
function entryDisplayName(entry: PermissionEntry): string {
  return entry.noCrawl ? `${entry.name} (hidden from search)` : entry.name;
}

export class ExcelExportService {
  async exportUserAccess(
    entries: PermissionEntry[],
    siteUrl: string,
    userDisplayName: string,
  ): Promise<void> {
    const Excel = await loadExcelJS();
    const wb = new Excel.Workbook();
    const ws = wb.addWorksheet('User Access');

    // Title
    const titleCell = ws.getCell('A1');
    titleCell.value = `User Access Report — ${userDisplayName}`;
    titleCell.font = { bold: true, size: 16, color: { argb: COLOR.titleFont } };

    const meta: [string, string][] = [
      ['Site URL', siteUrl],
      ['User', userDisplayName],
      ['Generated', new Date().toLocaleString()],
      ['Accessible Locations', String(entries.length)],
    ];
    meta.forEach(([label, value], i) => {
      const row = i + 3;
      ws.getCell(row, 1).value = label;
      ws.getCell(row, 1).font = { bold: true };
      ws.getCell(row, 2).value = value;
    });

    // Headers
    const headerRow = ws.getRow(8);
    ['Type', 'Name', 'Path', 'Permission Level'].forEach((h, i) => {
      const cell = headerRow.getCell(i + 1);
      cell.value = h;
      cell.fill = argbFill(COLOR.headerFill);
      cell.font = { bold: true, color: { argb: COLOR.headerFont } };
      cell.alignment = { horizontal: 'center', vertical: 'middle' };
    });
    headerRow.commit();

    ws.views = [{ state: 'frozen', xSplit: 0, ySplit: 8 }];

    entries.forEach((entry, idx) => {
      const row = ws.getRow(9 + idx);
      const typeCell = row.getCell(1);
      typeCell.value = typeLabel(entry.objectType);
      typeCell.fill = argbFill(typeFillArgb(entry.objectType));
      typeCell.font = {
        bold: true,
        color: {
          argb: entry.objectType === ObjectType.Folder ? COLOR.folderTypeFont : COLOR.headerFont,
        },
      };
      typeCell.alignment = { horizontal: 'center', vertical: 'middle' };

      const nameCell = row.getCell(2);
      nameCell.value = entryDisplayName(entry);
      nameCell.alignment = { indent: entry.depth, vertical: 'middle' };

      row.getCell(3).value = entry.serverRelativeUrl;
      row.getCell(3).alignment = { vertical: 'middle' };

      const roles = entry.uniquePermissions[0]?.roles ?? [];
      const roleCell = row.getCell(4);
      roleCell.value = roles.join(', ');
      roleCell.fill = argbFill(roleColor(roles, entry.uniquePermissions[0]?.roleTypeKinds));
      roleCell.alignment = { vertical: 'middle' };

      row.commit();
    });

    ws.getColumn(1).width = 10;
    ws.getColumn(2).width = 40;
    ws.getColumn(3).width = 60;
    ws.getColumn(4).width = 25;
    ws.autoFilter = { from: 'A8', to: 'D8' };

    const buffer = await wb.xlsx.writeBuffer();
    const blob = new Blob([buffer as ArrayBuffer], {
      type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    });
    const ts = new Date().toISOString().replace(/[-:T]/g, '').substring(0, 15).replace('.', '');
    const safeName = userDisplayName.replace(/[^a-zA-Z0-9_-]/g, '_');
    const filename = `SP_UserAccess_${safeName}_${ts}.xlsx`;
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = filename;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    setTimeout(() => URL.revokeObjectURL(url), 2000);
  }

  // ── CSV exports ───────────────────────────────────────────────────────────

  private csvEscape(v: string | number): string {
    const s = String(v);
    return s.includes(',') || s.includes('"') || s.includes('\n')
      ? `"${s.replace(/"/g, '""')}"`
      : s;
  }

  private downloadCsv(rows: string[][], filename: string): void {
    const content = rows.map((r) => r.map((c) => this.csvEscape(c)).join(',')).join('\r\n');
    const blob = new Blob(['﻿' + content], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = filename;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    setTimeout(() => URL.revokeObjectURL(url), 2000);
  }

  exportPermissionsCsv(entries: PermissionEntry[], siteUrl: string): void {
    const ts = new Date().toISOString().replace(/[-:T]/g, '').substring(0, 15).replace('.', '');
    const rows: string[][] = [
      ['Type', 'Path', 'Name', 'Permission Source', 'User / Group', 'Access Via', 'Principal Type', 'Permission Level', 'Site URL'],
    ];
    for (const entry of entries) {
      if (entry.uniquePermissions.length === 0) {
        rows.push([entry.objectType, entry.serverRelativeUrl, entryDisplayName(entry), entry.hasUniquePermissions ? 'Unique' : 'Inherited', '', '', '', '', siteUrl]);
      } else {
        for (const user of entry.uniquePermissions) {
          rows.push([
            entry.objectType,
            entry.serverRelativeUrl,
            entryDisplayName(entry),
            entry.hasUniquePermissions ? 'Unique' : 'Inherited',
            user.displayName,
            user.sourceGroup ?? 'Direct',
            friendlyPrincipalType(user.principalType),
            user.roles.join('; '),
            siteUrl,
          ]);
        }
      }
    }
    this.downloadCsv(rows, `SP_Permissions_${ts}.csv`);
  }

  exportUserAccessCsv(entries: PermissionEntry[], siteUrl: string, userDisplayName: string): void {
    const ts = new Date().toISOString().replace(/[-:T]/g, '').substring(0, 15).replace('.', '');
    const safeName = userDisplayName.replace(/[^a-zA-Z0-9_-]/g, '_');
    const rows: string[][] = [['Type', 'Name', 'Path', 'Permission Level']];
    for (const entry of entries) {
      rows.push([
        entry.objectType,
        entryDisplayName(entry),
        entry.serverRelativeUrl,
        (entry.uniquePermissions[0]?.roles ?? []).join('; '),
      ]);
    }
    this.downloadCsv(rows, `SP_UserAccess_${safeName}_${ts}.csv`);
  }

  async export(entries: PermissionEntry[], siteUrl: string): Promise<void> {
    const Excel = await loadExcelJS();
    const wb = new Excel.Workbook();
    this.addSummarySheet(wb, entries, siteUrl);
    this.addDetailsSheet(wb, entries);

    const buffer = await wb.xlsx.writeBuffer();
    const blob = new Blob([buffer as ArrayBuffer], {
      type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    });

    const ts = new Date()
      .toISOString()
      .replace(/[-:T]/g, '')
      .substring(0, 15)
      .replace('.', '');
    const filename = `SP_Permissions_${ts}.xlsx`;

    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = filename;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    setTimeout(() => URL.revokeObjectURL(url), 2000);
  }

  // ── Summary sheet ─────────────────────────────────────────────────────────

  private addSummarySheet(
    wb: ExcelJS.Workbook,
    entries: PermissionEntry[],
    siteUrl: string,
  ): void {
    const ws = wb.addWorksheet('Summary');

    const title = ws.getCell('A1');
    title.value = 'SharePoint Permissions Report';
    title.font = { bold: true, size: 16, color: { argb: COLOR.titleFont } };

    const data: [string, string | number][] = [
      ['Site URL', siteUrl],
      ['Generated', new Date().toLocaleString()],
      ['Objects Scanned', entries.length],
      ['Unique Permissions', entries.filter((e) => e.hasUniquePermissions).length],
      ['Inherited Permissions', entries.filter((e) => !e.hasUniquePermissions).length],
    ];

    data.forEach(([label, value], i) => {
      const row = i + 3;
      const labelCell = ws.getCell(row, 1);
      labelCell.value = label;
      labelCell.font = { bold: true };
      ws.getCell(row, 2).value = value;
    });

    ws.getColumn(1).width = 24;
    ws.getColumn(2).width = 60;
  }

  // ── Details sheet ─────────────────────────────────────────────────────────

  private addDetailsSheet(wb: ExcelJS.Workbook, entries: PermissionEntry[]): void {
    const ws = wb.addWorksheet('Permissions');

    const headers = [
      'Type',
      'Path',
      'Name',
      'Permission Source',
      'User / Group',
      'Access Via',
      'Principal Type',
      'Permission Level',
    ];

    const headerRow = ws.getRow(1);
    headers.forEach((h, i) => {
      const cell = headerRow.getCell(i + 1);
      cell.value = h;
      cell.fill = argbFill(COLOR.headerFill);
      cell.font = { bold: true, color: { argb: COLOR.headerFont } };
      cell.alignment = { horizontal: 'center', vertical: 'middle' };
    });
    headerRow.commit();

    // Freeze header row
    ws.views = [{ state: 'frozen', xSplit: 0, ySplit: 1 }];

    let rowIndex = 2;

    for (const entry of entries) {
      const users = entry.uniquePermissions;

      if (users.length === 0) {
        this.writeRow(ws, rowIndex, entry, null);
        rowIndex++;
        continue;
      }

      for (const user of users) {
        this.writeRow(ws, rowIndex, entry, user);
        rowIndex++;
      }

      // Thin bottom border after each entry group
      for (let c = 1; c <= 8; c++) {
        const cell = ws.getCell(rowIndex - 1, c);
        cell.border = {
          ...cell.border,
          bottom: { style: 'thin' },
        };
      }
    }

    // Column widths
    ws.getColumn(1).width = 10;
    ws.getColumn(2).width = 55;
    ws.getColumn(3).width = 35;
    ws.getColumn(4).width = 18;
    ws.getColumn(5).width = 35;
    ws.getColumn(6).width = 30;
    ws.getColumn(7).width = 16;
    ws.getColumn(8).width = 20;

    // Auto-filter on header row
    ws.autoFilter = { from: 'A1', to: 'H1' };
  }

  private writeRow(
    ws: ExcelJS.Worksheet,
    rowIndex: number,
    entry: PermissionEntry,
    user: UserPermissionInfo | null,
  ): void {
    const row = ws.getRow(rowIndex);

    // Col 1: Type badge
    const typeCell = row.getCell(1);
    typeCell.value = typeLabel(entry.objectType);
    typeCell.fill = argbFill(typeFillArgb(entry.objectType));
    typeCell.font = {
      bold: true,
      color: {
        argb:
          entry.objectType === ObjectType.Folder
            ? COLOR.folderTypeFont
            : COLOR.headerFont,
      },
    };
    typeCell.alignment = { horizontal: 'center', vertical: 'middle' };

    // Col 2: Path (indented by depth)
    const pathCell = row.getCell(2);
    pathCell.value = entry.serverRelativeUrl;
    pathCell.alignment = { indent: entry.depth, vertical: 'middle' };

    // Col 3: Name
    const nameCell = row.getCell(3);
    nameCell.value = entryDisplayName(entry);
    nameCell.font = {
      bold: entry.objectType === ObjectType.Site,
    };
    nameCell.alignment = { vertical: 'middle' };

    // Col 4: Permission Source
    const sourceCell = row.getCell(4);
    sourceCell.value = entry.hasUniquePermissions ? 'Unique' : 'Inherited';
    sourceCell.font = {
      italic: !entry.hasUniquePermissions,
      color: {
        argb: entry.hasUniquePermissions ? COLOR.uniqueFont : COLOR.inheritedFont,
      },
    };
    sourceCell.alignment = { horizontal: 'center', vertical: 'middle' };

    if (user) {
      row.getCell(5).value = user.displayName;
      row.getCell(5).alignment = { vertical: 'middle' };

      // Col 6: Access Via — group name if expanded from a group, otherwise "Direct"
      row.getCell(6).value = user.sourceGroup ?? 'Direct';
      row.getCell(6).alignment = { vertical: 'middle' };

      row.getCell(7).value = friendlyPrincipalType(user.principalType);
      row.getCell(7).alignment = { vertical: 'middle' };

      const roleCell = row.getCell(8);
      roleCell.value = user.roles.join(', ');
      roleCell.fill = argbFill(roleColor(user.roles, user.roleTypeKinds));
      roleCell.alignment = { vertical: 'middle' };
    }

    row.commit();
  }
}
