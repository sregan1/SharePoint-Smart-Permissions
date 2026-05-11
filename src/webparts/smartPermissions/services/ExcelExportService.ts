import ExcelJS from 'exceljs';
import { PermissionEntry, ObjectType } from '../models/models';

// Colours matching ExcelExportService.cs
const COLOR = {
  siteFill: 'FF0078D4',
  libraryFill: 'FF00BCF2',
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

function roleColor(roles: string[]): string {
  if (roles.some((r) => r.toLowerCase().includes('full control'))) return COLOR.roleFullControl;
  if (
    roles.some(
      (r) =>
        r.toLowerCase().includes('edit') ||
        r.toLowerCase().includes('contribute') ||
        r.toLowerCase().includes('design'),
    )
  )
    return COLOR.roleEdit;
  if (
    roles.some(
      (r) => r.toLowerCase().includes('read') || r.toLowerCase().includes('view'),
    )
  )
    return COLOR.roleRead;
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
    case ObjectType.Folder:  return COLOR.folderFill;
    case ObjectType.File:    return COLOR.fileFill;
    default:                 return COLOR.fileFill;
  }
}

export class ExcelExportService {
  async export(entries: PermissionEntry[], siteUrl: string): Promise<void> {
    const wb = new ExcelJS.Workbook();
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

      const startRow = rowIndex;
      for (const user of users) {
        this.writeRow(ws, rowIndex, entry, user);
        rowIndex++;
      }

      // Thin bottom border after each entry group
      for (let c = 1; c <= headers.length; c++) {
        const cell = ws.getCell(rowIndex - 1, c);
        cell.border = {
          ...cell.border,
          bottom: { style: 'thin' },
        };
      }
      void startRow; // suppress unused-variable warning
    }

    // Column widths matching the C# original
    ws.getColumn(1).width = 10;
    ws.getColumn(2).width = 55;
    ws.getColumn(3).width = 35;
    ws.getColumn(4).width = 18;
    ws.getColumn(5).width = 35;
    ws.getColumn(6).width = 16;
    ws.getColumn(7).width = 20;

    // Auto-filter on header row
    ws.autoFilter = { from: 'A1', to: 'G1' };
  }

  private writeRow(
    ws: ExcelJS.Worksheet,
    rowIndex: number,
    entry: PermissionEntry,
    user: { displayName: string; principalType: string; roles: string[] } | null,
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
    nameCell.value = entry.name;
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

      row.getCell(6).value = friendlyPrincipalType(user.principalType);
      row.getCell(6).alignment = { vertical: 'middle' };

      const roleCell = row.getCell(7);
      roleCell.value = user.roles.join(', ');
      roleCell.fill = argbFill(roleColor(user.roles));
      roleCell.alignment = { vertical: 'middle' };
    }

    row.commit();
  }
}
