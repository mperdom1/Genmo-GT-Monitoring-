import * as XLSX from 'xlsx';

const MONTH_ABBR = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];

function pad(n: number) {
  return n.toString().padStart(2, '0');
}

/** True si el formato numerico de la celda es de fecha u hora (ej. d-mmm, h:mm AM/PM). */
function isDateFormat(format: unknown): boolean {
  const z = String(format ?? '').replace(/\[[^\]]*\]|"[^"]*"|\\./g, '');
  return /[dmyhs]/i.test(z) && !/general/i.test(z);
}

/** Numero de serie de Excel -> partes de fecha/hora (sin zonas horarias). */
function serialToParts(serial: number) {
  const ms = Math.round((serial * 86400000) / 60000) * 60000;
  const d = new Date(Date.UTC(1899, 11, 30) + ms);
  return { y: d.getUTCFullYear(), m: d.getUTCMonth() + 1, d: d.getUTCDate(), H: d.getUTCHours(), M: d.getUTCMinutes() };
}

function clean(text: string) {
  return text.replace(/[\t\r\n]+/g, ' ').replace(/"/g, "'").trim();
}

/** Convierte una celda a texto con el mismo formato que se obtiene al copiar desde Google Sheets. */
export function cellToText(cell?: XLSX.CellObject): string {
  if (!cell || cell.v === undefined || cell.v === null) return '';

  if (cell.t === 'n' && cell.z && isDateFormat(cell.z)) {
    const serial = cell.v as number;
    const code = serialToParts(serial);
    if (serial < 1) {
      // Solo hora (ej. 8:00 AM)
      const period = code.H >= 12 ? 'PM' : 'AM';
      const hour12 = code.H % 12 || 12;
      return `${hour12}:${pad(code.M)} ${period}`;
    }
    // Fecha: "30-Mar" (el parser de la herramienta entiende este formato)
    return `${code.d}-${MONTH_ABBR[code.m - 1]}`;
  }

  if (cell.t === 'n') return String(cell.v);

  return clean(String(cell.v));
}

/** Año de la primera celda con fecha real en la fila indicada (para el cambio de año). */
export function firstDateInRow(ws: XLSX.WorkSheet, rowIndex: number): string {
  const ref = ws['!ref'];
  if (!ref) return '';
  const range = XLSX.utils.decode_range(ref);
  for (let c = range.s.c; c <= range.e.c; c++) {
    const cell = ws[XLSX.utils.encode_cell({ r: rowIndex, c })];
    if (cell && cell.t === 'n' && cell.z && isDateFormat(cell.z) && (cell.v as number) >= 1) {
      const code = serialToParts(cell.v as number);
      return `${code.y}-${pad(code.m)}-${pad(code.d)}`;
    }
  }
  return '';
}

/** Hoja semanal -> texto separado por tabuladores (igual que pegar en la Caja 1). */
export function sheetToTSV(ws: XLSX.WorkSheet): string {
  const ref = ws['!ref'];
  if (!ref) return '';
  const range = XLSX.utils.decode_range(ref);
  const lines: string[] = [];

  for (let r = range.s.r; r <= range.e.r; r++) {
    const cells: string[] = [];
    for (let c = range.s.c; c <= range.e.c; c++) {
      cells.push(cellToText(ws[XLSX.utils.encode_cell({ r, c })]));
    }
    while (cells.length > 0 && cells[cells.length - 1] === '') cells.pop();
    lines.push(cells.join('\t'));
  }

  while (lines.length > 0 && lines[lines.length - 1] === '') lines.pop();
  return lines.join('\n');
}

/** Hoja Usertimeline -> headcount (Emp ID, Full Name, Short Name, Extension), solo sitio GT si existe. */
export function buildHeadcountTSV(ws: XLSX.WorkSheet): { tsv: string; count: number } {
  const ref = ws['!ref'];
  if (!ref) return { tsv: '', count: 0 };
  const range = XLSX.utils.decode_range(ref);

  const norm = (s: string) => s.toLowerCase().replace(/[^a-z0-9]/g, '');
  let headerRow = -1;
  const col: Record<string, number> = {};

  for (let r = range.s.r; r <= Math.min(range.e.r, range.s.r + 10) && headerRow === -1; r++) {
    const found: Record<string, number> = {};
    for (let c = range.s.c; c <= range.e.c; c++) {
      const key = norm(cellToText(ws[XLSX.utils.encode_cell({ r, c })]));
      if (key === 'empid') found.empId = c;
      if (key === 'fullname') found.fullName = c;
      if (key === 'shortname') found.shortName = c;
      if (key === 'extension') found.extension = c;
      if (key === 'site') found.site = c;
    }
    if (found.empId !== undefined && found.fullName !== undefined && found.shortName !== undefined && found.extension !== undefined) {
      headerRow = r;
      Object.assign(col, found);
    }
  }

  if (headerRow === -1) return { tsv: '', count: 0 };

  const text = (r: number, c: number | undefined) => (c === undefined ? '' : cellToText(ws[XLSX.utils.encode_cell({ r, c })]));

  const all: string[][] = [];
  const gt: string[][] = [];
  for (let r = headerRow + 1; r <= range.e.r; r++) {
    const extension = text(r, col.extension);
    if (!extension) continue;
    const row = [text(r, col.empId), text(r, col.fullName), text(r, col.shortName), extension];
    all.push(row);
    if (text(r, col.site).toUpperCase() === 'GT') gt.push(row);
  }

  const rows = gt.length > 0 ? gt : all;
  const tsv = ['Emp ID\tFull Name\tShort Name\tExtension', ...rows.map((r) => r.join('\t'))].join('\n');
  return { tsv, count: rows.length };
}
