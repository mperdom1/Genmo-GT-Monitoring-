import React, { useState } from 'react';
import { ArrowRightLeft, ClipboardCopy, Download, FileSpreadsheet, Upload } from 'lucide-react';
import * as XLSX from 'xlsx';
import { getDownloadURL, ref as storageRef, uploadBytes } from 'firebase/storage';
import { storage } from './firebase';
import { fetchGoogleSheetData, listGoogleSheetWeeks, saveGoogleSheetWeek, type GoogleSheetRow } from './googleSheets';
import { buildHeadcountTSV, firstDateInRow, sheetToTSV } from './excel';

const MONTH_ABBR = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];

type OutputRow = {
  VisualCode: string;
  Campaign: string;
  Team: string;
  EmployeeName: string;
  Date: string;
  SchIn: string;
  SchOut: string;
  PnchIn: string;
  PnchOut: string;
  Staffed: number;
  Scheduled: number;
  Paused: number;
  Remark: string;
  ScheduleType: string;
  WorkType: string;
  BreakLunchScheduledDisplay: string;
  BreakLunchStaffedDisplay: string;
  BreakLunchRemark: string;
};

function formatDate(dateObj: Date) {
  const m = (dateObj.getMonth() + 1).toString().padStart(2, '0');
  const d = dateObj.getDate().toString().padStart(2, '0');
  const y = dateObj.getFullYear();
  return `${m}-${d}-${y}`;
}

function parseHeaderDateToken(token: string, fallbackYear: number) {
  const raw = (token || '').trim();
  if (!raw) return null;

  const match = raw.match(/^(\d{1,2})[-\s\/](\p{L}{3,})$/u);
  if (!match) return null;

  const day = parseInt(match[1], 10);
  const monthText = match[2].toLowerCase();
  const monthMap: Record<string, number> = {
    jan: 0,
    enero: 0,
    feb: 1,
    febrero: 1,
    mar: 2,
    marzo: 2,
    apr: 3,
    abril: 3,
    may: 4,
    mayo: 4,
    jun: 5,
    junio: 5,
    jul: 6,
    julio: 6,
    aug: 7,
    ago: 7,
    agosto: 7,
    sep: 8,
    sept: 8,
    septiembre: 8,
    oct: 9,
    octubre: 9,
    nov: 10,
    noviembre: 10,
    dec: 11,
    dic: 11,
    diciembre: 11,
  };

  const month = monthMap[monthText];
  if (month === undefined) return null;

  return new Date(fallbackYear, month, day);
}

// Las fechas del encabezado vienen sin año (ej. 31-Dec, 1-Jan). Si el mes baja de un día al
// siguiente, la semana cruzó de año y el resto de fechas pertenece al año siguiente.
function fixYearWrap(dates: (Date | null)[]) {
  let yearOffset = 0;
  let prevMonth = -1;
  return dates.map((d) => {
    if (!d) return null;
    if (prevMonth !== -1 && d.getMonth() < prevMonth) yearOffset++;
    prevMonth = d.getMonth();
    return new Date(d.getFullYear() + yearOffset, d.getMonth(), d.getDate());
  });
}

function normalizeName(name: string) {
  if (!name) return '';
  return name
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .toLowerCase()
    .trim()
    .replace(/\s+/g, ' ');
}

function parseTSV(tsv: string) {
  const rows: string[][] = [];
  let currentRow: string[] = [];
  let currentCell = '';
  let insideQuotes = false;

  for (let i = 0; i < tsv.length; i++) {
    const char = tsv[i];
    const nextChar = tsv[i + 1];

    if (char === '"') {
      if (insideQuotes && nextChar === '"') {
        currentCell += '"';
        i++;
      } else {
        insideQuotes = !insideQuotes;
      }
    } else if (char === '\t' && !insideQuotes) {
      currentRow.push(currentCell);
      currentCell = '';
    } else if (char === '\n' && !insideQuotes) {
      currentRow.push(currentCell);
      rows.push(currentRow);
      currentRow = [];
      currentCell = '';
    } else if (char === '\r' && !insideQuotes) {
      if (nextChar !== '\n') {
        currentRow.push(currentCell);
        rows.push(currentRow);
        currentRow = [];
        currentCell = '';
      }
    } else {
      currentCell += char;
    }
  }

  if (currentRow.length > 0 || currentCell !== '') {
    currentRow.push(currentCell);
    rows.push(currentRow);
  }

  return rows;
}

function parseVisualCode(attendanceIdRaw: string) {
  const match = attendanceIdRaw.match(/\((\d+)\)/);
  return match?.[1] ?? '';
}

function convertTime(timeStr: string) {
  const raw = (timeStr || '').trim();
  if (!raw) return '';

  const upper = raw.toUpperCase();
  if (upper === 'OFF' || upper === 'MAT' || upper === 'VACATION' || upper === 'VAC') {
    return '00:00';
  }

  if (/^\d{1,2}:\d{2}$/.test(raw)) {
    const [h, m] = raw.split(':');
    return `${h.padStart(2, '0')}:${m}`;
  }

  const ampm = raw.match(/^(\d{1,2}):(\d{2})\s*(AM|PM)$/i);
  if (ampm) {
    let h = parseInt(ampm[1], 10);
    const m = ampm[2];
    const period = ampm[3].toUpperCase();
    if (period === 'PM' && h !== 12) h += 12;
    if (period === 'AM' && h === 12) h = 0;
    return `${h.toString().padStart(2, '0')}:${m}`;
  }

  return raw;
}

function calculateMinutes(start: string, end: string) {
  if (!start || !end || (start === '00:00' && end === '00:00')) return 0;

  const [sh, sm] = start.split(':').map(Number);
  const [eh, em] = end.split(':').map(Number);

  if ([sh, sm, eh, em].some((v) => Number.isNaN(v))) return 0;

  let startMins = sh * 60 + sm;
  let endMins = eh * 60 + em;

  if (endMins < startMins) endMins += 24 * 60;

  return endMins - startMins;
}

function toEmployeeName(name: string) {
  const parts = normalizeName(name).split(' ').filter(Boolean);
  if (parts.length >= 3) return `${parts[0]}.${parts[2]}`;
  if (parts.length >= 2) return `${parts[0]}.${parts[1]}`;
  return parts[0] || '';
}


function normalizeHeader(header: string) {
  return (header || '').toLowerCase().replace(/[^a-z0-9]/g, '');
}





const STATUS_LABELS: Record<string, { remark: string; scheduleType: string; isOff: boolean }> = {
  PRESENT: { remark: 'Present', scheduleType: ' Regular', isOff: false },
  ABSENT: { remark: 'Absent', scheduleType: ' Regular', isOff: true },
  LATE: { remark: 'Late', scheduleType: ' Regular', isOff: false },
  UNDERTIME: { remark: 'Undertime', scheduleType: ' Regular', isOff: false },
  'LATE / UNDERTIME': { remark: 'Late / Undertime', scheduleType: ' Regular', isOff: false },
  'LATE/UNDERTIME': { remark: 'Late / Undertime', scheduleType: ' Regular', isOff: false },
  OFF: { remark: 'Rest Day', scheduleType: 'Rest Day', isOff: true },
  REST: { remark: 'Rest Day', scheduleType: 'Rest Day', isOff: true },
  'REST DAY': { remark: 'Rest Day', scheduleType: 'Rest Day', isOff: true },
  VAC: { remark: 'Leave', scheduleType: 'Vacation Leave', isOff: true },
  VACATION: { remark: 'Leave', scheduleType: 'Vacation Leave', isOff: true },
  'VACATION LEAVE': { remark: 'Leave', scheduleType: 'Vacation Leave', isOff: true },
  MAT: { remark: 'Leave', scheduleType: 'Maternity Leave', isOff: true },
  MATERNIDAD: { remark: 'Leave', scheduleType: 'Maternity Leave', isOff: true },
  'MATERNITY LEAVE': { remark: 'Leave', scheduleType: 'Maternity Leave', isOff: true },
  'BIRTHDAY LEAVE': { remark: 'Leave', scheduleType: 'Birthday Leave', isOff: true },
  MED: { remark: 'Leave', scheduleType: 'Extended Medical Leave', isOff: true },
  MEDICAL: { remark: 'Leave', scheduleType: 'Extended Medical Leave', isOff: true },
  'MED LEAVE': { remark: 'Leave', scheduleType: 'Extended Medical Leave', isOff: true },
  'MEDICAL LEAVE': { remark: 'Leave', scheduleType: 'Extended Medical Leave', isOff: true },
  EML: { remark: 'Leave', scheduleType: 'Extended Medical Leave', isOff: true },
  MLEAVE: { remark: 'Leave', scheduleType: 'Extended Medical Leave', isOff: true },
  LEAVE: { remark: 'Leave', scheduleType: 'Leave of Absence', isOff: true },
  'EXTENDED LEAVE': { remark: 'Leave', scheduleType: 'Leave of Absence', isOff: true },
  LOA: { remark: 'Leave', scheduleType: 'Leave of Absence', isOff: true },
  'LEAVE OF ABSENCE': { remark: 'Leave', scheduleType: 'Leave of Absence', isOff: true },
  HOLIDAY: { remark: 'Leave', scheduleType: 'Holiday', isOff: true },
  SUSPENSION: { remark: 'Leave', scheduleType: 'Suspension', isOff: true },
  TERM: { remark: 'Term', scheduleType: 'Term', isOff: true },
  TERMINATED: { remark: 'Term', scheduleType: 'Term', isOff: true },
  TERMINATION: { remark: 'Term', scheduleType: 'Term', isOff: true },
};

const TERMINATION_STATUSES = new Set(['TERM', 'TERMINATED', 'TERMINATION']);

function normalizeStatus(value: string) {
  return value.toUpperCase().trim().replace(/\s+/g, ' ');
}

function getStatusLabels(statusRaw: string) {
  const key = normalizeStatus(statusRaw);
  if (key === 'EML' || /\bMED(?:ICAL)?\b/.test(key)) {
    return { remark: 'Leave', scheduleType: 'Extended Medical Leave', isOff: true };
  }
  return STATUS_LABELS[key];
}

function isClockTime(value: string) {
  return /^(\d{1,2}:\d{2})(\s?(AM|PM))?$/i.test(value.trim());
}


function inferWeekStartFromRows(rows: string[][], fallbackYear = new Date().getFullYear()) {
  for (const row of rows) {
    for (const cell of row) {
      const parsed = parseHeaderDateToken(String(cell || ''), fallbackYear);
      if (parsed) return formatDate(parsed);
    }
  }
  return '';
}

function normalizeGoogleDate(value: unknown) {
  const raw = String(value ?? '').trim();
  if (!raw) return null;
  const direct = new Date(raw);
  if (!Number.isNaN(direct.getTime())) return direct;
  const match = raw.match(/^(\d{1,2})[\\/.-](\d{1,2})[\\/.-](\d{4})$/);
  if (match) {
    const parsed = new Date(Number(match[3]), Number(match[2]) - 1, Number(match[1]));
    return Number.isNaN(parsed.getTime()) ? null : parsed;
  }
  return null;
}

function buildGoogleHeadcountTSV(rows: GoogleSheetRow[], weekDate: string) {
  const target = weekDate ? new Date(`${weekDate}T00:00:00`) : null;
  const validRows = rows.filter((row) => {
    if (String(row.Site ?? '').trim().toUpperCase() !== 'GT') return false;
    if (String(row.Role ?? '').trim().toUpperCase() !== 'AGENT') return false;

    // Si esta fila tiene Last_Day, esta fila no representa al agente activo.
    // Si el mismo agente tiene otra fila con Last_Day en blanco, esa otra fila si puede ser valida.
    const lastDayRaw = String(row.Last_Day ?? '').trim();
    if (lastDayRaw) return false;

    if (!target) return true;
    const start = normalizeGoogleDate(row.Start);
    if (start && start > target) return false;
    return true;
  });
  const lines = validRows.map((row) => [
    String(row['Emp ID'] ?? '').trim(),
    String(row['Full Name'] ?? '').trim(),
    String(row['Short Name'] ?? '').trim(),
    String(row.Extension ?? '').trim(),
  ].join('\t')).filter((line) => line.replace(/\t/g, '').trim());
  return {
    tsv: ['Emp ID\tFull Name\tShort Name\tExtension', ...lines].join('\n'),
    count: lines.length,
  };
}

function normalizeGMInput(scheduleData: string[][], weekStartDate: string) {
  const gmHeaderIdx = scheduleData.findIndex((row) => {
    const upper = row.map((cell) => String(cell || '').toUpperCase().trim());
    return upper.includes('LOB') && upper.includes('GETTY NAME') && upper.includes('ID');
  });
  if (gmHeaderIdx === -1) return null;

  const header = scheduleData[gmHeaderIdx].map((cell) => String(cell || '').trim());
  const lobIdx = header.findIndex((h) => h.toUpperCase() === 'LOB');
  const gettyIdx = header.findIndex((h) => h.toUpperCase() === 'GETTY NAME');
  const idIdx = header.findIndex((h) => h.toUpperCase() === 'ID');
  const firstScheduleCol = Math.max(lobIdx, gettyIdx, idIdx) + 1;
  const days = ['Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat', 'Sun'];
  const groupSize = 8;

  if (header.length < firstScheduleCol + days.length * groupSize) {
    throw new Error('Detecte el formato GM-GT, pero no encontre los 8 campos completos de Mon a Sun.');
  }

  let startDate = weekStartDate || inferWeekStartFromRows(scheduleData.slice(0, gmHeaderIdx));
  if (!startDate) throw new Error('Detecte el formato GM-GT, pero no pude identificar la fecha inicial de la semana.');

  const start = new Date(`${startDate}T00:00:00`);
  const normalizedHeader = ['NAME', 'ATTENDANCE ID'];
  const dateHeader = ['', ''];
  days.forEach((day, dayIndex) => {
    normalizedHeader.push(`${day.toUpperCase()} IN`, `${day.toUpperCase()} OUT`);
    const date = new Date(start);
    date.setDate(start.getDate() + dayIndex);
    const token = `${date.getDate()}-${MONTH_ABBR[date.getMonth()]}`;
    dateHeader.push(token, token);
  });

  const rows: string[][] = [dateHeader, normalizedHeader];
  const pausesByIdDay: Record<string, string> = {};
  const pauseMinutesByIdDay: Record<string, number> = {};

  for (let r = gmHeaderIdx + 1; r < scheduleData.length; r++) {
    const source = scheduleData[r];
    const rawId = String(source[idIdx] || '').trim();
    if (!rawId) continue;

    const row: string[] = ['', rawId];
    for (let dayIndex = 0; dayIndex < days.length; dayIndex++) {
      const base = firstScheduleCol + dayIndex * groupSize;
      const shiftIn = String(source[base] || '').trim();
      const break1Start = String(source[base + 1] || '').trim();
      const break1End = String(source[base + 2] || '').trim();
      const lunchStart = String(source[base + 3] || '').trim();
      const lunchEnd = String(source[base + 4] || '').trim();
      const break2Start = String(source[base + 5] || '').trim();
      const break2End = String(source[base + 6] || '').trim();
      const shiftOut = String(source[base + 7] || '').trim();
      row.push(shiftIn, shiftOut);

      const pauseParts = [
        ['Break 1', break1Start, break1End],
        ['Lunch', lunchStart, lunchEnd],
        ['Break 2', break2Start, break2End],
      ].filter(([, from, to]) => from || to);

      if (pauseParts.length) {
        const key = `${rawId}|${dayIndex}`;
        pausesByIdDay[key] = pauseParts
          .map(([label, from, to]) => `${label}: ${from || ''} - ${to || ''}`)
          .join(' | ');
        pauseMinutesByIdDay[key] = pauseParts.reduce((total, [, from, to]) => {
          if (!from || !to || !isClockTime(from) || !isClockTime(to)) return total;
          return total + calculateMinutes(convertTime(from), convertTime(to));
        }, 0);
      }
    }
    rows.push(row);
  }

  return { rows, pausesByIdDay, pauseMinutesByIdDay, weekStartDate: startDate };
}

export default function App() {
  const [scheduleDataStr, setScheduleDataStr] = useState('');
  const [headcountDataStr, setHeadcountDataStr] = useState('');
  const [weekStartDate, setWeekStartDate] = useState('');
  const [campaignName, setCampaignName] = useState('Gen Mobile');
  const [defaultWorkType, setDefaultWorkType] = useState('WFH');
  const [defaultTeam, setDefaultTeam] = useState('GT-Gen Mobile -01');
  const [outputData, setOutputData] = useState<OutputRow[]>([]);
  const [workbook, setWorkbook] = useState<XLSX.WorkBook | null>(null);
  const [fileInfo, setFileInfo] = useState('');
  const [weekSheets, setWeekSheets] = useState<string[]>([]);
  const [selectedSheet, setSelectedSheet] = useState('');
  const [integrationStatus, setIntegrationStatus] = useState('');
  const [googleWeeks, setGoogleWeeks] = useState<string[]>([]);
  const [selectedGoogleWeek, setSelectedGoogleWeek] = useState('');
  const [reuseWeek, setReuseWeek] = useState('');
  const [googleUserTimelineRows, setGoogleUserTimelineRows] = useState<GoogleSheetRow[]>([]);

  const loadWeek = (wb: XLSX.WorkBook, sheetName: string) => {
    const ws = wb.Sheets[sheetName];
    if (!ws) return;
    const tsv = sheetToTSV(ws);
    const rows = tsv ? parseTSV(tsv) : [];
    const selectedDate = firstDateInRow(ws, 0) || firstDateInRow(ws, 1) ||
      inferWeekStartFromRows(rows.slice(0, 5));

    setSelectedSheet(sheetName);
    setScheduleDataStr(tsv);
    setOutputData([]);
    if (selectedDate) {
      setWeekStartDate(selectedDate);
      if (googleUserTimelineRows.length > 0) {
        const hc = buildGoogleHeadcountTSV(googleUserTimelineRows, selectedDate);
        setHeadcountDataStr(hc.tsv);
        setFileInfo((current) => current.replace(/, Usertimeline cargada desde Google Sheets \([^)]*\)/, `, Usertimeline cargada desde Google Sheets (${hc.count} agentes GT/Agent válidos para la semana)`));
      }
    }
  };

  const handleFile = async (file: File | undefined) => {
    if (!file) return;
    try {
      const buffer = await file.arrayBuffer();
      const wb = XLSX.read(buffer, { type: 'array', cellNF: true });
      const userTimelineName = wb.SheetNames.find((n) => n.trim().toLowerCase() === 'usertimeline');
      const weeks = wb.SheetNames.filter((n) => n !== userTimelineName);

      if (weeks.length === 0) {
        alert('El archivo no tiene hojas de horarios semanales.');
        return;
      }

      let info = `${file.name}: ${weeks.length} semana(s)`;
      if (userTimelineName) {
        const hc = buildHeadcountTSV(wb.Sheets[userTimelineName]);
        if (hc.count > 0) {
          setHeadcountDataStr(hc.tsv);
          info += `, headcount de la hoja ${userTimelineName} (${hc.count} filas)`;
        } else {
          info += `, pero no pude leer el headcount de ${userTimelineName} (revisa Emp ID, Full Name, Short Name, Extension)`;
        }
      } else {
        const selectedWeek = weeks[weeks.length - 1];
        const selectedWs = wb.Sheets[selectedWeek];
        const selectedTsv = selectedWs ? sheetToTSV(selectedWs) : '';
        const selectedRows = selectedTsv ? parseTSV(selectedTsv) : [];
        const selectedDate = firstDateInRow(selectedWs, 0) || firstDateInRow(selectedWs, 1) ||
          inferWeekStartFromRows(selectedRows.slice(0, 5));

        try {
          const googleTimeline = await fetchGoogleSheetData('Usertimeline');
          setGoogleUserTimelineRows(googleTimeline.rows);
          const hc = buildGoogleHeadcountTSV(googleTimeline.rows, selectedDate);
          if (hc.count > 0) {
            setHeadcountDataStr(hc.tsv);
            info += `, Usertimeline cargada desde Google Sheets (${hc.count} agentes GT/Agent válidos para la semana)`;
          } else {
            info += ', sin agentes GT/Agent válidos en Google Usertimeline para la semana';
          }
        } catch (googleError: any) {
          info += ', no se pudo cargar Usertimeline desde Google Sheets';
          console.error(googleError);
          alert(`El GM-GT no trae Usertimeline y no pude cargarla desde Google Sheets: ${googleError?.message || googleError}`);
        }
      }

      setWorkbook(wb);
      setWeekSheets(weeks);
      setFileInfo(info);
      loadWeek(wb, weeks[weeks.length - 1]);
    } catch (error: any) {
      console.error(error);
      alert(`No pude leer el archivo: ${error.message}`);
    }
  };

  const handleGenerate = () => {
    try {
      if (!scheduleDataStr) {
        alert('Pega los horarios en la Caja 1.');
        return;
      }

      const scheduleData = parseTSV(scheduleDataStr.trim());
      if (scheduleData.length === 0) {
        alert('La caja de horarios esta vacia.');
        return;
      }

      let gmPauseMap: Record<string, string> = {};
      let gmPauseMinutesMap: Record<string, number> = {};
      const gmNormalized = normalizeGMInput(scheduleData, weekStartDate);
      if (gmNormalized) {
        scheduleData.splice(0, scheduleData.length, ...gmNormalized.rows);
        gmPauseMap = gmNormalized.pausesByIdDay;
        gmPauseMinutesMap = gmNormalized.pauseMinutesByIdDay;
        if (!weekStartDate && gmNormalized.weekStartDate) setWeekStartDate(gmNormalized.weekStartDate);
      }

      const headerIdx = scheduleData.findIndex((row) => {
        const upper = row.map((cell) => String(cell || '').toUpperCase().trim());
        return upper.includes('NAME') && upper.some((c) => c.includes('MON IN'));
      });

      if (headerIdx === -1) {
        alert('No encontre el encabezado. Debe ser un GM-GT o una tabla Getty con Mon IN/OUT ... Sun IN/OUT.');
        return;
      }

      const headers = scheduleData[headerIdx].map((h) => String(h || '').toUpperCase().trim());
      const nameIdx = headers.findIndex((h) => h === 'NAME');
      const attendanceIdIdx = headers.findIndex((h) => h.includes('ATTENDANCE'));

      const days = ['Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat', 'Sun'];
      const dayCols = days.map((day) => {
        const inIdx = headers.findIndex((h) => h === `${day.toUpperCase()} IN`);
        const outIdx = headers.findIndex((h) => h === `${day.toUpperCase()} OUT`);
        return { day, inIdx, outIdx };
      });

      const hasMissingDayColumns = dayCols.some((d) => d.inIdx === -1 || d.outIdx === -1);
      if (nameIdx === -1 || attendanceIdIdx === -1 || hasMissingDayColumns) {
        alert('Faltan columnas requeridas: Name, Attendance ID y columnas IN/OUT de lunes a domingo.');
        return;
      }

      const fallbackYear = weekStartDate ? Number(weekStartDate.split('-')[0]) : new Date().getFullYear();
      const hasWeekStart = Boolean(weekStartDate);
      const weekStart = hasWeekStart
        ? new Date(
            Number(weekStartDate.split('-')[0]),
            Number(weekStartDate.split('-')[1]) - 1,
            Number(weekStartDate.split('-')[2])
          )
        : null;

      const dateHeaderRow = headerIdx > 0 ? scheduleData[headerIdx - 1] : [];

      const dayDatesByColumn = dayCols.map((dc) => {
        const fromHeader = parseHeaderDateToken((dateHeaderRow?.[dc.inIdx] || '').trim(), fallbackYear);
        return fromHeader;
      });

      const sequentialHeaderDates = (dateHeaderRow || [])
        .map((cell) => parseHeaderDateToken((cell || '').trim(), fallbackYear))
        .filter((d): d is Date => d !== null);

      // Some exports place the first date in col 0 and then leave blanks between days.
      // In that case, mapping by column index shifts dates by +1 day. Prefer sequential mapping.
      const dayDates = fixYearWrap(
        sequentialHeaderDates.length >= 7
          ? sequentialHeaderDates.slice(0, 7)
          : dayDatesByColumn
      );

      const hasHeaderDates = dayDates.some((d) => d !== null);
      if (!hasHeaderDates && !weekStart) {
        alert('No pude detectar las fechas de la fila superior (ej: 30-Mar, 1-Apr). Ingresa Week Start Date o pega esa fila.');
        return;
      }

      const rosterRows = headcountDataStr.trim() ? parseTSV(headcountDataStr.trim()) : [];
      type RosterEntry = { empId: string; shortName: string; fullName: string; extension: string };
      // Una extension puede tener varias personas en el historial (ej. 5199), por eso es una lista.
      const rosterByExtension: Record<string, RosterEntry[]> = {};
      const rosterByName: Record<string, RosterEntry[]> = {};

      if (rosterRows.length > 0) {
        let rHeaderIdx = rosterRows.findIndex((row) =>
          row.some((cell) => {
            const c = cell.toUpperCase();
            return c.includes('EMP ID') || c.includes('FULL NAME') || c.includes('SHORT NAME');
          })
        );
        if (rHeaderIdx === -1) rHeaderIdx = 0;

        const rHeaders = rosterRows[rHeaderIdx] || [];
        const normalizedHeaders = rHeaders.map((h) => normalizeHeader(h));
        const empIdIdx = normalizedHeaders.findIndex((h) => h === 'empid' || h === 'employeeid' || h === 'id');
        const fullNameIdx = normalizedHeaders.findIndex((h) => h === 'fullname' || h === 'name');
        const shortNameIdx = normalizedHeaders.findIndex((h) => h === 'shortname' || h === 'nickname' || h === 'alias');
        const extensionIdx = normalizedHeaders.findIndex((h) => h === 'extension' || h === 'ext');

        if (empIdIdx === -1 || fullNameIdx === -1 || shortNameIdx === -1 || extensionIdx === -1) {
          alert('No pude leer el headcount. Asegurate de pegar columnas: Emp ID, Full Name, Short Name, Extension.');
          return;
        }

        for (let i = rHeaderIdx + 1; i < rosterRows.length; i++) {
          const row = rosterRows[i];
          const empId = String(row[empIdIdx] || '').trim().replace(/\s+/g, '');
          const fullName = (row[fullNameIdx] || '').trim();
          const shortName = (row[shortNameIdx] || '').trim();
          const extension = String(row[extensionIdx] || '').trim().replace(/\D/g, '');

          if (extension && (empId || fullName || shortName)) {
            const entry: RosterEntry = { empId, shortName, fullName, extension };
            const list = (rosterByExtension[extension] ||= []);
            // La misma persona en varias filas del historial: gana la ultima fila.
            const sameIdx = empId ? list.findIndex((e) => e.empId === empId) : -1;
            if (sameIdx >= 0) list.splice(sameIdx, 1);
            list.push(entry);

            const nameKey = normalizeName(fullName);
            if (nameKey) (rosterByName[nameKey] ||= []).push(entry);
          }
        }
      }

      const out: OutputRow[] = [];
      const unmatchedAttendance = new Set<string>();
      const unknownStatuses = new Set<string>();
      const outOnlyRows: string[] = [];
      const noIdRows = new Set<string>();
      const idFromName = new Set<string>();
      const sharedExtensions = new Set<string>();

      const pickRoster = (extension: string, fullName: string): RosterEntry | undefined => {
        const candidates = rosterByExtension[extension] || [];
        if (candidates.length <= 1) return candidates[0];

        const target = normalizeName(fullName).split(' ').filter(Boolean);
        let best = candidates[candidates.length - 1];
        let bestScore = -1;
        candidates.forEach((c) => {
          const tokens = normalizeName(c.fullName || c.shortName).split(' ').filter(Boolean);
          const score = target.filter((t) => tokens.includes(t)).length;
          if (score >= bestScore) {
            best = c;
            bestScore = score;
          }
        });
        if (new Set(candidates.map((c) => c.empId)).size > 1) sharedExtensions.add(extension);
        return best;
      };

      for (let rowIdx = headerIdx + 1; rowIdx < scheduleData.length; rowIdx++) {
        const row = scheduleData[rowIdx];
        if (!row || row.length === 0) continue;

        const attendanceId = (row[attendanceIdIdx] || '').trim();
        const fullName = (row[nameIdx] || '').trim();
        if (!attendanceId && !fullName) continue;

        let extension = parseVisualCode(attendanceId);
        if (!attendanceId) {
          // Fila con nombre pero sin Attendance ID: se intenta ubicar por nombre en el headcount.
          const byName = rosterByName[normalizeName(fullName)];
          if (byName && byName.length > 0) {
            extension = byName[byName.length - 1].extension;
            idFromName.add(fullName);
          } else {
            noIdRows.add(fullName);
            continue;
          }
        }
        const rosterMatch = pickRoster(extension, fullName);

        if (headcountDataStr.trim() && !rosterMatch) {
          unmatchedAttendance.add(attendanceId || fullName);
          if (gmNormalized) continue;
        }

        const visualCode = rosterMatch?.empId || extension;
        const employeeName = rosterMatch?.shortName
          ? normalizeName(rosterMatch.shortName).replace(/\s+/g, '.')
          : toEmployeeName(fullName);

        for (let day = 0; day < 7; day++) {
          const inRaw = (row[dayCols[day].inIdx] || '').trim();
          const outRaw = (row[dayCols[day].outIdx] || '').trim();

          if (!inRaw && !outRaw) continue;

          const inStatus = normalizeStatus(inRaw);
          const outStatus = normalizeStatus(outRaw);
          if (!gmNormalized && (TERMINATION_STATUSES.has(inStatus) || TERMINATION_STATUSES.has(outStatus))) continue;

          const firstCell = inRaw || outRaw;
          const statusRaw = normalizeStatus(firstCell);
          const statusLabels = getStatusLabels(statusRaw);
          const statusOnlyCell = Boolean(statusLabels) && !isClockTime(inRaw);
          // Texto que no es hora ni un estado conocido (ej. "SICK"): no se debe tratar como "Present"
          const unknownStatus = !statusLabels && !isClockTime(firstCell);
          const isOff = unknownStatus ? true : (!inRaw || statusOnlyCell ? (statusLabels?.isOff ?? true) : false);

          if (unknownStatus) unknownStatuses.add(firstCell);
          if (!inRaw && isClockTime(outRaw)) outOnlyRows.push(`${fullName} (${days[day]})`);

          const schIn = isOff ? '00:00' : convertTime(inRaw);
          const schOut = isOff ? '00:00' : convertTime(outRaw);
          const minutes = isOff ? 0 : calculateMinutes(schIn, schOut);

          const date = dayDates[day]
            ? new Date(dayDates[day] as Date)
            : new Date(weekStart as Date);

          if (!dayDates[day]) {
            date.setDate(date.getDate() + day);
          }

          out.push({
            VisualCode: visualCode,
            Campaign: campaignName,
            Team: defaultTeam,
            EmployeeName: employeeName,
            Date: formatDate(date),
            SchIn: schIn,
            SchOut: schOut,
            PnchIn: '',
            PnchOut: '',
            Staffed: minutes,
            Scheduled: minutes,
            Paused: gmPauseMinutesMap[`${attendanceId}|${day}`] || 0,
            Remark: unknownStatus ? 'Check status' : (statusLabels?.remark ?? (isOff ? 'Rest Day' : 'Present')),
            ScheduleType: unknownStatus ? firstCell : (statusLabels?.scheduleType ?? (isOff ? 'Rest Day' : ' Regular')),
            WorkType: isOff ? '' : defaultWorkType,
            BreakLunchScheduledDisplay: gmPauseMap[`${attendanceId}|${day}`] || '',
            BreakLunchStaffedDisplay: '',
            BreakLunchRemark: '',
          });
        }
      }

      if (out.length === 0) {
        alert('No se generaron filas. Verifica que pegaste toda la tabla de horarios.');
        return;
      }

      setOutputData(out);

      const notes: string[] = [`Generado: ${out.length} filas.`];
      if (headcountDataStr.trim() && unmatchedAttendance.size > 0) {
        const sample = Array.from(unmatchedAttendance).slice(0, 8).join(', ');
        notes.push(`Ojo: ${unmatchedAttendance.size} Attendance no hizo match con headcount. Ejemplos: ${sample}`);
      }
      if (noIdRows.size > 0) {
        notes.push(`${noIdRows.size} agente(s) sin Attendance ID y sin match por nombre: NO se incluyeron: ${Array.from(noIdRows).slice(0, 8).join(', ')}`);
      }
      if (idFromName.size > 0) {
        notes.push(`${idFromName.size} agente(s) sin Attendance ID se ubicaron por nombre en el headcount: ${Array.from(idFromName).slice(0, 8).join(', ')}`);
      }
      if (sharedExtensions.size > 0) {
        notes.push(`Extension compartida por varias personas en el headcount (${Array.from(sharedExtensions).join(', ')}); elegi la persona por nombre. Revisa que sea la correcta.`);
      }
      if (unknownStatuses.size > 0) {
        notes.push(`Estados no reconocidos (quedaron como "Check status" en Remark): ${Array.from(unknownStatuses).slice(0, 8).join(', ')}`);
      }
      if (outOnlyRows.length > 0) {
        notes.push(`${outOnlyRows.length} dia(s) con hora de salida pero sin entrada (quedaron como descanso). Ejemplos: ${outOnlyRows.slice(0, 5).join(', ')}`);
      }
      alert(notes.join('\n\n'));
    } catch (error: any) {
      console.error(error);
      alert(`Ocurrio un error: ${error.message}`);
    }
  };

  const handleUploadTemplate = async (file: File | undefined) => {
    if (!file) return;
    try {
      const fileRef = storageRef(storage, `templates/${file.name}`);
      const snapshot = await uploadBytes(fileRef, file);
      const url = await getDownloadURL(snapshot.ref);
      setIntegrationStatus(`✓ Template guardado: ${file.name}`);
      // The file is also loaded into the current transformer below.
      console.info('Firebase template URL:', url);
    } catch (error: any) {
      console.error(error);
      setIntegrationStatus('No se pudo guardar el template en Firebase');
      alert(`No se pudo guardar el Excel en Firebase: ${error?.message || error}`);
    }
  };

  const handleGoogleWeeks = async () => {
    try {
      setIntegrationStatus('Cargando semanas de Google Sheets...');
      const weeks = await listGoogleSheetWeeks();
      setGoogleWeeks(weeks);
      setReuseWeek((current) => current || weeks[weeks.length - 1] || '');
      setIntegrationStatus(`✓ ${weeks.length} semana(s) encontradas en Google Sheets`);
    } catch (error: any) {
      console.error(error);
      setIntegrationStatus('No se pudo leer Google Sheets');
      alert(`No se pudo leer Google Sheets: ${error?.message || error}`);
    }
  };

  const handleReuseWeek = async () => {
    if (!reuseWeek) {
      alert('Selecciona una semana guardada.');
      return;
    }
    if (!weekStartDate) {
      alert('Selecciona primero la fecha de la nueva semana.');
      return;
    }

    try {
      setIntegrationStatus(`Cargando ${reuseWeek} y Usertimeline...`);
      const [source, timeline] = await Promise.all([
        fetchGoogleSheetData(reuseWeek),
        fetchGoogleSheetData('Usertimeline'),
      ]);

      const targetDate = new Date(`${weekStartDate}T00:00:00`);
      const activeAgents = new Map<string, GoogleSheetRow>();

      for (const agent of timeline.rows) {
        const site = String(agent.Site ?? '').trim().toUpperCase();
        const role = String(agent.Role ?? '').trim().toLowerCase();
        const startRaw = String(agent.Start ?? '').trim();
        const lastDayRaw = String(agent.Last_Day ?? '').trim();
        const startDate = startRaw ? normalizeGoogleDate(startRaw) : null;

        if (site !== 'GT' || role !== 'agent') continue;

        // Una fila con Last_Day informado se considera cerrada.
        // Si el mismo agente tiene otra fila con Last_Day en blanco, esa fila si lo mantiene activo.
        if (lastDayRaw) continue;
        if (startDate && startDate > targetDate) continue;

        const key = String(agent['Emp ID'] ?? agent['Getty Username'] ?? agent['Full Name'] ?? '').trim().toLowerCase();
        if (key) activeAgents.set(key, agent);
      }

      const filtered = source.rows.filter((row) => {
        const status = String(row.Remark ?? row.Status ?? '').trim().toUpperCase();
        const type = String(row.ScheduleType ?? '').trim().toUpperCase();
        if (status.includes('TERM') || type.includes('TERM')) return false;

        const keyCandidates = [
          row['VisualCode'],
          row['Emp ID'],
          row['Getty Username'],
          row['EmployeeName'],
          row['Full Name'],
        ].map((v) => String(v ?? '').trim().toLowerCase()).filter(Boolean);

        return keyCandidates.some((key) => activeAgents.has(key));
      });

      // Shift the reused week by 7 days while preserving the schedule/time assigned to each agent.
      const shifted = filtered.map((row) => {
        const copy: GoogleSheetRow = { ...row };
        const rawDate = String(copy.Date ?? '');
        const oldDate = new Date(rawDate);
        if (!Number.isNaN(oldDate.getTime())) {
          const shiftedDate = new Date(oldDate);
          shiftedDate.setDate(shiftedDate.getDate() + 7);
          copy.Date = formatDate(shiftedDate);
        } else {
          copy.Date = weekStartDate;
        }
        return copy;
      });

      setOutputData(shifted as OutputRow[]);
      setIntegrationStatus(`✓ Reutilizada ${reuseWeek}: ${shifted.length} filas limpias para ${weekStartDate}`);
    } catch (error: any) {
      console.error(error);
      setIntegrationStatus('No se pudo reutilizar la semana');
      alert(`No se pudo reutilizar la semana: ${error?.message || error}`);
    }
  };

  const handleSaveWeekToGoogle = async () => {
    if (outputData.length === 0) {
      alert('Primero genera o reutiliza un schedule.');
      return;
    }
    if (!weekStartDate) {
      alert('Selecciona Week Start Date.');
      return;
    }

    try {
      setIntegrationStatus('Guardando semana en Google Sheets...');
      const weekName = formatDate(new Date(`${weekStartDate}T00:00:00`));
      await saveGoogleSheetWeek(weekName, outputData as unknown as GoogleSheetRow[]);
      setIntegrationStatus(`✓ Semana ${weekName} guardada en Google Sheets`);
      await handleGoogleWeeks();
    } catch (error: any) {
      console.error(error);
      setIntegrationStatus('No se pudo guardar la semana en Google Sheets');
      alert(`No se pudo guardar la semana: ${error?.message || error}`);
    }
  };

  const handleCopy = () => {
    if (outputData.length === 0) return;

    const headers = Object.keys(outputData[0]);
    const tsv = [
      headers.join('\t'),
      ...outputData.map((row) =>
        headers
          .map((h) => {
            let val = String((row as any)[h] ?? '');
            if (val.includes('\n') || val.includes('\t') || val.includes('"')) {
              val = `"${val.replace(/"/g, '""')}"`;
            }
            return val;
          })
          .join('\t')
      ),
    ].join('\n');

    navigator.clipboard.writeText(tsv)
      .then(() => alert('Copiado al portapapeles.'))
      .catch(() => alert('No se pudo copiar automaticamente. Usa Export CSV.'));
  };

  const handleExportCSV = () => {
    if (outputData.length === 0) return;

    const headers = Object.keys(outputData[0]);
    const csv = [
      headers.join(','),
      ...outputData.map((row) =>
        headers
          .map((h) => {
            let val = String((row as any)[h] ?? '');
            if (val.includes(',') || val.includes('\n') || val.includes('"')) {
              val = `"${val.replace(/"/g, '""')}"`;
            }
            return val;
          })
          .join(',')
      ),
    ].join('\n');

    const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.setAttribute('href', url);
    link.setAttribute('download', `schedule_export_${formatDate(new Date())}.csv`);
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
  };

  return (
    <div className="min-h-screen bg-slate-50 text-slate-900 p-6 font-sans">
      <div className="max-w-7xl mx-auto space-y-6">
        <header className="flex items-center space-x-3 pb-4 border-b border-slate-200">
          <div className="p-2 bg-indigo-600 text-white rounded-lg">
            <ArrowRightLeft size={24} />
          </div>
          <h1 className="text-2xl font-semibold text-slate-800">Schedule Transformer</h1>
        </header>

        <div className="bg-white p-4 rounded-xl shadow-sm border border-slate-200 space-y-3">
          <label className="flex items-center space-x-2 text-sm font-medium text-slate-700">
            <Upload size={18} className="text-indigo-500" />
            <span>Subir Excel original / nuevo schedule</span>
          </label>
          <input
            type="file"
            accept=".xlsx,.xls"
            className="block w-full text-sm text-slate-600 file:mr-3 file:px-4 file:py-2 file:rounded-lg file:border-0 file:bg-indigo-50 file:text-indigo-700 file:font-medium hover:file:bg-indigo-100"
            onChange={(e) => { const file = e.target.files?.[0]; handleFile(file); handleUploadTemplate(file); }}
          />
          {fileInfo && <p className="text-xs text-slate-500">{fileInfo}</p>}
          {weekSheets.length > 0 && workbook && (
            <div className="flex flex-col space-y-1 max-w-sm">
              <label className="text-xs font-medium text-slate-500 uppercase tracking-wider">Semana (hoja)</label>
              <select
                className="px-3 py-2 border border-slate-300 rounded-lg focus:ring-2 focus:ring-indigo-500 outline-none"
                value={selectedSheet}
                onChange={(e) => loadWeek(workbook, e.target.value)}
              >
                {weekSheets.map((name) => (
                  <option key={name} value={name}>
                    {name.trim()}
                  </option>
                ))}
              </select>
            </div>
          )}
          <p className="text-xs text-slate-400">El Excel es opcional: si ya tienes una semana guardada, puedes reutilizarla desde Google Sheets.</p>
        </div>

        <div className="bg-white p-4 rounded-xl shadow-sm border border-slate-200 space-y-4">
          <div className="flex items-center justify-between gap-4">
            <div>
              <h2 className="font-medium text-slate-800">Semanas guardadas</h2>
              <p className="text-xs text-slate-500">Puedes reutilizar una semana anterior sin volver a subir el Excel.</p>
            </div>
            <button
              onClick={handleGoogleWeeks}
              className="px-4 py-2 bg-slate-100 border border-slate-300 hover:bg-slate-200 text-slate-700 rounded-lg text-sm font-medium"
            >
              Cargar semanas
            </button>
          </div>
          <div className="flex flex-wrap items-end gap-4">
            <div className="flex flex-col space-y-1 min-w-[260px]">
              <label className="text-xs font-medium text-slate-500 uppercase tracking-wider">Semana a reutilizar</label>
              <select
                className="px-3 py-2 border border-slate-300 rounded-lg outline-none"
                value={reuseWeek}
                >
                  <option value="">Selecciona una semana</option>
                  {googleWeeks.map((name) => <option key={name} value={name}>{name}</option>)}
                </select>
              <button
                onClick={handleReuseWeek}
                disabled={!reuseWeek}
                className="px-5 py-2.5 bg-amber-500 hover:bg-amber-600 disabled:opacity-50 text-white font-medium rounded-lg"
              >
                Reutilizar semana
              </button>
            </div>
            <button
              onClick={handleSaveWeekToGoogle}
              disabled={outputData.length === 0}
              className="px-5 py-2.5 bg-emerald-600 hover:bg-emerald-700 disabled:opacity-50 text-white font-medium rounded-lg"
            >
              Guardar semana en Google Sheets
            </button>
          </div>
        </div>

        <div className="bg-white p-5 rounded-xl shadow-sm border border-slate-200 flex flex-wrap items-end gap-4">
          <div className="flex flex-col space-y-1">
            <label className="text-xs font-medium text-slate-500 uppercase tracking-wider">Week Start Date (Monday)</label>
            <input
              type="date"
              className="px-3 py-2 border border-slate-300 rounded-lg focus:ring-2 focus:ring-indigo-500 outline-none"
              value={weekStartDate}
              onChange={(e) => setWeekStartDate(e.target.value)}
            />
          </div>
          <div className="flex flex-col space-y-1">
            <label className="text-xs font-medium text-slate-500 uppercase tracking-wider">Campaign Name</label>
            <input
              type="text"
              className="px-3 py-2 border border-slate-300 rounded-lg outline-none"
              value={campaignName}
              onChange={(e) => setCampaignName(e.target.value)}
            />
          </div>
          <div className="flex flex-col space-y-1">
            <label className="text-xs font-medium text-slate-500 uppercase tracking-wider">Default Team</label>
            <input
              type="text"
              className="px-3 py-2 border border-slate-300 rounded-lg outline-none"
              value={defaultTeam}
              onChange={(e) => setDefaultTeam(e.target.value)}
            />
          </div>
          <div className="flex flex-col space-y-1">
            <label className="text-xs font-medium text-slate-500 uppercase tracking-wider">Default Work Type</label>
            <input
              type="text"
              className="px-3 py-2 border border-slate-300 rounded-lg outline-none w-32"
              value={defaultWorkType}
              onChange={(e) => setDefaultWorkType(e.target.value)}
            />
          </div>
          <button
            onClick={handleGenerate}
            className="ml-auto px-6 py-2.5 bg-indigo-600 hover:bg-indigo-700 text-white font-medium rounded-lg"
          >
            <ArrowRightLeft size={18} className="inline mr-2" />
            Transform Data
          </button>
        </div>

        {outputData.length > 0 && (
          <div className="bg-white rounded-xl shadow-sm border border-slate-200 overflow-hidden">
            <div className="p-4 border-b border-slate-200 flex justify-between items-center bg-slate-50">
              <h2 className="font-medium text-slate-800">Generated Output ({outputData.length} rows)</h2>
              <div className="flex items-center gap-2">
                <button
                  onClick={handleCopy}
                  className="px-4 py-2 bg-white border border-slate-300 text-slate-700 rounded-lg text-sm"
                >
                  <ClipboardCopy size={16} className="inline mr-2" />Copy TSV
                </button>
                <button
                  onClick={handleExportCSV}
                  className="px-4 py-2 bg-emerald-50 border border-emerald-200 text-emerald-700 rounded-lg text-sm"
                >
                  <Download size={16} className="inline mr-2" />Export CSV
                </button>
              </div>
            </div>
            <div className="overflow-x-auto">
              <table className="w-full text-sm text-left whitespace-nowrap">
                <thead className="text-xs text-slate-500 uppercase bg-slate-50 border-b border-slate-200">
                  <tr>
                    {Object.keys(outputData[0]).map((key) => (
                      <th key={key} className="px-4 py-3 font-medium tracking-wider">{key}</th>
                    ))}
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-100">
                  {outputData.slice(0, 100).map((row, i) => (
                    <tr key={i}>
                      {Object.values(row).map((val: any, j) => (
                        <td key={j} className="px-4 py-2 text-slate-600">{val}</td>
                      ))}
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>
        )}
      </div>
    </div>
  );
}
