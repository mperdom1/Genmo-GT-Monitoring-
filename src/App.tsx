import React, { useState } from 'react';
import { ArrowRightLeft, ClipboardCopy, Download, FileSpreadsheet } from 'lucide-react';

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

function tokenizeName(name: string) {
  const stopwords = new Set(['de', 'del', 'la', 'las', 'los', 'y', 'da', 'do']);
  return normalizeName(name)
    .replace(/[^a-z0-9\s]/g, ' ')
    .split(' ')
    .map((t) => t.trim())
    .filter((t) => t && !stopwords.has(t));
}

function normalizeHeader(header: string) {
  return (header || '').toLowerCase().replace(/[^a-z0-9]/g, '');
}

function firstLastKey(tokens: string[]) {
  if (!tokens.length) return '';
  if (tokens.length === 1) return tokens[0];
  return `${tokens[0]} ${tokens[tokens.length - 1]}`;
}

function tokenDistance(a: string, b: string) {
  if (a === b) return 0;
  if (Math.abs(a.length - b.length) > 1) return 99;

  let i = 0;
  let j = 0;
  let edits = 0;

  while (i < a.length && j < b.length) {
    if (a[i] === b[j]) {
      i++;
      j++;
      continue;
    }

    edits++;
    if (edits > 1) return edits;

    if (a.length > b.length) {
      i++;
    } else if (b.length > a.length) {
      j++;
    } else {
      i++;
      j++;
    }
  }

  if (i < a.length || j < b.length) edits++;
  return edits;
}

function tokenMatchesLoose(a: string, b: string) {
  if (a === b) return true;
  if (a.length < 4 || b.length < 4) return false;
  return tokenDistance(a, b) <= 1;
}

function countLooseOverlap(source: string[], target: string[]) {
  let overlap = 0;
  const used = new Set<number>();

  for (const s of source) {
    const idx = target.findIndex((t, i) => !used.has(i) && tokenMatchesLoose(s, t));
    if (idx !== -1) {
      used.add(idx);
      overlap++;
    }
  }

  return overlap;
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
  SUSPENSION: { remark: 'Leave', scheduleType: 'Suspension', isOff: true },
};

function normalizeStatus(value: string) {
  return value.toUpperCase().trim().replace(/\s+/g, ' ');
}

function getStatusLabels(statusRaw: string) {
  const key = normalizeStatus(statusRaw);
  return STATUS_LABELS[key];
}

function isClockTime(value: string) {
  return /^(\d{1,2}:\d{2})(\s?(AM|PM))?$/i.test(value.trim());
}

export default function App() {
  const [scheduleDataStr, setScheduleDataStr] = useState('');
  const [headcountDataStr, setHeadcountDataStr] = useState('');
  const [weekStartDate, setWeekStartDate] = useState('');
  const [campaignName, setCampaignName] = useState('Gen Mobile');
  const [defaultWorkType, setDefaultWorkType] = useState('WFH');
  const [defaultTeam, setDefaultTeam] = useState('GT-Gen Mobile -01');
  const [outputData, setOutputData] = useState<OutputRow[]>([]);

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

      const headerIdx = scheduleData.findIndex((row) => {
        const upper = row.map((cell) => cell.toUpperCase().trim());
        return upper.includes('NAME') && upper.some((c) => c.includes('MON IN'));
      });

      if (headerIdx === -1) {
        alert('No encontre el encabezado. Debe incluir Name, Attendance ID, Mon IN, Mon Out... Sun IN, Sun Out.');
        return;
      }

      const headers = scheduleData[headerIdx].map((h) => h.toUpperCase().trim());
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
      const dayDates = sequentialHeaderDates.length >= 7
        ? sequentialHeaderDates.slice(0, 7)
        : dayDatesByColumn;

      const hasHeaderDates = dayDates.some((d) => d !== null);
      if (!hasHeaderDates && !weekStart) {
        alert('No pude detectar las fechas de la fila superior (ej: 30-Mar, 1-Apr). Ingresa Week Start Date o pega esa fila.');
        return;
      }

      const rosterRows = headcountDataStr.trim() ? parseTSV(headcountDataStr.trim()) : [];
      const rosterByExtension: Record<string, { empId: string; shortName: string }> = {};

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
            rosterByExtension[extension] = { empId, shortName };
          }
        }
      }

      const out: OutputRow[] = [];
      const unmatchedAttendance = new Set<string>();

      for (let rowIdx = headerIdx + 1; rowIdx < scheduleData.length; rowIdx++) {
        const row = scheduleData[rowIdx];
        if (!row || row.length === 0) continue;

        const fullName = (row[nameIdx] || '').trim();
        if (!fullName) continue;

        const attendanceId = (row[attendanceIdIdx] || '').trim();
        const extension = parseVisualCode(attendanceId);
        const rosterMatch = rosterByExtension[extension];

        if (headcountDataStr.trim() && !rosterMatch) {
          unmatchedAttendance.add(attendanceId || fullName);
        }

        const visualCode = rosterMatch?.empId || extension;
        const employeeName = rosterMatch?.shortName
          ? normalizeName(rosterMatch.shortName).replace(/\s+/g, '.')
          : toEmployeeName(fullName);

        for (let day = 0; day < 7; day++) {
          const inRaw = (row[dayCols[day].inIdx] || '').trim();
          const outRaw = (row[dayCols[day].outIdx] || '').trim();

          if (!inRaw && !outRaw) continue;

          const statusRaw = normalizeStatus(inRaw);
          const statusLabels = getStatusLabels(statusRaw);
          const statusOnlyCell = Boolean(statusLabels) && !isClockTime(inRaw);
          const isOff = !inRaw || statusOnlyCell ? (statusLabels?.isOff ?? true) : false;

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
            Paused: 0,
            Remark: statusLabels?.remark ?? (isOff ? 'Rest Day' : 'Present'),
            ScheduleType: statusLabels?.scheduleType ?? (isOff ? 'Rest Day' : ' Regular'),
            WorkType: isOff ? '' : defaultWorkType,
            BreakLunchScheduledDisplay: '',
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
      if (headcountDataStr.trim() && unmatchedAttendance.size > 0) {
        const sample = Array.from(unmatchedAttendance).slice(0, 8).join(', ');
        alert(`Generado: ${out.length} filas. Ojo: ${unmatchedAttendance.size} Attendance no hizo match con headcount. Ejemplos: ${sample}`);
        return;
      }
      alert(`Generado: ${out.length} filas.`);
    } catch (error: any) {
      console.error(error);
      alert(`Ocurrio un error: ${error.message}`);
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
            let val = String((row as any)[h] || '');
            if (val.includes('\n') || val.includes('\t') || val.includes('"')) {
              val = `"${val.replace(/"/g, '""')}"`;
            }
            return val;
          })
          .join('\t')
      ),
    ].join('\n');

    navigator.clipboard.writeText(tsv);
    alert('Copiado al portapapeles.');
  };

  const handleExportCSV = () => {
    if (outputData.length === 0) return;

    const headers = Object.keys(outputData[0]);
    const csv = [
      headers.join(','),
      ...outputData.map((row) =>
        headers
          .map((h) => {
            let val = String((row as any)[h] || '');
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

        <div className="grid grid-cols-1 gap-6">
          <div className="bg-white p-4 rounded-xl shadow-sm border border-slate-200 flex flex-col">
            <label className="flex items-center space-x-2 text-sm font-medium text-slate-700 mb-2">
              <FileSpreadsheet size={18} className="text-indigo-500" />
              <span>Caja 1: Horarios Semanales</span>
            </label>
            <textarea
              className="flex-1 min-h-[260px] p-3 text-sm border border-slate-300 rounded-lg focus:ring-2 focus:ring-indigo-500 focus:border-indigo-500 outline-none resize-y font-mono whitespace-pre"
              placeholder="Pega aqui la tabla de horarios (Name, Attendance ID, Mon IN/OUT ... Sun IN/OUT)"
              value={scheduleDataStr}
              onChange={(e) => setScheduleDataStr(e.target.value)}
            />
          </div>

          <div className="bg-white p-4 rounded-xl shadow-sm border border-slate-200 flex flex-col">
            <label className="flex items-center space-x-2 text-sm font-medium text-slate-700 mb-2">
              <FileSpreadsheet size={18} className="text-indigo-500" />
              <span>Caja 2 (opcional): Headcount (Emp ID, Full Name, Short Name, Extension)</span>
            </label>
            <textarea
              className="flex-1 min-h-[180px] p-3 text-sm border border-slate-300 rounded-lg focus:ring-2 focus:ring-indigo-500 focus:border-indigo-500 outline-none resize-y font-mono whitespace-pre"
              placeholder="Pega aqui el headcount para mapear VisualCode y EmployeeName"
              value={headcountDataStr}
              onChange={(e) => setHeadcountDataStr(e.target.value)}
            />
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
              className="px-3 py-2 border border-slate-300 rounded-lg focus:ring-2 focus:ring-indigo-500 outline-none"
              value={campaignName}
              onChange={(e) => setCampaignName(e.target.value)}
            />
          </div>
          <div className="flex flex-col space-y-1">
            <label className="text-xs font-medium text-slate-500 uppercase tracking-wider">Default Team</label>
            <input
              type="text"
              className="px-3 py-2 border border-slate-300 rounded-lg focus:ring-2 focus:ring-indigo-500 outline-none"
              value={defaultTeam}
              onChange={(e) => setDefaultTeam(e.target.value)}
            />
          </div>
          <div className="flex flex-col space-y-1">
            <label className="text-xs font-medium text-slate-500 uppercase tracking-wider">Default Work Type</label>
            <input
              type="text"
              className="px-3 py-2 border border-slate-300 rounded-lg focus:ring-2 focus:ring-indigo-500 outline-none w-32"
              value={defaultWorkType}
              onChange={(e) => setDefaultWorkType(e.target.value)}
            />
          </div>
          <button
            onClick={handleGenerate}
            className="ml-auto px-6 py-2.5 bg-indigo-600 hover:bg-indigo-700 text-white font-medium rounded-lg shadow-sm transition-colors flex items-center space-x-2"
          >
            <ArrowRightLeft size={18} />
            <span>Transform Data</span>
          </button>
        </div>

        {outputData.length > 0 && (
          <div className="bg-white rounded-xl shadow-sm border border-slate-200 overflow-hidden flex flex-col">
            <div className="p-4 border-b border-slate-200 flex justify-between items-center bg-slate-50">
              <h2 className="font-medium text-slate-800 flex items-center space-x-2">
                <FileSpreadsheet size={18} className="text-slate-500" />
                <span>Generated Output ({outputData.length} rows)</span>
              </h2>
              <div className="flex items-center space-x-2">
                <button
                  onClick={handleCopy}
                  className="px-4 py-2 bg-white border border-slate-300 hover:bg-slate-50 text-slate-700 font-medium rounded-lg shadow-sm transition-colors flex items-center space-x-2 text-sm"
                >
                  <ClipboardCopy size={16} />
                  <span>Copy TSV</span>
                </button>
                <button
                  onClick={handleExportCSV}
                  className="px-4 py-2 bg-emerald-50 border border-emerald-200 hover:bg-emerald-100 text-emerald-700 font-medium rounded-lg shadow-sm transition-colors flex items-center space-x-2 text-sm"
                >
                  <Download size={16} />
                  <span>Export CSV</span>
                </button>
              </div>
            </div>
            <div className="overflow-x-auto">
              <table className="w-full text-sm text-left whitespace-nowrap">
                <thead className="text-xs text-slate-500 uppercase bg-slate-50 border-b border-slate-200">
                  <tr>
                    {Object.keys(outputData[0]).map((key) => (
                      <th key={key} className="px-4 py-3 font-medium tracking-wider">
                        {key}
                      </th>
                    ))}
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-100">
                  {outputData.slice(0, 100).map((row, i) => (
                    <tr key={i} className="hover:bg-slate-50/50">
                      {Object.values(row).map((val: any, j) => (
                        <td key={j} className="px-4 py-2 text-slate-600">
                          {val}
                        </td>
                      ))}
                    </tr>
                  ))}
                </tbody>
              </table>
              {outputData.length > 100 && (
                <div className="p-3 text-center text-sm text-slate-500 bg-slate-50 border-t border-slate-200">
                  Showing first 100 rows. Export or copy to see all {outputData.length} rows.
                </div>
              )}
            </div>
          </div>
        )}
      </div>
    </div>
  );
}
