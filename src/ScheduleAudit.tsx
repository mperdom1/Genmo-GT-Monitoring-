import React, { useMemo, useState } from 'react';
import * as XLSX from 'xlsx';
import { AlertTriangle, ClipboardCopy, Download, FileDiff, Upload } from 'lucide-react';

type AuditChange = { sheet: string; employee: string; field: string; before: string; after: string; category: string };
type ParsedWorkbook = { fileName: string; sheets: Record<string, string[][]> };

const clean = (value: unknown) => String(value ?? '').replace(/\s+/g, ' ').trim();
const norm = (value: string) => value.normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase().replace(/[^a-z0-9]/g, '');
const isOff = (value: string) => /^(off|rest|rest day|day off)$/i.test(value);
const isLeave = (value: string) => /^(vac|vacation|vacation leave|igss|leave|medical leave|mat|eml)$/i.test(value);

async function readWorkbook(file: File): Promise<ParsedWorkbook> {
  const buffer = await file.arrayBuffer();
  const wb = XLSX.read(buffer, { type: 'array', cellDates: true });
  const sheets: Record<string, string[][]> = {};
  wb.SheetNames.forEach((name) => {
    sheets[name] = XLSX.utils.sheet_to_json(wb.Sheets[name], { header: 1, raw: false, defval: '' }) as string[][];
  });
  return { fileName: file.name, sheets };
}

function findHeaderIndex(rows: string[][]) {
  return rows.findIndex((row) => row.some((v) => /^(name|employee name|full name|agent name|employee)$/i.test(clean(v))) &&
    row.some((v) => /mon|monday|date|schedule|shift/i.test(clean(v))));
}

function headerLabel(value: string, index: number) {
  return clean(value) || `Column ${index + 1}`;
}

function buildChanges(oldBook: ParsedWorkbook, newBook: ParsedWorkbook): AuditChange[] {
  const changes: AuditChange[] = [];
  const oldSheetByNorm = new Map(Object.keys(oldBook.sheets).map((name) => [norm(name), name]));
  const matchedNewSheets = new Set<string>();

  Object.entries(newBook.sheets).forEach(([newSheetName, newRows]) => {
    const oldSheetName = oldSheetByNorm.get(norm(newSheetName));
    if (!oldSheetName) {
      changes.push({ sheet: newSheetName, employee: '—', field: 'Sheet / week', before: 'Not present', after: 'Added in new file', category: 'Week added' });
      return;
    }
    matchedNewSheets.add(norm(oldSheetName));
    const oldRows = oldBook.sheets[oldSheetName];
    const oldHeaderIndex = findHeaderIndex(oldRows);
    const newHeaderIndex = findHeaderIndex(newRows);
    if (oldHeaderIndex < 0 || newHeaderIndex < 0) {
      // Fallback: compare whole rows when the workbook does not have a recognizable schedule header.
      const rowKey = (r: string[]) => norm(clean(r[0])) || r.map(norm).join('|');
      const oldMap = new Map(oldRows.map((r, i) => [rowKey(r), { r, i }]));
      newRows.forEach((row, i) => {
        const key = rowKey(row);
        const prior = oldMap.get(key);
        if (!prior && row.some((v) => clean(v))) changes.push({ sheet: newSheetName, employee: clean(row[0]) || `Row ${i + 1}`, field: 'Row', before: 'Not present', after: row.map(clean).filter(Boolean).join(' | '), category: 'Row added' });
        else if (prior && prior.r.join('|') !== row.join('|')) {
          const max = Math.max(prior.r.length, row.length);
          for (let c = 0; c < max; c++) if (clean(prior.r[c]) !== clean(row[c])) changes.push({ sheet: newSheetName, employee: clean(row[0]) || `Row ${i + 1}`, field: `Column ${c + 1}`, before: clean(prior.r[c]) || '(blank)', after: clean(row[c]) || '(blank)', category: 'Schedule value changed' });
        }
      });
      return;
    }

    const oldHeaders = oldRows[oldHeaderIndex].map(headerLabel);
    const newHeaders = newRows[newHeaderIndex].map(headerLabel);
    const oldNameIdx = oldHeaders.findIndex((h) => /^(name|employee name|full name|agent name|employee)$/i.test(clean(h)));
    const newNameIdx = newHeaders.findIndex((h) => /^(name|employee name|full name|agent name|employee)$/i.test(clean(h)));
    const keyFor = (row: string[], nameIdx: number) => norm(clean(row[nameIdx]));
    const oldMap = new Map<string, { row: string[]; index: number }>();
    oldRows.slice(oldHeaderIndex + 1).forEach((row, i) => {
      const key = keyFor(row, oldNameIdx);
      if (key) oldMap.set(key, { row, index: i + oldHeaderIndex + 1 });
    });
    const seen = new Set<string>();
    newRows.slice(newHeaderIndex + 1).forEach((row, i) => {
      const employee = clean(row[newNameIdx]);
      const key = keyFor(row, newNameIdx);
      if (!key) return;
      seen.add(key);
      const prior = oldMap.get(key);
      if (!prior) {
        changes.push({ sheet: newSheetName, employee, field: 'Roster', before: 'Not present', after: 'Added in new version', category: 'Agent added' });
        return;
      }
      const oldCol = new Map(oldHeaders.map((h, idx) => [norm(h), idx]));
      newHeaders.forEach((header, colIdx) => {
        const oldIdx = oldCol.get(norm(header));
        if (oldIdx === undefined) return;
        const before = clean(prior.row[oldIdx]);
        const after = clean(row[colIdx]);
        if (before === after) return;
        const label = clean(header);
        const category = isLeave(after) && !isLeave(before) ? 'Leave added/changed'
          : isOff(before) !== isOff(after) ? 'Day off changed'
          : /break|lunch/i.test(label) && /off/i.test(after) ? 'Break/lunch marked OFF'
          : /break|lunch/i.test(label) ? 'Break/lunch changed'
          : /in|out|shift|schedule|time/i.test(label) ? 'Shift time changed'
          : 'Schedule value changed';
        changes.push({ sheet: newSheetName, employee, field: label, before: before || '(blank)', after: after || '(blank)', category });
      });
    });
    oldMap.forEach(({ row }, key) => {
      if (!seen.has(key)) changes.push({ sheet: newSheetName, employee: clean(row[oldNameIdx]), field: 'Roster', before: 'Present in previous version', after: 'Not present in new version', category: 'Agent removed' });
    });
  });

  Object.keys(oldBook.sheets).forEach((oldSheetName) => {
    if (!matchedNewSheets.has(norm(oldSheetName)) && !Object.keys(newBook.sheets).some((name) => norm(name) === norm(oldSheetName))) {
      changes.push({ sheet: oldSheetName, employee: '—', field: 'Sheet / week', before: 'Present in previous file', after: 'Not present in new file', category: 'Week removed' });
    }
  });
  return changes;
}

function csvCell(value: unknown) {
  const text = String(value ?? '');
  return /[",\n]/.test(text) ? `"${text.replace(/"/g, '""')}"` : text;
}

export default function ScheduleAudit() {
  const [oldBook, setOldBook] = useState<ParsedWorkbook | null>(null);
  const [newBook, setNewBook] = useState<ParsedWorkbook | null>(null);
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState('');
  const [oldDate, setOldDate] = useState('Oct 5');
  const [newDate, setNewDate] = useState('Oct 12');
  const changes = useMemo(() => oldBook && newBook ? buildChanges(oldBook, newBook) : [], [oldBook, newBook]);
  const retro = changes.filter((c) => /leave|day off|shift time|schedule value|break\/lunch/i.test(c.category));
  const employees = new Set(retro.map((c) => c.employee).filter((n) => n !== '—'));
  const agentDays = new Set(retro.map((c) => `${norm(c.employee)}|${norm(c.sheet)}|${norm(c.field)}`));
  const counts = (category: string) => changes.filter((c) => c.category === category).length;

  const loadFile = async (which: 'old' | 'new', file?: File) => {
    if (!file) return;
    setBusy(true); setError('');
    try {
      const parsed = await readWorkbook(file);
      which === 'old' ? setOldBook(parsed) : setNewBook(parsed);
    } catch (e: any) {
      setError(e?.message || 'Could not read this workbook.');
    } finally { setBusy(false); }
  };

  const emailDraft = useMemo(() => {
    if (!oldBook || !newBook) return '';
    const byCategory = new Map<string, AuditChange[]>();
    changes.forEach((c) => byCategory.set(c.category, [...(byCategory.get(c.category) || []), c]));
    const sections = Array.from(byCategory.entries()).map(([category, rows]) => {
      const lines = rows.slice(0, 18).map((c) => `- ${c.employee}: ${c.field} — ${c.before} → ${c.after} (${c.sheet})`);
      if (rows.length > 18) lines.push(`- ${rows.length - 18} additional change(s) are listed in the attached audit.`);
      return `**${category} (${rows.length})**\n${lines.join('\n')}`;
    });
    return `Hi Julio,\n\nThanks for sharing the updated schedules. I compared the new file (${newDate}) with the previous version (${oldDate}) and found changes that I would like to validate before loading it. This schedule is the official reference used to measure GT attendance and adherence, so changes to completed weeks may affect historical results.\n\n**Comparison summary**\n- Previous file: ${oldBook.fileName}\n- New file: ${newBook.fileName}\n- Agents with schedule changes: ${employees.size}\n- Changed schedule fields / records: ${changes.length}\n\n${sections.join('\n\n')}\n\nCould you please confirm whether these changes are expected and correct, particularly for weeks that have already ended? For any shift with breaks or lunch marked OFF, please also share the correct times so we can validate the schedule before loading it.\n\nThanks in advance,`;
  }, [oldBook, newBook, changes, oldDate, newDate, employees.size]);

  const downloadAudit = () => {
    const rows = [['Week / Sheet', 'Employee', 'Category', 'Field', 'Previous value', 'New value'], ...changes.map(c => [c.sheet, c.employee, c.category, c.field, c.before, c.after])];
    const blob = new Blob([rows.map(r => r.map(csvCell).join(',')).join('\r\n')], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob); const a = document.createElement('a');
    a.href = url; a.download = 'schedule_version_audit.csv'; a.click(); URL.revokeObjectURL(url);
  };
  const copyEmail = () => navigator.clipboard.writeText(emailDraft).then(() => alert('Email draft copied.')).catch(() => setError('Copy failed. Select and copy the draft manually.'));

  return <section className="bg-white p-5 rounded-xl shadow-sm border border-slate-200 space-y-5">
    <div className="flex items-start gap-3">
      <div className="p-2 bg-amber-50 text-amber-700 rounded-lg"><FileDiff size={22}/></div>
      <div className="flex-1"><h2 className="text-lg font-semibold text-slate-800">Schedule Version Intelligence</h2><p className="text-sm text-slate-500 mt-1">Compare two schedule workbooks, identify historical changes and prepare a validation email before loading the new version.</p></div>
    </div>
    <div className="grid md:grid-cols-2 gap-4">
      <div className="border border-slate-200 rounded-lg p-4 space-y-2">
        <label className="text-sm font-medium text-slate-700">Previous / official version</label>
        <input type="text" value={oldDate} onChange={e => setOldDate(e.target.value)} className="w-full px-3 py-2 border border-slate-300 rounded-md text-sm" aria-label="Previous version date"/>
        <label className="block text-xs text-slate-500">Upload the earlier Excel file</label>
        <input type="file" accept=".xlsx,.xls,.xlsm" onChange={e => loadFile('old', e.target.files?.[0])} className="block w-full text-sm"/>
        {oldBook && <p className="text-xs text-emerald-700">Loaded: {oldBook.fileName}</p>}
      </div>
      <div className="border border-slate-200 rounded-lg p-4 space-y-2">
        <label className="text-sm font-medium text-slate-700">New version to validate</label>
        <input type="text" value={newDate} onChange={e => setNewDate(e.target.value)} className="w-full px-3 py-2 border border-slate-300 rounded-md text-sm" aria-label="New version date"/>
        <label className="block text-xs text-slate-500">Upload the newly received Excel file</label>
        <input type="file" accept=".xlsx,.xls,.xlsm" onChange={e => loadFile('new', e.target.files?.[0])} className="block w-full text-sm"/>
        {newBook && <p className="text-xs text-emerald-700">Loaded: {newBook.fileName}</p>}
      </div>
    </div>
    {error && <p className="text-sm text-red-700 bg-red-50 p-3 rounded-md">{error}</p>}
    {busy && <p className="text-sm text-slate-500">Reading workbook…</p>}
    {oldBook && newBook && <>
      <div className="grid grid-cols-2 lg:grid-cols-4 gap-3">
        <div className="rounded-lg bg-slate-50 border p-3"><p className="text-xs text-slate-500">Changed records</p><p className="text-2xl font-semibold">{changes.length}</p></div>
        <div className="rounded-lg bg-slate-50 border p-3"><p className="text-xs text-slate-500">Agents affected</p><p className="text-2xl font-semibold">{employees.size}</p></div>
        <div className="rounded-lg bg-slate-50 border p-3"><p className="text-xs text-slate-500">Leave changes</p><p className="text-2xl font-semibold">{counts('Leave added/changed')}</p></div>
        <div className="rounded-lg bg-amber-50 border border-amber-200 p-3"><p className="text-xs text-amber-800">Review before loading</p><p className="text-2xl font-semibold text-amber-900">{retro.length}</p></div>
      </div>
      <div className="flex flex-wrap gap-2">
        <button onClick={downloadAudit} className="px-4 py-2 bg-indigo-600 text-white rounded-lg text-sm font-medium"><Download size={16} className="inline mr-2"/>Export comparison CSV</button>
        <button onClick={copyEmail} className="px-4 py-2 bg-slate-100 border rounded-lg text-sm font-medium"><ClipboardCopy size={16} className="inline mr-2"/>Copy email draft</button>
      </div>
      <div className="border border-slate-200 rounded-lg overflow-hidden">
        <div className="p-3 bg-slate-50 border-b flex items-center gap-2"><AlertTriangle size={17} className="text-amber-600"/><h3 className="font-medium">Detected differences</h3></div>
        <div className="overflow-auto max-h-[420px]"><table className="w-full text-sm text-left"><thead className="sticky top-0 bg-white"><tr>{['Week / sheet','Employee','Category','Field','Previous','New'].map(h=><th key={h} className="px-3 py-2 border-b font-medium text-slate-600">{h}</th>)}</tr></thead><tbody>{changes.slice(0,500).map((c,i)=><tr key={i} className="border-b last:border-0"><td className="px-3 py-2">{c.sheet}</td><td className="px-3 py-2 font-medium">{c.employee}</td><td className="px-3 py-2">{c.category}</td><td className="px-3 py-2">{c.field}</td><td className="px-3 py-2 text-slate-500">{c.before}</td><td className="px-3 py-2">{c.after}</td></tr>)}</tbody></table></div>
        {changes.length > 500 && <p className="p-3 text-xs text-slate-500">Showing first 500 changes; the export includes all detected differences.</p>}
        {changes.length === 0 && <p className="p-4 text-sm text-emerald-700">No differences were detected in matching sheet structures.</p>}
      </div>
      <div className="space-y-2"><h3 className="font-semibold text-slate-800">Email draft for Julio</h3><textarea readOnly value={emailDraft} rows={18} className="w-full rounded-lg border border-slate-300 p-3 text-sm font-mono leading-5"/></div>
    </>}
    <p className="text-xs text-slate-400">The audit compares matching sheet names and employee rows. Review the detected differences before relying on the result; this is a validation aid, not an automatic approval to load the schedule.</p>
  </section>;
}
