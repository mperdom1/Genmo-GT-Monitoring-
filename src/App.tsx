import React, { useState } from 'react';
import { ClipboardCopy, FileSpreadsheet, ArrowRightLeft, Users, Download } from 'lucide-react';

function convertTime(timeStr: string) {
  if (!timeStr) return '00:00';
  const upper = timeStr.trim().toUpperCase();
  if (upper === 'OFF') return '00:00';
  
  // Already in HH:MM or H:MM format
  if (/^\d{1,2}:\d{2}$/.test(timeStr.trim())) {
    const [h, m] = timeStr.trim().split(':');
    return `${h.padStart(2, '0')}:${m}`;
  }
  
  const match = timeStr.match(/(\d+):(\d+)\s*(AM|PM)/i);
  if (!match) return timeStr.trim(); // Fallback
  let hours = parseInt(match[1], 10);
  const minutes = match[2];
  const period = match[3].toUpperCase();
  
  if (period === 'PM' && hours !== 12) hours += 12;
  if (period === 'AM' && hours === 12) hours = 0;
  
  return `${hours.toString().padStart(2, '0')}:${minutes}`;
}

function calculateMinutes(start: string, end: string) {
  if (!start || !end || (start === '00:00' && end === '00:00')) return 0;
  const [sh, sm] = start.split(':').map(Number);
  const [eh, em] = end.split(':').map(Number);
  
  if (isNaN(sh) || isNaN(sm) || isNaN(eh) || isNaN(em)) return 0;
  
  let startMins = sh * 60 + sm;
  let endMins = eh * 60 + em;
  if (endMins < startMins) endMins += 24 * 60;
  return endMins - startMins;
}

function formatDate(dateObj: Date) {
  const m = (dateObj.getMonth() + 1).toString().padStart(2, '0');
  const d = dateObj.getDate().toString().padStart(2, '0');
  const y = dateObj.getFullYear();
  return `${m}-${d}-${y}`;
}

function formatDuration(minutes: number) {
  if (!minutes) return '00:00:00';
  const h = Math.floor(minutes / 60).toString().padStart(2, '0');
  const m = (minutes % 60).toString().padStart(2, '0');
  return `${h}:${m}:00`;
}

function normalizeName(name: string) {
  if (!name) return '';
  return name.normalize("NFD").replace(/[\u0300-\u036f]/g, "").toLowerCase().trim().replace(/\s+/g, ' ');
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

function isSimilar(s1: string, s2: string) {
  if (s1 === s2) return true;
  if (Math.abs(s1.length - s2.length) > 2) return false;
  let common = 0;
  let s2Copy = s2;
  for (let i = 0; i < s1.length; i++) {
    const char = s1[i];
    const idx = s2Copy.indexOf(char);
    if (idx !== -1) {
      common++;
      s2Copy = s2Copy.substring(0, idx) + s2Copy.substring(idx + 1);
    }
  }
  return common / Math.max(s1.length, s2.length) >= 0.8;
}

export default function App() {
  const [scheduleDataStr, setScheduleDataStr] = useState('');
  const [rosterDataStr, setRosterDataStr] = useState('');
  const [weekStartDate, setWeekStartDate] = useState('');
  const [campaignName, setCampaignName] = useState('Gen Mobile');
  const [defaultWorkType, setDefaultWorkType] = useState('WFH');
  const [defaultTeam, setDefaultTeam] = useState('GT-Gen Mobile -01');
  const [outputData, setOutputData] = useState<any[]>([]);

  const handleGenerate = () => {
    try {
      if (!scheduleDataStr || !rosterDataStr || !weekStartDate) {
        alert('Please provide Schedule Data, Roster Data, and a Week Start Date.');
        return;
      }

      const scheduleData = parseTSV(scheduleDataStr.trim());
      const rosterData = parseTSV(rosterDataStr.trim());

      if (scheduleData.length === 0 || rosterData.length === 0) {
        alert('Data is empty.');
        return;
      }

      // --- ROSTER PARSING ---
      let rHeaderIdx = rosterData.findIndex(row => row.some(cell => cell.toUpperCase().includes('EMP ID') || cell.toUpperCase().includes('FULL NAME')));
      if (rHeaderIdx === -1) rHeaderIdx = 0;
      const rHeaders = rosterData[rHeaderIdx] || [];
      
      let rEmpIdIdx = rHeaders.findIndex(h => h.toUpperCase().includes('EMP ID'));
      if (rEmpIdIdx === -1) rEmpIdIdx = 0;
      
      let rFullNameIdx = rHeaders.findIndex(h => h.toUpperCase().includes('FULL NAME'));
      if (rFullNameIdx === -1) rFullNameIdx = 1;
      
      let rShortNameIdx = rHeaders.findIndex(h => h.toUpperCase().includes('SHORT NAME'));
      if (rShortNameIdx === -1) rShortNameIdx = 2;
      
      let rSupervisorIdx = rHeaders.findIndex(h => h.toUpperCase().includes('SUPERVISOR'));
      if (rSupervisorIdx === -1) rSupervisorIdx = 8;

      const rosterList: any[] = [];
      for (let i = 0; i < rosterData.length; i++) {
        if (i === rHeaderIdx) continue;
        const row = rosterData[i];
        const fullName = row[rFullNameIdx]?.trim();
        if (fullName && fullName.toUpperCase() !== 'FULL NAME') {
          rosterList.push({
            originalName: fullName,
            normName: normalizeName(fullName),
            parts: normalizeName(fullName).split(' '),
            empId: row[rEmpIdIdx]?.trim() || '',
            shortName: row[rShortNameIdx]?.trim() || '',
            supervisor: row[rSupervisorIdx]?.trim() || ''
          });
        }
      }

      // --- SCHEDULE PARSING ---
      let sHeaderIdx = scheduleData.findIndex(row => row.some(cell => cell.toUpperCase().replace(/\s+/g, '').includes('ID-NAME') || cell.toUpperCase().replace(/\s+/g, '').includes('SHIFTSTART')));
      
      let idNameIdx = -1;
      let idIdx = -1;
      let shiftStartIndices: number[] = [];
      let shiftEndIndices: number[] = [];
      let b1StartIndices: number[] = [];
      let b1EndIndices: number[] = [];
      let lunchStartIndices: number[] = [];
      let lunchEndIndices: number[] = [];
      let b2StartIndices: number[] = [];
      let b2EndIndices: number[] = [];
      let breakLunchDisplayIndices: number[] = [];

      if (sHeaderIdx !== -1) {
        const sHeaders = scheduleData[sHeaderIdx] || [];
        idNameIdx = sHeaders.findIndex(h => h.toUpperCase().replace(/\s+/g, '').includes('ID-NAME') || h.toUpperCase().replace(/\s+/g, '').includes('NAME'));
        idIdx = sHeaders.findIndex(h => h.toUpperCase().replace(/\s+/g, '') === 'ID' || h.toUpperCase().replace(/\s+/g, '') === 'EMPID');

        sHeaders.forEach((h, i) => {
          const normH = h.toUpperCase().replace(/\s+/g, '');
          if (normH.includes('SHIFTSTART')) shiftStartIndices.push(i);
          if (normH.includes('SHIFTEND')) shiftEndIndices.push(i);
          if (normH.includes('BREAK1START')) b1StartIndices.push(i);
          if (normH.includes('BREAK1END')) b1EndIndices.push(i);
          if (normH.includes('LUNCHSTART')) lunchStartIndices.push(i);
          if (normH.includes('LUNCHEND')) lunchEndIndices.push(i);
          if (normH.includes('BREAK2START')) b2StartIndices.push(i);
          if (normH.includes('BREAK2END')) b2EndIndices.push(i);
          if (normH.includes('BREAKLUNCHSCHEDULEDDISPLAY')) breakLunchDisplayIndices.push(i);
        });
      }

      // Auto-detect columns if headers are missing or incomplete
      if (shiftStartIndices.length < 7 || shiftEndIndices.length < 7 || idNameIdx === -1) {
        // Find the first row that looks like it has time data
        const dataRow = scheduleData.find(row => row.some(cell => /^\d{1,2}:\d{2}/.test(cell.trim()) || cell.trim().toUpperCase() === 'OFF'));
        if (dataRow) {
          const firstTimeIdx = dataRow.findIndex(cell => /^\d{1,2}:\d{2}/.test(cell.trim()) || cell.trim().toUpperCase() === 'OFF');
          if (firstTimeIdx > 0) {
            if (idNameIdx === -1) idNameIdx = firstTimeIdx - 1; // Name is usually right before the first time
            
            shiftStartIndices = [];
            shiftEndIndices = [];
            for (let i = 0; i < 7; i++) {
              shiftStartIndices.push(firstTimeIdx + (i * 2));
              shiftEndIndices.push(firstTimeIdx + (i * 2) + 1);
            }
          }
        }
      }

      if (shiftStartIndices.length < 7 || shiftEndIndices.length < 7 || idNameIdx === -1) {
        alert('Could not detect the Schedule columns. Please ensure you copied the data correctly including the times.');
        return;
      }

      const [y, m, d] = weekStartDate.split('-');
      const startDate = new Date(parseInt(y), parseInt(m) - 1, parseInt(d));
      const newOutput = [];

      for (let i = 0; i < scheduleData.length; i++) {
        if (i === sHeaderIdx) continue;
        const row = scheduleData[i];
        if (!row || row.length < 5) continue;

        const fullName = row[idNameIdx]?.trim();
        // Skip if empty or looks like a sub-header (e.g., "Mon", "Tue")
        if (!fullName || fullName.toUpperCase().includes('ID-NAME') || fullName.toUpperCase() === 'MON' || fullName.toUpperCase() === 'TUE') continue;

        const normSch = normalizeName(fullName);
        const schParts = normSch.split(' ');

        let rosterInfo = rosterList.find(r => r.normName === normSch);

        // Fallback 1: Match by Short Name
        if (!rosterInfo) {
          rosterInfo = rosterList.find(r => normalizeName(r.shortName) === normSch);
        }

        // Fallback 2: All words in Roster's Short Name are in Schedule Name
        if (!rosterInfo) {
          rosterInfo = rosterList.find(r => {
            const rShortParts = normalizeName(r.shortName).split(' ');
            return rShortParts.length >= 2 && rShortParts.every((p: string) => schParts.includes(p));
          });
        }

        // Fallback 3: All words in Roster's Full Name are in Schedule Name
        if (!rosterInfo) {
          rosterInfo = rosterList.find(r => {
            return r.parts.length >= 2 && r.parts.every((p: string) => schParts.includes(p));
          });
        }

        // Fallback 4: First name and Third name
        if (!rosterInfo && schParts.length >= 3) {
          const target = `${schParts[0]} ${schParts[2]}`;
          rosterInfo = rosterList.find(r => r.parts.length >= 3 && `${r.parts[0]} ${r.parts[2]}` === target);
        }

        // Fallback 5: First name and Last name
        if (!rosterInfo && schParts.length >= 2) {
          const target = `${schParts[0]} ${schParts[schParts.length - 1]}`;
          rosterInfo = rosterList.find(r => r.parts.length >= 2 && `${r.parts[0]} ${r.parts[r.parts.length - 1]}` === target);
        }

        // Fallback 6: First name and Second name
        if (!rosterInfo && schParts.length >= 2) {
          const target = `${schParts[0]} ${schParts[1]}`;
          rosterInfo = rosterList.find(r => r.parts.length >= 2 && `${r.parts[0]} ${r.parts[1]}` === target);
        }

        // Fallback 7: Fuzzy match for typos (e.g. Gladis vs Gladys, Moterroso vs Monterroso)
        if (!rosterInfo && schParts.length >= 2) {
          rosterInfo = rosterList.find(r => {
            if (r.parts.length >= 2 && r.parts[0].length >= 4 && schParts[0].length >= 4) {
              const firstNameMatch = r.parts[0].substring(0, 4) === schParts[0].substring(0, 4);
              const surnameMatch = r.parts.some((rp: string, i: number) => {
                if (i === 0) return false;
                return schParts.some(sp => isSimilar(rp, sp));
              });
              return firstNameMatch && surnameMatch;
            }
            return false;
          });
        }

        rosterInfo = rosterInfo || {};
        const visualCode = rosterInfo.empId || row[idIdx]?.trim() || '';
        
        let employeeName = '';
        if (rosterInfo.shortName) {
           employeeName = normalizeName(rosterInfo.shortName).replace(/\s+/g, '.');
        } else {
           // LatAm naming convention fallback: Primer Nombre + Primer Apellido
           if (schParts.length >= 4) {
             employeeName = `${schParts[0]}.${schParts[2]}`;
           } else if (schParts.length >= 2) {
             employeeName = `${schParts[0]}.${schParts[1]}`;
           } else {
             employeeName = schParts[0] || '';
           }
        }
        
        // Always use the default team as requested
        const team = defaultTeam;

        let rawStarts = [];
        let rawEnds = [];
        let rawB1Starts = [];
        let rawB1Ends = [];
        let rawB2Starts = [];
        let rawB2Ends = [];
        let rawLStarts = [];
        let rawLEnds = [];

        for (let day = 0; day < 7; day++) {
          rawStarts.push(row[shiftStartIndices[day]]?.trim() || '');
          rawEnds.push(row[shiftEndIndices[day]]?.trim() || '');
          rawB1Starts.push(b1StartIndices[day] !== undefined ? row[b1StartIndices[day]]?.trim() || '' : '');
          rawB1Ends.push(b1EndIndices[day] !== undefined ? row[b1EndIndices[day]]?.trim() || '' : '');
          rawB2Starts.push(b2StartIndices[day] !== undefined ? row[b2StartIndices[day]]?.trim() || '' : '');
          rawB2Ends.push(b2EndIndices[day] !== undefined ? row[b2EndIndices[day]]?.trim() || '' : '');
          rawLStarts.push(lunchStartIndices[day] !== undefined ? row[lunchStartIndices[day]]?.trim() || '' : '');
          rawLEnds.push(lunchEndIndices[day] !== undefined ? row[lunchEndIndices[day]]?.trim() || '' : '');
        }

        // Check for "Starts then Ends" format (pasted incorrectly into interleaved columns)
        let invalidShifts = 0;
        for (let day = 0; day < 7; day++) {
          const s = rawStarts[day].toUpperCase();
          const e = rawEnds[day].toUpperCase();
          const isOffS = s === 'OFF' || s === '';
          const isOffE = e === 'OFF' || e === '';
          
          if (isOffS !== isOffE) {
            invalidShifts++;
          } else if (!isOffS && !isOffE && s === e) {
            invalidShifts++;
          }
        }

        if (invalidShifts >= 4) {
          // Reconstruct the original 14 values in order for each category
          const rearrange = (starts: string[], ends: string[]) => {
            const v = [];
            for (let day = 0; day < 7; day++) {
              v.push(starts[day]);
              v.push(ends[day]);
            }
            return [v.slice(0, 7), v.slice(7, 14)];
          };

          [rawStarts, rawEnds] = rearrange(rawStarts, rawEnds);
          [rawB1Starts, rawB1Ends] = rearrange(rawB1Starts, rawB1Ends);
          [rawB2Starts, rawB2Ends] = rearrange(rawB2Starts, rawB2Ends);
          [rawLStarts, rawLEnds] = rearrange(rawLStarts, rawLEnds);
        }

        for (let day = 0; day < 7; day++) {
          const sStart = rawStarts[day];
          const sEnd = rawEnds[day];

          if (!sStart && !sEnd) continue;
          if (sStart.toUpperCase() === 'MON' || sStart.toUpperCase() === 'TUE') continue;

          const isOff = sStart.toUpperCase() === 'OFF' || sStart === '';
          const schIn = convertTime(sStart);
          const schOut = convertTime(sEnd);
          const minutes = calculateMinutes(schIn, schOut);

          const b1In = convertTime(rawB1Starts[day]);
          const b1Out = convertTime(rawB1Ends[day]);
          const b2In = convertTime(rawB2Starts[day]);
          const b2Out = convertTime(rawB2Ends[day]);
          const lIn = convertTime(rawLStarts[day]);
          const lOut = convertTime(rawLEnds[day]);

          const breakMins = calculateMinutes(b1In, b1Out) + calculateMinutes(b2In, b2Out);
          const lunchMins = calculateMinutes(lIn, lOut);

          let breakLunchStr = '';
          if (breakLunchDisplayIndices[day] !== undefined && row[breakLunchDisplayIndices[day]]) {
            breakLunchStr = row[breakLunchDisplayIndices[day]].trim();
          } else if (!isOff && (breakMins > 0 || lunchMins > 0)) {
            breakLunchStr = `Lunch: ${formatDuration(lunchMins)}\nBreak: ${formatDuration(breakMins)}`;
          }

          const currentDate = new Date(startDate);
          currentDate.setDate(currentDate.getDate() + day);

          newOutput.push({
            VisualCode: visualCode,
            Campaign: campaignName,
            Team: team,
            EmployeeName: employeeName,
            Date: formatDate(currentDate),
            SchIn: schIn,
            SchOut: schOut,
            PnchIn: '',
            PnchOut: '',
            Staffed: isOff ? 0 : minutes,
            Scheduled: isOff ? 0 : minutes,
            Paused: 0,
            Remark: isOff ? 'Rest Day' : 'Present',
            ScheduleType: isOff ? 'Rest Day' : ' Regular',
            WorkType: isOff ? '' : defaultWorkType,
            BreakLunchScheduledDisplay: breakLunchStr,
            BreakLunchStaffedDisplay: '',
            BreakLunchRemark: ''
          });
        }
      }

      if (newOutput.length === 0) {
        alert('No data could be generated. Please check if the columns match the expected format.');
      } else {
        setOutputData(newOutput);
        alert(`Successfully generated ${newOutput.length} rows!`);
      }
    } catch (error: any) {
      console.error(error);
      alert(`An error occurred: ${error.message}`);
    }
  };

  const handleCopy = () => {
    if (outputData.length === 0) return;
    const headers = Object.keys(outputData[0]);
    const tsv = [
      headers.join('\t'),
      ...outputData.map(row => headers.map(h => {
        let val = String(row[h] || '');
        if (val.includes('\n') || val.includes('\t') || val.includes('"')) {
          val = `"${val.replace(/"/g, '""')}"`;
        }
        return val;
      }).join('\t'))
    ].join('\n');
    navigator.clipboard.writeText(tsv);
    alert('Copied to clipboard!');
  };

  const handleExportCSV = () => {
    if (outputData.length === 0) return;
    const headers = Object.keys(outputData[0]);
    const csv = [
      headers.join(','),
      ...outputData.map(row => headers.map(h => {
        let val = String(row[h] || '');
        if (val.includes(',') || val.includes('\n') || val.includes('"')) {
          val = `"${val.replace(/"/g, '""')}"`;
        }
        return val;
      }).join(','))
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

        <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
          <div className="bg-white p-4 rounded-xl shadow-sm border border-slate-200 flex flex-col">
            <label className="flex items-center space-x-2 text-sm font-medium text-slate-700 mb-2">
              <FileSpreadsheet size={18} className="text-indigo-500" />
              <span>Paste Schedule Data (Image 1)</span>
            </label>
            <textarea 
              className="flex-1 min-h-[200px] p-3 text-sm border border-slate-300 rounded-lg focus:ring-2 focus:ring-indigo-500 focus:border-indigo-500 outline-none resize-y font-mono whitespace-pre"
              placeholder="Paste TSV data here (copy directly from Excel)..."
              value={scheduleDataStr}
              onChange={e => setScheduleDataStr(e.target.value)}
            />
          </div>

          <div className="bg-white p-4 rounded-xl shadow-sm border border-slate-200 flex flex-col">
            <label className="flex items-center space-x-2 text-sm font-medium text-slate-700 mb-2">
              <Users size={18} className="text-emerald-500" />
              <span>Paste Timeline/Roster Data (Image 3)</span>
            </label>
            <textarea 
              className="flex-1 min-h-[200px] p-3 text-sm border border-slate-300 rounded-lg focus:ring-2 focus:ring-emerald-500 focus:border-emerald-500 outline-none resize-y font-mono whitespace-pre"
              placeholder="Paste TSV data here (copy directly from Excel)..."
              value={rosterDataStr}
              onChange={e => setRosterDataStr(e.target.value)}
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
              onChange={e => setWeekStartDate(e.target.value)}
            />
          </div>
          <div className="flex flex-col space-y-1">
            <label className="text-xs font-medium text-slate-500 uppercase tracking-wider">Campaign Name</label>
            <input 
              type="text" 
              className="px-3 py-2 border border-slate-300 rounded-lg focus:ring-2 focus:ring-indigo-500 outline-none"
              value={campaignName}
              onChange={e => setCampaignName(e.target.value)}
            />
          </div>
          <div className="flex flex-col space-y-1">
            <label className="text-xs font-medium text-slate-500 uppercase tracking-wider">Default Team</label>
            <input 
              type="text" 
              className="px-3 py-2 border border-slate-300 rounded-lg focus:ring-2 focus:ring-indigo-500 outline-none"
              value={defaultTeam}
              onChange={e => setDefaultTeam(e.target.value)}
            />
          </div>
          <div className="flex flex-col space-y-1">
            <label className="text-xs font-medium text-slate-500 uppercase tracking-wider">Default Work Type</label>
            <input 
              type="text" 
              className="px-3 py-2 border border-slate-300 rounded-lg focus:ring-2 focus:ring-indigo-500 outline-none w-32"
              value={defaultWorkType}
              onChange={e => setDefaultWorkType(e.target.value)}
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
                    {Object.keys(outputData[0]).map(key => (
                      <th key={key} className="px-4 py-3 font-medium tracking-wider">{key}</th>
                    ))}
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-100">
                  {outputData.slice(0, 100).map((row, i) => (
                    <tr key={i} className="hover:bg-slate-50/50">
                      {Object.values(row).map((val: any, j) => (
                        <td key={j} className="px-4 py-2 text-slate-600">{val}</td>
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
