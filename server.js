/**
 * Attendance Processor — rewritten with all QA fixes applied.
 *
 * Fixes vs. original:
 *  1. Auto header-row detection  (no more hardcoded skipRows)
 *  2. TSV single-column fallback  (punch file was TSV with .xls ext)
 *  3. Deterministic regularisation dedup  (latest Approval Time wins)
 *  4. UTC-only date arithmetic  (no DST one-day drift)
 *  5. Explicit Date-object guard in all date parsers
 *  6. Multer file-size limits  (10 MB per file)
 *  7. Centralised JSON error middleware  (no more HTML error pages)
 *  8. Single-punch business rule documented & preserved (pre/post noon)
 */

const express  = require('express');
const multer   = require('multer');
const XLSX     = require('xlsx');
const ExcelJS  = require('exceljs');
const path     = require('path');
const fs       = require('fs');
const os       = require('os');

const app  = express();
const PORT = 3000;

app.use(express.static(path.join(__dirname, 'public')));
app.use(express.json());

// ─── Upload middleware ────────────────────────────────────────────────────────

const upload = multer({
  dest: os.tmpdir(),
  limits: {
    fileSize:  10 * 1024 * 1024, // 10 MB per file
    files:     3,
    fields:    0,
  },
  fileFilter: (_req, file, cb) => {
    const allowed = ['.xlsx', '.xls', '.csv'];
    const ext = path.extname(file.originalname).toLowerCase();
    if (allowed.includes(ext)) cb(null, true);
    else cb(Object.assign(new Error(`Unsupported file type: ${ext}`), { code: 'UNSUPPORTED_TYPE' }));
  },
});

// ─── Month map ───────────────────────────────────────────────────────────────

const MONTH_MAP = {
  jan:1, feb:2, mar:3, apr:4, may:5, jun:6,
  jul:7, aug:8, sep:9, oct:10, nov:11, dec:12,
};

// ─── Date / time helpers ─────────────────────────────────────────────────────

/**
 * Guard: if val is already a JS Date, convert it to ISO string before regex work.
 * This protects against future cellDates:true changes in SheetJS options.
 */
function normaliseDateInput(val) {
  if (val instanceof Date) {
    // Always serialise as UTC date-only to avoid local-timezone shift.
    const y  = val.getUTCFullYear();
    const mo = String(val.getUTCMonth() + 1).padStart(2, '0');
    const d  = String(val.getUTCDate()).padStart(2, '0');
    return `${y}-${mo}-${d}`;
  }
  return val;
}

/**
 * Parse Punch / Regularisation date to YYYY-MM-DD.
 * Handles: "31 Jan 2026", "31-Jan-26", "31/01/2026", "2026-01-31", Excel serial.
 */
function parsePunchDate(val) {
  val = normaliseDateInput(val);
  if (!val) return null;
  const s = String(val).trim();
  if (!s) return null;

  // "31 Jan 2026"  or  "31-Jan-26"
  let m = s.match(/^(\d{1,2})[\s\-\/]([A-Za-z]{3})[\s\-\/](\d{2,4})$/);
  if (m) {
    const day = m[1].padStart(2, '0');
    const mon = MONTH_MAP[m[2].toLowerCase()];
    if (!mon) return null;
    let year = parseInt(m[3]);
    if (year < 100) year += 2000;
    return `${year}-${String(mon).padStart(2, '0')}-${day}`;
  }

  // YYYY-MM-DD or YYYY/MM/DD
  m = s.match(/^(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})$/);
  if (m) return `${m[1]}-${m[2].padStart(2, '0')}-${m[3].padStart(2, '0')}`;

  // DD/MM/YYYY or DD-MM-YYYY
  m = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})$/);
  if (m) return `${m[3]}-${m[2].padStart(2, '0')}-${m[1].padStart(2, '0')}`;

  // Excel serial number
  if (/^\d+(\.\d+)?$/.test(s)) {
    const d = XLSX.SSF.parse_date_code(parseFloat(s));
    if (d) return `${d.y}-${String(d.m).padStart(2, '0')}-${String(d.d).padStart(2, '0')}`;
  }

  return null;
}

const parseRegDate = parsePunchDate; // same formats

/**
 * Parse Leave date: "01/14/2026" → "2026-01-14"  (MM/DD/YYYY)
 */
function parseLeaveDate(val) {
  val = normaliseDateInput(val);
  if (!val) return null;
  const s = String(val).trim();

  // MM/DD/YYYY
  const m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (m) return `${m[3]}-${m[1].padStart(2, '0')}-${m[2].padStart(2, '0')}`;

  return parsePunchDate(s);
}

/**
 * Add `days` to a YYYY-MM-DD string using UTC-only arithmetic.
 * Avoids the DST one-day drift that JS local setDate() can cause.
 */
function addDateDaysUTC(dateStr, days) {
  const [y, mo, d] = dateStr.split('-').map(Number);
  const dt = new Date(Date.UTC(y, mo - 1, d + days));
  const yr = dt.getUTCFullYear();
  const mn = String(dt.getUTCMonth() + 1).padStart(2, '0');
  const dy = String(dt.getUTCDate()).padStart(2, '0');
  return `${yr}-${mn}-${dy}`;
}

/**
 * Parse time string → "HH:MM:SS" (24 h), or null.
 * Handles "17:57:00", "09:01 AM", "06:31 PM", Excel fraction, "-", blank.
 */
function parseTime(val) {
  if (!val) return null;
  const s = String(val).trim();
  if (!s || s === '-' || s === '0') return null;

  const m = s.match(/^(\d{1,2}):(\d{2})(?::(\d{2}))?\s*(AM|PM)?$/i);
  if (m) {
    let h     = parseInt(m[1]);
    const min = parseInt(m[2]);
    const sec = m[3] ? parseInt(m[3]) : 0;
    const ap  = m[4] ? m[4].toUpperCase() : null;
    if (ap === 'PM' && h < 12) h += 12;
    if (ap === 'AM' && h === 12) h = 0;
    return `${String(h).padStart(2, '0')}:${String(min).padStart(2, '0')}:${String(sec).padStart(2, '0')}`;
  }

  if (/^\d+\.\d+$/.test(s)) {
    const totalSec = Math.round(parseFloat(s) * 86400);
    const h  = Math.floor(totalSec / 3600);
    const mn = Math.floor((totalSec % 3600) / 60);
    const sc = totalSec % 60;
    return `${String(h).padStart(2, '0')}:${String(mn).padStart(2, '0')}:${String(sc).padStart(2, '0')}`;
  }

  return null;
}

function timeToSeconds(t) {
  if (!t) return null;
  const parts = t.split(':').map(Number);
  return parts[0] * 3600 + parts[1] * 60 + (parts[2] || 0);
}

function isMorning(timeStr) {
  const s = timeToSeconds(timeStr);
  return s !== null && s < 12 * 3600;
}

/**
 * Calculate working hours between check-in and check-out.
 * Returns a human-readable string like "8h 30m", or null if not calculable.
 */
function calcWorkingHours(checkIn, checkOut) {
  if (!checkIn || !checkOut || checkIn === '0' || checkOut === '0') return null;
  const inSec  = timeToSeconds(checkIn);
  const outSec = timeToSeconds(checkOut);
  if (inSec === null || outSec === null || outSec <= inSec) return null;
  const diffSec = outSec - inSec;
  const h = Math.floor(diffSec / 3600);
  const m = Math.floor((diffSec % 3600) / 60);
  return `${h}h ${String(m).padStart(2, '0')}m`;
}

/**
 * Strip leading apostrophe (Excel text prefix) and whitespace.
 */
function cleanCardNo(val) {
  if (val === null || val === undefined) return '';
  return String(val).trim().replace(/^'/, '');
}

/**
 * Extract trailing numeric token: "Riya Kohli 1774" → "1774"
 */
function extractEmpId(val) {
  const s = String(val).trim();
  if (/^\d+$/.test(s)) return s;
  const parts = s.split(/\s+/);
  const last  = parts[parts.length - 1];
  if (/^\d+$/.test(last)) return last;
  return s;
}

// ─── File parsing ─────────────────────────────────────────────────────────────

/**
 * SheetJS → array-of-arrays, then optionally split tab-delimited single-column rows.
 * This handles TSV content stored with an .xls extension (real punch file issue).
 */
function sheetToAoA(filePath) {
  const wb = XLSX.readFile(filePath, { raw: true, cellText: true, cellDates: false });
  const ws = wb.Sheets[wb.SheetNames[0]];
  let aoa  = XLSX.utils.sheet_to_json(ws, { defval: '', header: 1 });

  // TSV fallback: if every non-empty row has exactly 1 cell containing a tab,
  // the heuristic delimiter detection failed — split manually.
  const nonEmpty = aoa.filter(r => r.some(c => String(c).trim() !== ''));
  const allSingleTab = nonEmpty.length > 0 && nonEmpty.every(r =>
    r.length === 1 && String(r[0]).includes('\t')
  );

  if (allSingleTab) {
    aoa = aoa.map(r => String(r[0] || '').split('\t'));
  }

  return aoa;
}

/**
 * Find the header row by scanning the first `scanLimit` rows for a row that
 * contains ALL of the required column names (case-insensitive, trimmed).
 *
 * Returns { headers: string[], dataRows: any[][] } or throws a descriptive error.
 *
 * This replaces the brittle hardcoded skipRows parameter.
 */
function findHeaderAndData(aoa, requiredCols, label, scanLimit = 50) {
  for (let i = 0; i < Math.min(aoa.length, scanLimit); i++) {
    const row     = aoa[i];
    const trimmed = row.map(c => String(c).trim());
    // Check if every required column appears somewhere in this row
    const match = requiredCols.every(req =>
      trimmed.some(cell => cell.toLowerCase() === req.toLowerCase())
    );
    if (match) {
      const headers  = trimmed;
      const dataRows = aoa.slice(i + 1).filter(r =>
        r.some(c => String(c).trim() !== '')
      );
      return { headers, dataRows };
    }
  }

  // Not found — provide a helpful diagnostic
  const firstFew = aoa.slice(0, 5).map(r => r.map(c => String(c).trim()).join(' | ')).join('\n  ');
  throw new Error(
    `Could not find header row in ${label}.\n` +
    `Expected columns: ${requiredCols.join(', ')}\n` +
    `First 5 rows of file:\n  ${firstFew}`
  );
}

/**
 * Parse a file into array of row-objects, auto-detecting the header row.
 */
function parseFile(filePath, requiredCols, label) {
  const aoa              = sheetToAoA(filePath);
  const { headers, dataRows } = findHeaderAndData(aoa, requiredCols, label);

  return dataRows.map(row => {
    const obj = {};
    headers.forEach((h, i) => { obj[h] = row[i] !== undefined ? row[i] : ''; });
    return obj;
  });
}

// ─── Core processing ──────────────────────────────────────────────────────────

function processPunchData(punchRows, regularRows, leaveRows) {

  // ── Step 1: Punch map ──────────────────────────────────────────────────────
  if (!punchRows.length) throw new Error('Punch report is empty or could not be parsed.');

  const punchMap = {}; // "cardNo|YYYY-MM-DD" → { cardNo, date, times[] }

  for (const row of punchRows) {
    const cardNo = cleanCardNo(row['Card No']);
    if (!cardNo || cardNo === 'Card No') continue;

    const date = parsePunchDate(row['Punch Date']);
    const time = parseTime(row['Punch Time']);
    if (!date || !time) continue;

    const key = `${cardNo}|${date}`;
    if (!punchMap[key]) punchMap[key] = { cardNo, date, times: [] };
    punchMap[key].times.push(time);
  }

  if (Object.keys(punchMap).length === 0) {
    throw new Error(
      'No valid punch records found. ' +
      'Detected columns: ' + Object.keys(punchRows[0] || {}).join(', ') + '. ' +
      'Expected: "Card No", "Punch Date", "Punch Time".'
    );
  }

  // ── Step 2: Regularisation map ─────────────────────────────────────────────
  // FIX: deterministic dedup — latest Approval Time wins (tie-break: last row).
  const regMap  = {}; // "empId|YYYY-MM-DD" → { checkIn, checkOut, approvalTime }

  for (const row of regularRows) {
    const status = String(row['Approval Status'] || '').trim().toLowerCase();
    if (status !== 'approved') continue;

    const empId = cleanCardNo(row['Employee Id']);
    if (!empId) continue;
    const date = parseRegDate(row['Attendance Day']);
    if (!date) continue;

    const approvalTime = parseTime(row['Approval Time']) || parseTime(row['Created Time']) || '00:00:00';
    const key = `${empId}|${date}`;

    const existing = regMap[key];
    if (existing && timeToSeconds(existing.approvalTime) > timeToSeconds(approvalTime)) {
      continue; // existing record is newer — keep it
    }

    regMap[key] = {
      checkIn:      parseTime(row['New Check-In'])  || '0',
      checkOut:     parseTime(row['New Check-Out']) || '0',
      approvalTime,
    };
  }

  // ── Step 3: Leave map ──────────────────────────────────────────────────────
  // Stores: { fullDay: bool, session: string }
  // "Days/Hours Taken" = 1   → full day leave  (wipe punch times)
  // "Days/Hours Taken" = 0.5 → half day leave  (keep punch times, note session)
  const leaveMap = {}; // "empId|YYYY-MM-DD" → { fullDay, session }

  for (const row of leaveRows) {
    const approvalStatus = String(row['Approval Status'] || '').trim().toLowerCase();
    if (approvalStatus !== 'approved') continue;

    const empId    = extractEmpId(row['Employee ID']);
    if (!empId) continue;
    const fromDate = parseLeaveDate(row['From']);
    const toDate   = parseLeaveDate(row['To']) || fromDate;
    if (!fromDate) continue;

    // Parse duration — default to full day if column absent/unreadable
    const daysRaw = row['Days/Hours Taken'];
    const days    = (daysRaw !== undefined && daysRaw !== '') ? parseFloat(daysRaw) : 1;
    const fullDay = isNaN(days) || days >= 1;

    // Session only matters for half-day leaves
    const session = String(row['Session'] || '').trim(); // "First Half" | "Second Half" | ""

    let cur = fromDate;
    while (cur <= toDate) {
      leaveMap[`${empId}|${cur}`] = { fullDay, session };
      cur = addDateDaysUTC(cur, 1);
    }
  }

  // ── Step 4: Merge ──────────────────────────────────────────────────────────
  // Only include days that have an actual punch. Regularization and leave
  // are overrides only — they never create new rows on their own.
  const allKeys = new Set(Object.keys(punchMap));
  const outputRows = [];

  for (const key of allKeys) {
    let cardNo, date;
    if (punchMap[key]) {
      ({ cardNo, date } = punchMap[key]);
    } else {
      [cardNo, date] = key.split('|');
    }

    let checkIn  = '0';
    let checkOut = '0';
    let status   = 'Present';

    // Punch times
    if (punchMap[key]) {
      const times = [...punchMap[key].times].sort((a, b) => timeToSeconds(a) - timeToSeconds(b));
      if (times.length === 1) {
        // Business rule: single punch before noon → check-in; after noon → check-out.
        // NOTE: This will misclassify night-shift and half-day employees.
        //       Adjust this rule to match your HR policy if needed.
        if (isMorning(times[0])) { checkIn = times[0]; }
        else { checkOut = times[0]; }
      } else {
        checkIn  = times[0];
        checkOut = times[times.length - 1];
      }
    }

    // Regularisation override
    if (regMap[key]) {
      checkIn  = regMap[key].checkIn;
      checkOut = regMap[key].checkOut;
      status   = 'Regularized';
    }

    // Leave override (highest priority)
    if (leaveMap[key]) {
      const leave = leaveMap[key];
      if (leave.fullDay) {
        // Full day leave — wipe punch times entirely
        checkIn  = '0';
        checkOut = '0';
        status   = 'On Leave';
      } else {
        // Half day leave — keep whatever punch times exist, just update status.
        // The employee worked part of the day so check-in/out remain valid.
        const sessionLabel = leave.session ? ` (${leave.session})` : '';
        status = `Half Day Leave${sessionLabel}`;
      }
    }

    const workingHours = calcWorkingHours(checkIn, checkOut);
    outputRows.push({ cardNo, date, checkIn, checkOut, workingHours, status });
  }

  // ── Step 5: Sort — numeric Card No, then date ──────────────────────────────
  outputRows.sort((a, b) => {
    const na = parseInt(a.cardNo) || 0;
    const nb = parseInt(b.cardNo) || 0;
    if (na !== nb) return na - nb;
    return a.date < b.date ? -1 : a.date > b.date ? 1 : 0;
  });

  return {
    rows: outputRows,
    stats: {
      totalRecords:    outputRows.length,
      present:         outputRows.filter(r => r.status === 'Present').length,
      regularized:     outputRows.filter(r => r.status === 'Regularized').length,
      onLeave:         outputRows.filter(r => r.status === 'On Leave').length,
      halfDayLeave:    outputRows.filter(r => r.status.startsWith('Half Day Leave')).length,
      uniqueEmployees: [...new Set(outputRows.map(r => r.cardNo))].length,
    },
  };
}

// ─── Excel builder ────────────────────────────────────────────────────────────

/**
 * Collapse detail rows into one record per Card No with working / leave counts.
 */
function buildSummary(rows) {
  const map = {};
  for (const r of rows) {
    if (!map[r.cardNo]) map[r.cardNo] = { cardNo: r.cardNo, working: 0, onLeave: 0, halfDay: 0 };
    if (r.status === 'On Leave') {
      map[r.cardNo].onLeave++;
    } else if (r.status.startsWith('Half Day')) {
      // Employee worked half the day — counts as 0.5 in Present+Reg and 0.5 in Half Day Leave
      map[r.cardNo].halfDay   += 0.5;
      map[r.cardNo].working   += 0.5;
    } else {
      map[r.cardNo].working++;
    }
  }
  return Object.values(map).sort((a, b) => (parseInt(a.cardNo) || 0) - (parseInt(b.cardNo) || 0));
}

async function buildWorkbook(rows, stats) {
  const workbook = new ExcelJS.Workbook();
  workbook.creator = 'Attendance Processor';
  workbook.created = new Date();

  // ── Main sheet ─────────────────────────────────────────────────────────────
  const sheet = workbook.addWorksheet('Attendance Report', {
    views: [{ state: 'frozen', ySplit: 1 }],
  });

  sheet.columns = [
    { header: 'Card No',       key: 'cardNo',       width: 16 },
    { header: 'Punch Date',    key: 'date',         width: 14 },
    { header: 'Check In',      key: 'checkIn',      width: 12 },
    { header: 'Check Out',     key: 'checkOut',     width: 12 },
    { header: 'Working Hours', key: 'workingHours', width: 16 },
    { header: 'Status',        key: 'status',       width: 14 },
  ];

  const headerRow = sheet.getRow(1);
  headerRow.eachCell(cell => {
    cell.fill      = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1A1A2E' } };
    cell.font      = { bold: true, color: { argb: 'FFE94560' }, size: 11, name: 'Calibri' };
    cell.alignment = { vertical: 'middle', horizontal: 'center' };
    cell.border    = { bottom: { style: 'medium', color: { argb: 'FFE94560' } } };
  });
  headerRow.height = 22;

  const statusColors = {
    'Present':     { bg: 'FFFFFFFF', fg: 'FF1A1A2E' },
    'Regularized': { bg: 'FFFFF3CD', fg: 'FF856404' },
    'On Leave':    { bg: 'FFFFE0E0', fg: 'FF721C24' },
    'Half Day':    { bg: 'FFECE8FE', fg: 'FF5B2DA8' }, // purple tint for half-day
  };

  for (let i = 0; i < rows.length; i++) {
    const r       = rows[i];
    const excelRow = sheet.addRow({ cardNo: r.cardNo, date: r.date, checkIn: r.checkIn, checkOut: r.checkOut, workingHours: r.workingHours || '—', status: r.status });
    const colorKey = r.status.startsWith('Half Day') ? 'Half Day' : r.status;
    const colors   = statusColors[colorKey] || statusColors['Present'];
    const bgColor  = r.status !== 'Present' ? colors.bg : (i % 2 === 0 ? 'FFF8F9FA' : 'FFFFFFFF');

    excelRow.eachCell(cell => {
      cell.fill      = { type: 'pattern', pattern: 'solid', fgColor: { argb: bgColor } };
      cell.font      = { color: { argb: colors.fg }, size: 10, name: 'Calibri' };
      cell.alignment = { vertical: 'middle', horizontal: 'center' };
      cell.border    = { bottom: { style: 'thin', color: { argb: 'FFE0E0E0' } } };
    });
    excelRow.height = 18;
  }

  sheet.autoFilter = { from: 'A1', to: 'F1' };

  // ── Summary sheet ──────────────────────────────────────────────────────────
  const statsSheet = workbook.addWorksheet('Summary');
  statsSheet.columns = [
    { header: 'Metric', key: 'metric', width: 24 },
    { header: 'Value',  key: 'value',  width: 14 },
  ];

  const sh = statsSheet.getRow(1);
  sh.eachCell(cell => {
    cell.fill      = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1A1A2E' } };
    cell.font      = { bold: true, color: { argb: 'FFE94560' }, size: 11 };
    cell.alignment = { vertical: 'middle', horizontal: 'center' };
  });
  sh.height = 22;

  for (const [metric, value] of [
    ['Total Records',    stats.totalRecords],
    ['Unique Employees', stats.uniqueEmployees],
    ['Present',          stats.present],
    ['Regularized',      stats.regularized],
    ['Full Day Leave',   stats.onLeave],
    ['Half Day Leave',   stats.halfDayLeave],
  ]) {
    const sr = statsSheet.addRow({ metric, value });
    sr.eachCell(cell => {
      cell.alignment = { vertical: 'middle', horizontal: 'center' };
      cell.font      = { size: 10, name: 'Calibri' };
      cell.border    = { bottom: { style: 'thin', color: { argb: 'FFE0E0E0' } } };
    });
    sr.height = 18;
  }

  // ── Employee Summary sheet ─────────────────────────────────────────────────
  const empSheet = workbook.addWorksheet('Employee Summary', {
    views: [{ state: 'frozen', ySplit: 1 }],
  });

  empSheet.columns = [
    { header: 'Card No',                  key: 'cardNo',   width: 16 },
    { header: 'Present + Regularized',    key: 'working',  width: 24 },
    { header: 'Half Day Leave',           key: 'halfDay',  width: 18 },
    { header: 'Full Day Leave',           key: 'onLeave',  width: 18 },
  ];

  const empHeader = empSheet.getRow(1);
  empHeader.eachCell(cell => {
    cell.fill      = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1A1A2E' } };
    cell.font      = { bold: true, color: { argb: 'FFE94560' }, size: 11, name: 'Calibri' };
    cell.alignment = { vertical: 'middle', horizontal: 'center' };
    cell.border    = { bottom: { style: 'medium', color: { argb: 'FFE94560' } } };
  });
  empHeader.height = 22;

  const summary = buildSummary(rows);
  for (let i = 0; i < summary.length; i++) {
    const s  = summary[i];
    const er = empSheet.addRow({ cardNo: s.cardNo, working: s.working, halfDay: s.halfDay || 0, onLeave: s.onLeave });
    const bg = i % 2 === 0 ? 'FFF8F9FA' : 'FFFFFFFF';
    er.eachCell(cell => {
      cell.fill      = { type: 'pattern', pattern: 'solid', fgColor: { argb: bg } };
      cell.font      = { size: 10, name: 'Calibri' };
      cell.alignment = { vertical: 'middle', horizontal: 'center' };
      cell.border    = { bottom: { style: 'thin', color: { argb: 'FFE0E0E0' } } };
    });
    // Colour the numbers
    er.getCell('working').font = { size: 10, name: 'Calibri', color: { argb: 'FF1A7A4E' }, bold: true };
    if ((s.halfDay || 0) > 0)
      er.getCell('halfDay').font = { size: 10, name: 'Calibri', color: { argb: 'FF5B2DA8' }, bold: true };
    if (s.onLeave > 0)
      er.getCell('onLeave').font = { size: 10, name: 'Calibri', color: { argb: 'FFC0392B' }, bold: true };
    er.height = 18;
  }

  empSheet.autoFilter = { from: 'A1', to: 'D1' };

  return workbook;
}

// ─── Shared file parsing logic ────────────────────────────────────────────────

function parseUploads(files) {
  const punchFile = files['punch'][0];
  const punchRows = parseFile(
    punchFile.path,
    ['Card No', 'Punch Date', 'Punch Time'],
    'Punch Report'
  );

  let regularRows = [];
  if (files['regularization']?.[0]) {
    regularRows = parseFile(
      files['regularization'][0].path,
      ['Employee Id', 'Attendance Day', 'New Check-In', 'New Check-Out', 'Approval Status'],
      'Regularization Report'
    );
  }

  let leaveRows = [];
  if (files['leave']?.[0]) {
    leaveRows = parseFile(
      files['leave'][0].path,
      ['Employee ID', 'From', 'To', 'Approval Status'],
      'Leave Report'
    );
  }

  return processPunchData(punchRows, regularRows, leaveRows);
}

// ─── Routes ───────────────────────────────────────────────────────────────────

const uploadFields = upload.fields([
  { name: 'punch',          maxCount: 1 },
  { name: 'regularization', maxCount: 1 },
  { name: 'leave',          maxCount: 1 },
]);

// POST /api/process  → returns styled XLSX download
app.post('/api/process', uploadFields, async (req, res) => {
  const tmpFiles = Object.values(req.files || {}).flat().map(f => f.path);
  try {
    if (!req.files?.['punch']) return res.status(400).json({ error: 'Punch report file is required.' });

    const { rows, stats } = parseUploads(req.files);
    const workbook        = await buildWorkbook(rows, stats);

    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', `attachment; filename="attendance_report_${Date.now()}.xlsx"`);
    res.setHeader('X-Stats', JSON.stringify(stats));

    await workbook.xlsx.write(res);
    res.end();
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: err.message });
  } finally {
    tmpFiles.forEach(f => { try { fs.unlinkSync(f); } catch (_) {} });
  }
});

// POST /api/preview  → returns JSON (first 100 rows + stats)
app.post('/api/preview', uploadFields, async (req, res) => {
  const tmpFiles = Object.values(req.files || {}).flat().map(f => f.path);
  try {
    if (!req.files?.['punch']) return res.status(400).json({ error: 'Punch report file is required.' });

    const { rows, stats } = parseUploads(req.files);

    // Send ALL rows to client so search works across the full dataset.
    // buildSummary has correct half-day handling (shared with Excel builder).
    const summary = buildSummary(rows);

    res.json({ stats, rows, total: rows.length, summary });
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: err.message });
  } finally {
    tmpFiles.forEach(f => { try { fs.unlinkSync(f); } catch (_) {} });
  }
});

// ─── Centralised JSON error middleware ────────────────────────────────────────
// FIX: Multer errors (wrong type, file too large) are caught here and returned
//      as JSON instead of Express' default HTML error page.
// eslint-disable-next-line no-unused-vars
app.use((err, _req, res, _next) => {
  if (err.code === 'LIMIT_FILE_SIZE') {
    return res.status(413).json({ error: 'File too large. Maximum size is 10 MB per file.' });
  }
  if (err.code === 'LIMIT_FILE_COUNT') {
    return res.status(400).json({ error: 'Too many files uploaded.' });
  }
  if (err.code === 'UNSUPPORTED_TYPE') {
    return res.status(400).json({ error: err.message });
  }
  console.error('Unhandled error:', err);
  res.status(500).json({ error: err.message || 'Internal server error.' });
});

app.listen(PORT, () => {
  console.log(`\n✅  Attendance Processor running at http://localhost:${PORT}\n`);
});
