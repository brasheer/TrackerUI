// ============================================================================
// scratch-sheets.js — In-App Workbook Scratch Sheet Tabs, Excel-Like Range
//                     Selection, Auto-Scroll, Right-Click Cut/Copy/Paste,
//                     Excel Cell Fill Colors, Striped Rows, Data to Columns & Grid
// ============================================================================

// --- SCRATCH SHEET EXTENDED INTERACTION STATE (v1.16.0 - v1.19.0) ---
const SCRATCH_MAX_UNDO_STEPS = 50;
const SCRATCH_STRIPED_STORAGE_KEY = 'trackerui_scratch_striped_rows';
const SCRATCH_AUTO_INCREMENT_STORAGE_KEY = 'trackerui_scratch_auto_increment';
const SCRATCH_LAST_FILL_STORAGE_KEY = 'trackerui_scratch_last_fill_id';

const SCRATCH_FORMAT_DEFS = {
  auto: { id: 'auto', label: 'Auto', numFmt: null },
  text: { id: 'text', label: 'Text (@)', numFmt: '@' },
  date: { id: 'date', label: 'Date (MM/DD/YYYY)', numFmt: 'mm/dd/yyyy' },
  number: { id: 'number', label: 'Number (123)', numFmt: '0' },
  currency: { id: 'currency', label: 'Currency ($)', numFmt: '$#,##0.00' }
};

const scratchSelectionRangeMap = new Map(); // sheetName -> { anchorRow, anchorCol, focusRow, focusCol }
const scratchUndoStackMap = new Map(); // sheetName -> Array<{ changes, selectionBefore, selectionAfter, coalesceKey? }>
const scratchRedoStackMap = new Map(); // sheetName -> Array<{ changes, selectionBefore, selectionAfter }>
const scratchPendingColWidths = new Map(); // sheetName -> Map<colNumber, excelWidth>
const scratchPendingFills = new Map(); // sheetName -> Map<'row::col', { fill, font, numFmt? }>
const scratchSpillStateMap = new Map(); // sheetName -> Map<'anchorR::anchorC', Array<{ r, c }>>
const scratchSpilledCellsMap = new Map(); // sheetName -> Map<'r::c', { display, rawVal, anchorR, anchorC, formula, isError? }>
let scratchClipboardRangeState = null; // { sheetName, minRow, maxRow, minCol, maxCol } | null
let scratchInternalClipboardText = '';
let scratchInternalClipboardState = null; // { tsv, html, styleMatrix, rawEditMatrix, sourceRange } | null
let scratchDragState = null; // { mode: 'cells'|'rows'|'cols'|'fill', sheetName, clientX, clientY, rafId, lastExpandAt, sourceRange?, fillTargetRange? } | null
let scratchColResizeState = null; // { sheetName, col, thEl, resizerEl, startX, startWidthPx, currentWidthPx } | null
let lastScratchNativeCopyAt = 0;
let lastFormulaCoalesceMeta = null; // { key, at }
let scratchMutationVersion = 0;
let scratchStripedRowsEnabled = localStorage.getItem(SCRATCH_STRIPED_STORAGE_KEY) !== 'false';
let scratchAutoIncrementEnabled = localStorage.getItem(SCRATCH_AUTO_INCREMENT_STORAGE_KEY) === 'true';
let activeScratchFillPresetId = localStorage.getItem(SCRATCH_LAST_FILL_STORAGE_KEY) || 'gold_80';
let activeTextToColumnsState = null; // { sheetName, col, minRow, maxRow, sourceRows: [{ r, raw }] } | null
let activeScratchAutoFilterState = null; // { sheetName, headers: [{ col, letter, label }] } | null

// Standard Microsoft Excel Default Fill Colors (Cell Styles, Theme Accents & Standard Palette)
const EXCEL_DEFAULT_FILL_PALETTE = [
  // Group 1: Excel Cell Styles & Light Header Fills
  { id: 'gray_light', group: 'pastel', label: 'Light Gray (Header)', hex: '#D9D9D9', fillArgb: 'FFD9D9D9', fontArgb: 'FF000000', quick: true },
  { id: 'slate_light', group: 'pastel', label: 'Ice Slate (Header)', hex: '#E2E8F0', fillArgb: 'FFE2E8F0', fontArgb: 'FF0F172A', quick: false },
  { id: 'blue_80', group: 'pastel', label: 'Blue, Accent 1 (Lighter 80%)', hex: '#D9E1F2', fillArgb: 'FFD9E1F2', fontArgb: 'FF000000', quick: true },
  { id: 'sky_80', group: 'pastel', label: 'Blue, Accent 5 (Lighter 80%)', hex: '#DDEBF7', fillArgb: 'FFDDEBF7', fontArgb: 'FF000000', quick: false },
  { id: 'green_80', group: 'pastel', label: 'Green, Accent 6 (Lighter 80%)', hex: '#E2EFDA', fillArgb: 'FFE2EFDA', fontArgb: 'FF000000', quick: false },
  { id: 'gold_80', group: 'pastel', label: 'Gold, Accent 4 (Lighter 80%)', hex: '#FFF2CC', fillArgb: 'FFFFF2CC', fontArgb: 'FF000000', quick: true },
  { id: 'orange_80', group: 'pastel', label: 'Orange, Accent 2 (Lighter 80%)', hex: '#FCE4D6', fillArgb: 'FFFCE4D6', fontArgb: 'FF000000', quick: false },
  { id: 'blue_60', group: 'pastel', label: 'Blue, Accent 1 (Lighter 60%)', hex: '#B4C6E7', fillArgb: 'FFB4C6E7', fontArgb: 'FF000000', quick: false },
  { id: 'green_60', group: 'pastel', label: 'Green, Accent 6 (Lighter 60%)', hex: '#C6E0B4', fillArgb: 'FFC6E0B4', fontArgb: 'FF000000', quick: false },
  { id: 'excel_good', group: 'pastel', label: 'Excel Good (Light Green)', hex: '#C6EFCE', fillArgb: 'FFC6EFCE', fontArgb: 'FF006100', quick: true },
  { id: 'excel_neutral', group: 'pastel', label: 'Excel Neutral (Light Yellow)', hex: '#FFEB9C', fillArgb: 'FFFFEB9C', fontArgb: 'FF9C6500', quick: false },
  { id: 'excel_bad', group: 'pastel', label: 'Excel Bad (Light Red)', hex: '#FFC7CE', fillArgb: 'FFFFC7CE', fontArgb: 'FF9C0006', quick: true },

  // Group 2: Excel Standard & Bold Header Colors
  { id: 'navy_hdr', group: 'standard', label: 'Dark Blue (Header)', hex: '#1F4E78', fillArgb: 'FF1F4E78', fontArgb: 'FFFFFFFF', quick: true },
  { id: 'slate_hdr', group: 'standard', label: 'Blue-Gray, Text 2 (Header)', hex: '#44546A', fillArgb: 'FF44546A', fontArgb: 'FFFFFFFF', quick: false },
  { id: 'blue_acc1', group: 'standard', label: 'Blue, Accent 1', hex: '#4472C4', fillArgb: 'FF4472C4', fontArgb: 'FFFFFFFF', quick: false },
  { id: 'std_lt_blue', group: 'standard', label: 'Standard Light Blue', hex: '#00B0F0', fillArgb: 'FF00B0F0', fontArgb: 'FFFFFFFF', quick: false },
  { id: 'green_acc6', group: 'standard', label: 'Green, Accent 6', hex: '#70AD47', fillArgb: 'FF70AD47', fontArgb: 'FFFFFFFF', quick: false },
  { id: 'std_green', group: 'standard', label: 'Standard Green', hex: '#00B050', fillArgb: 'FF00B050', fontArgb: 'FFFFFFFF', quick: false },
  { id: 'std_yellow', group: 'standard', label: 'Standard Yellow Highlight', hex: '#FFFF00', fillArgb: 'FFFFFF00', fontArgb: 'FF000000', quick: true },
  { id: 'gold_acc4', group: 'standard', label: 'Gold, Accent 4', hex: '#FFC000', fillArgb: 'FFFFC000', fontArgb: 'FF000000', quick: false },
  { id: 'orange_acc2', group: 'standard', label: 'Orange, Accent 2', hex: '#ED7D31', fillArgb: 'FFED7D31', fontArgb: 'FFFFFFFF', quick: false },
  { id: 'std_red', group: 'standard', label: 'Standard Red', hex: '#FF0000', fillArgb: 'FFFF0000', fontArgb: 'FFFFFFFF', quick: false },
  { id: 'dark_red', group: 'standard', label: 'Standard Dark Red', hex: '#C00000', fillArgb: 'FFC00000', fontArgb: 'FFFFFFFF', quick: false },
  { id: 'std_purple', group: 'standard', label: 'Standard Purple', hex: '#7030A0', fillArgb: 'FF7030A0', fontArgb: 'FFFFFFFF', quick: false }
];

// Standard Office OpenXML 10-slot Theme Base Colors (ECMA-376 Office Theme)
const OFFICE_THEME_BASE_HEX = [
  '#FFFFFF', // 0: lt1
  '#000000', // 1: dk1
  '#E7E6E6', // 2: lt2
  '#44546A', // 3: dk2
  '#4472C4', // 4: accent1
  '#ED7D31', // 5: accent2
  '#A5A5A5', // 6: accent3
  '#FFC000', // 7: accent4
  '#5B9BD5', // 8: accent5
  '#70AD47'  // 9: accent6
];

function applyOpenXmlTintToHex(baseHex, tint) {
  const clean = String(baseHex || '').replace(/^#/, '');
  if (clean.length !== 6) return null;
  let r = parseInt(clean.slice(0, 2), 16);
  let g = parseInt(clean.slice(2, 4), 16);
  let b = parseInt(clean.slice(4, 6), 16);
  if (typeof tint === 'number' && Number.isFinite(tint) && tint !== 0) {
    const applyCh = (ch) => {
      const v = tint < 0 ? ch * (1 + tint) : ch * (1 - tint) + 255 * tint;
      return Math.max(0, Math.min(255, Math.round(v)));
    };
    r = applyCh(r);
    g = applyCh(g);
    b = applyCh(b);
  }
  return `#${r.toString(16).padStart(2, '0')}${g.toString(16).padStart(2, '0')}${b.toString(16).padStart(2, '0')}`.toUpperCase();
}

function resolveExcelColorToHex(colorObj) {
  if (!colorObj || typeof colorObj !== 'object') return null;
  if (typeof colorObj.argb === 'string') {
    const raw = colorObj.argb.trim().replace(/^#/, '');
    if (raw.length === 8) {
      const alpha = raw.slice(0, 2).toUpperCase();
      if (alpha === '00') return null;
      return `#${raw.slice(2).toUpperCase()}`;
    }
    if (raw.length === 6) {
      return `#${raw.toUpperCase()}`;
    }
  }
  if (typeof colorObj.theme === 'number' && colorObj.theme >= 0 && colorObj.theme < OFFICE_THEME_BASE_HEX.length) {
    return applyOpenXmlTintToHex(OFFICE_THEME_BASE_HEX[colorObj.theme], colorObj.tint || 0);
  }
  return null;
}

function getContrastingTextHexForBg(bgHex) {
  const clean = String(bgHex || '').replace(/^#/, '');
  if (clean.length !== 6) return '#0F172A';
  const r = parseInt(clean.slice(0, 2), 16);
  const g = parseInt(clean.slice(2, 4), 16);
  const b = parseInt(clean.slice(4, 6), 16);
  const lum = 0.299 * r + 0.587 * g + 0.114 * b;
  return lum > 148 ? '#0F172A' : '#FFFFFF';
}

function extractScratchCellVisualStyle(ws, rNum, cNum) {
  if (!ws) return { bgHex: null, fgHex: null, isBold: false, align: null };
  const row = ws.findRow ? ws.findRow(rNum) : null;
  if (!row) return { bgHex: null, fgHex: null, isBold: false, align: null };
  const cell = row.findCell ? row.findCell(cNum) : null;
  if (!cell) return { bgHex: null, fgHex: null, isBold: false, align: null };

  let bgHex = null;
  const fill = cell.fill;
  if (fill && fill.type === 'pattern' && fill.pattern === 'solid') {
    bgHex = resolveExcelColorToHex(fill.fgColor) || resolveExcelColorToHex(fill.bgColor);
  }

  const font = cell.font;
  const isBold = Boolean(font && font.bold);
  const explicitFontHex = font ? resolveExcelColorToHex(font.color) : null;
  const fgHex = bgHex
    ? (explicitFontHex || getContrastingTextHexForBg(bgHex))
    : explicitFontHex;
  const align = (cell.alignment && cell.alignment.horizontal) ? cell.alignment.horizontal : null;
  return { bgHex, fgHex, isBold, align };
}

function captureCellStyleSnapshot(ws, rNum, cNum) {
  if (!ws) return { fill: null, font: null, numFmt: null };
  const row = ws.findRow ? ws.findRow(rNum) : null;
  const cell = row && row.findCell ? row.findCell(cNum) : null;
  const numFmt = (cell && typeof cell.numFmt === 'string' && cell.numFmt.trim() !== '')
    ? cell.numFmt.trim()
    : ((cell && cell.style && typeof cell.style.numFmt === 'string' && cell.style.numFmt.trim() !== '')
        ? cell.style.numFmt.trim()
        : null);
  if (!cell || !cell.style) return { fill: null, font: null, numFmt };
  return {
    fill: cell.style.fill ? JSON.parse(JSON.stringify(cell.style.fill)) : null,
    font: cell.style.font ? JSON.parse(JSON.stringify(cell.style.font)) : null,
    numFmt
  };
}

function buildStyleSnapshotForFillPreset(beforeSnap, presetId) {
  const baseFont = beforeSnap && beforeSnap.font ? JSON.parse(JSON.stringify(beforeSnap.font)) : null;
  const baseNumFmt = beforeSnap && beforeSnap.numFmt ? beforeSnap.numFmt : null;
  if (!presetId || presetId === 'none') {
    if (baseFont && baseFont.color) {
      delete baseFont.color;
    }
    const cleanFont = baseFont && Object.keys(baseFont).length > 0 ? baseFont : null;
    return { fill: null, font: cleanFont, numFmt: baseNumFmt };
  }

  const preset = EXCEL_DEFAULT_FILL_PALETTE.find(p => p.id === presetId);
  if (!preset) return beforeSnap;

  const nextFont = baseFont || {};
  nextFont.color = { argb: preset.fontArgb };
  return {
    fill: {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: preset.fillArgb }
    },
    font: nextFont,
    numFmt: baseNumFmt
  };
}

function applyCellStyleSnapshotToExcelCell(ws, rNum, cNum, snap) {
  if (!ws) return;
  const row = ws.getRow(rNum);
  const cell = row.getCell(cNum);
  const nextStyle = cell.style ? JSON.parse(JSON.stringify(cell.style)) : {};

  if (snap && snap.fill) {
    nextStyle.fill = JSON.parse(JSON.stringify(snap.fill));
  } else {
    delete nextStyle.fill;
  }

  if (snap && snap.font && Object.keys(snap.font).length > 0) {
    nextStyle.font = JSON.parse(JSON.stringify(snap.font));
  } else {
    delete nextStyle.font;
  }

  if (snap && typeof snap.numFmt === 'string' && snap.numFmt.trim() !== '') {
    nextStyle.numFmt = snap.numFmt.trim();
  } else {
    delete nextStyle.numFmt;
  }

  cell.style = nextStyle;
  row.commit();
}

function applyVisualStyleToScratchTd(td, visual) {
  if (!td || !visual) return;
  td.classList.toggle('bold-cell', Boolean(visual.isBold));
  if (visual.bgHex) {
    td.classList.add('has-custom-fill');
    td.classList.remove('has-custom-fg');
    td.style.setProperty('--cell-fill-bg', visual.bgHex);
    td.style.setProperty('--cell-fill-fg', visual.fgHex || getContrastingTextHexForBg(visual.bgHex));
  } else {
    td.classList.remove('has-custom-fill');
    td.style.removeProperty('--cell-fill-bg');
    if (visual.fgHex) {
      td.classList.add('has-custom-fg');
      td.style.setProperty('--cell-fill-fg', visual.fgHex);
    } else {
      td.classList.remove('has-custom-fg');
      td.style.removeProperty('--cell-fill-fg');
    }
  }
}

// --- 6D. IN-APP WORKBOOK SCRATCH SHEET TABS & EDITABLE SPREADSHEET ENGINE ---
function colIndexToExcelLetter(colNumber) {
  let n = Math.max(1, Math.floor(colNumber));
  let letters = '';
  while (n > 0) {
    const rem = (n - 1) % 26;
    letters = String.fromCharCode(65 + rem) + letters;
    n = Math.floor((n - 1) / 26);
  }
  return letters;
}

function getCellCoordString(row, col) {
  return `${colIndexToExcelLetter(col)}${row}`;
}

function getRangeCoordString(range) {
  if (!range) return 'A1';
  if (range.minRow === range.maxRow && range.minCol === range.maxCol) {
    return getCellCoordString(range.minRow, range.minCol);
  }
  return `${getCellCoordString(range.minRow, range.minCol)}:${getCellCoordString(range.maxRow, range.maxCol)}`;
}

function isDateLikeValue(val) {
  return Boolean(val && typeof val === 'object' && (val instanceof Date || Object.prototype.toString.call(val) === '[object Date]'));
}

function formatExcelDateDisplay(dateObj) {
  if (!isDateLikeValue(dateObj) || isNaN(dateObj.getTime())) return '';
  const yyyy = dateObj.getUTCFullYear();
  const mm = String(dateObj.getUTCMonth() + 1).padStart(2, '0');
  const dd = String(dateObj.getUTCDate()).padStart(2, '0');
  return `${mm}/${dd}/${yyyy}`;
}

function parseFlexibleScratchDate(rawVal, { allowTwoDigitHyphen = false } = {}) {
  if (isDateLikeValue(rawVal) && !isNaN(rawVal.getTime())) {
    const year = rawVal.getUTCFullYear();
    const month = rawVal.getUTCMonth() + 1;
    const day = rawVal.getUTCDate();
    const dateUtc = new Date(Date.UTC(year, month - 1, day));
    return {
      __isDate: true,
      dateUtc,
      display: formatExcelDateDisplay(dateUtc),
      year,
      month,
      day,
      dayEpoch: Math.floor(dateUtc.getTime() / 86400000)
    };
  }
  if (typeof rawVal !== 'string') return null;
  const s = rawVal.trim();
  if (!s) return null;

  let year = 0;
  let month = 0;
  let day = 0;

  const isoMatch = s.match(/^(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})$/);
  if (isoMatch) {
    year = parseInt(isoMatch[1], 10);
    month = parseInt(isoMatch[2], 10);
    day = parseInt(isoMatch[3], 10);
  } else {
    const usMatch = s.match(/^(\d{1,2})([\/\-])(\d{1,2})\2(\d{2}|\d{4})$/);
    if (!usMatch) return null;
    const sep = usMatch[2];
    const yStr = usMatch[4];
    if (sep === '-' && yStr.length === 2 && !allowTwoDigitHyphen) {
      return null;
    }
    month = parseInt(usMatch[1], 10);
    day = parseInt(usMatch[3], 10);
    const yNum = parseInt(yStr, 10);
    year = yStr.length === 2 ? (yNum <= 69 ? 2000 + yNum : 1900 + yNum) : yNum;
  }

  if (year < 1900 || year > 2199 || month < 1 || month > 12 || day < 1 || day > 31) {
    return null;
  }
  const dt = new Date(Date.UTC(year, month - 1, day));
  if (dt.getUTCFullYear() !== year || dt.getUTCMonth() !== month - 1 || dt.getUTCDate() !== day) {
    return null;
  }
  return {
    __isDate: true,
    dateUtc: dt,
    display: formatExcelDateDisplay(dt),
    year,
    month,
    day,
    dayEpoch: Math.floor(dt.getTime() / 86400000)
  };
}

function detectScratchFormatIdFromNumFmt(numFmt) {
  const fmt = String(numFmt || '').trim();
  if (!fmt || fmt.toLowerCase() === 'general') return 'auto';
  if (fmt === '@') return 'text';
  const clean = fmt.replace(/"[^"]*"/g, '');
  if (/\$|currency/i.test(fmt)) return 'currency';
  if (/(?:^|[^a-z\\])(mm?|dd?|yy(?:yy)?)(?:[^a-z]|$)/i.test(clean)) return 'date';
  if (/^[#0.,;%?+\-()\s]+$/.test(clean)) return 'number';
  return 'auto';
}

function getScratchCellFormatId(ws, rNum, cNum) {
  if (!ws) return 'auto';
  const row = ws.findRow ? ws.findRow(rNum) : null;
  const cell = row && row.findCell ? row.findCell(cNum) : null;
  if (!cell) return 'auto';
  const fmt = (typeof cell.numFmt === 'string' && cell.numFmt.trim() !== '')
    ? cell.numFmt.trim()
    : ((cell.style && typeof cell.style.numFmt === 'string') ? cell.style.numFmt.trim() : '');
  const detected = detectScratchFormatIdFromNumFmt(fmt);
  if (detected !== 'auto') return detected;
  if (cell.value instanceof Date) return 'date';
  return 'auto';
}

function mapFormulaOutsideQuotes(formulaBody, transformFn) {
  const str = String(formulaBody || '');
  let out = '';
  let inQuotes = false;
  let chunkStart = 0;
  for (let i = 0; i < str.length; i++) {
    if (str[i] === '"') {
      if (inQuotes) {
        if (str[i + 1] === '"') {
          i++;
          continue;
        }
        out += str.slice(chunkStart, i + 1);
        chunkStart = i + 1;
        inQuotes = false;
      } else {
        out += transformFn(str.slice(chunkStart, i));
        chunkStart = i;
        inQuotes = true;
      }
    }
  }
  if (chunkStart < str.length) {
    out += inQuotes ? str.slice(chunkStart) : transformFn(str.slice(chunkStart));
  }
  return out;
}

function toHumanReadableFormula(formulaBody) {
  return mapFormulaOutsideQuotes(formulaBody, (seg) =>
    seg
      .replace(/_xlfn\._xlws\./gi, '')
      .replace(/_xlfn\./gi, '')
      .replace(/_xlpm\./gi, '')
  );
}

function toExcelOpenXmlFormula(formulaBody) {
  const clean = toHumanReadableFormula(formulaBody);
  return mapFormulaOutsideQuotes(clean, (seg) =>
    seg
      .replace(/\b(FILTER|SORT)\s*\(/gi, (_m, fn) => `_xlfn._xlws.${fn.toUpperCase()}(`)
      .replace(/\b(UNIQUE)\s*\(/gi, (_m, fn) => `_xlfn.${fn.toUpperCase()}(`)
  );
}

function formatScratchScalarValue(val, fmtId = 'auto') {
  if (val === null || val === undefined) return '';
  if (val && typeof val === 'object' && val.__isDate && val.display) return val.display;
  if (isDateLikeValue(val)) return formatExcelDateDisplay(val);
  if (typeof val === 'number') {
    if (!Number.isFinite(val)) return String(val);
    if (fmtId === 'currency') {
      const sign = val < 0 ? '-' : '';
      return `${sign}$${Math.abs(val).toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`;
    }
    return String(val);
  }
  if (typeof val === 'boolean') return val ? 'TRUE' : 'FALSE';
  return String(val);
}

function extractScratchCellInfo(ws, rNum, cNum, { skipSpillLookup = false } = {}) {
  if (!ws) return { display: '', rawEdit: '', isFormula: false, isSpilled: false, formatId: 'auto' };
  const sheetName = ws.name || '';
  const coordKey = `${rNum}::${cNum}`;
  const spillMap = !skipSpillLookup && sheetName ? scratchSpilledCellsMap.get(sheetName) : null;
  const spilledEntry = spillMap ? spillMap.get(coordKey) : null;

  const row = ws.findRow ? ws.findRow(rNum) : ws.getRow(rNum);
  const cell = row ? (row.findCell ? row.findCell(cNum) : row.getCell(cNum)) : null;
  const fmtStr = cell
    ? ((typeof cell.numFmt === 'string' && cell.numFmt.trim() !== '')
        ? cell.numFmt.trim()
        : ((cell.style && typeof cell.style.numFmt === 'string') ? cell.style.numFmt.trim() : ''))
    : '';
  let formatId = detectScratchFormatIdFromNumFmt(fmtStr);
  const val = cell ? cell.value : null;

  if (val instanceof Date && formatId === 'auto') {
    formatId = 'date';
  }

  // If this cell belongs to an active dynamic array spill (or cleared Excel cached spill), render live spill output
  if (spilledEntry && !(val && typeof val === 'object' && (val.formula !== undefined || val.sharedFormula !== undefined))) {
    const disp = formatScratchScalarValue(spilledEntry.rawVal, formatId);
    return {
      display: disp,
      rawEdit: disp,
      isFormula: false,
      isSpilled: true,
      anchorR: spilledEntry.anchorR,
      anchorC: spilledEntry.anchorC,
      formatId
    };
  }

  if (val === null || val === undefined) {
    return { display: '', rawEdit: '', isFormula: false, isSpilled: false, formatId };
  }
  if (val instanceof Date) {
    const dStr = formatExcelDateDisplay(val);
    return { display: dStr, rawEdit: dStr, isFormula: false, isSpilled: false, formatId };
  }
  if (typeof val === 'object') {
    if (val.formula !== undefined || val.sharedFormula !== undefined) {
      let resStr = '';
      if (val.result instanceof Date) {
        resStr = formatExcelDateDisplay(val.result);
      } else if (val.result !== undefined && val.result !== null && typeof val.result !== 'object') {
        resStr = formatScratchScalarValue(val.result, formatId);
      } else if (val.result && typeof val.result === 'object' && val.result.error) {
        resStr = String(val.result.error);
      }
      const rawFormula = val.formula
        ? `=${toHumanReadableFormula(val.formula)}`
        : resStr;
      return { display: resStr, rawEdit: rawFormula, isFormula: true, isSpilled: false, formatId };
    }
    if (Array.isArray(val.richText)) {
      const joined = val.richText.map(t => t.text || '').join('');
      return { display: joined, rawEdit: joined, isFormula: false, isSpilled: false, formatId };
    }
    if (val.text !== undefined && val.text !== null) {
      const s = String(val.text);
      return { display: s, rawEdit: s, isFormula: false, isSpilled: false, formatId };
    }
    if (val.error) {
      const eStr = String(val.error);
      return { display: eStr, rawEdit: eStr, isFormula: false, isSpilled: false, formatId };
    }
    return { display: '', rawEdit: '', isFormula: false, isSpilled: false, formatId };
  }

  if (typeof val === 'number') {
    return {
      display: formatScratchScalarValue(val, formatId),
      rawEdit: String(val),
      isFormula: false,
      isSpilled: false,
      formatId
    };
  }

  const str = String(val);
  return { display: str, rawEdit: str, isFormula: false, isSpilled: false, formatId };
}

function measureScratchWorksheetBounds(ws) {
  let maxRow = 0;
  let maxCol = 0;
  if (!ws) return { nonEmptyRows: 0, maxRow: 0, maxCol: 0 };

  const nonEmptyRowSet = new Set();

  ws.eachRow({ includeEmpty: false }, (row, rowNumber) => {
    row.eachCell({ includeEmpty: false }, (_cell, colNumber) => {
      const info = extractScratchCellInfo(ws, rowNumber, colNumber);
      const visual = extractScratchCellVisualStyle(ws, rowNumber, colNumber);
      if (info.display.trim() !== '' || info.isFormula) {
        nonEmptyRowSet.add(rowNumber);
        if (rowNumber > maxRow) maxRow = rowNumber;
        if (colNumber > maxCol) maxCol = colNumber;
      } else if (visual.bgHex) {
        if (rowNumber > maxRow) maxRow = rowNumber;
        if (colNumber > maxCol) maxCol = colNumber;
      }
    });
  });

  const spillMap = ws.name ? scratchSpilledCellsMap.get(ws.name) : null;
  if (spillMap && spillMap.size > 0) {
    for (const [coordKey, entry] of spillMap.entries()) {
      const disp = formatScratchScalarValue(entry.rawVal);
      if (disp.trim() !== '') {
        const [rStr, cStr] = coordKey.split('::');
        const r = parseInt(rStr, 10);
        const c = parseInt(cStr, 10);
        nonEmptyRowSet.add(r);
        if (r > maxRow) maxRow = r;
        if (c > maxCol) maxCol = c;
      }
    }
  }

  return { nonEmptyRows: nonEmptyRowSet.size, maxRow, maxCol };
}

function applyRawEditToExcelCell(ws, rNum, cNum, rawText) {
  if (!ws) return;
  const normalized = String(rawText ?? '').replace(/\r/g, '');
  const row = ws.getRow(rNum);
  const cell = row.getCell(cNum);
  const fmtStr = (typeof cell.numFmt === 'string' && cell.numFmt.trim() !== '')
    ? cell.numFmt.trim()
    : ((cell.style && typeof cell.style.numFmt === 'string') ? cell.style.numFmt.trim() : '');
  const formatId = detectScratchFormatIdFromNumFmt(fmtStr);

  if (normalized === '') {
    cell.value = null;
  } else if (normalized.startsWith('=') && normalized.length > 1) {
    cell.value = { formula: toExcelOpenXmlFormula(normalized.slice(1)) };
  } else if (formatId === 'text') {
    cell.value = normalized;
  } else if (formatId === 'date') {
    const parsed = parseFlexibleScratchDate(normalized, { allowTwoDigitHyphen: true });
    if (parsed) {
      cell.value = parsed.dateUtc;
      if (!fmtStr || fmtStr.toLowerCase() === 'general') {
        cell.numFmt = 'mm/dd/yyyy';
      }
    } else {
      cell.value = normalized;
    }
  } else if (formatId === 'number' || formatId === 'currency') {
    const stripped = normalized.trim().replace(/^\$/, '').replace(/,/g, '');
    const isNum =
      /^-?(?:0|[1-9]\d*)(?:\.\d+)?$/.test(stripped) &&
      stripped.replace(/[-.]/g, '').length <= 12;
    cell.value = isNum ? Number(stripped) : normalized;
  } else {
    const trimmed = normalized.trim();
    const parsedDate = parseFlexibleScratchDate(trimmed, { allowTwoDigitHyphen: false });
    if (parsedDate) {
      cell.value = parsedDate.dateUtc;
      if (!fmtStr || fmtStr.toLowerCase() === 'general') {
        cell.numFmt = 'mm/dd/yyyy';
      }
    } else {
      // Preserve leading-zero codes and 11+ digit carrier tracking numbers (FedEx/UPS/USPS) as exact strings
      // so Excel never formats them into floating-point scientific notation (e.g., 7.74829E+11).
      const isSafeNumber =
        /^-?(?:0|[1-9]\d*)(?:\.\d+)?$/.test(trimmed) &&
        trimmed.replace(/[-.]/g, '').length <= 10;
      cell.value = isSafeNumber ? Number(trimmed) : normalized;
    }
  }
  row.commit();
}

function hasAnyPendingScratchMutations(sheetName = null) {
  if (sheetName) {
    const hasCellEdits = scratchPendingEdits.has(sheetName) && scratchPendingEdits.get(sheetName).size > 0;
    const hasColEdits = scratchPendingColWidths.has(sheetName) && scratchPendingColWidths.get(sheetName).size > 0;
    const hasFillEdits = scratchPendingFills.has(sheetName) && scratchPendingFills.get(sheetName).size > 0;
    return hasCellEdits || hasColEdits || hasFillEdits;
  }
  for (const m of scratchPendingEdits.values()) {
    if (m.size > 0) return true;
  }
  for (const m of scratchPendingColWidths.values()) {
    if (m.size > 0) return true;
  }
  for (const m of scratchPendingFills.values()) {
    if (m.size > 0) return true;
  }
  return false;
}

function getPendingScratchMutationCount(sheetName) {
  const cellCount = scratchPendingEdits.has(sheetName) ? scratchPendingEdits.get(sheetName).size : 0;
  const colCount = scratchPendingColWidths.has(sheetName) ? scratchPendingColWidths.get(sheetName).size : 0;
  const fillCount = scratchPendingFills.has(sheetName) ? scratchPendingFills.get(sheetName).size : 0;
  return cellCount + colCount + fillCount;
}

function reapplyPendingScratchEditsToWorkbook(wb) {
  if (!wb) return;
  // Apply style/format snapshots first so applyRawEditToExcelCell sees each cell's numFmt
  if (scratchPendingFills.size > 0) {
    for (const [sheetName, fillMap] of scratchPendingFills.entries()) {
      const ws = wb.getWorksheet(sheetName);
      if (!ws) continue;
      for (const [coordKey, styleSnap] of fillMap.entries()) {
        const [rStr, cStr] = coordKey.split('::');
        const r = parseInt(rStr, 10);
        const c = parseInt(cStr, 10);
        if (r >= 1 && c >= 1) {
          applyCellStyleSnapshotToExcelCell(ws, r, c, styleSnap);
        }
      }
    }
  }
  if (scratchPendingEdits.size > 0) {
    for (const [sheetName, cellMap] of scratchPendingEdits.entries()) {
      const ws = wb.getWorksheet(sheetName);
      if (!ws) continue;
      for (const [coordKey, rawVal] of cellMap.entries()) {
        const [rStr, cStr] = coordKey.split('::');
        const r = parseInt(rStr, 10);
        const c = parseInt(cStr, 10);
        if (r >= 1 && c >= 1) {
          applyRawEditToExcelCell(ws, r, c, rawVal);
        }
      }
    }
  }
  if (scratchPendingColWidths.size > 0) {
    for (const [sheetName, colMap] of scratchPendingColWidths.entries()) {
      const ws = wb.getWorksheet(sheetName);
      if (!ws) continue;
      for (const [colNum, excelW] of colMap.entries()) {
        if (colNum >= 1 && excelW > 0) {
          ws.getColumn(colNum).width = excelW;
        }
      }
    }
  }
}

function getPersistedOpenScratchTabsMap() {
  try {
    const raw = localStorage.getItem(SCRATCH_OPEN_TABS_STORAGE_KEY);
    if (!raw) return {};
    const parsed = JSON.parse(raw);
    return (parsed && typeof parsed === 'object') ? parsed : {};
  } catch (_) {
    return {};
  }
}

function getScratchPersistenceWorkbookKey() {
  if (!workingFileHandle || !workingFileHandle.name) return '';
  if (!activeProfileId || activeProfileId === DEFAULT_TRACKER_PROFILE_ID) {
    return workingFileHandle.name;
  }
  return `${activeProfileId}::${workingFileHandle.name}`;
}

function persistOpenScratchTabsForWorkbook() {
  const storageWbKey = getScratchPersistenceWorkbookKey();
  if (!storageWbKey) return;
  try {
    const map = getPersistedOpenScratchTabsMap();
    map[storageWbKey] = [...openScratchTabs];
    localStorage.setItem(SCRATCH_OPEN_TABS_STORAGE_KEY, JSON.stringify(map));
  } catch (e) {
    console.warn('Could not persist open scratch tabs:', e);
  }
}

function restoreOpenScratchTabsForWorkbook() {
  const storageWbKey = getScratchPersistenceWorkbookKey();
  if (!storageWbKey) {
    openScratchTabs = [];
    activeWorkspaceTab = DAILY_TRACKER_WS_TAB;
    return;
  }
  const map = getPersistedOpenScratchTabsMap();
  const savedList = Array.isArray(map[storageWbKey])
    ? map[storageWbKey]
    : (Array.isArray(map[workingFileHandle.name]) ? map[workingFileHandle.name] : []);
  const validNamesByLower = new Map(scratchSheets.map(s => [s.name.toLowerCase(), s.name]));

  const restored = [];
  savedList.forEach(name => {
    const canonical = validNamesByLower.get(String(name || '').toLowerCase());
    if (canonical && !restored.includes(canonical)) {
      restored.push(canonical);
    }
  });

  // Ensure currently active scratch tab stays open if it still exists in the workbook
  if (activeWorkspaceTab !== DAILY_TRACKER_WS_TAB) {
    const activeCanonical = validNamesByLower.get(activeWorkspaceTab.toLowerCase());
    if (activeCanonical) {
      activeWorkspaceTab = activeCanonical;
      if (!restored.includes(activeCanonical)) {
        restored.push(activeCanonical);
      }
    } else {
      activeWorkspaceTab = DAILY_TRACKER_WS_TAB;
    }
  }

  openScratchTabs = restored;
  persistOpenScratchTabsForWorkbook();
}

function closeScratchMenu() {
  if (scratchSheetDropdown) {
    scratchSheetDropdown.classList.remove('open');
  }
  if (btnOpenScratchMenu) {
    btnOpenScratchMenu.setAttribute('aria-expanded', 'false');
  }
}

function renderScratchMenuDropdown() {
  if (!scratchSheetMenuList) return;
  if (scratchAvailableCountBadge) {
    scratchAvailableCountBadge.textContent = String(scratchSheets.length);
  }

  scratchSheetMenuList.innerHTML = '';
  if (scratchSheets.length === 0) {
    scratchSheetMenuList.innerHTML = `
      <div class="scratch-menu-empty">
        No custom scratch sheets in this workbook yet.<br>
        Click <strong>Create New Scratch Sheet</strong> below to add one (e.g., for Title Tracking).
      </div>
    `;
    return;
  }

  scratchSheets.forEach(info => {
    const isOpen = openScratchTabs.includes(info.name);
    const btn = document.createElement('button');
    btn.type = 'button';
    btn.className = 'scratch-menu-item';
    btn.setAttribute('role', 'menuitem');
    btn.innerHTML = `
      <span class="scratch-menu-item-name">${escapeHtml(info.name)}</span>
      <span class="scratch-menu-item-meta">
        <span>${info.nonEmptyRows} row${info.nonEmptyRows === 1 ? '' : 's'}</span>
        ${isOpen ? '<span class="scratch-open-pill">Open</span>' : ''}
      </span>
    `;
    btn.addEventListener('click', () => {
      closeScratchMenu();
      openScratchSheetInAppTab(info.name);
    });
    scratchSheetMenuList.appendChild(btn);
  });
}

function updateScratchSaveBadgeUI() {
  if (!scratchSaveStateBadge) return;
  const activeSheetDirty =
    activeWorkspaceTab !== DAILY_TRACKER_WS_TAB &&
    hasAnyPendingScratchMutations(activeWorkspaceTab);

  if (scratchIsSaving) {
    scratchSaveStateBadge.textContent = 'Saving...';
    scratchSaveStateBadge.className = 'scratch-save-badge dirty';
  } else if (activeSheetDirty) {
    const editCount = getPendingScratchMutationCount(activeWorkspaceTab);
    scratchSaveStateBadge.textContent = `Unsaved (${editCount}) • Auto-saving...`;
    scratchSaveStateBadge.className = 'scratch-save-badge dirty';
  } else {
    scratchSaveStateBadge.textContent = 'Saved';
    scratchSaveStateBadge.className = 'scratch-save-badge saved';
  }
}

function updateScratchUndoRedoButtonsUI() {
  const isScratch = activeWorkspaceTab !== DAILY_TRACKER_WS_TAB;
  const undoStack = isScratch ? (scratchUndoStackMap.get(activeWorkspaceTab) || []) : [];
  const redoStack = isScratch ? (scratchRedoStackMap.get(activeWorkspaceTab) || []) : [];

  if (btnScratchUndo) {
    btnScratchUndo.disabled = undoStack.length === 0;
    btnScratchUndo.title = undoStack.length > 0
      ? `Undo last edit (${undoStack.length} available • Ctrl+Z)`
      : 'Undo last edit (Ctrl+Z)';
  }
  if (btnScratchRedo) {
    btnScratchRedo.disabled = redoStack.length === 0;
    btnScratchRedo.title = redoStack.length > 0
      ? `Redo edit (${redoStack.length} available • Ctrl+Y)`
      : 'Redo edit (Ctrl+Y / Ctrl+Shift+Z)';
  }
}

function renderTrackerProfilesBar() {
  if (!trackerProfilesBar || !trackerProfilesList) return;
  const profilesToRender = Array.isArray(trackerProfiles) && trackerProfiles.length > 0
    ? trackerProfiles
    : [{ id: DEFAULT_TRACKER_PROFILE_ID, name: 'Daily Tracker', mirrorEnabled: true }];

  trackerProfilesBar.style.display = 'flex';
  trackerProfilesList.innerHTML = '';

  profilesToRender.forEach(prof => {
    const isCurrentProf = prof.id === activeProfileId;

    let openCount = null;
    let isCustomDept = false;
    if (isCurrentProf && workbook) {
      openCount = allRows.filter(r => !isCompletedStatus(r.status)).length;
      isCustomDept = Boolean(activeSchemaConfig && !activeSchemaConfig.isStandardSchema);
    } else if (!isCurrentProf && trackerProfileRuntimeCache.has(prof.id)) {
      const snap = trackerProfileRuntimeCache.get(prof.id);
      if (snap && snap.workbook && Array.isArray(snap.allRows)) {
        openCount = snap.allRows.filter(r => !isCompletedStatus(r.status)).length;
        isCustomDept = Boolean(snap.activeSchemaConfig && !snap.activeSchemaConfig.isStandardSchema);
      }
    }

    const canCloseProfile = prof.id !== DEFAULT_TRACKER_PROFILE_ID;
    const countBadgeHtml = openCount !== null
      ? `<span class="ws-tab-pill-count" title="${openCount} open request(s)">${openCount}</span>`
      : '';
    const customTagHtml = isCustomDept
      ? `<span class="ws-tab-schema-tag" title="Dynamic custom column schema active">Custom</span>`
      : '';
    const closeBtnHtml = canCloseProfile
      ? `<span class="ws-tab-close-btn" data-close-profile-tab="${escapeHtml(prof.id)}" title="Remove ${escapeHtml(prof.name)} workbook tab (Excel files stay untouched)" aria-label="Remove ${escapeHtml(prof.name)} tab">×</span>`
      : '';

    const profBtn = document.createElement('button');
    profBtn.type = 'button';
    profBtn.className = `wb-profile-tab${isCurrentProf ? ' active' : ''}`;
    profBtn.setAttribute('role', 'tab');
    profBtn.setAttribute('aria-selected', String(isCurrentProf));
    profBtn.title = `${prof.name}${prof.mirrorEnabled === false ? ' (Mirror Off)' : ''} — Double-click to rename`;
    profBtn.innerHTML = `
      <svg class="ui-icon" viewBox="0 0 20 20" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">
        <path d="M2.5 5.5A1.5 1.5 0 0 1 4 4h3.9a1.5 1.5 0 0 1 1.1.47l1.1 1.06a1.5 1.5 0 0 0 1.05.47H16a1.5 1.5 0 0 1 1.5 1.5v7A1.5 1.5 0 0 1 16 16H4a1.5 1.5 0 0 1-1.5-1.5v-9Z"></path>
      </svg>
      <span>${escapeHtml(prof.name)}</span>
      ${customTagHtml}
      ${countBadgeHtml}
      ${closeBtnHtml}
    `;

    profBtn.addEventListener('click', (e) => {
      const closeProfileEl = e.target.closest('[data-close-profile-tab]');
      if (closeProfileEl) {
        e.stopPropagation();
        if (typeof removeTrackerProfileInteractive === 'function') {
          removeTrackerProfileInteractive(prof.id);
        }
        return;
      }
      if (isCurrentProf) {
        if (activeWorkspaceTab !== DAILY_TRACKER_WS_TAB) {
          switchToWorkspaceTab(DAILY_TRACKER_WS_TAB);
        }
      } else if (typeof switchTrackerProfile === 'function') {
        switchTrackerProfile(prof.id);
      }
    });

    profBtn.addEventListener('dblclick', (e) => {
      if (e.target.closest('[data-close-profile-tab]')) return;
      if (typeof renameActiveTrackerProfileInteractive === 'function') {
        renameActiveTrackerProfileInteractive(prof.id);
      }
    });

    trackerProfilesList.appendChild(profBtn);
  });
}

function renderWorkspaceTabsBar() {
  renderTrackerProfilesBar();

  if (!workspaceTabsBar || !workspaceTabsList) return;
  if (!workbook) {
    workspaceTabsBar.style.display = 'none';
    return;
  }

  workspaceTabsBar.style.display = 'flex';
  if (scratchSheetMenuWrap) {
    scratchSheetMenuWrap.style.display = 'inline-flex';
  }
  renderScratchMenuDropdown();

  workspaceTabsList.innerHTML = '';

  // 1. Active Workbook's Primary Daily Requests Sheet View Tab
  const isTrackerViewActive = activeWorkspaceTab === DAILY_TRACKER_WS_TAB;
  const openCount = allRows.filter(r => !isCompletedStatus(r.status)).length;
  const trackerViewBtn = document.createElement('button');
  trackerViewBtn.type = 'button';
  trackerViewBtn.className = `ws-app-tab${isTrackerViewActive ? ' active' : ''}`;
  trackerViewBtn.setAttribute('role', 'tab');
  trackerViewBtn.setAttribute('aria-selected', String(isTrackerViewActive));
  trackerViewBtn.title = `View daily MM-DD-YYYY request tabs (${openCount} open request(s))`;
  trackerViewBtn.innerHTML = `
    <svg class="ui-icon" viewBox="0 0 20 20" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">
      <rect x="2.5" y="3" width="15" height="14" rx="2"></rect>
      <line x1="2.5" y1="7.5" x2="17.5" y2="7.5"></line>
      <line x1="7" y1="11.5" x2="13.5" y2="11.5"></line>
      <line x1="7" y1="14.2" x2="11.5" y2="14.2"></line>
    </svg>
    <span>Daily Requests</span>
    <span class="ws-tab-pill-count" title="${openCount} open request(s)">${openCount}</span>
  `;
  trackerViewBtn.addEventListener('click', () => {
    switchToWorkspaceTab(DAILY_TRACKER_WS_TAB);
  });
  workspaceTabsList.appendChild(trackerViewBtn);

  // 2. Open Custom / Scratch Sheet App Tabs (belonging to this active workbook)
  openScratchTabs.forEach(sheetName => {
    const isActive = activeWorkspaceTab === sheetName;
    const isDirty = hasAnyPendingScratchMutations(sheetName);
    const meta = scratchSheets.find(s => s.name === sheetName);
    const rowCount = meta ? meta.nonEmptyRows : 0;

    const tabBtn = document.createElement('button');
    tabBtn.type = 'button';
    tabBtn.className = `ws-app-tab${isActive ? ' active' : ''}`;
    tabBtn.setAttribute('role', 'tab');
    tabBtn.setAttribute('aria-selected', String(isActive));
    tabBtn.title = `Scratch Sheet: ${sheetName} (${rowCount} non-empty rows)`;
    tabBtn.innerHTML = `
      <svg class="ui-icon" viewBox="0 0 20 20" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">
        <rect x="3" y="3" width="14" height="14" rx="2"></rect>
        <line x1="3" y1="8" x2="17" y2="8"></line>
        <line x1="8" y1="8" x2="8" y2="17"></line>
      </svg>
      <span>${escapeHtml(sheetName)}</span>
      ${isDirty ? '<span class="ws-tab-dirty-dot" title="Unsaved edits pending"></span>' : `<span class="ws-tab-pill-count">${rowCount}</span>`}
      <span class="ws-tab-close-btn" data-close-scratch-tab="${escapeHtml(sheetName)}" title="Close ${escapeHtml(sheetName)} tab (sheet stays saved in workbook)" aria-label="Close ${escapeHtml(sheetName)} tab">×</span>
    `;

    tabBtn.addEventListener('click', (e) => {
      const closeTarget = e.target.closest('[data-close-scratch-tab]');
      if (closeTarget) {
        e.stopPropagation();
        closeScratchSheetAppTab(sheetName);
        return;
      }
      switchToWorkspaceTab(sheetName);
    });
    workspaceTabsList.appendChild(tabBtn);
  });
}

function closeScratchFillDropdown() {
  if (scratchFillDropdown) {
    scratchFillDropdown.classList.remove('open');
  }
  if (btnScratchFillDropdownToggle) {
    btnScratchFillDropdownToggle.setAttribute('aria-expanded', 'false');
  }
}

function closeScratchContextMenu() {
  if (scratchContextMenu) {
    scratchContextMenu.style.display = 'none';
  }
}

function syncScratchStripedRowsUI() {
  if (chkScratchStripedRows) {
    chkScratchStripedRows.checked = scratchStripedRowsEnabled;
  }
  if (scratchGridTable) {
    scratchGridTable.classList.toggle('striped-rows', scratchStripedRowsEnabled);
  }
}

function syncScratchAutoIncrementUI() {
  if (chkScratchAutoIncrement) {
    chkScratchAutoIncrement.checked = scratchAutoIncrementEnabled;
  }
}

function syncWorkspaceTabsUI() {
  stopScratchPointerDrag();
  closeScratchFillDropdown();
  closeScratchContextMenu();
  syncScratchStripedRowsUI();
  syncScratchAutoIncrementUI();
  renderWorkspaceTabsBar();

  const isDailyActive = activeWorkspaceTab === DAILY_TRACKER_WS_TAB;
  if (dailyTrackerWorkspace) {
    dailyTrackerWorkspace.style.display = isDailyActive ? 'flex' : 'none';
  }
  if (scratchSheetWorkspace) {
    scratchSheetWorkspace.style.display = isDailyActive ? 'none' : 'flex';
  }

  if (!isDailyActive) {
    renderActiveScratchSheetGrid();
  }
  updateScratchUndoRedoButtonsUI();
  updateStatusBarTelemetry();
}

async function switchToWorkspaceTab(targetTab) {
  stopScratchPointerDrag();
  closeScratchFillDropdown();
  closeScratchContextMenu();
  commitActiveScratchInlineEdit();
  if (targetTab === activeWorkspaceTab) return;

  if (hasAnyPendingScratchMutations()) {
    flushPendingScratchEditsToDisk({ reloadAfter: false, silentToast: true });
  }

  activeWorkspaceTab = targetTab;
  syncWorkspaceTabsUI();
}

function setActiveWorkspaceTab(targetTab) {
  return switchToWorkspaceTab(targetTab);
}

function openScratchSheetInAppTab(sheetName) {
  if (!workbook) return;
  const ws = workbook.getWorksheet(sheetName);
  if (!ws) return;

  if (!openScratchTabs.includes(ws.name)) {
    openScratchTabs.push(ws.name);
    persistOpenScratchTabsForWorkbook();
  }
  switchToWorkspaceTab(ws.name);
}

async function closeScratchSheetAppTab(sheetName) {
  stopScratchPointerDrag();
  commitActiveScratchInlineEdit();
  if (hasAnyPendingScratchMutations(sheetName)) {
    await flushPendingScratchEditsToDisk({ reloadAfter: false, silentToast: true });
  }

  const idx = openScratchTabs.indexOf(sheetName);
  if (idx !== -1) {
    openScratchTabs.splice(idx, 1);
    persistOpenScratchTabsForWorkbook();
  }

  if (activeWorkspaceTab === sheetName) {
    const fallbackTab = openScratchTabs[idx] || openScratchTabs[idx - 1] || DAILY_TRACKER_WS_TAB;
    activeWorkspaceTab = fallbackTab;
  }
  syncWorkspaceTabsUI();
}

function getScratchSheetGridDimensions(ws, sheetName) {
  const bounds = measureScratchWorksheetBounds(ws);
  const existing = scratchGridDimensionsMap.get(sheetName) || { rows: 0, cols: 0 };
  const rows = Math.max(bounds.maxRow + 12, existing.rows, 30);
  const cols = Math.max(bounds.maxCol + 3, existing.cols, 10);
  const nextDims = { rows, cols };
  scratchGridDimensionsMap.set(sheetName, nextDims);
  return { ...nextDims, nonEmptyRows: bounds.nonEmptyRows, dataMaxRow: bounds.maxRow, dataMaxCol: bounds.maxCol };
}

function getActiveScratchCell(sheetName) {
  const cur = scratchActiveCellMap.get(sheetName);
  if (cur && cur.row >= 1 && cur.col >= 1) return cur;
  const initial = { row: 1, col: 1 };
  scratchActiveCellMap.set(sheetName, initial);
  return initial;
}

function normalizeScratchRangeObj(raw) {
  const anchorRow = Math.max(1, Math.floor(raw.anchorRow || 1));
  const anchorCol = Math.max(1, Math.floor(raw.anchorCol || 1));
  const focusRow = Math.max(1, Math.floor(raw.focusRow || anchorRow));
  const focusCol = Math.max(1, Math.floor(raw.focusCol || anchorCol));
  const minRow = Math.min(anchorRow, focusRow);
  const maxRow = Math.max(anchorRow, focusRow);
  const minCol = Math.min(anchorCol, focusCol);
  const maxCol = Math.max(anchorCol, focusCol);
  const rowCount = maxRow - minRow + 1;
  const colCount = maxCol - minCol + 1;
  return {
    anchorRow,
    anchorCol,
    focusRow,
    focusCol,
    minRow,
    maxRow,
    minCol,
    maxCol,
    rowCount,
    colCount,
    cellCount: rowCount * colCount
  };
}

function getScratchSelectionRange(sheetName) {
  const existing = scratchSelectionRangeMap.get(sheetName);
  if (existing && existing.anchorRow >= 1 && existing.anchorCol >= 1) {
    return normalizeScratchRangeObj(existing);
  }
  const active = getActiveScratchCell(sheetName);
  const initial = {
    anchorRow: active.row,
    anchorCol: active.col,
    focusRow: active.row,
    focusCol: active.col
  };
  scratchSelectionRangeMap.set(sheetName, initial);
  return normalizeScratchRangeObj(initial);
}

function formatScratchStatNumber(num) {
  if (!Number.isFinite(num)) return '0';
  const rounded = Math.round(num * 100) / 100;
  return rounded.toLocaleString(undefined, { maximumFractionDigits: 2 });
}

function updateScratchSelectionStatsUI() {
  if (!scratchSelectionStats) return;
  if (activeWorkspaceTab === DAILY_TRACKER_WS_TAB || !workbook) {
    scratchSelectionStats.style.display = 'none';
    return;
  }
  const range = getScratchSelectionRange(activeWorkspaceTab);
  if (range.cellCount <= 1) {
    scratchSelectionStats.style.display = 'none';
    return;
  }

  const ws = workbook.getWorksheet(activeWorkspaceTab);
  if (!ws) {
    scratchSelectionStats.style.display = 'none';
    return;
  }

  const parts = [`${range.rowCount}R × ${range.colCount}C`];
  if (range.cellCount <= 25000) {
    let nonEmptyCount = 0;
    let numCount = 0;
    let numSum = 0;

    for (let r = range.minRow; r <= range.maxRow; r++) {
      const row = ws.findRow ? ws.findRow(r) : null;
      if (!row) continue;
      for (let c = range.minCol; c <= range.maxCol; c++) {
        const info = extractScratchCellInfo(ws, r, c);
        const trimmed = info.display.trim();
        if (trimmed === '') continue;
        nonEmptyCount++;

        const stripped = trimmed.replace(/^\$/, '').replace(/,/g, '');
        if (
          stripped !== '' &&
          /^-?(?:0|[1-9]\d*)(?:\.\d+)?$/.test(stripped) &&
          stripped.replace(/[-.]/g, '').length <= 12
        ) {
          const n = Number(stripped);
          if (Number.isFinite(n)) {
            numCount++;
            numSum += n;
          }
        }
      }
    }

    if (nonEmptyCount > 0) {
      parts.push(`Count: ${nonEmptyCount.toLocaleString()}`);
    }
    if (numCount >= 2) {
      parts.push(`Sum: ${formatScratchStatNumber(numSum)}`);
      parts.push(`Avg: ${formatScratchStatNumber(numSum / numCount)}`);
    }
  } else {
    parts.push(`${range.cellCount.toLocaleString()} cells`);
  }

  scratchSelectionStats.textContent = parts.join(' • ');
  scratchSelectionStats.style.display = 'inline-block';
}

// --- CELL DATA TYPE / FORMAT SELECTOR & LIVE FORMULA / SPILL ENGINE (v1.19.0) ---
function syncScratchFormatSelectUI() {
  if (!scratchCellFormatSelect || activeWorkspaceTab === DAILY_TRACKER_WS_TAB || !workbook) return;
  const ws = workbook.getWorksheet(activeWorkspaceTab);
  if (!ws) return;
  const range = getScratchSelectionRange(activeWorkspaceTab);
  const fmtId = getScratchCellFormatId(ws, range.anchorRow, range.anchorCol);
  scratchCellFormatSelect.value = SCRATCH_FORMAT_DEFS[fmtId] ? fmtId : 'auto';
}

function applyScratchFormatToSelection(sheetName, fmtId) {
  if (!workbook || !sheetName || sheetName === DAILY_TRACKER_WS_TAB) return 0;
  const def = SCRATCH_FORMAT_DEFS[fmtId] || SCRATCH_FORMAT_DEFS.auto;
  commitActiveScratchInlineEdit();
  const ws = workbook.getWorksheet(sheetName);
  if (!ws) return 0;

  const range = getScratchSelectionRange(sheetName);
  const dims = getScratchSheetGridDimensions(ws, sheetName);
  const isFullColumnSelection = range.minRow === 1 && range.maxRow >= dims.rows && range.colCount < dims.cols;
  const effectiveMaxRow = isFullColumnSelection ? Math.max(dims.dataMaxRow, dims.rows) : range.maxRow;

  if (isFullColumnSelection) {
    for (let c = range.minCol; c <= range.maxCol; c++) {
      ws.getColumn(c).numFmt = def.numFmt || undefined;
    }
  }

  const edits = [];
  for (let r = range.minRow; r <= effectiveMaxRow; r++) {
    for (let c = range.minCol; c <= range.maxCol; c++) {
      const beforeSnap = captureCellStyleSnapshot(ws, r, c);
      const nextSnap = {
        fill: beforeSnap.fill,
        font: beforeSnap.font,
        numFmt: def.numFmt
      };
      const info = extractScratchCellInfo(ws, r, c, { skipSpillLookup: true });
      const editObj = { r, c, styleSnap: nextSnap };
      if (!info.isFormula && info.rawEdit !== '' && !(isFullColumnSelection && r === 1)) {
        editObj.rawText = info.rawEdit;
      }
      edits.push(editObj);
    }
  }

  const count = applyScratchCellChanges(sheetName, edits, { recordUndo: true });
  syncScratchFormatSelectUI();
  if (count > 0) {
    showToast(`Set cell type to ${def.label} on ${count} cell${count === 1 ? '' : 's'}.`, 'info');
  }
  return count;
}

function excelLetterToColIndex(letters) {
  const clean = String(letters || '').trim().toUpperCase();
  if (!clean) return 0;
  let col = 0;
  for (let i = 0; i < clean.length; i++) {
    const code = clean.charCodeAt(i);
    if (code < 65 || code > 90) return 0;
    col = col * 26 + (code - 64);
  }
  return col;
}

function shiftFormulaCellReferences(formulaStr, rowDelta, colDelta) {
  if (!formulaStr || (rowDelta === 0 && colDelta === 0)) return formulaStr;
  return mapFormulaOutsideQuotes(formulaStr, (seg) =>
    seg.replace(/(^|[^A-Za-z0-9_])(\$?)([A-Za-z]{1,3})(\$?)(\d+)(?![A-Za-z0-9_(])/g, (full, prefix, absCol, colLetters, absRow, rowStr) => {
      const cNum = excelLetterToColIndex(colLetters);
      const rNum = parseInt(rowStr, 10);
      if (cNum < 1 || cNum > 16384 || rNum < 1 || rNum > 1048576) return full;
      const nextCol = absCol ? cNum : Math.max(1, cNum + colDelta);
      const nextRow = absRow ? rNum : Math.max(1, rNum + rowDelta);
      return `${prefix}${absCol}${colIndexToExcelLetter(nextCol)}${absRow}${nextRow}`;
    })
  );
}

function makeFormulaMatrix(rows) {
  return { __isMatrix: true, rows };
}

function isFormulaMatrix(val) {
  return Boolean(val && typeof val === 'object' && val.__isMatrix === true && Array.isArray(val.rows));
}

function makeFormulaError(code) {
  return { __isError: true, error: code };
}

function isFormulaError(val) {
  return Boolean(val && typeof val === 'object' && val.__isError === true);
}

function isEmptyFormulaVal(val) {
  return val === null || val === undefined || val === '';
}

function unwrapFormulaDate(val) {
  if (val instanceof Date && !isNaN(val.getTime())) {
    return parseFlexibleScratchDate(val);
  }
  if (val && typeof val === 'object' && val.__isDate) {
    return val;
  }
  return null;
}

function formulaValToString(val) {
  if (isFormulaError(val)) return val.error;
  if (val === null || val === undefined) return '';
  if (val instanceof Date) return formatExcelDateDisplay(val);
  if (typeof val === 'object' && val.__isDate) return val.display;
  if (typeof val === 'boolean') return val ? 'TRUE' : 'FALSE';
  return String(val);
}

function formulaValToNumber(val) {
  if (isFormulaError(val)) return val;
  if (val === null || val === undefined || val === '') return 0;
  if (typeof val === 'number') return val;
  if (typeof val === 'boolean') return val ? 1 : 0;
  if (val instanceof Date) return Math.floor(val.getTime() / 86400000);
  if (typeof val === 'object' && val.__isDate) return val.dayEpoch;
  const s = String(val).trim().replace(/^\$/, '').replace(/,/g, '');
  if (s !== '' && /^-?(?:0|[1-9]\d*)(?:\.\d+)?$/.test(s)) {
    return Number(s);
  }
  return makeFormulaError('#VALUE!');
}

function formulaValToBool(val) {
  if (isFormulaError(val)) return false;
  if (typeof val === 'boolean') return val;
  if (typeof val === 'number') return val !== 0;
  if (typeof val === 'string') {
    const up = val.trim().toUpperCase();
    if (up === 'TRUE') return true;
    if (up === 'FALSE' || up === '') return false;
    return false;
  }
  return Boolean(val);
}

function compareFormulaScalars(a, b, op) {
  if (isFormulaError(a)) return a;
  if (isFormulaError(b)) return b;

  const dA = unwrapFormulaDate(a);
  const dB = unwrapFormulaDate(b);
  if (dA && dB) {
    const cmp = dA.dayEpoch - dB.dayEpoch;
    return evalCmpOp(cmp, op);
  }

  const aNull = a === null || a === undefined;
  const bNull = b === null || b === undefined;
  if (aNull && bNull) return evalCmpOp(0, op);
  if (aNull) {
    if (b === '') return evalCmpOp(0, op);
    if (typeof b === 'number') return evalCmpOp(0 - b, op);
    if (typeof b === 'boolean') return evalCmpOp(0 - (b ? 1 : 0), op);
    return evalCmpOp(''.localeCompare(formulaValToString(b).toLowerCase()), op);
  }
  if (bNull) {
    if (a === '') return evalCmpOp(0, op);
    if (typeof a === 'number') return evalCmpOp(a - 0, op);
    if (typeof a === 'boolean') return evalCmpOp((a ? 1 : 0) - 0, op);
    return evalCmpOp(formulaValToString(a).toLowerCase().localeCompare(''), op);
  }

  if (typeof a === 'number' && typeof b === 'number') {
    return evalCmpOp(a - b, op);
  }
  if (typeof a === 'boolean' && typeof b === 'boolean') {
    return evalCmpOp((a ? 1 : 0) - (b ? 1 : 0), op);
  }
  const sA = formulaValToString(a).toLowerCase();
  const sB = formulaValToString(b).toLowerCase();
  if (typeof a !== typeof b) {
    if (op === '=' || op === '==' || op === '<>' || op === '!=') {
      // Numbers vs non-matching strings are unequal in Excel (e.g., 123 <> "Duplicate" -> TRUE)
      return (op === '<>' || op === '!=') ? true : false;
    }
  }
  return evalCmpOp(sA === sB ? 0 : (sA < sB ? -1 : 1), op);
}

function evalCmpOp(cmp, op) {
  switch (op) {
    case '=':
    case '==': return cmp === 0;
    case '<>':
    case '!=': return cmp !== 0;
    case '<': return cmp < 0;
    case '<=': return cmp <= 0;
    case '>': return cmp > 0;
    case '>=': return cmp >= 0;
    default: return false;
  }
}

function broadcastBinaryOp(left, right, fn) {
  const lMat = isFormulaMatrix(left);
  const rMat = isFormulaMatrix(right);
  if (!lMat && !rMat) {
    return fn(left, right);
  }
  const lRows = lMat ? left.rows : [[left]];
  const rRows = rMat ? right.rows : [[right]];
  const rowCount = Math.max(lRows.length, rRows.length);
  const colCount = Math.max(lRows[0] ? lRows[0].length : 1, rRows[0] ? rRows[0].length : 1);
  const out = [];
  for (let r = 0; r < rowCount; r++) {
    const lr = lRows[r % lRows.length];
    const rr = rRows[r % rRows.length];
    const rowOut = [];
    for (let c = 0; c < colCount; c++) {
      const lv = lr[c % lr.length];
      const rv = rr[c % rr.length];
      rowOut.push(fn(lv, rv));
    }
    out.push(rowOut);
  }
  return makeFormulaMatrix(out);
}

function broadcastUnaryOp(val, fn) {
  if (!isFormulaMatrix(val)) return fn(val);
  return makeFormulaMatrix(val.rows.map(row => row.map(cell => fn(cell))));
}

function tokenizeScratchFormula(expr) {
  const tokens = [];
  const s = String(expr || '').trim();
  let i = 0;
  const len = s.length;

  while (i < len) {
    const ch = s[i];
    if (/\s/.test(ch)) {
      i++;
      continue;
    }

    // String literal "..."
    if (ch === '"') {
      let str = '';
      i++;
      while (i < len) {
        if (s[i] === '"') {
          if (s[i + 1] === '"') {
            str += '"';
            i += 2;
            continue;
          }
          i++;
          break;
        }
        str += s[i];
        i++;
      }
      tokens.push({ type: 'STRING', value: str });
      continue;
    }

    // Sheet-quoted prefix 'Sheet Name'!
    if (ch === "'") {
      let sheet = '';
      let j = i + 1;
      while (j < len) {
        if (s[j] === "'") {
          if (s[j + 1] === "'") {
            sheet += "'";
            j += 2;
            continue;
          }
          j++;
          break;
        }
        sheet += s[j];
        j++;
      }
      if (s[j] === '!') {
        i = j + 1;
        tokens.push({ type: 'SHEET_PREFIX', value: sheet });
        continue;
      }
    }

    // Two-character comparison operators
    const two = s.slice(i, i + 2);
    if (two === '<>' || two === '!=' || two === '<=' || two === '>=') {
      tokens.push({ type: 'OP', value: two });
      i += 2;
      continue;
    }

    if ('=<>+-*/&'.includes(ch)) {
      tokens.push({ type: 'OP', value: ch });
      i++;
      continue;
    }

    if (ch === '(' || ch === ')' || ch === ',' || ch === ':') {
      tokens.push({ type: ch, value: ch });
      i++;
      continue;
    }

    // Number literal
    if (/\d/.test(ch) || (ch === '.' && /\d/.test(s[i + 1] || ''))) {
      let numStr = '';
      while (i < len && /[\d.]/.test(s[i])) {
        numStr += s[i++];
      }
      tokens.push({ type: 'NUMBER', value: Number(numStr) });
      continue;
    }

    // Identifier, function name, or cell reference ($A$1 / Sheet1!A1)
    if (/[A-Za-z_$]/.test(ch)) {
      let ident = '';
      while (i < len && /[A-Za-z0-9_.$]/.test(s[i])) {
        ident += s[i++];
      }
      if (s[i] === '!') {
        i++;
        tokens.push({ type: 'SHEET_PREFIX', value: ident });
        continue;
      }
      tokens.push({ type: 'IDENT', value: ident });
      continue;
    }

    // Unknown char: emit as error token
    tokens.push({ type: 'ERR', value: ch });
    i++;
  }

  return tokens;
}

function parseCellRefParts(tokenStr) {
  const m = String(tokenStr || '').match(/^\$?([A-Za-z]{1,3})\$?(\d+)$/);
  if (!m) return null;
  const col = excelLetterToColIndex(m[1]);
  const row = parseInt(m[2], 10);
  if (col < 1 || col > 16384 || row < 1 || row > 1048576) return null;
  return { row, col };
}

function readCellValueForFormula(ws, rNum, cNum, ctx) {
  if (!ws) return null;
  const sheetName = ws.name || '';
  const coordKey = `${sheetName}::${rNum}::${cNum}`;

  // Avoid infinite circular reference loops
  if (ctx.visited.has(coordKey)) {
    return makeFormulaError('#CIRC!');
  }

  // If already computed in spill cache for this sheet pass, return it
  const sheetSpills = ctx.spillBySheet.get(sheetName);
  if (sheetSpills && sheetSpills.has(`${rNum}::${cNum}`)) {
    return sheetSpills.get(`${rNum}::${cNum}`).rawVal;
  }

  // If this cell was a stale Excel cached spill cell cleared by an array formula on this sheet, return null
  const origSpillSet = ctx.origSpillBySheet.get(sheetName);
  if (origSpillSet && origSpillSet.has(`${rNum}::${cNum}`)) {
    return null;
  }

  const row = ws.findRow ? ws.findRow(rNum) : null;
  const cell = row && row.findCell ? row.findCell(cNum) : null;
  if (!cell || cell.value === null || cell.value === undefined) return null;

  const val = cell.value;
  if (val instanceof Date) {
    return parseFlexibleScratchDate(val);
  }
  if (typeof val === 'object') {
    if (val.formula !== undefined) {
      return evaluateWorksheetCellFormula(ws, rNum, cNum, ctx);
    }
    if (Array.isArray(val.richText)) {
      return val.richText.map(t => t.text || '').join('');
    }
    if (val.text !== undefined && val.text !== null) return String(val.text);
    if (val.error) return makeFormulaError(String(val.error));
    return null;
  }
  return val;
}

function evalFormulaTokens(tokens, ctx) {
  let pos = 0;

  const peek = () => tokens[pos] || null;
  const consume = () => tokens[pos++] || null;

  function parseComparison() {
    let left = parseConcat();
    while (peek() && peek().type === 'OP' && ['=', '<>', '!=', '<', '<=', '>', '>='].includes(peek().value)) {
      const op = consume().value;
      const right = parseConcat();
      left = broadcastBinaryOp(left, right, (a, b) => compareFormulaScalars(a, b, op));
    }
    return left;
  }

  function parseConcat() {
    let left = parseAddSub();
    while (peek() && peek().type === 'OP' && peek().value === '&') {
      consume();
      const right = parseAddSub();
      left = broadcastBinaryOp(left, right, (a, b) => {
        if (isFormulaError(a)) return a;
        if (isFormulaError(b)) return b;
        return formulaValToString(a) + formulaValToString(b);
      });
    }
    return left;
  }

  function parseAddSub() {
    let left = parseMulDiv();
    while (peek() && peek().type === 'OP' && (peek().value === '+' || peek().value === '-')) {
      const op = consume().value;
      const right = parseMulDiv();
      left = broadcastBinaryOp(left, right, (a, b) => {
        const na = formulaValToNumber(a);
        const nb = formulaValToNumber(b);
        if (isFormulaError(na)) return na;
        if (isFormulaError(nb)) return nb;
        return op === '+' ? na + nb : na - nb;
      });
    }
    return left;
  }

  function parseMulDiv() {
    let left = parseUnary();
    while (peek() && peek().type === 'OP' && (peek().value === '*' || peek().value === '/')) {
      const op = consume().value;
      const right = parseUnary();
      left = broadcastBinaryOp(left, right, (a, b) => {
        const na = formulaValToNumber(a);
        const nb = formulaValToNumber(b);
        if (isFormulaError(na)) return na;
        if (isFormulaError(nb)) return nb;
        if (op === '/') return nb === 0 ? makeFormulaError('#DIV/0!') : na / nb;
        return na * nb;
      });
    }
    return left;
  }

  function parseUnary() {
    if (peek() && peek().type === 'OP' && (peek().value === '+' || peek().value === '-')) {
      const op = consume().value;
      const val = parseUnary();
      if (op === '+') return val;
      return broadcastUnaryOp(val, (v) => {
        const n = formulaValToNumber(v);
        return isFormulaError(n) ? n : -n;
      });
    }
    return parsePrimary();
  }

  function parseArgExpressions() {
    const args = [];
    if (peek() && peek().type === ')') {
      consume();
      return args;
    }
    while (pos < tokens.length) {
      args.push(parseComparison());
      if (peek() && peek().type === ',') {
        consume();
        continue;
      }
      if (peek() && peek().type === ')') {
        consume();
        break;
      }
      break;
    }
    return args;
  }

  function flattenMatrixValues(argVal) {
    if (!isFormulaMatrix(argVal)) return [argVal];
    const flat = [];
    argVal.rows.forEach(r => r.forEach(c => flat.push(c)));
    return flat;
  }

  function evalBuiltInFunction(fnNameUpper, args) {
    switch (fnNameUpper) {
      case 'FILTER': {
        if (args.length < 2) return makeFormulaError('#VALUE!');
        const arrArg = args[0];
        const incArg = args[1];
        const ifEmptyArg = args.length >= 3 ? args[2] : undefined;
        const arrRows = isFormulaMatrix(arrArg) ? arrArg.rows : [[arrArg]];
        const incRows = isFormulaMatrix(incArg) ? incArg.rows : [[incArg]];

        const matchedRows = [];
        // Row-wise filter (standard N x M array with N x 1 include mask)
        if (incRows.length === arrRows.length || incRows.length === 1) {
          for (let r = 0; r < arrRows.length; r++) {
            const condCell = incRows.length === 1 ? incRows[0][r % incRows[0].length] : incRows[r][0];
            if (formulaValToBool(condCell)) {
              // Skip completely blank trailing rows in A2:C200 when all columns of the source row are blank
              const rowHasContent = arrRows[r].some(cell => !isEmptyFormulaVal(cell));
              if (rowHasContent) {
                matchedRows.push([...arrRows[r]]);
              }
            }
          }
        } else if (arrRows.length === 1 && incRows.length === 1 && incRows[0].length === arrRows[0].length) {
          const filteredRow = [];
          for (let c = 0; c < arrRows[0].length; c++) {
            if (formulaValToBool(incRows[0][c])) {
              filteredRow.push(arrRows[0][c]);
            }
          }
          if (filteredRow.length > 0) matchedRows.push(filteredRow);
        }

        if (matchedRows.length === 0) {
          if (ifEmptyArg !== undefined) {
            return isFormulaMatrix(ifEmptyArg) ? ifEmptyArg : makeFormulaMatrix([[ifEmptyArg]]);
          }
          return makeFormulaMatrix([[makeFormulaError('#CALC!')]]);
        }
        return makeFormulaMatrix(matchedRows);
      }

      case 'UNIQUE': {
        if (args.length < 1) return makeFormulaError('#VALUE!');
        const rows = isFormulaMatrix(args[0]) ? args[0].rows : [[args[0]]];
        const seen = new Set();
        const out = [];
        rows.forEach(r => {
          const key = r.map(v => formulaValToString(v).toLowerCase()).join('\u0001');
          if (!seen.has(key)) {
            seen.add(key);
            out.push([...r]);
          }
        });
        return makeFormulaMatrix(out.length > 0 ? out : [['']]);
      }

      case 'SORT': {
        if (args.length < 1) return makeFormulaError('#VALUE!');
        const rows = (isFormulaMatrix(args[0]) ? args[0].rows : [[args[0]]]).map(r => [...r]);
        const sortIdx = args.length >= 2 ? Math.max(1, Math.floor(Number(formulaValToNumber(args[1])) || 1)) : 1;
        const sortOrder = args.length >= 3 ? (Number(formulaValToNumber(args[2])) < 0 ? -1 : 1) : 1;
        const colIdx = sortIdx - 1;
        rows.sort((rA, rB) => {
          const a = rA[colIdx];
          const b = rB[colIdx];
          const eA = isEmptyFormulaVal(a);
          const eB = isEmptyFormulaVal(b);
          if (eA && eB) return 0;
          if (eA) return 1;
          if (eB) return -1;
          const dA = unwrapFormulaDate(a);
          const dB = unwrapFormulaDate(b);
          if (dA && dB) return (dA.dayEpoch - dB.dayEpoch) * sortOrder;
          if (typeof a === 'number' && typeof b === 'number') return (a - b) * sortOrder;
          return formulaValToString(a).localeCompare(formulaValToString(b), undefined, { numeric: true, sensitivity: 'base' }) * sortOrder;
        });
        return makeFormulaMatrix(rows);
      }

      case 'IF': {
        if (args.length < 2) return makeFormulaError('#VALUE!');
        const cond = args[0];
        const tVal = args[1];
        const fVal = args.length >= 3 ? args[2] : false;
        if (!isFormulaMatrix(cond) && !isFormulaMatrix(tVal) && !isFormulaMatrix(fVal)) {
          return formulaValToBool(cond) ? tVal : fVal;
        }
        return broadcastBinaryOp(cond, tVal, (cVal, tv) => (formulaValToBool(cVal) ? tv : (!isFormulaMatrix(fVal) ? fVal : '')));
      }

      case 'ISBLANK':
      case 'ISEMPTY': {
        if (args.length < 1) return makeFormulaError('#VALUE!');
        return broadcastUnaryOp(args[0], (v) => isEmptyFormulaVal(v));
      }

      case 'ISNUMBER': {
        if (args.length < 1) return makeFormulaError('#VALUE!');
        return broadcastUnaryOp(args[0], (v) => typeof v === 'number' && Number.isFinite(v));
      }

      case 'ISTEXT': {
        if (args.length < 1) return makeFormulaError('#VALUE!');
        return broadcastUnaryOp(args[0], (v) => typeof v === 'string' && v !== '');
      }

      case 'NOT': {
        if (args.length < 1) return makeFormulaError('#VALUE!');
        return broadcastUnaryOp(args[0], (v) => !formulaValToBool(v));
      }

      case 'AND': {
        for (const arg of args) {
          const vals = flattenMatrixValues(arg);
          if (vals.some(v => !formulaValToBool(v))) return false;
        }
        return args.length > 0;
      }

      case 'OR': {
        for (const arg of args) {
          const vals = flattenMatrixValues(arg);
          if (vals.some(v => formulaValToBool(v))) return true;
        }
        return false;
      }

      case 'TRIM': {
        if (args.length < 1) return '';
        return broadcastUnaryOp(args[0], (v) => formulaValToString(v).trim().replace(/\s+/g, ' '));
      }

      case 'UPPER': {
        if (args.length < 1) return '';
        return broadcastUnaryOp(args[0], (v) => formulaValToString(v).toUpperCase());
      }

      case 'LOWER': {
        if (args.length < 1) return '';
        return broadcastUnaryOp(args[0], (v) => formulaValToString(v).toLowerCase());
      }

      case 'LEN': {
        if (args.length < 1) return 0;
        return broadcastUnaryOp(args[0], (v) => formulaValToString(v).length);
      }

      case 'LEFT': {
        if (args.length < 1) return '';
        const n = args.length >= 2 ? Math.max(0, Math.floor(Number(formulaValToNumber(args[1])) || 0)) : 1;
        return broadcastUnaryOp(args[0], (v) => formulaValToString(v).slice(0, n));
      }

      case 'RIGHT': {
        if (args.length < 1) return '';
        const n = args.length >= 2 ? Math.max(0, Math.floor(Number(formulaValToNumber(args[1])) || 0)) : 1;
        return broadcastUnaryOp(args[0], (v) => {
          const s = formulaValToString(v);
          return n === 0 ? '' : s.slice(-n);
        });
      }

      case 'MID': {
        if (args.length < 3) return '';
        const start = Math.max(1, Math.floor(Number(formulaValToNumber(args[1])) || 1));
        const len = Math.max(0, Math.floor(Number(formulaValToNumber(args[2])) || 0));
        return broadcastUnaryOp(args[0], (v) => formulaValToString(v).slice(start - 1, start - 1 + len));
      }

      case 'SEARCH':
      case 'FIND': {
        if (args.length < 2) return makeFormulaError('#VALUE!');
        const caseSensitive = fnNameUpper === 'FIND';
        const startAt = args.length >= 3 ? Math.max(1, Math.floor(Number(formulaValToNumber(args[2])) || 1)) - 1 : 0;
        return broadcastBinaryOp(args[0], args[1], (needle, haystack) => {
          const nStr = formulaValToString(needle);
          const hStr = formulaValToString(haystack);
          const idx = caseSensitive
            ? hStr.indexOf(nStr, startAt)
            : hStr.toLowerCase().indexOf(nStr.toLowerCase(), startAt);
          return idx === -1 ? makeFormulaError('#VALUE!') : idx + 1;
        });
      }

      case 'COUNTA': {
        let cnt = 0;
        args.forEach(a => {
          flattenMatrixValues(a).forEach(v => {
            if (!isEmptyFormulaVal(v)) cnt++;
          });
        });
        return cnt;
      }

      case 'COUNTBLANK': {
        let cnt = 0;
        args.forEach(a => {
          flattenMatrixValues(a).forEach(v => {
            if (isEmptyFormulaVal(v)) cnt++;
          });
        });
        return cnt;
      }

      case 'COUNTIF': {
        if (args.length < 2) return 0;
        const vals = flattenMatrixValues(args[0]);
        const crit = args[1];
        const critStr = formulaValToString(crit);
        const opMatch = critStr.match(/^(<>|!=|<=|>=|=|<|>)(.*)$/);
        let cnt = 0;
        vals.forEach(v => {
          if (opMatch) {
            const op = opMatch[1];
            const rhsRaw = opMatch[2];
            const rhs = rhsRaw !== '' && /^-?(?:0|[1-9]\d*)(?:\.\d+)?$/.test(rhsRaw.trim())
              ? Number(rhsRaw.trim())
              : rhsRaw;
            if (compareFormulaScalars(v, rhs, op) === true) cnt++;
          } else if (typeof crit === 'string' && (crit.includes('*') || crit.includes('?'))) {
            const escaped = crit.replace(/[.+^${}()|[\]\\]/g, '\\$&').replace(/\*/g, '.*').replace(/\?/g, '.');
            if (new RegExp(`^${escaped}$`, 'i').test(formulaValToString(v))) cnt++;
          } else {
            const sV = formulaValToString(v).toLowerCase();
            const sC = critStr.toLowerCase();
            if (sV === sC) cnt++;
          }
        });
        return cnt;
      }

      case 'SUM': {
        let total = 0;
        args.forEach(a => {
          flattenMatrixValues(a).forEach(v => {
            if (typeof v === 'number' && Number.isFinite(v)) total += v;
          });
        });
        return total;
      }

      case 'AVERAGE': {
        let total = 0;
        let cnt = 0;
        args.forEach(a => {
          flattenMatrixValues(a).forEach(v => {
            if (typeof v === 'number' && Number.isFinite(v)) {
              total += v;
              cnt++;
            }
          });
        });
        return cnt === 0 ? makeFormulaError('#DIV/0!') : total / cnt;
      }

      case 'MIN':
      case 'MAX': {
        const nums = [];
        args.forEach(a => {
          flattenMatrixValues(a).forEach(v => {
            if (typeof v === 'number' && Number.isFinite(v)) nums.push(v);
          });
        });
        if (nums.length === 0) return 0;
        return fnNameUpper === 'MIN' ? Math.min(...nums) : Math.max(...nums);
      }

      case 'TODAY': {
        const now = new Date();
        return parseFlexibleScratchDate(new Date(Date.UTC(now.getFullYear(), now.getMonth(), now.getDate())));
      }

      default:
        return makeFormulaError('#NAME?');
    }
  }

  function parsePrimary() {
    const tok = peek();
    if (!tok) return null;

    if (tok.type === '(') {
      consume();
      const inner = parseComparison();
      if (peek() && peek().type === ')') consume();
      return inner;
    }

    if (tok.type === 'NUMBER') {
      consume();
      return tok.value;
    }

    if (tok.type === 'STRING') {
      consume();
      return tok.value;
    }

    if (tok.type === 'SHEET_PREFIX') {
      const targetSheetName = consume().value;
      const targetWs = ctx.wb ? ctx.wb.getWorksheet(targetSheetName) : null;
      const startTok = peek();
      if (!startTok || startTok.type !== 'IDENT') return makeFormulaError('#REF!');
      consume();
      const startCell = parseCellRefParts(startTok.value);
      if (!startCell) return makeFormulaError('#REF!');
      if (peek() && peek().type === ':') {
        consume();
        const endTok = peek();
        if (!endTok || endTok.type !== 'IDENT') return makeFormulaError('#REF!');
        consume();
        const endCell = parseCellRefParts(endTok.value);
        if (!endCell) return makeFormulaError('#REF!');
        if (!targetWs) return makeFormulaError('#REF!');
        return readRangeMatrix(targetWs, startCell, endCell, ctx);
      }
      if (!targetWs) return makeFormulaError('#REF!');
      return readCellValueForFormula(targetWs, startCell.row, startCell.col, ctx);
    }

    if (tok.type === 'IDENT') {
      consume();
      const ident = tok.value;
      const upper = ident.toUpperCase();

      // Function call
      if (peek() && peek().type === '(') {
        consume();
        const args = parseArgExpressions();
        return evalBuiltInFunction(upper, args);
      }

      if (upper === 'TRUE') return true;
      if (upper === 'FALSE') return false;

      // Range ref A2:C200 or single cell ref D2
      const startCell = parseCellRefParts(ident);
      if (startCell) {
        if (peek() && peek().type === ':') {
          consume();
          const endTok = peek();
          if (!endTok || endTok.type !== 'IDENT') return makeFormulaError('#REF!');
          consume();
          const endCell = parseCellRefParts(endTok.value);
          if (!endCell) return makeFormulaError('#REF!');
          return readRangeMatrix(ctx.ws, startCell, endCell, ctx);
        }
        return readCellValueForFormula(ctx.ws, startCell.row, startCell.col, ctx);
      }

      return makeFormulaError('#NAME?');
    }

    consume();
    return makeFormulaError('#VALUE!');
  }

  return parseComparison();
}

function readRangeMatrix(ws, startCell, endCell, ctx) {
  const minR = Math.min(startCell.row, endCell.row);
  const maxR = Math.max(startCell.row, endCell.row);
  const minC = Math.min(startCell.col, endCell.col);
  const maxC = Math.max(startCell.col, endCell.col);
  // Cap absurd full-sheet scans cleanly to avoid freezing if someone writes A1:A1048576
  const safeMaxR = Math.min(maxR, minR + 5000);
  const safeMaxC = Math.min(maxC, minC + 100);
  const rows = [];
  for (let r = minR; r <= safeMaxR; r++) {
    const rowArr = [];
    for (let c = minC; c <= safeMaxC; c++) {
      rowArr.push(readCellValueForFormula(ws, r, c, ctx));
    }
    rows.push(rowArr);
  }
  return makeFormulaMatrix(rows);
}

function toExcelScalarValue(val) {
  if (isFormulaError(val)) return { error: val.error };
  if (val && typeof val === 'object' && val.__isDate) return val.dateUtc;
  if (val === undefined) return '';
  return val;
}

function evaluateWorksheetCellFormula(ws, rNum, cNum, ctx) {
  const sheetName = ws.name || '';
  const coordKey = `${sheetName}::${rNum}::${cNum}`;
  if (ctx.computedAnchors.has(coordKey)) {
    return ctx.computedAnchors.get(coordKey);
  }

  const row = ws.findRow ? ws.findRow(rNum) : null;
  const cell = row && row.findCell ? row.findCell(cNum) : null;
  if (!cell || !cell.value || typeof cell.value !== 'object' || !cell.value.formula) {
    return null;
  }

  ctx.visited.add(coordKey);
  let evaluated;
  try {
    const humanFormula = toHumanReadableFormula(cell.value.formula);
    const tokens = tokenizeScratchFormula(humanFormula);
    evaluated = evalFormulaTokens(tokens, ctx);
  } catch (err) {
    console.warn(`Scratch formula error at ${getCellCoordString(rNum, cNum)}:`, err);
    evaluated = makeFormulaError('#VALUE!');
  } finally {
    ctx.visited.delete(coordKey);
  }

  ctx.computedAnchors.set(coordKey, evaluated);

  if (isFormulaMatrix(evaluated)) {
    const outRows = evaluated.rows;
    const R = outRows.length;
    const C = outRows[0] ? outRows[0].length : 1;

    // Check for blocking user values inside the spill rectangle
    let blocked = false;
    const origSpillSet = ctx.origSpillBySheet.get(sheetName);
    for (let dr = 0; dr < R && !blocked; dr++) {
      for (let dc = 0; dc < C; dc++) {
        if (dr === 0 && dc === 0) continue;
        const tr = rNum + dr;
        const tc = cNum + dc;
        if (origSpillSet && origSpillSet.has(`${tr}::${tc}`)) continue;
        const targetInfo = extractScratchCellInfo(ws, tr, tc, { skipSpillLookup: true });
        if (targetInfo.rawEdit.trim() !== '' || targetInfo.isFormula) {
          blocked = true;
          break;
        }
      }
    }

    if (blocked) {
      cell.value = { formula: cell.value.formula, result: { error: '#SPILL!' } };
      return makeFormulaError('#SPILL!');
    }

    let sheetSpills = ctx.spillBySheet.get(sheetName);
    if (!sheetSpills) {
      sheetSpills = new Map();
      ctx.spillBySheet.set(sheetName, sheetSpills);
    }

    const topLeftRaw = outRows[0] ? outRows[0][0] : '';
    cell.value = { formula: cell.value.formula, result: toExcelScalarValue(topLeftRaw) };
    const spillCoords = [];
    for (let dr = 0; dr < R; dr++) {
      for (let dc = 0; dc < C; dc++) {
        const tr = rNum + dr;
        const tc = cNum + dc;
        const rawVal = outRows[dr][dc];
        spillCoords.push({ r: tr, c: tc });
        sheetSpills.set(`${tr}::${tc}`, {
          rawVal: (rawVal && typeof rawVal === 'object' && rawVal.__isDate) ? rawVal.dateUtc : (isFormulaError(rawVal) ? rawVal.error : rawVal),
          anchorR: rNum,
          anchorC: cNum,
          formula: cell.value.formula
        });
      }
    }
    let anchorSpillMap = ctx.anchorSpillsBySheet.get(sheetName);
    if (!anchorSpillMap) {
      anchorSpillMap = new Map();
      ctx.anchorSpillsBySheet.set(sheetName, anchorSpillMap);
    }
    anchorSpillMap.set(`${rNum}::${cNum}`, spillCoords);

    return topLeftRaw;
  }

  const hasPriorResult = cell.value.result !== undefined && cell.value.result !== null;
  if (!(isFormulaError(evaluated) && evaluated.error === '#NAME?' && hasPriorResult)) {
    cell.value = { formula: cell.value.formula, result: toExcelScalarValue(evaluated) };
  }
  return evaluated;
}

function getOriginalExcelSpillSetForSheet(sheetName) {
  const set = new Set();
  if (typeof originalArchiveState === 'undefined' || !originalArchiveState || !originalArchiveState.arrayFormulasBySheet) {
    return set;
  }
  const cellMap = originalArchiveState.arrayFormulasBySheet.get(String(sheetName || '').toLowerCase());
  if (!cellMap) return set;

  for (const [anchorRef, info] of cellMap.entries()) {
    if (!info || !info.fTag) continue;
    const refMatch = info.fTag.match(/\bref="([A-Z]+\d+)(?::([A-Z]+\d+))?"/i);
    if (!refMatch || !refMatch[2]) continue;
    const p1 = parseCellRefParts(refMatch[1]);
    const p2 = parseCellRefParts(refMatch[2]);
    const anchorP = parseCellRefParts(anchorRef);
    if (!p1 || !p2 || !anchorP) continue;
    const minR = Math.min(p1.row, p2.row);
    const maxR = Math.max(p1.row, p2.row);
    const minC = Math.min(p1.col, p2.col);
    const maxC = Math.max(p1.col, p2.col);
    for (let r = minR; r <= maxR; r++) {
      for (let c = minC; c <= maxC; c++) {
        if (r === anchorP.row && c === anchorP.col) continue;
        set.add(`${r}::${c}`);
      }
    }
  }
  return set;
}

function recalculateScratchWorksheetFormulas(ws, sheetName) {
  if (!ws || !sheetName) return [];
  const prevSpills = scratchSpilledCellsMap.get(sheetName) || new Map();
  const prevSnapshot = new Map();
  for (const [k, v] of prevSpills.entries()) {
    prevSnapshot.set(k, formatScratchScalarValue(v.rawVal));
  }

  // Collect formula cells on this worksheet
  const formulaCoords = [];
  const prevAnchorResults = new Map();
  ws.eachRow({ includeEmpty: false }, (row, rowNumber) => {
    row.eachCell({ includeEmpty: false }, (cell, colNumber) => {
      if (cell && cell.value && typeof cell.value === 'object' && cell.value.formula !== undefined) {
        formulaCoords.push({ r: rowNumber, c: colNumber });
        prevAnchorResults.set(`${rowNumber}::${colNumber}`, extractScratchCellInfo(ws, rowNumber, colNumber, { skipSpillLookup: true }).display);
      }
    });
  });

  const origSpillSet = getOriginalExcelSpillSetForSheet(sheetName);
  if (formulaCoords.length === 0 && prevSpills.size === 0 && origSpillSet.size === 0) {
    return [];
  }

  const ctx = {
    wb: workbook || { getWorksheet: (name) => (name === sheetName ? ws : null) },
    ws,
    visited: new Set(),
    computedAnchors: new Map(),
    spillBySheet: new Map(),
    anchorSpillsBySheet: new Map(),
    origSpillBySheet: new Map([[sheetName, origSpillSet]])
  };

  formulaCoords.forEach(({ r, c }) => {
    evaluateWorksheetCellFormula(ws, r, c, ctx);
  });

  // Physically purge stale Excel cached spill cells on any original array anchor that is now evaluated
  // in-app (or was cleared) so ExcelJS never writes ghost cached cells back into the saved .xlsx archive.
  if (typeof originalArchiveState !== 'undefined' && originalArchiveState && originalArchiveState.arrayFormulasBySheet) {
    const origCellMap = originalArchiveState.arrayFormulasBySheet.get(String(sheetName || '').toLowerCase());
    if (origCellMap) {
      for (const [anchorRef, info] of origCellMap.entries()) {
        if (!info || !info.fTag) continue;
        const refMatch = info.fTag.match(/\bref="([A-Z]+\d+):([A-Z]+\d+)"/i);
        if (!refMatch) continue;
        const p1 = parseCellRefParts(refMatch[1]);
        const p2 = parseCellRefParts(refMatch[2]);
        const anchorP = parseCellRefParts(anchorRef);
        if (!p1 || !p2 || !anchorP) continue;

        const anchorEval = ctx.computedAnchors.get(`${sheetName}::${anchorP.row}::${anchorP.col}`);
        const stillUnsupportedDesktopFormula = Boolean(isFormulaError(anchorEval) && anchorEval.error === '#NAME?');
        if (stillUnsupportedDesktopFormula) continue;

        const minR = Math.min(p1.row, p2.row);
        const maxR = Math.max(p1.row, p2.row);
        const minC = Math.min(p1.col, p2.col);
        const maxC = Math.max(p1.col, p2.col);
        for (let r = minR; r <= maxR; r++) {
          const rObj = ws.findRow ? ws.findRow(r) : null;
          if (!rObj) continue;
          let rowTouched = false;
          for (let c = minC; c <= maxC; c++) {
            if (r === anchorP.row && c === anchorP.col) continue;
            const cObj = rObj.findCell ? rObj.findCell(c) : null;
            if (cObj && cObj.value !== null && cObj.value !== undefined) {
              if (!(typeof cObj.value === 'object' && cObj.value.formula !== undefined)) {
                cObj.value = null;
                rowTouched = true;
              }
            }
          }
          if (rowTouched && rObj.commit) rObj.commit();
        }
        info.fTag = info.fTag.replace(/\bref="[A-Z]+\d+:[A-Z]+\d+"/i, `ref="${anchorRef}"`);
      }
    }
  }

  const nextSpills = ctx.spillBySheet.get(sheetName) || new Map();
  // Mask any original Desktop Excel cached spill cells that were NOT repopulated in this recalculation
  origSpillSet.forEach(coordKey => {
    if (!nextSpills.has(coordKey)) {
      const [rStr, cStr] = coordKey.split('::');
      nextSpills.set(coordKey, {
        rawVal: '',
        anchorR: 0,
        anchorC: 0,
        formula: '',
        clearedOrigSpill: true,
        r: parseInt(rStr, 10),
        c: parseInt(cStr, 10)
      });
    }
  });

  scratchSpilledCellsMap.set(sheetName, nextSpills);
  scratchSpillStateMap.set(sheetName, ctx.anchorSpillsBySheet.get(sheetName) || new Map());

  // Compute all coordinates whose visible display value changed
  const changedSet = new Map();
  const addChanged = (r, c) => {
    changedSet.set(`${r}::${c}`, { r, c });
  };

  formulaCoords.forEach(({ r, c }) => {
    const key = `${r}::${c}`;
    const nextDisp = extractScratchCellInfo(ws, r, c).display;
    if (prevAnchorResults.get(key) !== nextDisp) {
      addChanged(r, c);
    }
  });

  for (const [k, oldDisp] of prevSnapshot.entries()) {
    const nextEntry = nextSpills.get(k);
    const nextDisp = nextEntry ? formatScratchScalarValue(nextEntry.rawVal) : '';
    if (oldDisp !== nextDisp) {
      const [rStr, cStr] = k.split('::');
      addChanged(parseInt(rStr, 10), parseInt(cStr, 10));
    }
  }
  for (const [k, nextEntry] of nextSpills.entries()) {
    if (!prevSnapshot.has(k)) {
      const [rStr, cStr] = k.split('::');
      addChanged(parseInt(rStr, 10), parseInt(cStr, 10));
    }
  }

  return Array.from(changedSet.values());
}

// --- FILL HANDLE AUTO-INCREMENT & SERIES EXTRAPOLATION ---
function computeScratchFillTargetRange(srcRange, hoverR, hoverC) {
  if (!srcRange) return null;
  const dDown = hoverR > srcRange.maxRow ? hoverR - srcRange.maxRow : 0;
  const dUp = hoverR < srcRange.minRow ? srcRange.minRow - hoverR : 0;
  const dRight = hoverC > srcRange.maxCol ? hoverC - srcRange.maxCol : 0;
  const dLeft = hoverC < srcRange.minCol ? srcRange.minCol - hoverC : 0;
  if (dDown === 0 && dUp === 0 && dRight === 0 && dLeft === 0) return null;

  if (Math.max(dDown, dUp) >= Math.max(dRight, dLeft)) {
    if (dDown > 0) {
      return { axis: 'vertical', dir: 1, minRow: srcRange.minRow, maxRow: hoverR, minCol: srcRange.minCol, maxCol: srcRange.maxCol };
    }
    return { axis: 'vertical', dir: -1, minRow: hoverR, maxRow: srcRange.maxRow, minCol: srcRange.minCol, maxCol: srcRange.maxCol };
  }
  if (dRight > 0) {
    return { axis: 'horizontal', dir: 1, minRow: srcRange.minRow, maxRow: srcRange.maxRow, minCol: srcRange.minCol, maxCol: hoverC };
  }
  return { axis: 'horizontal', dir: -1, minRow: srcRange.minRow, maxRow: srcRange.maxRow, minCol: hoverC, maxCol: srcRange.maxCol };
}

function buildFillSeriesValues(
  sourceItems,
  targetCount,
  dir,
  { axis = 'vertical', ctrlKey = false, autoIncrement = scratchAutoIncrementEnabled } = {}
) {
  const S = sourceItems.length;
  if (S === 0 || targetCount <= 0) return [];

  const results = [];
  const allNonEmpty = sourceItems.every(it => it.rawEdit.trim() !== '');
  const anyFormula = sourceItems.some(it => it.isFormula && it.rawEdit.startsWith('='));
  const anyExplicitText = sourceItems.some(it => it.formatId === 'text');
  const shouldAutoIncrement = Boolean(autoIncrement) !== Boolean(ctrlKey);

  // 1. Formulas: shift relative references by step offset
  if (anyFormula) {
    for (let m = 1; m <= targetCount; m++) {
      const idx = dir > 0 ? ((m - 1) % S) : (S - 1 - ((m - 1) % S));
      const src = sourceItems[idx];
      const deltaUnits = dir > 0
        ? (S - idx + (m - 1))
        : -(idx + m);
      const rDelta = axis === 'vertical' ? deltaUnits : 0;
      const cDelta = axis === 'horizontal' ? deltaUnits : 0;
      const nextText = src.isFormula && src.rawEdit.startsWith('=')
        ? shiftFormulaCellReferences(src.rawEdit, rDelta, cDelta)
        : src.rawEdit;
      results.push({ rawText: nextText, styleSnap: src.styleSnap });
    }
    return results;
  }

  // 2. Dates & Numbers (when Auto-increment enabled xor Ctrl held, and not explicit Text format)
  if (allNonEmpty && !anyExplicitText && shouldAutoIncrement) {
    const parsedDates = sourceItems.map(it => {
      const allowHyphen = it.formatId === 'date';
      const p = parseFlexibleScratchDate(it.rawEdit || it.display, { allowTwoDigitHyphen: allowHyphen });
      if (!p) return null;
      // Only treat as date when cell format is 'date' OR string actually contains a date separator
      if (it.formatId !== 'date' && !/[\/\-]/.test(it.rawEdit)) return null;
      return p;
    });
    if (parsedDates.every(Boolean)) {
      let stepDays = 1;
      if (S >= 2) {
        stepDays = Math.round((parsedDates[S - 1].dayEpoch - parsedDates[0].dayEpoch) / (S - 1));
      }
      for (let m = 1; m <= targetCount; m++) {
        const srcIdx = dir > 0 ? ((m - 1) % S) : (S - 1 - ((m - 1) % S));
        const baseSnap = sourceItems[srcIdx].styleSnap || { fill: null, font: null, numFmt: null };
        const epoch = dir > 0
          ? parsedDates[S - 1].dayEpoch + m * stepDays
          : parsedDates[0].dayEpoch - m * stepDays;
        const dt = new Date(epoch * 86400000);
        results.push({
          rawText: formatExcelDateDisplay(dt),
          styleSnap: {
            fill: baseSnap.fill,
            font: baseSnap.font,
            numFmt: baseSnap.numFmt || 'mm/dd/yyyy'
          }
        });
      }
      return results;
    }

    // 3a. Long exact integers (11-30 digits, e.g. sequential tracking numbers) using BigInt
    //     while applyRawEditToExcelCell continues storing 11+ digit outputs as text strings.
    const parsedBigInts = sourceItems.map(it => {
      const trimmed = it.rawEdit.trim();
      if (!/^-?(?:0|[1-9]\d*)$/.test(trimmed)) return null;
      const digitLen = trimmed.replace(/^-/, '').length;
      if (digitLen > 30) return null;
      return { val: BigInt(trimmed), digitLen };
    });
    if (parsedBigInts.every(Boolean) && parsedBigInts.some(x => x.digitLen > 10)) {
      let stepBig = 1n;
      if (S >= 2) {
        const diff = parsedBigInts[S - 1].val - parsedBigInts[0].val;
        const span = BigInt(S - 1);
        if (diff % span !== 0n) {
          stepBig = null;
        } else {
          stepBig = diff / span;
        }
      }
      if (stepBig !== null) {
        for (let m = 1; m <= targetCount; m++) {
          const srcIdx = dir > 0 ? ((m - 1) % S) : (S - 1 - ((m - 1) % S));
          const baseSnap = sourceItems[srcIdx].styleSnap || { fill: null, font: null, numFmt: null };
          const nextVal = dir > 0
            ? parsedBigInts[S - 1].val + BigInt(m) * stepBig
            : parsedBigInts[0].val - BigInt(m) * stepBig;
          results.push({
            rawText: nextVal.toString(),
            styleSnap: baseSnap
          });
        }
        return results;
      }
    }

    // 3b. Standard Numbers / Currency (<= 15 significant digits)
    const parsedNums = sourceItems.map(it => {
      const cleaned = it.rawEdit.trim().replace(/^\$/, '').replace(/,/g, '');
      const valid = /^-?(?:0|[1-9]\d*)(?:\.\d+)?$/.test(cleaned) && cleaned.replace(/[-.]/g, '').length <= 15;
      if (!valid) return null;
      const decPart = cleaned.includes('.') ? cleaned.split('.')[1].length : 0;
      return { num: Number(cleaned), decimals: decPart };
    });
    if (parsedNums.every(Boolean)) {
      let stepNum = 1;
      if (S >= 2) {
        stepNum = (parsedNums[S - 1].num - parsedNums[0].num) / (S - 1);
      }
      const maxDec = parsedNums.reduce((mx, x) => Math.max(mx, x.decimals), 0);
      const factor = Math.pow(10, Math.min(6, maxDec));
      for (let m = 1; m <= targetCount; m++) {
        const srcIdx = dir > 0 ? ((m - 1) % S) : (S - 1 - ((m - 1) % S));
        const baseSnap = sourceItems[srcIdx].styleSnap || { fill: null, font: null, numFmt: null };
        const rawVal = dir > 0
          ? parsedNums[S - 1].num + m * stepNum
          : parsedNums[0].num - m * stepNum;
        const rounded = Math.round(rawVal * factor) / factor;
        results.push({
          rawText: String(rounded),
          styleSnap: baseSnap
        });
      }
      return results;
    }
  }

  // 4. Plain Text / Explicit Text (@) / Non-increment copy fallback: cycle values without incrementing
  for (let m = 1; m <= targetCount; m++) {
    const srcIdx = dir > 0 ? ((m - 1) % S) : (S - 1 - ((m - 1) % S));
    const src = sourceItems[srcIdx];
    results.push({
      rawText: src.rawEdit,
      styleSnap: src.styleSnap
    });
  }
  return results;
}

function executeScratchFillHandleDrag(sheetName, sourceRange, fillTargetRange, { ctrlKey = false } = {}) {
  if (!workbook || !sheetName || !sourceRange || !fillTargetRange) return 0;
  const ws = workbook.getWorksheet(sheetName);
  if (!ws) return 0;

  const edits = [];
  if (fillTargetRange.axis === 'vertical') {
    const dir = fillTargetRange.dir;
    const targetCount = dir > 0
      ? (fillTargetRange.maxRow - sourceRange.maxRow)
      : (sourceRange.minRow - fillTargetRange.minRow);
    if (targetCount <= 0) return 0;

    for (let c = sourceRange.minCol; c <= sourceRange.maxCol; c++) {
      const sourceItems = [];
      for (let r = sourceRange.minRow; r <= sourceRange.maxRow; r++) {
        const info = extractScratchCellInfo(ws, r, c);
        sourceItems.push({
          rawEdit: info.rawEdit,
          display: info.display,
          isFormula: info.isFormula,
          formatId: getScratchCellFormatId(ws, r, c),
          styleSnap: captureCellStyleSnapshot(ws, r, c)
        });
      }
      const generated = buildFillSeriesValues(sourceItems, targetCount, dir, {
        axis: 'vertical',
        ctrlKey,
        autoIncrement: scratchAutoIncrementEnabled
      });
      for (let m = 0; m < generated.length; m++) {
        const targetRow = dir > 0 ? (sourceRange.maxRow + 1 + m) : (sourceRange.minRow - 1 - m);
        edits.push({ r: targetRow, c, rawText: generated[m].rawText, styleSnap: generated[m].styleSnap });
      }
    }
  } else {
    const dir = fillTargetRange.dir;
    const targetCount = dir > 0
      ? (fillTargetRange.maxCol - sourceRange.maxCol)
      : (sourceRange.minCol - fillTargetRange.minCol);
    if (targetCount <= 0) return 0;

    for (let r = sourceRange.minRow; r <= sourceRange.maxRow; r++) {
      const sourceItems = [];
      for (let c = sourceRange.minCol; c <= sourceRange.maxCol; c++) {
        const info = extractScratchCellInfo(ws, r, c);
        sourceItems.push({
          rawEdit: info.rawEdit,
          display: info.display,
          isFormula: info.isFormula,
          formatId: getScratchCellFormatId(ws, r, c),
          styleSnap: captureCellStyleSnapshot(ws, r, c)
        });
      }
      const generated = buildFillSeriesValues(sourceItems, targetCount, dir, {
        axis: 'horizontal',
        ctrlKey,
        autoIncrement: scratchAutoIncrementEnabled
      });
      for (let m = 0; m < generated.length; m++) {
        const targetCol = dir > 0 ? (sourceRange.maxCol + 1 + m) : (sourceRange.minCol - 1 - m);
        edits.push({ r, c: targetCol, rawText: generated[m].rawText, styleSnap: generated[m].styleSnap });
      }
    }
  }

  if (edits.length === 0) return 0;
  scratchClipboardRangeState = null;
  const nextSelection = {
    anchorRow: fillTargetRange.minRow,
    anchorCol: fillTargetRange.minCol,
    focusRow: fillTargetRange.maxRow,
    focusCol: fillTargetRange.maxCol
  };
  return applyScratchCellChanges(sheetName, edits, {
    recordUndo: true,
    selectionAfter: nextSelection
  });
}

function applyScratchSelectionDOM() {
  if (activeWorkspaceTab === DAILY_TRACKER_WS_TAB || !workbook || !scratchGridBody) return;
  const ws = workbook.getWorksheet(activeWorkspaceTab);
  if (!ws) return;

  const dims = getScratchSheetGridDimensions(ws, activeWorkspaceTab);
  const range = getScratchSelectionRange(activeWorkspaceTab);
  const clip =
    scratchClipboardRangeState && scratchClipboardRangeState.sheetName === activeWorkspaceTab
      ? scratchClipboardRangeState
      : null;
  const fillPreview =
    scratchDragState && scratchDragState.mode === 'fill' && scratchDragState.sheetName === activeWorkspaceTab
      ? scratchDragState.fillTargetRange
      : null;

  const prevClasses = [
    'selected-cell',
    'in-range-cell',
    'fill-handle-host',
    'range-edge-top',
    'range-edge-bottom',
    'range-edge-left',
    'range-edge-right',
    'copied-edge-top',
    'copied-edge-bottom',
    'copied-edge-left',
    'copied-edge-right',
    'in-fill-preview-cell',
    'fill-edge-top',
    'fill-edge-bottom',
    'fill-edge-left',
    'fill-edge-right'
  ];
  const selector = prevClasses.map(cls => `.${cls}`).join(',');
  scratchGridBody.querySelectorAll(selector).forEach(el => {
    el.classList.remove(...prevClasses);
  });
  scratchGridBody.querySelectorAll('.scratch-fill-handle').forEach(el => el.remove());
  scratchGridBody.querySelectorAll('.active-row-hdr').forEach(el => el.classList.remove('active-row-hdr'));

  if (scratchGridHead) {
    scratchGridHead.querySelectorAll('.active-col-hdr').forEach(el => el.classList.remove('active-col-hdr'));
    const cornerTh = scratchGridHead.querySelector('.scratch-corner-th');
    if (cornerTh) {
      const isWholeSheet =
        range.minRow === 1 &&
        range.maxRow >= dims.rows &&
        range.minCol === 1 &&
        range.maxCol >= dims.cols;
      cornerTh.classList.toggle('active-corner-hdr', isWholeSheet);
    }
    const hdrRow = scratchGridHead.rows[0];
    if (hdrRow) {
      const maxHdrCol = Math.min(range.maxCol, hdrRow.cells.length - 1);
      for (let c = range.minCol; c <= maxHdrCol; c++) {
        if (hdrRow.cells[c]) hdrRow.cells[c].classList.add('active-col-hdr');
      }
    }
  }

  const trs = scratchGridBody.rows;
  for (let i = 0; i < trs.length; i++) {
    const tr = trs[i];
    const rowHdr = tr.cells[0];
    if (!rowHdr) continue;
    const r = Number(rowHdr.dataset.rowHdr);
    if (!r) continue;

    if (r >= range.minRow && r <= range.maxRow) {
      rowHdr.classList.add('active-row-hdr');
      const maxCellIdx = Math.min(range.maxCol, tr.cells.length - 1);
      if (range.cellCount === 1) {
        const td = tr.cells[range.minCol];
        if (td) td.classList.add('selected-cell');
      } else {
        for (let c = range.minCol; c <= maxCellIdx; c++) {
          const td = tr.cells[c];
          if (!td) continue;
          td.classList.add('in-range-cell');
          if (r === range.anchorRow && c === range.anchorCol) {
            td.classList.add('selected-cell');
          }
          if (r === range.minRow) td.classList.add('range-edge-top');
          if (r === range.maxRow) td.classList.add('range-edge-bottom');
          if (c === range.minCol) td.classList.add('range-edge-left');
          if (c === range.maxCol) td.classList.add('range-edge-right');
        }
      }
    }

    if (clip && r >= clip.minRow && r <= clip.maxRow) {
      const maxClipCol = Math.min(clip.maxCol, tr.cells.length - 1);
      for (let c = clip.minCol; c <= maxClipCol; c++) {
        const td = tr.cells[c];
        if (!td) continue;
        if (r === clip.minRow) td.classList.add('copied-edge-top');
        if (r === clip.maxRow) td.classList.add('copied-edge-bottom');
        if (c === clip.minCol) td.classList.add('copied-edge-left');
        if (c === clip.maxCol) td.classList.add('copied-edge-right');
      }
    }

    if (fillPreview && r >= fillPreview.minRow && r <= fillPreview.maxRow) {
      const maxFillCol = Math.min(fillPreview.maxCol, tr.cells.length - 1);
      for (let c = fillPreview.minCol; c <= maxFillCol; c++) {
        const td = tr.cells[c];
        if (!td) continue;
        const inSrc = r >= range.minRow && r <= range.maxRow && c >= range.minCol && c <= range.maxCol;
        if (!inSrc) {
          td.classList.add('in-fill-preview-cell');
          if (r === fillPreview.minRow) td.classList.add('fill-edge-top');
          if (r === fillPreview.maxRow) td.classList.add('fill-edge-bottom');
          if (c === fillPreview.minCol) td.classList.add('fill-edge-left');
          if (c === fillPreview.maxCol) td.classList.add('fill-edge-right');
        }
      }
    }
  }

  // Attach Excel-style bottom-right corner Fill Handle square when not inline-editing
  if (!scratchInlineEditState && !fillPreview) {
    const cornerTd = scratchGridBody.querySelector(`td[data-r="${range.maxRow}"][data-c="${range.maxCol}"]`);
    if (cornerTd) {
      cornerTd.classList.add('fill-handle-host');
      const handleEl = document.createElement('span');
      handleEl.className = 'scratch-fill-handle';
      handleEl.dataset.fillHandle = 'true';
      handleEl.title = 'Drag corner square to copy/fill • Dates & Numbers auto-increment (hold Ctrl to copy without incrementing)';
      cornerTd.appendChild(handleEl);
    }
  }

  updateScratchSelectionStatsUI();
  syncScratchFormatSelectUI();
}

function setScratchSelectionRange(
  sheetName,
  rawRange,
  { scrollIntoView = false, scrollTarget = 'focus' } = {}
) {
  if (!workbook || !sheetName || sheetName === DAILY_TRACKER_WS_TAB) return;
  const ws = workbook.getWorksheet(sheetName);
  if (!ws) return;

  const norm = normalizeScratchRangeObj(rawRange);
  const dims = getScratchSheetGridDimensions(ws, sheetName);

  let needsReRender = false;
  let targetRows = dims.rows;
  let targetCols = dims.cols;
  if (norm.maxRow > dims.rows) {
    targetRows = norm.maxRow + 9;
    needsReRender = true;
  }
  if (norm.maxCol > dims.cols) {
    targetCols = norm.maxCol + 2;
    needsReRender = true;
  }
  if (needsReRender) {
    scratchGridDimensionsMap.set(sheetName, { rows: targetRows, cols: targetCols });
  }

  scratchSelectionRangeMap.set(sheetName, {
    anchorRow: norm.anchorRow,
    anchorCol: norm.anchorCol,
    focusRow: norm.focusRow,
    focusCol: norm.focusCol
  });
  scratchActiveCellMap.set(sheetName, { row: norm.anchorRow, col: norm.anchorCol });

  if (activeWorkspaceTab === sheetName) {
    if (needsReRender) {
      renderActiveScratchSheetGrid();
    } else {
      applyScratchSelectionDOM();
      syncScratchFormulaBarFromActiveCell();
    }

    if (scrollIntoView && scratchGridBody) {
      const targetR = scrollTarget === 'anchor' ? norm.anchorRow : norm.focusRow;
      const targetC = scrollTarget === 'anchor' ? norm.anchorCol : norm.focusCol;
      const td = scratchGridBody.querySelector(`td[data-r="${targetR}"][data-c="${targetC}"]`);
      if (td) {
        td.scrollIntoView({ block: 'nearest', inline: 'nearest' });
      }
    }
  }
}

function selectScratchCell(rNum, cNum, { scrollIntoView = false, extendSelection = false } = {}) {
  if (activeWorkspaceTab === DAILY_TRACKER_WS_TAB || !workbook) return;
  const clampedRow = Math.max(1, rNum);
  const clampedCol = Math.max(1, cNum);

  if (extendSelection) {
    const cur = getScratchSelectionRange(activeWorkspaceTab);
    setScratchSelectionRange(
      activeWorkspaceTab,
      {
        anchorRow: cur.anchorRow,
        anchorCol: cur.anchorCol,
        focusRow: clampedRow,
        focusCol: clampedCol
      },
      { scrollIntoView, scrollTarget: 'focus' }
    );
  } else {
    setScratchSelectionRange(
      activeWorkspaceTab,
      {
        anchorRow: clampedRow,
        anchorCol: clampedCol,
        focusRow: clampedRow,
        focusCol: clampedCol
      },
      { scrollIntoView, scrollTarget: 'anchor' }
    );
  }
}

function commitActiveScratchInlineEdit({ restoreGridFocus = false } = {}) {
  if (!scratchInlineEditState) return;
  const state = scratchInlineEditState;
  const inputHadFocus = Boolean(state.inputEl && document.activeElement === state.inputEl);
  scratchInlineEditState = null;

  const newRaw = state.inputEl ? state.inputEl.value : state.originalRaw;
  if (state.inputEl && state.inputEl.parentNode) {
    state.inputEl.remove();
  }

  if (newRaw !== state.originalRaw) {
    recordScratchCellEdit(state.sheetName, state.row, state.col, newRaw);
  } else {
    refreshSingleScratchCellDOM(state.sheetName, state.row, state.col);
  }

  if (
    (restoreGridFocus || inputHadFocus) &&
    activeWorkspaceTab !== DAILY_TRACKER_WS_TAB &&
    scratchGridWrap &&
    (!document.activeElement || document.activeElement === document.body)
  ) {
    scratchGridWrap.focus({ preventScroll: true });
  }
}

function cancelActiveScratchInlineEdit() {
  if (!scratchInlineEditState) return;
  const state = scratchInlineEditState;
  scratchInlineEditState = null;

  if (state.inputEl && state.inputEl.parentNode) {
    state.inputEl.remove();
  }
  refreshSingleScratchCellDOM(state.sheetName, state.row, state.col);
  syncScratchFormulaBarFromActiveCell();
  if (
    activeWorkspaceTab !== DAILY_TRACKER_WS_TAB &&
    scratchGridWrap &&
    (!document.activeElement || document.activeElement === document.body)
  ) {
    scratchGridWrap.focus({ preventScroll: true });
  }
}

function isUserActivelyInteractingWithScratch() {
  if (scratchInlineEditState) return true;
  if (scratchDragState || scratchColResizeState) return true;
  if (scratchFormulaInput && document.activeElement === scratchFormulaInput) return true;
  return false;
}

function scheduleScratchAutoSave(delayMs = 1800) {
  if (typeof isLockedByOtherTab !== 'undefined' && isLockedByOtherTab) {
    return;
  }
  if (scratchAutoSaveTimer) {
    clearTimeout(scratchAutoSaveTimer);
  }
  scratchAutoSaveTimer = setTimeout(() => {
    scratchAutoSaveTimer = null;
    if (typeof isLockedByOtherTab !== 'undefined' && isLockedByOtherTab) {
      return;
    }
    if (isUserActivelyInteractingWithScratch()) {
      // Never steal focus or commit mid-typing; wait until the user finishes editing the cell
      scheduleScratchAutoSave(1500);
      return;
    }
    flushPendingScratchEditsToDisk({ reloadAfter: false, silentToast: true, commitInline: false });
  }, delayMs);
}

async function flushPendingScratchEditsToDisk({
  reloadAfter = false,
  silentToast = false,
  commitInline = true
} = {}) {
  if (typeof isLockedByOtherTab !== 'undefined' && isLockedByOtherTab && silentToast) {
    return false;
  }
  if (commitInline) {
    if (scratchAutoSaveTimer) {
      clearTimeout(scratchAutoSaveTimer);
      scratchAutoSaveTimer = null;
    }
    commitActiveScratchInlineEdit({ restoreGridFocus: true });
  } else if (isUserActivelyInteractingWithScratch()) {
    scheduleScratchAutoSave(1500);
    return false;
  }

  if (!workbook || !workingFileHandle) return false;
  if (!hasAnyPendingScratchMutations() && silentToast) return true;
  if (scratchIsSaving) {
    if (hasAnyPendingScratchMutations()) {
      scheduleScratchAutoSave(1500);
    }
    return false;
  }

  const versionAtSaveStart = scratchMutationVersion;
  scratchIsSaving = true;
  updateScratchSaveBadgeUI();

  const saved = await saveWorkbookToDisk();
  scratchIsSaving = false;

  // Restore grid keyboard focus if browser dropped activeElement to body during async file write
  if (
    activeWorkspaceTab !== DAILY_TRACKER_WS_TAB &&
    !scratchInlineEditState &&
    scratchGridWrap &&
    (!document.activeElement || document.activeElement === document.body)
  ) {
    scratchGridWrap.focus({ preventScroll: true });
  }

  if (saved) {
    if (scratchMutationVersion === versionAtSaveStart) {
      scratchPendingEdits.clear();
      scratchPendingFills.clear();
      scratchPendingColWidths.clear();
    } else {
      // User made additional edits while disk I/O was in flight; queue follow-up auto-save
      scheduleScratchAutoSave(1800);
    }
    updateScratchSaveBadgeUI();
    renderWorkspaceTabsBar();
    if (reloadAfter) {
      await loadWorkbookFromDisk(sheetSelect.value, {
        preserveStatus: saved === 'mirror-locked',
        skipFlushPendingScratch: true
      });
    }
    if (!silentToast) {
      showToast('Scratch sheet saved to workbook & mirror.', 'success');
    }
    return true;
  } else {
    updateScratchSaveBadgeUI();
    return false;
  }
}

function refreshSingleScratchCellDOM(sheetName, rNum, cNum) {
  if (activeWorkspaceTab !== sheetName || !scratchGridBody || !workbook) return;
  const ws = workbook.getWorksheet(sheetName);
  if (!ws) return;
  const td = scratchGridBody.querySelector(`td[data-r="${rNum}"][data-c="${cNum}"]`);
  if (!td) return;

  const info = extractScratchCellInfo(ws, rNum, cNum);
  const visual = extractScratchCellVisualStyle(ws, rNum, cNum);
  const anchorSpills = scratchSpillStateMap.get(sheetName);
  const isArrayAnchor = Boolean(anchorSpills && anchorSpills.has(`${rNum}::${cNum}`));
  td.classList.toggle('formula-cell', info.isFormula);
  td.classList.toggle('formula-anchor-cell', isArrayAnchor);
  td.classList.toggle('spilled-array-cell', Boolean(info.isSpilled && info.anchorR > 0));
  applyVisualStyleToScratchTd(td, visual);
  if (info.isFormula) {
    td.title = `${info.rawEdit} → ${info.display}`;
  } else if (info.isSpilled && info.anchorR > 0) {
    td.title = `${info.display} (Spilled from ${getCellCoordString(info.anchorR, info.anchorC)})`;
  } else {
    td.title = info.display;
  }

  let span = td.querySelector('.scratch-cell-text');
  if (!span) {
    span = document.createElement('span');
    span.className = 'scratch-cell-text';
    td.innerHTML = '';
    td.appendChild(span);
  }
  span.textContent = info.display;
}

// --- BATCH CELL MUTATION ENGINE WITH 50-STEP UNDO / REDO (VALUES & EXCEL FILL COLORS) ---
function applyScratchCellChanges(
  sheetName,
  editsList,
  {
    recordUndo = true,
    coalesceKey = null,
    selectionAfter = null
  } = {}
) {
  if (!workbook || !sheetName || !Array.isArray(editsList) || editsList.length === 0) return 0;
  const ws = workbook.getWorksheet(sheetName);
  if (!ws) return 0;

  const selectionBefore = getScratchSelectionRange(sheetName);
  const recordedChanges = [];
  let maxTargetRow = 0;
  let maxTargetCol = 0;

  if (!scratchPendingEdits.has(sheetName)) {
    scratchPendingEdits.set(sheetName, new Map());
  }
  if (!scratchPendingFills.has(sheetName)) {
    scratchPendingFills.set(sheetName, new Map());
  }
  const pendingSheetMap = scratchPendingEdits.get(sheetName);
  const pendingFillMap = scratchPendingFills.get(sheetName);

  editsList.forEach(item => {
    const r = Math.max(1, Math.floor(item.r));
    const c = Math.max(1, Math.floor(item.c));
    const hasTextProp = Object.prototype.hasOwnProperty.call(item, 'rawText');
    const hasStyleProp = Object.prototype.hasOwnProperty.call(item, 'styleSnap');

    if (r > maxTargetRow) maxTargetRow = r;
    if (c > maxTargetCol) maxTargetCol = c;

    let textChanged = false;
    let beforeRaw = '';
    let afterRaw = '';
    if (hasTextProp) {
      beforeRaw = extractScratchCellInfo(ws, r, c, { skipSpillLookup: true }).rawEdit;
      afterRaw = String(item.rawText ?? '').replace(/\r/g, '');
      textChanged = beforeRaw !== afterRaw;
    }

    let styleChanged = false;
    let beforeStyle = null;
    let afterStyle = null;
    if (hasStyleProp) {
      beforeStyle = captureCellStyleSnapshot(ws, r, c);
      afterStyle = item.styleSnap
        ? JSON.parse(JSON.stringify(item.styleSnap))
        : { fill: null, font: null, numFmt: null };
      styleChanged = JSON.stringify(beforeStyle) !== JSON.stringify(afterStyle);
    }

    if (!textChanged && !styleChanged) return;

    // Apply style/format snapshot first so applyRawEditToExcelCell respects updated numFmt
    if (styleChanged) {
      applyCellStyleSnapshotToExcelCell(ws, r, c, afterStyle);
      pendingFillMap.set(`${r}::${c}`, afterStyle);
    }
    if (textChanged) {
      applyRawEditToExcelCell(ws, r, c, afterRaw);
      pendingSheetMap.set(`${r}::${c}`, afterRaw);
    }

    recordedChanges.push({
      r,
      c,
      textChanged,
      beforeRaw,
      afterRaw,
      styleChanged,
      beforeStyle,
      afterStyle
    });
  });

  if (recordedChanges.length === 0) {
    if (selectionAfter) {
      setScratchSelectionRange(sheetName, selectionAfter);
    }
    return 0;
  }

  if (recordUndo) {
    if (!scratchUndoStackMap.has(sheetName)) {
      scratchUndoStackMap.set(sheetName, []);
    }
    const undoStack = scratchUndoStackMap.get(sheetName);
    const now = Date.now();
    const canCoalesce =
      coalesceKey &&
      recordedChanges.length === 1 &&
      recordedChanges[0].textChanged &&
      !recordedChanges[0].styleChanged &&
      undoStack.length > 0 &&
      undoStack[undoStack.length - 1].coalesceKey === coalesceKey &&
      lastFormulaCoalesceMeta &&
      lastFormulaCoalesceMeta.key === coalesceKey &&
      now - lastFormulaCoalesceMeta.at < 2500;

    if (canCoalesce) {
      const top = undoStack[undoStack.length - 1];
      top.changes[0].afterRaw = recordedChanges[0].afterRaw;
      lastFormulaCoalesceMeta.at = now;
    } else {
      undoStack.push({
        changes: recordedChanges,
        selectionBefore: {
          anchorRow: selectionBefore.anchorRow,
          anchorCol: selectionBefore.anchorCol,
          focusRow: selectionBefore.focusRow,
          focusCol: selectionBefore.focusCol
        },
        selectionAfter: selectionAfter || {
          anchorRow: selectionBefore.anchorRow,
          anchorCol: selectionBefore.anchorCol,
          focusRow: selectionBefore.focusRow,
          focusCol: selectionBefore.focusCol
        },
        coalesceKey: coalesceKey || null
      });
      if (undoStack.length > SCRATCH_MAX_UNDO_STEPS) {
        undoStack.shift();
      }
      lastFormulaCoalesceMeta = coalesceKey ? { key: coalesceKey, at: now } : null;
    }
    scratchRedoStackMap.set(sheetName, []);
  }

  // Re-evaluate any =FILTER / =SORT / =UNIQUE / summary formulas affected by these edits
  const formulaAffectedCoords = recalculateScratchWorksheetFormulas(ws, sheetName);
  formulaAffectedCoords.forEach(({ r, c }) => {
    if (r > maxTargetRow) maxTargetRow = r;
    if (c > maxTargetCol) maxTargetCol = c;
  });

  // Update cached non-empty row count & data bounds metadata (including live spills)
  const bounds = measureScratchWorksheetBounds(ws);
  if (bounds.maxRow > maxTargetRow) maxTargetRow = bounds.maxRow;
  if (bounds.maxCol > maxTargetCol) maxTargetCol = bounds.maxCol;

  // Update grid dimensions if paste/edit/spill exceeded previously rendered/cached bounds
  const prevRenderedRows = scratchGridBody ? scratchGridBody.rows.length : 0;
  const prevRenderedCols = (scratchGridHead && scratchGridHead.rows && scratchGridHead.rows[0])
    ? Math.max(0, scratchGridHead.rows[0].cells.length - 1)
    : 0;
  const prevCachedDims = scratchGridDimensionsMap.get(sheetName) || {
    rows: prevRenderedRows,
    cols: prevRenderedCols
  };
  const curDims = getScratchSheetGridDimensions(ws, sheetName);
  const targetRows = Math.max(curDims.rows, maxTargetRow + 5);
  const targetCols = Math.max(curDims.cols, maxTargetCol + 2);
  if (targetRows > curDims.rows || targetCols > curDims.cols) {
    scratchGridDimensionsMap.set(sheetName, {
      rows: targetRows,
      cols: targetCols
    });
  }
  const needsExpand =
    targetRows > prevRenderedRows ||
    targetCols > prevRenderedCols ||
    targetRows > prevCachedDims.rows ||
    targetCols > prevCachedDims.cols ||
    maxTargetRow > prevRenderedRows ||
    maxTargetCol > prevRenderedCols;
  const touchedRow1 = recordedChanges.some(ch => ch.r === 1);

  const meta = scratchSheets.find(s => s.name === sheetName);
  if (meta) {
    meta.nonEmptyRows = bounds.nonEmptyRows;
    meta.maxRow = bounds.maxRow;
    meta.maxCol = bounds.maxCol;
  }

  if (selectionAfter) {
    scratchSelectionRangeMap.set(sheetName, {
      anchorRow: selectionAfter.anchorRow,
      anchorCol: selectionAfter.anchorCol,
      focusRow: selectionAfter.focusRow,
      focusCol: selectionAfter.focusCol
    });
    scratchActiveCellMap.set(sheetName, {
      row: selectionAfter.anchorRow,
      col: selectionAfter.anchorCol
    });
  }

  const hasActiveSearch = Boolean(scratchSearchInput && scratchSearchInput.value.trim().length > 0);
  if (activeWorkspaceTab === sheetName) {
    if (needsExpand || touchedRow1 || (recordedChanges.length + formulaAffectedCoords.length) > 40 || hasActiveSearch) {
      renderActiveScratchSheetGrid();
    } else {
      const toRefresh = new Map();
      recordedChanges.forEach(ch => toRefresh.set(`${ch.r}::${ch.c}`, { r: ch.r, c: ch.c }));
      formulaAffectedCoords.forEach(pt => toRefresh.set(`${pt.r}::${pt.c}`, pt));
      toRefresh.forEach(pt => refreshSingleScratchCellDOM(sheetName, pt.r, pt.c));
      applyScratchSelectionDOM();
      syncScratchFormulaBarFromActiveCell();
      if (scratchStatsMeta) {
        const nextDims = getScratchSheetGridDimensions(ws, sheetName);
        scratchStatsMeta.textContent = `${bounds.nonEmptyRows} active row${bounds.nonEmptyRows === 1 ? '' : 's'} (${nextDims.rows} × ${nextDims.cols} grid)`;
      }
    }
  }

  updateScratchUndoRedoButtonsUI();
  updateScratchSaveBadgeUI();
  renderWorkspaceTabsBar();
  scratchMutationVersion++;
  if (typeof window !== 'undefined') {
    window.__lastTrackerScratchEditAt = Date.now();
  }
  scheduleScratchAutoSave();
  return recordedChanges.length;
}

function recordScratchCellEdit(sheetName, rNum, cNum, rawText, { coalesceKey = null } = {}) {
  scratchClipboardRangeState = null;
  applyScratchCellChanges(sheetName, [{ r: rNum, c: cNum, rawText }], {
    recordUndo: true,
    coalesceKey
  });
}

function applyScratchFillPresetToSelection(sheetName, presetId) {
  if (!workbook || !sheetName || sheetName === DAILY_TRACKER_WS_TAB) return 0;
  commitActiveScratchInlineEdit();
  const ws = workbook.getWorksheet(sheetName);
  if (!ws) return 0;

  const range = getScratchSelectionRange(sheetName);
  const dims = getScratchSheetGridDimensions(ws, sheetName);
  // If entire columns were highlighted via header click, cap row fill to active data bounds (or at least row 1)
  // unless user explicitly has empty cells below they want filled; wait: if full column selected (rows === dims.rows),
  // capping to Math.max(range.anchorRow, dims.dataMaxRow, 1) prevents coloring 40 blank rows unintentionally unless whole column has data,
  // OR coloring exact range if range.rowCount < dims.rows! Let's check:
  const isFullColumnSelection = range.minRow === 1 && range.maxRow >= dims.rows && range.colCount < dims.cols;
  const effectiveMaxRow = isFullColumnSelection
    ? Math.max(1, dims.dataMaxRow)
    : range.maxRow;

  const edits = [];
  for (let r = range.minRow; r <= effectiveMaxRow; r++) {
    for (let c = range.minCol; c <= range.maxCol; c++) {
      const beforeSnap = captureCellStyleSnapshot(ws, r, c);
      const nextSnap = buildStyleSnapshotForFillPreset(beforeSnap, presetId);
      edits.push({ r, c, styleSnap: nextSnap });
    }
  }

  if (presetId && presetId !== 'none') {
    activeScratchFillPresetId = presetId;
    try {
      localStorage.setItem(SCRATCH_LAST_FILL_STORAGE_KEY, presetId);
    } catch (_) { /* ignore */ }
    syncScratchFillIndicatorUI();
  }

  const count = applyScratchCellChanges(sheetName, edits, { recordUndo: true });
  if (count > 0) {
    const presetObj = EXCEL_DEFAULT_FILL_PALETTE.find(p => p.id === presetId);
    const label = presetId === 'none' ? 'Cleared fill color' : `Applied "${presetObj ? presetObj.label : 'Fill'}"`;
    if (count > 1) {
      showToast(`${label} on ${count} cells (${getRangeCoordString(range)}).`, 'info');
    }
  }
  return count;
}

function undoScratchEdit() {
  if (activeWorkspaceTab === DAILY_TRACKER_WS_TAB || !workbook) return;
  commitActiveScratchInlineEdit();
  const sheetName = activeWorkspaceTab;
  const undoStack = scratchUndoStackMap.get(sheetName);
  if (!undoStack || undoStack.length === 0) return;

  const entry = undoStack.pop();
  lastFormulaCoalesceMeta = null;

  const reverseEdits = entry.changes.map(ch => {
    const rev = { r: ch.r, c: ch.c };
    if (ch.textChanged || ch.textChanged === undefined) {
      rev.rawText = ch.beforeRaw;
    }
    if (ch.styleChanged) {
      rev.styleSnap = ch.beforeStyle;
    }
    return rev;
  });

  applyScratchCellChanges(sheetName, reverseEdits, {
    recordUndo: false,
    selectionAfter: entry.selectionBefore
  });

  if (!scratchRedoStackMap.has(sheetName)) {
    scratchRedoStackMap.set(sheetName, []);
  }
  scratchRedoStackMap.get(sheetName).push(entry);
  updateScratchUndoRedoButtonsUI();
  syncScratchFormulaBarFromActiveCell();
}

function redoScratchEdit() {
  if (activeWorkspaceTab === DAILY_TRACKER_WS_TAB || !workbook) return;
  commitActiveScratchInlineEdit();
  const sheetName = activeWorkspaceTab;
  const redoStack = scratchRedoStackMap.get(sheetName);
  if (!redoStack || redoStack.length === 0) return;

  const entry = redoStack.pop();
  lastFormulaCoalesceMeta = null;

  const forwardEdits = entry.changes.map(ch => {
    const fwd = { r: ch.r, c: ch.c };
    if (ch.textChanged || ch.textChanged === undefined) {
      fwd.rawText = ch.afterRaw;
    }
    if (ch.styleChanged) {
      fwd.styleSnap = ch.afterStyle;
    }
    return fwd;
  });

  applyScratchCellChanges(sheetName, forwardEdits, {
    recordUndo: false,
    selectionAfter: entry.selectionAfter
  });

  if (!scratchUndoStackMap.has(sheetName)) {
    scratchUndoStackMap.set(sheetName, []);
  }
  scratchUndoStackMap.get(sheetName).push(entry);
  updateScratchUndoRedoButtonsUI();
  syncScratchFormulaBarFromActiveCell();
}

function syncScratchFormulaBarFromActiveCell() {
  if (activeWorkspaceTab === DAILY_TRACKER_WS_TAB || !workbook) return;
  const ws = workbook.getWorksheet(activeWorkspaceTab);
  if (!ws) return;
  const range = getScratchSelectionRange(activeWorkspaceTab);
  if (scratchCellCoord) {
    scratchCellCoord.textContent = getRangeCoordString(range);
  }
  if (scratchFormulaInput && document.activeElement !== scratchFormulaInput) {
    const info = extractScratchCellInfo(ws, range.anchorRow, range.anchorCol);
    scratchFormulaInput.value = info.rawEdit;
  }
}

function startScratchInlineEdit(rNum, cNum, initialChar = null) {
  if (activeWorkspaceTab === DAILY_TRACKER_WS_TAB || !workbook || !scratchGridBody) return;
  scratchClipboardRangeState = null;
  commitActiveScratchInlineEdit();
  selectScratchCell(rNum, cNum);

  const ws = workbook.getWorksheet(activeWorkspaceTab);
  if (!ws) return;
  const td = scratchGridBody.querySelector(`td[data-r="${rNum}"][data-c="${cNum}"]`);
  if (!td) return;

  const info = extractScratchCellInfo(ws, rNum, cNum);
  const input = document.createElement('input');
  input.type = 'text';
  input.className = 'scratch-inline-editor';
  input.value = initialChar !== null ? initialChar : (info.isSpilled ? '' : info.rawEdit);

  scratchInlineEditState = {
    sheetName: activeWorkspaceTab,
    row: rNum,
    col: cNum,
    originalRaw: info.rawEdit,
    inputEl: input,
    tdEl: td
  };

  // Defer any pending background auto-save as soon as the user starts editing another cell
  if (scratchAutoSaveTimer) {
    scheduleScratchAutoSave(1800);
  }

  if (scratchFormulaInput) {
    scratchFormulaInput.value = input.value;
  }

  input.addEventListener('input', () => {
    if (scratchFormulaInput) {
      scratchFormulaInput.value = input.value;
    }
    if (scratchAutoSaveTimer) {
      scheduleScratchAutoSave(1800);
    }
  });

  input.addEventListener('keydown', (e) => {
    if (e.key === 'Enter') {
      e.preventDefault();
      e.stopPropagation();
      commitActiveScratchInlineEdit({ restoreGridFocus: true });
      selectScratchCell(e.shiftKey ? Math.max(1, rNum - 1) : rNum + 1, cNum, { scrollIntoView: true });
      if (scratchGridWrap) scratchGridWrap.focus({ preventScroll: true });
    } else if (e.key === 'Tab') {
      e.preventDefault();
      e.stopPropagation();
      commitActiveScratchInlineEdit({ restoreGridFocus: true });
      selectScratchCell(rNum, e.shiftKey ? Math.max(1, cNum - 1) : cNum + 1, { scrollIntoView: true });
      if (scratchGridWrap) scratchGridWrap.focus({ preventScroll: true });
    } else if (e.key === 'Escape') {
      e.preventDefault();
      e.stopPropagation();
      cancelActiveScratchInlineEdit();
      if (scratchGridWrap) scratchGridWrap.focus({ preventScroll: true });
    }
  });

  input.addEventListener('blur', () => {
    commitActiveScratchInlineEdit();
  });

  td.appendChild(input);
  input.focus();
  if (initialChar !== null) {
    input.setSelectionRange(input.value.length, input.value.length);
  } else {
    input.select();
  }
}

function renderActiveScratchSheetGrid() {
  if (!scratchGridHead || !scratchGridBody || !workbook || activeWorkspaceTab === DAILY_TRACKER_WS_TAB) {
    return;
  }
  const ws = workbook.getWorksheet(activeWorkspaceTab);
  if (!ws) {
    activeWorkspaceTab = DAILY_TRACKER_WS_TAB;
    syncWorkspaceTabsUI();
    return;
  }

  const prevScrollTop = scratchGridWrap ? scratchGridWrap.scrollTop : 0;
  const prevScrollLeft = scratchGridWrap ? scratchGridWrap.scrollLeft : 0;

  recalculateScratchWorksheetFormulas(ws, activeWorkspaceTab);
  const dims = getScratchSheetGridDimensions(ws, activeWorkspaceTab);
  const anchorSpills = scratchSpillStateMap.get(activeWorkspaceTab);
  const range = getScratchSelectionRange(activeWorkspaceTab);
  const query = scratchSearchInput ? scratchSearchInput.value.trim().toLowerCase() : '';
  const filterRowsOnly = Boolean(chkScratchFilterRows && chkScratchFilterRows.checked && query.length > 0);

  if (scratchActiveSheetBadge) {
    scratchActiveSheetBadge.textContent = ws.name;
  }
  if (scratchStatsMeta) {
    scratchStatsMeta.textContent = `${dims.nonEmptyRows} active row${dims.nonEmptyRows === 1 ? '' : 's'} (${dims.rows} × ${dims.cols} grid)`;
  }
  syncScratchStripedRowsUI();
  updateScratchSaveBadgeUI();
  updateScratchUndoRedoButtonsUI();

  // Build Sticky Column Header Row (with clickable headers + drag-to-resize handles)
  let headHtml = `<tr><th class="scratch-corner-th" title="Click to select all cells (Ctrl+A)">#</th>`;
  for (let c = 1; c <= dims.cols; c++) {
    const excelCol = ws.getColumn(c);
    const colW = excelCol && excelCol.width
      ? Math.max(68, Math.min(500, Math.round(excelCol.width * 7.6)))
      : 132;
    const isActiveCol = c >= range.minCol && c <= range.maxCol;
    const colLetter = colIndexToExcelLetter(c);
    headHtml += `<th data-col-hdr="${c}" class="${isActiveCol ? 'active-col-hdr' : ''}" style="width:${colW}px; min-width:${colW}px;" title="Click or drag to select Column ${colLetter}"><span class="scratch-col-label">${colLetter}</span><span class="scratch-col-resizer" data-resize-col="${c}" title="Drag to resize Column ${colLetter} • Double-click to auto-fit"></span></th>`;
  }
  headHtml += `</tr>`;
  scratchGridHead.innerHTML = headHtml;

  // Check whether Row 1 looks like a header row
  let row1HasValues = false;
  for (let c = 1; c <= dims.cols; c++) {
    if (extractScratchCellInfo(ws, 1, c).display.trim() !== '') {
      row1HasValues = true;
      break;
    }
  }

  let matchCellCount = 0;
  const rowsHtml = [];

  for (let r = 1; r <= dims.rows; r++) {
    let rowMatchesSearch = false;
    const rowCellsHtml = [];

    for (let c = 1; c <= dims.cols; c++) {
      const info = extractScratchCellInfo(ws, r, c);
      const visual = extractScratchCellVisualStyle(ws, r, c);
      const isHit = query.length > 0 && (
        info.display.toLowerCase().includes(query) ||
        info.rawEdit.toLowerCase().includes(query)
      );
      if (isHit) {
        rowMatchesSearch = true;
        matchCellCount++;
      }

      const classes = ['scratch-cell'];
      if (r === 1 && row1HasValues) classes.push('header-like-cell');
      if (visual.isBold) classes.push('bold-cell');
      if (visual.bgHex) classes.push('has-custom-fill');
      else if (visual.fgHex) classes.push('has-custom-fg');
      if (info.isFormula) classes.push('formula-cell');
      if (anchorSpills && anchorSpills.has(`${r}::${c}`)) classes.push('formula-anchor-cell');
      if (info.isSpilled && info.anchorR > 0) classes.push('spilled-array-cell');
      if (isHit) classes.push('search-hit');

      let styleAttr = '';
      if (visual.bgHex) {
        styleAttr = ` style="--cell-fill-bg:${visual.bgHex}; --cell-fill-fg:${visual.fgHex || getContrastingTextHexForBg(visual.bgHex)};"`;
      } else if (visual.fgHex) {
        styleAttr = ` style="--cell-fill-fg:${visual.fgHex};"`;
      }
      const titleText = info.isFormula
        ? `${info.rawEdit} → ${info.display}`
        : (info.isSpilled && info.anchorR > 0 ? `${info.display} (Spilled from ${getCellCoordString(info.anchorR, info.anchorC)})` : info.display);
      rowCellsHtml.push(
        `<td data-r="${r}" data-c="${c}" class="${classes.join(' ')}"${styleAttr} title="${escapeHtml(titleText)}"><span class="scratch-cell-text">${escapeHtml(info.display)}</span></td>`
      );
    }

    // Keep Row 1 header visible when filtering rows, plus any matching row
    if (filterRowsOnly && r > 1 && !rowMatchesSearch) {
      continue;
    }

    rowsHtml.push(
      `<tr><th data-row-hdr="${r}" class="scratch-row-hdr" title="Click or drag to select Row ${r}">${r}</th>${rowCellsHtml.join('')}</tr>`
    );
  }

  scratchGridBody.innerHTML = rowsHtml.join('');

  if (scratchGridWrap) {
    scratchGridWrap.scrollTop = prevScrollTop;
    scratchGridWrap.scrollLeft = prevScrollLeft;
  }

  if (scratchSearchMatchCount) {
    if (query.length > 0) {
      scratchSearchMatchCount.style.display = 'inline-block';
      scratchSearchMatchCount.textContent = `${matchCellCount} match${matchCellCount === 1 ? '' : 'es'}`;
    } else {
      scratchSearchMatchCount.style.display = 'none';
    }
  }

  applyScratchSelectionDOM();
  syncScratchFormulaBarFromActiveCell();
}

// --- POINTER DRAG RANGE SELECTION + SMOOTH EDGE AUTO-SCROLL ---
function calcScratchAutoScrollVelocity(distanceIntoZone) {
  return Math.min(26, Math.max(3, Math.round(distanceIntoZone * 0.3)));
}

function updateScratchDragSelectionFromPoint(clientX, clientY) {
  if (!scratchDragState || !scratchGridWrap || !workbook) return;
  if (activeWorkspaceTab !== scratchDragState.sheetName) {
    stopScratchPointerDrag();
    return;
  }
  const ws = workbook.getWorksheet(activeWorkspaceTab);
  if (!ws) return;

  const rect = scratchGridWrap.getBoundingClientRect();
  const headerH = scratchGridHead ? (scratchGridHead.offsetHeight || 28) : 28;
  const rowHdrW = 48;
  const dims = getScratchSheetGridDimensions(ws, activeWorkspaceTab);
  const curRange = getScratchSelectionRange(activeWorkspaceTab);

  if (scratchDragState.mode === 'rows') {
    const probeX = Math.max(rect.left + 12, Math.min(rect.right - 8, rect.left + 20));
    const probeY = Math.max(rect.top + headerH + 3, Math.min(rect.bottom - 3, clientY));
    const hit = document.elementFromPoint(probeX, probeY);
    if (!hit) return;
    const rowEl = hit.closest('th[data-row-hdr], td[data-r]');
    if (!rowEl) return;
    const targetRow = Number(rowEl.dataset.rowHdr || rowEl.dataset.r);
    if (targetRow >= 1 && (targetRow !== curRange.focusRow || curRange.focusCol !== dims.cols)) {
      setScratchSelectionRange(activeWorkspaceTab, {
        anchorRow: curRange.anchorRow,
        anchorCol: 1,
        focusRow: targetRow,
        focusCol: dims.cols
      });
    }
    return;
  }

  if (scratchDragState.mode === 'cols') {
    const probeX = Math.max(rect.left + rowHdrW + 3, Math.min(rect.right - 3, clientX));
    const probeY = Math.max(rect.top + 6, Math.min(rect.bottom - 6, rect.top + 12));
    const hit = document.elementFromPoint(probeX, probeY);
    if (!hit) return;
    const colEl = hit.closest('th[data-col-hdr], td[data-c]');
    if (!colEl) return;
    const targetCol = Number(colEl.dataset.colHdr || colEl.dataset.c);
    if (targetCol >= 1 && (targetCol !== curRange.focusCol || curRange.focusRow !== dims.rows)) {
      setScratchSelectionRange(activeWorkspaceTab, {
        anchorRow: 1,
        anchorCol: curRange.anchorCol,
        focusRow: dims.rows,
        focusCol: targetCol
      });
    }
    return;
  }

  if (scratchDragState.mode === 'fill') {
    const probeX = Math.max(rect.left + rowHdrW + 3, Math.min(rect.right - 3, clientX));
    const probeY = Math.max(rect.top + headerH + 3, Math.min(rect.bottom - 3, clientY));
    const hit = document.elementFromPoint(probeX, probeY);
    if (!hit) return;
    const td = hit.closest('td[data-r][data-c]');
    if (!td) return;
    const r = Number(td.dataset.r);
    const c = Number(td.dataset.c);
    if (r >= 1 && c >= 1) {
      const nextTarget = computeScratchFillTargetRange(scratchDragState.sourceRange, r, c);
      const prevKey = scratchDragState.fillTargetRange ? JSON.stringify(scratchDragState.fillTargetRange) : '';
      const nextKey = nextTarget ? JSON.stringify(nextTarget) : '';
      if (prevKey !== nextKey) {
        scratchDragState.fillTargetRange = nextTarget;
        applyScratchSelectionDOM();
      }
    }
    return;
  }

  // Standard rectangular cell drag ('cells')
  const probeX = Math.max(rect.left + rowHdrW + 3, Math.min(rect.right - 3, clientX));
  const probeY = Math.max(rect.top + headerH + 3, Math.min(rect.bottom - 3, clientY));
  const hit = document.elementFromPoint(probeX, probeY);
  if (!hit) return;
  const td = hit.closest('td[data-r][data-c]');
  if (!td) return;
  const r = Number(td.dataset.r);
  const c = Number(td.dataset.c);
  if (r >= 1 && c >= 1 && (r !== curRange.focusRow || c !== curRange.focusCol)) {
    setScratchSelectionRange(activeWorkspaceTab, {
      anchorRow: curRange.anchorRow,
      anchorCol: curRange.anchorCol,
      focusRow: r,
      focusCol: c
    });
  }
}

function runScratchDragAutoScrollTick() {
  if (!scratchDragState || !scratchGridWrap || !workbook) return;
  const rect = scratchGridWrap.getBoundingClientRect();
  const headerH = scratchGridHead ? (scratchGridHead.offsetHeight || 28) : 28;
  const rowHdrW = 48;
  const edgeZone = 40;

  let vx = 0;
  let vy = 0;
  const { clientX, clientY, mode } = scratchDragState;

  if (mode !== 'cols') {
    const topZoneEdge = rect.top + headerH + edgeZone;
    const bottomZoneEdge = rect.bottom - edgeZone;
    if (clientY < topZoneEdge) {
      vy = -calcScratchAutoScrollVelocity(topZoneEdge - clientY);
    } else if (clientY > bottomZoneEdge) {
      vy = calcScratchAutoScrollVelocity(clientY - bottomZoneEdge);
    }
  }

  if (mode !== 'rows') {
    const leftZoneEdge = rect.left + rowHdrW + edgeZone;
    const rightZoneEdge = rect.right - edgeZone;
    if (clientX < leftZoneEdge) {
      vx = -calcScratchAutoScrollVelocity(leftZoneEdge - clientX);
    } else if (clientX > rightZoneEdge) {
      vx = calcScratchAutoScrollVelocity(clientX - rightZoneEdge);
    }
  }

  if (vx !== 0 || vy !== 0) {
    const prevTop = scratchGridWrap.scrollTop;
    const prevLeft = scratchGridWrap.scrollLeft;
    scratchGridWrap.scrollTop += vy;
    scratchGridWrap.scrollLeft += vx;

    // Auto-expand grid smoothly when dragging past the bottom/right scroll boundary
    const now = Date.now();
    if (now - scratchDragState.lastExpandAt > 200) {
      const ws = workbook.getWorksheet(activeWorkspaceTab);
      if (ws) {
        const dims = getScratchSheetGridDimensions(ws, activeWorkspaceTab);
        const atBottomLimit =
          vy > 0 &&
          clientY > rect.bottom + 8 &&
          scratchGridWrap.scrollTop === prevTop &&
          scratchGridWrap.scrollTop + scratchGridWrap.clientHeight >= scratchGridWrap.scrollHeight - 3;
        const atRightLimit =
          vx > 0 &&
          clientX > rect.right + 8 &&
          scratchGridWrap.scrollLeft === prevLeft &&
          scratchGridWrap.scrollLeft + scratchGridWrap.clientWidth >= scratchGridWrap.scrollWidth - 3;

        if (atBottomLimit || atRightLimit) {
          scratchDragState.lastExpandAt = now;
          scratchGridDimensionsMap.set(activeWorkspaceTab, {
            rows: atBottomLimit ? dims.rows + 5 : dims.rows,
            cols: atRightLimit ? dims.cols + 1 : dims.cols
          });
          renderActiveScratchSheetGrid();
          if (atBottomLimit) scratchGridWrap.scrollTop += 40;
          if (atRightLimit) scratchGridWrap.scrollLeft += 60;
        }
      }
    }

    updateScratchDragSelectionFromPoint(clientX, clientY);
  }

  if (scratchDragState) {
    scratchDragState.rafId = requestAnimationFrame(runScratchDragAutoScrollTick);
  }
}

function startScratchPointerDrag(mode, clientX, clientY, extraState = null) {
  stopScratchPointerDrag();
  if (activeWorkspaceTab === DAILY_TRACKER_WS_TAB) return;
  scratchDragState = {
    mode,
    sheetName: activeWorkspaceTab,
    clientX,
    clientY,
    rafId: null,
    lastExpandAt: 0,
    ...(extraState || {})
  };
  if (scratchGridWrap) {
    scratchGridWrap.classList.add('is-drag-selecting');
    if (mode === 'fill') {
      scratchGridWrap.classList.add('is-fill-dragging');
    }
  }
  scratchDragState.rafId = requestAnimationFrame(runScratchDragAutoScrollTick);
}

function stopScratchPointerDrag() {
  const wasFill = Boolean(scratchDragState && scratchDragState.mode === 'fill');
  if (scratchDragState) {
    if (scratchDragState.rafId) {
      cancelAnimationFrame(scratchDragState.rafId);
    }
    scratchDragState = null;
  }
  if (scratchGridWrap) {
    scratchGridWrap.classList.remove('is-drag-selecting');
    scratchGridWrap.classList.remove('is-fill-dragging');
  }
  if (wasFill) {
    applyScratchSelectionDOM();
  }
}

// --- COLUMN WIDTH DRAG-RESIZE & DOUBLE-CLICK AUTO-FIT ---
function recordScratchColumnWidthChange(sheetName, colNum, widthPx) {
  if (!workbook || !sheetName || sheetName === DAILY_TRACKER_WS_TAB) return;
  const ws = workbook.getWorksheet(sheetName);
  if (!ws) return;
  const clampedPx = Math.max(68, Math.min(500, Math.round(widthPx)));
  const excelWidth = +(clampedPx / 7.6).toFixed(1);
  ws.getColumn(colNum).width = excelWidth;

  if (!scratchPendingColWidths.has(sheetName)) {
    scratchPendingColWidths.set(sheetName, new Map());
  }
  scratchPendingColWidths.get(sheetName).set(colNum, excelWidth);
  scratchMutationVersion++;
  updateScratchSaveBadgeUI();
  renderWorkspaceTabsBar();
  scheduleScratchAutoSave();
}

function autoFitScratchColumnWidth(sheetName, colNum) {
  if (!workbook || !sheetName || sheetName === DAILY_TRACKER_WS_TAB || !scratchGridHead) return;
  const ws = workbook.getWorksheet(sheetName);
  if (!ws) return;
  const dims = getScratchSheetGridDimensions(ws, sheetName);

  let maxLen = colIndexToExcelLetter(colNum).length + 2;
  for (let r = 1; r <= dims.rows; r++) {
    const disp = extractScratchCellInfo(ws, r, colNum).display;
    if (disp.length > maxLen) {
      maxLen = disp.length;
    }
  }
  const fittedPx = Math.max(80, Math.min(440, Math.round(maxLen * 7.8 + 26)));
  const th = scratchGridHead.querySelector(`th[data-col-hdr="${colNum}"]`);
  if (th) {
    th.style.width = `${fittedPx}px`;
    th.style.minWidth = `${fittedPx}px`;
  }
  recordScratchColumnWidthChange(sheetName, colNum, fittedPx);
}

// --- KEYBOARD DATA-BOUNDARY JUMP (Ctrl+Arrow) & CLIPBOARD HELPERS ---
function computeScratchCtrlArrowTarget(ws, dims, startRow, startCol, dr, dc) {
  const hasVal = (r, c) => extractScratchCellInfo(ws, r, c).display.trim() !== '';
  const inBounds = (r, c) => r >= 1 && r <= dims.rows && c >= 1 && c <= dims.cols;

  let nextR = startRow + dr;
  let nextC = startCol + dc;
  if (!inBounds(nextR, nextC)) {
    return { row: startRow, col: startCol };
  }

  if (hasVal(startRow, startCol) && hasVal(nextR, nextC)) {
    // Walk to the last non-empty cell of the current contiguous block
    let curR = nextR;
    let curC = nextC;
    while (inBounds(curR + dr, curC + dc) && hasVal(curR + dr, curC + dc)) {
      curR += dr;
      curC += dc;
    }
    return { row: curR, col: curC };
  } else {
    // Skip empty cells until landing on the next non-empty cell (or grid edge)
    let curR = nextR;
    let curC = nextC;
    while (inBounds(curR, curC) && !hasVal(curR, curC)) {
      if (!inBounds(curR + dr, curC + dc)) {
        return { row: curR, col: curC };
      }
      curR += dr;
      curC += dc;
    }
    return { row: curR, col: curC };
  }
}

function resolveEffectiveScratchCopyRange(ws, range) {
  if (!ws || !range) return range;
  const dims = getScratchSheetGridDimensions(ws, ws.name || activeWorkspaceTab);
  const fullHeight = range.minRow === 1 && range.maxRow >= dims.rows;
  const fullWidth = range.minCol === 1 && range.maxCol >= dims.cols;
  if (!fullHeight && !fullWidth) return range;

  let lastRow = range.minRow;
  let lastCol = range.minCol;
  for (let r = range.minRow; r <= range.maxRow; r++) {
    for (let c = range.minCol; c <= range.maxCol; c++) {
      const info = extractScratchCellInfo(ws, r, c);
      const visual = extractScratchCellVisualStyle(ws, r, c);
      if (info.display.trim() !== '' || visual.bgHex || visual.isBold) {
        if (r > lastRow) lastRow = r;
        if (c > lastCol) lastCol = c;
      }
    }
  }
  return {
    ...range,
    maxRow: fullHeight ? Math.max(range.minRow, lastRow) : range.maxRow,
    maxCol: fullWidth ? Math.max(range.minCol, lastCol) : range.maxCol
  };
}

function serializeScratchSelectionToTSV(ws, range) {
  const lines = [];
  for (let r = range.minRow; r <= range.maxRow; r++) {
    const rowVals = [];
    for (let c = range.minCol; c <= range.maxCol; c++) {
      const disp = extractScratchCellInfo(ws, r, c).display.replace(/[\t\r\n]+/g, ' ');
      rowVals.push(disp);
    }
    lines.push(rowVals.join('\t'));
  }
  return lines.join('\r\n');
}

function serializeScratchSelectionToExcelHtml(ws, range) {
  if (!ws || !range) return '';
  const dims = getScratchSheetGridDimensions(ws, ws.name || activeWorkspaceTab);
  let row1HasValues = false;
  for (let c = 1; c <= dims.cols; c++) {
    if (extractScratchCellInfo(ws, 1, c).display.trim() !== '') {
      row1HasValues = true;
      break;
    }
  }

  const colTags = [];
  for (let c = range.minCol; c <= range.maxCol; c++) {
    const excelCol = ws.getColumn ? ws.getColumn(c) : null;
    const colW = excelCol && excelCol.width
      ? Math.max(68, Math.min(500, Math.round(excelCol.width * 7.6)))
      : 110;
    colTags.push(`<col width="${colW}" style="width:${colW}px;">`);
  }

  const trHtml = [];
  for (let r = range.minRow; r <= range.maxRow; r++) {
    const tdHtml = [];
    for (let c = range.minCol; c <= range.maxCol; c++) {
      const info = extractScratchCellInfo(ws, r, c);
      const visual = extractScratchCellVisualStyle(ws, r, c);
      const isAutoHeader = r === 1 && row1HasValues;

      let bgHex = visual.bgHex;
      let fgHex = visual.fgHex;
      let isBold = visual.isBold;
      if (isAutoHeader && !bgHex) {
        bgHex = '#D9E1F2';
        fgHex = fgHex || '#0F172A';
        isBold = true;
      } else if (bgHex) {
        fgHex = fgHex || getContrastingTextHexForBg(bgHex);
      } else if (fgHex && fgHex.toUpperCase() === '#FFFFFF') {
        fgHex = '#000000';
      }

      const styles = [
        'border:1px solid #BFBFBF',
        'padding:3px 8px',
        "font-family:Calibri, Aptos, 'Segoe UI', Arial, sans-serif",
        'font-size:11pt',
        'vertical-align:middle',
        'white-space:nowrap'
      ];
      if (bgHex) {
        styles.push(`background:${bgHex}`, `background-color:${bgHex}`);
      }
      styles.push(`color:${fgHex || '#000000'}`);
      if (isBold) {
        styles.push('font-weight:700');
      }
      if (visual.align) {
        styles.push(`text-align:${visual.align}`);
      }

      const bgAttr = bgHex ? ` bgcolor="${bgHex}"` : '';
      const dispTrim = info.display.trim();
      const cellContent = dispTrim === ''
        ? '&nbsp;'
        : escapeHtml(info.display).replace(/\r\n|\r|\n/g, '<br>');

      tdHtml.push(`<td${bgAttr} style="${styles.join('; ')};">${cellContent}</td>`);
    }
    trHtml.push(`<tr>${tdHtml.join('')}</tr>`);
  }

  return `<!--StartFragment--><table border="1" cellspacing="0" cellpadding="0" style="border-collapse:collapse; border:1px solid #BFBFBF; font-family:Calibri, Aptos, 'Segoe UI', Arial, sans-serif; font-size:11pt; color:#000000;"><colgroup>${colTags.join('')}</colgroup><tbody>${trHtml.join('')}</tbody></table><!--EndFragment-->`;
}

function buildScratchStyleMatrix(ws, range) {
  const matrix = [];
  for (let r = range.minRow; r <= range.maxRow; r++) {
    const rowSnaps = [];
    for (let c = range.minCol; c <= range.maxCol; c++) {
      rowSnaps.push(captureCellStyleSnapshot(ws, r, c));
    }
    matrix.push(rowSnaps);
  }
  return matrix;
}

function buildScratchRawEditMatrix(ws, range) {
  const matrix = [];
  for (let r = range.minRow; r <= range.maxRow; r++) {
    const rowVals = [];
    for (let c = range.minCol; c <= range.maxCol; c++) {
      const info = extractScratchCellInfo(ws, r, c);
      rowVals.push({ rawEdit: info.rawEdit, isFormula: info.isFormula });
    }
    matrix.push(rowVals);
  }
  return matrix;
}

function parseCssColorToArgb(cssColorStr) {
  const raw = String(cssColorStr || '').trim().toLowerCase();
  if (!raw || raw === 'transparent' || raw === 'none' || raw === 'inherit' || raw === 'initial' || raw === 'window' || raw === 'windowtext' || raw === 'auto') {
    return null;
  }
  const named = {
    white: 'FFFFFFFF',
    black: 'FF000000',
    red: 'FFFF0000',
    green: 'FF008000',
    blue: 'FF0000FF',
    yellow: 'FFFFFF00',
    gray: 'FF808080',
    grey: 'FF808080',
    orange: 'FFFFA500',
    purple: 'FF800080',
    navy: 'FF000080',
    silver: 'FFC0C0C0'
  };
  if (named[raw]) return named[raw];

  const hexMatch = raw.match(/^#([0-9a-f]{3}|[0-9a-f]{6}|[0-9a-f]{8})$/i);
  if (hexMatch) {
    const h = hexMatch[1].toUpperCase();
    if (h.length === 3) {
      return `FF${h[0]}${h[0]}${h[1]}${h[1]}${h[2]}${h[2]}`;
    }
    if (h.length === 6) {
      return `FF${h}`;
    }
    if (h.length === 8) {
      return h.slice(0, 2) === '00' ? null : h;
    }
  }

  const rgbMatch = raw.match(/^rgba?\(\s*(\d{1,3})\s*,\s*(\d{1,3})\s*,\s*(\d{1,3})(?:\s*[,/]\s*([0-9.]+))?\s*\)$/i);
  if (rgbMatch) {
    if (rgbMatch[4] !== undefined && parseFloat(rgbMatch[4]) === 0) return null;
    const toHex = (n) => Math.max(0, Math.min(255, parseInt(n, 10) || 0)).toString(16).padStart(2, '0').toUpperCase();
    return `FF${toHex(rgbMatch[1])}${toHex(rgbMatch[2])}${toHex(rgbMatch[3])}`;
  }
  return null;
}

function parseStyleDeclarationsMap(declText) {
  const out = {};
  String(declText || '').split(';').forEach(part => {
    const colonIdx = part.indexOf(':');
    if (colonIdx === -1) return;
    const prop = part.slice(0, colonIdx).trim().toLowerCase();
    const val = part.slice(colonIdx + 1).replace(/!important/gi, '').trim();
    if (prop) out[prop] = val;
  });
  return out;
}

function resolveColorStyleSnapshotFromProps({ bgRaw, fgRaw, weightRaw, tagName = 'td', hasBoldTag = false }) {
  let bgArgb = parseCssColorToArgb(bgRaw);
  if (bgArgb === 'FFFFFFFF') bgArgb = null;

  let fgArgb = parseCssColorToArgb(fgRaw);
  if (!bgArgb && (fgArgb === 'FF000000' || fgArgb === 'FF0F172A')) {
    fgArgb = null;
  }

  const wLower = String(weightRaw || '').trim().toLowerCase();
  const isBold = hasBoldTag || tagName === 'th' || wLower === 'bold' || wLower === 'bolder' || (parseInt(wLower, 10) >= 600);

  const fill = bgArgb
    ? { type: 'pattern', pattern: 'solid', fgColor: { argb: bgArgb } }
    : null;

  const fontObj = {};
  if (fgArgb) fontObj.color = { argb: fgArgb };
  else if (bgArgb) fontObj.color = { argb: getContrastingTextHexForBg(`#${bgArgb.slice(2)}`).replace('#', 'FF') };
  if (isBold) fontObj.bold = true;
  const font = Object.keys(fontObj).length > 0 ? fontObj : null;

  return { fill, font, hasVisualStyle: Boolean(fill || font) };
}

function extractHtmlTableCellText(cellEl) {
  if (!cellEl) return '';
  const BLOCK_TAGS = new Set(['P', 'DIV', 'LI', 'H1', 'H2', 'H3', 'H4', 'H5', 'H6', 'BLOCKQUOTE', 'PRE']);
  const SKIP_TAGS = new Set(['STYLE', 'SCRIPT', 'NOSCRIPT', 'META', 'LINK']);
  let out = '';

  const appendNewlineIfNeeded = () => {
    if (out.length > 0 && !out.endsWith('\n')) {
      out += '\n';
    }
  };

  const walk = (node, inPre = false) => {
    if (!node) return;
    if (node.nodeType === 3) {
      let txt = String(node.nodeValue || '').replace(/\u00A0/g, ' ');
      if (inPre) {
        txt = txt.replace(/\r\n/g, '\n').replace(/\r/g, '\n');
      } else {
        txt = txt.replace(/[\r\n\t]+/g, ' ').replace(/ {2,}/g, ' ');
      }
      out += txt;
      return;
    }
    if (node.nodeType !== 1) return;
    const tag = String(node.tagName || '').toUpperCase();
    if (SKIP_TAGS.has(tag)) return;
    if (tag === 'BR') {
      out += '\n';
      return;
    }
    const isBlock = BLOCK_TAGS.has(tag);
    if (isBlock) {
      appendNewlineIfNeeded();
    }
    const childInPre = inPre || tag === 'PRE';
    const children = node.childNodes ? Array.from(node.childNodes) : [];
    children.forEach(ch => walk(ch, childInPre));
    if (isBlock) {
      appendNewlineIfNeeded();
    }
  };

  walk(cellEl, false);

  const lines = out.split('\n').map(line => line.replace(/^[ \t]+|[ \t]+$/g, ''));
  while (lines.length > 0 && lines[0] === '') lines.shift();
  while (lines.length > 0 && lines[lines.length - 1] === '') lines.pop();
  return lines.join('\n');
}

function extractHtmlCellTextFallback(innerHtml) {
  let str = String(innerHtml || '');
  str = str.replace(/<(style|script)\b[^>]*>[\s\S]*?<\/\1>/gi, '');
  str = str.replace(/[\r\n\t]+/g, ' ');
  str = str.replace(/<br\s*\/?>/gi, '\n');
  str = str.replace(/<\/(?:p|div|li|h[1-6]|blockquote|pre)\s*>/gi, '\n');
  str = str.replace(/<(?:p|div|li|h[1-6]|blockquote|pre)\b[^>]*>/gi, '');
  str = str.replace(/<[^>]+>/g, '');
  str = str
    .replace(/&nbsp;/gi, ' ')
    .replace(/&#160;/gi, ' ')
    .replace(/&amp;/gi, '&')
    .replace(/&lt;/gi, '<')
    .replace(/&gt;/gi, '>')
    .replace(/&quot;/gi, '"')
    .replace(/&#0*39;|&#x0*27;/gi, "'")
    .replace(/&#0*10;|&#x0*a;/gi, '\n');
  const lines = str.split('\n').map(line => line.replace(/ {2,}/g, ' ').replace(/^[ \t]+|[ \t]+$/g, ''));
  while (lines.length > 0 && lines[0] === '') lines.shift();
  while (lines.length > 0 && lines[lines.length - 1] === '') lines.pop();
  return lines.join('\n');
}

function parseClipboardHtmlTableToMatrix(htmlText) {
  const raw = String(htmlText || '');
  if (!raw || !/<table[\s>]/i.test(raw)) return null;

  const classStyleMap = new Map();
  const styleBlockRegex = /<style[^>]*>([\s\S]*?)<\/style>/gi;
  let sMatch;
  while ((sMatch = styleBlockRegex.exec(raw)) !== null) {
    const ruleRegex = /([^{}]+)\{([^}]*)\}/g;
    let rMatch;
    while ((rMatch = ruleRegex.exec(sMatch[1])) !== null) {
      const selectors = rMatch[1].split(',');
      const decls = parseStyleDeclarationsMap(rMatch[2]);
      selectors.forEach(sel => {
        const clsMatch = sel.trim().match(/^(?:[a-z0-9*]+\s*)?\.([a-z0-9_-]+)$/i);
        if (clsMatch) {
          const prev = classStyleMap.get(clsMatch[1]) || {};
          classStyleMap.set(clsMatch[1], { ...prev, ...decls });
        }
      });
    }
  }

  let tableSourceHtml = raw;
  const fragMatch = raw.match(/<!--StartFragment-->([\s\S]*?)<!--EndFragment-->/i);
  if (fragMatch && /<table[\s>]/i.test(fragMatch[1])) {
    tableSourceHtml = fragMatch[1];
  }

  const extractBackgroundProp = (decls) => {
    if (!decls) return '';
    if (decls['background-color']) return decls['background-color'];
    if (decls.background) {
      const token = String(decls.background).split(/\s+/).find(tok => /^(#[0-9a-f]{3,8}|rgba?\(|white|black|red|green|blue|yellow|gray|grey|orange|purple|navy|silver)$/i.test(tok));
      if (token) return token;
    }
    return '';
  };

  let valueRows = null;
  let styleRows = null;
  let anyStyle = false;

  if (typeof DOMParser !== 'undefined') {
    try {
      const doc = new DOMParser().parseFromString(tableSourceHtml, 'text/html');
      const table = doc.querySelector('table');
      const trs = table
        ? Array.from(table.querySelectorAll('tr')).filter(tr => !tr.closest || tr.closest('table') === table)
        : [];
      if (trs.length > 0) {
        const occupied = [];
        valueRows = [];
        styleRows = [];
        trs.forEach((tr, rIdx) => {
          if (!occupied[rIdx]) occupied[rIdx] = [];
          if (!valueRows[rIdx]) valueRows[rIdx] = [];
          if (!styleRows[rIdx]) styleRows[rIdx] = [];
          const rawChildren = tr.children ? Array.from(tr.children) : Array.from(tr.querySelectorAll('th, td'));
          const cells = rawChildren.filter(el => /^(TH|TD)$/i.test(el.tagName) && (!el.closest || el.closest('tr') === tr));
          let cCursor = 0;
          cells.forEach(cell => {
            while (occupied[rIdx][cCursor]) cCursor++;
            const clsDecls = {};
            String(cell.getAttribute('class') || '').split(/\s+/).filter(Boolean).forEach(clsName => {
              const found = classStyleMap.get(clsName);
              if (found) Object.assign(clsDecls, found);
            });
            const inlineDecls = parseStyleDeclarationsMap(cell.getAttribute('style') || '');
            const merged = { ...clsDecls, ...inlineDecls };
            const bgRaw = extractBackgroundProp(merged) || cell.getAttribute('bgcolor') || '';
            const fgRaw = merged.color || cell.getAttribute('color') || '';
            const weightRaw = merged['font-weight'] || '';
            const hasBoldTag = Boolean(cell.querySelector('b, strong'));
            const snap = resolveColorStyleSnapshotFromProps({
              bgRaw,
              fgRaw,
              weightRaw,
              tagName: cell.tagName.toLowerCase(),
              hasBoldTag
            });
            if (snap.hasVisualStyle) anyStyle = true;
            const cellText = extractHtmlTableCellText(cell);
            const colspan = Math.min(500, Math.max(1, parseInt(cell.getAttribute('colspan') || '1', 10) || 1));
            const rowspan = Math.min(500, Math.max(1, parseInt(cell.getAttribute('rowspan') || '1', 10) || 1));
            const styleObj = { fill: snap.fill, font: snap.font };

            for (let dr = 0; dr < rowspan; dr++) {
              const targetR = rIdx + dr;
              if (!occupied[targetR]) occupied[targetR] = [];
              if (!valueRows[targetR]) valueRows[targetR] = [];
              if (!styleRows[targetR]) styleRows[targetR] = [];
              for (let dc = 0; dc < colspan; dc++) {
                const targetC = cCursor + dc;
                occupied[targetR][targetC] = true;
                valueRows[targetR][targetC] = (dr === 0 && dc === 0) ? cellText : '';
                styleRows[targetR][targetC] = styleObj;
              }
            }
            cCursor += colspan;
          });
        });
      }
    } catch (_) {
      valueRows = null;
      styleRows = null;
    }
  }

  if (!valueRows) {
    // Lightweight fallback parser when DOMParser is unavailable (e.g. non-browser environment)
    const tableMatch = tableSourceHtml.match(/<table\b[^>]*>[\s\S]*?<\/table>/i);
    const fallbackHtml = tableMatch ? tableMatch[0] : tableSourceHtml;
    const trMatches = fallbackHtml.match(/<tr\b[^>]*>[\s\S]*?<\/tr>/gi) || [];
    if (trMatches.length === 0) return null;
    const occupied = [];
    valueRows = [];
    styleRows = [];
    trMatches.forEach((trStr, rIdx) => {
      if (!occupied[rIdx]) occupied[rIdx] = [];
      if (!valueRows[rIdx]) valueRows[rIdx] = [];
      if (!styleRows[rIdx]) styleRows[rIdx] = [];
      const cellMatches = trStr.match(/<(td|th)\b[^>]*>[\s\S]*?<\/\1>/gi) || [];
      let cCursor = 0;
      cellMatches.forEach(cellStr => {
        const tagMatch = cellStr.match(/^<(td|th)\b([^>]*)>([\s\S]*?)<\/\1>/i);
        if (!tagMatch) return;
        while (occupied[rIdx][cCursor]) cCursor++;
        const tagName = tagMatch[1].toLowerCase();
        const attrs = tagMatch[2] || '';
        const inner = tagMatch[3] || '';
        const clsAttr = (attrs.match(/\bclass\s*=\s*["']([^"']*)["']/i) || [])[1] || '';
        const styleAttr = (attrs.match(/\bstyle\s*=\s*["']([^"']*)["']/i) || [])[1] || '';
        const bgcolorAttr = (attrs.match(/\bbgcolor\s*=\s*["']([^"']*)["']/i) || [])[1] || '';
        const colspanAttr = Math.min(500, Math.max(1, parseInt((attrs.match(/\bcolspan\s*=\s*["']?(\d+)/i) || [])[1] || '1', 10) || 1));
        const rowspanAttr = Math.min(500, Math.max(1, parseInt((attrs.match(/\browspan\s*=\s*["']?(\d+)/i) || [])[1] || '1', 10) || 1));

        const clsDecls = {};
        clsAttr.split(/\s+/).filter(Boolean).forEach(clsName => {
          const found = classStyleMap.get(clsName);
          if (found) Object.assign(clsDecls, found);
        });
        const inlineDecls = parseStyleDeclarationsMap(styleAttr);
        const merged = { ...clsDecls, ...inlineDecls };
        const bgRaw = extractBackgroundProp(merged) || bgcolorAttr;
        const fgRaw = merged.color || '';
        const weightRaw = merged['font-weight'] || '';
        const hasBoldTag = /<(b|strong)\b/i.test(inner);

        const snap = resolveColorStyleSnapshotFromProps({ bgRaw, fgRaw, weightRaw, tagName, hasBoldTag });
        if (snap.hasVisualStyle) anyStyle = true;
        const cellText = extractHtmlCellTextFallback(inner);
        const styleObj = { fill: snap.fill, font: snap.font };

        for (let dr = 0; dr < rowspanAttr; dr++) {
          const targetR = rIdx + dr;
          if (!occupied[targetR]) occupied[targetR] = [];
          if (!valueRows[targetR]) valueRows[targetR] = [];
          if (!styleRows[targetR]) styleRows[targetR] = [];
          for (let dc = 0; dc < colspanAttr; dc++) {
            const targetC = cCursor + dc;
            occupied[targetR][targetC] = true;
            valueRows[targetR][targetC] = (dr === 0 && dc === 0) ? cellText : '';
            styleRows[targetR][targetC] = styleObj;
          }
        }
        cCursor += colspanAttr;
      });
    });
  }

  if (!valueRows || valueRows.length === 0) return null;
  const maxCols = valueRows.reduce((mx, row) => Math.max(mx, Array.isArray(row) ? row.length : 0), 0);
  if (maxCols === 0) return null;
  for (let r = 0; r < valueRows.length; r++) {
    if (!valueRows[r]) valueRows[r] = [];
    if (!styleRows[r]) styleRows[r] = [];
    for (let c = 0; c < maxCols; c++) {
      if (valueRows[r][c] === undefined) valueRows[r][c] = '';
      if (styleRows[r][c] === undefined) styleRows[r][c] = { fill: null, font: null };
    }
  }

  return {
    values: valueRows,
    styles: anyStyle ? styleRows : null
  };
}

function writeScratchClipboardPayloadToSystem(tsv, html) {
  if (typeof navigator === 'undefined' || !navigator.clipboard) return;
  if (typeof navigator.clipboard.write === 'function' && typeof ClipboardItem !== 'undefined' && typeof Blob !== 'undefined' && html) {
    try {
      const item = new ClipboardItem({
        'text/plain': new Blob([tsv], { type: 'text/plain' }),
        'text/html': new Blob([html], { type: 'text/html' })
      });
      navigator.clipboard.write([item]).catch(() => {
        if (typeof navigator.clipboard.writeText === 'function') {
          navigator.clipboard.writeText(tsv).catch(() => {});
        }
      });
      return;
    } catch (_) {
      // fallback to writeText below
    }
  }
  if (typeof navigator.clipboard.writeText === 'function') {
    navigator.clipboard.writeText(tsv).catch(() => {});
  }
}

function clearScratchSelectionCells(sheetName, range) {
  if (!workbook) return 0;
  const ws = workbook.getWorksheet(sheetName);
  if (!ws) return 0;
  scratchClipboardRangeState = null;
  const edits = [];
  for (let r = range.minRow; r <= range.maxRow; r++) {
    for (let c = range.minCol; c <= range.maxCol; c++) {
      edits.push({ r, c, rawText: '' });
    }
  }
  return applyScratchCellChanges(sheetName, edits, { recordUndo: true });
}

function parseClipboardTSVMatrix(rawText) {
  const normalized = String(rawText ?? '').replace(/\r\n/g, '\n').replace(/\r/g, '\n');
  if (normalized === '') return [];

  const rows = [];
  let currentRow = [];
  let current = '';
  let inQuotes = false;
  let atFieldStart = true;

  for (let i = 0; i < normalized.length; i++) {
    const ch = normalized[i];
    if (inQuotes) {
      if (ch === '"') {
        if (normalized[i + 1] === '"') {
          current += '"';
          i++;
        } else {
          inQuotes = false;
        }
      } else {
        current += ch;
      }
      continue;
    }

    if (atFieldStart && ch === '"') {
      inQuotes = true;
      atFieldStart = false;
      continue;
    }

    if (ch === '\t') {
      currentRow.push(current);
      current = '';
      atFieldStart = true;
      continue;
    }

    if (ch === '\n') {
      currentRow.push(current);
      rows.push(currentRow);
      currentRow = [];
      current = '';
      atFieldStart = true;
      continue;
    }

    current += ch;
    atFieldStart = false;
  }

  if (inQuotes) {
    const lines = normalized.split('\n');
    if (lines.length > 1 && lines[lines.length - 1] === '') {
      lines.pop();
    }
    return lines.map(line => line.split('\t'));
  }

  currentRow.push(current);
  rows.push(currentRow);

  if (rows.length > 1) {
    const last = rows[rows.length - 1];
    if (last.length === 1 && last[0] === '') {
      rows.pop();
    }
  }
  return rows;
}

function transpose2DMatrix(matrix, padValue = '') {
  if (!Array.isArray(matrix) || matrix.length === 0) return [];
  const rowCount = matrix.length;
  const colCount = matrix.reduce((mx, row) => Math.max(mx, Array.isArray(row) ? row.length : 0), 0);
  if (colCount === 0) return [];
  const out = [];
  for (let c = 0; c < colCount; c++) {
    const newRow = [];
    for (let r = 0; r < rowCount; r++) {
      const row = matrix[r];
      const val = (Array.isArray(row) && c < row.length && row[c] !== undefined) ? row[c] : padValue;
      newRow.push(val);
    }
    out.push(newRow);
  }
  return out;
}

function performScratchCopyOrCut(sheetName, range, { isCut = false, writeSystemClipboard = true } = {}) {
  if (!workbook || !sheetName || sheetName === DAILY_TRACKER_WS_TAB) return '';
  const ws = workbook.getWorksheet(sheetName);
  if (!ws) return '';

  const effRange = isCut ? range : resolveEffectiveScratchCopyRange(ws, range);
  const tsv = serializeScratchSelectionToTSV(ws, effRange);
  const html = serializeScratchSelectionToExcelHtml(ws, effRange);
  const styleMatrix = buildScratchStyleMatrix(ws, effRange);
  const rawEditMatrix = buildScratchRawEditMatrix(ws, effRange);
  scratchInternalClipboardText = tsv;
  scratchInternalClipboardState = {
    tsv,
    html,
    styleMatrix,
    rawEditMatrix,
    sourceRange: {
      minRow: effRange.minRow,
      maxRow: effRange.maxRow,
      minCol: effRange.minCol,
      maxCol: effRange.maxCol
    }
  };

  if (writeSystemClipboard) {
    writeScratchClipboardPayloadToSystem(tsv, html);
  }

  if (isCut) {
    scratchClipboardRangeState = null;
    clearScratchSelectionCells(sheetName, range);
    syncScratchFormulaBarFromActiveCell();
    showToast(
      range.cellCount > 1
        ? `Cut ${range.cellCount} cells (${getRangeCoordString(range)}).`
        : `Cut cell ${getRangeCoordString(range)}.`,
      'info'
    );
  } else {
    scratchClipboardRangeState = {
      sheetName,
      minRow: effRange.minRow,
      maxRow: effRange.maxRow,
      minCol: effRange.minCol,
      maxCol: effRange.maxCol
    };
    applyScratchSelectionDOM();
    if (effRange.cellCount > 1 || !writeSystemClipboard) {
      showToast(
        effRange.cellCount > 1
          ? `Copied ${effRange.cellCount} cells (${getRangeCoordString(effRange)}).`
          : `Copied cell ${getRangeCoordString(effRange)}.`,
        'info'
      );
    }
  }
  return tsv;
}

function applyScratchPasteFromText(
  sheetName,
  rawText,
  htmlText = '',
  { clearCopyMarquee = false, transpose = false } = {}
) {
  if (!workbook || !sheetName || sheetName === DAILY_TRACKER_WS_TAB) return 0;
  const hasHtmlTable = Boolean(htmlText && /<table[\s>]/i.test(htmlText));
  if ((rawText === undefined || rawText === null || rawText === '') && !hasHtmlTable) return 0;

  const ws = workbook.getWorksheet(sheetName);
  if (!ws) return 0;

  const hasMatchingInternalClip = Boolean(
    rawText &&
    scratchInternalClipboardState &&
    scratchInternalClipboardState.tsv === rawText
  );
  let matrix = [];
  let styleMatrix = null;
  let rawEditMatrix = null;
  let srcRange = null;

  if (hasMatchingInternalClip) {
    if (Array.isArray(scratchInternalClipboardState.styleMatrix)) {
      styleMatrix = scratchInternalClipboardState.styleMatrix;
    }
    if (Array.isArray(scratchInternalClipboardState.rawEditMatrix)) {
      rawEditMatrix = scratchInternalClipboardState.rawEditMatrix;
      matrix = rawEditMatrix.map(row => row.map(cell => (cell ? cell.rawEdit : '')));
    } else {
      matrix = parseClipboardTSVMatrix(rawText);
    }
    srcRange = scratchInternalClipboardState.sourceRange || null;
  } else if (hasHtmlTable) {
    const parsedHtml = parseClipboardHtmlTableToMatrix(htmlText);
    if (parsedHtml && Array.isArray(parsedHtml.values) && parsedHtml.values.length > 0) {
      matrix = parsedHtml.values;
      styleMatrix = parsedHtml.styles || null;
    } else {
      matrix = parseClipboardTSVMatrix(rawText || '');
    }
  } else {
    matrix = parseClipboardTSVMatrix(rawText || '');
  }

  if (matrix.length === 0) return 0;

  if (transpose) {
    matrix = transpose2DMatrix(matrix, '');
    if (styleMatrix) styleMatrix = transpose2DMatrix(styleMatrix, null);
    if (rawEditMatrix) rawEditMatrix = transpose2DMatrix(rawEditMatrix, null);
  }

  const curRange = getScratchSelectionRange(sheetName);
  const isSingleValueClipboard = matrix.length === 1 && matrix[0].length === 1;
  const edits = [];
  let nextSelection = null;

  if (isSingleValueClipboard && curRange.cellCount > 1) {
    // Excel-style Repeat-Fill: fill entire highlighted multi-cell range with the single copied value
    const singleRawEntry = rawEditMatrix && rawEditMatrix[0] ? rawEditMatrix[0][0] : null;
    const baseVal = singleRawEntry ? singleRawEntry.rawEdit : matrix[0][0];
    const singleIsFormula = Boolean(singleRawEntry && singleRawEntry.isFormula && baseVal.startsWith('='));
    const singleStyle = styleMatrix && styleMatrix[0] ? styleMatrix[0][0] : undefined;
    for (let r = curRange.minRow; r <= curRange.maxRow; r++) {
      for (let c = curRange.minCol; c <= curRange.maxCol; c++) {
        const finalVal = (singleIsFormula && srcRange)
          ? shiftFormulaCellReferences(baseVal, r - srcRange.minRow, c - srcRange.minCol)
          : baseVal;
        const editObj = { r, c, rawText: finalVal };
        if (singleStyle !== undefined) {
          editObj.styleSnap = singleStyle || { fill: null, font: null, numFmt: null };
        }
        edits.push(editObj);
      }
    }
    nextSelection = {
      anchorRow: curRange.anchorRow,
      anchorCol: curRange.anchorCol,
      focusRow: curRange.focusRow,
      focusCol: curRange.focusCol
    };
  } else {
    // Paste 2D block starting from top-left of current selection
    const startRow = curRange.minRow;
    const startCol = curRange.minCol;
    let maxR = startRow;
    let maxC = startCol;

    matrix.forEach((cols, rOffset) => {
      const styleRow = styleMatrix && styleMatrix[rOffset] ? styleMatrix[rOffset] : null;
      const rawRow = rawEditMatrix && rawEditMatrix[rOffset] ? rawEditMatrix[rOffset] : null;
      cols.forEach((cellVal, cOffset) => {
        const r = startRow + rOffset;
        const c = startCol + cOffset;
        if (r > maxR) maxR = r;
        if (c > maxC) maxC = c;
        const rawEntry = rawRow && rawRow[cOffset] ? rawRow[cOffset] : null;
        let finalVal = rawEntry ? rawEntry.rawEdit : cellVal;
        if (rawEntry && rawEntry.isFormula && finalVal.startsWith('=') && srcRange) {
          const origR = transpose ? (srcRange.minRow + cOffset) : (srcRange.minRow + rOffset);
          const origC = transpose ? (srcRange.minCol + rOffset) : (srcRange.minCol + cOffset);
          finalVal = shiftFormulaCellReferences(finalVal, r - origR, c - origC);
        }
        const editObj = { r, c, rawText: finalVal };
        if (styleMatrix) {
          const cellStyle = styleRow && styleRow[cOffset] ? styleRow[cOffset] : null;
          editObj.styleSnap = cellStyle || { fill: null, font: null, numFmt: null };
        }
        edits.push(editObj);
      });
    });

    nextSelection = {
      anchorRow: startRow,
      anchorCol: startCol,
      focusRow: maxR,
      focusCol: maxC
    };
  }

  if (clearCopyMarquee) {
    scratchClipboardRangeState = null;
  }
  const changed = applyScratchCellChanges(sheetName, edits, {
    recordUndo: true,
    selectionAfter: nextSelection
  });

  if (clearCopyMarquee) {
    applyScratchSelectionDOM();
  }

  if (transpose && edits.length > 0) {
    showToast(
      `Pasted & transposed ${edits.length} cell${edits.length === 1 ? '' : 's'} (${matrix.length}×${matrix[0].length}) into "${sheetName}".`,
      'info'
    );
  } else if (edits.length > 1) {
    showToast(`Pasted ${edits.length} cells into "${sheetName}".`, 'info');
  }
  return changed;
}

async function handleScratchContextPaste({ transpose = false } = {}) {
  if (activeWorkspaceTab === DAILY_TRACKER_WS_TAB || !workbook) return;
  const targetSheet = activeWorkspaceTab;
  let clipText = null;
  let clipHtml = '';

  if (navigator.clipboard && typeof navigator.clipboard.read === 'function') {
    try {
      const items = await navigator.clipboard.read();
      for (const item of items) {
        if (!clipText && item.types && item.types.includes('text/plain')) {
          const blob = await item.getType('text/plain');
          clipText = await blob.text();
        }
        if (!clipHtml && item.types && item.types.includes('text/html')) {
          const blob = await item.getType('text/html');
          clipHtml = await blob.text();
        }
      }
    } catch (_) {
      clipText = null;
    }
  }

  if ((clipText === null || clipText === '') && navigator.clipboard && typeof navigator.clipboard.readText === 'function') {
    try {
      clipText = await navigator.clipboard.readText();
    } catch (_) {
      clipText = null;
    }
  }

  if ((clipText === null || clipText === '') && !clipHtml && scratchInternalClipboardText !== '') {
    clipText = scratchInternalClipboardText;
    if (scratchInternalClipboardState && scratchInternalClipboardState.tsv === clipText) {
      clipHtml = scratchInternalClipboardState.html || '';
    }
  }

  const hasHtmlTable = Boolean(clipHtml && /<table[\s>]/i.test(clipHtml));
  if ((clipText !== null && clipText !== '') || hasHtmlTable) {
    applyScratchPasteFromText(targetSheet, clipText || '', clipHtml, { transpose });
    if (scratchGridWrap) scratchGridWrap.focus({ preventScroll: true });
  } else {
    showToast('Allow clipboard read permission when prompted, or press Ctrl+V to paste.', 'info');
  }
}

function transposeScratchSelection(sheetName, range) {
  if (!workbook || !sheetName || sheetName === DAILY_TRACKER_WS_TAB || !range) return 0;
  const ws = workbook.getWorksheet(sheetName);
  if (!ws) return 0;

  const effRange = resolveEffectiveScratchCopyRange(ws, range);
  const rowCount = effRange.maxRow - effRange.minRow + 1;
  const colCount = effRange.maxCol - effRange.minCol + 1;
  if (rowCount * colCount <= 1) {
    showToast('Select 2 or more cells to transpose.', 'info');
    return 0;
  }

  const srcGrid = [];
  for (let r = effRange.minRow; r <= effRange.maxRow; r++) {
    const rowItems = [];
    for (let c = effRange.minCol; c <= effRange.maxCol; c++) {
      const info = extractScratchCellInfo(ws, r, c);
      rowItems.push({
        srcR: r,
        srcC: c,
        rawEdit: info.rawEdit,
        isFormula: info.isFormula,
        styleSnap: captureCellStyleSnapshot(ws, r, c)
      });
    }
    srcGrid.push(rowItems);
  }

  const edits = [];
  const writtenCoords = new Set();

  for (let rOff = 0; rOff < rowCount; rOff++) {
    for (let cOff = 0; cOff < colCount; cOff++) {
      const destR = effRange.minRow + cOff;
      const destC = effRange.minCol + rOff;
      const cellSnap = srcGrid[rOff][cOff];
      let nextRaw = cellSnap.rawEdit;
      if (cellSnap.isFormula && nextRaw.startsWith('=')) {
        nextRaw = shiftFormulaCellReferences(nextRaw, destR - cellSnap.srcR, destC - cellSnap.srcC);
      }
      edits.push({
        r: destR,
        c: destC,
        rawText: nextRaw,
        styleSnap: cellSnap.styleSnap
      });
      writtenCoords.add(`${destR}::${destC}`);
    }
  }

  for (let r = effRange.minRow; r <= effRange.maxRow; r++) {
    for (let c = effRange.minCol; c <= effRange.maxCol; c++) {
      if (!writtenCoords.has(`${r}::${c}`)) {
        edits.push({
          r,
          c,
          rawText: '',
          styleSnap: { fill: null, font: null, numFmt: null }
        });
      }
    }
  }

  scratchClipboardRangeState = null;
  const nextSelection = {
    anchorRow: effRange.minRow,
    anchorCol: effRange.minCol,
    focusRow: effRange.minRow + colCount - 1,
    focusCol: effRange.minCol + rowCount - 1
  };
  const changed = applyScratchCellChanges(sheetName, edits, {
    recordUndo: true,
    selectionAfter: nextSelection
  });
  syncScratchFormulaBarFromActiveCell();
  showToast(
    `Transposed ${rowCount}×${colCount} → ${colCount}×${rowCount} (${getRangeCoordString(normalizeScratchRangeObj(nextSelection))}).`,
    'success'
  );
  return changed;
}

// --- EXCEL DEFAULT FILL PALETTE & RIGHT-CLICK CONTEXT MENU ---
function syncScratchFillIndicatorUI() {
  const preset = EXCEL_DEFAULT_FILL_PALETTE.find(p => p.id === activeScratchFillPresetId) || EXCEL_DEFAULT_FILL_PALETTE[5];
  if (scratchFillActiveBar && preset) {
    scratchFillActiveBar.style.background = preset.hex;
  }
  if (btnScratchApplyFill && preset) {
    btnScratchApplyFill.title = `Apply "${preset.label}" fill to selected cell(s)`;
  }
  if (scratchFillDropdown) {
    scratchFillDropdown.querySelectorAll('.scratch-fill-swatch').forEach(btn => {
      btn.classList.toggle('active-swatch', btn.dataset.fillId === activeScratchFillPresetId);
    });
  }
}

function initScratchFillPaletteUI() {
  const makeSwatchBtn = (preset, clsName) => {
    const btn = document.createElement('button');
    btn.type = 'button';
    btn.className = clsName;
    btn.dataset.fillId = preset.id;
    btn.style.background = preset.hex;
    btn.title = preset.label;
    btn.setAttribute('aria-label', preset.label);
    return btn;
  };

  if (scratchFillPastelGrid) {
    scratchFillPastelGrid.innerHTML = '';
    EXCEL_DEFAULT_FILL_PALETTE.filter(p => p.group === 'pastel').forEach(preset => {
      scratchFillPastelGrid.appendChild(makeSwatchBtn(preset, 'scratch-fill-swatch'));
    });
  }
  if (scratchFillStandardGrid) {
    scratchFillStandardGrid.innerHTML = '';
    EXCEL_DEFAULT_FILL_PALETTE.filter(p => p.group === 'standard').forEach(preset => {
      scratchFillStandardGrid.appendChild(makeSwatchBtn(preset, 'scratch-fill-swatch'));
    });
  }
  if (scratchCtxQuickSwatches) {
    scratchCtxQuickSwatches.innerHTML = '';
    const noneBtn = document.createElement('button');
    noneBtn.type = 'button';
    noneBtn.className = 'scratch-ctx-swatch-btn scratch-swatch-none';
    noneBtn.dataset.fillId = 'none';
    noneBtn.title = 'No Fill (Clear Color)';
    noneBtn.setAttribute('aria-label', 'No Fill');
    scratchCtxQuickSwatches.appendChild(noneBtn);

    EXCEL_DEFAULT_FILL_PALETTE.filter(p => p.quick).slice(0, 7).forEach(preset => {
      scratchCtxQuickSwatches.appendChild(makeSwatchBtn(preset, 'scratch-ctx-swatch-btn'));
    });
  }

  syncScratchFillIndicatorUI();
  syncScratchStripedRowsUI();
  syncScratchAutoIncrementUI();
}

function openScratchContextMenuAt(clientX, clientY) {
  if (!scratchContextMenu || activeWorkspaceTab === DAILY_TRACKER_WS_TAB || !workbook) return;
  const range = getScratchSelectionRange(activeWorkspaceTab);
  if (scratchCtxRangeLabel) {
    scratchCtxRangeLabel.textContent = `Selection • ${getRangeCoordString(range)}`;
  }
  const transposeSelBtn = scratchContextMenu.querySelector('[data-ctx-action="transpose-selection"]');
  if (transposeSelBtn) {
    transposeSelBtn.disabled = range.cellCount <= 1;
  }

  scratchContextMenu.style.display = 'block';
  scratchContextMenu.style.left = `${clientX}px`;
  scratchContextMenu.style.top = `${clientY}px`;

  const menuRect = scratchContextMenu.getBoundingClientRect();
  const maxLeft = Math.max(8, window.innerWidth - menuRect.width - 8);
  const maxTop = Math.max(8, window.innerHeight - menuRect.height - 8);
  scratchContextMenu.style.left = `${Math.min(clientX, maxLeft)}px`;
  scratchContextMenu.style.top = `${Math.min(clientY, maxTop)}px`;
}

// --- DATA TO COLUMNS (TEXT TO COLUMNS) WIZARD & ENGINE ---
function splitScratchLineByDelimiters(rawText, config) {
  const str = String(rawText ?? '').replace(/\r\n/g, '\n').replace(/\r/g, '\n');
  const delims = Array.isArray(config.delims)
    ? config.delims.filter(d => typeof d === 'string' && d.length > 0).sort((a, b) => b.length - a.length)
    : [];
  if (delims.length === 0) {
    return [config.trim ? str.trim() : str];
  }

  const quoteCount = (str.match(/"/g) || []).length;
  const allowQuotedFields = !delims.includes('\n') && quoteCount >= 2 && quoteCount % 2 === 0;
  const tokens = [];
  let current = '';
  let inQuotes = false;
  let lastWasDelim = false;
  let i = 0;

  while (i < str.length) {
    const ch = str[i];
    if (allowQuotedFields && ch === '"') {
      if (inQuotes && str[i + 1] === '"') {
        current += '"';
        i += 2;
        continue;
      }
      if (!inQuotes && current.trim() === '') {
        inQuotes = true;
        current = '';
        i++;
        continue;
      }
      if (inQuotes) {
        inQuotes = false;
        i++;
        continue;
      }
    }

    if (!inQuotes) {
      let matchedDelim = null;
      for (let dIdx = 0; dIdx < delims.length; dIdx++) {
        const d = delims[dIdx];
        if (str.startsWith(d, i)) {
          matchedDelim = d;
          break;
        }
      }
      if (matchedDelim !== null) {
        if (!(config.mergeConsecutive && lastWasDelim)) {
          tokens.push(config.trim ? current.trim() : current);
          current = '';
        }
        lastWasDelim = true;
        i += matchedDelim.length;
        continue;
      }
    }

    lastWasDelim = false;
    current += ch;
    i++;
  }

  tokens.push(config.trim ? current.trim() : current);
  return tokens;
}

function getActiveTextToColumnsConfig() {
  const delims = [];
  if (chkTtcTab && chkTtcTab.checked) delims.push('\t');
  if (chkTtcComma && chkTtcComma.checked) delims.push(',');
  if (chkTtcSemicolon && chkTtcSemicolon.checked) delims.push(';');
  if (chkTtcPipe && chkTtcPipe.checked) delims.push('|');
  if (chkTtcHyphen && chkTtcHyphen.checked) delims.push('-');
  if (chkTtcSpace && chkTtcSpace.checked) delims.push(' ');
  if (chkTtcNewline && chkTtcNewline.checked) delims.push('\n');
  if (chkTtcOther && chkTtcOther.checked && ttcOtherInput && ttcOtherInput.value.length > 0) {
    delims.push(ttcOtherInput.value);
  }
  return {
    delims,
    trim: Boolean(!chkTtcTrim || chkTtcTrim.checked),
    mergeConsecutive: Boolean(chkTtcConsecutive && chkTtcConsecutive.checked),
    direction: (radTtcDown && radTtcDown.checked) ? 'down' : 'across'
  };
}

function autoDetectDelimitersForSamples(samples) {
  const sampleSlice = Array.isArray(samples) ? samples.slice(0, 15) : [];
  const hasNewline = sampleSlice.some(s => /[\r\n]/.test(String(s || '')));
  const hasBlankLines = hasNewline && sampleSlice.some(s => /\r?\n[ \t]*\r?\n/.test(String(s || '')));
  const joined = sampleSlice.join('\n');
  const hasTab = joined.includes('\t');
  const hasComma = joined.includes(',');
  const hasPipe = joined.includes('|');
  const hasSemi = joined.includes(';');
  const hasHyphen = joined.includes(' - ') || /[A-Za-z0-9]-[A-Za-z0-9]/.test(joined);

  if (chkTtcTab) chkTtcTab.checked = false;
  if (chkTtcComma) chkTtcComma.checked = false;
  if (chkTtcSemicolon) chkTtcSemicolon.checked = false;
  if (chkTtcPipe) chkTtcPipe.checked = false;
  if (chkTtcHyphen) chkTtcHyphen.checked = false;
  if (chkTtcSpace) chkTtcSpace.checked = false;
  if (chkTtcNewline) chkTtcNewline.checked = false;
  if (chkTtcOther) chkTtcOther.checked = false;
  if (chkTtcConsecutive) chkTtcConsecutive.checked = false;
  if (radTtcAcross) radTtcAcross.checked = true;
  if (radTtcDown) radTtcDown.checked = false;

  if (hasNewline) {
    if (chkTtcNewline) chkTtcNewline.checked = true;
    if (hasBlankLines && chkTtcConsecutive) chkTtcConsecutive.checked = true;
    if (sampleSlice.length === 1 && radTtcDown) {
      radTtcDown.checked = true;
      if (radTtcAcross) radTtcAcross.checked = false;
    }
  } else if (hasTab) {
    if (chkTtcTab) chkTtcTab.checked = true;
  } else if (hasComma) {
    if (chkTtcComma) chkTtcComma.checked = true;
  } else if (hasPipe) {
    if (chkTtcPipe) chkTtcPipe.checked = true;
  } else if (hasSemi) {
    if (chkTtcSemicolon) chkTtcSemicolon.checked = true;
  } else if (hasHyphen) {
    if (chkTtcHyphen) chkTtcHyphen.checked = true;
  } else if (joined.includes(' ')) {
    if (chkTtcSpace) chkTtcSpace.checked = true;
    if (chkTtcConsecutive) chkTtcConsecutive.checked = true;
  } else {
    if (chkTtcComma) chkTtcComma.checked = true;
  }
}

function closeTextToColumnsModal() {
  activeTextToColumnsState = null;
  if (textToColumnsModal) {
    textToColumnsModal.classList.remove('open');
  }
}

function renderTextToColumnsPreview() {
  if (!activeTextToColumnsState || !workbook || !ttcPreviewBox) return;
  const { sheetName, col, sourceRows } = activeTextToColumnsState;
  const ws = workbook.getWorksheet(sheetName);
  if (!ws) return;

  const config = getActiveTextToColumnsConfig();
  const isDown = config.direction === 'down';
  const colLetter = colIndexToExcelLetter(col);

  if (ttcSubtitle) {
    ttcSubtitle.textContent = isDown
      ? `Splitting ${sourceRows.length} non-empty cell${sourceRows.length === 1 ? '' : 's'} from Column ${colLetter} downward into consecutive rows`
      : `Splitting ${sourceRows.length} non-empty cell${sourceRows.length === 1 ? '' : 's'} from Column ${colLetter} into adjacent columns to the right`;
  }
  if (btnConfirmTextToColumns) {
    btnConfirmTextToColumns.textContent = isDown ? 'Split into Rows' : 'Split into Columns';
  }

  if (config.delims.length === 0) {
    ttcPreviewBox.innerHTML = `<div class="ttc-preview-empty">Select at least one delimiter above to preview split ${isDown ? 'rows' : 'columns'}.</div>`;
    if (ttcOverwriteBanner) ttcOverwriteBanner.style.display = 'none';
    if (btnConfirmTextToColumns) btnConfirmTextToColumns.disabled = true;
    return;
  }

  if (btnConfirmTextToColumns) {
    btnConfirmTextToColumns.disabled = false;
  }

  if (isDown) {
    const flatParts = [];
    sourceRows.forEach(item => {
      splitScratchLineByDelimiters(item.raw, config).forEach(part => {
        flatParts.push(part);
      });
    });

    const startRow = sourceRows[0].r;
    const sourceRowSet = new Set(sourceRows.map(s => s.r));
    let overwriteCount = 0;
    let firstOverwrittenRow = 0;
    let lastOverwrittenRow = 0;

    for (let i = 0; i < flatParts.length; i++) {
      const targetR = startRow + i;
      if (sourceRowSet.has(targetR)) continue;
      const existingDisp = extractScratchCellInfo(ws, targetR, col).display.trim();
      if (existingDisp !== '') {
        overwriteCount++;
        if (!firstOverwrittenRow) firstOverwrittenRow = targetR;
        lastOverwrittenRow = targetR;
      }
    }

    if (ttcOverwriteBanner) {
      if (overwriteCount > 0) {
        const rowSpanLabel = firstOverwrittenRow === lastOverwrittenRow
          ? `Row ${firstOverwrittenRow}`
          : `Rows ${firstOverwrittenRow}–${lastOverwrittenRow}`;
        ttcOverwriteBanner.textContent = `Heads up: Splitting down will overwrite ${overwriteCount} existing non-empty cell${overwriteCount === 1 ? '' : 's'} in Column ${colLetter} (${rowSpanLabel}). You can undo anytime with Ctrl+Z.`;
        ttcOverwriteBanner.style.display = 'block';
      } else {
        ttcOverwriteBanner.style.display = 'none';
      }
    }

    if (ttcPreviewTitle) {
      ttcPreviewTitle.textContent = `Data Preview (${sourceRows.length} cell${sourceRows.length === 1 ? '' : 's'} → ${flatParts.length} row${flatParts.length === 1 ? '' : 's'} in Col ${colLetter})`;
    }

    const previewLimit = Math.min(12, flatParts.length);
    const headHtml = `<tr><th class="ttc-row-num-col">#</th><th>Col ${colLetter}</th></tr>`;
    const bodyRows = [];
    for (let i = 0; i < previewLimit; i++) {
      const val = flatParts[i] !== undefined ? flatParts[i] : '';
      bodyRows.push(`<tr><td class="ttc-row-num-col">${startRow + i}</td><td title="${escapeHtml(val)}">${escapeHtml(val)}</td></tr>`);
    }
    if (flatParts.length > previewLimit) {
      bodyRows.push(`<tr><td class="ttc-row-num-col">…</td><td style="color:var(--text-muted);">+${flatParts.length - previewLimit} more row(s)</td></tr>`);
    }
    ttcPreviewBox.innerHTML = `<table class="ttc-preview-table"><thead>${headHtml}</thead><tbody>${bodyRows.join('')}</tbody></table>`;
    return;
  }

  let maxCols = 1;
  const splitRows = sourceRows.map(item => {
    const parts = splitScratchLineByDelimiters(item.raw, config);
    if (parts.length > maxCols) maxCols = parts.length;
    return { r: item.r, parts };
  });

  // Check if splitting into adjacent columns (col + 1 .. col + maxCols - 1) overwrites existing non-empty cells
  let overwriteCount = 0;
  if (maxCols > 1) {
    splitRows.forEach(sr => {
      for (let cOff = 1; cOff < sr.parts.length; cOff++) {
        const existingDisp = extractScratchCellInfo(ws, sr.r, col + cOff).display.trim();
        if (existingDisp !== '') {
          overwriteCount++;
        }
      }
    });
  }

  if (ttcOverwriteBanner) {
    if (overwriteCount > 0) {
      const firstTargetLetter = colIndexToExcelLetter(col + 1);
      const lastTargetLetter = colIndexToExcelLetter(col + maxCols - 1);
      const colSpanLabel = firstTargetLetter === lastTargetLetter
        ? `Column ${firstTargetLetter}`
        : `Columns ${firstTargetLetter}–${lastTargetLetter}`;
      ttcOverwriteBanner.textContent = `Heads up: Splitting will overwrite ${overwriteCount} existing non-empty cell${overwriteCount === 1 ? '' : 's'} in ${colSpanLabel}. You can undo anytime with Ctrl+Z.`;
      ttcOverwriteBanner.style.display = 'block';
    } else {
      ttcOverwriteBanner.style.display = 'none';
    }
  }

  if (ttcPreviewTitle) {
    ttcPreviewTitle.textContent = `Data Preview (${sourceRows.length} non-empty row${sourceRows.length === 1 ? '' : 's'} → ${maxCols} column${maxCols === 1 ? '' : 's'})`;
  }

  const previewLimit = Math.min(8, splitRows.length);
  const displayCols = Math.min(maxCols, 12);
  let headHtml = `<tr><th class="ttc-row-num-col">#</th>`;
  for (let cOff = 0; cOff < displayCols; cOff++) {
    headHtml += `<th>Col ${colIndexToExcelLetter(col + cOff)}</th>`;
  }
  headHtml += `</tr>`;

  const bodyRows = [];
  for (let i = 0; i < previewLimit; i++) {
    const sr = splitRows[i];
    let cells = `<td class="ttc-row-num-col">${sr.r}</td>`;
    for (let cOff = 0; cOff < displayCols; cOff++) {
      const val = sr.parts[cOff] !== undefined ? sr.parts[cOff] : '';
      cells += `<td title="${escapeHtml(val)}">${escapeHtml(val)}</td>`;
    }
    bodyRows.push(`<tr>${cells}</tr>`);
  }

  ttcPreviewBox.innerHTML = `<table class="ttc-preview-table"><thead>${headHtml}</thead><tbody>${bodyRows.join('')}</tbody></table>`;
}

function openTextToColumnsModal() {
  closeScratchContextMenu();
  closeScratchFillDropdown();
  if (activeWorkspaceTab === DAILY_TRACKER_WS_TAB || !workbook) return;
  commitActiveScratchInlineEdit();

  const sheetName = activeWorkspaceTab;
  const ws = workbook.getWorksheet(sheetName);
  if (!ws) return;

  const dims = getScratchSheetGridDimensions(ws, sheetName);
  const range = getScratchSelectionRange(sheetName);
  const col = range.minCol;

  let minRow = range.minRow;
  let maxRow = range.maxRow;
  if (range.minRow === 1 && range.maxRow >= dims.rows) {
    maxRow = Math.max(1, dims.dataMaxRow);
  }

  let sourceRows = [];
  for (let r = minRow; r <= maxRow; r++) {
    const disp = extractScratchCellInfo(ws, r, col).display;
    if (disp.trim() !== '') {
      sourceRows.push({ r, raw: disp });
    }
  }

  // Convenience: if user had only 1 empty cell selected, check if that column has data rows we can target
  if (sourceRows.length === 0 && range.cellCount === 1 && dims.dataMaxRow >= 1) {
    minRow = 1;
    maxRow = dims.dataMaxRow;
    for (let r = minRow; r <= maxRow; r++) {
      const disp = extractScratchCellInfo(ws, r, col).display;
      if (disp.trim() !== '') {
        sourceRows.push({ r, raw: disp });
      }
    }
  }

  if (sourceRows.length === 0) {
    showModalAlert({
      title: 'No Text to Split',
      badge: `Column ${colIndexToExcelLetter(col)}`,
      message: `Select one or more non-empty cells (or click Column ${colIndexToExcelLetter(col)} header) containing delimited text to split across columns.`
    });
    return;
  }

  activeTextToColumnsState = {
    sheetName,
    col,
    minRow: sourceRows[0].r,
    maxRow: sourceRows[sourceRows.length - 1].r,
    sourceRows
  };

  const coordBadge = sourceRows.length === 1
    ? getCellCoordString(sourceRows[0].r, col)
    : `${getCellCoordString(sourceRows[0].r, col)}:${getCellCoordString(sourceRows[sourceRows.length - 1].r, col)}`;

  if (ttcSourceBadge) {
    ttcSourceBadge.textContent = `Source • ${coordBadge}`;
  }
  if (ttcSubtitle) {
    ttcSubtitle.textContent = `Splitting ${sourceRows.length} non-empty cell${sourceRows.length === 1 ? '' : 's'} from Column ${colIndexToExcelLetter(col)} into adjacent columns to the right`;
  }

  autoDetectDelimitersForSamples(sourceRows.map(s => s.raw));
  renderTextToColumnsPreview();

  if (textToColumnsModal) {
    textToColumnsModal.classList.add('open');
  }
}

function executeTextToColumnsSplit() {
  if (!activeTextToColumnsState || !workbook) return;
  const { sheetName, col, minRow, maxRow, sourceRows } = activeTextToColumnsState;
  const config = getActiveTextToColumnsConfig();
  if (config.delims.length === 0) return;

  if (config.direction === 'down') {
    const flatParts = [];
    sourceRows.forEach(item => {
      splitScratchLineByDelimiters(item.raw, config).forEach(part => {
        flatParts.push(part);
      });
    });

    const startRow = sourceRows[0].r;
    const edits = flatParts.map((val, idx) => ({
      r: startRow + idx,
      c: col,
      rawText: val
    }));

    const endWrittenRow = startRow + flatParts.length - 1;
    sourceRows.forEach(item => {
      if (item.r > endWrittenRow) {
        edits.push({ r: item.r, c: col, rawText: '' });
      }
    });

    const endRow = startRow + Math.max(1, flatParts.length) - 1;
    const nextSelection = {
      anchorRow: startRow,
      anchorCol: col,
      focusRow: endRow,
      focusCol: col
    };

    closeTextToColumnsModal();
    applyScratchCellChanges(sheetName, edits, {
      recordUndo: true,
      selectionAfter: nextSelection
    });

    showToast(
      `Split ${sourceRows.length} cell${sourceRows.length === 1 ? '' : 's'} down into ${flatParts.length} row${flatParts.length === 1 ? '' : 's'} (${getRangeCoordString(normalizeScratchRangeObj(nextSelection))}).`,
      'success'
    );
    if (scratchGridWrap) {
      scratchGridWrap.focus({ preventScroll: true });
    }
    return;
  }

  const edits = [];
  let maxSplitCols = 1;

  sourceRows.forEach(item => {
    const parts = splitScratchLineByDelimiters(item.raw, config);
    if (parts.length > maxSplitCols) {
      maxSplitCols = parts.length;
    }
    parts.forEach((val, offset) => {
      edits.push({
        r: item.r,
        c: col + offset,
        rawText: val
      });
    });
  });

  const endCol = col + Math.max(1, maxSplitCols) - 1;
  const nextSelection = {
    anchorRow: minRow,
    anchorCol: col,
    focusRow: maxRow,
    focusCol: endCol
  };

  closeTextToColumnsModal();
  applyScratchCellChanges(sheetName, edits, {
    recordUndo: true,
    selectionAfter: nextSelection
  });

  showToast(
    `Split ${sourceRows.length} row${sourceRows.length === 1 ? '' : 's'} across ${maxSplitCols} column${maxSplitCols === 1 ? '' : 's'} (${getRangeCoordString(normalizeScratchRangeObj(nextSelection))}).`,
    'success'
  );
  if (scratchGridWrap) {
    scratchGridWrap.focus({ preventScroll: true });
  }
}

// --- AUTO-FILTER SIDE LISTS (=FILTER ARRAY BUILDER) ---
function closeScratchAutoFilterModal() {
  activeScratchAutoFilterState = null;
  if (scratchAutoFilterModal) {
    scratchAutoFilterModal.classList.remove('open');
  }
}

function buildAutoFilterPlanFromUI() {
  if (!activeScratchAutoFilterState || !workbook) return null;
  const { sheetName, headers } = activeScratchAutoFilterState;
  const ws = workbook.getWorksheet(sheetName);
  if (!ws) return null;

  const presetMode = afPresetModeSelect ? afPresetModeSelect.value : 'three_bucket';
  let retStartCol = parseInt(afReturnStartCol ? afReturnStartCol.value : '1', 10) || 1;
  let retEndCol = parseInt(afReturnEndCol ? afReturnEndCol.value : '1', 10) || 1;
  if (retEndCol < retStartCol) {
    const tmp = retStartCol;
    retStartCol = retEndCol;
    retEndCol = tmp;
  }
  const critCol = parseInt(afCriteriaCol ? afCriteriaCol.value : '1', 10) || 1;
  const destStartCol = parseInt(afDestStartCol ? afDestStartCol.value : '6', 10) || 6;
  const startRow = Math.max(2, parseInt(afStartRowInput ? afStartRowInput.value : '2', 10) || 2);
  const maxRow = Math.max(startRow, Math.min(5000, parseInt(afMaxRowInput ? afMaxRowInput.value : '200', 10) || 200));

  const retStartLetter = colIndexToExcelLetter(retStartCol);
  const retEndLetter = colIndexToExcelLetter(retEndCol);
  const critLetter = colIndexToExcelLetter(critCol);
  const retRange = `${retStartLetter}${startRow}:${retEndLetter}${maxRow}`;
  const critRange = `${critLetter}${startRow}:${critLetter}${maxRow}`;
  const nonEmptyGuard = `(${retStartLetter}${startRow}:${retStartLetter}${maxRow}<>"")`;

  const getHeaderLabel = (c) => {
    const found = headers.find(h => h.col === c);
    return (found && found.rawLabel) ? found.rawLabel : `Col ${colIndexToExcelLetter(c)}`;
  };
  const baseHeaders = [];
  for (let c = retStartCol; c <= retEndCol; c++) {
    baseHeaders.push(getHeaderLabel(c));
  }
  const critHeader = getHeaderLabel(critCol);

  const buckets = [];
  let cursorCol = destStartCol;

  if (presetMode === 'three_bucket') {
    const keyword = (afYellowKeywordInput ? afYellowKeywordInput.value : 'Duplicate').trim() || 'Duplicate';
    const kwEsc = keyword.replace(/"/g, '""');
    const includeCritInGreen = Boolean(!afIncludeCriteriaInGreen || afIncludeCriteriaInGreen.checked);

    // 1. Red Bucket: Empty Criteria
    buckets.push({
      key: 'red',
      title: `Missing ${critHeader} (Red)`,
      badgeClass: 'af-red',
      fillPresetId: 'excel_bad',
      startCol: cursorCol,
      headers: baseHeaders.map((h, i) => (i === 0 ? `${h} — No ${critHeader}` : h)),
      formulas: [
        {
          colOffset: 0,
          formula: `=FILTER(${retRange}, ${nonEmptyGuard}*(${critRange}=""), "")`
        }
      ]
    });
    cursorCol += baseHeaders.length + 1;

    // 2. Yellow Bucket: Keyword Match (e.g. "Duplicate")
    buckets.push({
      key: 'yellow',
      title: `${keyword} ${critHeader} (Yellow)`,
      badgeClass: 'af-yellow',
      fillPresetId: 'excel_neutral',
      startCol: cursorCol,
      headers: baseHeaders.map((h, i) => (i === 0 ? `${h} — ${keyword}` : h)),
      formulas: [
        {
          colOffset: 0,
          formula: `=FILTER(${retRange}, ${critRange}="${kwEsc}", "")`
        }
      ]
    });
    cursorCol += baseHeaders.length + 1;

    // 3. Green Bucket: Occupied Criteria & Not Keyword
    const greenCond = `(${critRange}<>"")*(${critRange}<>"${kwEsc}")`;
    const critAlreadyIncluded = critCol >= retStartCol && critCol <= retEndCol;
    const greenHeaders = [...baseHeaders];
    const greenFormulas = [];

    if (includeCritInGreen && !critAlreadyIncluded && critCol === retEndCol + 1) {
      const greenRange = `${retStartLetter}${startRow}:${critLetter}${maxRow}`;
      greenHeaders.push(critHeader);
      greenFormulas.push({
        colOffset: 0,
        formula: `=FILTER(${greenRange}, ${greenCond}, "")`
      });
    } else if (includeCritInGreen && !critAlreadyIncluded) {
      greenFormulas.push({
        colOffset: 0,
        formula: `=FILTER(${retRange}, ${greenCond}, "")`
      });
      greenFormulas.push({
        colOffset: baseHeaders.length,
        formula: `=FILTER(${critRange}, ${greenCond}, "")`
      });
      greenHeaders.push(critHeader);
    } else {
      greenFormulas.push({
        colOffset: 0,
        formula: `=FILTER(${retRange}, ${greenCond}, "")`
      });
    }

    buckets.push({
      key: 'green',
      title: `${critHeader} Available (Green)`,
      badgeClass: 'af-green',
      fillPresetId: 'excel_good',
      startCol: cursorCol,
      headers: greenHeaders.map((h, i) => (i === 0 ? `${h} — Available` : h)),
      formulas: greenFormulas
    });
  } else {
    const op = afCustomOperatorSelect ? afCustomOperatorSelect.value : 'isempty';
    const val = (afCustomValueInput ? afCustomValueInput.value : '').trim();
    const valEsc = val.replace(/"/g, '""');
    const colorId = afCustomColorSelect ? afCustomColorSelect.value : 'none';

    let condExpr = `${nonEmptyGuard}*(${critRange}="")`;
    let suffix = 'Filtered';
    if (op === 'isempty') {
      condExpr = `${nonEmptyGuard}*(${critRange}="")`;
      suffix = `Missing ${critHeader}`;
    } else if (op === 'isnotempty') {
      condExpr = `${critRange}<>""`;
      suffix = `Has ${critHeader}`;
    } else if (op === 'equals') {
      condExpr = `${critRange}="${valEsc}"`;
      suffix = `${critHeader} = ${val}`;
    } else if (op === 'notequals') {
      condExpr = `${nonEmptyGuard}*(${critRange}<>"${valEsc}")`;
      suffix = `${critHeader} <> ${val}`;
    } else if (op === 'contains') {
      condExpr = `ISNUMBER(SEARCH("${valEsc}", ${critRange}))`;
      suffix = `${critHeader} ~ "${val}"`;
    }

    buckets.push({
      key: 'custom',
      title: `Filtered (${suffix})`,
      badgeClass: colorId === 'excel_bad' ? 'af-red' : (colorId === 'excel_neutral' ? 'af-yellow' : (colorId === 'excel_good' ? 'af-green' : '')),
      fillPresetId: colorId,
      startCol: cursorCol,
      headers: baseHeaders.map((h, i) => (i === 0 ? `${h} — ${suffix}` : h)),
      formulas: [
        {
          colOffset: 0,
          formula: `=FILTER(${retRange}, ${condExpr}, "")`
        }
      ]
    });
  }

  return { sheetName, startRow, maxRow, buckets };
}

function renderScratchAutoFilterPreview() {
  if (!activeScratchAutoFilterState || !afFormulaPreviewBox) return;
  const presetMode = afPresetModeSelect ? afPresetModeSelect.value : 'three_bucket';
  if (afThreeBucketWrap) afThreeBucketWrap.style.display = presetMode === 'three_bucket' ? 'block' : 'none';
  if (afSingleCustomWrap) afSingleCustomWrap.style.display = presetMode === 'single_custom' ? 'block' : 'none';

  const plan = buildAutoFilterPlanFromUI();
  if (!plan) return;

  afFormulaPreviewBox.innerHTML = plan.buckets.map(b => {
    const anchorCoord = getCellCoordString(plan.startRow, b.startCol);
    const colSpan = `${colIndexToExcelLetter(b.startCol)}:${colIndexToExcelLetter(b.startCol + b.headers.length - 1)}`;
    const formulasHtml = b.formulas.map(f => `<code>${escapeHtml(f.formula)}</code>`).join(' ');
    return `
      <div class="af-preview-row">
        <span class="af-legend-chip ${b.badgeClass}">${escapeHtml(b.title)} (${escapeHtml(colSpan)} • @${escapeHtml(anchorCoord)})</span>
        ${formulasHtml}
      </div>
    `;
  }).join('');
}

function openScratchAutoFilterModal() {
  closeScratchContextMenu();
  closeScratchFillDropdown();
  if (activeWorkspaceTab === DAILY_TRACKER_WS_TAB || !workbook) return;
  commitActiveScratchInlineEdit();

  const sheetName = activeWorkspaceTab;
  const ws = workbook.getWorksheet(sheetName);
  if (!ws) return;
  const dims = getScratchSheetGridDimensions(ws, sheetName);
  const scanCols = Math.max(dims.dataMaxCol, 10);
  const headers = [];
  for (let c = 1; c <= scanCols; c++) {
    const rawLabel = extractScratchCellInfo(ws, 1, c, { skipSpillLookup: true }).display.trim();
    const letter = colIndexToExcelLetter(c);
    headers.push({
      col: c,
      letter,
      rawLabel,
      label: rawLabel ? `Col ${letter} — ${rawLabel}` : `Col ${letter}`
    });
  }

  activeScratchAutoFilterState = { sheetName, headers };
  if (afSheetBadge) afSheetBadge.textContent = sheetName;

  const optsHtml = headers.map(h => `<option value="${h.col}">${escapeHtml(h.label)}</option>`).join('');
  [afReturnStartCol, afReturnEndCol, afCriteriaCol].forEach(sel => {
    if (sel) sel.innerHTML = optsHtml;
  });

  if (afDestStartCol) {
    const destHeaders = [];
    const maxDest = Math.max(scanCols + 8, 20);
    for (let c = 1; c <= maxDest; c++) {
      const found = headers.find(h => h.col === c);
      const lbl = found && found.rawLabel ? `Col ${found.letter} (${found.rawLabel})` : `Col ${colIndexToExcelLetter(c)}`;
      destHeaders.push(`<option value="${c}">${escapeHtml(lbl)}</option>`);
    }
    afDestStartCol.innerHTML = destHeaders.join('');
  }

  // Smart column auto-detection
  let detectedCritCol = Math.max(2, Math.min(dims.dataMaxCol || 4, scanCols));
  const critMatch = headers.find(h => /track|carrier|mail|status|title\s*#/i.test(h.rawLabel));
  if (critMatch) {
    detectedCritCol = critMatch.col;
  }
  const retEndDefault = detectedCritCol > 1 ? detectedCritCol - 1 : 1;
  const destDefault = Math.max(dims.dataMaxCol + 2, detectedCritCol + 2, 6);

  if (afReturnStartCol) afReturnStartCol.value = '1';
  if (afReturnEndCol) afReturnEndCol.value = String(retEndDefault);
  if (afCriteriaCol) afCriteriaCol.value = String(detectedCritCol);
  if (afDestStartCol) afDestStartCol.value = String(destDefault);
  if (afStartRowInput) afStartRowInput.value = '2';
  if (afMaxRowInput) afMaxRowInput.value = String(Math.max(200, dims.dataMaxRow + 50));

  renderScratchAutoFilterPreview();
  if (scratchAutoFilterModal) {
    scratchAutoFilterModal.classList.add('open');
  }
}

function executeScratchAutoFilterSetup() {
  const plan = buildAutoFilterPlanFromUI();
  if (!plan || !workbook) return;
  const ws = workbook.getWorksheet(plan.sheetName);
  if (!ws) return;

  const dims = getScratchSheetGridDimensions(ws, plan.sheetName);
  const headerRowNum = Math.max(1, plan.startRow - 1);
  const tintEndRow = Math.max(30, Math.min(plan.maxRow, Math.max(dims.dataMaxRow + 15, 35)));
  const edits = [];
  let maxEndCol = plan.buckets[0].startCol;

  plan.buckets.forEach(bucket => {
    const colCount = bucket.headers.length;
    const endCol = bucket.startCol + colCount - 1;
    if (endCol > maxEndCol) maxEndCol = endCol;

    // 1. Headers on row above startRow + fill tint down tintEndRow
    bucket.headers.forEach((hdrText, cOff) => {
      const c = bucket.startCol + cOff;
      const baseSnap = captureCellStyleSnapshot(ws, headerRowNum, c);
      let hdrSnap = baseSnap;
      if (bucket.fillPresetId && bucket.fillPresetId !== 'none') {
        hdrSnap = buildStyleSnapshotForFillPreset(baseSnap, bucket.fillPresetId);
      }
      hdrSnap = {
        fill: hdrSnap.fill,
        font: { ...(hdrSnap.font || {}), bold: true },
        numFmt: hdrSnap.numFmt || null
      };
      edits.push({ r: headerRowNum, c, rawText: hdrText, styleSnap: hdrSnap });

      // Ensure formula / data rows start clean (empty values except formula anchor) & tinted
      for (let r = plan.startRow; r <= tintEndRow; r++) {
        const cellBefore = captureCellStyleSnapshot(ws, r, c);
        let cellSnap = cellBefore;
        if (bucket.fillPresetId && bucket.fillPresetId !== 'none') {
          cellSnap = buildStyleSnapshotForFillPreset(cellBefore, bucket.fillPresetId);
        }
        const formulaSpec = (r === plan.startRow)
          ? bucket.formulas.find(f => f.colOffset === cOff)
          : null;
        edits.push({
          r,
          c,
          rawText: formulaSpec ? formulaSpec.formula : '',
          styleSnap: cellSnap
        });
      }
    });
  });

  const firstBucket = plan.buckets[0];
  const nextSelection = {
    anchorRow: plan.startRow,
    anchorCol: firstBucket.startCol,
    focusRow: plan.startRow,
    focusCol: maxEndCol
  };

  closeScratchAutoFilterModal();
  applyScratchCellChanges(plan.sheetName, edits, {
    recordUndo: true,
    selectionAfter: nextSelection
  });

  showToast(
    `Created ${plan.buckets.length} auto-filtered side list${plan.buckets.length === 1 ? '' : 's'} using live =FILTER formulas.`,
    'success'
  );
  if (scratchGridWrap) {
    scratchGridWrap.focus({ preventScroll: true });
  }
}

function migrateScratchSheetRuntimeMaps(oldName, newName) {
  const maps = [
    scratchGridDimensionsMap,
    scratchActiveCellMap,
    scratchSelectionRangeMap,
    scratchUndoStackMap,
    scratchRedoStackMap,
    scratchPendingColWidths,
    scratchPendingFills,
    scratchSpillStateMap,
    scratchSpilledCellsMap
  ];
  maps.forEach(m => {
    if (m.has(oldName)) {
      if (newName) m.set(newName, m.get(oldName));
      m.delete(oldName);
    }
  });
  if (scratchClipboardRangeState && scratchClipboardRangeState.sheetName === oldName) {
    if (newName) {
      scratchClipboardRangeState.sheetName = newName;
    } else {
      scratchClipboardRangeState = null;
    }
  }
}

function resetScratchWorkbookRuntimeState() {
  stopScratchPointerDrag();
  closeScratchFillDropdown();
  closeScratchContextMenu();
  scratchPendingEdits.clear();
  scratchPendingFills.clear();
  scratchPendingColWidths.clear();
  scratchGridDimensionsMap.clear();
  scratchActiveCellMap.clear();
  scratchSelectionRangeMap.clear();
  scratchUndoStackMap.clear();
  scratchRedoStackMap.clear();
  scratchSpillStateMap.clear();
  scratchSpilledCellsMap.clear();
  scratchClipboardRangeState = null;
  lastFormulaCoalesceMeta = null;
}

function validateScratchSheetNameInput(rawVal, ignoreCurrentName = null) {
  const trimmed = String(rawVal || '').trim();
  if (!trimmed) return 'Please enter a sheet name.';
  if (trimmed.length > 31) return 'Excel worksheet names cannot exceed 31 characters.';
  if (/[:\\/?*[\]]/.test(trimmed)) {
    return 'Sheet names cannot contain : \\ / ? * [ or ].';
  }
  if (trimmed.startsWith("'") || trimmed.endsWith("'")) {
    return 'Sheet names cannot start or end with an apostrophe.';
  }
  if (DATE_TAB_REGEX.test(trimmed)) {
    return 'MM-DD-YYYY names are reserved for Daily Tracker tabs. Use a descriptive name (e.g., "Title Tracking").';
  }
  if (isProtectedReferenceSheetName(trimmed) || isTemplatePlaceholderTab(trimmed)) {
    return `"${trimmed}" is a reserved tracker reference sheet name.`;
  }
  if (workbook) {
    let duplicate = false;
    workbook.eachSheet(ws => {
      if (
        (!ignoreCurrentName || ws.name.toLowerCase() !== ignoreCurrentName.toLowerCase()) &&
        ws.name.trim().toLowerCase() === trimmed.toLowerCase()
      ) {
        duplicate = true;
      }
    });
    if (duplicate) {
      return `A worksheet named "${trimmed}" already exists in this workbook.`;
    }
  }
  return null;
}

async function promptCreateNewScratchSheet() {
  closeScratchMenu();
  closeToolsMenu();
  if (!workbook || !workingFileHandle) {
    return showModalAlert({
      title: 'No Active Workbook',
      message: 'Please connect an Agent Folder and open a workbook first.'
    });
  }

  const defaultSuggested = scratchSheets.some(s => s.name.toLowerCase() === 'title tracking')
    ? `Scratch ${scratchSheets.length + 1}`
    : 'Title Tracking';

  const result = await showModalPrompt({
    title: 'Create New Scratch Sheet',
    badge: workingFileHandle.name,
    message: 'Create a custom worksheet inside your workbook and open it in an app tab (saved directly to Excel & mirrored to View-Only):',
    selectLabel: 'Starter Layout',
    selectOptions: [
      {
        value: 'title_buckets',
        label: 'Title Tracking + Auto-Filtered Red / Yellow / Green Side Lists (Lot #, Buyer #, LLY Date, Tracking #)'
      },
      {
        value: 'title_headers',
        label: 'Title Tracking Columns (Lot #, Yard #, Broker / Master #, Title Status, Carrier / Tracking #, Notes)'
      },
      {
        value: 'blank',
        label: 'Blank Spreadsheet Grid'
      }
    ],
    defaultSelectValue: 'title_buckets',
    inputLabel: 'Scratch Sheet Name',
    defaultValue: defaultSuggested,
    placeholder: 'e.g., Title Tracking, Pending Titles, Scratch',
    confirmText: 'Create & Open Tab',
    validate: (val) => validateScratchSheetNameInput(val)
  });

  if (!result) return;
  const newSheetName = (typeof result === 'object' ? result.value : result).trim();
  const starterLayout = (typeof result === 'object' && result.selectValue) ? result.selectValue : 'blank';

  await loadWorkbookFromDisk(sheetSelect.value);
  const recheckErr = validateScratchSheetNameInput(newSheetName);
  if (recheckErr) {
    return showModalAlert({ title: 'Cannot Create Scratch Sheet', message: recheckErr });
  }

  const newWs = workbook.addWorksheet(newSheetName);

  if (starterLayout === 'title_buckets') {
    const mainHeaders = [
      { header: 'Lot #', width: 15, numFmt: '@' },
      { header: 'Buyer #', width: 14, numFmt: '@' },
      { header: 'LLY Date', width: 14, numFmt: 'mm/dd/yyyy' },
      { header: 'Tracking #', width: 24, numFmt: '@' }
    ];
    const headerRow = newWs.getRow(1);
    headerRow.height = 22;
    mainHeaders.forEach((colDef, idx) => {
      const colObj = newWs.getColumn(idx + 1);
      colObj.width = colDef.width;
      colObj.numFmt = colDef.numFmt;
      const cell = headerRow.getCell(idx + 1);
      cell.value = colDef.header;
      cell.font = { bold: true, color: { argb: 'FF0F172A' } };
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE2E8F0' } };
    });

    const bucketDefs = [
      {
        startCol: 6,
        fillArgb: 'FFFFC7CE',
        fontArgb: 'FF9C0006',
        headers: ['Lot # — No Tracking', 'Buyer #', 'LLY Date'],
        formula: '=_xlfn._xlws.FILTER(A2:C200, (A2:A200<>"")*(D2:D200=""), "")'.slice(1)
      },
      {
        startCol: 10,
        fillArgb: 'FFFFEB9C',
        fontArgb: 'FF9C6500',
        headers: ['Lot # — Duplicate', 'Buyer #', 'LLY Date'],
        formula: '=_xlfn._xlws.FILTER(A2:C200, D2:D200="Duplicate", "")'.slice(1)
      },
      {
        startCol: 14,
        fillArgb: 'FFC6EFCE',
        fontArgb: 'FF006100',
        headers: ['Lot # — Available', 'Buyer #', 'LLY Date', 'Tracking #'],
        formula: '=_xlfn._xlws.FILTER(A2:D200, (D2:D200<>"")*(D2:D200<>"Duplicate"), "")'.slice(1)
      }
    ];
    bucketDefs.forEach(b => {
      b.headers.forEach((hText, offset) => {
        const c = b.startCol + offset;
        newWs.getColumn(c).width = offset === 3 ? 24 : 16;
        const hCell = headerRow.getCell(c);
        hCell.value = hText;
        hCell.font = { bold: true, color: { argb: b.fontArgb } };
        hCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: b.fillArgb } };
        for (let r = 2; r <= 30; r++) {
          const rowObj = newWs.getRow(r);
          const dCell = rowObj.getCell(c);
          dCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: b.fillArgb } };
          dCell.font = { color: { argb: b.fontArgb } };
          if (r === 2 && offset === 0) {
            dCell.value = { formula: b.formula };
          }
          rowObj.commit();
        }
      });
    });
    headerRow.commit();
    scratchActiveCellMap.set(newSheetName, { row: 2, col: 1 });
    scratchSelectionRangeMap.set(newSheetName, { anchorRow: 2, anchorCol: 1, focusRow: 2, focusCol: 1 });
  } else if (starterLayout === 'title_headers') {
    const starterHeaders = [
      { header: 'Lot #', width: 16 },
      { header: 'Yard #', width: 11 },
      { header: 'Broker / Master #', width: 26 },
      { header: 'Title Status', width: 18 },
      { header: 'Carrier / Tracking #', width: 26 },
      { header: 'Date Updated', width: 15 },
      { header: 'Notes', width: 36 }
    ];
    const headerRow = newWs.getRow(1);
    headerRow.height = 22;
    starterHeaders.forEach((colDef, idx) => {
      newWs.getColumn(idx + 1).width = colDef.width;
      const cell = headerRow.getCell(idx + 1);
      cell.value = colDef.header;
      cell.font = { bold: true, color: { argb: 'FF0F172A' } };
      cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FFE2E8F0' }
      };
    });
    headerRow.commit();
    scratchActiveCellMap.set(newSheetName, { row: 2, col: 1 });
    scratchSelectionRangeMap.set(newSheetName, { anchorRow: 2, anchorCol: 1, focusRow: 2, focusCol: 1 });
  } else {
    for (let c = 1; c <= 10; c++) {
      newWs.getColumn(c).width = 18;
    }
    scratchActiveCellMap.set(newSheetName, { row: 1, col: 1 });
    scratchSelectionRangeMap.set(newSheetName, { anchorRow: 1, anchorCol: 1, focusRow: 1, focusCol: 1 });
  }

  const saved = await saveWorkbookToDisk();
  if (saved) {
    if (!openScratchTabs.includes(newSheetName)) {
      openScratchTabs.push(newSheetName);
      persistOpenScratchTabsForWorkbook();
    }
    activeWorkspaceTab = newSheetName;
    await loadWorkbookFromDisk(sheetSelect.value, {
      preserveStatus: saved === 'mirror-locked',
      skipFlushPendingScratch: true
    });
    showToast(`Created scratch sheet "${newSheetName}" and opened in tab.`, 'success');
  }
}

async function handleRenameActiveScratchSheet() {
  if (!workbook || activeWorkspaceTab === DAILY_TRACKER_WS_TAB) return;
  const currentName = activeWorkspaceTab;
  if (!isCustomScratchSheetName(currentName)) return;

  stopScratchPointerDrag();
  commitActiveScratchInlineEdit();
  if (hasAnyPendingScratchMutations()) {
    await flushPendingScratchEditsToDisk({ reloadAfter: false, silentToast: true });
  }

  const rawNew = await showModalPrompt({
    title: 'Rename Scratch Sheet',
    badge: currentName,
    message: `Enter a new worksheet name for "${currentName}":`,
    inputLabel: 'New Scratch Sheet Name',
    defaultValue: currentName,
    confirmText: 'Rename Sheet',
    validate: (val) => validateScratchSheetNameInput(val, currentName)
  });
  if (!rawNew) return;
  const nextName = (typeof rawNew === 'object' ? rawNew.value : rawNew).trim();
  if (nextName === currentName) return;

  await loadWorkbookFromDisk(sheetSelect.value);
  const ws = workbook.getWorksheet(currentName);
  if (!ws) return;

  ws.name = nextName;
  const saved = await saveWorkbookToDisk();
  if (saved) {
    openScratchTabs = openScratchTabs.map(n => (n === currentName ? nextName : n));
    persistOpenScratchTabsForWorkbook();
    migrateScratchSheetRuntimeMaps(currentName, nextName);
    activeWorkspaceTab = nextName;
    await loadWorkbookFromDisk(sheetSelect.value, {
      preserveStatus: saved === 'mirror-locked',
      skipFlushPendingScratch: true
    });
    showToast(`Renamed scratch sheet "${currentName}" → "${nextName}".`, 'success');
  }
}

async function handleDeleteActiveScratchSheet() {
  if (!workbook || activeWorkspaceTab === DAILY_TRACKER_WS_TAB) return;
  const targetName = activeWorkspaceTab;
  if (!isCustomScratchSheetName(targetName)) return;

  stopScratchPointerDrag();
  commitActiveScratchInlineEdit();
  const meta = scratchSheets.find(s => s.name === targetName);
  const rowCount = meta ? meta.nonEmptyRows : 0;

  const confirmed = await showModalConfirm({
    title: `Delete Scratch Sheet "${targetName}"?`,
    badge: `${rowCount} non-empty row(s)`,
    message:
      `Are you sure you want to permanently delete worksheet "${targetName}" from "${workingFileHandle.name}"?\n\n` +
      `(Tip: If you only want to hide the tab in the app without deleting the Excel sheet, click the "×" button on the tab instead.)`,
    confirmText: 'Delete Worksheet Permanently',
    variant: 'danger'
  });
  if (!confirmed) return;

  scratchPendingEdits.delete(targetName);
  scratchPendingFills.delete(targetName);
  scratchPendingColWidths.delete(targetName);
  await loadWorkbookFromDisk(sheetSelect.value, { skipFlushPendingScratch: true });
  const ws = workbook.getWorksheet(targetName);
  if (!ws) return;

  if (workbook.worksheets.length <= 1) {
    return showModalAlert({
      title: 'Cannot Delete Only Worksheet',
      message: 'Excel workbooks must contain at least one worksheet.'
    });
  }

  workbook.removeWorksheet(ws.id);
  const saved = await saveWorkbookToDisk();
  if (saved) {
    openScratchTabs = openScratchTabs.filter(n => n !== targetName);
    persistOpenScratchTabsForWorkbook();
    migrateScratchSheetRuntimeMaps(targetName, null);
    activeWorkspaceTab = openScratchTabs[0] || DAILY_TRACKER_WS_TAB;
    await loadWorkbookFromDisk(sheetSelect.value, {
      preserveStatus: saved === 'mirror-locked',
      skipFlushPendingScratch: true
    });
    showToast(`Deleted scratch worksheet "${targetName}".`, 'success');
  }
}

// ============================================================================
// EVENT LISTENERS — SCRATCH MENU, TOOLBAR, FILL PALETTE, CONTEXT MENU & GRID
// ============================================================================
initScratchFillPaletteUI();

if (btnOpenScratchMenu) {
  btnOpenScratchMenu.addEventListener('click', (e) => {
    e.stopPropagation();
    closeToolsMenu();
    closeScratchFillDropdown();
    closeScratchContextMenu();
    const isOpen = scratchSheetDropdown && scratchSheetDropdown.classList.contains('open');
    if (isOpen) {
      closeScratchMenu();
    } else if (scratchSheetDropdown) {
      renderScratchMenuDropdown();
      scratchSheetDropdown.classList.add('open');
      btnOpenScratchMenu.setAttribute('aria-expanded', 'true');
    }
  });
}

document.addEventListener('click', (e) => {
  if (scratchSheetMenuWrap && !scratchSheetMenuWrap.contains(e.target)) {
    closeScratchMenu();
  }
  if (scratchFillMenuWrap && !scratchFillMenuWrap.contains(e.target)) {
    closeScratchFillDropdown();
  }
  if (scratchContextMenu && !scratchContextMenu.contains(e.target)) {
    closeScratchContextMenu();
  }
});

if (chkScratchStripedRows) {
  chkScratchStripedRows.addEventListener('change', () => {
    scratchStripedRowsEnabled = Boolean(chkScratchStripedRows.checked);
    try {
      localStorage.setItem(SCRATCH_STRIPED_STORAGE_KEY, String(scratchStripedRowsEnabled));
    } catch (_) { /* ignore */ }
    syncScratchStripedRowsUI();
    if (scratchGridWrap) scratchGridWrap.focus({ preventScroll: true });
  });
}

if (chkScratchAutoIncrement) {
  chkScratchAutoIncrement.addEventListener('change', () => {
    scratchAutoIncrementEnabled = Boolean(chkScratchAutoIncrement.checked);
    try {
      localStorage.setItem(SCRATCH_AUTO_INCREMENT_STORAGE_KEY, String(scratchAutoIncrementEnabled));
    } catch (_) { /* ignore */ }
    syncScratchAutoIncrementUI();
    if (scratchGridWrap) scratchGridWrap.focus({ preventScroll: true });
  });
}

if (btnScratchApplyFill) {
  btnScratchApplyFill.addEventListener('click', () => {
    closeScratchFillDropdown();
    closeScratchContextMenu();
    applyScratchFillPresetToSelection(activeWorkspaceTab, activeScratchFillPresetId);
    if (scratchGridWrap) scratchGridWrap.focus({ preventScroll: true });
  });
}

if (btnScratchFillDropdownToggle) {
  btnScratchFillDropdownToggle.addEventListener('click', (e) => {
    e.stopPropagation();
    closeToolsMenu();
    closeScratchMenu();
    closeScratchContextMenu();
    const isOpen = scratchFillDropdown && scratchFillDropdown.classList.contains('open');
    if (isOpen) {
      closeScratchFillDropdown();
    } else if (scratchFillDropdown) {
      syncScratchFillIndicatorUI();
      scratchFillDropdown.classList.add('open');
      btnScratchFillDropdownToggle.setAttribute('aria-expanded', 'true');
    }
  });
}

if (scratchFillDropdown) {
  scratchFillDropdown.addEventListener('click', (e) => {
    const fillBtn = e.target.closest('[data-fill-id]');
    if (!fillBtn) return;
    const presetId = fillBtn.dataset.fillId;
    closeScratchFillDropdown();
    applyScratchFillPresetToSelection(activeWorkspaceTab, presetId);
    if (scratchGridWrap) scratchGridWrap.focus({ preventScroll: true });
  });
}

if (scratchCellFormatSelect) {
  scratchCellFormatSelect.addEventListener('change', () => {
    if (activeWorkspaceTab === DAILY_TRACKER_WS_TAB || !workbook) return;
    applyScratchFormatToSelection(activeWorkspaceTab, scratchCellFormatSelect.value);
    if (scratchGridWrap) scratchGridWrap.focus({ preventScroll: true });
  });
}

if (btnScratchTextToColumns) {
  btnScratchTextToColumns.addEventListener('click', openTextToColumnsModal);
}

if (btnScratchAutoFilter) {
  btnScratchAutoFilter.addEventListener('click', openScratchAutoFilterModal);
}

// Auto-Filter Side Lists Modal Listeners
[
  afPresetModeSelect,
  afReturnStartCol,
  afReturnEndCol,
  afCriteriaCol,
  afDestStartCol,
  afStartRowInput,
  afMaxRowInput,
  afYellowKeywordInput,
  afIncludeCriteriaInGreen,
  afCustomOperatorSelect,
  afCustomValueInput,
  afCustomColorSelect
].forEach(el => {
  if (el) {
    el.addEventListener('change', renderScratchAutoFilterPreview);
    if (el.tagName === 'INPUT' && el.type !== 'checkbox') {
      el.addEventListener('input', renderScratchAutoFilterPreview);
    }
  }
});

if (btnCancelAutoFilter) {
  btnCancelAutoFilter.addEventListener('click', closeScratchAutoFilterModal);
}

if (scratchAutoFilterModal) {
  scratchAutoFilterModal.addEventListener('click', (e) => {
    if (e.target === scratchAutoFilterModal) {
      closeScratchAutoFilterModal();
    }
  });
}

if (scratchAutoFilterForm) {
  scratchAutoFilterForm.addEventListener('submit', (e) => {
    e.preventDefault();
    executeScratchAutoFilterSetup();
  });
}

// Data to Columns Modal Listeners
[chkTtcComma, chkTtcTab, chkTtcSemicolon, chkTtcPipe, chkTtcHyphen, chkTtcSpace, chkTtcNewline, chkTtcOther, chkTtcTrim, chkTtcConsecutive, radTtcAcross, radTtcDown].forEach(el => {
  if (el) {
    el.addEventListener('change', renderTextToColumnsPreview);
  }
});

if (ttcOtherInput) {
  ttcOtherInput.addEventListener('input', () => {
    if (chkTtcOther && ttcOtherInput.value.length > 0 && !chkTtcOther.checked) {
      chkTtcOther.checked = true;
    }
    renderTextToColumnsPreview();
  });
}

if (btnCancelTextToColumns) {
  btnCancelTextToColumns.addEventListener('click', closeTextToColumnsModal);
}

if (textToColumnsModal) {
  textToColumnsModal.addEventListener('click', (e) => {
    if (e.target === textToColumnsModal) {
      closeTextToColumnsModal();
    }
  });
}

if (textToColumnsForm) {
  textToColumnsForm.addEventListener('submit', (e) => {
    e.preventDefault();
    executeTextToColumnsSplit();
  });
}

// Right-Click Context Menu Actions
if (scratchContextMenu) {
  scratchContextMenu.addEventListener('click', (e) => {
    const swatchBtn = e.target.closest('[data-fill-id]');
    if (swatchBtn) {
      const presetId = swatchBtn.dataset.fillId;
      closeScratchContextMenu();
      applyScratchFillPresetToSelection(activeWorkspaceTab, presetId);
      if (scratchGridWrap) scratchGridWrap.focus({ preventScroll: true });
      return;
    }

    const actionBtn = e.target.closest('[data-ctx-action]');
    if (!actionBtn || activeWorkspaceTab === DAILY_TRACKER_WS_TAB || !workbook) return;
    const action = actionBtn.dataset.ctxAction;
    const range = getScratchSelectionRange(activeWorkspaceTab);
    closeScratchContextMenu();

    switch (action) {
      case 'cut':
        performScratchCopyOrCut(activeWorkspaceTab, range, { isCut: true, writeSystemClipboard: true });
        if (scratchGridWrap) scratchGridWrap.focus({ preventScroll: true });
        break;
      case 'copy':
        performScratchCopyOrCut(activeWorkspaceTab, range, { isCut: false, writeSystemClipboard: true });
        if (scratchGridWrap) scratchGridWrap.focus({ preventScroll: true });
        break;
      case 'paste':
        handleScratchContextPaste();
        break;
      case 'paste-transpose':
        handleScratchContextPaste({ transpose: true });
        break;
      case 'transpose-selection':
        transposeScratchSelection(activeWorkspaceTab, range);
        if (scratchGridWrap) scratchGridWrap.focus({ preventScroll: true });
        break;
      case 'text-to-columns':
        openTextToColumnsModal();
        break;
      case 'auto-filter-lists':
        openScratchAutoFilterModal();
        break;
      case 'clear-contents': {
        const cleared = clearScratchSelectionCells(activeWorkspaceTab, range);
        syncScratchFormulaBarFromActiveCell();
        if (cleared > 1) {
          showToast(`Cleared ${cleared} cells (${getRangeCoordString(range)}).`, 'info');
        }
        if (scratchGridWrap) scratchGridWrap.focus({ preventScroll: true });
        break;
      }
      case 'clear-fill':
        applyScratchFillPresetToSelection(activeWorkspaceTab, 'none');
        if (scratchGridWrap) scratchGridWrap.focus({ preventScroll: true });
        break;
      default:
        break;
    }
  });
}

if (btnNewScratchSheetFromBar) {
  btnNewScratchSheetFromBar.addEventListener('click', promptCreateNewScratchSheet);
}
if (menuNewScratchSheet) {
  menuNewScratchSheet.addEventListener('click', promptCreateNewScratchSheet);
}

if (scratchSearchInput) {
  scratchSearchInput.addEventListener('input', () => {
    commitActiveScratchInlineEdit();
    renderActiveScratchSheetGrid();
  });
}
if (chkScratchFilterRows) {
  chkScratchFilterRows.addEventListener('change', () => {
    commitActiveScratchInlineEdit();
    renderActiveScratchSheetGrid();
  });
}

if (btnScratchUndo) {
  btnScratchUndo.addEventListener('click', () => {
    undoScratchEdit();
    if (scratchGridWrap) scratchGridWrap.focus({ preventScroll: true });
  });
}
if (btnScratchRedo) {
  btnScratchRedo.addEventListener('click', () => {
    redoScratchEdit();
    if (scratchGridWrap) scratchGridWrap.focus({ preventScroll: true });
  });
}

if (btnScratchAddRows) {
  btnScratchAddRows.addEventListener('click', () => {
    if (activeWorkspaceTab === DAILY_TRACKER_WS_TAB || !workbook) return;
    commitActiveScratchInlineEdit();
    const ws = workbook.getWorksheet(activeWorkspaceTab);
    if (!ws) return;
    const dims = getScratchSheetGridDimensions(ws, activeWorkspaceTab);
    scratchGridDimensionsMap.set(activeWorkspaceTab, { rows: dims.rows + 10, cols: dims.cols });
    renderActiveScratchSheetGrid();
  });
}

if (btnScratchAddCol) {
  btnScratchAddCol.addEventListener('click', () => {
    if (activeWorkspaceTab === DAILY_TRACKER_WS_TAB || !workbook) return;
    commitActiveScratchInlineEdit();
    const ws = workbook.getWorksheet(activeWorkspaceTab);
    if (!ws) return;
    const dims = getScratchSheetGridDimensions(ws, activeWorkspaceTab);
    scratchGridDimensionsMap.set(activeWorkspaceTab, { rows: dims.rows, cols: dims.cols + 1 });
    renderActiveScratchSheetGrid();
  });
}

if (btnScratchRenameSheet) {
  btnScratchRenameSheet.addEventListener('click', handleRenameActiveScratchSheet);
}
if (btnScratchDeleteSheet) {
  btnScratchDeleteSheet.addEventListener('click', handleDeleteActiveScratchSheet);
}
if (btnScratchSaveNow) {
  btnScratchSaveNow.addEventListener('click', () => {
    flushPendingScratchEditsToDisk({ reloadAfter: false, silentToast: false });
  });
}

if (scratchFormulaInput) {
  scratchFormulaInput.addEventListener('input', () => {
    if (activeWorkspaceTab === DAILY_TRACKER_WS_TAB || !workbook) return;
    const { row, col } = getActiveScratchCell(activeWorkspaceTab);
    if (scratchInlineEditState && scratchInlineEditState.inputEl) {
      scratchInlineEditState.inputEl.value = scratchFormulaInput.value;
    } else {
      recordScratchCellEdit(activeWorkspaceTab, row, col, scratchFormulaInput.value, {
        coalesceKey: `fx::${activeWorkspaceTab}::${row}::${col}`
      });
    }
  });

  scratchFormulaInput.addEventListener('keydown', (e) => {
    if (activeWorkspaceTab === DAILY_TRACKER_WS_TAB) return;
    const { row, col } = getActiveScratchCell(activeWorkspaceTab);
    if (e.key === 'Enter') {
      e.preventDefault();
      lastFormulaCoalesceMeta = null;
      selectScratchCell(e.shiftKey ? Math.max(1, row - 1) : row + 1, col, { scrollIntoView: true });
      if (scratchGridWrap) scratchGridWrap.focus({ preventScroll: true });
    } else if (e.key === 'Tab') {
      e.preventDefault();
      lastFormulaCoalesceMeta = null;
      selectScratchCell(row, e.shiftKey ? Math.max(1, col - 1) : col + 1, { scrollIntoView: true });
      if (scratchGridWrap) scratchGridWrap.focus({ preventScroll: true });
    } else if (e.key === 'Escape') {
      e.preventDefault();
      lastFormulaCoalesceMeta = null;
      if (scratchGridWrap) scratchGridWrap.focus({ preventScroll: true });
      syncScratchFormulaBarFromActiveCell();
    }
  });
}

// Column Header Selection + Column Width Drag-Resize + Corner Select-All
if (scratchGridHead) {
  scratchGridHead.addEventListener('mousedown', (e) => {
    if (e.button !== 0 || activeWorkspaceTab === DAILY_TRACKER_WS_TAB || !workbook) return;
    closeScratchContextMenu();
    closeScratchFillDropdown();
    const ws = workbook.getWorksheet(activeWorkspaceTab);
    if (!ws) return;

    // 1. Column Width Drag-Resize Handle
    const resizer = e.target.closest('.scratch-col-resizer[data-resize-col]');
    if (resizer) {
      e.preventDefault();
      e.stopPropagation();
      commitActiveScratchInlineEdit();
      const colNum = parseInt(resizer.dataset.resizeCol, 10);
      const thEl = resizer.closest('th[data-col-hdr]');
      if (!thEl || !colNum) return;
      const startW = thEl.getBoundingClientRect().width;
      resizer.classList.add('resizing');
      scratchColResizeState = {
        sheetName: activeWorkspaceTab,
        col: colNum,
        thEl,
        resizerEl: resizer,
        startX: e.clientX,
        startWidthPx: startW,
        currentWidthPx: startW
      };
      return;
    }

    // 2. Top-Left Corner '#' Select-All Button
    const cornerTh = e.target.closest('th.scratch-corner-th');
    if (cornerTh) {
      e.preventDefault();
      commitActiveScratchInlineEdit();
      if (scratchGridWrap) scratchGridWrap.focus({ preventScroll: true });
      const dims = getScratchSheetGridDimensions(ws, activeWorkspaceTab);
      setScratchSelectionRange(activeWorkspaceTab, {
        anchorRow: 1,
        anchorCol: 1,
        focusRow: dims.rows,
        focusCol: dims.cols
      });
      return;
    }

    // 3. Whole-Column Click / Shift+Click / Drag Selection
    const colTh = e.target.closest('th[data-col-hdr]');
    if (colTh) {
      e.preventDefault();
      commitActiveScratchInlineEdit();
      if (scratchGridWrap) scratchGridWrap.focus({ preventScroll: true });
      const c = parseInt(colTh.dataset.colHdr, 10);
      if (!c) return;
      const dims = getScratchSheetGridDimensions(ws, activeWorkspaceTab);
      const anchorCol = e.shiftKey ? getScratchSelectionRange(activeWorkspaceTab).anchorCol : c;
      setScratchSelectionRange(activeWorkspaceTab, {
        anchorRow: 1,
        anchorCol,
        focusRow: dims.rows,
        focusCol: c
      });
      startScratchPointerDrag('cols', e.clientX, e.clientY);
    }
  });

  scratchGridHead.addEventListener('dblclick', (e) => {
    const resizer = e.target.closest('.scratch-col-resizer[data-resize-col]');
    if (!resizer || activeWorkspaceTab === DAILY_TRACKER_WS_TAB) return;
    e.preventDefault();
    e.stopPropagation();
    const colNum = parseInt(resizer.dataset.resizeCol, 10);
    if (colNum >= 1) {
      autoFitScratchColumnWidth(activeWorkspaceTab, colNum);
    }
  });
}

// Cell & Row-Header Selection + Drag Highlighting + Double-Click Inline Edit
if (scratchGridBody) {
  scratchGridBody.addEventListener('mousedown', (e) => {
    if (e.button !== 0 || activeWorkspaceTab === DAILY_TRACKER_WS_TAB || !workbook) return;
    closeScratchContextMenu();
    closeScratchFillDropdown();
    const ws = workbook.getWorksheet(activeWorkspaceTab);
    if (!ws) return;

    // 1. Whole-Row Header Click / Shift+Click / Drag
    const rowTh = e.target.closest('th[data-row-hdr]');
    if (rowTh) {
      e.preventDefault();
      commitActiveScratchInlineEdit();
      if (scratchGridWrap) scratchGridWrap.focus({ preventScroll: true });
      const r = parseInt(rowTh.dataset.rowHdr, 10);
      if (!r) return;
      const dims = getScratchSheetGridDimensions(ws, activeWorkspaceTab);
      const anchorRow = e.shiftKey ? getScratchSelectionRange(activeWorkspaceTab).anchorRow : r;
      setScratchSelectionRange(activeWorkspaceTab, {
        anchorRow,
        anchorCol: 1,
        focusRow: r,
        focusCol: dims.cols
      });
      startScratchPointerDrag('rows', e.clientX, e.clientY);
      return;
    }

    // 2. Corner Fill Handle Drag
    const fillHandle = e.target.closest('.scratch-fill-handle');
    if (fillHandle) {
      e.preventDefault();
      e.stopPropagation();
      commitActiveScratchInlineEdit();
      if (scratchGridWrap) scratchGridWrap.focus({ preventScroll: true });
      const curRange = getScratchSelectionRange(activeWorkspaceTab);
      startScratchPointerDrag('fill', e.clientX, e.clientY, {
        sourceRange: { ...curRange },
        fillTargetRange: null
      });
      return;
    }

    // 3. Grid Cell Click / Shift+Click / Drag Range Selection
    const td = e.target.closest('td[data-r][data-c]');
    if (!td) return;
    if (scratchInlineEditState && scratchInlineEditState.tdEl === td) return;

    e.preventDefault();
    commitActiveScratchInlineEdit();
    if (scratchGridWrap) scratchGridWrap.focus({ preventScroll: true });

    const r = parseInt(td.dataset.r, 10);
    const c = parseInt(td.dataset.c, 10);
    selectScratchCell(r, c, { extendSelection: e.shiftKey });
    startScratchPointerDrag('cells', e.clientX, e.clientY);
  });

  scratchGridBody.addEventListener('dblclick', (e) => {
    const fillHandle = e.target.closest('.scratch-fill-handle');
    if (fillHandle && activeWorkspaceTab !== DAILY_TRACKER_WS_TAB && workbook) {
      e.preventDefault();
      e.stopPropagation();
      const ws = workbook.getWorksheet(activeWorkspaceTab);
      if (!ws) return;
      const curRange = getScratchSelectionRange(activeWorkspaceTab);
      const dims = getScratchSheetGridDimensions(ws, activeWorkspaceTab);
      // Find contiguous filled rows in adjacent left (or right) column
      const probeCols = [curRange.minCol - 1, curRange.maxCol + 1].filter(col => col >= 1 && col <= dims.cols);
      let targetEndRow = curRange.maxRow;
      for (const pCol of probeCols) {
        let r = curRange.maxRow + 1;
        while (r <= Math.max(dims.dataMaxRow, dims.rows) && extractScratchCellInfo(ws, r, pCol).display.trim() !== '') {
          if (r > targetEndRow) targetEndRow = r;
          r++;
        }
        if (targetEndRow > curRange.maxRow) break;
      }
      if (targetEndRow > curRange.maxRow) {
        executeScratchFillHandleDrag(
          activeWorkspaceTab,
          curRange,
          {
            axis: 'vertical',
            dir: 1,
            minRow: curRange.minRow,
            maxRow: targetEndRow,
            minCol: curRange.minCol,
            maxCol: curRange.maxCol
          },
          { ctrlKey: Boolean(e.ctrlKey || e.metaKey) }
        );
      }
      return;
    }

    const td = e.target.closest('td[data-r][data-c]');
    if (!td) return;
    const r = parseInt(td.dataset.r, 10);
    const c = parseInt(td.dataset.c, 10);
    startScratchInlineEdit(r, c);
  });
}

// Global Pointer Tracking for Cell/Header Dragging & Column Width Resizing
window.addEventListener('mousemove', (e) => {
  if (scratchColResizeState) {
    const deltaX = e.clientX - scratchColResizeState.startX;
    const nextW = Math.max(68, Math.min(500, Math.round(scratchColResizeState.startWidthPx + deltaX)));
    scratchColResizeState.currentWidthPx = nextW;
    scratchColResizeState.thEl.style.width = `${nextW}px`;
    scratchColResizeState.thEl.style.minWidth = `${nextW}px`;
    return;
  }

  if (scratchDragState) {
    scratchDragState.clientX = e.clientX;
    scratchDragState.clientY = e.clientY;
    updateScratchDragSelectionFromPoint(e.clientX, e.clientY);
  }
});

window.addEventListener('mouseup', (e) => {
  if (scratchColResizeState) {
    const st = scratchColResizeState;
    scratchColResizeState = null;
    if (st.resizerEl) st.resizerEl.classList.remove('resizing');
    if (Math.abs(st.currentWidthPx - st.startWidthPx) >= 2) {
      recordScratchColumnWidthChange(st.sheetName, st.col, st.currentWidthPx);
    }
  }
  if (scratchDragState) {
    const dragSnap = scratchDragState;
    if (dragSnap.mode === 'fill' && dragSnap.fillTargetRange) {
      const src = dragSnap.sourceRange;
      const target = dragSnap.fillTargetRange;
      const sheet = dragSnap.sheetName;
      stopScratchPointerDrag();
      executeScratchFillHandleDrag(sheet, src, target, {
        ctrlKey: Boolean(e && (e.ctrlKey || e.metaKey))
      });
    } else {
      stopScratchPointerDrag();
    }
  }
});

window.addEventListener('blur', () => {
  if (scratchDragState) {
    stopScratchPointerDrag();
  }
});

// Keyboard Navigation, Shortcuts & Clipboard Copy/Cut/Paste on Grid
if (scratchGridWrap) {
  scratchGridWrap.addEventListener('keydown', (e) => {
    if (activeWorkspaceTab === DAILY_TRACKER_WS_TAB || !workbook) return;
    if (scratchInlineEditState) return; // Handled by inline input

    const ws = workbook.getWorksheet(activeWorkspaceTab);
    if (!ws) return;
    const range = getScratchSelectionRange(activeWorkspaceTab);
    const dims = getScratchSheetGridDimensions(ws, activeWorkspaceTab);
    const keyLower = e.key.toLowerCase();
    const mod = e.ctrlKey || e.metaKey;

    // Save: Ctrl/Cmd + S
    if (mod && keyLower === 's') {
      e.preventDefault();
      flushPendingScratchEditsToDisk({ reloadAfter: false, silentToast: false });
      return;
    }

    // Undo: Ctrl/Cmd + Z (without Shift)
    if (mod && keyLower === 'z' && !e.shiftKey) {
      e.preventDefault();
      undoScratchEdit();
      return;
    }

    // Redo: Ctrl/Cmd + Y or Ctrl/Cmd + Shift + Z
    if (mod && (keyLower === 'y' || (keyLower === 'z' && e.shiftKey))) {
      e.preventDefault();
      redoScratchEdit();
      return;
    }

    // Select All: Ctrl/Cmd + A (Data bounds first, then entire grid on second press)
    if (mod && keyLower === 'a') {
      e.preventDefault();
      const hasData = dims.dataMaxRow >= 1 && dims.dataMaxCol >= 1;
      const alreadyCoversData =
        hasData &&
        range.minRow === 1 &&
        range.maxRow === dims.dataMaxRow &&
        range.minCol === 1 &&
        range.maxCol === dims.dataMaxCol;

      if (hasData && !alreadyCoversData) {
        setScratchSelectionRange(activeWorkspaceTab, {
          anchorRow: 1,
          anchorCol: 1,
          focusRow: dims.dataMaxRow,
          focusCol: dims.dataMaxCol
        });
      } else {
        setScratchSelectionRange(activeWorkspaceTab, {
          anchorRow: 1,
          anchorCol: 1,
          focusRow: dims.rows,
          focusCol: dims.cols
        });
      }
      return;
    }

    // Copy / Cut fallback if browser does not dispatch native copy/cut event
    if (mod && (keyLower === 'c' || keyLower === 'x')) {
      const isCut = keyLower === 'x';
      const snapRange = { ...range };
      const snapSheet = activeWorkspaceTab;
      setTimeout(() => {
        if (Date.now() - lastScratchNativeCopyAt < 80) return;
        performScratchCopyOrCut(snapSheet, snapRange, { isCut, writeSystemClipboard: true });
      }, 15);
      return;
    }

    // Arrow Navigation (Single move, Shift+Arrow range extend, Ctrl+Arrow data jump, Ctrl+Shift+Arrow)
    if (
      e.key === 'ArrowUp' ||
      e.key === 'ArrowDown' ||
      e.key === 'ArrowLeft' ||
      e.key === 'ArrowRight'
    ) {
      e.preventDefault();
      const dr = e.key === 'ArrowUp' ? -1 : e.key === 'ArrowDown' ? 1 : 0;
      const dc = e.key === 'ArrowLeft' ? -1 : e.key === 'ArrowRight' ? 1 : 0;
      const baseRow = e.shiftKey ? range.focusRow : range.anchorRow;
      const baseCol = e.shiftKey ? range.focusCol : range.anchorCol;

      let nextRow = Math.max(1, baseRow + dr);
      let nextCol = Math.max(1, baseCol + dc);
      if (mod) {
        const jumped = computeScratchCtrlArrowTarget(ws, dims, baseRow, baseCol, dr, dc);
        nextRow = jumped.row;
        nextCol = jumped.col;
      }

      selectScratchCell(nextRow, nextCol, {
        scrollIntoView: true,
        extendSelection: e.shiftKey
      });
      return;
    }

    switch (e.key) {
      case 'Tab':
        e.preventDefault();
        selectScratchCell(
          range.anchorRow,
          e.shiftKey ? Math.max(1, range.anchorCol - 1) : range.anchorCol + 1,
          { scrollIntoView: true }
        );
        return;
      case 'Enter':
        e.preventDefault();
        if (
          scratchClipboardRangeState &&
          scratchInternalClipboardText !== ''
        ) {
          applyScratchPasteFromText(
            activeWorkspaceTab,
            scratchInternalClipboardText,
            scratchInternalClipboardState ? scratchInternalClipboardState.html : '',
            { clearCopyMarquee: true }
          );
          return;
        }
        selectScratchCell(
          e.shiftKey ? Math.max(1, range.anchorRow - 1) : range.anchorRow + 1,
          range.anchorCol,
          { scrollIntoView: true }
        );
        return;
      case 'F2':
        e.preventDefault();
        startScratchInlineEdit(range.anchorRow, range.anchorCol);
        return;
      case 'Delete':
      case 'Backspace': {
        e.preventDefault();
        const cleared = clearScratchSelectionCells(activeWorkspaceTab, range);
        syncScratchFormulaBarFromActiveCell();
        if (range.cellCount > 1 && cleared > 1) {
          showToast(`Cleared ${cleared} cells (${getRangeCoordString(range)}).`, 'info');
        }
        return;
      }
      case 'Escape':
        if (scratchClipboardRangeState) {
          e.preventDefault();
          scratchClipboardRangeState = null;
          applyScratchSelectionDOM();
          return;
        }
        if (range.cellCount > 1) {
          e.preventDefault();
          selectScratchCell(range.anchorRow, range.anchorCol);
          return;
        }
        break;
      default:
        break;
    }

    // Printable character starts immediate inline cell entry on the anchor cell (Excel-style)
    if (e.key.length === 1 && !mod && !e.altKey) {
      e.preventDefault();
      startScratchInlineEdit(range.anchorRow, range.anchorCol, e.key);
    }
  });

  // Right-Click Context Menu on Scratch Grid (Cut / Copy / Paste / Fill / Data to Columns)
  scratchGridWrap.addEventListener('contextmenu', (e) => {
    if (activeWorkspaceTab === DAILY_TRACKER_WS_TAB || !workbook) return;
    // Allow native browser right-click menu when actively editing text inside an inline cell input
    if (e.target.closest('.scratch-inline-editor')) return;

    e.preventDefault();
    closeToolsMenu();
    closeScratchMenu();
    closeScratchFillDropdown();
    commitActiveScratchInlineEdit();

    const ws = workbook.getWorksheet(activeWorkspaceTab);
    if (!ws) return;
    const dims = getScratchSheetGridDimensions(ws, activeWorkspaceTab);
    const curRange = getScratchSelectionRange(activeWorkspaceTab);

    const td = e.target.closest('td[data-r][data-c]');
    const rowTh = e.target.closest('th[data-row-hdr]');
    const colTh = e.target.closest('th[data-col-hdr]');
    const cornerTh = e.target.closest('th.scratch-corner-th');

    if (td) {
      const r = parseInt(td.dataset.r, 10);
      const c = parseInt(td.dataset.c, 10);
      const insideSelection =
        r >= curRange.minRow &&
        r <= curRange.maxRow &&
        c >= curRange.minCol &&
        c <= curRange.maxCol;
      if (!insideSelection && r >= 1 && c >= 1) {
        selectScratchCell(r, c);
      }
    } else if (rowTh) {
      const r = parseInt(rowTh.dataset.rowHdr, 10);
      const rowAlreadySelected =
        curRange.minCol === 1 &&
        curRange.maxCol >= dims.cols &&
        r >= curRange.minRow &&
        r <= curRange.maxRow;
      if (!rowAlreadySelected && r >= 1) {
        setScratchSelectionRange(activeWorkspaceTab, {
          anchorRow: r,
          anchorCol: 1,
          focusRow: r,
          focusCol: dims.cols
        });
      }
    } else if (colTh) {
      const c = parseInt(colTh.dataset.colHdr, 10);
      const colAlreadySelected =
        curRange.minRow === 1 &&
        curRange.maxRow >= dims.rows &&
        c >= curRange.minCol &&
        c <= curRange.maxCol;
      if (!colAlreadySelected && c >= 1) {
        setScratchSelectionRange(activeWorkspaceTab, {
          anchorRow: 1,
          anchorCol: c,
          focusRow: dims.rows,
          focusCol: c
        });
      }
    } else if (cornerTh) {
      setScratchSelectionRange(activeWorkspaceTab, {
        anchorRow: 1,
        anchorCol: 1,
        focusRow: dims.rows,
        focusCol: dims.cols
      });
    }

    scratchGridWrap.focus({ preventScroll: true });
    openScratchContextMenuAt(e.clientX, e.clientY);
  });

  scratchGridWrap.addEventListener('scroll', () => {
    closeScratchContextMenu();
  });

  // Multi-Cell Excel HTML + TSV Copy Handler (Ctrl+C)
  scratchGridWrap.addEventListener('copy', (e) => {
    if (activeWorkspaceTab === DAILY_TRACKER_WS_TAB || !workbook) return;
    if (scratchInlineEditState) return;

    const ws = workbook.getWorksheet(activeWorkspaceTab);
    if (!ws) return;
    const range = getScratchSelectionRange(activeWorkspaceTab);
    const effRange = resolveEffectiveScratchCopyRange(ws, range);
    const tsv = serializeScratchSelectionToTSV(ws, effRange);
    const html = serializeScratchSelectionToExcelHtml(ws, effRange);
    const styleMatrix = buildScratchStyleMatrix(ws, effRange);
    const rawEditMatrix = buildScratchRawEditMatrix(ws, effRange);
    scratchInternalClipboardText = tsv;
    scratchInternalClipboardState = {
      tsv,
      html,
      styleMatrix,
      rawEditMatrix,
      sourceRange: {
        minRow: effRange.minRow,
        maxRow: effRange.maxRow,
        minCol: effRange.minCol,
        maxCol: effRange.maxCol
      }
    };

    if (e.clipboardData) {
      e.preventDefault();
      e.clipboardData.setData('text/plain', tsv);
      if (html) e.clipboardData.setData('text/html', html);
      lastScratchNativeCopyAt = Date.now();
    }

    scratchClipboardRangeState = {
      sheetName: activeWorkspaceTab,
      minRow: effRange.minRow,
      maxRow: effRange.maxRow,
      minCol: effRange.minCol,
      maxCol: effRange.maxCol
    };
    applyScratchSelectionDOM();
    if (effRange.cellCount > 1) {
      showToast(`Copied ${effRange.cellCount} cells (${getRangeCoordString(effRange)}).`, 'info');
    }
  });

  // Multi-Cell Excel HTML + TSV Cut Handler (Ctrl+X)
  scratchGridWrap.addEventListener('cut', (e) => {
    if (activeWorkspaceTab === DAILY_TRACKER_WS_TAB || !workbook) return;
    if (scratchInlineEditState) return;

    const ws = workbook.getWorksheet(activeWorkspaceTab);
    if (!ws) return;
    const range = getScratchSelectionRange(activeWorkspaceTab);
    const tsv = serializeScratchSelectionToTSV(ws, range);
    const html = serializeScratchSelectionToExcelHtml(ws, range);
    const styleMatrix = buildScratchStyleMatrix(ws, range);
    const rawEditMatrix = buildScratchRawEditMatrix(ws, range);
    scratchInternalClipboardText = tsv;
    scratchInternalClipboardState = {
      tsv,
      html,
      styleMatrix,
      rawEditMatrix,
      sourceRange: {
        minRow: range.minRow,
        maxRow: range.maxRow,
        minCol: range.minCol,
        maxCol: range.maxCol
      }
    };

    if (e.clipboardData) {
      e.preventDefault();
      e.clipboardData.setData('text/plain', tsv);
      if (html) e.clipboardData.setData('text/html', html);
      lastScratchNativeCopyAt = Date.now();
    }

    scratchClipboardRangeState = null;
    clearScratchSelectionCells(activeWorkspaceTab, range);
    syncScratchFormulaBarFromActiveCell();
    if (range.cellCount > 1) {
      showToast(`Cut ${range.cellCount} cells (${getRangeCoordString(range)}).`, 'info');
    }
  });

  // Multi-Cell TSV / Excel / HTML Clipboard Paste Handler (Ctrl+V + 1-Cell Repeat-Fill)
  function isScratchPasteBlockedByInputEl(el) {
    if (!el || !el.tagName) return false;
    if (el.isContentEditable) return true;
    const tag = el.tagName.toUpperCase();
    if (tag === 'TEXTAREA' || tag === 'SELECT') return true;
    if (tag === 'INPUT') {
      const type = String(el.type || 'text').toLowerCase();
      return !['checkbox', 'radio', 'button', 'submit', 'reset', 'range', 'color', 'file'].includes(type);
    }
    return false;
  }

  function handleScratchGridPasteEvent(e) {
    const cb = e.clipboardData || window.clipboardData;
    if (!cb) return false;
    const text = cb.getData('text/plain');
    const html = cb.getData('text/html') || '';
    const hasHtmlTable = Boolean(html && /<table[\s>]/i.test(html));
    if ((text === undefined || text === null || text === '') && !hasHtmlTable) return false;
    e.preventDefault();

    applyScratchPasteFromText(activeWorkspaceTab, text || '', html);
    return true;
  }

  scratchGridWrap.addEventListener('paste', (e) => {
    if (activeWorkspaceTab === DAILY_TRACKER_WS_TAB || !workbook) return;
    if (scratchInlineEditState) return; // Let inline editor handle single-cell text paste
    handleScratchGridPasteEvent(e);
  });

  // Document-level paste fallback when Scratch Sheet is active but focus slipped onto toolbar/body
  document.addEventListener('paste', (e) => {
    if (e.defaultPrevented) return;
    if (activeWorkspaceTab === DAILY_TRACKER_WS_TAB || !workbook) return;
    if (scratchInlineEditState) return;
    if (typeof document.querySelector === 'function' && document.querySelector('.modal-backdrop.open')) return;
    if (isScratchPasteBlockedByInputEl(e.target) || isScratchPasteBlockedByInputEl(document.activeElement)) return;
    handleScratchGridPasteEvent(e);
  });
}

// Global Ctrl+S & Escape handlers while viewing a Scratch Sheet Tab
document.addEventListener('keydown', (e) => {
  if (e.key === 'Escape') {
    if (scratchAutoFilterModal && scratchAutoFilterModal.classList.contains('open')) {
      e.preventDefault();
      closeScratchAutoFilterModal();
      return;
    }
    if (textToColumnsModal && textToColumnsModal.classList.contains('open')) {
      e.preventDefault();
      closeTextToColumnsModal();
      return;
    }
    if (scratchContextMenu && scratchContextMenu.style.display !== 'none') {
      e.preventDefault();
      closeScratchContextMenu();
      return;
    }
    if (scratchFillDropdown && scratchFillDropdown.classList.contains('open')) {
      e.preventDefault();
      closeScratchFillDropdown();
      return;
    }
  }
  if (
    activeWorkspaceTab !== DAILY_TRACKER_WS_TAB &&
    (e.ctrlKey || e.metaKey) &&
    e.key.toLowerCase() === 's'
  ) {
    e.preventDefault();
    flushPendingScratchEditsToDisk({ reloadAfter: false, silentToast: false });
  }
});

// Closeable / Acknowledgeable Scratch Sheet Formula Warning Persistence
const SCRATCH_FORMULA_NOTICE_ACK_KEY = 'trackerui_scratch_formula_notice_ack_v1';

function syncScratchFormulaNoticeVisibility() {
  if (!scratchFormulaNotice) return;
  const isAcked = localStorage.getItem(SCRATCH_FORMULA_NOTICE_ACK_KEY) === 'true';
  scratchFormulaNotice.style.display = isAcked ? 'none' : 'flex';
}

if (btnAckScratchFormulaNotice) {
  btnAckScratchFormulaNotice.addEventListener('click', () => {
    localStorage.setItem(SCRATCH_FORMULA_NOTICE_ACK_KEY, 'true');
    syncScratchFormulaNoticeVisibility();
    showToast('Scratch sheet formula warning acknowledged and hidden.', 'info');
  });
}

syncScratchFormulaNoticeVisibility();
renderTrackerProfilesBar();
