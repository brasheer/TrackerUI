// ============================================================================
// spreadsheet-engine.js — Pure OpenXML Styling, Value/Date Coercion, Formula
//                         AST & Spill Engine, Fill-Handle Series & Clipboard
// ============================================================================

const SCRATCH_FORMAT_DEFS = {
  auto: { id: 'auto', label: 'Auto', numFmt: null },
  text: { id: 'text', label: 'Text (@)', numFmt: '@' },
  date: { id: 'date', label: 'Date (MM/DD/YYYY)', numFmt: 'mm/dd/yyyy' },
  number: { id: 'number', label: 'Number (123)', numFmt: '0' },
  currency: { id: 'currency', label: 'Currency ($)', numFmt: '$#,##0.00' }
};

const scratchSpillStateMap = new Map(); // sheetName -> Map<'anchorR::anchorC', Array<{ r, c }>>
const scratchSpilledCellsMap = new Map(); // sheetName -> Map<'r::c', { display, rawVal, anchorR, anchorC, formula, isError? }>

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
  const emptyVisual = { bgHex: null, fgHex: null, isBold: false, isExplicitUnbold: false, isItalic: false, isUnderline: false, align: null };
  if (!ws) return emptyVisual;
  const row = ws.findRow ? ws.findRow(rNum) : null;
  if (!row) return emptyVisual;
  const cell = row.findCell ? row.findCell(cNum) : null;
  if (!cell) return emptyVisual;

  let bgHex = null;
  const fill = cell.fill;
  if (fill && fill.type === 'pattern' && fill.pattern === 'solid') {
    bgHex = resolveExcelColorToHex(fill.fgColor) || resolveExcelColorToHex(fill.bgColor);
  }

  const font = cell.font;
  const isBold = Boolean(font && font.bold === true);
  const isExplicitUnbold = Boolean(font && font.bold === false);
  const isItalic = Boolean(font && font.italic);
  const isUnderline = Boolean(font && font.underline);
  const explicitFontHex = font ? resolveExcelColorToHex(font.color) : null;
  const fgHex = bgHex
    ? (explicitFontHex || getContrastingTextHexForBg(bgHex))
    : explicitFontHex;
  const align = (cell.alignment && cell.alignment.horizontal) ? cell.alignment.horizontal : null;
  return { bgHex, fgHex, isBold, isExplicitUnbold, isItalic, isUnderline, align };
}

function normalizeScratchStyleSnapshot(snap) {
  return {
    fill: snap && snap.fill ? JSON.parse(JSON.stringify(snap.fill)) : null,
    font: snap && snap.font && Object.keys(snap.font).length > 0 ? JSON.parse(JSON.stringify(snap.font)) : null,
    numFmt: (snap && typeof snap.numFmt === 'string' && snap.numFmt.trim() !== '') ? snap.numFmt.trim() : null,
    alignment: snap && snap.alignment && Object.keys(snap.alignment).length > 0 ? JSON.parse(JSON.stringify(snap.alignment)) : null
  };
}

function captureCellStyleSnapshot(ws, rNum, cNum) {
  if (!ws) return { fill: null, font: null, numFmt: null, alignment: null };
  const row = ws.findRow ? ws.findRow(rNum) : null;
  const cell = row && row.findCell ? row.findCell(cNum) : null;
  const numFmt = (cell && typeof cell.numFmt === 'string' && cell.numFmt.trim() !== '')
    ? cell.numFmt.trim()
    : ((cell && cell.style && typeof cell.style.numFmt === 'string' && cell.style.numFmt.trim() !== '')
        ? cell.style.numFmt.trim()
        : null);
  if (!cell || !cell.style) return { fill: null, font: null, numFmt, alignment: null };
  return {
    fill: cell.style.fill ? JSON.parse(JSON.stringify(cell.style.fill)) : null,
    font: cell.style.font && Object.keys(cell.style.font).length > 0 ? JSON.parse(JSON.stringify(cell.style.font)) : null,
    numFmt,
    alignment: cell.style.alignment && Object.keys(cell.style.alignment).length > 0 ? JSON.parse(JSON.stringify(cell.style.alignment)) : null
  };
}

function buildStyleSnapshotForFillPreset(beforeSnap, presetId) {
  const baseFont = beforeSnap && beforeSnap.font ? JSON.parse(JSON.stringify(beforeSnap.font)) : null;
  const baseNumFmt = beforeSnap && beforeSnap.numFmt ? beforeSnap.numFmt : null;
  const baseAlign = beforeSnap && beforeSnap.alignment ? JSON.parse(JSON.stringify(beforeSnap.alignment)) : null;
  if (!presetId || presetId === 'none') {
    if (baseFont && baseFont.color) {
      delete baseFont.color;
    }
    const cleanFont = baseFont && Object.keys(baseFont).length > 0 ? baseFont : null;
    return { fill: null, font: cleanFont, numFmt: baseNumFmt, alignment: baseAlign };
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
    numFmt: baseNumFmt,
    alignment: baseAlign
  };
}

function buildStyleSnapshotForTextFormatting(prevSnap, stylePatch = {}) {
  const baseFill = prevSnap && prevSnap.fill ? JSON.parse(JSON.stringify(prevSnap.fill)) : null;
  const baseNumFmt = prevSnap && prevSnap.numFmt ? prevSnap.numFmt : null;
  const nextFont = prevSnap && prevSnap.font ? JSON.parse(JSON.stringify(prevSnap.font)) : {};
  const nextAlign = prevSnap && prevSnap.alignment ? JSON.parse(JSON.stringify(prevSnap.alignment)) : {};
  const forceExplicitUnbold = Boolean(stylePatch && stylePatch.explicitUnbold);

  if (stylePatch && stylePatch.clearFormatting) {
    if (forceExplicitUnbold) {
      nextFont.bold = false;
    } else {
      delete nextFont.bold;
    }
    delete nextFont.italic;
    delete nextFont.underline;
    delete nextAlign.horizontal;
  } else if (stylePatch && typeof stylePatch === 'object') {
    if (stylePatch.bold !== undefined) {
      if (stylePatch.bold) {
        nextFont.bold = true;
      } else if (forceExplicitUnbold) {
        nextFont.bold = false;
      } else {
        delete nextFont.bold;
      }
    }
    if (stylePatch.italic !== undefined) {
      if (stylePatch.italic) nextFont.italic = true;
      else delete nextFont.italic;
    }
    if (stylePatch.underline !== undefined) {
      if (stylePatch.underline) nextFont.underline = true;
      else delete nextFont.underline;
    }
    if (stylePatch.align !== undefined) {
      const normAlign = String(stylePatch.align || '').trim().toLowerCase();
      if (['left', 'center', 'right'].includes(normAlign)) {
        nextAlign.horizontal = normAlign;
      } else {
        delete nextAlign.horizontal;
      }
    }
  }

  return {
    fill: baseFill,
    font: Object.keys(nextFont).length > 0 ? nextFont : null,
    numFmt: baseNumFmt,
    alignment: Object.keys(nextAlign).length > 0 ? nextAlign : null
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

  if (snap && snap.alignment && Object.keys(snap.alignment).length > 0) {
    nextStyle.alignment = JSON.parse(JSON.stringify(snap.alignment));
  } else if (snap && Object.prototype.hasOwnProperty.call(snap, 'alignment')) {
    delete nextStyle.alignment;
  }

  cell.style = nextStyle;
  row.commit();
}

function applyVisualStyleToScratchTd(td, visual) {
  if (!td || !visual) return;
  td.classList.toggle('bold-cell', Boolean(visual.isBold));
  td.classList.toggle('unbold-cell', Boolean(visual.isExplicitUnbold));
  td.classList.toggle('italic-cell', Boolean(visual.isItalic));
  td.classList.toggle('underline-cell', Boolean(visual.isUnderline));
  td.style.fontStyle = visual.isItalic ? 'italic' : '';
  td.style.textDecoration = visual.isUnderline ? 'underline' : '';
  td.style.textAlign = visual.align || '';
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

function extractScratchCellInfo(ws, rNum, cNum, { skipSpillLookup = false, spilledCellsMap = null } = {}) {
  if (!ws) return { display: '', rawEdit: '', isFormula: false, isSpilled: false, formatId: 'auto' };
  const sheetName = ws.name || '';
  const coordKey = `${rNum}::${cNum}`;
  const spillMapRef = spilledCellsMap || scratchSpilledCellsMap;
  const spillMap = !skipSpillLookup && sheetName && spillMapRef ? spillMapRef.get(sheetName) : null;
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

function measureScratchWorksheetBounds(ws, ctx = {}) {
  let maxRow = 0;
  let maxCol = 0;
  if (!ws) return { nonEmptyRows: 0, maxRow: 0, maxCol: 0 };

  const nonEmptyRowSet = new Set();

  ws.eachRow({ includeEmpty: false }, (row, rowNumber) => {
    row.eachCell({ includeEmpty: false }, (_cell, colNumber) => {
      const info = extractScratchCellInfo(ws, rowNumber, colNumber, { spilledCellsMap: ctx && ctx.spilledCellsMap });
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

  const spillMapRef = (ctx && ctx.spilledCellsMap) || scratchSpilledCellsMap;
  const spillMap = ws.name && spillMapRef ? spillMapRef.get(ws.name) : null;
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
  } else if (formatId === 'text') {
    cell.value = normalized;
  } else if (normalized.startsWith("'") && normalized.length > 1) {
    // Leading apostrophe escapes formula / number coercion in Auto mode, matching Excel
    cell.value = normalized.slice(1);
  } else if (normalized.startsWith('=') && normalized.length > 1) {
    cell.value = { formula: toExcelOpenXmlFormula(normalized.slice(1)) };
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
      /^[+-]?(?:0|[1-9]\d*)(?:\.\d+)?$/.test(stripped) &&
      stripped.replace(/[+-.]/g, '').length <= 12;
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

// --- FORMULA TOKENIZER, AST PARSER & DYNAMIC ARRAY SPILL ENGINE ---
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

  let dA = unwrapFormulaDate(a);
  let dB = unwrapFormulaDate(b);
  if (dA && !dB && typeof b === 'string' && b.trim() !== '') {
    dB = parseFlexibleScratchDate(b.trim(), { allowTwoDigitHyphen: true });
  } else if (dB && !dA && typeof a === 'string' && a.trim() !== '') {
    dA = parseFlexibleScratchDate(a.trim(), { allowTwoDigitHyphen: true });
  }
  if (dA && dB) {
    const cmp = dA.dayEpoch - dB.dayEpoch;
    return evalCmpOp(cmp, op);
  }
  if (dA || dB) {
    if (op === '=' || op === '==' || op === '<>' || op === '!=') {
      const sDate = formulaValToString(a).trim().toLowerCase();
      const sOther = formulaValToString(b).trim().toLowerCase();
      if (sDate === sOther) return evalCmpOp(0, op);
      return (op === '<>' || op === '!=') ? true : false;
    }
  }

  const aNull = a === null || a === undefined;
  const bNull = b === null || b === undefined;
  if (aNull && bNull) return evalCmpOp(0, op);
  if (aNull) {
    if (b === '') return evalCmpOp(0, op);
    if (typeof b === 'number') return evalCmpOp(0 - b, op);
    if (typeof b === 'boolean') return evalCmpOp(0 - (b ? 1 : 0), op);
    const sB = formulaValToString(b).toLowerCase();
    return evalCmpOp('' === sB ? 0 : ('' < sB ? -1 : 1), op);
  }
  if (bNull) {
    if (a === '') return evalCmpOp(0, op);
    if (typeof a === 'number') return evalCmpOp(a - 0, op);
    if (typeof a === 'boolean') return evalCmpOp((a ? 1 : 0) - 0, op);
    const sA = formulaValToString(a).toLowerCase();
    return evalCmpOp(sA === '' ? 0 : (sA < '' ? -1 : 1), op);
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

function broadcastTernaryOp(cond, tVal, fVal, fn) {
  const cMat = isFormulaMatrix(cond);
  const tMat = isFormulaMatrix(tVal);
  const fMat = isFormulaMatrix(fVal);
  if (!cMat && !tMat && !fMat) {
    return fn(cond, tVal, fVal);
  }
  const cRows = cMat ? cond.rows : [[cond]];
  const tRows = tMat ? tVal.rows : [[tVal]];
  const fRows = fMat ? fVal.rows : [[fVal]];
  const rowCount = Math.max(cRows.length, tRows.length, fRows.length);
  const colCount = Math.max(
    cRows[0] ? cRows[0].length : 1,
    tRows[0] ? tRows[0].length : 1,
    fRows[0] ? fRows[0].length : 1
  );
  const out = [];
  for (let r = 0; r < rowCount; r++) {
    const cr = cRows[r % cRows.length];
    const tr = tRows[r % tRows.length];
    const fr = fRows[r % fRows.length];
    const rowOut = [];
    for (let c = 0; c < colCount; c++) {
      rowOut.push(fn(cr[c % cr.length], tr[c % tr.length], fr[c % fr.length]));
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
    if (two === '<>' || two === '!=' || two === '<=' || two === '>=' || two === '==') {
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
        return broadcastTernaryOp(cond, tVal, fVal, (cVal, tv, fv) => (formulaValToBool(cVal) ? tv : fv));
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
        const opMatch = critStr.match(/^(<>|!=|<=|>=|==|=|<|>)(.*)$/);
        let cnt = 0;
        vals.forEach(v => {
          if (opMatch) {
            const op = opMatch[1];
            const rhsRaw = opMatch[2];
            if ((op === '=' || op === '==') && rhsRaw === '') {
              if (isEmptyFormulaVal(v)) cnt++;
              return;
            }
            if ((op === '<>' || op === '!=') && rhsRaw === '') {
              if (!isEmptyFormulaVal(v)) cnt++;
              return;
            }
            const rhs = rhsRaw !== '' && /^-?(?:0|[1-9]\d*)(?:\.\d+)?$/.test(rhsRaw.trim())
              ? Number(rhsRaw.trim())
              : rhsRaw;
            if (compareFormulaScalars(v, rhs, op) === true) cnt++;
          } else if (typeof crit === 'string' && (crit.includes('*') || crit.includes('?'))) {
            const escaped = crit.replace(/[.+^${}()|[\]\\]/g, '\\$&').replace(/\*/g, '.*').replace(/\?/g, '.');
            if (new RegExp(`^${escaped}$`, 'i').test(formulaValToString(v))) cnt++;
          } else {
            const sV = formulaValToString(v).trim().toLowerCase();
            const sC = critStr.trim().toLowerCase();
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

function getOriginalExcelSpillSetForSheet(sheetName, ctx = {}) {
  const set = new Set();
  const archiveState = (ctx && ctx.archiveState) || (typeof originalArchiveState !== 'undefined' ? originalArchiveState : null);
  if (!archiveState || !archiveState.arrayFormulasBySheet) {
    return set;
  }
  const cellMap = archiveState.arrayFormulasBySheet.get(String(sheetName || '').toLowerCase());
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

function recalculateScratchWorksheetFormulas(ws, sheetNameOrCtx = '', maybeCtx = {}) {
  let sheetName = '';
  let opts = maybeCtx || {};
  if (sheetNameOrCtx && typeof sheetNameOrCtx === 'object') {
    opts = sheetNameOrCtx;
    sheetName = opts.sheetName || (ws && ws.name) || '';
  } else {
    sheetName = sheetNameOrCtx || (ws && ws.name) || '';
  }
  if (!ws || !sheetName) return [];
  const spilledCellsMapRef = (opts && opts.spilledCellsMap) || scratchSpilledCellsMap;
  const spillStateMapRef = (opts && opts.spillStateMap) || scratchSpillStateMap;
  const archiveStateRef = (opts && opts.archiveState) || (typeof originalArchiveState !== 'undefined' ? originalArchiveState : null);
  const wbRef = (opts && opts.wb) || (typeof workbook !== 'undefined' ? workbook : null);
  const prevSpills = spilledCellsMapRef.get(sheetName) || new Map();
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

  const origSpillSet = getOriginalExcelSpillSetForSheet(sheetName, { archiveState: archiveStateRef });
  if (formulaCoords.length === 0 && prevSpills.size === 0 && origSpillSet.size === 0) {
    return [];
  }

  const ctx = {
    wb: wbRef || { getWorksheet: (name) => (name === sheetName ? ws : null) },
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
  if (archiveStateRef && archiveStateRef.arrayFormulasBySheet) {
    const origCellMap = archiveStateRef.arrayFormulasBySheet.get(String(sheetName || '').toLowerCase());
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

  spilledCellsMapRef.set(sheetName, nextSpills);
  spillStateMapRef.set(sheetName, ctx.anchorSpillsBySheet.get(sheetName) || new Map());

  // Compute all coordinates whose visible display value changed
  const changedSet = new Map();
  const addChanged = (r, c) => {
    changedSet.set(`${r}::${c}`, { r, c });
  };

  formulaCoords.forEach(({ r, c }) => {
    const key = `${r}::${c}`;
    const nextDisp = extractScratchCellInfo(ws, r, c, { spilledCellsMap: spilledCellsMapRef }).display;
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
  { axis = 'vertical', ctrlKey = false, autoIncrement = (typeof scratchAutoIncrementEnabled !== 'undefined' ? scratchAutoIncrementEnabled : false) } = {}
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
            numFmt: baseSnap.numFmt || 'mm/dd/yyyy',
            alignment: baseSnap.alignment || null
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

// --- CLIPBOARD TSV / EXCEL HTML SERIALIZERS & MATRIX PARSERS ---
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

function serializeScratchSelectionToExcelHtml(ws, range, opts = {}) {
  if (!ws || !range) return '';
  const maxScanCol = (opts && opts.maxScanCol) || Math.max(range.maxCol, ws.columnCount || 0, 10);
  let row1HasValues = Boolean(opts && typeof opts.row1HasValues === 'boolean' && opts.row1HasValues);
  if (!opts || typeof opts.row1HasValues !== 'boolean') {
    for (let c = 1; c <= maxScanCol; c++) {
      if (extractScratchCellInfo(ws, 1, c, opts).display.trim() !== '') {
        row1HasValues = true;
        break;
      }
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
      const info = extractScratchCellInfo(ws, r, c, opts);
      const visual = extractScratchCellVisualStyle(ws, r, c);
      const isAutoHeader = r === 1 && row1HasValues && !visual.isExplicitUnbold;

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
      if (visual.isItalic) {
        styles.push('font-style:italic');
      }
      if (visual.isUnderline) {
        styles.push('text-decoration:underline');
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

function resolveColorStyleSnapshotFromProps({
  bgRaw,
  fgRaw,
  weightRaw,
  styleRaw = '',
  decoRaw = '',
  alignRaw = '',
  tagName = 'td',
  hasBoldTag = false,
  hasItalicTag = false,
  hasUnderlineTag = false
}) {
  let bgArgb = parseCssColorToArgb(bgRaw);
  if (bgArgb === 'FFFFFFFF') bgArgb = null;

  let fgArgb = parseCssColorToArgb(fgRaw);
  if (!bgArgb && (fgArgb === 'FF000000' || fgArgb === 'FF0F172A')) {
    fgArgb = null;
  }

  const wLower = String(weightRaw || '').trim().toLowerCase();
  const isBold = hasBoldTag || wLower === 'bold' || wLower === 'bolder' || (parseInt(wLower, 10) >= 600) || (tagName === 'th' && wLower === '');
  const isItalic = hasItalicTag || /\b(?:italic|oblique)\b/i.test(String(styleRaw || ''));
  const isUnderline = hasUnderlineTag || /\bunderline\b/i.test(String(decoRaw || ''));
  const alignLower = String(alignRaw || '').trim().toLowerCase();
  const alignment = ['left', 'center', 'right'].includes(alignLower) ? { horizontal: alignLower } : null;

  const fill = bgArgb
    ? { type: 'pattern', pattern: 'solid', fgColor: { argb: bgArgb } }
    : null;

  const fontObj = {};
  if (fgArgb) fontObj.color = { argb: fgArgb };
  else if (bgArgb) fontObj.color = { argb: getContrastingTextHexForBg(`#${bgArgb.slice(2)}`).replace('#', 'FF') };
  if (isBold) fontObj.bold = true;
  if (isItalic) fontObj.italic = true;
  if (isUnderline) fontObj.underline = true;
  const font = Object.keys(fontObj).length > 0 ? fontObj : null;

  return { fill, font, alignment, hasVisualStyle: Boolean(fill || font || alignment) };
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
            const styleRaw = merged['font-style'] || '';
            const decoRaw = merged['text-decoration'] || merged['text-decoration-line'] || '';
            const alignRaw = merged['text-align'] || cell.getAttribute('align') || '';
            const hasBoldTag = Boolean(cell.querySelector('b, strong'));
            const hasItalicTag = Boolean(cell.querySelector('i, em'));
            const hasUnderlineTag = Boolean(cell.querySelector('u, ins'));
            const snap = resolveColorStyleSnapshotFromProps({
              bgRaw,
              fgRaw,
              weightRaw,
              styleRaw,
              decoRaw,
              alignRaw,
              tagName: cell.tagName.toLowerCase(),
              hasBoldTag,
              hasItalicTag,
              hasUnderlineTag
            });
            if (snap.hasVisualStyle) anyStyle = true;
            const cellText = extractHtmlTableCellText(cell);
            const colspan = Math.min(500, Math.max(1, parseInt(cell.getAttribute('colspan') || '1', 10) || 1));
            const rowspan = Math.min(500, Math.max(1, parseInt(cell.getAttribute('rowspan') || '1', 10) || 1));
            const styleObj = { fill: snap.fill, font: snap.font, alignment: snap.alignment || null };

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
        const readAttr = (name) => {
          const m = attrs.match(new RegExp(`\\b${name}\\s*=\\s*(?:"([^"]*)"|'([^']*)')`, 'i'));
          return m ? (m[1] !== undefined ? m[1] : (m[2] || '')) : '';
        };
        const clsAttr = readAttr('class');
        const styleAttr = readAttr('style');
        const bgcolorAttr = readAttr('bgcolor');
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
        const fgRaw = merged.color || readAttr('color');
        const weightRaw = merged['font-weight'] || '';
        const styleRaw = merged['font-style'] || '';
        const decoRaw = merged['text-decoration'] || merged['text-decoration-line'] || '';
        const alignRaw = merged['text-align'] || readAttr('align');
        const hasBoldTag = /<(b|strong)\b/i.test(inner);
        const hasItalicTag = /<(i|em)\b/i.test(inner);
        const hasUnderlineTag = /<(u|ins)\b/i.test(inner);

        const snap = resolveColorStyleSnapshotFromProps({
          bgRaw,
          fgRaw,
          weightRaw,
          styleRaw,
          decoRaw,
          alignRaw,
          tagName,
          hasBoldTag,
          hasItalicTag,
          hasUnderlineTag
        });
        if (snap.hasVisualStyle) anyStyle = true;
        const cellText = extractHtmlCellTextFallback(inner);
        const styleObj = { fill: snap.fill, font: snap.font, alignment: snap.alignment || null };

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
      if (styleRows[r][c] === undefined) styleRows[r][c] = { fill: null, font: null, alignment: null };
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

// --- DELIMITED LINE TOKENIZER ---
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
