import { normalizePhoneText } from './phone.ts';

export interface ParsedSpreadsheetResult {
  /** Newline-separated phone numbers / chat IDs ready for the bulk recipients textarea. */
  recipientsText: string;
  /** Per-recipient template variables keyed by normalized `<digits>@c.us` (or chat ID). */
  variablesByChatId: Record<string, Record<string, string>>;
  /** Sanitized column names available as `{{variable}}` placeholders. */
  detectedColumns: string[];
  /** Raw parsed 2D rows from the spreadsheet, so re-toggling +91 prefix can re-key variables. */
  rows: string[][];
}

const MIN_PHONE_DIGITS = 6;
const MAX_PHONE_DIGITS = 15;
const CHAT_ID_PATTERN = /^[0-9][0-9:-]*@(c\.us|g\.us|s\.whatsapp\.net|lid|newsletter|broadcast)$/i;
const PHONE_HEADER_HINTS = /^(phone|mobile|number|whatsapp|wa|contact|tel|telephone|cell|msisdn|phone_number|mob|ph)$/i;

/**
 * Format a raw cell string into a normalized `<digits>@c.us` or chat ID, or null if not a valid recipient.
 */
export function normalizeRecipientCell(rawCell: string, autoPrefix91 = false): { chatId: string; displayPhone: string } | null {
  const trimmed = normalizePhoneText(rawCell).trim().replace(/^["']|["']$/g, '').trim();
  if (!trimmed) return null;
  if (trimmed.includes('@')) {
    return CHAT_ID_PATTERN.test(trimmed) ? { chatId: trimmed, displayPhone: trimmed } : null;
  }
  // Handle Excel scientific notation for 10-15 digit phone numbers (e.g. "9.1987654321E+11")
  let candidate = trimmed;
  if (/^[0-9]+(\.[0-9]+)?e\+[0-9]+$/i.test(candidate)) {
    const num = Number(candidate);
    if (Number.isFinite(num) && num > 0) {
      candidate = Math.round(num).toFixed(0);
    }
  }
  let digits = candidate.replace(/[^0-9]/g, '');
  const isInternational = trimmed.startsWith('+') || digits.startsWith('00');
  if (digits.startsWith('00') && digits.length >= 12) {
    digits = digits.substring(2);
  }
  if (autoPrefix91 && digits.length === 13 && digits.startsWith('910')) {
    digits = '91' + digits.substring(3);
  }
  if (autoPrefix91 && !isInternational && digits.length === 11 && digits.startsWith('0')) {
    digits = '91' + digits.substring(1);
  } else if (autoPrefix91 && !isInternational && digits.length === 10) {
    digits = '91' + digits;
  }
  if (digits.length < MIN_PHONE_DIGITS || digits.length > MAX_PHONE_DIGITS) return null;
  return { chatId: `${digits}@c.us`, displayPhone: digits };
}

/**
 * Sanitize a header cell into a template placeholder identifier (e.g. "First Name" -> "first_name").
 */
function sanitizeColumnName(raw: string, fallbackIndex: number): string {
  const cleaned = raw
    .trim()
    .toLowerCase()
    .replace(/[^a-z0-9_]+/g, '_')
    .replace(/^_+|_+$/g, '');
  return cleaned || `col${fallbackIndex + 1}`;
}

/**
 * Extract recipients and per-recipient template variables from 2D spreadsheet rows.
 */
export function extractContactsFromRows(rows: string[][], autoPrefix91 = false): ParsedSpreadsheetResult {
  const nonEmptyRows = rows.filter(r => r.some(cell => cell.trim() !== ''));
  if (nonEmptyRows.length === 0) {
    return { recipientsText: '', variablesByChatId: {}, detectedColumns: [], rows: [] };
  }

  const maxCols = nonEmptyRows.reduce((max, r) => Math.max(max, r.length), 0);

  // Determine which column holds phone numbers by scoring each column across up to the first 50 rows.
  let bestPhoneCol = -1;
  let bestPhoneHits = 0;
  const sampleLimit = Math.min(nonEmptyRows.length, 50);

  for (let col = 0; col < maxCols; col++) {
    let hits = 0;
    const headerVal = (nonEmptyRows[0]?.[col] ?? '').trim();
    const headerBonus = PHONE_HEADER_HINTS.test(headerVal.replace(/[^a-zA-Z_]/g, '')) ? 5 : 0;
    for (let rowIdx = 0; rowIdx < sampleLimit; rowIdx++) {
      const cell = nonEmptyRows[rowIdx]?.[col] ?? '';
      if (normalizeRecipientCell(cell, autoPrefix91)) hits++;
    }
    if (hits > 0 && hits + headerBonus > bestPhoneHits) {
      bestPhoneHits = hits + headerBonus;
      bestPhoneCol = col;
    }
  }

  // If only 1 column or no clear structured table, still check if row 0 is a header.
  const firstRowPhoneCell = bestPhoneCol >= 0 ? (nonEmptyRows[0]?.[bestPhoneCol] ?? '') : (nonEmptyRows[0]?.[0] ?? '');
  const hasHeaderRow = nonEmptyRows.length > 1 && normalizeRecipientCell(firstRowPhoneCell, autoPrefix91) === null;

  const columnNames: string[] = [];
  if (maxCols > 1 && bestPhoneCol >= 0) {
    const usedNames = new Set<string>();
    for (let col = 0; col < maxCols; col++) {
      const rawName = hasHeaderRow ? (nonEmptyRows[0]?.[col] ?? '') : `col${col + 1}`;
      let name = sanitizeColumnName(rawName, col);
      if (usedNames.has(name)) name = `${name}_${col + 1}`;
      usedNames.add(name);
      columnNames.push(name);
    }
  }

  const startRow = hasHeaderRow ? 1 : 0;
  const displayPhones: string[] = [];
  const seenChatIds = new Set<string>();
  const variablesByChatId: Record<string, Record<string, string>> = {};

  for (let r = startRow; r < nonEmptyRows.length; r++) {
    const row = nonEmptyRows[r];
    if (bestPhoneCol >= 0 && maxCols > 1) {
      const parsed = normalizeRecipientCell(row[bestPhoneCol] ?? '', autoPrefix91);
      if (!parsed || seenChatIds.has(parsed.chatId)) continue;
      seenChatIds.add(parsed.chatId);
      displayPhones.push(parsed.displayPhone);

      const vars: Record<string, string> = {};
      for (let col = 0; col < maxCols; col++) {
        const val = (row[col] ?? '').trim();
        const colName = columnNames[col];
        if (colName) vars[colName] = val;
        vars[`col${col + 1}`] = val;
      }
      variablesByChatId[parsed.chatId] = vars;
    } else {
      // Unstructured or single-column sheet: scan every cell in the row for phone numbers
      for (const cell of row) {
        const parsed = normalizeRecipientCell(cell, autoPrefix91);
        if (!parsed || seenChatIds.has(parsed.chatId)) continue;
        seenChatIds.add(parsed.chatId);
        displayPhones.push(parsed.displayPhone);
      }
    }
  }

  const detectedColumns = columnNames.filter((_, idx) => idx !== bestPhoneCol);

  return {
    recipientsText: displayPhones.join('\n'),
    variablesByChatId,
    detectedColumns,
    rows: nonEmptyRows,
  };
}

/**
 * Parse CSV/TSV text into a 2D array of rows and cells (handling quoted fields).
 */
export function parseCsvRows(text: string): string[][] {
  const lines = text.split(/\r\n?|\n/);
  const rows: string[][] = [];
  for (const line of lines) {
    if (!line.trim()) continue;
    const delimiter = line.includes('\t') ? '\t' : line.includes(';') && !line.includes(',') ? ';' : ',';
    const cells: string[] = [];
    let current = '';
    let inQuotes = false;
    for (let i = 0; i < line.length; i++) {
      const ch = line[i];
      if (ch === '"') {
        if (inQuotes && line[i + 1] === '"') {
          current += '"';
          i++;
        } else {
          inQuotes = !inQuotes;
        }
      } else if (ch === delimiter && !inQuotes) {
        cells.push(current.trim());
        current = '';
      } else {
        current += ch;
      }
    }
    cells.push(current.trim());
    rows.push(cells);
  }
  return rows;
}

/**
 * Convert an Excel column reference like "A1", "B2", "AA10" to a 0-based column index.
 */
function excelColToIndex(cellRef: string): number {
  const match = cellRef.match(/^[A-Za-z]+/);
  if (!match) return -1;
  const letters = match[0].toUpperCase();
  let col = 0;
  for (let i = 0; i < letters.length; i++) {
    col = col * 26 + (letters.charCodeAt(i) - 64);
  }
  return col - 1;
}

/**
 * Minimal ZIP reader for `.xlsx` files using the standard Web `DecompressionStream('deflate-raw')`.
 */
async function unzipXlsxEntries(buffer: ArrayBuffer): Promise<Map<string, string>> {
  const bytes = new Uint8Array(buffer);
  const view = new DataView(buffer);
  const entries = new Map<string, string>();
  const decoder = new TextDecoder('utf-8');

  // Locate End of Central Directory (EOCD) signature 0x06054b50 from the end of the file
  let eocdOffset = -1;
  const minEocdOffset = Math.max(0, bytes.length - 65557);
  for (let i = bytes.length - 22; i >= minEocdOffset; i--) {
    if (view.getUint32(i, true) === 0x06054b50) {
      eocdOffset = i;
      break;
    }
  }
  if (eocdOffset < 0) {
    throw new Error('Invalid Excel (.xlsx) archive');
  }

  const totalEntries = view.getUint16(eocdOffset + 10, true);
  const cdOffset = view.getUint32(eocdOffset + 16, true);
  let ptr = cdOffset;

  for (let i = 0; i < totalEntries && ptr + 46 <= bytes.length; i++) {
    if (view.getUint32(ptr, true) !== 0x02014b50) break;
    const compressionMethod = view.getUint16(ptr + 10, true);
    const compressedSize = view.getUint32(ptr + 20, true);
    const fileNameLen = view.getUint16(ptr + 28, true);
    const extraLen = view.getUint16(ptr + 30, true);
    const commentLen = view.getUint16(ptr + 32, true);
    const localHeaderOffset = view.getUint32(ptr + 42, true);

    const fileName = decoder.decode(bytes.subarray(ptr + 46, ptr + 46 + fileNameLen));
    ptr += 46 + fileNameLen + extraLen + commentLen;

    if (
      fileName !== 'xl/sharedStrings.xml' &&
      !/^xl\/worksheets\/sheet[0-9]+\.xml$/i.test(fileName)
    ) {
      continue;
    }

    if (localHeaderOffset + 30 > bytes.length || view.getUint32(localHeaderOffset, true) !== 0x04034b50) {
      continue;
    }
    const localNameLen = view.getUint16(localHeaderOffset + 26, true);
    const localExtraLen = view.getUint16(localHeaderOffset + 28, true);
    const dataStart = localHeaderOffset + 30 + localNameLen + localExtraLen;
    const compressedSlice = bytes.subarray(dataStart, dataStart + compressedSize);

    if (compressionMethod === 0) {
      entries.set(fileName, decoder.decode(compressedSlice));
    } else if (compressionMethod === 8) {
      const ds = new DecompressionStream('deflate-raw');
      const writer = ds.writable.getWriter();
      writer.write(compressedSlice).catch(() => {});
      writer.close().catch(() => {});
      const decompressedBuffer = await new Response(ds.readable).arrayBuffer();
      entries.set(fileName, decoder.decode(decompressedBuffer));
    }
  }

  return entries;
}

/**
 * Parse an `.xlsx` ArrayBuffer into 2D string rows across all worksheets.
 */
export async function parseXlsxRows(buffer: ArrayBuffer): Promise<string[][]> {
  const entries = await unzipXlsxEntries(buffer);
  const parser = new DOMParser();

  const sharedStrings: string[] = [];
  const sharedStringsXml = entries.get('xl/sharedStrings.xml');
  if (sharedStringsXml) {
    const doc = parser.parseFromString(sharedStringsXml, 'application/xml');
    const siNodes = doc.getElementsByTagName('si');
    for (let i = 0; i < siNodes.length; i++) {
      const tNodes = siNodes[i].getElementsByTagName('t');
      let text = '';
      for (let j = 0; j < tNodes.length; j++) {
        text += tNodes[j].textContent ?? '';
      }
      sharedStrings.push(text);
    }
  }

  const sheetNames = [...entries.keys()]
    .filter(k => /^xl\/worksheets\/sheet[0-9]+\.xml$/i.test(k))
    .sort((a, b) => a.localeCompare(b, undefined, { numeric: true }));

  const allRows: string[][] = [];
  for (const sheetName of sheetNames) {
    const xml = entries.get(sheetName);
    if (!xml) continue;
    const doc = parser.parseFromString(xml, 'application/xml');
    const rowNodes = doc.getElementsByTagName('row');
    for (let i = 0; i < rowNodes.length; i++) {
      const cellNodes = rowNodes[i].getElementsByTagName('c');
      const rowCells: string[] = [];
      for (let j = 0; j < cellNodes.length; j++) {
        const cell = cellNodes[j];
        const ref = cell.getAttribute('r') ?? '';
        const colIdx = ref ? excelColToIndex(ref) : j;
        const type = cell.getAttribute('t') ?? '';
        let val = '';
        if (type === 'inlineStr') {
          const tNodes = cell.getElementsByTagName('t');
          for (let k = 0; k < tNodes.length; k++) {
            val += tNodes[k].textContent ?? '';
          }
        } else {
          const vNode = cell.getElementsByTagName('v')[0];
          const rawVal = vNode?.textContent ?? '';
          if (type === 's') {
            const idx = parseInt(rawVal, 10);
            val = !Number.isNaN(idx) && sharedStrings[idx] !== undefined ? sharedStrings[idx] : rawVal;
          } else {
            val = rawVal;
          }
        }
        if (colIdx >= 0) {
          while (rowCells.length <= colIdx) rowCells.push('');
          rowCells[colIdx] = val.trim();
        }
      }
      if (rowCells.some(c => c !== '')) {
        allRows.push(rowCells);
      }
    }
  }

  return allRows;
}

/**
 * Parse an uploaded `.xlsx`, `.csv`, or `.txt` file into recipients + template variables.
 */
export async function parseSpreadsheetFile(file: File, autoPrefix91 = false): Promise<ParsedSpreadsheetResult> {
  const isXlsx =
    file.name.toLowerCase().endsWith('.xlsx') ||
    file.type === 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';

  if (isXlsx) {
    const buffer = await file.arrayBuffer();
    const rows = await parseXlsxRows(buffer);
    return extractContactsFromRows(rows, autoPrefix91);
  }

  const text = await file.text();
  const rows = parseCsvRows(text);
  return extractContactsFromRows(rows, autoPrefix91);
}

