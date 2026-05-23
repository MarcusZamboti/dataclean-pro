/**
 * Data Cleaner Engine
 * Processes CSV/XLSX files to ensure data quality:
 * - Error detection and correction
 * - Duplicate removal
 * - Missing value filling
 * - Standardization and organization
 */

import Papa from 'papaparse';
import * as XLSX from 'xlsx';
import { v4 as uuidv4 } from 'uuid';
import fs from 'fs';
import path from 'path';
import os from 'os';

export interface CleaningLog {
  type: 'error_fix' | 'duplicate_removed' | 'missing_filled' | 'standardization' | 'column_split' | 'info';
  column?: string;
  row?: number;
  original?: string;
  corrected?: string;
  description: string;
}

export interface CleaningConfig {
  standardizeColumns: boolean;
  splitColumns: boolean;
  fixErrors: boolean;
  removeDuplicates: boolean;
  fillMissing: boolean;
  missingStrategy: 'smart' | 'none';
}


export interface CleaningResult {
  id: string;
  originalFileName: string;
  originalFormat: 'csv' | 'xlsx';
  rowCount: number;
  columnCount: number;
  cleanedRowCount: number;
  cleanedColumnCount: number;
  columns: string[];
  originalPreview: Record<string, unknown>[];
  cleanedPreview: Record<string, unknown>[];
  logs: CleaningLog[];
  stats: {
    errorsFixed: number;
    duplicatesRemoved: number;
    missingValuesFilled: number;
    columnsStandardized: number;
    columnsSplit: number;
  };
  downloadUrl: string;
  cleanedData: Record<string, unknown>[];
}

// ============ HELPER FUNCTIONS ============

function normalizeColumnName(name: string): string {
  return name
    .trim()
    .toLowerCase()
    .replace(/[áàãâä]/g, 'a')
    .replace(/[éèêë]/g, 'e')
    .replace(/[íìîï]/g, 'i')
    .replace(/[óòõôö]/g, 'o')
    .replace(/[úùûü]/g, 'u')
    .replace(/[ç]/g, 'c')
    .replace(/[ñ]/g, 'n')
    .replace(/[^a-z0-9_]/g, '_')
    .replace(/_+/g, '_')
    .replace(/^_|_$/g, '');
}

function removeSpecialChars(value: string): string {
  // Remove zero-width chars, BOM, and other invisible characters
  return value
    .replace(/[\u200B-\u200D\uFEFF\u00AD]/g, '')
    .replace(/\u00A0/g, ' ') // non-breaking space -> regular space
    .trim();
}

function isDateLike(value: string): boolean {
  const v = value.trim();
  // Common date patterns: DD/MM/YYYY, YYYY-MM-DD, DD-MM-YYYY, etc.
  const datePatterns = [
    /^\d{1,2}[\/\-\.]\d{1,2}[\/\-\.]\d{2,4}$/,
    /^\d{4}[\/\-\.]\d{1,2}[\/\-\.]\d{1,2}$/,
    /^\d{1,2}\s+(jan|fev|mar|abr|mai|jun|jul|ago|set|out|nov|dez|jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)\w*\s+\d{2,4}$/i,
  ];
  if (datePatterns.some(p => p.test(v))) return true;

  // Accept Excel serial numbers as dates too (commonly 5 digits between 30000 and 65000)
  if (/^\d{5}$/.test(v)) {
    const num = parseInt(v);
    return num >= 30000 && num <= 65000;
  }
  return false;
}

function isNumericLike(value: string): boolean {
  const cleaned = value.trim().replace(/[.,\s]/g, '');
  return /^\d+$/.test(cleaned) && cleaned.length > 0;
}

function parseDate(value: string): string | null {
  const v = value.trim();

  // 1. Try Excel Serial Number (e.g., "43831" -> 01/01/2020)
  if (/^\d{5}$/.test(v)) {
    const serial = parseInt(v);
    if (serial >= 30000 && serial <= 65000) {
      const utc_days = Math.floor(serial - 25569);
      const utc_value = utc_days * 86400;
      const date_info = new Date(utc_value * 1000);
      const timezoneOffset = date_info.getTimezoneOffset() * 60 * 1000;
      const date = new Date(date_info.getTime() + timezoneOffset);

      if (!isNaN(date.getTime())) {
        const day = String(date.getDate()).padStart(2, '0');
        const month = String(date.getMonth() + 1).padStart(2, '0');
        const year = date.getFullYear();
        return `${day}/${month}/${year}`;
      }
    }
  }

  // 2. Try ISO or YYYY-MM-DD or YYYY/MM/DD (American standard 1)
  let match = v.match(/^(\d{4})[\/\-\.](\d{1,2})[\/\-\.](\d{1,2})(?:\s.*)?$/);
  if (match) {
    const [, year, month, day] = match;
    const date = new Date(`${year}-${month.padStart(2, '0')}-${day.padStart(2, '0')}`);
    if (!isNaN(date.getTime())) {
      // Return Brazilian format DD/MM/YYYY
      return `${day.padStart(2, '0')}/${month.padStart(2, '0')}/${year}`;
    }
  }

  // 2. Try DD/MM/YYYY or MM/DD/YYYY (delimiters / - .)
  match = v.match(/^(\d{1,2})[\/\-\.](\d{1,2})[\/\-\.](\d{2,4})(?:\s.*)?$/);
  if (match) {
    const [, term1, term2, yearTerm] = match;
    const y = yearTerm.length === 2 ? (parseInt(yearTerm) > 50 ? `19${yearTerm}` : `20${yearTerm}`) : yearTerm;
    
    const t1 = parseInt(term1);
    const t2 = parseInt(term2);

    // If first term is > 12, it must be the day, so it's Brazilian format DD/MM/YYYY
    if (t1 > 12) {
      const date = new Date(`${y}-${term2.padStart(2, '0')}-${term1.padStart(2, '0')}`);
      if (!isNaN(date.getTime())) {
        return `${term1.padStart(2, '0')}/${term2.padStart(2, '0')}/${y}`;
      }
    }
    // If second term is > 12, it must be the day, so it's American format MM/DD/YYYY
    else if (t2 > 12) {
      const date = new Date(`${y}-${term1.padStart(2, '0')}-${term2.padStart(2, '0')}`);
      if (!isNaN(date.getTime())) {
        // Convert to Brazilian format DD/MM/YYYY
        return `${term2.padStart(2, '0')}/${term1.padStart(2, '0')}/${y}`;
      }
    }
    // If both terms are <= 12, assume DD/MM/YYYY as standard Brazilian
    else {
      const date = new Date(`${y}-${term2.padStart(2, '0')}-${term1.padStart(2, '0')}`);
      if (!isNaN(date.getTime())) {
        return `${term1.padStart(2, '0')}/${term2.padStart(2, '0')}/${y}`;
      }
    }
  }

  // 3. Try standard Date parsing for other text formats
  const ts = Date.parse(v);
  if (!isNaN(ts)) {
    const date = new Date(ts);
    const day = String(date.getDate()).padStart(2, '0');
    const month = String(date.getMonth() + 1).padStart(2, '0');
    const year = date.getFullYear();
    return `${day}/${month}/${year}`;
  }

  return null;
}

function parseNumber(value: string): number | null {
  const v = value.trim();
  // Handle Brazilian number format: 1.234,56 -> 1234.56
  if (/^\d{1,3}(\.\d{3})*,\d+$/.test(v)) {
    return parseFloat(v.replace(/\./g, '').replace(',', '.'));
  }
  // Handle US format: 1,234.56 -> 1234.56
  if (/^\d{1,3}(,\d{3})*\.\d+$/.test(v)) {
    return parseFloat(v.replace(/,/g, ''));
  }
  // Simple number
  const num = parseFloat(v.replace(',', '.'));
  return isNaN(num) ? null : num;
}

function formatNumber(value: number): string {
  return value.toString();
}

function inferColumnType(values: unknown[]): 'date' | 'number' | 'text' {
  let dateCount = 0;
  let numberCount = 0;
  let textCount = 0;

  for (const val of values) {
    if (val === null || val === undefined || val === '') continue;
    const str = String(val).trim();
    if (!str) continue;

    if (isDateLike(str)) dateCount++;
    else if (isNumericLike(str)) numberCount++;
    else textCount++;
  }

  const total = dateCount + numberCount + textCount;
  if (total === 0) return 'text';

  if (dateCount / total > 0.6) return 'date';
  if (numberCount / total > 0.6) return 'number';
  return 'text';
}

// ============ COLUMN SPLITTING ============

/**
 * Detects columns where values consistently contain a separator (like /)
 * and splits them into two separate columns.
 *
 * Detection criteria:
 * - At least 40% of non-empty values in the column contain the separator
 * - Each value with the separator has exactly 2 parts when split
 * - The separator is NOT used in a date context (e.g., DD/MM/YYYY)
 */
function splitCompositeColumns(
  data: Record<string, unknown>[],
  logs: CleaningLog[]
): Record<string, unknown>[] {
  if (data.length === 0) return data;

  const columns = Object.keys(data[0]);
  const separators = ['/', '|', ';'];
  const columnsToRemove: Set<string> = new Set();
  const columnsToAdd: Record<string, { col1: string; col2: string; part1s: (string | null)[]; part2s: (string | null)[] }> = {};

  for (const col of columns) {
    const values = data.map(row => row[col]);
    const nonEmptyValues = values.filter(v => v !== null && v !== undefined && v !== '');

    if (nonEmptyValues.length === 0) continue;

    for (const sep of separators) {
      // Check how many values contain this separator
      const valuesWithSep = nonEmptyValues.filter(v => {
        const str = String(v).trim();
        return str.includes(sep);
      });

      const ratio = valuesWithSep.length / nonEmptyValues.length;

      // Need at least 40% of values to contain the separator
      if (ratio < 0.4) continue;

      // Check that split produces exactly 2 parts for most values
      let validSplitCount = 0;
      const part1s: (string | null)[] = [];
      const part2s: (string | null)[] = [];

      for (const val of values) {
        if (val === null || val === undefined || val === '') {
          part1s.push(null);
          part2s.push(null);
          continue;
        }

        const str = String(val).trim();

        if (!str.includes(sep)) {
          part1s.push(str);
          part2s.push(null);
          continue;
        }

        const parts = str.split(sep);
        if (parts.length === 2) {
          validSplitCount++;
          part1s.push(parts[0].trim());
          part2s.push(parts[1].trim());
        } else if (parts.length > 2) {
          // For more than 2 parts, check if it looks like a date (3 parts with numbers)
          const allNumeric = parts.every(p => /^\d+$/.test(p.trim()) || /^\d+[.,]\d+$/.test(p.trim()));
          if (parts.length === 3 && allNumeric) {
            // Looks like a date (DD/MM/YYYY), don't split
            part1s.push(str);
            part2s.push(null);
          } else {
            // Split first part and join the rest
            validSplitCount++;
            part1s.push(parts[0].trim());
            part2s.push(parts.slice(1).join(sep).trim());
          }
        } else {
          part1s.push(str);
          part2s.push(null);
        }
      }

      const splitRatio = validSplitCount / nonEmptyValues.length;
      if (splitRatio < 0.4) continue;

      // Skip if this looks like a date column (DD/MM/YYYY pattern)
      const datePatternCount = nonEmptyValues.filter(v => {
        const str = String(v).trim();
        return /^\d{1,2}[\/.-]\d{1,2}[\/.-]\d{2,4}$/.test(str) || /^\d{4}[\/.-]\d{1,2}[\/.-]\d{1,2}$/.test(str);
      }).length;

      if (datePatternCount / nonEmptyValues.length > 0.5) continue;

      // Generate names for new columns
      const baseName = col;
      let col1Name: string;
      let col2Name: string;

      // Try to infer names from the column name itself
      // e.g., "preco_original_preco_com_desconto" -> split by common patterns
      const nameParts = baseName.split(/_+/);

      // Check if the column name contains duplicated parts suggesting two columns
      // e.g., "preco_original_preco_com_desconto" has "preco" repeated
      const halfLength = Math.floor(nameParts.length / 2);
      if (nameParts.length >= 4) {
        // Try to find a natural split point
        col1Name = nameParts.slice(0, halfLength).join('_');
        col2Name = nameParts.slice(halfLength).join('_');
      } else if (nameParts.length === 2) {
        col1Name = `${nameParts[0]}_1`;
        col2Name = `${nameParts[1]}_2`;
      } else {
        col1Name = `${baseName}_1`;
        col2Name = `${baseName}_2`;
      }

      columnsToRemove.add(col);
      columnsToAdd[col] = { col1: col1Name, col2: col2Name, part1s, part2s };

      logs.push({
        type: 'column_split',
        column: col,
        original: col,
        corrected: `${col1Name} | ${col2Name}`,
        description: `Coluna "${col}" com dados compostos (separador "${sep}") dividida em duas novas colunas: "${col1Name}" e "${col2Name}"`,
      });

      break; // Only use the first matching separator per column
    }
  }

  // Apply column splits
  if (columnsToRemove.size > 0) {
    return data.map((row, idx) => {
      const newRow: Record<string, unknown> = {};

      for (const col of columns) {
        if (columnsToRemove.has(col)) {
          const split = columnsToAdd[col];
          newRow[split.col1] = split.part1s[idx] ?? '';
          newRow[split.col2] = split.part2s[idx] ?? '';
        } else {
          newRow[col] = row[col];
        }
      }

      return newRow;
    });
  }

  return data;
}

// ============ MAIN PROCESSING FUNCTIONS ============

function standardizeColumns(
  data: Record<string, unknown>[],
  logs: CleaningLog[]
): { data: Record<string, unknown>[]; columnMap: Record<string, string> } {
  if (data.length === 0) return { data, columnMap: {} };

  const originalColumns = Object.keys(data[0]);
  const columnMap: Record<string, string> = {};
  const seenNames: Record<string, number> = {};

  for (const col of originalColumns) {
    let newName = normalizeColumnName(col);

    // Handle duplicate column names after normalization
    if (seenNames[newName] !== undefined) {
      seenNames[newName]++;
      newName = `${newName}_${seenNames[newName]}`;
    } else {
      seenNames[newName] = 0;
    }

    columnMap[col] = newName;

    if (col !== newName) {
      logs.push({
        type: 'standardization',
        column: col,
        original: col,
        corrected: newName,
        description: `Coluna "${col}" renomeada para "${newName}" (padronização de nome)`,
      });
    }
  }

  const standardized = data.map(row => {
    const newRow: Record<string, unknown> = {};
    for (const [oldCol, newCol] of Object.entries(columnMap)) {
      newRow[newCol] = row[oldCol];
    }
    return newRow;
  });

  return { data: standardized, columnMap };
}

function detectAndFixErrors(
  data: Record<string, unknown>[],
  logs: CleaningLog[]
): Record<string, unknown>[] {
  if (data.length === 0) return data;

  const columns = Object.keys(data[0]);

  // Infer column types
  const columnTypes: Record<string, 'date' | 'number' | 'text'> = {};
  for (const col of columns) {
    columnTypes[col] = inferColumnType(data.map(row => row[col]));
  }

  return data.map((row, rowIndex) => {
    const newRow = { ...row };

    for (const col of columns) {
      const val = row[col];
      if (val === null || val === undefined || val === '') continue;

      const str = String(val);

      // Remove invisible/special characters
      const cleaned = removeSpecialChars(str);
      if (cleaned !== str) {
        logs.push({
          type: 'error_fix',
          column: col,
          row: rowIndex + 1,
          original: str,
          corrected: cleaned,
          description: `Caracteres especiais/invisíveis removidos na linha ${rowIndex + 1}, coluna "${col}"`,
        });
        newRow[col] = cleaned;
      }

      // Type-specific corrections
      const workingVal = String(newRow[col]);

      if (columnTypes[col] === 'number') {
        const num = parseNumber(workingVal);
        if (num !== null) {
          const formatted = formatNumber(num);
          if (formatted !== workingVal.trim()) {
            logs.push({
              type: 'error_fix',
              column: col,
              row: rowIndex + 1,
              original: workingVal,
              corrected: formatted,
              description: `Número formatado na linha ${rowIndex + 1}, coluna "${col}": "${workingVal}" → "${formatted}"`,
            });
            newRow[col] = formatted;
          }
        }
      }

      if (columnTypes[col] === 'date') {
        const dateStr = parseDate(workingVal);
        if (dateStr) {
          if (dateStr !== workingVal.trim()) {
            logs.push({
              type: 'error_fix',
              column: col,
              row: rowIndex + 1,
              original: workingVal,
              corrected: dateStr,
              description: `Data padronizada na linha ${rowIndex + 1}, coluna "${col}": "${workingVal}" → "${dateStr}" (formato brasileiro)`,
            });
            newRow[col] = dateStr;
          }
        }
      }

      // Text normalization: capitalize first letter, trim whitespace, and fix inverted names
      if (columnTypes[col] === 'text') {
        let textVal = workingVal.trim().replace(/\s+/g, ' ');
        
        // Detect and fix inverted names "LastName, FirstName" -> "FirstName LastName"
        if (/^([^,]+),\s*([^,]+)$/.test(textVal)) {
          const match = textVal.match(/^([^,]+),\s*([^,]+)$/);
          if (match) {
            const [, lastName, firstName] = match;
            // Additional safety check: ensure both parts look like names (letters, accents, hyphens, spaces)
            const nameRegex = /^[A-Za-zÀ-ÖØ-öø-ÿ\s'-]+$/;
            if (nameRegex.test(lastName.trim()) && nameRegex.test(firstName.trim())) {
              const correctedName = `${firstName.trim()} ${lastName.trim()}`;
              logs.push({
                type: 'error_fix',
                column: col,
                row: rowIndex + 1,
                original: workingVal,
                corrected: correctedName,
                description: `Nome invertido "Sobrenome, Nome" corrigido na linha ${rowIndex + 1}, coluna "${col}": "${workingVal}" → "${correctedName}"`,
              });
              textVal = correctedName;
              newRow[col] = correctedName;
            }
          }
        }

        if (textVal !== workingVal) {
          if (newRow[col] === workingVal) {
            logs.push({
              type: 'error_fix',
              column: col,
              row: rowIndex + 1,
              original: workingVal,
              corrected: textVal,
              description: `Espaços desnecessários removidos na linha ${rowIndex + 1}, coluna "${col}"`,
            });
            newRow[col] = textVal;
          }
        }
      }
    }

    return newRow;
  });
}

function removeDuplicates(
  data: Record<string, unknown>[],
  logs: CleaningLog[]
): Record<string, unknown>[] {
  const seen = new Set<string>();
  const result: Record<string, unknown>[] = [];

  for (let i = 0; i < data.length; i++) {
    const rowKey = JSON.stringify(data[i]);
    if (seen.has(rowKey)) {
      logs.push({
        type: 'duplicate_removed',
        row: i + 1,
        description: `Linha duplicada removida (linha ${i + 1}, mantendo primeira ocorrência)`,
      });
    } else {
      seen.add(rowKey);
      result.push(data[i]);
    }
  }

  return result;
}

function fillMissingValues(
  data: Record<string, unknown>[],
  logs: CleaningLog[]
): Record<string, unknown>[] {
  if (data.length === 0) return data;

  const columns = Object.keys(data[0]);

  for (const col of columns) {
    const values = data.map(row => row[col]);
    const nonEmptyValues = values.filter(v => v !== null && v !== undefined && v !== '');
    const emptyIndices = values
      .map((v, i) => (v === null || v === undefined || v === '' ? i : -1))
      .filter(i => i !== -1);

    if (emptyIndices.length === 0) continue;

    const colType = inferColumnType(values);

    if (colType === 'number') {
      // Use median for numeric columns
      const numValues = nonEmptyValues
        .map(v => parseNumber(String(v)))
        .filter((v): v is number => v !== null)
        .sort((a, b) => a - b);

      if (numValues.length > 0) {
        const median = numValues[Math.floor(numValues.length / 2)];
        for (const idx of emptyIndices) {
          logs.push({
            type: 'missing_filled',
            column: col,
            row: idx + 1,
            original: '(vazio)',
            corrected: String(median),
            description: `Campo vazio preenchido com mediana (${median}) na linha ${idx + 1}, coluna "${col}"`,
          });
          data[idx][col] = String(median);
        }
      }
    } else if (colType === 'date') {
      // Forward fill for dates
      for (const idx of emptyIndices) {
        let prevVal: string | null = null;
        for (let j = idx - 1; j >= 0; j--) {
          if (data[j][col] !== null && data[j][col] !== undefined && data[j][col] !== '') {
            prevVal = String(data[j][col]);
            break;
          }
        }
        if (prevVal) {
          logs.push({
            type: 'missing_filled',
            column: col,
            row: idx + 1,
            original: '(vazio)',
            corrected: prevVal,
            description: `Campo vazio preenchido com valor anterior válido na linha ${idx + 1}, coluna "${col}"`,
          });
          data[idx][col] = prevVal;
        } else {
          // Try next value
          let nextVal: string | null = null;
          for (let j = idx + 1; j < data.length; j++) {
            if (data[j][col] !== null && data[j][col] !== undefined && data[j][col] !== '') {
              nextVal = String(data[j][col]);
              break;
            }
          }
          if (nextVal) {
            logs.push({
              type: 'missing_filled',
              column: col,
              row: idx + 1,
              original: '(vazio)',
              corrected: nextVal,
              description: `Campo vazio preenchido com valor seguinte válido na linha ${idx + 1}, coluna "${col}"`,
            });
            data[idx][col] = nextVal;
          }
        }
      }
    } else {
      // Text: use most frequent value
      const freq: Record<string, number> = {};
      for (const v of nonEmptyValues) {
        const key = String(v).trim();
        freq[key] = (freq[key] || 0) + 1;
      }

      let mostFrequent = '';
      let maxFreq = 0;
      for (const [key, count] of Object.entries(freq)) {
        if (count > maxFreq) {
          maxFreq = count;
          mostFrequent = key;
        }
      }

      if (mostFrequent) {
        for (const idx of emptyIndices) {
          logs.push({
            type: 'missing_filled',
            column: col,
            row: idx + 1,
            original: '(vazio)',
            corrected: mostFrequent,
            description: `Campo vazio preenchido com valor mais frequente ("${mostFrequent}") na linha ${idx + 1}, coluna "${col}"`,
          });
          data[idx][col] = mostFrequent;
        }
      }
    }

    // Linear interpolation for numeric columns as secondary strategy
    if (colType === 'number') {
      const numValues = data.map(row => {
        const v = row[col];
        if (v === null || v === undefined || v === '') return null;
        return parseNumber(String(v));
      });

      // Check if there are still empty values
      const stillEmpty = numValues
        .map((v, i) => (v === null ? i : -1))
        .filter(i => i !== -1);

      if (stillEmpty.length > 0 && numValues.some(v => v !== null)) {
        for (const idx of stillEmpty) {
          // Find previous and next valid values
          let prevIdx = -1;
          let nextIdx = -1;
          for (let j = idx - 1; j >= 0; j--) {
            if (numValues[j] !== null) { prevIdx = j; break; }
          }
          for (let j = idx + 1; j < numValues.length; j++) {
            if (numValues[j] !== null) { nextIdx = j; break; }
          }

          let fillValue: number | null = null;
          if (prevIdx >= 0 && nextIdx >= 0) {
            // Interpolate
            const prev = numValues[prevIdx]!;
            const next = numValues[nextIdx]!;
            fillValue = Math.round((prev + next) / 2);
          } else if (prevIdx >= 0) {
            fillValue = numValues[prevIdx];
          } else if (nextIdx >= 0) {
            fillValue = numValues[nextIdx];
          }

          if (fillValue !== null) {
            logs.push({
              type: 'missing_filled',
              column: col,
              row: idx + 1,
              original: '(vazio)',
              corrected: String(fillValue),
              description: `Campo vazio preenchido por interpolação (${fillValue}) na linha ${idx + 1}, coluna "${col}"`,
            });
            data[idx][col] = String(fillValue);
          }
        }
      }
    }
  }

  return data;
}

// ============ MAIN EXPORT ============

export function processFileData(
  rawData: Record<string, unknown>[],
  fileName: string,
  format: 'csv' | 'xlsx',
  config?: CleaningConfig
): CleaningResult {
  const logs: CleaningLog[] = [];
  const id = uuidv4();

  const originalRowCount = rawData.length;
  const originalColumnCount = rawData.length > 0 ? Object.keys(rawData[0]).length : 0;

  // Store original preview (first 10 rows)
  const originalPreview = rawData.slice(0, 10);

  logs.push({
    type: 'info',
    description: `Iniciando processamento do arquivo "${fileName}" (${originalRowCount} linhas, ${originalColumnCount} colunas)`,
  });

  const defaultConfig: CleaningConfig = {
    standardizeColumns: true,
    splitColumns: true,
    fixErrors: true,
    removeDuplicates: true,
    fillMissing: true,
    missingStrategy: 'smart',
  };
  const actualConfig = config || defaultConfig;

  // Step 1: Standardize columns
  let data = rawData;
  let columnsStandardized = 0;
  if (actualConfig.standardizeColumns) {
    const standardizeResult = standardizeColumns(data, logs);
    data = standardizeResult.data;
    columnsStandardized = logs.filter(l => l.type === 'standardization').length;
  }

  // Step 2: Split composite columns (e.g., "595/368,9" -> two columns)
  let columnsSplit = 0;
  if (actualConfig.splitColumns) {
    data = splitCompositeColumns(data, logs);
    columnsSplit = logs.filter(l => l.type === 'column_split').length;
  }

  // Step 3: Detect and fix errors
  let errorsFixed = 0;
  if (actualConfig.fixErrors) {
    data = detectAndFixErrors(data, logs);
    errorsFixed = logs.filter(l => l.type === 'error_fix').length;
  }

  // Step 4: Remove duplicates
  let duplicatesRemoved = 0;
  if (actualConfig.removeDuplicates) {
    const beforeCount = data.length;
    data = removeDuplicates(data, logs);
    duplicatesRemoved = beforeCount - data.length;
  }

  // Step 5: Fill missing values
  let missingValuesFilled = 0;
  if (actualConfig.fillMissing && actualConfig.missingStrategy === 'smart') {
    data = fillMissingValues(data, logs);
    missingValuesFilled = logs.filter(l => l.type === 'missing_filled').length;
  }

  // Store cleaned preview
  const cleanedPreview = data.slice(0, 10);

  const columns = data.length > 0 ? Object.keys(data[0]) : [];

  // Save processed file
  const downloadUrl = `/api/download?id=${id}`;
  const processedDir = path.join(os.tmpdir(), 'dataclean-pro-processed');
  if (!fs.existsSync(processedDir)) {
    fs.mkdirSync(processedDir, { recursive: true });
  }

  if (format === 'xlsx') {
    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.json_to_sheet(data);
    
    // Auto-fit column widths for maximum legibility and professional design
    if (data.length > 0) {
      const cols = Object.keys(data[0]);
      ws['!cols'] = cols.map(col => {
        const maxLength = Math.max(
          col.length,
          ...data.map(row => String(row[col] ?? '').length)
        );
        // Minimum width of 10, maximum of 50 to keep layout organized
        return { wch: Math.max(10, Math.min(50, maxLength + 3)) };
      });
    }

    XLSX.utils.book_append_sheet(wb, ws, 'Dados Limpos');
    const buf = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });
    fs.writeFileSync(path.join(processedDir, `${id}.xlsx`), buf);
  } else {
    const csv = Papa.unparse(data);
    fs.writeFileSync(path.join(processedDir, `${id}.csv`), csv, 'utf-8');
  }

  // Also save the cleaning log
  fs.writeFileSync(
    path.join(processedDir, `${id}.json`),
    JSON.stringify({ id, fileName, logs, columns, format }),
    'utf-8'
  );

  logs.push({
    type: 'info',
    description: `Processamento concluído. ${data.length} linhas resultantes, ${columns.length} colunas.`,
  });

  return {
    id,
    originalFileName: fileName,
    originalFormat: format,
    rowCount: originalRowCount,
    columnCount: originalColumnCount,
    cleanedRowCount: data.length,
    cleanedColumnCount: columns.length,
    columns,
    originalPreview,
    cleanedPreview,
    logs,
    stats: {
      errorsFixed,
      duplicatesRemoved,
      missingValuesFilled,
      columnsStandardized,
      columnsSplit,
    },
    cleanedData: data,
    downloadUrl,
  };
}

export function readFileContent(
  buffer: Buffer,
  fileName: string
): { data: Record<string, unknown>[]; format: 'csv' | 'xlsx' } {
  const ext = fileName.toLowerCase().split('.').pop();

  if (ext === 'xlsx') {
    const workbook = XLSX.read(buffer, { type: 'buffer' });
    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];
    const data = XLSX.utils.sheet_to_json(sheet, { defval: '' });
    return { data, format: 'xlsx' };
  } else if (ext === 'csv') {
    const text = buffer.toString('utf-8');
    const result = Papa.parse(text, {
      header: true,
      skipEmptyLines: true,
      dynamicTyping: false,
    });
    return { data: result.data as Record<string, unknown>[], format: 'csv' };
  } else {
    throw new Error(`Formato de arquivo não suportado: .${ext}. Use .csv ou .xlsx`);
  }
}

export function cleanOldProcessedFiles(): void {
  try {
    const processedDir = path.join(os.tmpdir(), 'dataclean-pro-processed');
    if (!fs.existsSync(processedDir)) return;

    const files = fs.readdirSync(processedDir);
    const now = Date.now();
    const oneHour = 60 * 60 * 1000;

    for (const file of files) {
      const filePath = path.join(processedDir, file);
      // Skip if it's a directory
      if (fs.statSync(filePath).isDirectory()) continue;
      
      const stat = fs.statSync(filePath);
      if (now - stat.mtimeMs > oneHour) {
        fs.unlinkSync(filePath);
      }
    }
  } catch (error) {
    console.error('Erro ao limpar arquivos antigos:', error);
  }
}
