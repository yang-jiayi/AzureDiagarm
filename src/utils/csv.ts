// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

const SPREADSHEET_FORMULA_PREFIX = /^[\u0000-\u0020]*[=+\-@]/;

/**
 * Encode a text field for CSV and neutralize spreadsheet formulas from
 * imported or user-authored labels.
 */
export function csvTextCell(value: unknown, alwaysQuote = false): string {
  const text = value == null ? '' : String(value);
  const safeValue = SPREADSHEET_FORMULA_PREFIX.test(text) ? `'${text}` : text;
  const escaped = safeValue.replace(/"/g, '""');
  return alwaysQuote || /[",\r\n]/.test(safeValue)
    ? `"${escaped}"`
    : escaped;
}
