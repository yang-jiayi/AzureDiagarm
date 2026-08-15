// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

/**
 * Excel does not read a `.csv` as UTF-8 unless the file opens with a byte
 * order mark. Without one it falls back to the system ANSI code page, so a
 * Japanese service name — the common case here — arrives as mojibake, and the
 * user's only clue is that the spreadsheet looks broken. The `charset=utf-8`
 * on the blob's MIME type does not help: it is discarded the moment the file
 * lands on disk.
 *
 * Every CSV this app hands to the user goes through here so none of them can
 * forget the BOM. Other tools (pandas, `csv` in Python, Numbers, Sheets) all
 * skip a leading BOM, so adding it costs nothing anywhere else.
 */
export const UTF8_BOM = '\ufeff';

export function csvBlob(csv: string): Blob {
  const withBom = csv.startsWith(UTF8_BOM) ? csv : `${UTF8_BOM}${csv}`;
  return new Blob([withBom], { type: 'text/csv;charset=utf-8' });
}
