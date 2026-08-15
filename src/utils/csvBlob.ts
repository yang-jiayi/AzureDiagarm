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

/**
 * BOM-prefix a CSV that is *not* being handed out as a standalone Blob — a
 * `.csv` entry inside a ZIP, say. Extracting it produces exactly the file the
 * standalone download would have, so Excel treats the two identically.
 */
export function csvText(csv: string): string {
  return csv.startsWith(UTF8_BOM) ? csv : `${UTF8_BOM}${csv}`;
}

export function csvBlob(csv: string): Blob {
  return new Blob([csvText(csv)], { type: 'text/csv;charset=utf-8' });
}
