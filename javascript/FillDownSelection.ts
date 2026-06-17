/**
 * Fill Down (Power Query style) on the current selection.
 * -------------------------------------------------------------------
 * For each COLUMN in the selection, every blank cell is filled with the
 * nearest non-blank value ABOVE it (within the selection). When a new
 * non-blank value appears it becomes the new fill source and continues
 * downward - exactly like Power Query's Fill > Down.
 *
 * Result is written back as STATIC VALUES only.
 *
 * Notes:
 *  - Leading blanks (above the first value in a column) stay blank.
 *  - Whitespace-only cells (" "), 0, false and error values count as
 *    VALUES, not blanks. Only truly empty cells and "" get filled.
 *  - Multi-area (Ctrl-selected) selections are each handled.
 *  - Dates come back from getValues() as serial numbers, so a filled
 *    date cell shows the number unless its format is already a date.
 *    See the note at the bottom if you need formats carried down too.
 */
function main(workbook: ExcelScript.Workbook): void {
  const areas = workbook.getSelectedRanges().getAreas();
  for (const area of areas) {
    fillDownArea(area);
  }
}

/** Fill a single contiguous range, column by column. */
function fillDownArea(range: ExcelScript.Range): void {
  const rowCount = range.getRowCount();
  const colCount = range.getColumnCount();
  if (rowCount * colCount < 2) return; // single cell: nothing to fill

  const values = range.getValues();
  let changed = false;

  for (let c = 0; c < colCount; c++) {
    let lastVal: string | number | boolean = "";
    let hasVal = false;

    for (let r = 0; r < rowCount; r++) {
      const v = values[r][c];
      if (isBlank(v)) {
        if (hasVal) {
          values[r][c] = lastVal;
          changed = true;
        }
      } else {
        lastVal = v;
        hasVal = true;
      }
    }
  }

  if (changed) {
    range.setValues(values);
  }
}

/**
 * Blank = a truly empty cell or a zero-length string.
 * getValues() returns empty cells as "" and error cells as their error
 * string (e.g. "#N/A"), so the single check below treats errors,
 * whitespace, 0 and false as values - matching the VBA version.
 */
function isBlank(value: string | number | boolean): boolean {
  return value === "";
}

/*
 * ---------------------------------------------------------------------
 * CARRY NUMBER FORMAT (optional)
 * If you fill date/percent/currency columns and want the filled blanks
 * to display correctly, replace the `if (changed) range.setValues(...)`
 * block with a per-column copy that also applies the source format.
 * The simplest robust approach: after setValues, copy the format of the
 * first non-blank cell in each column down its filled cells using
 * range.getCell(r, c).setNumberFormatLocal(sourceFormat). Ask if you
 * want that version wired in - it trades a bit of speed for fidelity.
 * ---------------------------------------------------------------------
 */
