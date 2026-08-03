/**
 * Reads the "Data" tab of THIS spreadsheet (the consolidated Elmsall
 * Weekly sheet the Databricks job writes to), groups its rows by the
 * Area value in column J, and appends each group to the "Data" tab of
 * that area's own site spreadsheet.
 *
 * Column J is expected to hold one of the AREA_TARGETS keys below,
 * matching the "Area" label the Databricks notebook writes on every row.
 * Rows with any other/blank value in column J are left alone.
 *
 * Once all 5 destinations have been written successfully, A2:J on this
 * sheet's Data tab is wiped (header row 1 is left alone) so the sheet is
 * clean for the next Databricks run. If any destination write fails, the
 * wipe is skipped so no data is lost.
 *
 * Run distributeToSiteSheets() manually, or attach it to a time-driven
 * trigger (Apps Script editor -> Triggers) to run automatically after
 * each Databricks run.
 */

const AREA_TARGETS = {
  'MPF - PiE':               '1_w2Er6K9z5W_08313SLfXWs9bMItubo8d8Hvc3iXlu4',
  'MPF - E3 Topup':          '1HfoJbvUhUxCMahELPDMTzqX7BqmT-2NYAVcA-rRln4w',
  'MPF - Sorter 6 Packing':  '1HrfeWQm6F5uqHqmJ5wxOJvX3uLhzAMXWTOvtEFP7s2w',
  'MPF - E3 Packing':        '1JKb63sTUC2oc56CSFB9GlrhyipGN-A7Xkd_Vhb77HnQ',
  'MPF - Parcel Sort':       '1pMRUkhGZ-ee5JVTp1abSh_Tbn7f_P2LrOZtb2sTYcvA',
};

const AREA_COLUMN_INDEX = 10; // column J
const SOURCE_TAB_NAME = 'Data';
const DEST_TAB_NAME = 'Data';

function distributeToSiteSheets() {
  const sourceSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SOURCE_TAB_NAME);
  if (!sourceSheet) {
    throw new Error(`"${SOURCE_TAB_NAME}" tab not found in this spreadsheet.`);
  }

  const lastRow = sourceSheet.getLastRow();
  if (lastRow < 1) {
    Logger.log('No rows in Data tab; nothing to distribute.');
    return;
  }

  const values = sourceSheet.getRange(1, 1, lastRow, sourceSheet.getLastColumn()).getValues();

  // Group rows by column J (Area), only for areas we have a destination for
  const rowsByArea = {};
  for (const area of Object.keys(AREA_TARGETS)) {
    rowsByArea[area] = [];
  }

  for (const row of values) {
    const area = row[AREA_COLUMN_INDEX - 1];
    if (Object.prototype.hasOwnProperty.call(rowsByArea, area)) {
      rowsByArea[area].push(row);
    }
  }

  let anyFailures = false;

  for (const [area, sheetId] of Object.entries(AREA_TARGETS)) {
    const rows = rowsByArea[area];
    if (rows.length === 0) {
      Logger.log(`[${area}] no matching rows, skipped.`);
      continue;
    }

    const destSpreadsheet = SpreadsheetApp.openById(sheetId);
    const destSheet = destSpreadsheet.getSheetByName(DEST_TAB_NAME);
    if (!destSheet) {
      Logger.log(`[${area}] FAILED -- "${DEST_TAB_NAME}" tab not found in destination spreadsheet.`);
      anyFailures = true;
      continue;
    }

    const destLastRow = destSheet.getLastRow();
    destSheet.getRange(destLastRow + 1, 1, rows.length, rows[0].length).setValues(rows);
    Logger.log(`[${area}] appended ${rows.length} row(s) to ${destSpreadsheet.getName()} (from row ${destLastRow + 1}).`);
  }

  if (anyFailures) {
    Logger.log('One or more destinations failed; leaving this Data tab as-is (not wiped).');
    return;
  }

  if (lastRow > 1) {
    sourceSheet.getRange(2, 1, lastRow - 1, AREA_COLUMN_INDEX).clearContent();
    Logger.log(`Wiped A2:J${lastRow} on this sheet's Data tab.`);
  }
}
