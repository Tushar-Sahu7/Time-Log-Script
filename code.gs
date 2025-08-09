// ===================================
//  CONFIGURATION
// ===================================
const CONFIG = {
  HEADER_ROWS: 3, // Number of header rows to skip
  ID_COLUMN_INDEX: 16, // Column Q (0-indexed) for Event ID
  DATE_COLUMN_INDEX: 0, // Column A (0-indexed) for Date
};

// ===================================
//  MENU
// ===================================
/**
 * Adds a custom menu to the spreadsheet UI upon opening.
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("⌚ Time Log")
    .addItem("🕛 Sync Today's Time Log", "syncToday")
    .addItem("📅 Sync Time Log by Date", "openDatePickerUI")
    .addItem("🔁 Refresh Sync of Active Sheet", "refreshTimeLog")
    .addToUi();
}

// ===================================
//  UNIFIED DATA-FETCHING & HELPER FUNCTIONS
// ===================================

/**
 * The single, unified function to fetch and process calendar data.
 * It handles fetching, splitting multi-day events, and formatting rows.
 * @param {Date} startDate The start of the date range.
 * @param {Date} endDate The end of the date range.
 * @returns {Array<Array<Object>>} An array of rows ready for the sheet.
 */
function _fetchAndProcessEvents(startDate, endDate) {
  const tz = Session.getScriptTimeZone();
  const calendarId = CalendarApp.getDefaultCalendar().getId();
  let events;

  const effectiveStartDate = new Date(startDate);
  effectiveStartDate.setHours(0, 0, 0, 0);

  const effectiveEndDate = new Date(endDate);
  effectiveEndDate.setHours(23, 59, 59, 999);

  try {
    const eventList =
      Calendar.Events.list(calendarId, {
        timeMin: effectiveStartDate.toISOString(),
        timeMax: effectiveEndDate.toISOString(),
        singleEvents: true,
        orderBy: "startTime",
        // --- IMPLEMENTED: Selective Field Fetching for better performance ---
        fields: "items(id,summary,description,start,end,location,hangoutLink)",
      }).items || [];
    events = eventList.filter((event) => event.start.dateTime);
  } catch (err) {
    SpreadsheetApp.getUi().alert(
      `Could not connect to Google Calendar: ${err.message}`
    );
    return [];
  }

  const allRows = [];
  for (const event of events) {
    const originalStart = new Date(event.start.dateTime);
    const originalEnd = new Date(event.end.dateTime);
    let loopStart = new Date(originalStart);

    while (loopStart < originalEnd) {
      const dayEnd = new Date(loopStart);
      dayEnd.setHours(24, 0, 0, 0);
      const segmentEnd = new Date(Math.min(dayEnd, originalEnd));
      if (loopStart >= segmentEnd) break;

      if (loopStart >= effectiveStartDate && loopStart <= effectiveEndDate) {
        const row = _formatEventSegment(event, loopStart, segmentEnd, tz);
        allRows.push(row);
      }

      loopStart = new Date(dayEnd);
    }
  }
  return allRows;
}

/**
 * Helper to format a single event segment into a row array.
 */
function _formatEventSegment(event, start, end, tz) {
  const durationMs = end.getTime() - start.getTime();
  const hrs = Math.floor(durationMs / 3600000);
  const mins = Math.floor((durationMs % 3600000) / 60000);
  const durationStr = `${String(hrs).padStart(2, "0")}:${String(mins).padStart(
    2,
    "0"
  )}`;

  return [
    Utilities.formatDate(start, tz, "yyyy-MM-dd"),
    Utilities.formatDate(start, tz, "HH:mm"),
    Utilities.formatDate(end, tz, "HH:mm"),
    event.summary || "(No Title)",
    durationStr,
    "",
    "",
    "",
    "",
    "", // Placeholder columns F-J
    event.description || "",
    Utilities.formatDate(start, tz, "w"),
    start.toLocaleString("en-US", { month: "long", timeZone: tz }),
    start.getFullYear(),
    start.toLocaleString("en-US", { weekday: "long", timeZone: tz }),
    event.hangoutLink || event.location || "",
    event.id,
  ];
}

// ===================================
//  CORE SYNC & REFRESH LOGIC
// ===================================

/**
 * High-precision, high-speed refresh with robust locking and error handling.
 */
function refreshTimeLog() {
  const lock = LockService.getScriptLock();
  if (lock.tryLock(0)) {
    try {
      const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
      const ui = SpreadsheetApp.getUi();
      const tz = Session.getScriptTimeZone();

      const CALENDAR_MANAGED_COLS = [1, 2, 3, 4, 5, 11, 12, 13, 14, 15, 16, 17];

      try {
        getActiveSheetMonthYear();
      } catch (err) {
        ui.alert(err.message);
        return; // Stop execution if validation fails
      }

      const lastRow = sheet.getLastRow();
      if (lastRow <= CONFIG.HEADER_ROWS) {
        ui.alert("⛔ No data to refresh.");
        return;
      }

      ui.alert("⏳ Refreshing data... Please wait.");

      // --- The entire refresh logic is now safely wrapped ---
      const dataRange = sheet.getRange(
        CONFIG.HEADER_ROWS + 1,
        1,
        lastRow - CONFIG.HEADER_ROWS,
        sheet.getLastColumn()
      );
      const dataRowsAndCols = dataRange.getValues();
      const existingDataMap = new Map();
      const uniqueDates = new Set();

      dataRowsAndCols.forEach((row, index) => {
        const eventId = row[CONFIG.ID_COLUMN_INDEX];
        const eventDate = row[CONFIG.DATE_COLUMN_INDEX];
        if (eventId && eventDate instanceof Date) {
          const dateString = Utilities.formatDate(eventDate, tz, "yyyy-MM-dd");
          const key = `${eventId}_${dateString}`;
          uniqueDates.add(dateString);
          const standardizedRow = [...row];
          standardizedRow[0] = Utilities.formatDate(row[0], tz, "yyyy-MM-dd");
          if (row[1] instanceof Date)
            standardizedRow[1] = Utilities.formatDate(row[1], tz, "HH:mm");
          if (row[2] instanceof Date)
            standardizedRow[2] = Utilities.formatDate(row[2], tz, "HH:mm");
          const contentSnapshot = CALENDAR_MANAGED_COLS.map(
            (col) => standardizedRow[col - 1]
          ).join("|");
          existingDataMap.set(key, {
            rowNum: index + CONFIG.HEADER_ROWS + 1,
            snapshot: contentSnapshot,
          });
        }
      });

      if (uniqueDates.size === 0) {
        ui.alert("No valid dates found to refresh.");
        return;
      }

      const dateArray = [...uniqueDates].map((ds) => new Date(ds));
      const minDate = new Date(Math.min(...dateArray));
      const maxDate = new Date(Math.max(...dateArray));
      const freshEventRows = _fetchAndProcessEvents(minDate, maxDate);

      // --- IMPLEMENTED: Safeguard against accidental mass deletion ---
      const existingEventCount = existingDataMap.size;
      const freshEventCount = freshEventRows.length;
      const DELETION_THRESHOLD = 0.8; // Cancels if refresh would delete over 80% of data.

      if (
        existingEventCount > 10 &&
        freshEventCount < existingEventCount * (1 - DELETION_THRESHOLD)
      ) {
        ui.alert(
          `🛑 Refresh Canceled: Safety Precaution\n\n` +
            `The script detected that this refresh would delete over ${
              DELETION_THRESHOLD * 100
            }% of your data ` +
            `(${existingEventCount} rows -> ${freshEventCount} rows).\n\n` +
            `This can happen if the Calendar API has a temporary issue. Please check your calendar and try again.`
        );
        return; // Stop the function to prevent data loss.
      }
      // --- END OF SAFEGUARD ---

      const rowsToAdd = [];
      const rowsToUpdate = [];
      const freshKeys = new Set();

      for (const row of freshEventRows) {
        const key = `${row[CONFIG.ID_COLUMN_INDEX]}_${
          row[CONFIG.DATE_COLUMN_INDEX]
        }`;
        freshKeys.add(key);
        const newSnapshot = CALENDAR_MANAGED_COLS.map(
          (col) => row[col - 1]
        ).join("|");
        const existingEvent = existingDataMap.get(key);
        if (existingEvent) {
          if (existingEvent.snapshot !== newSnapshot) {
            rowsToUpdate.push({
              rowNum: existingEvent.rowNum,
              newRowData: row,
            });
          }
        } else {
          if (uniqueDates.has(row[CONFIG.DATE_COLUMN_INDEX])) {
            rowsToAdd.push(row);
          }
        }
      }

      const rowsToDelete = [...existingDataMap.keys()]
        .filter((key) => !freshKeys.has(key))
        .map((key) => existingDataMap.get(key).rowNum);

      for (const update of rowsToUpdate) {
        for (const colIndex of CALENDAR_MANAGED_COLS) {
          sheet
            .getRange(update.rowNum, colIndex)
            .setValue(update.newRowData[colIndex - 1]);
        }
      }

      if (rowsToAdd.length > 0) {
        sheet
          .getRange(
            sheet.getLastRow() + 1,
            1,
            rowsToAdd.length,
            rowsToAdd[0].length
          )
          .setValues(rowsToAdd);
      }

      rowsToDelete
        .sort((a, b) => b - a)
        .forEach((rowNum) => sheet.deleteRow(rowNum));

      ui.alert(
        `✅ Refresh Complete!\n\n- Added: ${rowsToAdd.length}\n- Modified: ${rowsToUpdate.length}\n- Deleted: ${rowsToDelete.length}`
      );
    } catch (e) {
      SpreadsheetApp.getUi().alert(
        "❌ An unexpected error occurred during refresh. Please try again.\n\nError: " +
          e.message
      );
    } finally {
      lock.releaseLock();
    }
  } else {
    SpreadsheetApp.getUi().alert(
      "⚠️ Another operation is already in progress. Please wait for it to finish and then try again."
    );
  }
}

/**
 * Unified sync function with robust locking and error handling.
 */
function processSelectedDateRange(startStr, endStr) {
  const lock = LockService.getScriptLock();
  if (lock.tryLock(0)) {
    try {
      const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
      const ui = SpreadsheetApp.getUi();
      const tz = Session.getScriptTimeZone();
      const startDate = new Date(startStr);
      const endDate = new Date(endStr);
      const isSingleDay = startStr === endStr;

      try {
        const activeMonthYear = getActiveSheetMonthYear();
        const startMatches =
          startDate.getMonth() === activeMonthYear.month &&
          startDate.getFullYear() === activeMonthYear.year;
        const endMatches =
          endDate.getMonth() === activeMonthYear.month &&
          endDate.getFullYear() === activeMonthYear.year;
        if (!startMatches || !endMatches) {
          const displayDate = Utilities.formatDate(
            startDate,
            tz,
            "dd MMM yyyy"
          );
          const errorMsg = isSingleDay
            ? `❌ Sync date (${displayDate}) does not match the active sheet (${sheet.getName()}).`
            : `❌ The date range must be within the month of the active sheet (${sheet.getName()}).`;
          ui.alert(errorMsg);
          return;
        }
      } catch (err) {
        ui.alert(err.message);
        return;
      }

      const existingKeys = new Set(
        sheet
          .getDataRange()
          .getValues()
          .slice(CONFIG.HEADER_ROWS)
          .map((r) =>
            r[CONFIG.ID_COLUMN_INDEX] &&
            r[CONFIG.DATE_COLUMN_INDEX] instanceof Date
              ? `${r[CONFIG.ID_COLUMN_INDEX]}_${Utilities.formatDate(
                  r[CONFIG.DATE_COLUMN_INDEX],
                  tz,
                  "yyyy-MM-dd"
                )}`
              : null
          )
          .filter(Boolean)
      );

      const allFetchedRows = _fetchAndProcessEvents(startDate, endDate);
      const newRows = allFetchedRows.filter((row) => {
        const key = `${row[CONFIG.ID_COLUMN_INDEX]}_${
          row[CONFIG.DATE_COLUMN_INDEX]
        }`;
        return !existingKeys.has(key);
      });

      if (newRows.length > 0) {
        sheet
          .getRange(
            sheet.getLastRow() + 1,
            1,
            newRows.length,
            newRows[0].length
          )
          .setValues(newRows);
      }

      let summaryMessage = "";
      if (isSingleDay) {
        const displayDate = Utilities.formatDate(startDate, tz, "dd MMM yyyy");
        if (newRows.length > 0) {
          summaryMessage = `✅ Synced ${newRows.length} new event(s) for ${displayDate}.`;
        } else if (allFetchedRows.length > 0) {
          summaryMessage = `✅ All events for ${displayDate} were already synced.`;
        } else {
          summaryMessage = `⚠️ No events found for ${displayDate}.`;
        }
      } else {
        const displayStart = Utilities.formatDate(startDate, tz, "dd MMM");
        const displayEnd = Utilities.formatDate(endDate, tz, "dd MMM yyyy");
        summaryMessage = `✅ Range sync complete: ${displayStart} - ${displayEnd}.\n\n`;
        summaryMessage +=
          newRows.length > 0
            ? `Total new events synced: ${newRows.length}`
            : "All events in this period were already synced.";
      }
      ui.alert(summaryMessage);
    } catch (e) {
      SpreadsheetApp.getUi().alert(
        "❌ An unexpected error occurred during sync. Please try again.\n\nError: " +
          e.message
      );
    } finally {
      lock.releaseLock();
    }
  } else {
    SpreadsheetApp.getUi().alert(
      "⚠️ Another operation is already in progress. Please wait for it to finish and then try again."
    );
  }
}

// ===================================
//  UI & UTILITY FUNCTIONS
// ===================================

function openDatePickerUI() {
  try {
    getActiveSheetMonthYear();
  } catch (err) {
    SpreadsheetApp.getUi().alert(err.message);
    return;
  }
  const html = HtmlService.createHtmlOutputFromFile("DatePicker")
    .setWidth(320)
    .setHeight(300);
  SpreadsheetApp.getUi().showModalDialog(html, "Date Picker");
}

function getActiveSheetMonthYear() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const name = sheet.getName().trim();
  const parts = name.split(" ");
  if (parts.length !== 2)
    throw new Error(
      `❌ Invalid sheet: "${name}". Select a sheet like "August 25".`
    );
  const [monthName, yearStr] = parts;
  const month = {
    January: 0,
    February: 1,
    March: 2,
    April: 3,
    May: 4,
    June: 5,
    July: 6,
    August: 7,
    September: 8,
    October: 9,
    November: 10,
    December: 11,
  }[monthName.charAt(0).toUpperCase() + monthName.slice(1).toLowerCase()];
  const year = parseInt("20" + yearStr, 10);
  if (month === undefined || isNaN(year))
    throw new Error(`❌ Invalid name format: "${name}". Use "Month YY".`);
  return { month, year };
}

function getDaysInActiveSheetMonth() {
  const { month, year } = getActiveSheetMonthYear();
  const daysInMonth = new Date(year, month + 1, 0).getDate();
  const tz = Session.getScriptTimeZone();
  const options = [];
  for (let d = 1; d <= daysInMonth; d++) {
    const date = new Date(year, month, d);
    options.push({
      value: Utilities.formatDate(date, tz, "yyyy-MM-dd"),
      display: Utilities.formatDate(date, tz, "d MMM yyyy"),
    });
  }
  return options;
}

function syncToday() {
  const today = new Date();
  const todayStr = Utilities.formatDate(
    today,
    Session.getScriptTimeZone(),
    "yyyy-MM-dd"
  );
  processSelectedDateRange(todayStr, todayStr);
}

function processSelectedDate(dateStr) {
  processSelectedDateRange(dateStr, dateStr);
}
