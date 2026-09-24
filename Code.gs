const CONFIG = Object.freeze({
    sheetName: "Raw Entries",
    whoopSheetName: "WHOOP Daily",
    afterMidnightCutoffHour: 6,
    whoopLookbackRows: 10,
});

const COL = Object.freeze({
    TIMESTAMP: 1,
    NAME: 2,
    WAKE_TIME: 3,
    BED_TIME: 4,
    BREAKFAST: 5,
    MORNING_COMPLETE: 6,
    HRV: 7,
    RHR: 8,
    LOG_DATE: 9,
    BREAKFAST_TIME: 10,
    FINAL_BREAKFAST: 11,
    BREAKFAST_FOOD: 12,
    COUCH_SLEEP: 13,
    MOVED_BEFORE_SLEEPY: 14,
    SUPPORTS_USED: 15,
    WHOOP_BEDTIME: 16,
    ACTUAL_BEDTIME: 17,
    BEDTIME_VS_WHOOP: 18,
    RECOVERY_SCORE: 19,
    SLEEP_PERFORMANCE: 20,
    SLEEP_CONSISTENCY: 21,
    SLEEP_EFFICIENCY: 22,
    RESPIRATORY_RATE: 23,
    TOTAL_SLEEP: 24,
    TIME_IN_BED: 25,
    DEEP_SLEEP: 26,
    REM_SLEEP: 27,
    SLEEP_NEED: 28,
});

const WHOOP_COL = Object.freeze({
    LOG_DATE: 1,
    BEDTIME: 6,
    WAKE_TIME: 7,
    HRV: 8,
    RHR: 9,
    RECOVERY_SCORE: 10,
    SLEEP_PERFORMANCE: 11,
    SLEEP_CONSISTENCY: 12,
    SLEEP_EFFICIENCY: 13,
    RESPIRATORY_RATE: 14,
    TOTAL_SLEEP: 15,
    TIME_IN_BED: 16,
    DEEP_SLEEP: 19,
    REM_SLEEP: 20,
    SLEEP_NEED: 21,
});

const WHOOP_RAW_FIELDS = Object.freeze([
    [COL.WAKE_TIME, WHOOP_COL.WAKE_TIME, "Wake Time", "h:mm AM/PM"],
    [COL.BED_TIME, WHOOP_COL.BEDTIME, "Bed Time", "h:mm AM/PM"],
    [COL.HRV, WHOOP_COL.HRV, "HRV", "0"],
    [COL.RHR, WHOOP_COL.RHR, "RHR", "0"],
    [COL.RECOVERY_SCORE, WHOOP_COL.RECOVERY_SCORE, "Recovery Score", "0"],
    [
        COL.SLEEP_PERFORMANCE,
        WHOOP_COL.SLEEP_PERFORMANCE,
        "Sleep Performance",
        "0",
    ],
    [
        COL.SLEEP_CONSISTENCY,
        WHOOP_COL.SLEEP_CONSISTENCY,
        "Sleep Consistency",
        "0",
    ],
    [COL.SLEEP_EFFICIENCY, WHOOP_COL.SLEEP_EFFICIENCY, "Sleep Efficiency", "0"],
    [
        COL.RESPIRATORY_RATE,
        WHOOP_COL.RESPIRATORY_RATE,
        "Respiratory Rate",
        "0.0",
    ],
    [COL.TOTAL_SLEEP, WHOOP_COL.TOTAL_SLEEP, "Total Sleep", "0.00"],
    [COL.TIME_IN_BED, WHOOP_COL.TIME_IN_BED, "Time in Bed", "0.00"],
    [COL.DEEP_SLEEP, WHOOP_COL.DEEP_SLEEP, "Deep Sleep", "0"],
    [COL.REM_SLEEP, WHOOP_COL.REM_SLEEP, "REM Sleep", "0"],
    [COL.SLEEP_NEED, WHOOP_COL.SLEEP_NEED, "Sleep Need", "0.00"],
]);

function doGet() {
    return jsonResponse_({ ok: true, service: "WHOOP daily log" });
}

function doPost(e) {
    const lock = LockService.getScriptLock();

    try {
        lock.waitLock(10000);

        const payload = parsePayload_(e);
        validateRequest_(payload);

        const spreadsheet = openConfiguredSpreadsheet_();
        const sheet = spreadsheet.getSheetByName(CONFIG.sheetName);
        if (!sheet) throw new Error(`Sheet not found: ${CONFIG.sheetName}`);

        const timezone =
            spreadsheet.getSpreadsheetTimeZone() || Session.getScriptTimeZone();
        const now = new Date();
        const action = normalizeAction_(payload.action);
        const name =
            cleanText_(payload.name) ||
            getRequiredScriptProperty_("DEFAULT_NAME");
        const logDateKey = getLogicalDateKey_(payload, action, now, timezone);

        // For the daily Shortcut, let WHOOP establish today's base row first.
        // This makes WHOOP the owner of wake/bed times and objective metrics,
        // while the Shortcut only adds the behavioral answers below.
        if (!action) {
            hydrateAllRawEntriesFromWhoop_(spreadsheet, 1);
        }

        let row = findRowByNameAndDate_(sheet, name, logDateKey, timezone);
        const created = row === null;

        if (created) {
            // Keep a safe fallback for days when WHOOP has not published a
            // sleep yet, so a daily check-in is never discarded.
            row = Math.max(sheet.getLastRow() + 1, 2);
            sheet.getRange(row, COL.TIMESTAMP).setValue(now);
            sheet.getRange(row, COL.NAME).setValue(name);
        }

        const updatedColumns = action
            ? applyNfcUpdate_(sheet, row, payload, action, now, timezone)
            : applyDailyUpdate_(sheet, row, payload, now, timezone);
        const whoop = action
            ? null
            : hydrateWhoopData_(
                  spreadsheet,
                  sheet,
                  row,
                  name,
                  logDateKey,
                  timezone,
                  updatedColumns
              );

        SpreadsheetApp.flush();

        return jsonResponse_({
            ok: true,
            result: created ? "created" : "updated",
            row,
            name,
            logDate: logDateKey,
            action: action || "daily",
            updatedColumns,
            whoop,
        });
    } catch (error) {
        console.error(error);
        return jsonResponse_({ ok: false, error: error.message });
    } finally {
        try {
            lock.releaseLock();
        } catch (_) {
            // The lock was never acquired or was already released.
        }
    }
}

function applyNfcUpdate_(sheet, row, payload, action, now, timezone) {
    const eventTime = parseClockValue_(payload.eventTime, now, timezone);
    if (!eventTime)
        throw new Error("NFC request is missing a valid eventTime.");

    if (action === "wake") {
        sheet.getRange(row, COL.WAKE_TIME).setValue(eventTime);
        return ["Wake Time"];
    }

    if (action === "bedtime") {
        sheet.getRange(row, COL.BED_TIME).setValue(eventTime);
        sheet.getRange(row, COL.ACTUAL_BEDTIME).setValue(eventTime);
        return ["Bed Time", "Actual Bedtime Tonight"];
    }

    throw new Error(`Unsupported NFC action: ${action}`);
}

function applyDailyUpdate_(sheet, row, payload, now, timezone) {
    const updated = [];

    // A daily submission owns the timestamp. An NFC-only row can therefore be
    // created in the morning and later become the completed daily check-in row.
    setCell_(sheet, row, COL.TIMESTAMP, now, "Timestamp", updated);
    setCell_(
        sheet,
        row,
        COL.NAME,
        cleanText_(payload.name) || getRequiredScriptProperty_("DEFAULT_NAME"),
        "Name",
        updated
    );

    // Wake time, bedtime, HRV, and RHR are intentionally not read from this
    // payload. WHOOP owns those fields and hydrates them after this update.

    if (hasKey_(payload, "breakfast")) {
        const breakfast = cleanText_(payload.breakfast);
        setCell_(sheet, row, COL.BREAKFAST, breakfast, "Breakfast", updated);
        setCell_(
            sheet,
            row,
            COL.FINAL_BREAKFAST,
            finalBreakfast_(breakfast),
            "Final Breakfast",
            updated
        );
    }

    const localHour = Number(Utilities.formatDate(now, timezone, "H"));
    setCell_(
        sheet,
        row,
        COL.MORNING_COMPLETE,
        localHour < 11 ? "Yes" : "No",
        "Morning Complete",
        updated
    );

    setTextIfPresent_(
        sheet,
        row,
        COL.BREAKFAST_FOOD,
        payload,
        "breakfastFood",
        "Breakfast Food",
        updated
    );
    setTextIfPresent_(
        sheet,
        row,
        COL.COUCH_SLEEP,
        payload,
        "couchSleep",
        "Fell Asleep on Couch?",
        updated
    );
    setTextIfPresent_(
        sheet,
        row,
        COL.MOVED_BEFORE_SLEEPY,
        payload,
        "movedBeforeSleepy",
        "Moved to Bed Before Sleepy?",
        updated
    );
    setTextIfPresent_(
        sheet,
        row,
        COL.SUPPORTS_USED,
        payload,
        "supportsUsed",
        "Supports Used",
        updated
    );

    if (hasNonBlankValue_(payload, "whoopBedtime")) {
        const whoopTime = parseClockValue_(payload.whoopBedtime, now, timezone);
        if (whoopTime) {
            setCell_(
                sheet,
                row,
                COL.WHOOP_BEDTIME,
                whoopTime,
                "WHOOP Suggested Bedtime",
                updated
            );
        } else {
            // This value is optional and Shortcuts may send a localized full
            // date string. Never abort the rest of the daily update for it.
            console.warn(
                "Skipping invalid optional whoopBedtime value: %s",
                cleanText_(payload.whoopBedtime)
            );
        }
    }

    return updated;
}

function hydrateWhoopData_(
    spreadsheet,
    rawSheet,
    currentRawRow,
    name,
    currentDateKey,
    timezone,
    updated
) {
    const whoopSheet = spreadsheet.getSheetByName(CONFIG.whoopSheetName);
    if (!whoopSheet || whoopSheet.getLastRow() < 2) return null;

    const rowCount = Math.min(
        whoopSheet.getLastRow() - 1,
        CONFIG.whoopLookbackRows
    );
    const startRow = whoopSheet.getLastRow() - rowCount + 1;
    const defaultName = getRequiredScriptProperty_("DEFAULT_NAME");
    const whoopRows = whoopSheet
        .getRange(startRow, 1, rowCount, WHOOP_COL.SLEEP_NEED)
        .getValues();
    const rawRowByDate = indexRawRowsByDate_(rawSheet, name, timezone);
    let currentSummary = null;

    whoopRows.forEach((whoopRow) => {
        const whoopDateKey = dateKeyFromValue_(
            whoopRow[WHOOP_COL.LOG_DATE - 1],
            timezone
        );
        if (!whoopDateKey) return;

        // doPost already resolved the exact row for this submission. Use it
        // directly for today's WHOOP record instead of depending on the
        // ARRAYFORMULA-backed Log Date column to find the row a second time.
        // A web request only needs the WHOOP record for the submitted date.
        // Historical rows are handled by the hourly sync/manual repair.
        if (whoopDateKey !== currentDateKey) return;
        setWhoopFields_(rawSheet, currentRawRow, whoopRow, updated);

        const previousDayRow =
            rawRowByDate[shiftDateKey_(whoopDateKey, -1, timezone)];
        if (previousDayRow) {
            setWhoopCellIfPresent_(
                rawSheet,
                previousDayRow,
                COL.ACTUAL_BEDTIME,
                whoopRow[WHOOP_COL.BEDTIME - 1],
                "Actual Bedtime Tonight",
                updated,
                "h:mm AM/PM"
            );
        }

        currentSummary = whoopSummary_(whoopRow, whoopDateKey);
    });

    return currentSummary;
}

/**
 * Manually creates or hydrates rows from every WHOOP Daily record.
 * Safe to run repeatedly because rows are joined by name and logical date.
 */
function hydrateRawEntriesFromWhoop() {
    const lock = LockService.getScriptLock();
    lock.waitLock(30000);

    try {
        const spreadsheet = openConfiguredSpreadsheet_();
        const result = hydrateAllRawEntriesFromWhoop_(spreadsheet);
        SpreadsheetApp.flush();
        console.log(JSON.stringify(result, null, 2));
        return result;
    } finally {
        lock.releaseLock();
    }
}

/**
 * Internal upsert used by both the manual action and the WHOOP trigger. WHOOP
 * owns the base row and objective metrics; the Shortcut later fills the
 * behavioral fields on that same name/date row. The caller owns locking and
 * flushing.
 */
function hydrateAllRawEntriesFromWhoop_(spreadsheet, maxRows) {
    const rawSheet = spreadsheet.getSheetByName(CONFIG.sheetName);
    const whoopSheet = spreadsheet.getSheetByName(CONFIG.whoopSheetName);
    if (!rawSheet) throw new Error(`Sheet not found: ${CONFIG.sheetName}`);
    if (!whoopSheet || whoopSheet.getLastRow() < 2) {
        return { createdRows: 0, matchedRows: 0, updatedColumns: [] };
    }

    const timezone =
        spreadsheet.getSpreadsheetTimeZone() || Session.getScriptTimeZone();
    const availableRows = whoopSheet.getLastRow() - 1;
    const requestedRows = Number(maxRows);
    const rowCount =
        Number.isFinite(requestedRows) && requestedRows > 0
            ? Math.min(availableRows, Math.floor(requestedRows))
            : availableRows;
    const startRow = whoopSheet.getLastRow() - rowCount + 1;
    const whoopRows = whoopSheet
        .getRange(
            startRow,
            1,
            rowCount,
            WHOOP_COL.SLEEP_NEED
        )
        .getValues();
    const rawRowByDate = indexRawRowsByDate_(
        rawSheet,
        defaultName,
        timezone
    );
    const updatedColumns = [];
    const matchedDates = {};
    let createdRows = 0;

    whoopRows.forEach((whoopRow) => {
        const whoopDateKey = dateKeyFromValue_(
            whoopRow[WHOOP_COL.LOG_DATE - 1],
            timezone
        );
        if (!whoopDateKey) return;

        let sameDayRow = rawRowByDate[whoopDateKey];
        if (!sameDayRow) {
            sameDayRow = Math.max(rawSheet.getLastRow() + 1, 2);
            const wakeTime = whoopRow[WHOOP_COL.WAKE_TIME - 1];
            const logDate = whoopRow[WHOOP_COL.LOG_DATE - 1];
            rawSheet
                .getRange(sameDayRow, COL.TIMESTAMP)
                .setValue(wakeTime || logDate);
            rawSheet
                .getRange(sameDayRow, COL.NAME)
                .setValue(defaultName);
            rawRowByDate[whoopDateKey] = sameDayRow;
            createdRows += 1;
        }

        setWhoopFields_(rawSheet, sameDayRow, whoopRow, updatedColumns);
        matchedDates[whoopDateKey] = true;

        const previousDayRow =
            rawRowByDate[shiftDateKey_(whoopDateKey, -1, timezone)];
        if (previousDayRow) {
            setWhoopCellIfPresent_(
                rawSheet,
                previousDayRow,
                COL.ACTUAL_BEDTIME,
                whoopRow[WHOOP_COL.BEDTIME - 1],
                "Actual Bedtime Tonight",
                updatedColumns,
                "h:mm AM/PM"
            );
        }
    });

    return {
        createdRows: createdRows,
        matchedRows: Object.keys(matchedDates).length,
        updatedColumns: updatedColumns,
    };
}

function indexRawRowsByDate_(sheet, name, timezone) {
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return {};

    const rows = sheet.getRange(2, 1, lastRow - 1, COL.LOG_DATE).getValues();
    const wantedName = name.toLowerCase();
    return rows.reduce((result, row, index) => {
        if (cleanText_(row[COL.NAME - 1]).toLowerCase() !== wantedName)
            return result;
        const dateKey = dateKeyFromValue_(
            row[COL.LOG_DATE - 1] || row[COL.TIMESTAMP - 1],
            timezone
        );
        if (dateKey) result[dateKey] = index + 2;
        return result;
    }, {});
}

function setWhoopCellIfPresent_(
    sheet,
    row,
    column,
    value,
    label,
    updated,
    format
) {
    if (value === "" || value === null || value === undefined) return;
    const range = sheet.getRange(row, column).setValue(value);
    if (format) range.setNumberFormat(format);
    if (!updated.includes(label)) updated.push(label);
}

/** Writes the WHOOP-owned columns in three batches instead of cell by cell. */
function setWhoopFields_(sheet, row, whoopRow, updated) {
    const blocks = [
        {
            column: COL.WAKE_TIME,
            values: [
                whoopRow[WHOOP_COL.WAKE_TIME - 1],
                whoopRow[WHOOP_COL.BEDTIME - 1],
            ],
            labels: ["Wake Time", "Bed Time"],
            formats: ["h:mm AM/PM", "h:mm AM/PM"],
        },
        {
            column: COL.HRV,
            values: [
                whoopRow[WHOOP_COL.HRV - 1],
                whoopRow[WHOOP_COL.RHR - 1],
            ],
            labels: ["HRV", "RHR"],
            formats: ["0", "0"],
        },
        {
            column: COL.RECOVERY_SCORE,
            values: [
                whoopRow[WHOOP_COL.RECOVERY_SCORE - 1],
                whoopRow[WHOOP_COL.SLEEP_PERFORMANCE - 1],
                whoopRow[WHOOP_COL.SLEEP_CONSISTENCY - 1],
                whoopRow[WHOOP_COL.SLEEP_EFFICIENCY - 1],
                whoopRow[WHOOP_COL.RESPIRATORY_RATE - 1],
                whoopRow[WHOOP_COL.TOTAL_SLEEP - 1],
                whoopRow[WHOOP_COL.TIME_IN_BED - 1],
                whoopRow[WHOOP_COL.DEEP_SLEEP - 1],
                whoopRow[WHOOP_COL.REM_SLEEP - 1],
                whoopRow[WHOOP_COL.SLEEP_NEED - 1],
            ],
            labels: [
                "Recovery Score",
                "Sleep Performance",
                "Sleep Consistency",
                "Sleep Efficiency",
                "Respiratory Rate",
                "Total Sleep",
                "Time in Bed",
                "Deep Sleep",
                "REM Sleep",
                "Sleep Need",
            ],
            formats: [
                "0",
                "0",
                "0",
                "0",
                "0.0",
                "0.00",
                "0.00",
                "0",
                "0",
                "0.00",
            ],
        },
    ];

    blocks.forEach((block) => {
        if (
            block.values.every(
                (value) => value === "" || value === null || value === undefined
            )
        ) {
            return;
        }

        sheet
            .getRange(row, block.column, 1, block.values.length)
            .setValues([block.values])
            .setNumberFormats([block.formats]);

        block.values.forEach((value, index) => {
            if (value === "" || value === null || value === undefined) return;
            if (!updated.includes(block.labels[index])) {
                updated.push(block.labels[index]);
            }
        });
    });
}

function whoopSummary_(row, logDate) {
    return {
        logDate,
        bedtime: valueOrNull_(row[WHOOP_COL.BEDTIME - 1]),
        wakeTime: valueOrNull_(row[WHOOP_COL.WAKE_TIME - 1]),
        hrv: valueOrNull_(row[WHOOP_COL.HRV - 1]),
        rhr: valueOrNull_(row[WHOOP_COL.RHR - 1]),
        recoveryScore: valueOrNull_(row[WHOOP_COL.RECOVERY_SCORE - 1]),
        sleepPerformance: valueOrNull_(row[WHOOP_COL.SLEEP_PERFORMANCE - 1]),
        sleepConsistency: valueOrNull_(row[WHOOP_COL.SLEEP_CONSISTENCY - 1]),
        sleepEfficiency: valueOrNull_(row[WHOOP_COL.SLEEP_EFFICIENCY - 1]),
        respiratoryRate: valueOrNull_(row[WHOOP_COL.RESPIRATORY_RATE - 1]),
        totalSleepHours: valueOrNull_(row[WHOOP_COL.TOTAL_SLEEP - 1]),
        timeInBedHours: valueOrNull_(row[WHOOP_COL.TIME_IN_BED - 1]),
        deepSleepMinutes: valueOrNull_(row[WHOOP_COL.DEEP_SLEEP - 1]),
        remSleepMinutes: valueOrNull_(row[WHOOP_COL.REM_SLEEP - 1]),
        sleepNeedHours: valueOrNull_(row[WHOOP_COL.SLEEP_NEED - 1]),
    };
}

function findRowByNameAndDate_(sheet, name, targetDateKey, timezone) {
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return null;

    // Read through Log Date. Column I is an ARRAYFORMULA, so getValues() gives
    // its calculated dates even though only I2 contains the formula itself.
    const rows = sheet.getRange(2, 1, lastRow - 1, COL.LOG_DATE).getValues();
    const wantedName = name.toLowerCase();

    for (let i = rows.length - 1; i >= 0; i -= 1) {
        const row = rows[i];
        if (cleanText_(row[COL.NAME - 1]).toLowerCase() !== wantedName)
            continue;

        const logDate = row[COL.LOG_DATE - 1] || row[COL.TIMESTAMP - 1];
        if (dateKeyFromValue_(logDate, timezone) === targetDateKey)
            return i + 2;
    }

    return null;
}

function getLogicalDateKey_(payload, action, now, timezone) {
    if (hasNonBlankValue_(payload, "logDate")) {
        const explicit = dateKeyFromValue_(payload.logDate, timezone);
        if (!explicit) throw new Error("logDate is not a valid date.");
        return explicit;
    }

    let logicalDate = now;
    const localHour = Number(Utilities.formatDate(now, timezone, "H"));

    // A 12:30 AM bedtime belongs to the day whose evening just ended. This also
    // lines up with the sheet's bedtime-vs-WHOOP formula, which treats times
    // before 6 AM as after-midnight bedtimes.
    if (action === "bedtime" && localHour < CONFIG.afterMidnightCutoffHour) {
        logicalDate = new Date(now.getTime() - 24 * 60 * 60 * 1000);
    }

    return Utilities.formatDate(logicalDate, timezone, "yyyy-MM-dd");
}

function parsePayload_(e) {
    if (!e || !e.postData || !e.postData.contents) {
        throw new Error("Missing POST body.");
    }

    try {
        return JSON.parse(e.postData.contents);
    } catch (_) {
        throw new Error("POST body must be valid JSON.");
    }
}

function validateRequest_(payload) {
    const action = normalizeAction_(payload.action);
    const propertyName = action ? "NFC_REQUEST_SECRET" : "DAILY_REQUEST_SECRET";
    const expectedSecret = getRequiredScriptProperty_(propertyName);
    if (cleanText_(payload.secret) !== expectedSecret)
        throw new Error("Unauthorized.");
}

function openConfiguredSpreadsheet_() {
    return SpreadsheetApp.openById(
        getRequiredScriptProperty_("SPREADSHEET_ID")
    );
}

function getRequiredScriptProperty_(name) {
    const value = String(
        PropertiesService.getScriptProperties().getProperty(name) || ""
    ).trim();
    if (!value) throw new Error(`Missing required Script Property: ${name}`);
    return value;
}

function normalizeAction_(value) {
    const action = cleanText_(value).toLowerCase();
    if (!action) return "";
    if (action === "wake") return "wake";
    if (action === "bed" || action === "bedtime") return "bedtime";
    return action;
}

function parseClockValue_(value, baseDate, timezone) {
    if (value instanceof Date && !Number.isNaN(value.getTime())) return value;

    const text = cleanText_(value);
    if (!text) return null;

    // Accept both a plain clock time and a full date/time string emitted by
    // Apple Shortcuts, for example "9/24/2026, 10:55:00 PM".
    const match = text.match(/(?:^|[T,\s])(\d{1,2}):(\d{2})(?::(\d{2}))?\s*(AM|PM)?/i);
    if (!match) return null;

    let hour = Number(match[1]);
    const minute = Number(match[2]);
    const second = Number(match[3] || 0);
    const meridiem = (match[4] || "").toUpperCase();

    if (meridiem) {
        if (hour < 1 || hour > 12) return null;
        if (hour === 12) hour = 0;
        if (meridiem === "PM") hour += 12;
    }

    if (hour > 23 || minute > 59 || second > 59) return null;

    const datePart = Utilities.formatDate(baseDate, timezone, "yyyy-MM-dd");
    return Utilities.parseDate(
        `${datePart} ${pad2_(hour)}:${pad2_(minute)}:${pad2_(second)}`,
        timezone,
        "yyyy-MM-dd HH:mm:ss"
    );
}

function dateKeyFromValue_(value, timezone) {
    if (value instanceof Date && !Number.isNaN(value.getTime())) {
        return Utilities.formatDate(value, timezone, "yyyy-MM-dd");
    }

    // ARRAYFORMULA date results can be returned by Apps Script as a raw
    // Google Sheets serial number instead of a Date object. Convert the whole
    // day using the Sheets epoch so it can still join to WHOOP Daily.
    if (typeof value === "number" && Number.isFinite(value)) {
        const sheetsEpoch = Date.UTC(1899, 11, 30);
        const date = new Date(
            sheetsEpoch + Math.floor(value) * 24 * 60 * 60 * 1000
        );
        return Utilities.formatDate(date, "UTC", "yyyy-MM-dd");
    }

    const text = cleanText_(value);
    if (!text) return "";

    let match = text.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
    if (match) return `${match[1]}-${pad2_(match[2])}-${pad2_(match[3])}`;

    match = text.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
    if (match) return `${match[3]}-${pad2_(match[1])}-${pad2_(match[2])}`;

    return "";
}

function finalBreakfast_(breakfast) {
    const value = cleanText_(breakfast).toLowerCase();
    if (value === "yes") return "Yes";
    if (value === "no" || value === "no/skipping") return "No";
    if (value === "no/not yet") return "Pending";
    return breakfast;
}

function setTimeIfPresent_(
    sheet,
    row,
    column,
    value,
    baseDate,
    timezone,
    label,
    updated
) {
    if (!cleanText_(value)) return;
    const parsed = parseClockValue_(value, baseDate, timezone);
    if (!parsed) throw new Error(`${label} is not a valid time.`);
    setCell_(sheet, row, column, parsed, label, updated);
}

function setTextIfPresent_(sheet, row, column, payload, key, label, updated) {
    if (!hasKey_(payload, key)) return;
    setCell_(sheet, row, column, cleanText_(payload[key]), label, updated);
}

function setCell_(sheet, row, column, value, label, updated) {
    sheet.getRange(row, column).setValue(value);
    updated.push(label);
}

function hasKey_(object, key) {
    return Object.prototype.hasOwnProperty.call(object, key);
}

function hasNonBlankValue_(object, key) {
    return hasKey_(object, key) && cleanText_(object[key]) !== "";
}

function cleanText_(value) {
    if (value === null || value === undefined) return "";
    return String(value).trim();
}

function valueOrNull_(value) {
    return value === "" || value === null || value === undefined ? null : value;
}

function pad2_(value) {
    return String(value).padStart(2, "0");
}

function shiftDateKey_(dateKey, days, timezone) {
    const date = Utilities.parseDate(
        `${dateKey} 12:00:00`,
        timezone,
        "yyyy-MM-dd HH:mm:ss"
    );
    date.setDate(date.getDate() + days);
    return Utilities.formatDate(date, timezone, "yyyy-MM-dd");
}

function jsonResponse_(body) {
    return ContentService.createTextOutput(JSON.stringify(body)).setMimeType(
        ContentService.MimeType.JSON
    );
}
