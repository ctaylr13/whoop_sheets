/**
 * WHOOP OAuth + daily recovery/sleep synchronization.
 *
 * Required Apps Script setup:
 *   1. Add the Google Workspace OAuth2 library with identifier `OAuth2`.
 *   2. Add Script Properties named WHOOP_CLIENT_ID and WHOOP_CLIENT_SECRET.
 *   3. Register the URL printed by showWhoopRedirectUri() in the WHOOP app.
 *   4. Run showWhoopAuthorizationUrl(), open the logged URL, and approve access.
 *   5. Run backfillWhoopDaily() once, then installWhoopSyncTrigger().
 */

const WHOOP_SYNC_CONFIG = Object.freeze({
    authorizationUrl: "https://api.prod.whoop.com/oauth/oauth2/auth",
    tokenUrl: "https://api.prod.whoop.com/oauth/oauth2/token",
    apiBaseUrl: "https://api.prod.whoop.com/developer/v2",
    scopes: ["read:recovery", "read:sleep", "offline"],
    sheetName: "WHOOP Daily",
    routineLookbackDays: 14,
    backfillDays: 180,
    pageSize: 25,
    maxPages: 20,
});

const WHOOP_DAILY_HEADERS = Object.freeze([
    "Log Date",
    "WHOOP Sleep ID",
    "WHOOP Cycle ID",
    "Sleep Start",
    "Sleep End",
    "Bedtime",
    "Wake Time",
    "HRV (ms)",
    "RHR (bpm)",
    "Recovery Score (%)",
    "Sleep Performance (%)",
    "Sleep Consistency (%)",
    "Sleep Efficiency (%)",
    "Respiratory Rate",
    "Total Sleep (h)",
    "Time in Bed (h)",
    "Awake (min)",
    "Light Sleep (min)",
    "Deep Sleep (min)",
    "REM Sleep (min)",
    "Sleep Need (h)",
    "SpO2 (%)",
    "Skin Temp (°C)",
    "WHOOP Updated At",
    "Last Synced At",
]);

/** Returns the configured OAuth2 service. Keep this function private. */
function getWhoopService_() {
    const credentials = getWhoopCredentials_();

    return OAuth2.createService("WHOOP")
        .setAuthorizationBaseUrl(WHOOP_SYNC_CONFIG.authorizationUrl)
        .setTokenUrl(WHOOP_SYNC_CONFIG.tokenUrl)
        .setClientId(credentials.clientId)
        .setClientSecret(credentials.clientSecret)
        .setCallbackFunction("authCallback")
        .setPropertyStore(PropertiesService.getUserProperties())
        .setCache(CacheService.getUserCache())
        .setLock(LockService.getUserLock())
        .setScope(WHOOP_SYNC_CONFIG.scopes);
}

/** Logs and returns the redirect URI that must be registered with WHOOP. */
function showWhoopRedirectUri() {
    const redirectUri = OAuth2.getRedirectUri();
    console.log("WHOOP redirect URI: %s", redirectUri);
    return redirectUri;
}

/** Logs and returns the URL used to authorize this script with WHOOP. */
function showWhoopAuthorizationUrl() {
    const service = getWhoopService_();
    if (service.hasAccess()) {
        console.log("WHOOP is already authorized.");
        return "WHOOP is already authorized.";
    }

    const authorizationUrl = service.getAuthorizationUrl();
    console.log(
        "Open this URL and grant access to WHOOP: %s",
        authorizationUrl
    );
    return authorizationUrl;
}

/** OAuth2 callback invoked by Apps Script after WHOOP authorization. */
function authCallback(request) {
    const authorized = getWhoopService_().handleCallback(request);
    const message = authorized
        ? "WHOOP authorization succeeded. You can close this tab and run backfillWhoopDaily()."
        : "WHOOP authorization was denied. Close this tab and run showWhoopAuthorizationUrl() to try again.";

    return HtmlService.createHtmlOutput(
        '<!doctype html><html><body style="font:16px system-ui;padding:32px">' +
            "<h2>" +
            (authorized ? "WHOOP connected" : "WHOOP not connected") +
            "</h2>" +
            "<p>" +
            message +
            "</p>" +
            "</body></html>"
    );
}

/** Returns a safe connection summary without exposing credentials or tokens. */
function whoopStatus() {
    const service = getWhoopService_();
    const result = {
        authorized: service.hasAccess(),
        redirectUri: OAuth2.getRedirectUri(),
        syncTriggerInstalled: ScriptApp.getProjectTriggers().some(function (
            trigger
        ) {
            return trigger.getHandlerFunction() === "syncWhoopDaily";
        }),
    };
    console.log(JSON.stringify(result, null, 2));
    return result;
}

/** Removes the stored WHOOP access and refresh tokens for the current user. */
function disconnectWhoop() {
    getWhoopService_().reset();
    console.log("WHOOP authorization was removed.");
}

/** Routine sync used by the hourly trigger. */
function syncWhoopDaily() {
    return syncWhoopDaily_(WHOOP_SYNC_CONFIG.routineLookbackDays);
}

/** One-time historical load. Safe to run again because rows are upserted by sleep ID. */
function backfillWhoopDaily() {
    return syncWhoopDaily_(WHOOP_SYNC_CONFIG.backfillDays);
}

/** Installs one hourly trigger after removing older copies of the same trigger. */
function installWhoopSyncTrigger() {
    removeWhoopSyncTriggers();
    ScriptApp.newTrigger("syncWhoopDaily").timeBased().everyHours(1).create();
    console.log("Installed an hourly WHOOP sync trigger.");
}

/** Removes all triggers that call syncWhoopDaily. */
function removeWhoopSyncTriggers() {
    ScriptApp.getProjectTriggers().forEach(function (trigger) {
        if (trigger.getHandlerFunction() === "syncWhoopDaily") {
            ScriptApp.deleteTrigger(trigger);
        }
    });
}

function syncWhoopDaily_(lookbackDays) {
    const lock = LockService.getScriptLock();
    lock.waitLock(30000);

    try {
        const service = getWhoopService_();
        if (!service.hasAccess()) {
            throw new Error(
                "WHOOP is not authorized. Run showWhoopAuthorizationUrl(), open the logged URL, and grant access."
            );
        }

        const end = new Date();
        const start = new Date(
            end.getTime() - lookbackDays * 24 * 60 * 60 * 1000
        );
        const sleeps = whoopFetchCollection_(
            "/activity/sleep",
            start,
            end,
            service
        ).filter(function (sleep) {
            return sleep && sleep.nap !== true;
        });
        const recoveries = whoopFetchCollection_(
            "/recovery",
            start,
            end,
            service
        );
        const recoveryBySleepId = {};
        const recoveryByCycleId = {};

        recoveries.forEach(function (recovery) {
            if (recovery.sleep_id)
                recoveryBySleepId[String(recovery.sleep_id)] = recovery;
            if (recovery.cycle_id !== null && recovery.cycle_id !== undefined) {
                recoveryByCycleId[String(recovery.cycle_id)] = recovery;
            }
        });

        const spreadsheet = openConfiguredSpreadsheet_();
        const sheet = getOrCreateWhoopSheet_(spreadsheet);
        const timezone =
            spreadsheet.getSpreadsheetTimeZone() || Session.getScriptTimeZone();
        const syncedAt = new Date();
        const existingRows = indexWhoopRowsBySleepId_(sheet);
        let inserted = 0;
        let updated = 0;

        sleeps.forEach(function (sleep) {
            if (!sleep.id) return;

            const sleepId = String(sleep.id);
            const recovery =
                recoveryBySleepId[sleepId] ||
                recoveryByCycleId[String(sleep.cycle_id)] ||
                null;
            const values = whoopDailyRow_(sleep, recovery, timezone, syncedAt);
            const row = existingRows[sleepId];

            if (row) {
                sheet
                    .getRange(row, 1, 1, WHOOP_DAILY_HEADERS.length)
                    .setValues([values]);
                updated += 1;
            } else {
                const newRow = Math.max(sheet.getLastRow() + 1, 2);
                sheet
                    .getRange(newRow, 1, 1, WHOOP_DAILY_HEADERS.length)
                    .setValues([values]);
                existingRows[sleepId] = newRow;
                inserted += 1;
            }
        });

        formatWhoopSheet_(sheet);
        if (sheet.getLastRow() > 2) {
            sheet
                .getRange(
                    2,
                    1,
                    sheet.getLastRow() - 1,
                    WHOOP_DAILY_HEADERS.length
                )
                .sort([
                    { column: 1, ascending: true },
                    { column: 5, ascending: true },
                ]);
        }
        SpreadsheetApp.flush();

        // Upsert WHOOP-owned base rows immediately after each sync. Existing
        // rows are enriched; missing dates get a base row whose behavioral
        // fields remain blank until the Shortcut updates them.
        // Routine syncs only need the recent window. A historical backfill
        // still hydrates its full requested range, and the separate manual
        // hydrateRawEntriesFromWhoop() function remains available for a full
        // repair without slowing every web request and hourly trigger.
        const rawHydration = hydrateAllRawEntriesFromWhoop_(
            spreadsheet,
            Math.max(CONFIG.whoopLookbackRows, lookbackDays + 2)
        );
        SpreadsheetApp.flush();

        const result = {
            ok: true,
            lookbackDays: lookbackDays,
            sleepsFound: sleeps.length,
            recoveriesFound: recoveries.length,
            inserted: inserted,
            updated: updated,
            rawEntriesCreated: rawHydration.createdRows,
            rawEntriesMatched: rawHydration.matchedRows,
            rawColumnsUpdated: rawHydration.updatedColumns,
        };
        console.log(JSON.stringify(result, null, 2));
        return result;
    } finally {
        lock.releaseLock();
    }
}

function whoopFetchCollection_(path, start, end, service) {
    const records = [];
    let nextToken = "";

    for (let page = 0; page < WHOOP_SYNC_CONFIG.maxPages; page += 1) {
        const query = [
            "limit=" + WHOOP_SYNC_CONFIG.pageSize,
            "start=" + encodeURIComponent(start.toISOString()),
            "end=" + encodeURIComponent(end.toISOString()),
        ];
        if (nextToken) query.push("nextToken=" + encodeURIComponent(nextToken));

        const payload = whoopGetJson_(
            WHOOP_SYNC_CONFIG.apiBaseUrl + path + "?" + query.join("&"),
            service
        );
        Array.prototype.push.apply(records, payload.records || []);
        nextToken = payload.next_token || "";
        if (!nextToken) return records;
    }

    throw new Error(
        "WHOOP returned more pages than the configured safety limit."
    );
}

function whoopGetJson_(url, service) {
    const response = UrlFetchApp.fetch(url, {
        method: "get",
        headers: { Authorization: "Bearer " + service.getAccessToken() },
        muteHttpExceptions: true,
    });
    const status = response.getResponseCode();
    const body = response.getContentText();

    if (status < 200 || status >= 300) {
        throw new Error(
            "WHOOP API request failed (" + status + "): " + body.slice(0, 500)
        );
    }

    try {
        return JSON.parse(body);
    } catch (_) {
        throw new Error("WHOOP API returned an invalid JSON response.");
    }
}

function whoopDailyRow_(sleep, recovery, spreadsheetTimezone, syncedAt) {
    const sleepScore = sleep.score || {};
    const stages = sleepScore.stage_summary || {};
    const sleepNeeded = sleepScore.sleep_needed || {};
    const recoveryScore = recovery && recovery.score ? recovery.score : {};
    const timezoneOffset = sleep.timezone_offset || "";
    const totalSleepMilli = sumNumbers_([
        stages.total_light_sleep_time_milli,
        stages.total_slow_wave_sleep_time_milli,
        stages.total_rem_sleep_time_milli,
    ]);
    const sleepNeedMilli = sumNumbers_([
        sleepNeeded.baseline_milli,
        sleepNeeded.need_from_sleep_debt_milli,
        sleepNeeded.need_from_recent_strain_milli,
        sleepNeeded.need_from_recent_nap_milli,
    ]);

    return [
        whoopLocalDate_(sleep.end, timezoneOffset, spreadsheetTimezone),
        String(sleep.id),
        valueOrBlank_(sleep.cycle_id),
        dateOrBlank_(sleep.start),
        dateOrBlank_(sleep.end),
        whoopLocalWallClock_(sleep.start, timezoneOffset, spreadsheetTimezone),
        whoopLocalWallClock_(sleep.end, timezoneOffset, spreadsheetTimezone),
        valueOrBlank_(recoveryScore.hrv_rmssd_milli),
        valueOrBlank_(recoveryScore.resting_heart_rate),
        valueOrBlank_(recoveryScore.recovery_score),
        valueOrBlank_(sleepScore.sleep_performance_percentage),
        valueOrBlank_(sleepScore.sleep_consistency_percentage),
        valueOrBlank_(sleepScore.sleep_efficiency_percentage),
        valueOrBlank_(sleepScore.respiratory_rate),
        millisecondsToHours_(totalSleepMilli),
        millisecondsToHours_(stages.total_in_bed_time_milli),
        millisecondsToMinutes_(stages.total_awake_time_milli),
        millisecondsToMinutes_(stages.total_light_sleep_time_milli),
        millisecondsToMinutes_(stages.total_slow_wave_sleep_time_milli),
        millisecondsToMinutes_(stages.total_rem_sleep_time_milli),
        millisecondsToHours_(sleepNeedMilli),
        valueOrBlank_(recoveryScore.spo2_percentage),
        valueOrBlank_(recoveryScore.skin_temp_celsius),
        latestDate_(sleep.updated_at, recovery && recovery.updated_at),
        syncedAt,
    ];
}

function getOrCreateWhoopSheet_(spreadsheet) {
    let sheet = spreadsheet.getSheetByName(WHOOP_SYNC_CONFIG.sheetName);
    if (!sheet) sheet = spreadsheet.insertSheet(WHOOP_SYNC_CONFIG.sheetName);

    const currentHeader = sheet
        .getRange(1, 1, 1, WHOOP_DAILY_HEADERS.length)
        .getDisplayValues()[0];
    const headerIsBlank = currentHeader.every(function (value) {
        return value === "";
    });

    if (headerIsBlank) {
        sheet
            .getRange(1, 1, 1, WHOOP_DAILY_HEADERS.length)
            .setValues([WHOOP_DAILY_HEADERS]);
    } else if (currentHeader.join("|") !== WHOOP_DAILY_HEADERS.join("|")) {
        throw new Error(
            "The WHOOP Daily headers do not match Whoop.gs. Do not overwrite the sheet; compare row 1 first."
        );
    }

    sheet.setFrozenRows(1);
    return sheet;
}

function indexWhoopRowsBySleepId_(sheet) {
    const result = {};
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return result;

    sheet
        .getRange(2, 2, lastRow - 1, 1)
        .getDisplayValues()
        .forEach(function (row, index) {
            const sleepId = String(row[0] || "").trim();
            if (sleepId) result[sleepId] = index + 2;
        });
    return result;
}

function formatWhoopSheet_(sheet) {
    const dataRows = Math.max(sheet.getLastRow() - 1, 1);
    sheet.getRange(2, 1, dataRows, 1).setNumberFormat("yyyy-mm-dd");
    sheet.getRange(2, 4, dataRows, 2).setNumberFormat("yyyy-mm-dd h:mm AM/PM");
    sheet.getRange(2, 6, dataRows, 2).setNumberFormat("h:mm AM/PM");
    sheet.getRange(2, 8, dataRows, 6).setNumberFormat("0.0");
    sheet.getRange(2, 14, dataRows, 1).setNumberFormat("0.0");
    sheet.getRange(2, 15, dataRows, 2).setNumberFormat("0.00");
    sheet.getRange(2, 17, dataRows, 4).setNumberFormat("0");
    sheet.getRange(2, 21, dataRows, 1).setNumberFormat("0.00");
    sheet.getRange(2, 22, dataRows, 2).setNumberFormat("0.0");
    sheet
        .getRange(2, 24, dataRows, 2)
        .setNumberFormat("yyyy-mm-dd h:mm:ss AM/PM");
}

function getWhoopCredentials_() {
    return {
        clientId: getRequiredScriptProperty_("WHOOP_CLIENT_ID"),
        clientSecret: getRequiredScriptProperty_("WHOOP_CLIENT_SECRET"),
    };
}

function whoopLocalDate_(isoValue, timezoneOffset, fallbackTimezone) {
    const date = dateOrBlank_(isoValue);
    if (!date) return "";
    const timezone = timezoneOffset ? "GMT" + timezoneOffset : fallbackTimezone;
    const key = Utilities.formatDate(date, timezone, "yyyy-MM-dd");
    return Utilities.parseDate(key, fallbackTimezone, "yyyy-MM-dd");
}

function whoopLocalWallClock_(isoValue, timezoneOffset, fallbackTimezone) {
    const date = dateOrBlank_(isoValue);
    if (!date) return "";
    const timezone = timezoneOffset ? "GMT" + timezoneOffset : fallbackTimezone;
    const localText = Utilities.formatDate(
        date,
        timezone,
        "yyyy-MM-dd HH:mm:ss"
    );
    return Utilities.parseDate(
        localText,
        fallbackTimezone,
        "yyyy-MM-dd HH:mm:ss"
    );
}

function dateOrBlank_(value) {
    if (!value) return "";
    const date = new Date(value);
    return Number.isNaN(date.getTime()) ? "" : date;
}

function latestDate_(first, second) {
    const firstDate = dateOrBlank_(first);
    const secondDate = dateOrBlank_(second);
    if (!firstDate) return secondDate;
    if (!secondDate) return firstDate;
    return firstDate.getTime() >= secondDate.getTime() ? firstDate : secondDate;
}

function sumNumbers_(values) {
    let hasNumber = false;
    const total = values.reduce(function (sum, value) {
        if (typeof value !== "number" || Number.isNaN(value)) return sum;
        hasNumber = true;
        return sum + value;
    }, 0);
    return hasNumber ? total : "";
}

function millisecondsToHours_(value) {
    return typeof value === "number" && !Number.isNaN(value)
        ? value / (60 * 60 * 1000)
        : "";
}

function millisecondsToMinutes_(value) {
    return typeof value === "number" && !Number.isNaN(value)
        ? value / (60 * 1000)
        : "";
}

function valueOrBlank_(value) {
    return value === null || value === undefined ? "" : value;
}
