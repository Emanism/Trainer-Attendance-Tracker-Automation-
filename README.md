# Trainer Attendance Tracker Automation (Google Calendar → Google Sheets)

This Google Apps Script automates tracking of trainer attendance for bootcamps by syncing Google Calendar event data into a Google Sheet. The script extracts trainer names and bootcamp codes from calendar events, matches them to the correct date and bootcamp columns in the sheet, and annotates each cell with the trainer’s name and session duration.

---

## Features

- **Automated Sync:** Periodically fetches events from Google Calendar and writes to a Google Sheet.
- **Flexible Time Filtering:** Only considers events within specified time ranges (default: 7–9 PM, with Saturday handled as 12–2 PM).
- **Intelligent Trainer Extraction:** Robust regex-based parsing to extract trainer names from event titles/descriptions and, as fallback, from guest lists.
- **Bootcamp Code Recognition:** Detects bootcamp or cohort codes using a priority list of regex patterns for maximum accuracy.
- **Duration Annotation:** Writes the trainer’s name and session duration (e.g., “Alice (1.5 hour)”) into the appropriate cell.
- **Automated Notifications:** Sends email notifications on updates or errors to the active user.
- **Easy Scheduling:** Includes helper functions for setting up daily time-based triggers for full automation.

---

## How It Works

1. **Configuration:**  
   Edit the `CONFIG` object at the top of the script to specify your:
   - Spreadsheet ID & sheet name
   - Calendar ID (use "primary" for your default)
   - Time filters (`START_HOUR`, `END_HOUR`)
   - Lookback and lookahead windows (in days)

2. **Event Fetching & Filtering:**  
   - Fetches all events within the date window.
   - Filters by event time (with special handling for Saturdays).

3. **Trainer & Bootcamp Extraction:**  
   - Uses a prioritized list of regex patterns to extract bootcamp codes from event text.
   - Extracts trainer names using common patterns (Mr/Ms/Dr, "led by", "trainer:", etc.), or falls back to the first non-organizer guest.

4. **Sheet Update:**  
   - Finds the correct row based on event date and column based on bootcamp code.
   - Annotates the cell with "Trainer Name (X hour)" where X is the rounded session duration.

5. **Notifications:**  
   - If updates are made, sends an email to the user.
   - If errors occur, sends an error notification.

6. **Automation:**  
   - `setupDailyTriggerWithTimeFilter()` sets up a daily trigger at 22:00.
   - `removeTriggers()` removes all triggers for this script.

---

## Setup

1. **Copy the Script:**  
   Paste the script into your Google Apps Script project.

2. **Edit the CONFIG:**  
   ```js
   const CONFIG = {
     SPREADSHEET_ID: 'your-sheet-id',
     SHEET_NAME:      'Trainer Sheet',
     CALENDAR_ID:     'primary',
     START_HOUR:      19,
     END_HOUR:        21,
     DAYS_BEHIND:     30,
     DAYS_AHEAD:      30
   };
   ```

3. **Prepare the Sheet:**  
   - Sheet should have a date column (column B) and bootcamp columns (headers) as needed.
   - Dates must be in Date format or recognized as strings.

4. **Authorize & Run:**  
   - Run `runCalendarSheetsAutomation()` manually the first time to authorize permissions.
   - Use `setupDailyTriggerWithTimeFilter()` to automate future runs.

---

## Customization

- **Update bootcamp code patterns:**  
  Edit the `codePatterns` array to support new bootcamp naming conventions.
- **Adjust time windows:**  
  Modify `START_HOUR`, `END_HOUR`, or Saturday logic as needed.
- **Change notification behavior:**  
  Edit or remove `sendNotificationEmail` and `sendErrorNotificationEmail` as desired.

---

## Troubleshooting

- **No updates:**  
  - Check that your calendar and sheet IDs are correct.
  - Make sure events and columns exist for the given bootcamp/date.
- **Trainer/bootcamp not found:**  
  - Adjust regex patterns or ensure event titles/descriptions are consistent.
- **Script errors:**  
  - Check logs in the Apps Script editor.
  - Use the email notifications for details on failures.

---

## Example Usage

- **Run once:**  
  In Apps Script, run `runCalendarSheetsAutomation()` to sync calendar events for the specified window.
- **Automate daily:**  
  Run `setupDailyTriggerWithTimeFilter()` to schedule daily refreshes at 22:00.

---

## License

MIT License

---

## Author

- Shahzadi Eman
- Shahzadieman1122@gmail.com

---
