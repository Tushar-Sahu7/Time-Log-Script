# Time Log Script: Google Calendar to Sheets -> An Intelligent Time Log Automation

This Google Apps Script provides a powerful, precise, and user-friendly way to sync your Google Calendar events to a Google Sheet for advanced time tracking and analysis. It's more than just a simple sync; it's an intelligent automation designed to be both fast and careful with your data.

This script was born from a simple idea: "How can I get my calendar events into a spreadsheet without the manual work?" What started as a single button evolved into a robust tool that handles complex scenarios gracefully. This documentation tells the story of its creation and gives you everything you need to use and customize it.

## ✨ Features

- **One-Click Sync:** Add events from any date or date range to your sheet with a simple, user-friendly date picker.
- **Intelligent Refresh:** The star of the show. The refresh function doesn't just wipe and replace data. It intelligently compares your sheet with your calendar and performs a three-way sync:
  - **➕ Adds** new events.
  - **✏️ Modifies** events that have been changed (time, title, etc.).
  - **❌ Deletes** events from the sheet that were deleted from the calendar.
- **Data Protection:** The refresh is surgically precise. It only updates columns that are managed by the calendar, **leaving any manual notes or data you've entered in other columns completely untouched.**
- **Performance Optimized:** The script is designed to be fast. It fetches all required data from Google Calendar in a **single batch operation**, avoiding the slow, loop-based approach that plagues many similar scripts.
- **Smart & Safe:** The script includes "guard clauses" to protect your data. It automatically checks that you are syncing events to the correct sheet, preventing you from accidentally adding February's events to your January log.
- **Handles Complex Events:** Multi-day events are automatically and accurately split into individual entries for each day, perfect for daily time logging.

## 🚀 How It Works: The Developer's Journey

This script didn't start out this smart. It evolved by solving one problem at a time.

#### **1. From a Simple Button to a Date Picker**

The first idea was a "Sync Today" button. But what if you forgot to sync yesterday? The solution was to create a `processSelectedDateRange` function and a simple HTML pop-up, giving the user full control over what they sync.

#### **2. The Bottleneck: "Why is it so slow?"**

The first version of the date range sync was slow because it called the Google Calendar API once for every single day in the range (30 days = 30 API calls).

- **The Fix: Batching.** We created a single, unified function, `_fetchAndProcessEvents`, that fetches the entire date range in **one API call**. This was the single biggest performance improvement and is now the core of the entire script.

#### **3. The "Aha!" Moment: The Intelligent Refresh**

Syncing new events was easy, but updating them was hard. Simply deleting all old data and writing new data would destroy any manual notes a user had made.

- **The Fix: The "Snapshot" Method.** For every event, the script creates a unique key (`EventID_Date`) and a "snapshot" string of all the calendar-managed data (`"2025-08-08|09:00|Team Meeting..."`). By comparing the old snapshot with the new one, the script knows _exactly_ what changed. This allows it to perform a precise update, only changing the columns it's supposed to.

## 🛠️ Installation Guide

Setting up this automation is a simple 3-step process.

#### **Step 1: Create Your Google Sheet**

1. Create a new Google Sheet.
2. Name your sheets according to a `"Month YY"` format (e.g., **"August 25"**, **"September 25"**). This is crucial for the script's validation logic.
3. Set up your header rows. The script assumes you have **3 header rows**, but you can configure this.

#### **Step 2: Add the Script**

1. In your Google Sheet, go to **Extensions > Apps Script**.
2. Delete any placeholder code in the `Code.gs` file and **paste the entire `code.gs` script** from this repository.
3. Click the **+** icon to add a new file, select **HTML**, and name it `DatePicker.html`.
4. Delete the placeholder code and **paste the entire `DatePicker.html` code** from this repository.
5. Save both files.

#### **Step 3: Enable the Calendar API**

1. In the Apps Script editor, click on **Services** in the left-hand menu.
2. Find **Google Calendar API** in the list, click **Add**. This enables the "advanced" service that allows for efficient batch fetching.
3. You're done! Refresh your Google Sheet, and you should see the new **`⌚ Time Log`** menu appear.

### **A Note on Permissions (Important First-Time Step)**

The very first time you try to run any function from the `⌚ Time Log` menu, Google will show you a pop-up window titled **"Authorization required"**. This is a standard and essential security step for all Google Apps Scripts.

1. Click **"Review permissions"**.
2. Choose the Google Account you want to use with this sheet.
3. You will likely see a screen saying **"Google hasn’t verified this app"**. This is completely normal for personal scripts that aren't on the public marketplace. It does not mean the script is dangerous.
4. Click on the small **"Advanced"** link, and then click on **"Go to [Your Script Name] (unsafe)"**.
5. Finally, review the permissions the script needs (like viewing your calendar and managing your sheets) and click **"Allow"**.

You only have to do this once. This process gives _your script_, running in _your account_, permission to work with _your data_ on your behalf.

## ⚙️ Configuration & Customization

This script is designed to be adapted to your specific needs. Here’s how to safely modify it.

#### **1. The `CONFIG` Object**

This is your main control panel at the top of `code.gs`.

```
const CONFIG = {
  HEADER_ROWS: 3,       // Number of header rows to skip
  ID_COLUMN_INDEX: 16,    // Column Q (0-indexed) for Event ID
  DATE_COLUMN_INDEX: 0    // Column A (0-indexed) for Date
};
```

- `HEADER_ROWS`: If you use more or fewer header rows, change this number.
- `ID_COLUMN_INDEX` & `DATE_COLUMN_INDEX`: If you move the Event ID or Date columns, update these values. **Remember, columns are 0-indexed (Column A = 0, B = 1, etc.).**

#### **2. The `_formatEventSegment` Function**

This function is a mirror of your sheet's column structure. If you reorder your columns, you must reorder the lines in this function's `return` statement to match.

```
// Example: This array directly maps to your sheet's columns.
return [
  Utilities.formatDate(start, tz, "yyyy-MM-dd"), // Column A
  Utilities.formatDate(start, tz, "HH:mm"),      // Column B
  // ...and so on
];
```

#### **3. The `CALENDAR_MANAGED_COLS` Array (Most Important!)**

Found inside the `refreshTimeLog` function, this array is the key to protecting your manual data. It tells the script which columns it is allowed to overwrite during a refresh.

```
const CALENDAR_MANAGED_COLS = [1, 2, 3, 4, 5, 11, 12, 13, 14, 15, 16, 17];
```

- These numbers are **1-based** (Column A = 1, B = 2, etc.).
- If you add a new column that should be synced from the calendar, **add its number to this array**.
- If you have a column for manual notes that you **never** want the script to touch, **make sure its number is NOT in this list.**
