# TIQ Assistant - Technical Overview & Security Summary

## What is TIQ Assistant?

TIQ Assistant is a **local desktop timesheet helper** that runs as a system tray application. It helps users track their work hours by:
- Displaying scheduled popup reminders at configurable times (default: 12:15 and 18:15)
- Reading calendar meetings from local Outlook to auto-suggest time entries
- Exporting timesheet data to Excel files

---

## Core Technologies

| Component | Technology | Purpose |
|-----------|------------|---------|
| Desktop UI | **PyQt6** | Native Windows application with system tray |
| Scheduler | **APScheduler** | Cron-style scheduling for timed popups |
| Database | **SQLite** | Local data storage (no external server) |
| Calendar | **pywin32** (COM) | Read-only access to local Outlook calendar |
| Export | **xlsxwriter/openpyxl** | Generate Excel files locally |

---

## How the System Tray Widget Works

1. **Startup**: The app sets a Windows App User Model ID and creates a `QSystemTrayIcon`
2. **Background Operation**: Runs silently with `setQuitOnLastWindowClosed(False)` - closing windows doesn't exit the app
3. **Notifications**: Uses native Windows toast notifications via `QSystemTrayIcon.showMessage()`
4. **Menu**: Right-click opens a context menu for manual access to features
5. **Popups**: Modal `QDialog` windows that appear at scheduled times

### Scheduler Details

The scheduler uses APScheduler's `BackgroundScheduler` with:
- **Cron triggers** for weekday-only popups (Monday-Friday)
- **Misfire grace time** of 15 minutes (handles laptop sleep/wake scenarios)
- **Thread-safe signal emission** via `QTimer.singleShot()` to safely communicate with the Qt UI thread
- **Snooze support** - users can delay reminders by 15 minutes

---

## How Outlook Calendar Integration Works

The Outlook reader (`src/tiq_assistant/integrations/outlook_reader.py`) uses **Windows COM automation** via `pywin32`:

```python
import win32com.client
outlook = win32com.client.Dispatch("Outlook.Application")
namespace = outlook.GetNamespace("MAPI")
calendar = namespace.GetDefaultFolder(9)  # 9 = olFolderCalendar
```

### Key Points

| Aspect | Description |
|--------|-------------|
| **Read-only access** | Only reads calendar items, never modifies them |
| **Local only** | Communicates with locally-installed Outlook desktop, not cloud services |
| **No credentials stored** | Uses Windows COM which inherits the logged-in user's session |
| **No network requests** | All data comes from the local Outlook cache/mailbox |
| **Graceful fallback** | If Outlook isn't installed or access is denied, the app continues working without calendar features |

### What Data is Read

- Meeting subject/title
- Start and end times
- Duration
- Whether it's a Teams meeting
- Whether it's recurring
- Organizer name
- Location
- Meeting body (truncated to 2000 characters)

### What is NOT Accessed

- Email messages
- Contacts
- Tasks
- Other mailbox folders
- Other users' calendars

---

## Data Storage

All data is stored in a **local SQLite database** at:
```
~/.tiq_assistant/tiq_assistant.db
```

### Database Tables

| Table | Purpose |
|-------|---------|
| `projects` | User-defined projects/tickets for time entries |
| `timesheet_entries` | Logged time entries with hours, descriptions, etc. |
| `user_settings` | Consultant ID, default location, default project |
| `schedule_config` | Popup times, target hours per session |
| `recent_projects` | Quick-access list of recently used projects |
| `outlook_meetings` | Cached meeting data for project matching |
| `holidays` | Turkish national holidays (full and half days) |
| `skipped_days` | Days marked as sick leave, vacation, etc. |

**No external database connections** - everything is file-based and local.

---

## Export Functionality

Timesheet entries are exported to Excel files stored locally:

```
~/Documents/TIQ Timesheets/January_2026_v1.xlsx
```

### Export Features

- Auto-incrementing version numbers (v1, v2, v3...)
- Standard Excel format compatible with corporate timesheet systems
- Columns: Consultant ID, Date, Hours, Ticket No, Project, Activity, Location, Description

---

## Security & Privacy Assurances

| Concern | Status |
|---------|--------|
| **Network requests** | **None.** The app makes no HTTP/HTTPS requests. |
| **Cloud storage** | **None.** All data stored locally in SQLite. |
| **Credentials** | **Not stored.** Outlook access uses Windows session. |
| **Data export** | **Local only.** Excel files saved to Documents folder. |
| **External APIs** | **None called.** No JIRA/Teams/etc. integrations require auth. |
| **Permissions** | Only file system (for SQLite + exports) and Outlook COM (read-only). |
| **Auto-start** | Optional Windows auto-start via a visible shortcut in the user's Startup folder (`shell:startup`), user-controlled in Settings. |
| **Telemetry** | **None.** No usage data, crash reports, or analytics collected. |

---

## Why This Application is Safe

### 1. 100% Offline Operation
The application doesn't make any network requests. All functionality works without internet connection. You can verify this by monitoring network traffic while using the app.

### 2. Read-Only Outlook Access
The COM interface only reads calendar items. It cannot:
- Send emails
- Modify calendar events
- Delete anything
- Access other mailbox data (contacts, tasks, etc.)

### 3. Local Database
SQLite stores data in a single file in the user's profile directory. There are:
- No database servers
- No remote connections
- No cloud sync

### 4. No Secret Storage
No passwords, API keys, or tokens are stored anywhere. Outlook access inherits authentication from the Windows login session.

### 5. Transparent Exports
Excel files are saved to a visible, user-accessible folder (`Documents/TIQ Timesheets/`) with clear, descriptive filenames.

### 6. Standard Libraries
Uses well-known, widely-audited Python packages:
- **PyQt6** - Qt framework bindings (used by thousands of applications)
- **APScheduler** - Standard Python scheduling library
- **openpyxl/xlsxwriter** - Common Excel libraries
- **pywin32** - Official Python Windows extensions

### 7. No Background Data Collection
The app doesn't:
- Phone home to any server
- Collect usage statistics
- Send crash reports
- Track user behavior

---

## File Locations Summary

| Type | Path |
|------|------|
| Database | `~/.tiq_assistant/tiq_assistant.db` |
| Excel exports | `~/Documents/TIQ Timesheets/*.xlsx` |
| Application logs | Console only (not persisted to disk) |
| Configuration | Stored in SQLite database |

---

## Source Code Structure

```
src/tiq_assistant/
├── desktop/                    # PyQt6 desktop application
│   ├── app.py                  # Main entry point, coordinates all components
│   ├── tray.py                 # System tray icon and context menu
│   ├── scheduler.py            # APScheduler for timed popup reminders
│   ├── icon.py                 # Custom teal clock icon generation
│   └── windows/
│       ├── main_window.py      # Dashboard with workday overview
│       ├── day_entry_dialog.py # Time entry popup for specific days
│       └── settings_dialog.py  # Configuration UI
│
├── integrations/
│   └── outlook_reader.py       # Read-only Outlook COM access
│
├── services/
│   ├── matching_service.py     # Match meetings to projects via keywords
│   ├── timesheet_service.py    # Business logic for entries
│   └── hour_suggestion_service.py
│
├── storage/
│   └── sqlite_store.py         # SQLite database operations
│
├── exporters/
│   └── excel_exporter.py       # Excel file generation
│
├── core/
│   ├── models.py               # Data models (Project, Entry, etc.)
│   └── holidays.py             # Turkish holiday definitions
│
└── web/
    └── streamlit_app.py        # Alternative web interface (optional)
```

---

## Verification Steps

If you want to verify the security claims:

1. **No network activity**: Use Windows Resource Monitor or Wireshark to confirm no network connections are made.

2. **Database inspection**: Open `~/.tiq_assistant/tiq_assistant.db` with any SQLite viewer (like DB Browser for SQLite) to see exactly what's stored.

3. **Source code review**: All code is in the `src/tiq_assistant/` directory and can be reviewed.

4. **Outlook access scope**: The COM calls are limited to `GetDefaultFolder(9)` (calendar only) and read operations.
