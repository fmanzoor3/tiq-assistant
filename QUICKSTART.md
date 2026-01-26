# TIQ Assistant - Quick Start Guide

## First-Time Setup

### Personal Laptop (with IDE/virtual environment)

If you're developing in VS Code or an IDE with the project open:
```cmd
pip install -r requirements.txt
pip install -e .
```

### Work Laptop (fresh install from ZIP)

If this is a fresh install or new machine, run these commands first:
```cmd
cd "C:\path\to\tiq-assistant"
pip install -r requirements.txt
pip install -e .
```

This installs all dependencies and registers the package so Python can find `tiq_assistant`.

---

## Running the App

### Desktop App (System Tray)

**Personal Laptop** (from project directory):
```cmd
python -m tiq_assistant
```

**Work Laptop** (from any directory):
```cmd
cd "C:\path\to\tiq-assistant"
python -m tiq_assistant
```

This starts the app in your system tray (near the clock). Features:
- Automatic popups at 12:15 (morning) and 18:15 (afternoon) on weekdays
- Right-click tray icon for manual access
- Syncs with Outlook calendar for meeting detection (requires Outlook desktop)
- Calendar Import tab to fetch meetings for any date range
- Workday overview showing expected hours and progress per day

### Web App (Streamlit)
```cmd
cd "C:\path\to\tiq-assistant"
python -m tiq_assistant web
```

Opens a browser-based interface for managing projects and viewing entries.

---

## Updating the Code

When you download a new version (ZIP from JupyterHub):

1. Extract and replace the existing folder
2. Run `pip install -e .` again to update the package registration
3. Run `python -m tiq_assistant`

---

## Stopping the App

- **Desktop**: Right-click tray icon → "Exit", or press `Ctrl+C` in terminal
- **Web**: Press `Ctrl+C` in terminal

---

## Testing During Development

### Test the Popup Window Directly
```cmd
python -c "from PyQt6.QtWidgets import QApplication; from datetime import date; from tiq_assistant.core.models import SessionType; from tiq_assistant.desktop.windows.time_entry_popup import TimeEntryPopup; app = QApplication([]); popup = TimeEntryPopup(SessionType.MORNING, date.today()); popup.show(); app.exec()"
```

### Test Outlook Integration
```cmd
python -c "from tiq_assistant.integrations.outlook_reader import get_outlook_reader; from datetime import date; reader = get_outlook_reader(); meetings = reader.get_meetings_for_date(date.today()) if reader.is_available() else []; [print(f'{m.display_time} - {m.subject}') for m in meetings]"
```

### Test Database Connection
```cmd
python -c "from tiq_assistant.storage.sqlite_store import get_store; store = get_store(); projects = store.get_projects(); print(f'Found {len(projects)} projects'); [print(f'  - {p.name}') for p in projects]"
```

### Test Schedule Config
```cmd
python -c "from tiq_assistant.storage.sqlite_store import get_store; config = get_store().get_schedule_config(); print(f'Morning: {config.morning_popup_time}'); print(f'Afternoon: {config.afternoon_popup_time}')"
```

---

## Project Structure

```
TIQ Assistant/
├── src/tiq_assistant/
│   ├── desktop/           # PyQt6 desktop app (NEW)
│   │   ├── app.py         # Main entry point
│   │   ├── tray.py        # System tray icon
│   │   ├── scheduler.py   # APScheduler for timed popups
│   │   └── windows/
│   │       ├── time_entry_popup.py  # The popup dialog
│   │       └── settings_dialog.py   # Settings UI
│   │
│   ├── web/               # Streamlit web app (ORIGINAL)
│   │   └── streamlit_app.py
│   │
│   ├── integrations/      # External integrations
│   │   └── outlook_reader.py  # Outlook COM automation
│   │
│   ├── services/          # Business logic
│   │   ├── matching_service.py      # JIRA key matching
│   │   ├── timesheet_service.py     # Entry management
│   │   └── hour_suggestion_service.py  # Smart hour calc
│   │
│   ├── storage/           # Data persistence
│   │   └── sqlite_store.py  # SQLite database
│   │
│   ├── core/              # Data models
│   │   └── models.py
│   │
│   └── exporters/         # Excel export
│       └── excel_exporter.py
│
└── Documents/TIQ Timesheets/  # Export location (created automatically)
    └── Timesheet_2026-01.xlsx
```

---

## Key Features

### Morning Popup (12:15)
- Accounts for 3 hours (9:30 - 12:15)
- Quick project selection
- Auto-detects morning meetings from Outlook
- Smart hour suggestions

### Afternoon Popup (18:15)
- Accounts for 5 hours (13:30 - 18:15)
- Full day summary
- Export to monthly Excel file
- Review all day's entries

### Workday Overview
- Shows all workdays in the selected month
- Displays expected vs filled hours per day
- Excludes weekends and Turkish national holidays
- Half-day holidays (like Ramazan/Kurban Bayramı arife) show 4h expected
- Color-coded: green = complete, yellow = partial, red = past unfilled
- Click a day to auto-fill the date and suggested hours in the add form

### Settings (Right-click tray → Settings)
- Change popup times
- Adjust target hours
- Configure workday schedule
- Enable/disable Windows auto-start

---

## Troubleshooting

### "Outlook not available"
- Make sure Outlook desktop is installed
- Outlook doesn't need to be open, but must be configured

### Popup doesn't appear at scheduled time
- Check if it's a weekday
- Verify times in Settings
- Make sure the app is running (check system tray)

### Can't see tray icon
- Click the ^ arrow in system tray to show hidden icons
- The icon might be in the overflow area
