"""Streamlit web interface for TIQ Assistant."""

import streamlit as st
import pandas as pd
from datetime import date, datetime, timedelta
from pathlib import Path
from io import BytesIO
import tempfile

from tiq_assistant.core.models import (
    Project, TimesheetEntry, UserSettings,
    ActivityCode, EntryStatus, EntrySource
)
from tiq_assistant.storage.sqlite_store import get_store
from tiq_assistant.parsers.outlook_parser import parse_outlook_calendar
from tiq_assistant.services.matching_service import get_matching_service
from tiq_assistant.services.timesheet_service import get_timesheet_service
from tiq_assistant.exporters.excel_exporter import ExcelExporter


# Page configuration
st.set_page_config(
    page_title="TIQ Assistant",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)


def init_session_state():
    """Initialize session state variables."""
    if "store" not in st.session_state:
        st.session_state.store = get_store()
    if "calendar_events" not in st.session_state:
        st.session_state.calendar_events = []
    if "generated_entries" not in st.session_state:
        st.session_state.generated_entries = []
    if "settings" not in st.session_state:
        st.session_state.settings = st.session_state.store.get_settings()


def main():
    """Main application entry point."""
    init_session_state()

    # Sidebar navigation
    st.sidebar.title("📊 TIQ Assistant")
    st.sidebar.markdown("---")

    page = st.sidebar.radio(
        "Navigation",
        ["🏠 Dashboard", "📁 Projects", "📅 Calendar Import", "📝 Timesheet", "⚙️ Settings"],
    )

    # Route to appropriate page
    if page == "🏠 Dashboard":
        show_dashboard()
    elif page == "📁 Projects":
        show_projects()
    elif page == "📅 Calendar Import":
        show_calendar_import()
    elif page == "📝 Timesheet":
        show_timesheet()
    elif page == "⚙️ Settings":
        show_settings()


def show_dashboard():
    """Dashboard page with overview."""
    st.title("Dashboard")

    col1, col2, col3 = st.columns(3)

    # Get stats
    store = st.session_state.store
    projects = store.get_projects()
    today = date.today()
    week_start = today - timedelta(days=today.weekday())
    week_entries = store.get_entries(start_date=week_start, end_date=today)

    with col1:
        st.metric("Active Projects", len(projects))

    with col2:
        week_hours = sum(e.hours for e in week_entries)
        st.metric("Hours This Week", week_hours)

    with col3:
        draft_entries = [e for e in week_entries if e.status == EntryStatus.DRAFT]
        st.metric("Draft Entries", len(draft_entries))

    st.markdown("---")

    # Recent entries
    st.subheader("Recent Entries")
    recent_entries = store.get_entries(
        start_date=today - timedelta(days=7),
        end_date=today
    )

    if recent_entries:
        df = pd.DataFrame([e.to_export_row() for e in recent_entries])
        st.dataframe(df, use_container_width=True, hide_index=True)
    else:
        st.info("No recent entries. Import calendar data or add manual entries.")

    # Quick actions hint
    st.markdown("---")
    st.info("Use the sidebar to navigate: **Calendar Import** to add meetings, **Timesheet** to manage entries, **Projects** to set up ticket mappings.")


def show_projects():
    """Projects management page."""
    st.title("Projects & Tickets")

    store = st.session_state.store
    projects = store.get_projects()

    # Add new project
    with st.expander("➕ Add New Project", expanded=False):
        with st.form("add_project"):
            name = st.text_input("Project Name *", placeholder="BI BÜYÜK VERI PLATFORM SUPPORT")
            ticket_number = st.text_input("Ticket No (Numeric ID) *", placeholder="2019135",
                                          help="Required - The unique numeric ID for this project")
            jira_key = st.text_input("JIRA Key (Optional)", placeholder="PEMP-948",
                                     help="Optional - If project has a JIRA key for calendar matching")
            keywords = st.text_input("Keywords (comma-separated)", placeholder="Agent Bot, big data",
                                     help="Keywords to match calendar events to this project")
            location = st.text_input("Default Location", value="ANKARA")

            if st.form_submit_button("Add Project"):
                if name and ticket_number:
                    project = Project(
                        name=name,
                        ticket_number=ticket_number,
                        jira_key=jira_key if jira_key else None,
                        keywords=[k.strip() for k in keywords.split(",") if k.strip()],
                        default_location=location,
                    )
                    store.save_project(project)
                    st.success(f"Project '{name}' added!")
                    st.rerun()
                else:
                    st.error("Project Name and Ticket No are required")

    st.markdown("---")

    # List projects as a table
    if projects:
        st.subheader("Existing Projects")

        project_data = []
        for p in projects:
            project_data.append({
                "Project Name": p.name,
                "Ticket No": p.ticket_number,
                "JIRA Key": p.jira_key or "-",
                "Keywords": ", ".join(p.keywords) if p.keywords else "-",
                "Location": p.default_location,
            })

        df = pd.DataFrame(project_data)
        st.dataframe(df, use_container_width=True, hide_index=True)

        # Delete section
        st.markdown("---")
        st.subheader("Manage Projects")

        for project in projects:
            col1, col2, col3 = st.columns([3, 1, 1])
            with col1:
                st.write(f"**{project.name}** (Ticket No: {project.ticket_number})")
            with col2:
                if project.jira_key:
                    st.caption(f"JIRA: {project.jira_key}")
            with col3:
                if st.button("🗑️ Delete", key=f"del_{project.id}"):
                    store.delete_project(project.id)
                    st.rerun()
    else:
        st.info("No projects yet. Add one above.")


def show_calendar_import():
    """Calendar import page."""
    st.title("Calendar Import")

    # File upload
    uploaded_file = st.file_uploader(
        "Upload Outlook Calendar Export (Excel)",
        type=["xlsx", "xls"],
        help="Export your calendar from Outlook as an Excel file"
    )

    if uploaded_file:
        # Save to temp file and parse
        with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
            tmp.write(uploaded_file.getvalue())
            tmp_path = tmp.name

        try:
            settings = st.session_state.settings
            events = parse_outlook_calendar(
                tmp_path,
                skip_canceled=settings.skip_canceled_meetings
            )
            st.session_state.calendar_events = events
            st.success(f"Loaded {len(events)} calendar events!")

            # Match events
            matching_service = get_matching_service()
            matching_service.match_events(events)

            # Show unmatched JIRA keys
            unmatched = matching_service.get_unmatched_jira_keys(events)
            if unmatched:
                st.warning(f"Found {len(unmatched)} JIRA keys not in database: {', '.join(unmatched)}")
                st.info("Add these tickets to a project to enable automatic matching.")

        except Exception as e:
            st.error(f"Failed to parse calendar file: {e}")

        # Clean up temp file
        Path(tmp_path).unlink(missing_ok=True)

    # Show parsed events
    if st.session_state.calendar_events:
        st.markdown("---")

        events = st.session_state.calendar_events
        settings = st.session_state.settings
        store = st.session_state.store

        # Filter by minimum duration
        min_hours = st.slider("Minimum duration (hours)", 0.0, 4.0, 0.5, 0.25)
        filtered_events = [e for e in events if float(e.duration_hours) >= min_hours]

        # Separate matched and unmatched
        matched_events = [e for e in filtered_events if e.match_confidence > 0]
        unmatched_events = [e for e in filtered_events if e.match_confidence == 0]

        # Helper to round hours to nearest integer (min 1)
        def round_hours(decimal_hours):
            return max(1, round(float(decimal_hours)))

        # Helper to create event data for display
        def create_event_row(e, idx):
            return {
                "idx": idx,
                "Select": False,
                "Date": e.start_date.strftime("%d.%m.%Y"),
                "Subject": e.subject[:60] + "..." if len(e.subject) > 60 else e.subject,
                "Hours": round_hours(e.duration_hours),
                "JIRA Key": e.matched_jira_key or "",
                "Project": "",
                "event": e,
            }

        # ============ MATCHED EVENTS SECTION ============
        st.subheader(f"✅ Matched Events ({len(matched_events)})")

        if matched_events:
            st.caption("These events have been automatically matched to projects.")

            # Create data for matched events
            matched_data = []
            for i, e in enumerate(matched_events):
                # Get project info
                project_name = ""
                ticket_num = ""
                if e.matched_project_id:
                    project = store.get_project(e.matched_project_id)
                    if project:
                        project_name = project.name
                        ticket_num = project.ticket_number

                matched_data.append({
                    "Select": True,  # Default selected
                    "Date": e.start_date.strftime("%d.%m.%Y"),
                    "Subject": e.subject[:50] + "..." if len(e.subject) > 50 else e.subject,
                    "Hours": round_hours(e.duration_hours),
                    "JIRA Key": e.matched_jira_key or "-",
                    "Project": project_name[:30] + "..." if len(project_name) > 30 else project_name,
                    "Ticket No": ticket_num,
                })

            matched_df = pd.DataFrame(matched_data)

            # Editable table with selection
            edited_matched = st.data_editor(
                matched_df,
                use_container_width=True,
                hide_index=True,
                column_config={
                    "Select": st.column_config.CheckboxColumn("Select", default=True),
                    "Hours": st.column_config.NumberColumn("Hours", min_value=1, max_value=24, step=1),
                },
                disabled=["Date", "Subject", "JIRA Key", "Project", "Ticket No"],
                key="matched_editor",
            )

            # Quick add button for matched
            col1, col2 = st.columns([3, 1])
            with col1:
                selected_count = edited_matched["Select"].sum()
                st.write(f"**{selected_count}** events selected")
            with col2:
                if st.button("➕ Add Selected Matched", type="primary", use_container_width=True):
                    entries_to_add = []
                    for i, row in edited_matched.iterrows():
                        if row["Select"]:
                            event = matched_events[i]
                            # Get project info
                            project_name = None
                            ticket_number = None
                            if event.matched_project_id:
                                project = store.get_project(event.matched_project_id)
                                if project:
                                    project_name = project.name
                                    ticket_number = project.ticket_number

                            entry = TimesheetEntry(
                                consultant_id=settings.consultant_id,
                                entry_date=event.start_date,
                                hours=int(row["Hours"]),
                                ticket_number=ticket_number,
                                project_name=project_name,
                                activity_code=settings.meeting_activity_code,
                                location=settings.default_location,
                                description=event.to_timesheet_description(),
                                status=EntryStatus.DRAFT,
                                source=EntrySource.CALENDAR,
                                source_event_id=event.id,
                                source_jira_key=event.matched_jira_key,
                            )
                            entries_to_add.append(entry)

                    if entries_to_add:
                        for entry in entries_to_add:
                            store.save_entry(entry)
                        st.success(f"Added {len(entries_to_add)} entries!")
                        st.rerun()
        else:
            st.info("No matched events. Add projects with JIRA keys or keywords to enable matching.")

        st.markdown("---")

        # ============ UNMATCHED EVENTS SECTION ============
        st.subheader(f"❓ Unmatched Events ({len(unmatched_events)})")

        if unmatched_events:
            st.caption("These events need manual project/ticket assignment.")

            # Get all projects for dropdown
            projects = store.get_projects()
            project_options = [""] + [p.name for p in projects]

            unmatched_data = []
            for i, e in enumerate(unmatched_events):
                unmatched_data.append({
                    "Select": False,
                    "Date": e.start_date.strftime("%d.%m.%Y"),
                    "Subject": e.subject[:50] + "..." if len(e.subject) > 50 else e.subject,
                    "Hours": round_hours(e.duration_hours),
                    "Project": "",
                })

            unmatched_df = pd.DataFrame(unmatched_data)

            edited_unmatched = st.data_editor(
                unmatched_df,
                use_container_width=True,
                hide_index=True,
                column_config={
                    "Select": st.column_config.CheckboxColumn("Select", default=False),
                    "Hours": st.column_config.NumberColumn("Hours", min_value=1, max_value=24, step=1),
                    "Project": st.column_config.SelectboxColumn("Project", options=project_options),
                },
                disabled=["Date", "Subject"],
                key="unmatched_editor",
            )

            # Add button for unmatched
            col1, col2 = st.columns([3, 1])
            with col1:
                selected_unmatched = edited_unmatched["Select"].sum()
                st.write(f"**{selected_unmatched}** events selected")
            with col2:
                if st.button("➕ Add Selected Unmatched", use_container_width=True):
                    entries_to_add = []
                    for i, row in edited_unmatched.iterrows():
                        if row["Select"]:
                            event = unmatched_events[i]
                            project_name = row["Project"] if row["Project"] else None

                            # Get ticket number from project
                            ticket_number = None
                            if project_name:
                                project = next((p for p in projects if p.name == project_name), None)
                                if project:
                                    ticket_number = project.ticket_number

                            entry = TimesheetEntry(
                                consultant_id=settings.consultant_id,
                                entry_date=event.start_date,
                                hours=int(row["Hours"]),
                                ticket_number=ticket_number,
                                project_name=project_name,
                                activity_code=settings.meeting_activity_code,
                                location=settings.default_location,
                                description=event.to_timesheet_description(),
                                status=EntryStatus.DRAFT,
                                source=EntrySource.CALENDAR,
                                source_event_id=event.id,
                            )
                            entries_to_add.append(entry)

                    if entries_to_add:
                        for entry in entries_to_add:
                            store.save_entry(entry)
                        st.success(f"Added {len(entries_to_add)} entries!")
                        st.rerun()
                    else:
                        st.warning("No events selected")
        else:
            st.info("All events are matched!")


def get_month_range(target_date: date = None) -> tuple[date, date]:
    """Get the first and last day of the month for a given date."""
    if target_date is None:
        target_date = date.today()
    first_day = target_date.replace(day=1)
    # Get last day by going to next month and subtracting a day
    if first_day.month == 12:
        last_day = first_day.replace(year=first_day.year + 1, month=1, day=1) - timedelta(days=1)
    else:
        last_day = first_day.replace(month=first_day.month + 1, day=1) - timedelta(days=1)
    return first_day, last_day


def show_timesheet():
    """Timesheet management page."""
    st.title("Timesheet")

    store = st.session_state.store
    settings = st.session_state.settings
    timesheet_service = get_timesheet_service()

    # Month selector
    today = date.today()
    months = ["January", "February", "March", "April", "May", "June",
              "July", "August", "September", "October", "November", "December"]

    col1, col2 = st.columns(2)
    with col1:
        selected_month = st.selectbox("Month", months, index=today.month - 1)
    with col2:
        selected_year = st.selectbox("Year", range(today.year - 1, today.year + 2), index=1)

    # Calculate date range for selected month
    month_idx = months.index(selected_month) + 1
    month_start, month_end = get_month_range(date(selected_year, month_idx, 1))

    st.caption(f"Showing entries from {month_start.strftime('%d.%m.%Y')} to {month_end.strftime('%d.%m.%Y')}")

    # Tabs
    tab1, tab2, tab3 = st.tabs(["📋 Entries", "➕ Add Entry", "📤 Export"])

    with tab1:
        # Get entries for selected month
        entries = store.get_entries(start_date=month_start, end_date=month_end)

        # Add generated entries if any
        if st.session_state.generated_entries:
            st.info(f"You have {len(st.session_state.generated_entries)} unsaved generated entries.")
            if st.button("Save Generated Entries"):
                for entry in st.session_state.generated_entries:
                    store.save_entry(entry)
                st.session_state.generated_entries = []
                st.success("Entries saved!")
                st.rerun()

            if st.button("Discard Generated Entries"):
                st.session_state.generated_entries = []
                st.rerun()

            # Show generated entries
            st.subheader("Generated Entries (Unsaved)")
            gen_df = pd.DataFrame([e.to_export_row() for e in st.session_state.generated_entries])
            st.dataframe(gen_df, use_container_width=True, hide_index=True)

        # Show saved entries
        st.subheader("Saved Entries")
        if entries:
            entry_data = []
            for e in entries:
                row = e.to_export_row()
                row["ID"] = e.id
                row["Status"] = e.status.value
                entry_data.append(row)

            df = pd.DataFrame(entry_data)

            # Editable dataframe
            edited_df = st.data_editor(
                df,
                use_container_width=True,
                hide_index=True,
                disabled=["ID", "Status"],
                column_config={
                    "ID": st.column_config.TextColumn("ID", width="small"),
                },
            )

            col1, col2 = st.columns(2)
            with col1:
                if st.button("💾 Save Changes"):
                    for idx, row in edited_df.iterrows():
                        entry = store.get_entry(row["ID"])
                        if entry:
                            entry.description = row["Activity"]
                            entry.hours = int(row["Workhour"])
                            entry.ticket_number = row["Ticket No"] if row["Ticket No"] else None
                            entry.project_name = row["Project"] if row["Project"] else None
                            store.save_entry(entry)
                    st.success("Changes saved!")

            with col2:
                selected_ids = st.multiselect(
                    "Select entries to delete",
                    options=[e.id for e in entries],
                    format_func=lambda x: next(
                        (f"{e.entry_date} - {e.description[:30]}..." for e in entries if e.id == x),
                        x
                    ),
                )
                if selected_ids and st.button("🗑️ Delete Selected"):
                    for eid in selected_ids:
                        store.delete_entry(eid)
                    st.success("Entries deleted!")
                    st.rerun()
        else:
            st.info("No entries for the selected date range.")

    with tab2:
        st.subheader("Add Manual Entry")

        with st.form("add_entry"):
            entry_date = st.date_input("Date", value=date.today())
            hours = st.number_input("Hours", min_value=1, max_value=24, value=8)
            description = st.text_input("Description", placeholder="Work description")

            # Project selection
            projects = store.get_projects()
            project_options = [""] + [p.name for p in projects]
            selected_project = st.selectbox("Project", project_options)

            # Show ticket number for selected project
            ticket_number = ""
            if selected_project:
                project = next((p for p in projects if p.name == selected_project), None)
                if project:
                    ticket_number = project.ticket_number
                    st.caption(f"Ticket No: {ticket_number}" + (f" (JIRA: {project.jira_key})" if project.jira_key else ""))

            activity_code = st.selectbox(
                "Activity Code",
                options=[code.value for code in ActivityCode],
                index=0,
            )
            location = st.text_input("Location", value=settings.default_location)

            if st.form_submit_button("Add Entry"):
                if description:
                    entry = timesheet_service.create_manual_entry(
                        entry_date=entry_date,
                        hours=hours,
                        description=description,
                        project_name=selected_project if selected_project else None,
                        ticket_number=ticket_number if ticket_number else None,
                        activity_code=ActivityCode(activity_code),
                        location=location,
                        settings=settings,
                    )
                    store.save_entry(entry)
                    st.success("Entry added!")
                    st.rerun()
                else:
                    st.error("Description is required")

    with tab3:
        st.subheader(f"Export Timesheet - {selected_month} {selected_year}")

        # Get entries for export (uses the month selected at top)
        export_entries = store.get_entries(start_date=month_start, end_date=month_end)

        if export_entries:
            st.write(f"Found {len(export_entries)} entries to export.")

            # Preview
            with st.expander("Preview Export", expanded=False):
                preview_df = pd.DataFrame([e.to_export_row() for e in export_entries])
                st.dataframe(preview_df, use_container_width=True, hide_index=True)

            # Export options
            aggregate = st.checkbox(
                "Aggregate entries by date/project/ticket",
                value=False,
                help="Combine hours for similar entries on the same day"
            )

            if aggregate:
                export_entries = timesheet_service.aggregate_entries(export_entries)
                st.info(f"Aggregated to {len(export_entries)} entries.")

            # Export button
            if st.button("📥 Download Excel", type="primary", use_container_width=True):
                exporter = ExcelExporter()

                # Create in-memory file
                output = BytesIO()
                with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
                    exporter.export_to_new_file(export_entries, tmp.name)
                    with open(tmp.name, "rb") as f:
                        output.write(f.read())
                    Path(tmp.name).unlink()

                output.seek(0)

                # Generate filename
                filename = f"timesheet_{export_start.strftime('%Y%m%d')}_{export_end.strftime('%Y%m%d')}.xlsx"

                st.download_button(
                    label="📥 Download",
                    data=output,
                    file_name=filename,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )

                # Mark as exported
                store.mark_entries_exported([e.id for e in export_entries])
                st.success("Timesheet exported!")
        else:
            st.info("No entries to export for the selected date range.")


def show_settings():
    """Settings page."""
    st.title("Settings")

    store = st.session_state.store
    settings = st.session_state.settings

    with st.form("settings"):
        st.subheader("User Settings")
        consultant_id = st.text_input("Consultant ID", value=settings.consultant_id)
        default_location = st.text_input("Default Location", value=settings.default_location)

        st.subheader("Activity Codes")
        default_activity = st.selectbox(
            "Default Activity Code",
            options=[code.value for code in ActivityCode],
            index=[code.value for code in ActivityCode].index(settings.default_activity_code.value),
        )
        meeting_activity = st.selectbox(
            "Meeting Activity Code",
            options=[code.value for code in ActivityCode],
            index=[code.value for code in ActivityCode].index(settings.meeting_activity_code.value),
        )

        st.subheader("Matching Settings")
        min_confidence = st.slider(
            "Minimum Match Confidence",
            0.0, 1.0,
            value=settings.min_match_confidence,
            step=0.1,
        )
        skip_canceled = st.checkbox(
            "Skip Canceled Meetings",
            value=settings.skip_canceled_meetings,
        )
        min_duration = st.number_input(
            "Minimum Meeting Duration (minutes)",
            min_value=0,
            max_value=60,
            value=settings.min_meeting_duration_minutes,
        )

        if st.form_submit_button("Save Settings"):
            new_settings = UserSettings(
                consultant_id=consultant_id,
                default_location=default_location,
                default_activity_code=ActivityCode(default_activity),
                meeting_activity_code=ActivityCode(meeting_activity),
                min_match_confidence=min_confidence,
                skip_canceled_meetings=skip_canceled,
                min_meeting_duration_minutes=min_duration,
            )
            store.save_settings(new_settings)
            st.session_state.settings = new_settings
            st.success("Settings saved!")


if __name__ == "__main__":
    main()
