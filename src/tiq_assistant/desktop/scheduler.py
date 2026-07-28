"""Scheduler for timed popup windows using APScheduler.

Reliability notes
-----------------
APScheduler's ``BackgroundScheduler`` alone is not enough on a laptop: if the
machine sleeps through a trigger time (e.g. lunch with the lid closed), the job
is missed and, past the misfire grace window, dropped. To make the popups
actually appear we add two safety nets on top of the cron jobs:

1. A wall-clock ``QTimer`` that ticks every minute and fires any due reminder
   that hasn't been shown yet today. This is immune to sleep/wake drift because
   it compares the real current time to the scheduled time.
2. A one-shot catch-up check on startup, so opening the app after a missed
   reminder still surfaces it.

Each reminder is recorded as "shown" per (date, session) so it fires exactly
once per day regardless of which mechanism triggers it.
"""

import logging
from datetime import datetime, date, time
from typing import Optional, Callable

from PyQt6.QtCore import QObject, pyqtSignal, QTimer
from apscheduler.schedulers.background import BackgroundScheduler
from apscheduler.triggers.cron import CronTrigger
from apscheduler.events import EVENT_JOB_EXECUTED, EVENT_JOB_ERROR

from tiq_assistant.core.models import ScheduleConfig, SessionType
from tiq_assistant.storage.sqlite_store import get_store

logger = logging.getLogger(__name__)


class SchedulerManager(QObject):
    """
    Manages scheduled tasks for the desktop app.

    Uses APScheduler for cron-style scheduling with proper handling of:
    - Laptop sleep/wake scenarios
    - Missed job execution (misfire grace time)
    - Thread-safe Qt signal emission

    Signals:
        morning_popup_due: Emitted when it's time for morning time entry
        afternoon_popup_due: Emitted when it's time for afternoon time entry
    """

    morning_popup_due = pyqtSignal()
    afternoon_popup_due = pyqtSignal()

    # Snooze durations in minutes
    SNOOZE_DURATION = 15

    # How often the wall-clock fallback checks for due reminders.
    FALLBACK_INTERVAL_MS = 60 * 1000  # 1 minute

    def __init__(self, parent: Optional[QObject] = None):
        super().__init__(parent)

        self._scheduler: Optional[BackgroundScheduler] = None
        self._config: Optional[ScheduleConfig] = None
        self._snoozed_morning: Optional[datetime] = None
        self._snoozed_afternoon: Optional[datetime] = None

        # Wall-clock fallback timer (sleep/wake resilient).
        self._fallback_timer: Optional[QTimer] = None

        # Tracks which reminders have already been shown today so each fires
        # once per day: {(date, SessionType)}. Reset when the date rolls over.
        self._shown: set[tuple[date, SessionType]] = set()

    def start(self, config: Optional[ScheduleConfig] = None) -> None:
        """
        Start the scheduler with the given configuration.

        Args:
            config: Schedule configuration. If None, loads from database.
        """
        if self._scheduler is not None and self._scheduler.running:
            self.stop()

        # Load config
        if config is None:
            store = get_store()
            config = store.get_schedule_config()
        self._config = config

        # Create scheduler
        self._scheduler = BackgroundScheduler(
            job_defaults={
                'coalesce': True,  # Combine multiple missed executions
                'misfire_grace_time': 900,  # 15 minutes grace period
            }
        )

        # Add job listeners for logging
        self._scheduler.add_listener(
            self._on_job_event,
            EVENT_JOB_EXECUTED | EVENT_JOB_ERROR
        )

        # Schedule jobs
        self._schedule_jobs()

        # Start the scheduler
        self._scheduler.start()
        logger.info("Scheduler started")

        # Start the wall-clock fallback timer (immune to sleep/wake drift).
        self._start_fallback_timer()

        # Catch up on any reminder already due today (e.g. app opened after the
        # scheduled time, or after waking from sleep).
        self._check_due_reminders(catch_up=True)

    def _start_fallback_timer(self) -> None:
        """Start (or restart) the per-minute wall-clock fallback timer."""
        if self._fallback_timer is None:
            self._fallback_timer = QTimer(self)
            self._fallback_timer.timeout.connect(self._check_due_reminders)
        self._fallback_timer.start(self.FALLBACK_INTERVAL_MS)

    def stop(self) -> None:
        """Stop the scheduler."""
        if self._fallback_timer is not None:
            self._fallback_timer.stop()
        if self._scheduler is not None:
            self._scheduler.shutdown(wait=False)
            self._scheduler = None
            logger.info("Scheduler stopped")

    def reschedule(self, config: ScheduleConfig) -> None:
        """
        Reschedule jobs with new configuration.

        Args:
            config: New schedule configuration
        """
        self._config = config
        if self._scheduler is not None and self._scheduler.running:
            # Remove existing jobs
            self._scheduler.remove_all_jobs()
            # Add jobs with new config
            self._schedule_jobs()
            logger.info("Scheduler rescheduled with new config")

    def snooze_morning(self) -> None:
        """Snooze the morning popup for 15 minutes."""
        self._schedule_snooze(SessionType.MORNING)

    def snooze_afternoon(self) -> None:
        """Snooze the afternoon popup for 15 minutes."""
        self._schedule_snooze(SessionType.AFTERNOON)

    def _schedule_jobs(self) -> None:
        """Schedule the morning and afternoon popup jobs."""
        if self._config is None or self._scheduler is None:
            return

        # Parse times
        morning_hour, morning_min = self._parse_time(self._config.morning_popup_time)
        afternoon_hour, afternoon_min = self._parse_time(self._config.afternoon_popup_time)

        # Schedule morning popup (weekdays only)
        if self._config.morning_popup_enabled:
            self._scheduler.add_job(
                self._trigger_morning_popup,
                CronTrigger(
                    hour=morning_hour,
                    minute=morning_min,
                    day_of_week='mon-fri'
                ),
                id='morning_popup',
                replace_existing=True,
            )
            logger.info(f"Scheduled morning popup at {self._config.morning_popup_time}")

        # Schedule afternoon popup (weekdays only)
        if self._config.afternoon_popup_enabled:
            self._scheduler.add_job(
                self._trigger_afternoon_popup,
                CronTrigger(
                    hour=afternoon_hour,
                    minute=afternoon_min,
                    day_of_week='mon-fri'
                ),
                id='afternoon_popup',
                replace_existing=True,
            )
            logger.info(f"Scheduled afternoon popup at {self._config.afternoon_popup_time}")

    def _schedule_snooze(self, session: SessionType) -> None:
        """Schedule a snoozed reminder."""
        if self._scheduler is None:
            return

        from datetime import timedelta

        snooze_time = datetime.now() + timedelta(minutes=self.SNOOZE_DURATION)
        job_id = f'snooze_{session.value}'

        # Remove existing snooze job if any
        existing = self._scheduler.get_job(job_id)
        if existing:
            self._scheduler.remove_job(job_id)

        # Schedule snooze. Snooze must bypass the once-per-day dedup (the
        # reminder was already shown), so it uses a dedicated re-fire path.
        if session == SessionType.MORNING:
            self._scheduler.add_job(
                self._snooze_fire_morning,
                'date',
                run_date=snooze_time,
                id=job_id,
            )
        else:
            self._scheduler.add_job(
                self._snooze_fire_afternoon,
                'date',
                run_date=snooze_time,
                id=job_id,
            )

        logger.info(f"Snoozed {session.value} popup for {self.SNOOZE_DURATION} minutes")

    def _snooze_fire_morning(self) -> None:
        """Re-fire the morning popup after a snooze (bypasses daily dedup)."""
        QTimer.singleShot(0, self.morning_popup_due.emit)

    def _snooze_fire_afternoon(self) -> None:
        """Re-fire the afternoon popup after a snooze (bypasses daily dedup)."""
        QTimer.singleShot(0, self.afternoon_popup_due.emit)

    def _trigger_morning_popup(self) -> None:
        """Trigger the morning popup signal (thread-safe, deduped, suppressed)."""
        self._fire(SessionType.MORNING)

    def _trigger_afternoon_popup(self) -> None:
        """Trigger the afternoon popup signal (thread-safe, deduped, suppressed)."""
        self._fire(SessionType.AFTERNOON)

    def _fire(self, session: SessionType) -> None:
        """Emit a popup signal once per day, suppressing non-working days.

        Safe to call from either the APScheduler thread or the Qt fallback
        timer -- signal emission is marshalled onto the main thread via
        ``QTimer.singleShot``.
        """
        today = datetime.now().date()

        # Roll over the "shown" set at midnight.
        self._prune_shown(today)

        key = (today, session)
        if key in self._shown:
            return  # Already shown today by cron, fallback, or catch-up.

        if not self._should_fire_today(today):
            logger.info("Suppressing %s popup on %s (non-working day).",
                        session.value, today)
            # Mark as "shown" so we don't re-evaluate every minute.
            self._shown.add(key)
            return

        self._shown.add(key)
        logger.info("Firing %s popup for %s", session.value, today)

        if session == SessionType.MORNING:
            QTimer.singleShot(0, self.morning_popup_due.emit)
        else:
            QTimer.singleShot(0, self.afternoon_popup_due.emit)

    def _should_fire_today(self, today: date) -> bool:
        """Return True if reminders should fire on ``today``.

        Suppresses weekends, full-day national holidays, and days the user has
        explicitly marked as skipped (sick leave, vacation, etc.).
        """
        # Weekend (cron already excludes these, but the fallback timer doesn't).
        if today.weekday() >= 5:
            return False

        try:
            store = get_store()
            is_skipped, _ = store.is_day_skipped(today)
            if is_skipped:
                return False
        except Exception as e:  # noqa: BLE001
            logger.debug("Skip-day check failed: %s", e)

        try:
            from tiq_assistant.core.holidays import get_holiday_service
            if get_holiday_service().is_full_day_holiday(today):
                return False
        except Exception as e:  # noqa: BLE001
            logger.debug("Holiday check failed: %s", e)

        return True

    def _prune_shown(self, today: date) -> None:
        """Drop 'shown' entries from previous days."""
        stale = [k for k in self._shown if k[0] != today]
        for k in stale:
            self._shown.discard(k)

    def _check_due_reminders(self, catch_up: bool = False) -> None:
        """Wall-clock check: fire any reminder whose time has passed today.

        Runs every minute (and once on startup with ``catch_up=True``). Because
        it compares the real current time to the configured popup times, it
        catches reminders missed while the machine was asleep.

        A reminder is considered "due" from its scheduled time until the end of
        the workday, so a lunch reminder missed at 12:15 still appears when the
        laptop wakes at 13:00 -- but not the next morning (the shown-set and the
        date check prevent that).
        """
        if self._config is None:
            return

        now = datetime.now()
        today = now.date()
        self._prune_shown(today)

        if not self._should_fire_today(today):
            return

        current = now.time()

        # Morning reminder: due from its time until lunch ends.
        if self._config.morning_popup_enabled:
            m_time = self._safe_time(self._config.morning_popup_time)
            m_deadline = self._safe_time(self._config.lunch_end, fallback=time(13, 30))
            if m_time is not None and m_time <= current < m_deadline:
                self._fire(SessionType.MORNING)

        # Afternoon reminder: due from its time until the day's end + buffer.
        if self._config.afternoon_popup_enabled:
            a_time = self._safe_time(self._config.afternoon_popup_time)
            a_deadline = self._safe_time(self._config.workday_end, fallback=time(23, 59))
            # Give a generous tail so an end-of-day reminder isn't lost if the
            # workday_end and popup time coincide.
            if a_time is not None and a_time <= current:
                # Only during the same day; treat everything after a_time (until
                # midnight) as still due for catch-up.
                self._fire(SessionType.AFTERNOON)

    def _safe_time(self, time_str: str, fallback: Optional[time] = None) -> Optional[time]:
        """Parse 'HH:MM' into a time, returning ``fallback`` on error."""
        try:
            h, m = self._parse_time(time_str)
            return time(h, m)
        except Exception:
            return fallback

    def _parse_time(self, time_str: str) -> tuple[int, int]:
        """Parse a time string like '12:30' into (hour, minute)."""
        parts = time_str.split(':')
        return int(parts[0]), int(parts[1])

    def _on_job_event(self, event) -> None:
        """Handle scheduler job events."""
        if hasattr(event, 'exception') and event.exception:
            logger.error(f"Scheduler job error: {event.exception}")
        else:
            logger.debug(f"Scheduler job executed: {event.job_id}")

    def get_next_run_times(self) -> dict:
        """Get the next scheduled run times for each job."""
        result = {}
        if self._scheduler is None:
            return result

        for job in self._scheduler.get_jobs():
            if job.next_run_time:
                result[job.id] = job.next_run_time

        return result

    @property
    def is_running(self) -> bool:
        """Check if the scheduler is running."""
        return self._scheduler is not None and self._scheduler.running
