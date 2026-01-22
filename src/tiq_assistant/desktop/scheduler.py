"""Scheduler for timed popup windows using APScheduler."""

import logging
from datetime import datetime
from typing import Optional, Callable

from PyQt6.QtCore import QObject, pyqtSignal, QTimer
from apscheduler.schedulers.background import BackgroundScheduler
from apscheduler.triggers.cron import CronTrigger
from apscheduler.events import EVENT_JOB_EXECUTED, EVENT_JOB_ERROR

from tiq_assistant.core.models import SessionType, ScheduleConfig
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

    def __init__(self, parent: Optional[QObject] = None):
        super().__init__(parent)

        self._scheduler: Optional[BackgroundScheduler] = None
        self._config: Optional[ScheduleConfig] = None
        self._snoozed_morning: Optional[datetime] = None
        self._snoozed_afternoon: Optional[datetime] = None

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

    def stop(self) -> None:
        """Stop the scheduler."""
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

        # Schedule snooze
        if session == SessionType.MORNING:
            self._scheduler.add_job(
                self._trigger_morning_popup,
                'date',
                run_date=snooze_time,
                id=job_id,
            )
        else:
            self._scheduler.add_job(
                self._trigger_afternoon_popup,
                'date',
                run_date=snooze_time,
                id=job_id,
            )

        logger.info(f"Snoozed {session.value} popup for {self.SNOOZE_DURATION} minutes")

    def _trigger_morning_popup(self) -> None:
        """Trigger the morning popup signal (thread-safe)."""
        # Use QTimer to emit signal on the main thread
        QTimer.singleShot(0, self.morning_popup_due.emit)

    def _trigger_afternoon_popup(self) -> None:
        """Trigger the afternoon popup signal (thread-safe)."""
        QTimer.singleShot(0, self.afternoon_popup_due.emit)

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
