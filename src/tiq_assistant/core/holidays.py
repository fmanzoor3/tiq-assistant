"""Turkish national holidays service for TIQ Assistant.

This module manages national holidays (Ulusal Bayram ve Genel Tatil Günleri)
that should be excluded from workday calculations. It also handles half-day
holidays where only 4 hours of work are expected instead of 8.

Based on the 2026 Enerjisa holiday calendar.
"""

from dataclasses import dataclass
from datetime import date
from enum import Enum
from typing import Optional


class HolidayType(Enum):
    """Type of holiday."""
    FULL_DAY = "full_day"  # No work expected (0 hours)
    HALF_DAY = "half_day"  # Only 4 hours of work expected


@dataclass
class Holiday:
    """A national holiday."""
    date: date
    name: str
    holiday_type: HolidayType = HolidayType.FULL_DAY

    @property
    def expected_hours(self) -> int:
        """Get expected work hours for this holiday."""
        if self.holiday_type == HolidayType.FULL_DAY:
            return 0
        return 4  # Half day


# 2026 Turkish National Holidays (Ulusal Bayram ve Genel Tatil Günleri)
# Only includes official holidays, NOT Ortak Tatil or Köprü İzin (optional)
HOLIDAYS_2026 = [
    # Yılbaşı (New Year)
    Holiday(date(2026, 1, 1), "Yılbaşı"),

    # Ramazan Bayramı (19-22 Mart)
    # 19 Mart is half day (Perşembe 0,5 Gün)
    Holiday(date(2026, 3, 19), "Ramazan Bayramı Arifesi", HolidayType.HALF_DAY),
    Holiday(date(2026, 3, 20), "Ramazan Bayramı 1. Gün"),
    Holiday(date(2026, 3, 21), "Ramazan Bayramı 2. Gün"),
    Holiday(date(2026, 3, 22), "Ramazan Bayramı 3. Gün"),

    # Ulusal Egemenlik ve Çocuk Bayramı (23 Nisan)
    Holiday(date(2026, 4, 23), "Ulusal Egemenlik ve Çocuk Bayramı"),

    # Emek ve Dayanışma Günü (1 Mayıs)
    Holiday(date(2026, 5, 1), "Emek ve Dayanışma Günü"),

    # Atatürk'ü Anma, Gençlik ve Spor Bayramı (19 Mayıs)
    Holiday(date(2026, 5, 19), "Atatürk'ü Anma, Gençlik ve Spor Bayramı"),

    # Kurban Bayramı (26-30 Mayıs)
    # 25 Mayıs is half day (arife), 26-30 are full holidays
    Holiday(date(2026, 5, 25), "Kurban Bayramı Arifesi", HolidayType.HALF_DAY),
    Holiday(date(2026, 5, 26), "Kurban Bayramı 1. Gün"),
    Holiday(date(2026, 5, 27), "Kurban Bayramı 2. Gün"),
    Holiday(date(2026, 5, 28), "Kurban Bayramı 3. Gün"),
    Holiday(date(2026, 5, 29), "Kurban Bayramı 4. Gün"),
    Holiday(date(2026, 5, 30), "Kurban Bayramı 5. Gün"),

    # Demokrasi ve Milli Birlik Günü (15 Temmuz)
    Holiday(date(2026, 7, 15), "Demokrasi ve Milli Birlik Günü"),

    # Zafer Bayramı (30 Ağustos)
    Holiday(date(2026, 8, 30), "Zafer Bayramı"),

    # Cumhuriyet Bayramı (28-29 Ekim)
    # 28 Ekim is half day (1,5 gün means 28th afternoon + 29th full)
    Holiday(date(2026, 10, 28), "Cumhuriyet Bayramı Arifesi", HolidayType.HALF_DAY),
    Holiday(date(2026, 10, 29), "Cumhuriyet Bayramı"),
]


class HolidayService:
    """Service for managing national holidays."""

    def __init__(self, holidays: Optional[list[Holiday]] = None):
        """
        Initialize the holiday service.

        Args:
            holidays: List of holidays. If None, uses default 2026 holidays.
        """
        self._holidays = holidays or HOLIDAYS_2026
        # Build lookup dict for fast access
        self._holiday_map: dict[date, Holiday] = {h.date: h for h in self._holidays}

    def is_holiday(self, check_date: date) -> bool:
        """Check if a date is a national holiday (full or half day)."""
        return check_date in self._holiday_map

    def is_full_day_holiday(self, check_date: date) -> bool:
        """Check if a date is a full-day holiday (no work expected)."""
        holiday = self._holiday_map.get(check_date)
        return holiday is not None and holiday.holiday_type == HolidayType.FULL_DAY

    def is_half_day_holiday(self, check_date: date) -> bool:
        """Check if a date is a half-day holiday (4 hours work expected)."""
        holiday = self._holiday_map.get(check_date)
        return holiday is not None and holiday.holiday_type == HolidayType.HALF_DAY

    def get_holiday(self, check_date: date) -> Optional[Holiday]:
        """Get the holiday for a date, or None if not a holiday."""
        return self._holiday_map.get(check_date)

    def get_expected_hours(self, check_date: date) -> int:
        """
        Get expected work hours for a date.

        Returns:
            8 for regular workdays
            4 for half-day holidays
            0 for full-day holidays or weekends
        """
        # Check weekends first
        if check_date.weekday() >= 5:  # Saturday = 5, Sunday = 6
            return 0

        # Check holidays
        holiday = self._holiday_map.get(check_date)
        if holiday:
            return holiday.expected_hours

        # Regular workday
        return 8

    def is_workday(self, check_date: date) -> bool:
        """
        Check if a date is a workday (not weekend, not full-day holiday).

        Half-day holidays ARE considered workdays (with reduced hours).
        """
        # Weekend check
        if check_date.weekday() >= 5:
            return False

        # Full-day holiday check
        if self.is_full_day_holiday(check_date):
            return False

        return True

    def get_workdays_in_month(self, year: int, month: int) -> list[tuple[date, int]]:
        """
        Get all workdays in a month with their expected hours.

        Args:
            year: The year
            month: The month (1-12)

        Returns:
            List of (date, expected_hours) tuples for each workday
        """
        from calendar import monthrange

        _, days_in_month = monthrange(year, month)
        workdays = []

        for day in range(1, days_in_month + 1):
            d = date(year, month, day)
            if self.is_workday(d):
                workdays.append((d, self.get_expected_hours(d)))

        return workdays

    def get_total_expected_hours_in_month(self, year: int, month: int) -> int:
        """Get the total expected work hours for a month."""
        workdays = self.get_workdays_in_month(year, month)
        return sum(hours for _, hours in workdays)

    def get_holidays_in_range(self, start_date: date, end_date: date) -> list[Holiday]:
        """Get all holidays within a date range."""
        return [
            h for h in self._holidays
            if start_date <= h.date <= end_date
        ]


# Global singleton instance
_holiday_service: Optional[HolidayService] = None


def get_holiday_service() -> HolidayService:
    """Get the global holiday service instance."""
    global _holiday_service
    if _holiday_service is None:
        _holiday_service = HolidayService()
    return _holiday_service
