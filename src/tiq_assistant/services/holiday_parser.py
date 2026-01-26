"""Holiday parser service for extracting holidays from PDF/JPG files.

This module processes uploaded holiday calendar files (PDF, JPG, PNG) and
extracts the Turkish national holidays (Ulusal Bayram ve Genel Tatil Günleri).
"""

import re
from datetime import date
from pathlib import Path
from typing import Optional

# Turkish month names for parsing
TURKISH_MONTHS = {
    "ocak": 1, "şubat": 2, "mart": 3, "nisan": 4,
    "mayıs": 5, "haziran": 6, "temmuz": 7, "ağustos": 8,
    "eylül": 9, "ekim": 10, "kasım": 11, "aralık": 12,
    # English fallbacks
    "january": 1, "february": 2, "march": 3, "april": 4,
    "may": 5, "june": 6, "july": 7, "august": 8,
    "september": 9, "october": 10, "november": 11, "december": 12,
}

# Known Turkish national holidays (names to look for)
NATIONAL_HOLIDAY_KEYWORDS = [
    "yılbaşı", "new year",
    "ramazan bayramı", "ramadan",
    "ulusal egemenlik", "çocuk bayramı", "23 nisan",
    "emek ve dayanışma", "1 mayıs", "işçi bayramı",
    "atatürk'ü anma", "gençlik ve spor", "19 mayıs",
    "kurban bayramı", "sacrifice",
    "demokrasi ve milli birlik", "15 temmuz",
    "zafer bayramı", "30 ağustos",
    "cumhuriyet bayramı", "29 ekim",
]

# Keywords indicating half-day
HALF_DAY_KEYWORDS = [
    "arife", "0,5 gün", "0.5 gün", "yarım gün", "half day",
]


class HolidayParseResult:
    """Result of parsing a holiday file."""

    def __init__(self):
        self.holidays: list[tuple[date, str, str]] = []  # (date, name, type)
        self.source_file: str = ""
        self.year: Optional[int] = None
        self.errors: list[str] = []
        self.raw_text: str = ""

    @property
    def count(self) -> int:
        return len(self.holidays)


def extract_text_from_image(file_path: Path) -> str:
    """
    Extract text from an image file using available methods.

    For now, returns empty string - user will manually enter holidays.
    In future, could integrate pytesseract or cloud OCR.
    """
    # TODO: Integrate OCR (pytesseract, Google Cloud Vision, etc.)
    # For now, we'll rely on the pre-defined 2026 holidays
    # and allow manual entry
    return ""


def extract_text_from_pdf(file_path: Path) -> str:
    """
    Extract text from a PDF file.

    For now, returns empty string - user will manually enter holidays.
    In future, could integrate PyPDF2 or pdfplumber.
    """
    # TODO: Integrate PDF text extraction
    return ""


def parse_holiday_text(text: str, year: int = None) -> list[tuple[date, str, str]]:
    """
    Parse holiday information from extracted text.

    Returns list of (date, name, holiday_type) tuples.
    """
    if not year:
        year = date.today().year

    holidays = []
    # TODO: Implement text parsing logic
    # For now, this is a placeholder
    return holidays


def get_default_holidays_for_year(year: int) -> list[tuple[date, str, str]]:
    """
    Get the default Turkish national holidays for a given year.

    Note: Islamic holidays (Ramazan, Kurban) dates change each year
    based on the lunar calendar. These need to be updated annually.
    """
    # Fixed holidays (same date every year)
    fixed_holidays = [
        (date(year, 1, 1), "Yılbaşı", "full_day"),
        (date(year, 4, 23), "Ulusal Egemenlik ve Çocuk Bayramı", "full_day"),
        (date(year, 5, 1), "Emek ve Dayanışma Günü", "full_day"),
        (date(year, 5, 19), "Atatürk'ü Anma, Gençlik ve Spor Bayramı", "full_day"),
        (date(year, 7, 15), "Demokrasi ve Milli Birlik Günü", "full_day"),
        (date(year, 8, 30), "Zafer Bayramı", "full_day"),
        (date(year, 10, 28), "Cumhuriyet Bayramı Arifesi", "half_day"),
        (date(year, 10, 29), "Cumhuriyet Bayramı", "full_day"),
    ]

    # Islamic holidays for 2026 (approximate - need annual update)
    if year == 2026:
        islamic_holidays = [
            # Ramazan Bayramı 2026 (March 19-22)
            (date(2026, 3, 19), "Ramazan Bayramı Arifesi", "half_day"),
            (date(2026, 3, 20), "Ramazan Bayramı 1. Gün", "full_day"),
            (date(2026, 3, 21), "Ramazan Bayramı 2. Gün", "full_day"),
            (date(2026, 3, 22), "Ramazan Bayramı 3. Gün", "full_day"),
            # Kurban Bayramı 2026 (May 26-30)
            (date(2026, 5, 25), "Kurban Bayramı Arifesi", "half_day"),
            (date(2026, 5, 26), "Kurban Bayramı 1. Gün", "full_day"),
            (date(2026, 5, 27), "Kurban Bayramı 2. Gün", "full_day"),
            (date(2026, 5, 28), "Kurban Bayramı 3. Gün", "full_day"),
            (date(2026, 5, 29), "Kurban Bayramı 4. Gün", "full_day"),
            (date(2026, 5, 30), "Kurban Bayramı 5. Gün", "full_day"),
        ]
        return fixed_holidays + islamic_holidays
    elif year == 2025:
        islamic_holidays = [
            # Ramazan Bayramı 2025 (March 30 - April 1)
            (date(2025, 3, 29), "Ramazan Bayramı Arifesi", "half_day"),
            (date(2025, 3, 30), "Ramazan Bayramı 1. Gün", "full_day"),
            (date(2025, 3, 31), "Ramazan Bayramı 2. Gün", "full_day"),
            (date(2025, 4, 1), "Ramazan Bayramı 3. Gün", "full_day"),
            # Kurban Bayramı 2025 (June 6-9)
            (date(2025, 6, 5), "Kurban Bayramı Arifesi", "half_day"),
            (date(2025, 6, 6), "Kurban Bayramı 1. Gün", "full_day"),
            (date(2025, 6, 7), "Kurban Bayramı 2. Gün", "full_day"),
            (date(2025, 6, 8), "Kurban Bayramı 3. Gün", "full_day"),
            (date(2025, 6, 9), "Kurban Bayramı 4. Gün", "full_day"),
        ]
        return fixed_holidays + islamic_holidays

    # For other years, return only fixed holidays
    # Islamic holidays need to be added manually
    return fixed_holidays


def parse_holiday_file(file_path: Path, year: int = None) -> HolidayParseResult:
    """
    Parse a holiday file (PDF, JPG, PNG) and extract holidays.

    Args:
        file_path: Path to the file
        year: Year for the holidays (default: current year)

    Returns:
        HolidayParseResult with extracted holidays
    """
    result = HolidayParseResult()
    result.source_file = file_path.name
    result.year = year or date.today().year

    if not file_path.exists():
        result.errors.append(f"File not found: {file_path}")
        return result

    suffix = file_path.suffix.lower()

    # Extract text based on file type
    if suffix in [".jpg", ".jpeg", ".png", ".bmp", ".gif"]:
        result.raw_text = extract_text_from_image(file_path)
    elif suffix == ".pdf":
        result.raw_text = extract_text_from_pdf(file_path)
    else:
        result.errors.append(f"Unsupported file type: {suffix}")
        return result

    # If we got text, try to parse it
    if result.raw_text:
        result.holidays = parse_holiday_text(result.raw_text, result.year)

    # If no holidays found via OCR, use defaults for the year
    if not result.holidays:
        result.holidays = get_default_holidays_for_year(result.year)
        if not result.raw_text:
            result.errors.append(
                "Could not extract text from file. Using default holidays for the year. "
                "You can manually adjust the holidays below."
            )

    return result
