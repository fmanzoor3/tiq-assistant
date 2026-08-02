"""Regression tests for locale-safe Outlook date-range filtering.

These guard the fix for the bug where some months (e.g. May, June) returned
no meetings because the Restrict date string was parsed ambiguously by locale.
The correctness guarantee now lives in the Python-side filter in _scan_items,
which these tests exercise with a fake COM items collection.
"""

from datetime import datetime, date

from tiq_assistant.integrations.outlook_reader import OutlookReader


class _FakeDT:
    def __init__(self, dt):
        self.year, self.month, self.day = dt.year, dt.month, dt.day
        self.hour, self.minute, self.second = dt.hour, dt.minute, dt.second


class _FakeItem:
    def __init__(self, subject, start, end):
        self.Subject = subject
        self.Start = _FakeDT(start)
        self.End = _FakeDT(end)
        self.AllDayEvent = False
        self.Location = ""
        self.Body = ""
        self.IsOnlineMeeting = False
        self.IsRecurring = False
        self.Organizer = "organizer"


class _FakeItems:
    IncludeRecurrences = True

    def __init__(self, items):
        self._items = sorted(
            items,
            key=lambda i: (i.Start.year, i.Start.month, i.Start.day, i.Start.hour),
        )
        self._i = 0

    def Sort(self, *a, **k):
        pass

    def Restrict(self, _r):
        # Worst case: pretend Restrict returns everything (as a bad locale
        # filter effectively would). The Python filter must still be correct.
        return self

    def GetFirst(self):
        self._i = 0
        return self._items[0] if self._items else None

    def GetNext(self):
        self._i += 1
        return self._items[self._i] if self._i < len(self._items) else None


def _reader_with(items):
    r = OutlookReader.__new__(OutlookReader)
    r._calendar = type("C", (), {"Items": _FakeItems(items)})()
    return r


ALL_ITEMS = [
    _FakeItem("April", datetime(2026, 4, 15, 10), datetime(2026, 4, 15, 11)),
    _FakeItem("May-1", datetime(2026, 5, 4, 9), datetime(2026, 5, 4, 10)),
    _FakeItem("May-2", datetime(2026, 5, 20, 14), datetime(2026, 5, 20, 15)),
    _FakeItem("June", datetime(2026, 6, 2, 9), datetime(2026, 6, 2, 10)),
]


def test_may_range_returns_only_may():
    r = _reader_with(ALL_ITEMS)
    res = r._scan_items(datetime(2026, 5, 1), datetime(2026, 6, 1), restriction="[Start] >= 'x'")
    assert sorted(m.subject for m in res) == ["May-1", "May-2"]


def test_june_range_returns_only_june():
    r = _reader_with(ALL_ITEMS)
    res = r._scan_items(datetime(2026, 6, 1), datetime(2026, 7, 1), restriction=None)
    assert [m.subject for m in res] == ["June"]


def test_format_outlook_date_has_four_digit_year():
    for d in [date(2026, 5, 1), date(2026, 6, 1), date(2026, 1, 3)]:
        s = OutlookReader._format_outlook_date(d)
        assert "2026" in s
