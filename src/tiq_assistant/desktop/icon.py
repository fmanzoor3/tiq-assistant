"""App icon creation for TIQ Assistant."""

from PyQt6.QtGui import QIcon, QPixmap, QPainter, QColor, QPen, QBrush
from PyQt6.QtCore import Qt


def create_app_icon() -> QIcon:
    """Create the custom clock icon for TIQ Assistant app.

    This can be used for the system tray, main window, and dialogs.
    Returns a teal clock icon that represents the timesheet app.
    """
    # Create a 64x64 pixmap for good quality on high-DPI displays
    size = 64
    pixmap = QPixmap(size, size)
    pixmap.fill(Qt.GlobalColor.transparent)

    painter = QPainter(pixmap)
    painter.setRenderHint(QPainter.RenderHint.Antialiasing)

    # Colors
    teal = QColor("#0D9488")  # Teal color matching the app theme
    white = QColor("#FFFFFF")
    dark_teal = QColor("#0F766E")

    center = size // 2
    radius = (size // 2) - 4

    # Draw clock face (filled circle)
    painter.setPen(QPen(dark_teal, 3))
    painter.setBrush(QBrush(teal))
    painter.drawEllipse(center - radius, center - radius, radius * 2, radius * 2)

    # Draw clock hands
    painter.setPen(QPen(white, 4, Qt.PenStyle.SolidLine, Qt.PenCapStyle.RoundCap))

    # Hour hand (pointing to ~10 o'clock position)
    painter.drawLine(center, center, center - 10, center - 14)

    # Minute hand (pointing to ~2 o'clock position)
    painter.setPen(QPen(white, 3, Qt.PenStyle.SolidLine, Qt.PenCapStyle.RoundCap))
    painter.drawLine(center, center, center + 14, center - 10)

    # Draw center dot
    painter.setPen(Qt.PenStyle.NoPen)
    painter.setBrush(QBrush(white))
    painter.drawEllipse(center - 3, center - 3, 6, 6)

    painter.end()

    return QIcon(pixmap)
