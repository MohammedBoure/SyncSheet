"""Small vector icon set drawn with Qt for the spreadsheet ribbon."""

from __future__ import annotations

from functools import lru_cache

from PySide6.QtCore import QPointF, QRectF, Qt
from PySide6.QtGui import QColor, QFont, QIcon, QPainter, QPainterPath, QPen, QPixmap, QPolygonF

GREEN = "#107c41"
GREEN_DARK = "#0b5f32"
GREEN_LIGHT = "#dff3e8"
INK = "#1f2937"
MUTED = "#6b7280"
RED = "#d92d20"
BLUE = "#2563eb"
AMBER = "#d97706"


@lru_cache(maxsize=128)
def app_icon(name: str, size: int = 28) -> QIcon:
    pixmap = QPixmap(size, size)
    pixmap.fill(Qt.transparent)
    painter = QPainter(pixmap)
    painter.setRenderHint(QPainter.Antialiasing, True)
    painter.setRenderHint(QPainter.TextAntialiasing, True)
    _draw_icon(painter, name, size)
    painter.end()
    return QIcon(pixmap)


def _pen(color: str, width: float = 1.8) -> QPen:
    pen = QPen(QColor(color), width)
    pen.setCapStyle(Qt.RoundCap)
    pen.setJoinStyle(Qt.RoundJoin)
    return pen


def _rounded(p: QPainter, rect: QRectF, fill: str, stroke: str = GREEN, width: float = 1.5) -> None:
    p.setPen(_pen(stroke, width))
    p.setBrush(QColor(fill))
    p.drawRoundedRect(rect, 4, 4)


def _plus(p: QPainter, center: QPointF, radius: float = 4.2, color: str = GREEN_DARK) -> None:
    p.setPen(_pen(color, 2.1))
    p.drawLine(QPointF(center.x() - radius, center.y()), QPointF(center.x() + radius, center.y()))
    p.drawLine(QPointF(center.x(), center.y() - radius), QPointF(center.x(), center.y() + radius))


def _x(p: QPainter, center: QPointF, radius: float = 4.0, color: str = RED) -> None:
    p.setPen(_pen(color, 2.1))
    p.drawLine(QPointF(center.x() - radius, center.y() - radius), QPointF(center.x() + radius, center.y() + radius))
    p.drawLine(QPointF(center.x() + radius, center.y() - radius), QPointF(center.x() - radius, center.y() + radius))


def _sheet(p: QPainter, rect: QRectF) -> None:
    _rounded(p, rect, "#ffffff", GREEN)
    p.setPen(_pen("#9ca3af", 1))
    for offset in (0.35, 0.65):
        x = rect.left() + rect.width() * offset
        p.drawLine(QPointF(x, rect.top() + 3), QPointF(x, rect.bottom() - 3))
    for offset in (0.38, 0.68):
        y = rect.top() + rect.height() * offset
        p.drawLine(QPointF(rect.left() + 3, y), QPointF(rect.right() - 3, y))


def _document(p: QPainter, rect: QRectF) -> None:
    path = QPainterPath()
    path.moveTo(rect.left(), rect.top())
    path.lineTo(rect.right() - 6, rect.top())
    path.lineTo(rect.right(), rect.top() + 6)
    path.lineTo(rect.right(), rect.bottom())
    path.lineTo(rect.left(), rect.bottom())
    path.closeSubpath()
    p.setPen(_pen(GREEN, 1.6))
    p.setBrush(QColor("#ffffff"))
    p.drawPath(path)
    p.setBrush(QColor(GREEN_LIGHT))
    p.drawPolygon(
        QPolygonF(
            [
                QPointF(rect.right() - 6, rect.top()),
                QPointF(rect.right(), rect.top() + 6),
                QPointF(rect.right() - 6, rect.top() + 6),
            ]
        )
    )


def _magnifier(p: QPainter, plus: bool | None) -> None:
    p.setPen(_pen(INK, 2))
    p.setBrush(Qt.NoBrush)
    p.drawEllipse(QRectF(6, 5, 12, 12))
    p.drawLine(QPointF(16, 16), QPointF(23, 23))
    if plus is None:
        p.setPen(_pen(GREEN, 2))
        p.drawArc(QRectF(6, 5, 12, 12), 30 * 16, 260 * 16)
        return
    p.setPen(_pen(GREEN if plus else RED, 2))
    p.drawLine(QPointF(9.5, 11), QPointF(14.5, 11))
    if plus:
        p.drawLine(QPointF(12, 8.5), QPointF(12, 13.5))


def _draw_icon(p: QPainter, name: str, size: int) -> None:
    p.scale(size / 28, size / 28)
    if name == "new":
        _document(p, QRectF(7, 4, 14, 19))
        _plus(p, QPointF(20, 20), 3.8)
    elif name == "open":
        p.setPen(_pen(GREEN, 1.6))
        p.setBrush(QColor(GREEN_LIGHT))
        p.drawRoundedRect(QRectF(3, 9, 21, 13), 3, 3)
        p.setBrush(QColor("#ffffff"))
        p.drawRoundedRect(QRectF(5, 6, 9, 6), 2, 2)
        p.setBrush(QColor("#f8fafc"))
        p.drawPolygon(QPolygonF([QPointF(5, 12), QPointF(25, 12), QPointF(21, 23), QPointF(3, 23)]))
    elif name in {"save", "save_as"}:
        _rounded(p, QRectF(5, 4, 18, 20), "#ffffff", GREEN)
        p.setBrush(QColor(GREEN))
        p.setPen(Qt.NoPen)
        p.drawRect(QRectF(9, 5.5, 9, 5))
        p.setBrush(QColor(GREEN_LIGHT))
        p.drawRoundedRect(QRectF(9, 15, 10, 6), 2, 2)
        if name == "save_as":
            p.setPen(_pen(AMBER, 2))
            p.drawLine(QPointF(18, 22), QPointF(24, 16))
    elif name == "csv":
        _sheet(p, QRectF(4, 5, 20, 18))
        p.setFont(QFont("Segoe UI", 7, QFont.Bold))
        p.setPen(QColor(GREEN_DARK))
        p.drawText(QRectF(4, 10, 20, 10), Qt.AlignCenter, "CSV")
    elif name == "sheet_add":
        _sheet(p, QRectF(4, 5, 18, 17))
        _plus(p, QPointF(22, 21), 3.7)
    elif name == "sheet_rename":
        _sheet(p, QRectF(4, 5, 18, 17))
        p.setPen(_pen(AMBER, 2))
        p.drawLine(QPointF(17, 23), QPointF(24, 16))
    elif name == "sheet_delete":
        _sheet(p, QRectF(4, 5, 18, 17))
        _x(p, QPointF(22, 21), 3.7)
    elif name == "row_insert":
        _sheet(p, QRectF(4, 5, 20, 18))
        p.setPen(_pen(GREEN, 2.2))
        p.drawLine(QPointF(7, 14), QPointF(21, 14))
        _plus(p, QPointF(22, 7), 3)
    elif name == "row_delete":
        _sheet(p, QRectF(4, 5, 20, 18))
        p.setPen(_pen(RED, 2.2))
        p.drawLine(QPointF(7, 14), QPointF(21, 14))
        _x(p, QPointF(22, 7), 3)
    elif name == "column_insert":
        _sheet(p, QRectF(4, 5, 20, 18))
        p.setPen(_pen(GREEN, 2.2))
        p.drawLine(QPointF(14, 8), QPointF(14, 20))
        _plus(p, QPointF(22, 7), 3)
    elif name == "column_delete":
        _sheet(p, QRectF(4, 5, 20, 18))
        p.setPen(_pen(RED, 2.2))
        p.drawLine(QPointF(14, 8), QPointF(14, 20))
        _x(p, QPointF(22, 7), 3)
    elif name == "clear":
        p.setPen(_pen(INK, 1.6))
        p.setBrush(QColor("#ffffff"))
        p.drawRoundedRect(QRectF(7, 8, 15, 10), 3, 3)
        p.setPen(_pen(RED, 2))
        p.drawLine(QPointF(9, 20), QPointF(22, 7))
    elif name in {"undo", "redo"}:
        p.setPen(_pen(GREEN, 2.2))
        if name == "undo":
            p.drawArc(QRectF(6, 8, 16, 12), 30 * 16, 250 * 16)
            p.drawLine(QPointF(8, 9), QPointF(4, 13))
            p.drawLine(QPointF(8, 9), QPointF(12, 13))
        else:
            p.drawArc(QRectF(6, 8, 16, 12), 140 * 16, -250 * 16)
            p.drawLine(QPointF(20, 9), QPointF(24, 13))
            p.drawLine(QPointF(20, 9), QPointF(16, 13))
    elif name == "chart":
        p.setPen(_pen(INK, 1.5))
        p.drawLine(QPointF(5, 22), QPointF(24, 22))
        p.drawLine(QPointF(5, 22), QPointF(5, 5))
        p.setPen(Qt.NoPen)
        p.setBrush(QColor(BLUE))
        p.drawRoundedRect(QRectF(8, 14, 3.5, 7), 1, 1)
        p.setBrush(QColor(GREEN))
        p.drawRoundedRect(QRectF(14, 9, 3.5, 12), 1, 1)
        p.setBrush(QColor(AMBER))
        p.drawRoundedRect(QRectF(20, 12, 3.5, 9), 1, 1)
    elif name == "selection":
        p.setPen(_pen(GREEN, 1.8))
        p.setBrush(QColor("#ffffff"))
        p.drawRoundedRect(QRectF(5, 5, 18, 18), 3, 3)
        p.setPen(_pen("#cbd5e1", 1))
        p.drawLine(QPointF(11, 6), QPointF(11, 22))
        p.drawLine(QPointF(17, 6), QPointF(17, 22))
        p.drawLine(QPointF(6, 11), QPointF(22, 11))
        p.drawLine(QPointF(6, 17), QPointF(22, 17))
        p.setPen(_pen(BLUE, 2.2))
        p.setBrush(Qt.NoBrush)
        p.drawRoundedRect(QRectF(10, 10, 8, 8), 2, 2)
    elif name == "stats":
        p.setPen(_pen(INK, 1.5))
        p.drawLine(QPointF(5, 22), QPointF(24, 22))
        p.drawLine(QPointF(5, 22), QPointF(5, 5))
        p.setPen(_pen(BLUE, 2))
        p.drawLine(QPointF(8, 18), QPointF(12, 13))
        p.drawLine(QPointF(12, 13), QPointF(16, 15))
        p.drawLine(QPointF(16, 15), QPointF(22, 8))
        p.setBrush(QColor("#ffffff"))
        for point in (QPointF(8, 18), QPointF(12, 13), QPointF(16, 15), QPointF(22, 8)):
            p.drawEllipse(point, 2, 2)
    elif name in {"network_host", "network_join", "network_leave"}:
        p.setPen(_pen(GREEN, 1.8))
        p.setBrush(QColor("#ffffff"))
        p.drawEllipse(QRectF(4, 5, 7, 7))
        p.drawEllipse(QRectF(17, 5, 7, 7))
        p.drawEllipse(QRectF(10.5, 17, 7, 7))
        p.drawLine(QPointF(10, 10), QPointF(14, 18))
        p.drawLine(QPointF(18, 10), QPointF(15, 18))
        if name == "network_host":
            _plus(p, QPointF(14, 14), 3.4, BLUE)
        elif name == "network_join":
            p.setPen(_pen(BLUE, 2.1))
            p.drawLine(QPointF(8, 22), QPointF(19, 22))
            p.drawLine(QPointF(19, 22), QPointF(15, 18))
            p.drawLine(QPointF(19, 22), QPointF(15, 26))
        else:
            _x(p, QPointF(14, 14), 3.7, RED)
    elif name == "settings":
        p.setPen(_pen(GREEN, 2))
        p.setBrush(QColor("#ffffff"))
        p.drawEllipse(QRectF(8, 8, 12, 12))
        p.setBrush(QColor(GREEN_LIGHT))
        p.drawEllipse(QRectF(11.5, 11.5, 5, 5))
        p.setPen(_pen(INK, 2))
        for start, end in (
            (QPointF(14, 3), QPointF(14, 7)),
            (QPointF(14, 21), QPointF(14, 25)),
            (QPointF(3, 14), QPointF(7, 14)),
            (QPointF(21, 14), QPointF(25, 14)),
            (QPointF(6, 6), QPointF(9, 9)),
            (QPointF(19, 19), QPointF(22, 22)),
            (QPointF(22, 6), QPointF(19, 9)),
            (QPointF(9, 19), QPointF(6, 22)),
        ):
            p.drawLine(start, end)
    elif name == "zoom_in":
        _magnifier(p, True)
    elif name == "zoom_out":
        _magnifier(p, False)
    elif name == "zoom_reset":
        _magnifier(p, None)
    elif name in {"bold", "italic", "underline"}:
        p.setFont(QFont("Segoe UI", 16, QFont.Bold if name == "bold" else QFont.Normal, italic=name == "italic"))
        p.setPen(QColor(INK))
        letter = "B" if name == "bold" else "I" if name == "italic" else "U"
        p.drawText(QRectF(4, 3, 20, 20), Qt.AlignCenter, letter)
        if name == "underline":
            p.setPen(_pen(GREEN, 2))
            p.drawLine(QPointF(9, 23), QPointF(19, 23))
    elif name == "text_color":
        p.setFont(QFont("Segoe UI", 15, QFont.Bold))
        p.setPen(QColor(INK))
        p.drawText(QRectF(5, 3, 18, 18), Qt.AlignCenter, "A")
        p.setPen(_pen(RED, 3))
        p.drawLine(QPointF(8, 23), QPointF(20, 23))
    elif name == "fill_color":
        p.setPen(_pen(INK, 1.6))
        p.setBrush(QColor("#ffffff"))
        p.drawRoundedRect(QRectF(8, 7, 12, 9), 2, 2)
        p.setBrush(QColor("#fde68a"))
        p.setPen(Qt.NoPen)
        p.drawRoundedRect(QRectF(6, 19, 16, 4), 2, 2)
        p.setPen(_pen(AMBER, 2))
        p.drawLine(QPointF(9, 7), QPointF(20, 18))
    elif name == "formula":
        p.setFont(QFont("Segoe UI", 13, QFont.Bold, italic=True))
        p.setPen(QColor(GREEN_DARK))
        p.drawText(QRectF(3, 4, 22, 19), Qt.AlignCenter, "fx")
    else:
        p.setPen(_pen(GREEN, 2))
        p.drawEllipse(QRectF(6, 6, 16, 16))
