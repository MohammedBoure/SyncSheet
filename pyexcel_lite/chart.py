"""Lightweight chart drawing widgets for selected spreadsheet data."""

from __future__ import annotations

from dataclasses import dataclass

from PySide6.QtCore import QPointF, QRectF, Qt
from PySide6.QtGui import QColor, QFont, QPainter, QPen
from PySide6.QtWidgets import QSizePolicy, QWidget


@dataclass(frozen=True)
class ChartPoint:
    label: str
    value: float


class ChartWidget(QWidget):
    def __init__(self):
        super().__init__()
        self.points: list[ChartPoint] = []
        self.chart_type = "Bar"
        self.title = "Chart"
        self.setMinimumHeight(220)
        self.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)

    def set_chart(self, points: list[ChartPoint], chart_type: str = "Bar", title: str = "Chart") -> None:
        self.points = points
        self.chart_type = chart_type
        self.title = title or "Chart"
        self.update()

    def paintEvent(self, _event) -> None:
        painter = QPainter(self)
        painter.setRenderHint(QPainter.Antialiasing, True)
        painter.fillRect(self.rect(), QColor("#ffffff"))
        painter.setPen(QColor("#1f2937"))
        painter.setFont(QFont("Segoe UI", 10, QFont.Bold))
        painter.drawText(QRectF(8, 6, self.width() - 16, 24), Qt.AlignLeft | Qt.AlignVCenter, self.title)
        chart_rect = QRectF(10, 38, self.width() - 20, self.height() - 48)
        if not self.points:
            painter.setFont(QFont("Segoe UI", 9))
            painter.setPen(QColor("#6b7280"))
            painter.drawText(chart_rect, Qt.AlignCenter, "Select numeric cells to create a chart")
            painter.end()
            return
        if self.chart_type == "Line":
            self._draw_line(painter, chart_rect)
        elif self.chart_type == "Pie":
            self._draw_pie(painter, chart_rect)
        else:
            self._draw_bar(painter, chart_rect)
        painter.end()

    def _draw_axes(self, painter: QPainter, rect: QRectF) -> QRectF:
        plot = rect.adjusted(28, 8, -8, -28)
        painter.setPen(QPen(QColor("#cbd5e1"), 1))
        painter.drawLine(plot.bottomLeft(), plot.bottomRight())
        painter.drawLine(plot.bottomLeft(), plot.topLeft())
        return plot

    def _scale(self, value: float, minimum: float, maximum: float, height: float) -> float:
        if maximum == minimum:
            return height / 2
        return height - ((value - minimum) / (maximum - minimum)) * height

    def _draw_bar(self, painter: QPainter, rect: QRectF) -> None:
        points = self.points[:30]
        values = [point.value for point in points]
        minimum = min(0, min(values))
        maximum = max(values)
        plot = self._draw_axes(painter, rect)
        gap = 5
        width = max(2, (plot.width() - gap * (len(points) + 1)) / max(1, len(points)))
        zero_y = plot.top() + self._scale(0, minimum, maximum, plot.height())
        colors = ["#107c41", "#2563eb", "#d97706", "#7c3aed", "#d92d20"]
        for index, point in enumerate(points):
            x = plot.left() + gap + index * (width + gap)
            y = plot.top() + self._scale(point.value, minimum, maximum, plot.height())
            top = min(y, zero_y)
            height = max(2, abs(zero_y - y))
            painter.setPen(Qt.NoPen)
            painter.setBrush(QColor(colors[index % len(colors)]))
            painter.drawRoundedRect(QRectF(x, top, width, height), 2, 2)
        painter.setPen(QColor("#64748b"))
        painter.setFont(QFont("Segoe UI", 7))
        for index, point in enumerate(points[:8]):
            x = plot.left() + gap + index * (width + gap)
            painter.drawText(QRectF(x - 8, plot.bottom() + 3, width + 16, 18), Qt.AlignCenter, point.label[:8])

    def _draw_line(self, painter: QPainter, rect: QRectF) -> None:
        points = self.points[:100]
        values = [point.value for point in points]
        minimum = min(values)
        maximum = max(values)
        plot = self._draw_axes(painter, rect)
        if len(points) == 1:
            x_step = 1
        else:
            x_step = plot.width() / (len(points) - 1)
        coordinates: list[QPointF] = []
        for index, point in enumerate(points):
            x = plot.left() + index * x_step
            y = plot.top() + self._scale(point.value, minimum, maximum, plot.height())
            coordinates.append(QPointF(x, y))
        painter.setPen(QPen(QColor("#107c41"), 2))
        for first, second in zip(coordinates, coordinates[1:]):
            painter.drawLine(first, second)
        painter.setBrush(QColor("#107c41"))
        painter.setPen(Qt.NoPen)
        for point in coordinates:
            painter.drawEllipse(point, 3, 3)

    def _draw_pie(self, painter: QPainter, rect: QRectF) -> None:
        points = [point for point in self.points[:12] if point.value > 0]
        total = sum(point.value for point in points)
        if total <= 0:
            return
        diameter = min(rect.width(), rect.height()) - 12
        pie_rect = QRectF(rect.left() + 8, rect.top() + 8, diameter, diameter)
        colors = ["#107c41", "#2563eb", "#d97706", "#7c3aed", "#d92d20", "#0891b2"]
        start_angle = 0
        for index, point in enumerate(points):
            span_angle = int(point.value / total * 360 * 16)
            painter.setBrush(QColor(colors[index % len(colors)]))
            painter.setPen(Qt.NoPen)
            painter.drawPie(pie_rect, start_angle, span_angle)
            start_angle += span_angle
        painter.setFont(QFont("Segoe UI", 8))
        legend_x = pie_rect.right() + 12
        legend_y = pie_rect.top()
        for index, point in enumerate(points[:6]):
            y = legend_y + index * 20
            painter.setBrush(QColor(colors[index % len(colors)]))
            painter.drawRect(QRectF(legend_x, y + 4, 10, 10))
            painter.setPen(QColor("#1f2937"))
            painter.drawText(QRectF(legend_x + 15, y, rect.right() - legend_x - 18, 18), Qt.AlignLeft | Qt.AlignVCenter, point.label[:14])
