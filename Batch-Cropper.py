from __future__ import annotations

import datetime
import io
import math
import os
import sys
from dataclasses import dataclass
from pathlib import Path

from PIL import Image, ImageGrab, ImageOps
from PySide6.QtCore import QByteArray, QMimeData, QPoint, QRect, QSize, Qt, Signal
from PySide6.QtGui import (
    QAction,
    QColor,
    QCursor,
    QDragEnterEvent,
    QDropEvent,
    QFont,
    QImage,
    QKeySequence,
    QMouseEvent,
    QPainter,
    QPainterPath,
    QPen,
    QPixmap,
    QResizeEvent,
    QWheelEvent,
)
from PySide6.QtWidgets import (
    QApplication,
    QCheckBox,
    QFrame,
    QGridLayout,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QScrollArea,
    QSizePolicy,
    QSplitter,
    QVBoxLayout,
    QWidget,
)


# GUIの主要パラメータ
APP_WINDOW_SIZE = (1120, 800)
WINDOW_RESIZE_SCALE_STEP = 0.1
THUMBNAIL_SIZE = (104, 104)
THUMBNAIL_SELECTED_BORDER_COLOR = QColor(66, 184, 255)
THUMBNAIL_SELECTED_BORDER_WIDTH = 3
RIGHT_PANEL_WIDTH = 282
HANDLE_SIZE = 10
MIN_CROP_SIZE = 4
OVERLAY_ALPHA = 145
MAX_HISTORY = 10
EDGE_THRESHOLD = 5

# image-trimming-toolと共通のカラーパレット
PREVIEW_BACKGROUND_COLOR = QColor(35, 39, 47)
SURFACE_COLOR = "#20242c"
CARD_COLOR = "#292e38"
INPUT_COLOR = "#1f232b"
TEXT_COLOR = "#f4f6f8"
MUTED_TEXT_COLOR = "#aeb6c2"
ACCENT_COLOR = QColor(66, 184, 255)
ACCENT_HEX = "#42b8ff"

# クリップボード／スナップショット画像の保存先（空欄なら Pictures/Batch-Cropper）
IMPORT_SAVE_DIR = r""
LOG_FILENAME = f"trim_log_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"
LOG_ENABLED = False

APP_STYLE = f"""
QMainWindow, QWidget#rootWidget {{
    background: {SURFACE_COLOR}; color: {TEXT_COLOR};
    font-family: "Yu Gothic UI", "Meiryo UI", sans-serif; font-size: 14px;
}}
QFrame#sidePanel {{ background: {SURFACE_COLOR}; border-left: 1px solid #353b47; }}
QLabel {{ color: {TEXT_COLOR}; }}
QLabel#appTitle {{ color: {TEXT_COLOR}; font-size: 22px; font-weight: 700; }}
QLabel#appSubtitle {{ color: {MUTED_TEXT_COLOR}; font-size: 11px; }}
QLabel#sectionTitle {{ color: {TEXT_COLOR}; font-size: 12px; font-weight: 700; }}
QLineEdit {{
    min-height: 36px; padding: 0 10px; border: 1px solid #434b59;
    border-radius: 8px; background: {INPUT_COLOR}; color: {TEXT_COLOR}; font-size: 14px;
    selection-background-color: {ACCENT_HEX};
}}
QLineEdit:focus {{ border: 1px solid {ACCENT_HEX}; }}
QCheckBox {{ color: {TEXT_COLOR}; spacing: 8px; }}
QCheckBox::indicator {{ width: 18px; height: 18px; }}
QCheckBox::indicator:unchecked {{ background: {INPUT_COLOR}; border: 1px solid #596272; border-radius: 5px; }}
QCheckBox::indicator:checked {{ background: {ACCENT_HEX}; border: 1px solid {ACCENT_HEX}; border-radius: 5px; }}
QPushButton {{
    min-height: 40px; padding: 0 14px; border: 1px solid #454e5d;
    border-radius: 9px; background: #353c48; color: {TEXT_COLOR}; font-weight: 600;
}}
QPushButton:hover {{ background: #414a58; border-color: #596576; }}
QPushButton:pressed {{ background: #2c323c; }}
QPushButton#primaryButton {{ background: {ACCENT_HEX}; border-color: {ACCENT_HEX}; color: #07131b; }}
QPushButton#primaryButton:hover {{ background: #69c7ff; }}
QScrollArea#thumbnailArea {{ background: #1b1f26; border: 1px solid #303641; border-radius: 10px; }}
QScrollBar:horizontal {{ height: 10px; background: #1b1f26; margin: 1px 4px; }}
QScrollBar::handle:horizontal {{ min-width: 32px; background: #454e5d; border-radius: 4px; }}
QScrollBar::handle:horizontal:hover {{ background: #596576; }}
QScrollBar::add-line:horizontal, QScrollBar::sub-line:horizontal {{ width: 0; }}
QSplitter::handle {{ background: #353b47; width: 1px; }}
"""


def _resolve_log_path() -> Path:
    try:
        return Path(__file__).with_name(LOG_FILENAME)
    except Exception:
        return Path(LOG_FILENAME)


LOG_PATH = _resolve_log_path()


def _log_debug(message: str) -> None:
    if not LOG_ENABLED:
        return
    try:
        timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S.%f")[:-3]
        with open(LOG_PATH, "a", encoding="utf-8") as log_file:
            log_file.write(f"{timestamp} {message}\n")
    except Exception:
        pass


def resolve_import_dir() -> Path:
    """取り込み画像の保存先を作成して返す。"""
    base = str(IMPORT_SAVE_DIR or "").strip()
    try:
        target = Path(base).expanduser() if base else Path.home() / "Pictures" / "Batch-Cropper"
        target.mkdir(parents=True, exist_ok=True)
    except Exception:
        target = Path.home()
    return target


def build_unique_path(base_dir: Path, prefix: str, ext: str = ".png") -> Path:
    """同名ファイルを上書きしない保存パスを返す。"""
    suffix = f".{ext.lstrip('.')}"
    timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    file_path = base_dir / f"{prefix}_{timestamp}{suffix}"
    counter = 1
    while file_path.exists():
        file_path = base_dir / f"{prefix}_{timestamp}_{counter}{suffix}"
        counter += 1
    return file_path


def add_bc_suffix(path: str) -> str:
    """出力ファイル名の末尾を _bc にそろえる。"""
    base, ext = os.path.splitext(path)
    if base.endswith("_bc"):
        base = base[:-3]
    return base + "_bc" + ext


def pil_to_qimage(image: Image.Image) -> QImage:
    """Pillow画像をデータ所有権が独立したQImageへ変換する。"""
    normalized = ImageOps.exif_transpose(image)
    if normalized.mode != "RGBA":
        normalized = normalized.convert("RGBA")
    width, height = normalized.size
    raw = normalized.tobytes("raw", "RGBA")
    return QImage(raw, width, height, width * 4, QImage.Format.Format_RGBA8888).copy()


@dataclass
class CropRect:
    """Qtの描画APIから独立したクロップ矩形。"""

    x: int = 0
    y: int = 0
    width: int = 0
    height: int = 0

    @property
    def right(self) -> int:
        return self.x + self.width

    @property
    def bottom(self) -> int:
        return self.y + self.height

    def copy(self) -> "CropRect":
        return CropRect(self.x, self.y, self.width, self.height)


class ThumbnailLabel(QLabel):
    """クリック可能なサムネイル。"""

    clicked = Signal(int)

    def __init__(self, index: int, parent: QWidget | None = None):
        super().__init__(parent)
        self.index = index
        self.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.setFixedSize(THUMBNAIL_SIZE[0] + 12, THUMBNAIL_SIZE[1] + 12)
        self.setCursor(Qt.CursorShape.PointingHandCursor)

    def mouseReleaseEvent(self, event: QMouseEvent) -> None:
        if event.button() == Qt.MouseButton.LeftButton:
            self.clicked.emit(self.index)
        super().mouseReleaseEvent(event)


class ThumbnailPanel(QScrollArea):
    """横スクロール可能なサムネイル一覧。"""

    def __init__(self, select_callback, parent: QWidget | None = None):
        super().__init__(parent)
        self.setObjectName("thumbnailArea")
        self.setWidgetResizable(True)
        self.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        self.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self.setFixedHeight(THUMBNAIL_SIZE[1] + 34)
        self._select_callback = select_callback
        self._pixmap_cache: dict[int, tuple[int, QPixmap]] = {}
        self._thumb_widgets: list[ThumbnailLabel] = []
        self._image_ids: list[int] = []
        self._selected_index: int | None = None
        self._content = QWidget()
        self._content.setStyleSheet("background: #1b1f26;")
        self._layout = QHBoxLayout(self._content)
        self._layout.setContentsMargins(8, 7, 8, 7)
        self._layout.setSpacing(7)
        self._layout.addStretch(1)
        self.setWidget(self._content)

    def _clear_layout(self) -> None:
        while self._layout.count():
            item = self._layout.takeAt(0)
            widget = item.widget()
            if widget is not None:
                widget.deleteLater()

    def _style_thumbnail(self, widget: ThumbnailLabel, selected: bool) -> None:
        border = THUMBNAIL_SELECTED_BORDER_WIDTH if selected else 1
        color = THUMBNAIL_SELECTED_BORDER_COLOR.name() if selected else "#454e5d"
        background = CARD_COLOR if selected else "#353c48"
        widget.setStyleSheet(
            f"QLabel {{ background: {background}; border: {border}px solid {color}; border-radius: 9px; }}"
        )

    def update_thumbnails(self, images: list[Image.Image], selected_index: int | None = None) -> None:
        if selected_index is not None and not 0 <= selected_index < len(images):
            selected_index = None
        image_ids = [id(image) for image in images]
        rebuild = image_ids != self._image_ids or len(images) != len(self._thumb_widgets)
        if rebuild:
            self._clear_layout()
            self._thumb_widgets.clear()
            valid_cache: dict[int, tuple[int, QPixmap]] = {}
            for index, image in enumerate(images):
                image_id = image_ids[index]
                cached = self._pixmap_cache.get(index)
                if cached and cached[0] == image_id:
                    pixmap = cached[1]
                else:
                    qimage = pil_to_qimage(image)
                    pixmap = QPixmap.fromImage(qimage).scaled(
                        QSize(*THUMBNAIL_SIZE),
                        Qt.AspectRatioMode.KeepAspectRatio,
                        Qt.TransformationMode.SmoothTransformation,
                    )
                valid_cache[index] = (image_id, pixmap)
                label = ThumbnailLabel(index, self._content)
                label.setPixmap(pixmap)
                label.clicked.connect(self._select_callback)
                self._style_thumbnail(label, index == selected_index)
                self._layout.addWidget(label)
                self._thumb_widgets.append(label)
            self._layout.addStretch(1)
            self._pixmap_cache = valid_cache
            self._image_ids = image_ids
        elif selected_index != self._selected_index:
            for index, widget in enumerate(self._thumb_widgets):
                self._style_thumbnail(widget, index == selected_index)
        self._selected_index = selected_index

    def clear_cache(self) -> None:
        self._pixmap_cache.clear()
        self._image_ids.clear()


class PreviewPanel(QWidget):
    """画像、選択範囲、リサイズハンドルを描画するプレビュー。"""

    crop_changed = Signal(tuple)

    def __init__(self, parent: QWidget | None = None):
        super().__init__(parent)
        self.setMinimumSize(420, 320)
        self.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
        self.setMouseTracking(True)
        self.setAutoFillBackground(False)
        self.current_image: Image.Image | None = None
        self._cached_pixmap: QPixmap | None = None
        self._cached_size = (0, 0)
        self._cached_image_id: int | None = None
        self.display_w = self.display_h = 0
        self.offset_x = self.offset_y = 0
        self.old_display_w = self.old_display_h = 0
        self.crop_rect: tuple[int, int, int, int] | None = None
        self.mode = "idle"
        self.drag_handle: str | None = None
        self.drag_start = QPoint()
        self.original_rect: CropRect | None = None
        self.fixed_aspect = True
        self.crop_aspect = "1:1"

    def _event_to_display_point(self, event: QMouseEvent) -> QPoint:
        position = event.position().toPoint()
        return QPoint(position.x() - self.offset_x, position.y() - self.offset_y)

    def _clamp_display_point(self, point: QPoint) -> QPoint:
        return QPoint(max(0, min(point.x(), self.display_w)), max(0, min(point.y(), self.display_h)))

    def _point_in_display(self, point: QPoint) -> bool:
        return 0 <= point.x() <= self.display_w and 0 <= point.y() <= self.display_h

    @staticmethod
    def _rect_contains_point(rect_tuple: tuple[int, int, int, int], point: QPoint) -> bool:
        x, y, width, height = rect_tuple
        return x <= point.x() <= x + width and y <= point.y() <= y + height

    def _iter_handle_rects_display(self):
        if not self.crop_rect:
            return
        x, y, width, height = self.crop_rect
        half = HANDLE_SIZE // 2
        points = {
            "top_left": (x, y),
            "top": (x + width / 2, y),
            "top_right": (x + width, y),
            "right": (x + width, y + height / 2),
            "bottom_right": (x + width, y + height),
            "bottom": (x + width / 2, y + height),
            "bottom_left": (x, y + height),
            "left": (x, y + height / 2),
        }
        for name, (point_x, point_y) in points.items():
            yield name, QRect(round(point_x) - half, round(point_y) - half, HANDLE_SIZE, HANDLE_SIZE)

    def _hit_test_handle(self, point: QPoint) -> str | None:
        for name, rect in self._iter_handle_rects_display() or ():
            if rect.contains(point):
                return name
        return None

    def _get_aspect_ratio(self) -> float | None:
        if not self.fixed_aspect:
            return None
        try:
            width_ratio, height_ratio = map(float, self.crop_aspect.split(":"))
            if width_ratio <= 0 or height_ratio <= 0:
                return None
            return width_ratio / height_ratio
        except (TypeError, ValueError):
            return None

    def _ensure_within_display(self, rect: CropRect | None) -> CropRect:
        if rect is None or self.display_w <= 0 or self.display_h <= 0:
            return CropRect()
        rect = rect.copy()
        rect.width = max(MIN_CROP_SIZE, min(rect.width, self.display_w))
        rect.height = max(MIN_CROP_SIZE, min(rect.height, self.display_h))
        rect.x = max(0, min(rect.x, self.display_w - rect.width))
        rect.y = max(0, min(rect.y, self.display_h - rect.height))
        return rect

    def _ensure_min_size(self, rect: CropRect) -> CropRect:
        rect.width = max(MIN_CROP_SIZE, rect.width)
        rect.height = max(MIN_CROP_SIZE, rect.height)
        return self._ensure_within_display(rect)

    def _create_rect_with_ratio(self, anchor: QPoint, current: QPoint, ratio: float) -> CropRect:
        delta_x = current.x() - anchor.x()
        delta_y = current.y() - anchor.y()
        width = abs(delta_x)
        height = abs(delta_y)
        if width == 0 and height == 0:
            return CropRect(anchor.x(), anchor.y(), 0, 0)
        if height == 0:
            height = round(width / ratio)
        if width == 0:
            width = round(height * ratio)
        if width / height > ratio:
            width = round(height * ratio)
        else:
            height = round(width / ratio)
        end_x = anchor.x() + (width if delta_x >= 0 else -width)
        end_y = anchor.y() + (height if delta_y >= 0 else -height)
        return self._ensure_within_display(
            CropRect(min(anchor.x(), end_x), min(anchor.y(), end_y), abs(end_x - anchor.x()), abs(end_y - anchor.y()))
        )

    def _create_rect(self, anchor: QPoint, current: QPoint) -> CropRect:
        anchor = self._clamp_display_point(anchor)
        current = self._clamp_display_point(current)
        ratio = self._get_aspect_ratio()
        if ratio:
            rect = self._create_rect_with_ratio(anchor, current, ratio)
        else:
            rect = CropRect(
                min(anchor.x(), current.x()),
                min(anchor.y(), current.y()),
                abs(current.x() - anchor.x()),
                abs(current.y() - anchor.y()),
            )
        return self._ensure_min_size(rect)

    def _rect_from_crop(self) -> CropRect | None:
        return CropRect(*map(round, self.crop_rect)) if self.crop_rect else None

    def _update_selection_move(self, delta_x: int, delta_y: int) -> None:
        if self.original_rect:
            rect = self.original_rect.copy()
            rect.x += delta_x
            rect.y += delta_y
            rect = self._ensure_min_size(rect)
            self.crop_rect = (rect.x, rect.y, rect.width, rect.height)

    def _resize_free(self, point: QPoint, handle: str) -> CropRect:
        rect = self.original_rect.copy()
        left, top, right, bottom = rect.x, rect.y, rect.right, rect.bottom
        if "left" in handle:
            left = min(point.x(), right - MIN_CROP_SIZE)
        if "right" in handle:
            right = max(point.x(), left + MIN_CROP_SIZE)
        if "top" in handle:
            top = min(point.y(), bottom - MIN_CROP_SIZE)
        if "bottom" in handle:
            bottom = max(point.y(), top + MIN_CROP_SIZE)
        left, right = max(0, left), min(self.display_w, right)
        top, bottom = max(0, top), min(self.display_h, bottom)
        return self._ensure_min_size(CropRect(left, top, right - left, bottom - top))

    def _resize_corner_with_ratio(self, point: QPoint, handle: str, origin: CropRect, ratio: float) -> CropRect:
        anchors = {
            "top_left": (origin.right, origin.bottom, -1, -1),
            "top_right": (origin.x, origin.bottom, 1, -1),
            "bottom_left": (origin.right, origin.y, -1, 1),
            "bottom_right": (origin.x, origin.y, 1, 1),
        }
        anchor_x, anchor_y, horizontal, vertical = anchors[handle]
        width = max(MIN_CROP_SIZE, abs(point.x() - anchor_x))
        height = max(MIN_CROP_SIZE, abs(point.y() - anchor_y))
        if width / height > ratio:
            width = round(height * ratio)
        else:
            height = round(width / ratio)
        left = anchor_x - width if horizontal < 0 else anchor_x
        top = anchor_y - height if vertical < 0 else anchor_y
        return self._ensure_min_size(CropRect(left, top, width, height))

    def _resize_edge_with_ratio(self, point: QPoint, handle: str, origin: CropRect, ratio: float) -> CropRect:
        if handle in ("left", "right"):
            anchor_x = origin.right if handle == "left" else origin.x
            width = abs(anchor_x - point.x())
            height = round(width / ratio)
            left = anchor_x - width if handle == "left" else anchor_x
            top = round(origin.y + origin.height / 2 - height / 2)
        else:
            anchor_y = origin.bottom if handle == "top" else origin.y
            height = abs(anchor_y - point.y())
            width = round(height * ratio)
            left = round(origin.x + origin.width / 2 - width / 2)
            top = anchor_y - height if handle == "top" else anchor_y
        return self._ensure_min_size(CropRect(left, top, width, height))

    def _update_selection_resize(self, point: QPoint) -> None:
        if not self.original_rect or not self.drag_handle:
            return
        ratio = self._get_aspect_ratio()
        if not ratio:
            rect = self._resize_free(point, self.drag_handle)
        elif self.drag_handle in ("top_left", "top_right", "bottom_left", "bottom_right"):
            rect = self._resize_corner_with_ratio(point, self.drag_handle, self.original_rect, ratio)
        else:
            rect = self._resize_edge_with_ratio(point, self.drag_handle, self.original_rect, ratio)
        self.crop_rect = (rect.x, rect.y, rect.width, rect.height)

    def _update_cursor(self, point: QPoint) -> None:
        handle = self._hit_test_handle(point)
        cursors = {
            "top_left": Qt.CursorShape.SizeFDiagCursor,
            "bottom_right": Qt.CursorShape.SizeFDiagCursor,
            "top_right": Qt.CursorShape.SizeBDiagCursor,
            "bottom_left": Qt.CursorShape.SizeBDiagCursor,
            "left": Qt.CursorShape.SizeHorCursor,
            "right": Qt.CursorShape.SizeHorCursor,
            "top": Qt.CursorShape.SizeVerCursor,
            "bottom": Qt.CursorShape.SizeVerCursor,
        }
        if handle:
            shape = cursors[handle]
        elif self._point_in_display(point):
            shape = Qt.CursorShape.SizeAllCursor if self.crop_rect and self._rect_contains_point(self.crop_rect, point) else Qt.CursorShape.CrossCursor
        else:
            shape = Qt.CursorShape.ArrowCursor
        self.setCursor(QCursor(shape))

    def apply_aspect_ratio_to_selection(self) -> None:
        ratio = self._get_aspect_ratio()
        if not ratio or not self.crop_rect:
            return
        x, y, width, height = self.crop_rect
        center_x, center_y = x + width / 2, y + height / 2
        height = min(self.display_h, max(MIN_CROP_SIZE, round(width / ratio)))
        width = min(self.display_w, max(MIN_CROP_SIZE, round(height * ratio)))
        if width >= self.display_w:
            height = max(MIN_CROP_SIZE, round(width / ratio))
        rect = self._ensure_min_size(CropRect(round(center_x - width / 2), round(center_y - height / 2), width, height))
        self.crop_rect = (rect.x, rect.y, rect.width, rect.height)

    def set_image(self, image: Image.Image) -> None:
        previous_rect = self.crop_rect
        self.current_image = image.copy()
        self._cached_pixmap = None
        self.update_display_geometry()
        self.crop_rect = self.clip_rect(*previous_rect) if previous_rect else None
        if self.crop_rect is None:
            self.init_crop_rect()
        self.mode = "idle"
        self.update()

    def update_display_geometry(self) -> None:
        if not self.current_image:
            return
        panel_width, panel_height = self.width(), self.height()
        image_width, image_height = self.current_image.size
        scale = min(panel_width / image_width, panel_height / image_height)
        self.display_w = max(1, int(image_width * scale))
        self.display_h = max(1, int(image_height * scale))
        self.offset_x = (panel_width - self.display_w) // 2
        self.offset_y = (panel_height - self.display_h) // 2
        self._cached_pixmap = None

    def resizeEvent(self, event: QResizeEvent) -> None:
        self.old_display_w, self.old_display_h = self.display_w, self.display_h
        self.update_display_geometry()
        if self.crop_rect and self.old_display_w and self.old_display_h:
            self.rescale_crop_rect()
        super().resizeEvent(event)

    def _ensure_cached_pixmap(self) -> None:
        if not self.current_image or self.display_w <= 0 or self.display_h <= 0:
            self._cached_pixmap = None
            return
        cache_key = (self.display_w, self.display_h)
        if self._cached_pixmap is None or self._cached_size != cache_key or self._cached_image_id != id(self.current_image):
            qimage = pil_to_qimage(self.current_image)
            self._cached_pixmap = QPixmap.fromImage(qimage).scaled(
                self.display_w,
                self.display_h,
                Qt.AspectRatioMode.IgnoreAspectRatio,
                Qt.TransformationMode.SmoothTransformation,
            )
            self._cached_size = cache_key
            self._cached_image_id = id(self.current_image)

    def paintEvent(self, event) -> None:
        painter = QPainter(self)
        painter.fillRect(self.rect(), PREVIEW_BACKGROUND_COLOR)
        if not self.current_image:
            painter.setPen(QColor(MUTED_TEXT_COLOR))
            painter.setFont(QFont("Segoe UI", 13))
            painter.drawText(self.rect(), Qt.AlignmentFlag.AlignCenter, "画像をここへドロップ\nまたは Ctrl+V で貼り付け")
            return
        self._ensure_cached_pixmap()
        if not self._cached_pixmap:
            return
        painter.drawPixmap(self.offset_x, self.offset_y, self._cached_pixmap)
        if not self.crop_rect:
            return
        crop_x, crop_y, crop_width, crop_height = self.crop_rect
        crop_qrect = QRect(crop_x + self.offset_x, crop_y + self.offset_y, crop_width, crop_height)
        overlay = QPainterPath()
        overlay.addRect(self.rect())
        cutout = QPainterPath()
        cutout.addRect(crop_qrect)
        overlay = overlay.subtracted(cutout)
        painter.fillPath(overlay, QColor(0, 0, 0, OVERLAY_ALPHA))
        painter.setPen(QPen(ACCENT_COLOR, 2))
        painter.drawRect(crop_qrect)
        painter.setBrush(QColor(248, 250, 252))
        painter.setPen(QPen(ACCENT_COLOR, 1))
        for _, handle_rect in self._iter_handle_rects_display() or ():
            handle_rect.translate(self.offset_x, self.offset_y)
            painter.drawRect(handle_rect)

    def render_to_pixmap(self) -> QPixmap:
        pixmap = QPixmap(self.size())
        self.render(pixmap)
        return pixmap

    def copy_original_to_clipboard(self) -> bool:
        if not self.current_image:
            return False
        buffer = io.BytesIO()
        ImageOps.exif_transpose(self.current_image).save(buffer, format="PNG")
        mime_data = QMimeData()
        mime_data.setData("image/png", QByteArray(buffer.getvalue()))
        QApplication.clipboard().setMimeData(mime_data)
        return True

    def init_crop_rect(self) -> None:
        if self.display_w <= 0 or self.display_h <= 0:
            self.crop_rect = None
            return
        ratio = self._get_aspect_ratio()
        if ratio:
            width = self.display_w // 4
            height = round(width / ratio)
            if height > self.display_h // 2:
                height = self.display_h // 4
                width = round(height * ratio)
        else:
            width = self.display_w // 4
            height = width
        rect = self._ensure_min_size(
            CropRect((self.display_w - width) // 2, (self.display_h - height) // 2, width, height)
        )
        self.crop_rect = (rect.x, rect.y, rect.width, rect.height)
        self._emit_crop_changed()

    def clip_rect(self, x: int, y: int, width: int, height: int) -> tuple[int, int, int, int]:
        rect = self._ensure_min_size(CropRect(round(x), round(y), round(width), round(height)))
        return rect.x, rect.y, rect.width, rect.height

    def mousePressEvent(self, event: QMouseEvent) -> None:
        if event.button() != Qt.MouseButton.LeftButton or not self.current_image:
            return
        point = self._event_to_display_point(event)
        handle = self._hit_test_handle(point)
        if handle and self.crop_rect:
            self.mode, self.drag_handle = "resizing", handle
            self.original_rect = self._rect_from_crop()
        elif self.crop_rect and self._rect_contains_point(self.crop_rect, point):
            self.mode, self.drag_handle = "moving", "inside"
            self.original_rect = self._rect_from_crop()
        elif self._point_in_display(point):
            self.mode, self.drag_handle = "creating", None
            self.original_rect = None
            point = self._clamp_display_point(point)
            self.crop_rect = (point.x(), point.y(), 0, 0)
        else:
            return
        self.drag_start = self._clamp_display_point(point)
        self.grabMouse()
        self.update()

    def mouseMoveEvent(self, event: QMouseEvent) -> None:
        if not self.current_image:
            return
        point = self._event_to_display_point(event)
        if event.buttons() & Qt.MouseButton.LeftButton and self.mode != "idle":
            point = self._clamp_display_point(point)
            if self.mode == "creating":
                rect = self._create_rect(self.drag_start, point)
                self.crop_rect = (rect.x, rect.y, rect.width, rect.height)
            elif self.mode == "moving":
                self._update_selection_move(point.x() - self.drag_start.x(), point.y() - self.drag_start.y())
            elif self.mode == "resizing":
                self._update_selection_resize(point)
            self._emit_crop_changed()
            self.update()
        else:
            self._update_cursor(point)

    def mouseReleaseEvent(self, event: QMouseEvent) -> None:
        if event.button() == Qt.MouseButton.LeftButton:
            if self.mouseGrabber() is self:
                self.releaseMouse()
            self.mode = "idle"
            self.drag_handle = None
            self.original_rect = None
            self.drag_start = QPoint()
            if self.crop_rect:
                self.crop_rect = self.clip_rect(*self.crop_rect)
            self._emit_crop_changed()
            self.update()

    def leaveEvent(self, event) -> None:
        if self.mode == "idle":
            self.unsetCursor()
        super().leaveEvent(event)

    def _emit_crop_changed(self) -> None:
        box = self.get_crop_box()
        if box:
            self.crop_changed.emit(box)

    def rescale_crop_rect(self) -> None:
        if not self.crop_rect or not self.old_display_w or not self.old_display_h:
            return
        x, y, width, height = self.crop_rect
        scale_x = self.display_w / self.old_display_w
        scale_y = self.display_h / self.old_display_h
        rect = self._ensure_min_size(
            CropRect(round(x * scale_x), round(y * scale_y), round(width * scale_x), round(height * scale_y))
        )
        self.crop_rect = (rect.x, rect.y, rect.width, rect.height)
        self._emit_crop_changed()

    def get_crop_box(self) -> tuple[int, int, int, int] | None:
        if not self.crop_rect or not self.current_image or self.display_w <= 0 or self.display_h <= 0:
            return None
        x, y, width, height = self.crop_rect
        image_width, image_height = self.current_image.size
        scale_x, scale_y = image_width / self.display_w, image_height / self.display_h
        x1 = max(0, min(image_width - 1, math.floor(x * scale_x)))
        y1 = max(0, min(image_height - 1, math.floor(y * scale_y)))
        x2 = max(x1 + 1, min(image_width, math.ceil((x + width) * scale_x)))
        y2 = max(y1 + 1, min(image_height, math.ceil((y + height) * scale_y)))
        return x1, y1, x2, y2


class ControlPanel(QFrame):
    """座標入力と一括編集操作をまとめた右パネル。"""

    def __init__(self, main: "MainWindow", parent: QWidget | None = None):
        super().__init__(parent)
        self.setObjectName("sidePanel")
        self.setMinimumWidth(RIGHT_PANEL_WIDTH)
        self.setMaximumWidth(RIGHT_PANEL_WIDTH + 40)
        self.main = main
        self.textcs: dict[str, QLineEdit] = {}
        self._is_editing = False

        layout = QVBoxLayout(self)
        layout.setContentsMargins(22, 24, 22, 22)
        layout.setSpacing(12)

        title = QLabel("Batch Cropper")
        title.setObjectName("appTitle")
        subtitle = QLabel("複数の画像を同じ範囲で一括編集")
        subtitle.setObjectName("appSubtitle")
        layout.addWidget(title)
        layout.addWidget(subtitle)
        layout.addSpacing(10)

        section = QLabel("切り抜き座標")
        section.setObjectName("sectionTitle")
        layout.addWidget(section)
        coordinate_grid = QGridLayout()
        coordinate_grid.setHorizontalSpacing(8)
        coordinate_grid.setVerticalSpacing(8)
        for index, (label_text, key) in enumerate((("XS", "xs"), ("YS", "ys"), ("XE", "xe"), ("YE", "ye"))):
            label = QLabel(label_text)
            label.setStyleSheet(f"color:{MUTED_TEXT_COLOR}; font-weight:700;")
            edit = QLineEdit()
            edit.setAlignment(Qt.AlignmentFlag.AlignCenter)
            edit.setPlaceholderText("—")
            edit.returnPressed.connect(lambda current_key=key: self.on_coord_enter(current_key))
            self.textcs[key] = edit
            row, column = divmod(index, 2)
            cell = QVBoxLayout()
            cell.setSpacing(3)
            cell.addWidget(label)
            cell.addWidget(edit)
            coordinate_grid.addLayout(cell, row, column)
        layout.addLayout(coordinate_grid)

        aspect_row = QHBoxLayout()
        self.cb_aspect = QCheckBox("縦横比を固定")
        self.cb_aspect.setChecked(True)
        self.cb_aspect.toggled.connect(self.on_aspect_toggle)
        self.tc_aspect = QLineEdit("1:1")
        self.tc_aspect.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.tc_aspect.setFixedWidth(72)
        self.tc_aspect.returnPressed.connect(self.on_aspect_toggle)
        aspect_row.addWidget(self.cb_aspect)
        aspect_row.addStretch(1)
        aspect_row.addWidget(self.tc_aspect)
        layout.addLayout(aspect_row)
        layout.addSpacing(8)

        trim_button = self._button("トリミングして保存", main.on_trim_all, "primaryButton")
        reduce_button = self._button("PNGを減色して保存", main.on_png_reduce)
        undo_button = self._button("ひとつ前に戻す", main.on_revert_all)
        snapshot_button = self._button("スナップショットを追加", main.on_snapshot)
        clear_selected_button = self._button("選択画像をリストから外す", main.on_clear_selected)
        clear_button = self._button("すべてクリア", main.on_clear_all)
        for button in (trim_button, reduce_button, undo_button, snapshot_button):
            layout.addWidget(button)
        layout.addStretch(1)
        layout.addWidget(clear_selected_button)
        layout.addWidget(clear_button)

    @staticmethod
    def _button(text: str, callback, object_name: str = "") -> QPushButton:
        button = QPushButton(text)
        if object_name:
            button.setObjectName(object_name)
        button.clicked.connect(callback)
        return button

    def set_crop_box(self, box: tuple[int, int, int, int]) -> None:
        if self._is_editing:
            return
        for key, value in zip(("xs", "ys", "xe", "ye"), box):
            self.textcs[key].setText(str(value))

    def clear_coordinates(self) -> None:
        for edit in self.textcs.values():
            edit.clear()

    def on_aspect_toggle(self, *_args) -> None:
        self.main.preview.fixed_aspect = self.cb_aspect.isChecked()
        self.main.preview.crop_aspect = self.tc_aspect.text().strip()
        if self.main.preview.fixed_aspect:
            if not self.main.preview._get_aspect_ratio():
                QMessageBox.warning(self, "入力エラー", "縦横比は 1:1 の形式で入力してください。")
                self.cb_aspect.setChecked(False)
                return
            if self.main.preview.crop_rect:
                self.main.preview.apply_aspect_ratio_to_selection()
            else:
                self.main.preview.init_crop_rect()
        self.main.preview._emit_crop_changed()
        self.main.preview.update()

    def on_coord_enter(self, key: str) -> None:
        self._is_editing = True
        try:
            box = self.get_validated_box()
            if box is None or not self.main.preview.current_image:
                return
            old_box = self.main.preview.get_crop_box() or box
            xs, ys, xe, ye = box
            if self.cb_aspect.isChecked():
                ratio = self.main.preview._get_aspect_ratio()
                if not ratio:
                    return
                old_xs, old_ys, old_xe, old_ye = old_box
                if key == "xs":
                    xe = old_xe + (xs - old_xs)
                elif key == "ys":
                    ye = old_ye + (ys - old_ys)
                elif key == "xe":
                    ye = ys + (xe - xs) / ratio
                elif key == "ye":
                    xe = xs + (ye - ys) * ratio
            preview = self.main.preview
            image_width, image_height = preview.current_image.size
            rect = CropRect(
                round(xs * preview.display_w / image_width),
                round(ys * preview.display_h / image_height),
                round((xe - xs) * preview.display_w / image_width),
                round((ye - ys) * preview.display_h / image_height),
            )
            rect = preview._ensure_min_size(rect)
            preview.crop_rect = (rect.x, rect.y, rect.width, rect.height)
            preview.update()
        finally:
            self._is_editing = False
            box = self.main.preview.get_crop_box()
            if box:
                self.set_crop_box(box)

    def get_validated_box(self) -> tuple[int, int, int, int] | None:
        try:
            box = tuple(int(self.textcs[key].text()) for key in ("xs", "ys", "xe", "ye"))
        except ValueError:
            QMessageBox.warning(self, "入力エラー", "座標は整数で入力してください。")
            return None
        xs, ys, xe, ye = box
        if xe <= xs or ye <= ys:
            QMessageBox.warning(self, "入力エラー", "XE > XS かつ YE > YS にしてください。")
            return None
        return box


class MainWindow(QMainWindow):
    """アプリケーションのメインウィンドウ。"""

    def __init__(self):
        super().__init__()
        self.setWindowTitle("Batch-Cropper")
        self.resize(*APP_WINDOW_SIZE)
        self.setMinimumSize(*APP_WINDOW_SIZE)
        self.setAcceptDrops(True)
        self._default_aspect = APP_WINDOW_SIZE[0] / APP_WINDOW_SIZE[1]
        self.file_paths: list[str] = []
        self.images: list[Image.Image] = []
        self.history: list[dict] = []
        self.reduced_flags: list[bool] = []
        self.selected_index = -1

        root = QWidget()
        root.setObjectName("rootWidget")
        root_layout = QVBoxLayout(root)
        root_layout.setContentsMargins(12, 12, 12, 12)
        root_layout.setSpacing(0)
        splitter = QSplitter(Qt.Orientation.Horizontal)
        splitter.setChildrenCollapsible(False)
        splitter.setHandleWidth(1)

        left = QWidget()
        left_layout = QVBoxLayout(left)
        left_layout.setContentsMargins(0, 0, 12, 0)
        left_layout.setSpacing(10)
        self.preview = PreviewPanel()
        self.thumbnails = ThumbnailPanel(self.on_select_thumbnail)
        left_layout.addWidget(self.preview, 1)
        left_layout.addWidget(self.thumbnails)
        self.ctrl = ControlPanel(self)

        splitter.addWidget(left)
        splitter.addWidget(self.ctrl)
        splitter.setStretchFactor(0, 1)
        splitter.setStretchFactor(1, 0)
        splitter.setSizes([APP_WINDOW_SIZE[0] - RIGHT_PANEL_WIDTH, RIGHT_PANEL_WIDTH])
        root_layout.addWidget(splitter)
        self.setCentralWidget(root)

        self.preview.crop_changed.connect(self.ctrl.set_crop_box)
        self._create_shortcuts()
        _log_debug(f"APP start log={LOG_PATH}")

    def _create_shortcuts(self) -> None:
        copy_action = QAction(self)
        copy_action.setShortcut(QKeySequence.StandardKey.Copy)
        copy_action.triggered.connect(self.on_copy_preview_original)
        self.addAction(copy_action)
        paste_action = QAction(self)
        paste_action.setShortcut(QKeySequence.StandardKey.Paste)
        paste_action.triggered.connect(self.on_paste_from_clipboard)
        self.addAction(paste_action)

    def dragEnterEvent(self, event: QDragEnterEvent) -> None:
        if event.mimeData().hasUrls():
            event.acceptProposedAction()

    def dropEvent(self, event: QDropEvent) -> None:
        paths = [url.toLocalFile() for url in event.mimeData().urls() if url.isLocalFile()]
        self.add_files(paths)
        event.acceptProposedAction()

    def wheelEvent(self, event: QWheelEvent) -> None:
        if event.angleDelta().y() == 0:
            return
        direction = 1 if event.angleDelta().y() > 0 else -1
        screen = self.screen()
        available = screen.availableGeometry()
        min_width, min_height = APP_WINDOW_SIZE
        max_width = min(available.width(), round(available.height() * self._default_aspect))
        target_width = round(self.width() * (1 + WINDOW_RESIZE_SCALE_STEP * direction))
        target_width = max(min_width, min(target_width, max_width))
        target_height = round(target_width / self._default_aspect)
        self.resize(target_width, target_height)
        self.move(available.center() - self.rect().center())
        event.accept()

    def _show_error(self, title: str, message: str) -> None:
        QMessageBox.critical(self, title, message)

    def _show_info(self, title: str, message: str) -> None:
        QMessageBox.information(self, title, message)

    def _get_clipboard_image(self) -> Image.Image | None:
        mime_data = QApplication.clipboard().mimeData()
        if mime_data.hasImage():
            qimage = QApplication.clipboard().image().convertToFormat(QImage.Format.Format_RGBA8888)
            raw = qimage.bits().tobytes()
            return Image.frombytes("RGBA", (qimage.width(), qimage.height()), raw)
        try:
            grabbed = ImageGrab.grabclipboard()
        except Exception:
            grabbed = None
        if isinstance(grabbed, Image.Image):
            return grabbed
        if isinstance(grabbed, list):
            for item in grabbed:
                if os.path.splitext(item)[1].lower() in (".jpg", ".jpeg", ".png", ".bmp", ".tiff", ".tif"):
                    try:
                        with Image.open(item) as image:
                            return image.copy()
                    except Exception:
                        continue
        return None

    def _save_import_image(self, image: Image.Image, prefix: str, ext: str = ".png") -> Path | None:
        file_path = build_unique_path(resolve_import_dir(), prefix, ext)
        try:
            image.save(file_path)
            return file_path
        except Exception as error:
            self._show_error("保存エラー", f"画像を保存できませんでした。\n{error}")
            return None

    def on_paste_from_clipboard(self) -> None:
        image = self._get_clipboard_image()
        if image is None:
            self._show_info("貼り付け", "クリップボードに画像がありません。")
            return
        file_path = self._save_import_image(image, "clipboard")
        if file_path:
            self.file_paths.append(str(file_path))
            self.images.append(image.copy())
            self.reduced_flags.append(False)
            self.history.clear()
            self.push_history()
            self.selected_index = len(self.images) - 1
            self.update_ui()

    def add_files(self, paths: list[str]) -> None:
        errors = []
        for path in paths:
            extension = os.path.splitext(path)[1].lower()
            if extension not in (".jpg", ".jpeg", ".png", ".bmp", ".tiff", ".tif") or path in self.file_paths:
                continue
            try:
                with Image.open(path) as source:
                    image = source.copy()
                    image.info.update(source.info)
                self.file_paths.append(path)
                self.images.append(image)
                self.reduced_flags.append(False)
            except Exception as error:
                errors.append(f"{os.path.basename(path)}: {error}")
        if errors:
            self._show_error("読み込みエラー", "\n".join(errors))
        if self.images:
            self.history.clear()
            self.push_history()
            self.selected_index = 0
            self.update_ui()

    def update_ui(self) -> None:
        if 0 <= self.selected_index < len(self.images):
            self.preview.set_image(self.images[self.selected_index])
            box = self.preview.get_crop_box()
            if box:
                self.ctrl.set_crop_box(box)
        self.thumbnails.update_thumbnails(self.images, self.selected_index)

    def on_copy_preview_original(self) -> None:
        if not 0 <= self.selected_index < len(self.file_paths):
            return
        extension = os.path.splitext(self.file_paths[self.selected_index])[1].lower()
        if extension in (".jpg", ".jpeg"):
            self._show_info("コピーを中止しました", "容量が増えるためJPEG画像はクリップボードへコピーできません。")
            return
        if self.reduced_flags[self.selected_index]:
            self._show_info("コピーを中止しました", "減色後の画像はクリップボード上で容量低減効果がありません。")
            return
        self.preview.copy_original_to_clipboard()

    def on_select_thumbnail(self, index: int) -> None:
        self.selected_index = index
        self.update_ui()

    def on_trim_all(self) -> None:
        crop_box = self.ctrl.get_validated_box()
        if crop_box is None:
            return
        new_paths, new_images, new_flags, errors = [], [], [], []
        for index, (path, image) in enumerate(zip(self.file_paths, self.images)):
            try:
                trimmed = image.crop(crop_box)
                extension = os.path.splitext(path)[1].lower()
                out_path = add_bc_suffix(path)
                save_options = {}
                if extension in (".jpg", ".jpeg"):
                    save_options["quality"] = image.info.get("quality", 80)
                    if trimmed.mode not in ("RGB", "L"):
                        trimmed = trimmed.convert("RGB")
                elif extension in (".tif", ".tiff"):
                    compression = getattr(image, "tag_v2", {}).get(259) if hasattr(image, "tag_v2") else None
                    if compression in (3, 4):
                        trimmed = trimmed.convert("1")
                        save_options["compression"] = "group3" if compression == 3 else "group4"
                trimmed.save(out_path, **save_options)
                with Image.open(out_path) as reopened:
                    new_images.append(reopened.copy())
                new_paths.append(out_path)
                new_flags.append(self.reduced_flags[index])
            except Exception as error:
                errors.append(f"{os.path.basename(path)}: {error}")
        if new_images:
            self.file_paths, self.images, self.reduced_flags = new_paths, new_images, new_flags
            self.push_history()
            self.selected_index = min(max(self.selected_index, 0), len(self.images) - 1)
            self.update_ui()
        if errors:
            self._show_error("トリミングエラー", "\n".join(errors))

    def on_png_reduce(self) -> None:
        if not self.file_paths:
            self._show_info("PNG減色", "減色対象のファイルがありません。")
            return
        if any(os.path.splitext(path)[1].lower() != ".png" for path in self.file_paths):
            self._show_info("PNG減色", "減色はPNGファイルのみ対応しています。")
            return
        new_paths, new_images, errors = [], [], []
        for path, image in zip(self.file_paths, self.images):
            try:
                rgba = image.convert("RGBA")
                alpha = rgba.getchannel("A")
                quantized = rgba.convert("RGB").quantize(colors=256, method=Image.Quantize.MEDIANCUT, dither=Image.Dither.FLOYDSTEINBERG)
                if alpha.getextrema() != (255, 255):
                    quantized = quantized.convert("RGBA")
                    quantized.putalpha(alpha)
                out_path = add_bc_suffix(path)
                quantized.save(out_path, optimize=True)
                with Image.open(out_path) as reopened:
                    new_images.append(reopened.copy())
                new_paths.append(out_path)
            except Exception as error:
                errors.append(f"{os.path.basename(path)}: {error}")
        if new_paths:
            self.file_paths, self.images = new_paths, new_images
            self.reduced_flags = [True] * len(new_images)
            self.push_history()
            self.selected_index = min(max(self.selected_index, 0), len(self.images) - 1)
            self.update_ui()
        if errors:
            self._show_error("PNG減色エラー", "\n".join(errors))

    def on_snapshot(self) -> None:
        try:
            screenshot = ImageGrab.grab(all_screens=True)
        except Exception as error:
            self._show_error("スナップショットエラー", str(error))
            return
        file_path = self._save_import_image(screenshot, "snapshot")
        if file_path:
            self.file_paths.append(str(file_path))
            self.images.append(screenshot.copy())
            self.reduced_flags.append(False)
            self.history.clear()
            self.push_history()
            self.selected_index = len(self.images) - 1
            self.update_ui()

    def on_revert_all(self) -> None:
        if len(self.history) < 2:
            return
        current_paths = list(self.file_paths)
        self.history.pop()
        previous = self.history[-1]
        previous_paths = previous["paths"]
        for path in current_paths:
            base, _ = os.path.splitext(path)
            if path not in previous_paths and base.endswith("_bc"):
                try:
                    os.remove(path)
                except OSError:
                    pass
        for path, image in zip(previous_paths, previous["images"]):
            if os.path.splitext(path)[0].endswith("_bc"):
                try:
                    image.save(path)
                except OSError:
                    pass
        self.file_paths = list(previous_paths)
        self.images = [image.copy() for image in previous["images"]]
        self.reduced_flags = list(previous.get("flags", [False] * len(self.images)))
        self.selected_index = min(max(self.selected_index, 0), len(self.images) - 1)
        self.update_ui()

    def push_history(self) -> None:
        self.history.append(
            {
                "paths": list(self.file_paths),
                "images": [image.copy() for image in self.images],
                "flags": list(self.reduced_flags),
            }
        )
        if len(self.history) > MAX_HISTORY:
            self.history.pop(0)

    def on_clear_selected(self) -> None:
        if not 0 <= self.selected_index < len(self.images):
            self._show_info("選択画像クリア", "リストから外す画像がありません。")
            return
        index = self.selected_index
        del self.file_paths[index]
        del self.images[index]
        del self.reduced_flags[index]
        self.selected_index = min(index, len(self.images) - 1) if self.images else -1
        self.push_history()
        if self.images:
            self.update_ui()
        else:
            self._reset_display()

    def _reset_display(self) -> None:
        self.preview.current_image = None
        self.preview._cached_pixmap = None
        self.preview.crop_rect = None
        self.preview.update()
        self.thumbnails.clear_cache()
        self.thumbnails.update_thumbnails([], None)
        self.ctrl.clear_coordinates()

    def on_clear_all(self) -> None:
        self.file_paths.clear()
        self.images.clear()
        self.history.clear()
        self.reduced_flags.clear()
        self.selected_index = -1
        self._reset_display()


def main() -> int:
    """Qtイベントループを開始する。"""
    Image.MAX_IMAGE_PIXELS = None
    app = QApplication.instance() or QApplication(sys.argv)
    app.setApplicationName("Batch-Cropper")
    app.setStyle("Fusion")
    app.setStyleSheet(APP_STYLE)
    window = MainWindow()
    # ファイルを引数で渡した場合は起動時に読み込む
    if len(sys.argv) > 1:
        window.add_files(sys.argv[1:])
    window.show()
    return app.exec()


if __name__ == "__main__":
    raise SystemExit(main())
