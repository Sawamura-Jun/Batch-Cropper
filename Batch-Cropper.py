import wx
import os
import datetime
import math
import io
import struct
from pathlib import Path
from PIL import Image, ImageGrab
from PIL import Image, ImageOps

# Save directory for clipboard/snapshot imports (empty -> Desktop)
IMPORT_SAVE_DIR = r""

# 定数定義
APP_WINDOW_SIZE = (1120,800)    # デフォルトサイズ
WINDOW_RESIZE_SCALE_STEP = 0.1  # ホイールリサイズの刻み幅（スケール比）
THUMBNAIL_SIZE = (100, 100)
MAX_HISTORY = 10
LOG_FILENAME = f"trim_log_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"
LOG_ENABLED = False
EDGE_THRESHOLD = 5
BACK_GROUND_COLOR = wx.Colour(100, 100, 100)
RIGHT_PANEL_WIDTH = 250
HANDLE_SIZE = 10
MIN_CROP_SIZE = 4
OVERLAY_ALPHA = 100

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
        ts = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S.%f")[:-3]
        with open(LOG_PATH, "a", encoding="utf-8") as f:
            f.write(f"{ts} {message}\n")
    except Exception:
        pass


def resolve_import_dir() -> Path:
    """
    IMPORT_SAVE_DIR が空なら Pictures/Batch-Cropper（Windows想定）、それ以外は指定を使用。
    """
    base = ""
    if IMPORT_SAVE_DIR is not None:
        base = str(IMPORT_SAVE_DIR).strip()

    target = None
    if base:
        try:
            target = Path(base).expanduser()
        except Exception:
            target = None

    if target is None and not base:
        target = Path.home() / "Pictures" / "Batch-Cropper"
    elif target is None:
        target = Path.home()

    try:
        target.mkdir(parents=True, exist_ok=True)
    except Exception:
        target = Path.home()
        target.mkdir(parents=True, exist_ok=True)

    return target


def build_unique_path(base_dir: Path, prefix: str, ext: str = ".png") -> Path:
    """同名ファイルがある場合に連番でずらした保存パスを返す。"""
    ext = f".{ext.lstrip('.')}"
    timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    file_path = base_dir / f"{prefix}_{timestamp}{ext}"
    counter = 1
    while file_path.exists():
        file_path = base_dir / f"{prefix}_{timestamp}_{counter}{ext}"
        counter += 1
    return file_path

def add_bc_suffix(path: str) -> str:
    """_bcのみとし既存の連番は除去する"""
    base, ext = os.path.splitext(path)
    if base.endswith("_bc"):
        base = base[:-3]
    return base + "_bc" + ext

class ThumbnailPanel(wx.Panel):
    """下ペイン：横スクロール可能なサムネイル一覧"""
    def __init__(self, parent, select_callback):
        super().__init__(parent)
        self.select_callback = select_callback
        self._bitmap_cache = {}
        self.scrolled = wx.ScrolledWindow(self, style=wx.HSCROLL)
        self.scrolled.SetScrollRate(5, 5)
        self.scrolled.SetMinSize((-1, THUMBNAIL_SIZE[1] + 10))  # 高さ固定
        self.hbox = wx.BoxSizer(wx.HORIZONTAL)
        self.scrolled.SetSizer(self.hbox)
        sizer = wx.BoxSizer(wx.VERTICAL)
        sizer.Add(self.scrolled, 1, wx.EXPAND)
        self.SetSizer(sizer)

    def update_thumbnails(self, images):
        if not images:
            self._bitmap_cache.clear()
        _log_debug(f"THUMBNAIL update start count={len(images)} cache={len(self._bitmap_cache)}")
        self.hbox.Clear(True)
        for idx, img in enumerate(images):
            img_id = id(img)
            cached = self._bitmap_cache.get(idx)
            if cached and cached[0] == img_id:
                _log_debug(f"THUMBNAIL hit idx={idx} id={img_id}")
                bmp = cached[1]
            else:
                if cached:
                    _log_debug(f"THUMBNAIL miss idx={idx} id={img_id} cached_id={cached[0]}")
                else:
                    _log_debug(f"THUMBNAIL miss idx={idx} id={img_id} cached_id=None")
                thumb = img.convert('RGB').copy()
                thumb.thumbnail(THUMBNAIL_SIZE, Image.LANCZOS)
                w, h = thumb.size
                raw = thumb.convert('RGB').tobytes()
                bmp = wx.Bitmap.FromBuffer(w, h, raw)
                self._bitmap_cache[idx] = (img_id, bmp)
            sb = wx.StaticBitmap(self.scrolled, bitmap=bmp)
            sb.idx = idx
            sb.Bind(wx.EVT_LEFT_UP, self.OnThumbClick)
            self.hbox.Add(sb, 0, wx.ALL, 5)
        self.hbox.Layout()
        self.scrolled.Layout()
        self.scrolled.FitInside()
        _log_debug("THUMBNAIL update end")

    def OnThumbClick(self, evt):
        # スクロール位置を変えず選択のみ
        idx = evt.GetEventObject().idx
        self.select_callback(idx)

class PreviewPanel(wx.Panel):
    """中央プレビュー：画像表示＋ラバーバンド"""
    def __init__(self, parent):
        super().__init__(parent)
        self.SetBackgroundColour(BACK_GROUND_COLOR)
        self.SetBackgroundStyle(wx.BG_STYLE_PAINT)
        self.Bind(wx.EVT_ERASE_BACKGROUND, lambda e: None)
        self.current_image = None
        self._cached_bitmap = None
        self._cached_size = (0, 0)
        self._cached_img_id = None
        self.display_w = self.display_h = 0
        self.offset_x = self.offset_y = 0
        self.old_display_w = self.old_display_h = 0
        self.crop_rect = None
        self.mode = "idle"
        self.drag_handle = None
        self.drag_start = wx.Point()
        self.original_rect = None
        self.fixed_aspect = True
        self.crop_aspect = "1:1"
        self.clip = wx.Clipboard()  #クリップボードの設定
        self._current_cursor = wx.Cursor(wx.CURSOR_ARROW)
        self.Bind(wx.EVT_PAINT, self.OnPaint)
        self.Bind(wx.EVT_SIZE, self.OnSize)
        self.Bind(wx.EVT_LEFT_DOWN, self.OnLeftDown)
        self.Bind(wx.EVT_MOTION, self.OnMouseMove)
        self.Bind(wx.EVT_LEFT_UP, self.OnLeftUp)
        self.Bind(wx.EVT_LEAVE_WINDOW, self.OnMouseLeave)

    def _event_to_display_point(self, event):
        x, y = event.GetPosition()
        return wx.Point(x - self.offset_x, y - self.offset_y)

    def _clamp_display_point(self, point):
        return wx.Point(
            max(0, min(point.x, self.display_w)),
            max(0, min(point.y, self.display_h))
        )

    def _point_in_display(self, point):
        return 0 <= point.x <= self.display_w and 0 <= point.y <= self.display_h

    def _rect_contains_point(self, rect_tuple, point):
        x, y, w, h = rect_tuple
        return x <= point.x <= x + w and y <= point.y <= y + h

    def _iter_handle_rects_display(self):
        if not self.crop_rect:
            return []
        half = HANDLE_SIZE // 2
        x, y, w, h = self.crop_rect
        left = x
        top = y
        right = x + w
        bottom = y + h
        center_x = x + w / 2
        center_y = y + h / 2
        points = {
            "top_left": (left, top),
            "top": (center_x, top),
            "top_right": (right, top),
            "right": (right, center_y),
            "bottom_right": (right, bottom),
            "bottom": (center_x, bottom),
            "bottom_left": (left, bottom),
            "left": (left, center_y),
        }
        for name, (px, py) in points.items():
            yield name, wx.Rect(int(round(px)) - half, int(round(py)) - half, HANDLE_SIZE, HANDLE_SIZE)

    def _iter_handle_rects_panel(self):
        for name, rect in self._iter_handle_rects_display():
            rect.Offset(self.offset_x, self.offset_y)
            yield name, rect

    def _hit_test_handle(self, point):
        for name, rect in self._iter_handle_rects_display():
            if rect.Contains(point):
                return name
        return None

    def _get_aspect_ratio(self):
        if not self.fixed_aspect:
            return None
        try:
            wr, hr = map(float, self.crop_aspect.split(':'))
            if hr == 0:
                return None
            return wr / hr
        except Exception:
            return None

    def _ensure_within_display(self, rect):
        if rect is None:
            return wx.Rect()
        rect = wx.Rect(rect)
        if self.display_w <= 0 or self.display_h <= 0:
            return wx.Rect()
        if rect.width > self.display_w:
            rect.width = self.display_w
        if rect.height > self.display_h:
            rect.height = self.display_h
        if rect.x < 0:
            rect.x = 0
        if rect.y < 0:
            rect.y = 0
        if rect.Right > self.display_w:
            rect.x = self.display_w - rect.width
        if rect.Bottom > self.display_h:
            rect.y = self.display_h - rect.height
        rect.width = max(MIN_CROP_SIZE, rect.width)
        rect.height = max(MIN_CROP_SIZE, rect.height)
        rect.x = max(0, min(rect.x, self.display_w - rect.width))
        rect.y = max(0, min(rect.y, self.display_h - rect.height))
        return rect

    def _ensure_min_size(self, rect):
        rect = self._ensure_within_display(rect)
        if rect.width < MIN_CROP_SIZE:
            rect.width = MIN_CROP_SIZE
        if rect.height < MIN_CROP_SIZE:
            rect.height = MIN_CROP_SIZE
        return self._ensure_within_display(rect)

    def _create_rect_with_ratio(self, anchor, current, ratio):
        dx = current.x - anchor.x
        dy = current.y - anchor.y
        abs_dx = abs(dx)
        abs_dy = abs(dy)
        if abs_dx == 0 and abs_dy == 0:
            return wx.Rect(anchor.x, anchor.y, 0, 0)
        if abs_dy == 0:
            abs_dy = int(round(abs_dx / ratio))
        if abs_dx == 0:
            abs_dx = int(round(abs_dy * ratio))
        current_ratio = abs_dx / abs_dy if abs_dy else ratio
        if current_ratio > ratio:
            abs_dx = int(round(abs_dy * ratio))
        else:
            abs_dy = int(round(abs_dx / ratio))
        x2 = anchor.x + (abs_dx if dx >= 0 else -abs_dx)
        y2 = anchor.y + (abs_dy if dy >= 0 else -abs_dy)
        left = min(anchor.x, x2)
        top = min(anchor.y, y2)
        rect = wx.Rect(left, top, abs(x2 - anchor.x), abs(y2 - anchor.y))
        return self._ensure_within_display(rect)

    def _create_rect(self, anchor, current):
        anchor = self._clamp_display_point(anchor)
        current = self._clamp_display_point(current)
        ratio = self._get_aspect_ratio()
        if ratio:
            rect = self._create_rect_with_ratio(anchor, current, ratio)
        else:
            left = min(anchor.x, current.x)
            top = min(anchor.y, current.y)
            width = abs(current.x - anchor.x)
            height = abs(current.y - anchor.y)
            rect = wx.Rect(left, top, width, height)
        return self._ensure_min_size(rect)

    def _rect_from_crop(self):
        if not self.crop_rect:
            return None
        x, y, w, h = self.crop_rect
        return wx.Rect(int(round(x)), int(round(y)), int(round(w)), int(round(h)))

    def _update_selection_creation(self, anchor, current):
        rect = self._create_rect(anchor, current)
        self.crop_rect = (rect.x, rect.y, rect.width, rect.height)

    def _update_selection_move(self, dx, dy):
        if not self.original_rect:
            return
        rect = wx.Rect(self.original_rect)
        rect.x += dx
        rect.y += dy
        rect = self._ensure_min_size(rect)
        self.crop_rect = (rect.x, rect.y, rect.width, rect.height)

    def _rect_from_horizontal_anchor(self, anchor_x, width, origin, to_left, ratio):
        width = max(MIN_CROP_SIZE, min(width, self.display_w))
        height = max(MIN_CROP_SIZE, int(round(width / ratio)))
        if height > self.display_h:
            height = self.display_h
            width = max(MIN_CROP_SIZE, int(round(height * ratio)))
        center_y = origin.y + origin.height / 2
        top = int(round(center_y - height / 2))
        top = max(0, min(top, self.display_h - height))
        bottom = top + height
        if to_left:
            right = min(anchor_x, self.display_w)
            left = max(0, right - width)
        else:
            left = max(0, anchor_x)
            right = min(self.display_w, left + width)
            left = right - width
        return wx.Rect(int(left), int(top), int(right - left), int(bottom - top))

    def _rect_from_vertical_anchor(self, anchor_y, height, origin, to_top, ratio):
        height = max(MIN_CROP_SIZE, min(height, self.display_h))
        width = max(MIN_CROP_SIZE, int(round(height * ratio)))
        if width > self.display_w:
            width = self.display_w
            height = max(MIN_CROP_SIZE, int(round(width / ratio)))
        center_x = origin.x + origin.width / 2
        left = int(round(center_x - width / 2))
        left = max(0, min(left, self.display_w - width))
        right = left + width
        if to_top:
            bottom = min(anchor_y, self.display_h)
            top = max(0, bottom - height)
        else:
            top = max(0, anchor_y)
            bottom = min(self.display_h, top + height)
            top = bottom - height
        return wx.Rect(int(left), int(top), int(right - left), int(bottom - top))

    def _resize_corner_with_ratio(self, point, handle, origin, ratio):
        if handle == "top_left":
            anchor = wx.Point(origin.Right, origin.Bottom)
            horizontal = -1
            vertical = -1
        elif handle == "top_right":
            anchor = wx.Point(origin.x, origin.Bottom)
            horizontal = 1
            vertical = -1
        elif handle == "bottom_left":
            anchor = wx.Point(origin.Right, origin.y)
            horizontal = -1
            vertical = 1
        else:
            anchor = wx.Point(origin.x, origin.y)
            horizontal = 1
            vertical = 1
        dx = (point.x - anchor.x) * horizontal
        dy = (point.y - anchor.y) * vertical
        dx = max(MIN_CROP_SIZE, min(abs(dx), self.display_w))
        dy = max(MIN_CROP_SIZE, min(abs(dy), self.display_h))
        if dy == 0:
            dy = int(round(dx / ratio))
        if dx == 0:
            dx = int(round(dy * ratio))
        width = dx
        height = int(round(width / ratio))
        if height > dy:
            height = dy
            width = int(round(height * ratio))
        height = max(MIN_CROP_SIZE, height)
        width = max(MIN_CROP_SIZE, width)
        if horizontal < 0:
            left = anchor.x - width
            right = anchor.x
        else:
            left = anchor.x
            right = anchor.x + width
        if vertical < 0:
            top = anchor.y - height
            bottom = anchor.y
        else:
            top = anchor.y
            bottom = anchor.y + height
        rect = wx.Rect(int(left), int(top), int(right - left), int(bottom - top))
        return self._ensure_min_size(rect)

    def _resize_with_ratio(self, point, handle, ratio):
        if not self.original_rect:
            return wx.Rect()
        origin = wx.Rect(self.original_rect)
        if handle == "left":
            anchor_x = origin.Right
            width = max(MIN_CROP_SIZE, min(anchor_x - point.x, anchor_x))
            rect = self._rect_from_horizontal_anchor(anchor_x, width, origin, to_left=True, ratio=ratio)
        elif handle == "right":
            anchor_x = origin.x
            width = max(MIN_CROP_SIZE, min(point.x - anchor_x, self.display_w - anchor_x))
            rect = self._rect_from_horizontal_anchor(anchor_x, width, origin, to_left=False, ratio=ratio)
        elif handle == "top":
            anchor_y = origin.Bottom
            height = max(MIN_CROP_SIZE, min(anchor_y - point.y, anchor_y))
            rect = self._rect_from_vertical_anchor(anchor_y, height, origin, to_top=True, ratio=ratio)
        elif handle == "bottom":
            anchor_y = origin.y
            height = max(MIN_CROP_SIZE, min(point.y - anchor_y, self.display_h - anchor_y))
            rect = self._rect_from_vertical_anchor(anchor_y, height, origin, to_top=False, ratio=ratio)
        else:
            rect = self._resize_corner_with_ratio(point, handle, origin, ratio)
        return self._ensure_min_size(rect)

    def _resize_free(self, point, handle):
        rect = wx.Rect(self.original_rect)
        left = rect.x
        top = rect.y
        right = rect.Right
        bottom = rect.Bottom
        if "left" in handle:
            left = min(point.x, right - MIN_CROP_SIZE)
        if "right" in handle:
            right = max(point.x, left + MIN_CROP_SIZE)
        if "top" in handle:
            top = min(point.y, bottom - MIN_CROP_SIZE)
        if "bottom" in handle:
            bottom = max(point.y, top + MIN_CROP_SIZE)
        left = max(0, min(left, self.display_w))
        right = max(0, min(right, self.display_w))
        top = max(0, min(top, self.display_h))
        bottom = max(0, min(bottom, self.display_h))
        width = max(MIN_CROP_SIZE, right - left)
        height = max(MIN_CROP_SIZE, bottom - top)
        return wx.Rect(int(left), int(top), int(width), int(height))

    def _update_selection_resize(self, point):
        if not self.original_rect or not self.drag_handle:
            return
        ratio = self._get_aspect_ratio()
        if ratio:
            rect = self._resize_with_ratio(point, self.drag_handle, ratio)
        else:
            rect = self._resize_free(point, self.drag_handle)
        rect = self._ensure_min_size(rect)
        self.crop_rect = (rect.x, rect.y, rect.width, rect.height)

    def _update_cursor(self, point):
        handle = self._hit_test_handle(point)
        if handle in ("top_left", "bottom_right"):
            cursor = wx.Cursor(wx.CURSOR_SIZENWSE)
        elif handle in ("top_right", "bottom_left"):
            cursor = wx.Cursor(wx.CURSOR_SIZENESW)
        elif handle in ("left", "right"):
            cursor = wx.Cursor(wx.CURSOR_SIZEWE)
        elif handle in ("top", "bottom"):
            cursor = wx.Cursor(wx.CURSOR_SIZENS)
        elif self._point_in_display(point):
            if self.crop_rect and self._rect_contains_point(self.crop_rect, point):
                cursor = wx.Cursor(wx.CURSOR_SIZING)
            else:
                cursor = wx.Cursor(wx.CURSOR_CROSS)
        else:
            cursor = wx.Cursor(wx.CURSOR_ARROW)
        if cursor != self._current_cursor:
            self.SetCursor(cursor)
            self._current_cursor = cursor

    def ApplyAspectRatioToSelection(self):
        if not self.fixed_aspect or not self.crop_rect:
            return
        ratio = self._get_aspect_ratio()
        if not ratio:
            return
        x, y, w, h = self.crop_rect
        center_x = x + w / 2
        center_y = y + h / 2
        width = w
        height = h
        desired_height = int(round(width / ratio))
        desired_width = int(round(height * ratio))
        if desired_height <= self.display_h:
            height = desired_height
        if height <= 0:
            height = MIN_CROP_SIZE
        width = int(round(height * ratio))
        if width > self.display_w:
            width = self.display_w
            height = int(round(width / ratio))
        if height > self.display_h:
            height = self.display_h
            width = int(round(height * ratio))
        width = max(MIN_CROP_SIZE, width)
        height = max(MIN_CROP_SIZE, height)
        left = int(round(center_x - width / 2))
        top = int(round(center_y - height / 2))
        rect = self._ensure_min_size(wx.Rect(left, top, width, height))
        self.crop_rect = (rect.x, rect.y, rect.width, rect.height)

    def SetImage(self, img):
        # 変更前は常に InitCropRect() していたため、前回の crop_rect を保持するように修正
        prev_rect = self.crop_rect
        self.current_image = img.copy()
        self._cached_bitmap = None
        self.UpdateDisplayGeometry()
        if prev_rect:
            # 前回の位置・サイズをクリップして再利用
            self.crop_rect = self.ClipRect(*prev_rect)
        else:
            self.InitCropRect()
        self.mode = "idle"
        self.drag_handle = None
        self.original_rect = None
        self.drag_start = wx.Point()
        self.Refresh()

    def UpdateDisplayGeometry(self):
        if not self.current_image:
            return
        pw, ph = self.GetClientSize()
        iw, ih = self.current_image.size
        scale = min(pw/iw, ph/ih)
        self.display_w = int(iw * scale)
        self.display_h = int(ih * scale)
        self.offset_x = (pw - self.display_w)//2
        self.offset_y = (ph - self.display_h)//2
        self._cached_bitmap = None

    def OnSize(self, evt):
        # 表示領域サイズ変更に伴いクロップ矩形を再スケール
        self.old_display_w, self.old_display_h = self.display_w, self.display_h
        self.UpdateDisplayGeometry()
        if self.crop_rect:
            self.RescaleCropRect()
        self.Refresh()
        evt.Skip()

    def OnPaint(self, evt):
        dc = wx.BufferedPaintDC(self)
        self._render_to_dc(dc)

    def _ensure_cached_bitmap(self):
        if (not self.current_image or self.display_w <= 0 or self.display_h <= 0):
            self._cached_bitmap = None
            self._cached_size = (0, 0)
            self._cached_img_id = None
            return
        if (self._cached_bitmap is None or
            self._cached_size != (self.display_w, self.display_h) or
            self._cached_img_id != id(self.current_image)):
            tmp = self.current_image.resize((self.display_w, self.display_h), Image.BILINEAR)
            raw = tmp.convert('RGB').tobytes()
            self._cached_bitmap = wx.Bitmap.FromBuffer(self.display_w, self.display_h, raw)
            self._cached_size = (self.display_w, self.display_h)
            self._cached_img_id = id(self.current_image)

    def _render_to_dc(self, dc):
        dc.SetBackground(wx.Brush(self.GetBackgroundColour()))
        dc.Clear()
        if not self.current_image or self.display_w <= 0 or self.display_h <= 0:
            return
        self._ensure_cached_bitmap()
        if not self._cached_bitmap:
            return
        pos_x = self.offset_x
        pos_y = self.offset_y
        dc.DrawBitmap(self._cached_bitmap, pos_x, pos_y)
        if self.crop_rect:
            crop_x = self.crop_rect[0] + pos_x
            crop_y = self.crop_rect[1] + pos_y
            crop_w = self.crop_rect[2]
            crop_h = self.crop_rect[3]
            gc = wx.GraphicsContext.Create(dc)
            if gc:
                overlay_path = gc.CreatePath()
                panel_w, panel_h = self.GetClientSize()
                overlay_path.AddRectangle(0, 0, panel_w, panel_h)
                rect_path = gc.CreatePath()
                rect_path.AddRectangle(crop_x, crop_y, crop_w, crop_h)
                overlay_path.AddPath(rect_path)
                overlay_path.CloseSubpath()
                gc.SetBrush(wx.Brush(wx.Colour(0, 0, 0, OVERLAY_ALPHA), wx.BRUSHSTYLE_SOLID))
                gc.FillPath(overlay_path, wx.ODDEVEN_RULE)
                gc.SetPen(wx.Pen(wx.Colour(255, 0, 0), 1, wx.PENSTYLE_SOLID))
                gc.StrokePath(rect_path)
            try:
                target_dc = wx.GCDC(dc)
            except Exception:
                target_dc = dc
            target_dc.SetPen(wx.Pen(wx.Colour(255, 0, 0), 1))
            target_dc.SetBrush(wx.Brush(wx.Colour(255, 255, 255)))
            for _, handle_rect in self._iter_handle_rects_panel():
                target_dc.DrawRectangle(handle_rect)

    def RenderToBitmap(self):
        width, height = self.GetClientSize()
        if width <= 0 or height <= 0:
            return None
        bmp = wx.Bitmap(width, height)
        dc = wx.MemoryDC()
        dc.SelectObject(bmp)
        self._render_to_dc(dc)
        dc.SelectObject(wx.NullBitmap)
        return bmp

    def pil_to_wx_image(self, pil_img: Image.Image) -> wx.Image:
        """
        PillowのImageをwx.Imageに変換する。
        - RGB / RGBA対応（L等はRGBへ変換）
        - EXIFの回転も補正（必要な場合）
        """
        # （任意）EXIFの向きを正規化
        pil = ImageOps.exif_transpose(pil_img)

        # wx.ImageはRGB + 独立αを想定
        if pil.mode == "RGBA":
            r, g, b, a = pil.split()
            rgb_bytes   = Image.merge("RGB", (r, g, b)).tobytes()
            alpha_bytes = a.tobytes()

            w, h = pil.size
            wx_img = wx.Image(w, h)
            # SetData/SetAlphaは所有権をwx側へ渡すのでbytearrayを使うのが安全
            wx_img.SetData(bytearray(rgb_bytes))
            wx_img.SetAlpha(bytearray(alpha_bytes))
            return wx_img

        else:
            if pil.mode != "RGB":
                pil = pil.convert("RGB")
            w, h = pil.size
            wx_img = wx.Image(w, h)
            wx_img.SetData(bytearray(pil.tobytes()))
            return wx_img

    def CopyOriginalToClipboard(self):

        if not self.current_image:
            return False
        pil_image = self.current_image.copy()
        width, height = pil_image.size
        if width <= 0 or height <= 0:
            return False

        pil_image = ImageOps.exif_transpose(pil_image)

        buffer = io.BytesIO()
        # Always encode as PNG for clipboard use
        pil_image.save(buffer, format="PNG")
        png_bytes = buffer.getvalue()
        buffer.close()

        if not self.clip.Open():
            print('Clipbord not open')
            return False

        try:
            data_object = wx.CustomDataObject(wx.DataFormat(wx.DF_PNG))
            data_object.SetData(png_bytes)
            self.clip.SetData(data_object)
            self.clip.Flush()                   # このコードは必要
            print('Copy to clipboard')
            return True
        finally:
            self.clip.Close()

    def InitCropRect(self):
        disp_w = self.display_w
        disp_h = self.display_h
        try:
            if self.fixed_aspect:
                wr, hr = map(float, self.crop_aspect.split(':'))
                if wr < hr:
                    rect_h = disp_h // 4
                    rect_w = int(rect_h * (wr / hr))
                else:
                    rect_w = disp_w // 4
                    rect_h = int(rect_w * (hr / wr)) if wr != 0 else disp_h // 4
            else:
                rect_w = disp_w // 4
                rect_h = rect_w
        except Exception:
            rect_w = disp_w // 4
            rect_h = rect_w
        center_x = disp_w / 2
        center_y = disp_h / 2
        new_x = center_x - rect_w / 2
        new_y = center_y - rect_h / 2
        rect = wx.Rect(int(round(new_x)), int(round(new_y)), int(round(rect_w)), int(round(rect_h)))
        rect = self._ensure_min_size(rect)
        self.crop_rect = (rect.x, rect.y, rect.width, rect.height)
        self.ApplyAspectRatioToSelection()
        self.mode = "idle"
        self.drag_handle = None
        self.original_rect = None
        self.drag_start = wx.Point()
        # ???????????????????????
        self.UpdateControls()

    def HitTestEdge(self,mx,my,rx,ry,rw,rh):
        near_left = abs(mx - rx) < EDGE_THRESHOLD
        near_right = abs(mx - (rx + rw)) < EDGE_THRESHOLD
        near_top = abs(my - ry) < EDGE_THRESHOLD
        near_bottom = abs(my - (ry + rh)) < EDGE_THRESHOLD
        inside_x = rx <= mx <= rx + rw
        inside_y = ry <= my <= ry + rh
        if near_left and near_top:
            return 'TOP_LEFT'
        if near_right and near_top:
            return 'TOP_RIGHT'
        if near_left and near_bottom:
            return 'BOTTOM_LEFT'
        if near_right and near_bottom:
            return 'BOTTOM_RIGHT'
        if near_left and inside_y: return 'LEFT'
        if near_right and inside_y: return 'RIGHT'
        if near_top and inside_x: return 'TOP'
        if near_bottom and inside_x: return 'BOTTOM'
        return None

    def ClipRect(self, x, y, w, h):
        rect = wx.Rect(int(round(x)), int(round(y)), int(round(w)), int(round(h)))
        rect = self._ensure_min_size(rect)
        return (rect.x, rect.y, rect.width, rect.height)

    def OnLeftDown(self, evt):
        if not self.current_image:
            return
        display_point = self._event_to_display_point(evt)
        handle = self._hit_test_handle(display_point)
        if handle and self.crop_rect:
            self.mode = "resizing"
            self.drag_handle = handle
            self.original_rect = self._rect_from_crop()
            self.drag_start = self._clamp_display_point(display_point)
        elif self.crop_rect and self._rect_contains_point(self.crop_rect, display_point):
            self.mode = "moving"
            self.drag_handle = "inside"
            self.original_rect = self._rect_from_crop()
            self.drag_start = self._clamp_display_point(display_point)
        elif self._point_in_display(display_point):
            self.mode = "creating"
            self.drag_handle = None
            anchor = self._clamp_display_point(display_point)
            self.drag_start = anchor
            self.original_rect = None
            self.crop_rect = (anchor.x, anchor.y, 0, 0)
        else:
            self.mode = "idle"
            self.drag_handle = None
            self.original_rect = None
            return
        if not self.HasCapture():
            self.CaptureMouse()
        self.Refresh(False)

    def OnMouseMove(self, evt):
        if not self.current_image:
            return
        display_point = self._event_to_display_point(evt)
        if evt.Dragging() and evt.LeftIsDown() and self.mode != "idle":
            point = self._clamp_display_point(display_point)
            if self.mode == "creating":
                self._update_selection_creation(self.drag_start, point)
            elif self.mode == "moving" and self.original_rect:
                dx = point.x - self.drag_start.x
                dy = point.y - self.drag_start.y
                self._update_selection_move(dx, dy)
            elif self.mode == "resizing" and self.original_rect:
                self._update_selection_resize(point)
            self.UpdateControls()
            self.Refresh(False)
            return
        self._update_cursor(display_point)

    def OnLeftUp(self, evt):
        previous_mode = self.mode
        if self.HasCapture():
            self.ReleaseMouse()
        self.mode = "idle"
        self.drag_handle = None
        self.original_rect = None
        self.drag_start = wx.Point()
        if previous_mode == "creating" and self.crop_rect:
            rect = self._ensure_min_size(self._rect_from_crop())
            self.crop_rect = (rect.x, rect.y, rect.width, rect.height)
        self.UpdateControls()
        self._update_cursor(self._event_to_display_point(evt))
        evt.Skip()

    def OnMouseLeave(self, evt):
        if self.mode == "idle":
            self.SetCursor(wx.Cursor(wx.CURSOR_ARROW))
            self._current_cursor = wx.Cursor(wx.CURSOR_ARROW)

    def UpdateControls(self):
        if getattr(self, '_is_editing', False):
            return    # ユーザー入力中は反映を抑制
        frame = wx.GetTopLevelParent(self)
        box = self.GetCropBox()
        if box is None:
            return
        xs, ys, xe, ye = box
        frame.ctrl.textcs['xs'].SetValue(str(xs))
        frame.ctrl.textcs['ys'].SetValue(str(ys))
        frame.ctrl.textcs['xe'].SetValue(str(xe))
        frame.ctrl.textcs['ye'].SetValue(str(ye))

    def RescaleCropRect(self):
        if not self.crop_rect or self.old_display_w == 0 or self.old_display_h == 0:
            return
        x, y, w, h = self.crop_rect
        cx_old = x + w / 2
        cy_old = y + h / 2
        sx = self.display_w / self.old_display_w
        sy = self.display_h / self.old_display_h
        new_w = w * sx
        new_h = h * sy
        cx_new = cx_old * sx
        cy_new = cy_old * sy
        new_x = cx_new - new_w / 2
        new_y = cy_new - new_h / 2
        rect = self._ensure_min_size(wx.Rect(int(round(new_x)), int(round(new_y)), int(round(new_w)), int(round(new_h))))
        self.crop_rect = (rect.x, rect.y, rect.width, rect.height)

    def GetCropBox(self):
        # crop_rect が未定義、または画像がない場合は None を返す
        if not self.crop_rect or not self.current_image:
            return None

        # 表示用矩形 (x, y, w, h)
        x, y, w, h = self.crop_rect
        # 元画像サイズ
        iw, ih = self.current_image.size

        # 表示サイズとの比率
        scale_x = iw / self.display_w
        scale_y = ih / self.display_h

        # 元画像上の座標に変換
        orig_x1 = int(math.floor(x * scale_x))
        orig_y1 = int(math.floor(y * scale_y))
        orig_x2 = int(math.ceil((x + w) * scale_x))
        orig_y2 = int(math.ceil((y + h) * scale_y))

        orig_x1 = max(0, min(iw - 1, orig_x1))
        orig_y1 = max(0, min(ih - 1, orig_y1))
        orig_x2 = max(orig_x1 + 1, min(iw, orig_x2))
        orig_y2 = max(orig_y1 + 1, min(ih, orig_y2))

        return (orig_x1, orig_y1, orig_x2, orig_y2)

class ControlPanel(wx.Panel):
    """右コントロールパネル"""
    def __init__(self,parent,main):
        super().__init__(parent)
        self.SetMinSize((RIGHT_PANEL_WIDTH,-1))
        self.main=main
        font=wx.Font(17,wx.FONTFAMILY_DEFAULT,wx.FONTSTYLE_NORMAL,wx.FONTWEIGHT_NORMAL)
        self.textcs={}
        s=wx.BoxSizer(wx.VERTICAL)
        g=wx.GridSizer(4,3,10,10)

        for lbl,key in [('XS','xs'),('YS','ys'),('XE','xe'),('YE','ye')]:
            st=wx.StaticText(self,label=f"{lbl}:")
            st.SetFont(font)
            g.Add(st,0,wx.ALIGN_CENTER_VERTICAL)
            tc=wx.TextCtrl(self,style=wx.TE_CENTER|wx.TE_PROCESS_ENTER,size=(140,35))
            tc.SetFont(font)
            tc.Bind(wx.EVT_TEXT_ENTER,self.OnCoordEnter)
            self.textcs[key]=tc
            g.Add(tc,0)
            # 空をセット
            st=wx.StaticText(self,label="")
            #st.SetFont(font)
            g.Add(st,0,wx.ALIGN_CENTER_VERTICAL)

        s.Add(g,0,wx.ALL|wx.EXPAND,10)
        hb=wx.BoxSizer(wx.HORIZONTAL)
        self.cb_aspect=wx.CheckBox(self,label="縦横比ロック")
        self.cb_aspect.SetValue(True)
        self.cb_aspect.Bind(wx.EVT_CHECKBOX,self.OnAspectToggle)
        hb.Add(self.cb_aspect,0,wx.ALIGN_CENTER_VERTICAL)
        self.tc_aspect=wx.TextCtrl(self,value="1:1",style=wx.TE_CENTER|wx.TE_PROCESS_ENTER,size=(131,35))
        self.tc_aspect.SetFont(font)
        self.tc_aspect.Bind(wx.EVT_TEXT_ENTER,self.OnAspectToggle)
        hb.Add(self.tc_aspect,0,wx.LEFT,5)
        s.Add(hb,0,wx.LEFT|wx.RIGHT|wx.EXPAND,10)
        b1=wx.Button(self,label="もどる",size=(100,45))
        b1.SetFont(font)
        b1.Bind(wx.EVT_BUTTON,lambda e:main.OnRevertAll())
        b2=wx.Button(self,label="トリミング＆保存",size=(100,45))
        b2.SetFont(font)
        b2.Bind(wx.EVT_BUTTON,lambda e:main.OnTrimAll())
        s.Add(b1,0,wx.ALL|wx.EXPAND,10)
        s.Add(b2,0,wx.LEFT|wx.RIGHT|wx.BOTTOM|wx.EXPAND,10)
        b_png_reduce = wx.Button(self, label="PNG減色＆保存", size=(100,45))
        b_png_reduce.SetFont(font)
        b_png_reduce.Bind(wx.EVT_BUTTON, lambda e: main.OnPngReduce())
        s.Add(b_png_reduce,0,wx.LEFT|wx.RIGHT|wx.BOTTOM|wx.EXPAND,10)
        b3 = wx.Button(self, label="スナップショット", size=(100,45))
        b3.SetFont(font)
        b3.Bind(wx.EVT_BUTTON, lambda e: main.OnSnapshot())
        s.Add(b3,0,wx.LEFT|wx.RIGHT|wx.BOTTOM|wx.EXPAND,10)
        # クリアボタン追加
        b4 = wx.Button(self, label="クリア", size=(100,45))
        b4.SetFont(font)
        b4.Bind(wx.EVT_BUTTON, lambda e: main.OnClearAll())
        s.Add(b4,0,wx.LEFT|wx.RIGHT|wx.BOTTOM|wx.EXPAND,10)
        self.SetSizer(s)
        self._is_editing = False

    def OnAspectToggle(self,event):
        lock=self.cb_aspect.GetValue()
        self.main.preview.fixed_aspect=lock
        if lock:
            self.main.preview.crop_aspect=self.tc_aspect.GetValue()
            if self.main.preview.crop_rect:
                self.main.preview.ApplyAspectRatioToSelection()
            else:
                self.main.preview.InitCropRect()
        self.main.preview.Refresh()
        # コントロールの座標表示を更新
        self.main.preview.UpdateControls()

    def OnCoordEnter(self, event):
        self._is_editing = True
        try:
            tc = event.GetEventObject()
            key = [k for k, v in self.textcs.items() if v is tc][0]
            box = self.GetValidatedBox()
            if box is None:
                return
            # 既存の矩形を取得
            old_box = self.main.preview.GetCropBox()
            if old_box:
                old_xs, old_ys, old_xe, old_ye = old_box
            else:
                old_xs, old_ys, old_xe, old_ye = box
            xs, ys, xe, ye = box
            # 縦横比ロック時の新ロジック
            if self.cb_aspect.GetValue():
                wr, hr = map(float, self.tc_aspect.GetValue().split(':'))
                if key == 'xs':
                    # 固定サイズでX方向平行移動
                    delta = xs - old_xs
                    xe = old_xe + delta
                elif key == 'ys':
                    # 固定サイズでY方向平行移動
                    delta = ys - old_ys
                    ye = old_ye + delta
                elif key == 'xe':
                    # XS固定でサイズ調整
                    new_w = xe - xs
                    new_h = new_w * (hr / wr)
                    ye = ys + new_h
                elif key == 'ye':
                    # YS固定でサイズ調整
                    new_h = ye - ys
                    new_w = new_h * (wr / hr)
                    xe = xs + new_w
                # 指定アスペクトとずれていればロック解除
                if abs((xe - xs) / (ye - ys) - wr / hr) > 0.01:
                    self.cb_aspect.SetValue(False)

            # 入力値を表示座標に変換し、ラバーバンドに反映
            disp = self.main.preview
            iw, ih = disp.current_image.size
            sw, sh = disp.display_w, disp.display_h
            scale_x = sw / iw
            scale_y = sh / ih
            dxs = round(xs * scale_x)
            dys = round(ys * scale_y)
            dxe = round(xe * scale_x)
            dye = round(ye * scale_y)
            rect = wx.Rect(dxs, dys, dxe - dxs, dye - dys)
            rect = disp._ensure_min_size(rect)
            disp.crop_rect = (rect.x, rect.y, rect.width, rect.height)
            disp.Refresh()
            disp.UpdateControls()
        finally:
            self._is_editing = False

    def GetValidatedBox(self):
        try:
            xs=int(self.textcs['xs'].GetValue())
            ys=int(self.textcs['ys'].GetValue())
            xe=int(self.textcs['xe'].GetValue())
            ye=int(self.textcs['ye'].GetValue())
        except ValueError:
            wx.MessageBox("座標は整数で入力してください。","エラー",wx.OK|wx.ICON_ERROR)
            return None
        if xe<=xs or ye<=ys:
            wx.MessageBox("XE>XSかつYE>YSにしてください。","エラー",wx.OK|wx.ICON_ERROR)
            return None
        return (xs,ys,xe,ye)

class FileDropTarget(wx.FileDropTarget):
    def __init__(self,main):
        super().__init__()
        self.main=main

    def OnDropFiles(self,x,y,files):
        self.main.AddFiles(files)
        return True

class MainFrame(wx.Frame):
    def __init__(self):
        super().__init__(None,title="Batch-Cropper",size=APP_WINDOW_SIZE)    
        self.SetMinSize(APP_WINDOW_SIZE)
        self._default_aspect = APP_WINDOW_SIZE[0] / APP_WINDOW_SIZE[1]
        self.file_paths=[]
        self.images=[]
        self.history=[]
        self.reduced_flags=[]
        self.selected_index=-1
        self.splitter=wx.SplitterWindow(self)
        left=wx.Panel(self.splitter)
        right=wx.Panel(self.splitter,size=(RIGHT_PANEL_WIDTH,-1))
        self.splitter.SplitVertically(left,right,sashPosition=1200-RIGHT_PANEL_WIDTH)
        self.Bind(wx.EVT_SIZE,self.OnFrameResize)
        self.Bind(wx.EVT_MOUSEWHEEL,self.OnMouseWheelResize)
        lv=wx.BoxSizer(wx.VERTICAL)
        self.preview=PreviewPanel(left)
        self.thumbnails=ThumbnailPanel(left,self.OnSelectThumbnail)
        lv.Add(self.preview,1,wx.EXPAND|wx.ALL,5)
        lv.Add(self.thumbnails,0,wx.EXPAND|wx.ALL,5)
        left.SetSizer(lv)
        rv=wx.BoxSizer(wx.VERTICAL)
        self.ctrl=ControlPanel(right,self)
        rv.Add(self.ctrl,0,wx.EXPAND)
        right.SetSizer(rv)
        self.SetDropTarget(FileDropTarget(self))

        accel = wx.AcceleratorTable([
            (wx.ACCEL_CTRL, ord('C'), wx.ID_COPY),
            (wx.ACCEL_CTRL, ord('V'), wx.ID_PASTE),
        ])
        self.SetAcceleratorTable(accel)
        self.Bind(wx.EVT_MENU, self.OnCopyPreviewOriginal, id=wx.ID_COPY)
        self.Bind(wx.EVT_MENU, self.OnPasteFromClipboard, id=wx.ID_PASTE)

        self.Centre()
        self.Show()
        _log_debug(f"APP start log={LOG_PATH}")

    def _get_clipboard_image(self):
        """クリップボードから画像を取得する。画像が無ければ None。"""
        try:
            data = ImageGrab.grabclipboard()
        except Exception as e:
            wx.LogError(f"Clipboard access failed: {e}")
            return None

        if isinstance(data, Image.Image):
            return data

        if isinstance(data, list):
            for item in data:
                ext = os.path.splitext(item)[1].lower()
                if ext in ('.jpg', '.jpeg', '.png', '.bmp', '.tiff', '.tif'):
                    try:
                        with Image.open(item) as img_file:
                            return img_file.copy()
                    except Exception:
                        continue

        wx.LogWarning("Clipboard does not contain an image.")
        return None

    def _save_import_image(self, image: Image.Image, prefix: str, ext: str = ".png"):
        """指定パスに画像を保存し、保存パス(Path)を返す。失敗時は None。"""
        save_dir = resolve_import_dir()
        file_path = build_unique_path(save_dir, prefix, ext)
        try:
            image.save(file_path)
        except Exception as e:
            wx.LogError(f"Failed to save image: {e}")
            return None
        return file_path

    def OnPasteFromClipboard(self, evt):
        """Ctrl+V でクリップボードの画像を取り込む。"""
        img = self._get_clipboard_image()
        if img is None:
            return

        file_path = self._save_import_image(img, "clipboard", ".png")
        if file_path is None:
            return

        self.file_paths.append(str(file_path))
        self.images.append(img.copy())
        self.reduced_flags.append(False)
        # 取り込みは加工のUndo対象外にしたいので、ベース状態だけを履歴に残す
        self.history.clear()
        self.PushHistory()
        self.selected_index = len(self.images) - 1
        self.UpdateUI()

    def OnFrameResize(self,evt):
        w,h=evt.GetSize()
        self.splitter.SetSashPosition(w-RIGHT_PANEL_WIDTH)
        evt.Skip()

    def _get_current_display_rect(self):
        """ウィンドウが属するモニターのクライアント領域を返す"""
        idx = wx.Display.GetFromWindow(self)
        if idx == wx.NOT_FOUND:
            idx = 0
        display = wx.Display(idx)
        try:
            return display.GetClientArea()
        except Exception:
            w, h = wx.GetDisplaySize()
            return wx.Rect(0, 0, w, h)

    def OnMouseWheelResize(self, evt):
        rotation = evt.GetWheelRotation()
        if rotation == 0:
            return
        direction = 1 if rotation > 0 else -1
        display_rect = self._get_current_display_rect()
        min_w, min_h = APP_WINDOW_SIZE
        aspect = self._default_aspect
        max_w = min(display_rect.width, int(round(display_rect.height * aspect)))
        max_w = max(min_w, max_w)
        current_w = self.GetSize().width
        target_w = current_w * (1 + WINDOW_RESIZE_SCALE_STEP * direction)
        target_w = max(min_w, min(target_w, max_w))
        target_w = int(round(target_w))
        target_h = int(round(target_w / aspect))
        pos_x = int(display_rect.x + (display_rect.width - target_w) // 2)
        pos_y = int(display_rect.y + (display_rect.height - target_h) // 2)
        self.SetSize(wx.Size(target_w, target_h))
        self.Move(wx.Point(pos_x, pos_y))

    def AddFiles(self,paths):
        for p in paths:
            ext = os.path.splitext(p)[1].lower()
            if ext in ('.jpg','.jpeg','.png','.bmp','.tiff', '.tif') and p not in self.file_paths:
                try:
                    # タグ情報を残すため、ここでは convert をせずにそのまま保持
                    img = Image.open(p)
                    self.file_paths.append(p)
                    self.images.append(img)
                    self.reduced_flags.append(False)
                except: wx.LogError(f"読み込み失敗:{p}")
        if self.images:
            self.history.clear()
            self.PushHistory()
            self.selected_index=0
            self.UpdateUI()

    def UpdateUI(self):
        _log_debug(f"UI update selected={self.selected_index} images={len(self.images)}")
        if 0 <= self.selected_index < len(self.images):
            _log_debug(f"UI current id={id(self.images[self.selected_index])} path={self.file_paths[self.selected_index]}")
            self.preview.SetImage(self.images[self.selected_index])
            box = self.preview.GetCropBox()
            if box:
                xs, ys, xe, ye = box
                self.ctrl.textcs['xs'].SetValue(str(xs))
                self.ctrl.textcs['ys'].SetValue(str(ys))
                self.ctrl.textcs['xe'].SetValue(str(xe))
                self.ctrl.textcs['ye'].SetValue(str(ye))
        self.thumbnails.update_thumbnails(self.images)      # このコードは必要


    def OnCopyPreviewOriginal(self, evt):
        if not (0 <= self.selected_index < len(self.file_paths)):
            wx.LogWarning('コピー対象のプレビュー画像がありません。')
            return

        ext = os.path.splitext(self.file_paths[self.selected_index])[1].lower()
        if ext in ('.jpg', '.jpeg'):
            wx.MessageBox("容量が増えるためJpeg画像はクリップボードにコピーできません\n挿入から画像を取り込んでください", "コピーを中止しました", wx.OK | wx.ICON_INFORMATION)
            return
        if self.reduced_flags and self.reduced_flags[self.selected_index]:
            wx.MessageBox("減色後のクリップボードへのコピーは容量低減効果がありません\n挿入から画像を取り込んでください", "コピーを中止しました", wx.OK | wx.ICON_INFORMATION)
            return

        if not self.preview.CopyOriginalToClipboard():
            wx.LogWarning('コピー対象のプレビュー画像がありません。')

    def OnSelectThumbnail(self, idx):
        self.selected_index = idx
        self.preview.SetImage(self.images[idx])
        box = self.preview.GetCropBox()
        if box:
            xs, ys, xe, ye = box
            self.ctrl.textcs['xs'].SetValue(str(xs))
            self.ctrl.textcs['ys'].SetValue(str(ys))
            self.ctrl.textcs['xe'].SetValue(str(xe))
            self.ctrl.textcs['ye'].SetValue(str(ye))

    def OnTrimAll(self):
        _log_debug("TRIM start")
        box = self.ctrl.GetValidatedBox()
        if box is None:
            return
        crop_box = box
        logs = []
        new_paths = []
        new_images = []
        new_flags = []
        for path, img in zip(self.file_paths, self.images):
            try:
                # ① トリミング
                trimmed = img.crop(crop_box)

                # ② 保存先と拡張子判定
                _, ext = os.path.splitext(path)
                ext_lower = ext.lower()
                out_path = add_bc_suffix(path)

                save_kwargs = {}

                if ext_lower in ('.jpg', '.jpeg'):
                    # JPEG は quality を指定
                    save_kwargs['quality'] = img.info.get('quality', 80)

                elif ext_lower in ('.tif', '.tiff'):
                    # TIFF の Compression タグを取得 (259番)
                    comp = None
                    try:
                        comp = img.tag_v2.get(259)
                    except Exception:
                        pass
                    print(comp)
                    # Group-3 (3) / Group-4 (4) のときは 1-bit モードに変換して再圧縮
                    if comp == 3:
                        trimmed = trimmed.convert('1')
                        save_kwargs['compression'] = 'group3'
                    elif comp == 4:
                        trimmed = trimmed.convert('1')
                        save_kwargs['compression'] = 'group4'
                    # それ以外の TIFF（LZW, PackBits など）は何も追加せず

                # ③ 保存
                trimmed.save(out_path, **save_kwargs)

                # ④ メモリ内イメージ更新
                with Image.open(out_path) as reopened:
                    new_images.append(reopened.copy())
                new_paths.append(out_path)
                new_flags.append(self.reduced_flags[self.file_paths.index(path)])
                logs.append(f"{os.path.basename(path)} → OK")
            except Exception as e:
                logs.append(f"{os.path.basename(path)} → ERROR: {e}")

        # 履歴に登録＆UI更新
        if new_images:
            self.file_paths = new_paths
            self.images = new_images
            self.reduced_flags = new_flags
            self.PushHistory()
            if not (0 <= self.selected_index < len(self.images)):
                self.selected_index = 0
            self.UpdateUI()
        _log_debug(f"TRIM end new_count={len(new_images)}")
        # （必要なら logs をファイル出力 or ダイアログ表示）


    def OnPngReduce(self):
        """PNG reduce for PNG files, save with _bc and reload"""
        if not self.file_paths:
            wx.MessageBox("減色対象のファイルがありません。", "情報", wx.OK | wx.ICON_INFORMATION)
            return

        non_png = [p for p in self.file_paths if os.path.splitext(p)[1].lower() != '.png']
        if non_png:
            wx.MessageBox("減色はPNGファイルのみ対応しています。", "情報", wx.OK | wx.ICON_INFORMATION)
            return

        new_paths = []
        new_images = []
        new_flags = []
        for path, img in zip(self.file_paths, self.images):
            try:
                rgba = img.convert("RGBA")
                alpha = rgba.getchannel("A")
                rgb = rgba.convert("RGB")
                quantized = rgb.quantize(colors=256, method=Image.MEDIANCUT, dither=Image.FLOYDSTEINBERG)
                # アルファがあれば戻す（RGBA保存時は減色効果は控えめだが表示を優先）
                if alpha.getextrema() != (255, 255):
                    quantized = quantized.convert("RGBA")
                    quantized.putalpha(alpha)
                out_path = add_bc_suffix(path)
                quantized.save(out_path, optimize=True)
                with Image.open(out_path) as reopened:
                    new_images.append(reopened.copy())
                new_paths.append(out_path)
                new_flags.append(True)
            except Exception as e:
                wx.LogError(f"減色に失敗しました: {path}: {e}")

        if not new_paths:
            return

        self.file_paths = new_paths
        self.images = new_images
        self.reduced_flags = new_flags
        self.PushHistory()
        if not (0 <= self.selected_index < len(self.images)):
            self.selected_index = 0
        self.UpdateUI()


    def OnSnapshot(self):
        """Capture full screen, save it, and load into the tool."""
        _log_debug("SNAPSHOT start")
        try:
            screenshot = ImageGrab.grab(all_screens=True)
        except Exception as e:
            wx.LogError(f"Failed to capture snapshot: {e}")
            return

        save_dir = resolve_import_dir()
        file_path = build_unique_path(save_dir, "snapshot", ".png")
        try:
            screenshot.save(file_path)
        except Exception as e:
            wx.LogError(f"Failed to save snapshot: {e}")
            return

        self.file_paths.append(str(file_path))
        img_copy = screenshot.copy()
        self.images.append(img_copy)
        self.reduced_flags.append(False)
        # 取り込みは加工のUndo対象外にしたいので、ベース状態だけを履歴に残す
        self.history.clear()
        self.PushHistory()
        self.selected_index = len(self.images) - 1
        self.UpdateUI()
        _log_debug(f"SNAPSHOT added path={file_path} id={id(img_copy)} images={len(self.images)}")

    def OnRevertAll(self):
        if len(self.history) < 2:
            # wx.MessageBox("No more history", "Info", wx.OK)
            return
        # 現在の状態を破棄し、前の状態へ戻す
        current_paths = list(self.file_paths)
        self.history.pop()
        previous = self.history[-1]
        prev_paths = previous["paths"]
        # いまのステップで生成された _bc のみ削除
        for p in current_paths:
            base, _ = os.path.splitext(p)
            if p not in prev_paths and base.endswith("_bc"):
                try:
                    os.remove(p)
                except Exception:
                    pass
        # 巻き戻し先が _bc ならファイル内容も復元しておく
        for p, img in zip(prev_paths, previous["images"]):
            base, _ = os.path.splitext(p)
            if base.endswith("_bc"):
                try:
                    img.save(p)
                except Exception:
                    pass
        self.file_paths = list(prev_paths)
        self.images = [img.copy() for img in previous["images"]]
        self.reduced_flags = list(previous.get("flags", [False]*len(self.images)))
        if not (0 <= self.selected_index < len(self.images)):
            self.selected_index = 0
        self.UpdateUI()

    def PushHistory(self):
        """現在のファイルパスと画像を履歴に保存"""
        snapshot = {
            "paths": list(self.file_paths),
            "images": [img.copy() for img in self.images],
            "flags": list(self.reduced_flags),
        }
        self.history.append(snapshot)
        if len(self.history) > MAX_HISTORY:
            self.history.pop(0)

    def OnClearAll(self):
        """読み込んだファイルを破棄して初期状態に戻す"""
        _log_debug(f"CLEAR start images={len(self.images)} cache={len(self.thumbnails._bitmap_cache)}")
        self.file_paths.clear()
        self.images.clear()
        self.history.clear()
        self.reduced_flags.clear()
        self.selected_index = -1
        # プレビューをリセット
        self.preview.current_image = None
        self.preview._cached_bitmap = None
        self.preview.crop_rect = None
        self.preview.Refresh()
        # サムネイルをクリア
        self.thumbnails._bitmap_cache.clear()
        self.thumbnails.update_thumbnails([])
        # コントロールパネルのテキストボックスを空に
        for tc in self.ctrl.textcs.values():
            tc.SetValue("")
        _log_debug(f"CLEAR end images={len(self.images)} cache={len(self.thumbnails._bitmap_cache)}")

if __name__=='__main__':
    Image.MAX_IMAGE_PIXELS=None
    app=wx.App(False)
    MainFrame()
    app.MainLoop()
    
