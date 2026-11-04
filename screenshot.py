"""Sidenote の窓を写さずに画面を撮る（Windows）。

押した瞬間だけ Sidenote の窓を画面の外へ逃がし、撮ったらすぐ元の位置へ戻す。隠すのではなく動かすので、
窓の消える動きを待たずに済み、ほかの窓の手前後ろも変わらない。窓の後ろに隠れていた部分も写る。
"""
import ctypes
import re
import time
from ctypes import wintypes

from PIL import ImageGrab

user32 = ctypes.WinDLL("user32", use_last_error=True)
dwmapi = ctypes.WinDLL("dwmapi")
user32.GetWindow.restype = wintypes.HWND
user32.GetWindow.argtypes = [wintypes.HWND, wintypes.UINT]
user32.GetWindowLongW.argtypes = [wintypes.HWND, ctypes.c_int]
user32.MonitorFromWindow.restype = wintypes.HANDLE
user32.MonitorFromWindow.argtypes = [wintypes.HWND, wintypes.DWORD]
user32.GetTopWindow.restype = wintypes.HWND
user32.GetTopWindow.argtypes = [wintypes.HWND]
user32.SetWindowPos.argtypes = [wintypes.HWND, wintypes.HWND, ctypes.c_int, ctypes.c_int, ctypes.c_int, ctypes.c_int, ctypes.c_uint]
user32.SetThreadDpiAwarenessContext.restype = ctypes.c_void_p
user32.SetThreadDpiAwarenessContext.argtypes = [ctypes.c_void_p]

SWP_NOSIZE, SWP_NOZORDER, SWP_NOACTIVATE = 0x0001, 0x0004, 0x0010
GW_HWNDNEXT, GWL_EXSTYLE, WS_EX_TOOLWINDOW = 2, -20, 0x00000080
OFFSCREEN = -32000
REPAINT_WAIT = 0.15  # 窓をどけたあと、後ろの窓が描き直されるのを待つ秒数
PER_MONITOR_V2 = ctypes.c_void_p(-4)
# Edge の InPrivate で開くと題名が「Sidenote - [InPrivate]」になる。同じ名前のフォルダの窓（エクスプローラー）は動かさない
TITLE = re.compile(r"^Sidenote( - .*InPrivate.*)?$")
NOT_CONTENT = {"Progman", "WorkerW", "Shell_TrayWnd", "Shell_SecondaryTrayWnd"}


class MONITORINFO(ctypes.Structure):
    _fields_ = [("cbSize", wintypes.DWORD), ("rcMonitor", wintypes.RECT), ("rcWork", wintypes.RECT), ("dwFlags", wintypes.DWORD)]


def _title(hwnd):
    buf = ctypes.create_unicode_buffer(256)
    user32.GetWindowTextW(hwnd, buf, 256)
    return buf.value


def _class(hwnd):
    buf = ctypes.create_unicode_buffer(256)
    user32.GetClassNameW(hwnd, buf, 256)
    return buf.value


def sidenote_windows():
    """画面に出ている Sidenote の窓（Chrome か Edge のアプリの窓）を返す"""
    found = []

    @ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.HWND, wintypes.LPARAM)
    def each(hwnd, _):
        if user32.IsWindowVisible(hwnd) and _class(hwnd) == "Chrome_WidgetWin_1" and TITLE.match(_title(hwnd)):
            found.append(hwnd)
        return True

    user32.EnumWindows(each, 0)
    return found


def _rect(hwnd):
    r = wintypes.RECT()
    user32.GetWindowRect(hwnd, ctypes.byref(r))
    return r


def _is_content(hwnd):
    """使う人が見ている窓らしいか（壁紙・タスクバー・隠れた窓を除く）"""
    if not user32.IsWindowVisible(hwnd) or user32.IsIconic(hwnd) or not _title(hwnd) or _class(hwnd) in NOT_CONTENT:
        return False
    if user32.GetWindowLongW(hwnd, GWL_EXSTYLE) & WS_EX_TOOLWINDOW:
        return False
    cloaked = wintypes.DWORD(0)
    dwmapi.DwmGetWindowAttribute(wintypes.HWND(hwnd), 14, ctypes.byref(cloaked), ctypes.sizeof(cloaked))  # 14 = DWMWA_CLOAKED
    if cloaked.value:
        return False
    r = _rect(hwnd)
    return r.right - r.left >= 200 and r.bottom - r.top >= 200


def _monitor_rect(hwnd):
    info = MONITORINFO()
    info.cbSize = ctypes.sizeof(info)
    user32.GetMonitorInfoW(user32.MonitorFromWindow(hwnd, 2), ctypes.byref(info))  # 2 = いちばん近い画面
    m = info.rcMonitor
    return m.left, m.top, m.right, m.bottom


def target_monitor(skip):
    """撮る画面を決める。Sidenote のすぐ後ろの窓（資料など、直前まで使っていた窓）がある画面を選ぶ"""
    hwnd = user32.GetTopWindow(None)
    while hwnd:
        if hwnd not in skip and _is_content(hwnd):
            return _monitor_rect(hwnd)
        hwnd = user32.GetWindow(hwnd, GW_HWNDNEXT)
    if skip:
        return _monitor_rect(skip[0])
    return (0, 0, user32.GetSystemMetrics(0), user32.GetSystemMetrics(1))


def capture(path, bbox=None):
    """Sidenote の窓をどけて画面を1枚撮り、PNG で保存する。窓は失敗しても必ず元の位置へ戻す"""
    old = user32.SetThreadDpiAwarenessContext(PER_MONITOR_V2)  # 画面ごとに拡大率が違っても位置がずれないように
    moved = []
    try:
        windows = sidenote_windows()
        bbox = bbox or target_monitor(windows)
        for h in windows:
            r = _rect(h)
            moved.append((h, r.left, r.top, r.right - r.left, r.bottom - r.top))
            user32.SetWindowPos(h, None, OFFSCREEN, OFFSCREEN, 0, 0, SWP_NOSIZE | SWP_NOZORDER | SWP_NOACTIVATE)
        if moved:
            dwmapi.DwmFlush()
            time.sleep(REPAINT_WAIT)
        image = ImageGrab.grab(bbox=bbox, all_screens=True)
    finally:
        for h, x, y, w, ht in moved:
            user32.SetWindowPos(h, None, x, y, w, ht, SWP_NOZORDER | SWP_NOACTIVATE)
        user32.SetThreadDpiAwarenessContext(old)
    image.save(path, "PNG")  # 圧縮に時間がかかるので、窓を戻してから保存する
    return image.size
