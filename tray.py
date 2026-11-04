"""Sidenote の常駐アイコン（画面右下の通知領域）。pywin32 で作る。

左クリックで窓を開き、右クリックで「開く／終了」のメニューを出す。
記録中は、アイコンを icon_rec.ico（橙の点付き）に替え、マウスを乗せたときの説明にも「記録中」と出す。
周りの人や本人が、いま記録しているかどうかをいつでも分かるようにするため。
"""
import threading
from pathlib import Path

import win32api
import win32con
import win32gui

WM_TRAY = win32con.WM_USER + 20
WM_SETREC = win32con.WM_USER + 21  # 別のスレッドから「記録中かどうか」を知らせる合図
ID_OPEN, ID_QUIT = 1023, 1024


class Tray:
    def __init__(self, icon_path, tip, on_open, on_quit, rec_icon_path=None):
        self.icon_path, self.tip = str(icon_path), tip
        p = Path(rec_icon_path) if rec_icon_path else Path(icon_path).with_name("icon_rec.ico")
        self.rec_icon_path = str(p)
        self.on_open, self.on_quit = on_open, on_quit
        self.hwnd = None
        self.recording = False

    def _load(self, path):
        try:
            return win32gui.LoadImage(0, path, win32con.IMAGE_ICON, 0, 0,
                                      win32con.LR_LOADFROMFILE | win32con.LR_DEFAULTSIZE)
        except win32gui.error:
            return win32gui.LoadIcon(0, win32con.IDI_APPLICATION)

    def _icon_args(self):
        hicon = self._load(self.rec_icon_path if self.recording else self.icon_path)
        tip = f"{self.tip}（記録中）" if self.recording else self.tip
        flags = win32gui.NIF_ICON | win32gui.NIF_MESSAGE | win32gui.NIF_TIP
        return (self.hwnd, 0, flags, WM_TRAY, hicon, tip)

    def _add_icon(self):
        win32gui.Shell_NotifyIcon(win32gui.NIM_ADD, self._icon_args())

    def set_recording(self, on):
        """記録中かどうかでアイコンと説明を替える。どのスレッドから呼んでもよい"""
        if self.hwnd:
            win32gui.PostMessage(self.hwnd, WM_SETREC, 1 if on else 0, 0)
        else:
            self.recording = bool(on)

    def _menu(self):
        menu = win32gui.CreatePopupMenu()
        win32gui.AppendMenu(menu, win32con.MF_STRING, ID_OPEN, "開く")
        win32gui.AppendMenu(menu, win32con.MF_SEPARATOR, 0, "")
        win32gui.AppendMenu(menu, win32con.MF_STRING, ID_QUIT, "終了")
        win32gui.SetMenuDefaultItem(menu, ID_OPEN, False)
        pos = win32gui.GetCursorPos()
        win32gui.SetForegroundWindow(self.hwnd)  # これが無いとメニューの外を押しても閉じない
        win32gui.TrackPopupMenu(menu, win32con.TPM_LEFTALIGN, pos[0], pos[1], 0, self.hwnd, None)
        win32gui.PostMessage(self.hwnd, win32con.WM_NULL, 0, 0)

    def run(self, quit_event: threading.Event):
        """通知領域にアイコンを出し、quit_event が立つまでメッセージを回す（呼んだスレッドで動く）"""
        # エクスプローラーが再起動すると通知領域のアイコンが消えるので、この合図で出し直す
        taskbar_created = win32gui.RegisterWindowMessage("TaskbarCreated")

        def proc(hwnd, msg, wp, lp):
            if msg == WM_TRAY:
                if lp in (win32con.WM_LBUTTONUP, win32con.WM_LBUTTONDBLCLK):
                    self.on_open()
                elif lp == win32con.WM_RBUTTONUP:
                    self._menu()
            elif msg == WM_SETREC:
                self.recording = bool(wp)
                win32gui.Shell_NotifyIcon(win32gui.NIM_MODIFY, self._icon_args())
            elif msg == win32con.WM_COMMAND:
                cmd = win32api.LOWORD(wp)
                if cmd == ID_OPEN:
                    self.on_open()
                elif cmd == ID_QUIT:
                    self.on_quit()
            elif msg == win32con.WM_QUERYENDSESSION:
                return True
            elif msg == win32con.WM_ENDSESSION and wp:
                # Windows の終了・再起動。記録中なら止めて WAV を書き終えるまで少し待つ（待てる長さは Windows 次第）
                self.on_quit()
                quit_event.wait(15)
                return 0
            elif msg == taskbar_created:
                self._add_icon()
            elif msg == win32con.WM_CLOSE:
                win32gui.Shell_NotifyIcon(win32gui.NIM_DELETE, (hwnd, 0))
                win32gui.DestroyWindow(hwnd)
            elif msg == win32con.WM_DESTROY:
                win32gui.PostQuitMessage(0)
            return win32gui.DefWindowProc(hwnd, msg, wp, lp)

        wc = win32gui.WNDCLASS()
        wc.hInstance = win32api.GetModuleHandle(None)
        wc.lpszClassName = "SidenoteTray"
        wc.lpfnWndProc = proc
        win32gui.RegisterClass(wc)
        self.hwnd = win32gui.CreateWindow(wc.lpszClassName, "Sidenote tray", 0, 0, 0, 0, 0, 0, 0,
                                          wc.hInstance, None)
        self._add_icon()

        def watch():  # 終了の合図（記録を書き終えたあと）が来たら、アイコンを消してメッセージの回転を止める
            quit_event.wait()
            win32gui.PostMessage(self.hwnd, win32con.WM_CLOSE, 0, 0)

        threading.Thread(target=watch, daemon=True).start()
        win32gui.PumpMessages()
