"""窓を隠したときに画面の右端に出す、小さなつまみ（Windows）。

Chrome や Edge の窓は外から「常に手前」にできないので、Sidenote 自身の小さな窓を別に出す。
つまみは画面共有とスクリーンショットに写らないようにする（Windows の SetWindowDisplayAffinity）。
上を押すと窓が戻り、下を押すと画面を撮る。上下にドラッグして位置を変えられる。
"""
import ctypes
import queue
import threading
from ctypes import wintypes

WIDTH, HEIGHT, ZONE = 30, 88, 44
COLOR, COLOR_REC, COLOR_DONE = "#4a657e", "#c4772f", "#3f8f5f"
WDA_EXCLUDEFROMCAPTURE = 0x11
GA_ROOT = 2


class MONITORINFO(ctypes.Structure):
    _fields_ = [("cbSize", wintypes.DWORD), ("rcMonitor", wintypes.RECT), ("rcWork", wintypes.RECT), ("dwFlags", wintypes.DWORD)]


def work_area(hwnd):
    """窓 hwnd がある画面の、タスクバーを除いた範囲"""
    user32 = ctypes.windll.user32
    user32.MonitorFromWindow.restype = wintypes.HANDLE
    user32.MonitorFromWindow.argtypes = [wintypes.HWND, wintypes.DWORD]
    user32.GetMonitorInfoW.argtypes = [wintypes.HANDLE, ctypes.POINTER(MONITORINFO)]
    info = MONITORINFO()
    info.cbSize = ctypes.sizeof(info)
    user32.GetMonitorInfoW(user32.MonitorFromWindow(hwnd, 2), ctypes.byref(info))
    w = info.rcWork
    return w.left, w.top, w.right, w.bottom


class EdgeTab:
    def __init__(self, on_open, on_shot, on_move=lambda ratio: None):
        self.on_open, self.on_shot, self.on_move = on_open, on_shot, on_move
        self.q = queue.Queue()
        self.thread = None
        self.recording = False
        self.flashing = 0

    def _ensure(self):
        if not self.thread:
            self.thread = threading.Thread(target=self._run, daemon=True)
            self.thread.start()

    def show(self, area, ratio=0.5):
        """画面の範囲 area の右端、上から ratio の高さに出す"""
        self._ensure()
        self.q.put(("show", (area, ratio)))

    def hide(self):
        if self.thread:
            self.q.put(("hide", None))

    def set_recording(self, on):
        self.recording = on
        if self.thread:
            self.q.put(("color", None))

    def flash(self):
        if self.thread:
            self.q.put(("flash", None))

    def _run(self):
        import tkinter as tk
        root = tk.Tk()
        root.title("Sidenote つまみ")
        root.overrideredirect(True)
        root.attributes("-topmost", True)
        root.withdraw()
        scale = root.winfo_fpixels("1i") / 96
        w, h, zone = [int(v * scale) for v in (WIDTH, HEIGHT, ZONE)]
        c = tk.Canvas(root, width=w, height=h, highlightthickness=0, bg=COLOR, cursor="hand2")
        c.pack()
        font = ("Segoe UI Symbol", 11)
        c.create_text(w / 2, zone / 2, text="◀", fill="white", font=font)
        c.create_line(4, zone, w - 4, zone, fill="#ffffff")
        c.create_text(w / 2, zone + (h - zone) / 2, text="📷", fill="white", font=font)
        state = {"area": None, "shown": False, "drag": None}
        root.update_idletasks()

        def recolor():
            c.configure(bg=COLOR_DONE if self.flashing else COLOR_REC if self.recording else COLOR)

        def top_hwnd():
            return wintypes.HWND(ctypes.windll.user32.GetAncestor(wintypes.HWND(root.winfo_id()), GA_ROOT))

        def place(y):
            # 枠なしの窓は Tk の geometry では動かなかった（実測）ので、Windows に直接頼む
            l, t, r, b = state["area"]
            state["y"] = max(t, min(y, b - h))
            set_pos = ctypes.windll.user32.SetWindowPos
            set_pos.argtypes = [wintypes.HWND, wintypes.HWND, ctypes.c_int, ctypes.c_int, ctypes.c_int, ctypes.c_int, ctypes.c_uint]
            set_pos(top_hwnd(), wintypes.HWND(-1), r - w, state["y"], w, h, 0x0010)  # -1 = 常に手前 / 0x10 = 前面の窓を奪わない

        def press(e):
            state["drag"] = (e.y_root, state["y"], False, e.y)

        def motion(e):
            d = state["drag"]
            if not d:
                return
            moved = d[2] or abs(e.y_root - d[0]) > 4
            state["drag"] = (d[0], d[1], moved, d[3])
            if moved:
                place(d[1] + e.y_root - d[0])

        def release(e):
            d, state["drag"] = state["drag"], None
            if not d:
                return
            if d[2]:  # ドラッグだった。位置を覚える
                t, b = state["area"][1], state["area"][3]
                self.on_move((state["y"] - t) / max(1, b - t - h))
            elif d[3] < zone:
                self.on_open()
            else:
                self.on_shot()

        c.bind("<ButtonPress-1>", press)
        c.bind("<B1-Motion>", motion)
        c.bind("<ButtonRelease-1>", release)

        def end_flash():
            self.flashing = max(0, self.flashing - 1)
            recolor()

        def poll():
            try:
                while True:
                    kind, val = self.q.get_nowait()
                    if kind == "show":
                        state["area"], ratio = val
                        t, b = state["area"][1], state["area"][3]
                        root.deiconify()
                        root.update_idletasks()
                        place(t + int((b - t - h) * ratio))
                        # 画面共有とスクリーンショットに写らないようにする（古い Windows で失敗しても動きは変わらない）
                        ctypes.windll.user32.SetWindowDisplayAffinity(top_hwnd(), WDA_EXCLUDEFROMCAPTURE)
                        # 押しても、使っている窓（発表中の資料など）から入力の先を奪わない
                        style = ctypes.windll.user32.GetWindowLongW(top_hwnd(), -20)
                        ctypes.windll.user32.SetWindowLongW(top_hwnd(), -20, style | 0x08000000)  # WS_EX_NOACTIVATE
                        state["shown"] = True
                        recolor()
                    elif kind == "hide":
                        root.withdraw()
                        state["shown"] = False
                    elif kind == "color":
                        recolor()
                    elif kind == "flash":
                        self.flashing += 1
                        recolor()
                        root.after(700, end_flash)
            except queue.Empty:
                pass
            if state["shown"]:
                root.attributes("-topmost", True)  # 全画面の資料などが前に出ても埋もれないように
            root.after(150, poll)

        poll()
        root.mainloop()
