"""Sidenote — 画面との窓口（このPCの中だけで通じるサーバー）と起動の入口。

画面は ui.html（HTML）で、Chrome（無ければ Edge）の「アプリとして開く」機能（--app）で表示する。アドレス欄やタブの無い窓になる。
Python は 127.0.0.1 で待ち受け、画面からの操作（開始・停止・質問など）を受け、出来事（発言・残り時間・答えなど）を
流し続ける（Server-Sent Events）。録音・文字起こし・Claude・engram（選べる機能）の中身は meeting_notes.py にある。
他のアプリやウェブページから勝手に操作されないよう、起動のたびに作る合言葉（token）が無い要求は断る。
起動: pythonw meeting_app.py （スタートメニューとデスクトップの「Sidenote」）。
Windows の起動時はスタートアップの「Sidenote」が --background で常駐させる（窓は出さない）。
常駐中は画面右下の通知領域にアイコンが出る（記録中は橙の点付きのアイコンになる）。
窓を閉じても記録は裏で続く。終わるのはアイコンの「終了」だけ。
"""
import ctypes
import datetime as dt
import hashlib
import http.server
import json
import os
import queue
import re
import secrets
import subprocess
import sys
import threading
import time
import urllib.parse
import urllib.request
import webbrowser
from ctypes import wintypes
from pathlib import Path

import meeting_notes as mn

HERE = Path(__file__).parent
UI_FILE = HERE / "ui.html"
# 画面は起動したときのものを使い続ける。動いている中身と画面の版が食い違うと、ボタンが押せなくなるなどの誤動作になる
UI_HTML = UI_FILE.read_text(encoding="utf-8")
STARTED = time.time()
ICON_FILE = HERE / "icon.ico"
STATE_FILE = mn.SETTINGS.parent / "server.json"
PROFILE_ROOT = Path(os.path.expandvars(r"%LOCALAPPDATA%\Sidenote"))  # いつものブラウザとは別の、Sidenote 専用の設定フォルダを置く場所
NO_WINDOW = mn.NO_WINDOW


LOCAL = urllib.request.build_opener(urllib.request.ProxyHandler({}))  # このパソコンの中の通信は、職場のプロキシを通さない


def _find_exe(folder, exe):
    """ブラウザの場所を探す。パソコンによって入っている場所が違う。見つからなければ None"""
    candidates = [Path(os.environ.get(v, "")) / folder / "Application" / exe
                  for v in ("ProgramFiles", "ProgramFiles(x86)", "LOCALAPPDATA") if os.environ.get(v)]
    try:
        import winreg
        for root in (winreg.HKEY_LOCAL_MACHINE, winreg.HKEY_CURRENT_USER):
            try:
                with winreg.OpenKey(root, rf"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\{exe}") as k:
                    candidates.append(Path(winreg.QueryValue(k, None)))
            except OSError:
                pass
    except ImportError:
        pass
    return next((str(p) for p in candidates if p.is_file()), None)


def find_browser():
    """画面を出すブラウザを選ぶ。Chrome があれば Chrome、無ければ Edge。返り値は (場所, 名前)。どちらも無ければ (None, "")"""
    chrome = _find_exe(Path("Google") / "Chrome", "chrome.exe")
    if chrome:
        return chrome, "chrome"
    edge = _find_exe(Path("Microsoft") / "Edge", "msedge.exe")
    return (edge, "edge") if edge else (None, "")


def tag_window(hwnd):
    """窓に「これは Sidenote だ」という情報を付ける。無いと、タスクバーにピン留めしたときに
    Windows が Chrome（または Edge）として覚え、アイコンも起動先もそちらになってしまう"""
    from win32com.propsys import propsys, pscon
    pythonw = Path(sys.executable).with_name("pythonw.exe")
    store = propsys.SHGetPropertyStoreForWindow(hwnd, propsys.IID_IPropertyStore)
    for key, value in ((pscon.PKEY_AppUserModel_ID, "RicoAiLab.Sidenote"),
                       (pscon.PKEY_AppUserModel_RelaunchCommand, f'"{pythonw}" "{HERE / "meeting_app.py"}"'),
                       (pscon.PKEY_AppUserModel_RelaunchDisplayNameResource, "Sidenote"),
                       (pscon.PKEY_AppUserModel_RelaunchIconResource, f"{ICON_FILE},0")):
        store.SetValue(key, propsys.PROPVARIANTType(value))
    store.Commit()


def pick_files(folder=False, title="Sidenote の保存先を選ぶ"):
    """ファイル（またはフォルダ）を選ぶ画面を、別の小さな Python で出して結果を受け取る"""
    code = ("import tkinter as tk, tkinter.filedialog as fd, json\n"
            "r = tk.Tk(); r.withdraw(); r.attributes('-topmost', True)\n")
    if folder:
        code += f"print(json.dumps(fd.askdirectory(title={title!r}, initialdir={str(mn.DATA)!r})))"
    else:
        code += ("print(json.dumps(list(fd.askopenfilenames(title='会議の資料を選ぶ', filetypes=["
                 "('資料', '*.pdf *.docx *.xlsx *.pptx *.doc *.xls *.ppt *.txt *.md *.csv *.png *.jpg *.jpeg'),"
                 "('すべて', '*.*')]))))")
    r = subprocess.run([sys.executable, "-X", "utf8", "-c", code], capture_output=True, text=True,
                       encoding="utf-8", creationflags=NO_WINDOW)
    try:
        return json.loads(r.stdout.strip() or "null")
    except ValueError:
        return None


def payload(m):
    if not m:
        return None
    return {"id": m.folder.name, "title": m.meta.get("title", ""), "lines": m.lines(), "chats": m.chats(),
            "docs": [p.name for p in m.docs()], "recorded_ms": m.meta.get("recorded_ms", 0),
            "clean": m.clean_lines(), "sound": m.meta.get("sound"), "has_audio": bool(mn.audio_parts(m.folder)),
            "has_clean_audio": (m.folder / mn.CLEAN_WAV).exists(),
            "engram_id": m.meta.get("engram_id"), "folder": str(m.folder)}


def update_check():
    """起動のあとで Sidenote のファイルが新しくなっていたら知らせる。窓を閉じるだけでは裏で動き続けるので、古いままになるため"""
    try:
        newest = max(p.stat().st_mtime for p in (*HERE.glob("*.py"), UI_FILE))
    except OSError:
        return []
    if newest <= STARTED:
        return []
    return [{"name": "新しい版", "ok": False, "need": False,
             "help": "ファイルが更新されています。通知領域のアイコンを右クリックして「終了」を押し、開き直すと新しい版になります"}]


class Hub:
    """画面と中身の橋渡し。録音・質問などの仕事を受け、出来事を画面へ流す"""

    def __init__(self):
        self.clients = []
        self.clients_lock = threading.Lock()
        self.meeting = None
        self.recorder = None
        self.rec_events = queue.Queue()
        self.rec_started = 0.0
        self.auto_stopped = False
        self.ever_connected = False
        self.last_seen = time.time()
        self.quit = threading.Event()
        self.finishing = False
        self.importing = False  # 音声ファイルの文字起こしの最中
        self.job_lock = threading.Lock()
        self.settings_lock = threading.Lock()
        self.shot_lock = threading.Lock()
        self.tab = None  # 窓を隠したときの右端のつまみ
        self.hidden = []  # 隠している窓
        self.sound_pending = None  # 回を開いていないときに選んだ拾う音
        self.last_open = 0.0  # 窓を最後に開いた時刻
        self.ask_finish = lambda: None
        self.ask_job = None  # いま答えを待っている質問（中止ボタン用）
        self.fix_job = None  # いま動いている補正（中止ボタン用）
        self.on_rec = lambda on: None  # 記録の開始・停止を通知領域のアイコンに知らせる（main で差し替える）
        # いま選んでいる会議名のフォルダ。新しい回はこの中に作る（保存先そのもの＝会議名なし）
        p = Path(mn.load_settings().get("series_path") or mn.DATA)
        self.series = p if p.is_dir() else mn.DATA
        threading.Thread(target=self._pump, daemon=True).start()
        threading.Thread(target=self._timer, daemon=True).start()
        if not os.environ.get("MEETING_NOTES_NO_WINDOW"):
            threading.Thread(target=self._tag_windows, daemon=True).start()

    def _tag_windows(self):
        """開いた Sidenote の窓に、ピン留めのための情報を付ける。窓はブラウザが作るので、見張って見つける。
        付けられなくても、ピン留めのアイコンが変わるだけなので、ほかの動きは止めない"""
        tagged = set()
        try:
            import pythoncom
            import screenshot
            pythoncom.CoInitialize()
        except Exception:
            return
        while True:
            time.sleep(1)
            try:
                for h in screenshot.sidenote_windows():
                    if h not in tagged:
                        tag_window(h)
                        tagged.add(h)
            except Exception:
                pass

    def emit(self, kind, data=None):
        with self.clients_lock:
            for q in list(self.clients):
                q.put((kind, data))

    def status(self, text):
        self.emit("status", text)

    def settings(self):
        s = mn.load_settings()
        return {"theme": s.get("theme", "system"), "limit": s.get("limit", "2時間"),
                "diarize": s.get("diarize", True), "mic": s.get("mic", ""), "past": s.get("past", False),
                "pin": s.get("pin", True), "split": s.get("split"),
                "claude_on": s.get("claude_on", True), "engram_on": s.get("engram_on", True),
                "denoise_on": s.get("denoise_on", False), "rate": s.get("rate", 1),
                "chat_closed": s.get("chat_closed", False), "sound_default": s.get("sound_default", "mic")}

    def state(self):
        mics = mn.list_mics()
        if os.environ.get("MEETING_NOTES_TEST_FILE"):  # 試験用: 音声ファイルをマイクの代わりに選べるようにする
            mics.insert(0, "file:" + os.environ["MEETING_NOTES_TEST_FILE"])
        return {"settings": self.settings(), "limits": list(mn.LIMITS), "mics": mics, "data_dir": str(mn.DATA),
                "series": [{"path": str(p), "name": p.name if p.parent == mn.DATA else f"{p.name}（{p.parent}）"}
                           for p in mn.list_series()],
                "series_cur": "" if self.series == mn.DATA else str(self.series),
                "meetings": [{"id": m.folder.name, "label": m.label} for m in mn.all_meetings(self.series)],
                "current": payload(self.meeting), "recording": bool(self.recorder),
                "engram": bool(mn.engram_path()),  # engram がある人にだけ「engramへ」ボタンを出す
                "denoise_ok": mn.denoise_available(),  # 雑音除去が入っている speech.cpp か
                "asking": bool(self.ask_job), "fixing": bool(self.fix_job),  # 窓を開き直しても中止ボタンを出す
                "ai": mn.ai_status(),  # AI の接続先。使えないときは質問・補正を隠す（文字起こしは使える）
                "checks": mn.check_ready(mics) + update_check()}  # 準備の確認。足りないものがあれば画面の上に案内を出す

    # 録音の出来事を画面向けに変える
    def _pump(self):
        while True:
            kind, val = self.rec_events.get()
            if kind == "line":
                self.emit("line", val)
            elif kind == "level":
                self.emit("level", round(val, 3))
            elif kind == "error":
                self.status(val)
            elif kind == "stopped":
                r, self.recorder = self.recorder, None
                m = r.meeting if r else self.meeting
                problem = r.problem if r else ""
                try:  # どこで失敗しても、画面への知らせと終了の合図までは必ず進む
                    self._rec_changed(False)
                    m.write_readable()
                except Exception as e:
                    problem = problem or f"後片付けの途中で失敗しました: {str(e)[:150]}"
                self.emit("stopped", {"recorded_ms": m.meta.get("recorded_ms", 0)})
                if problem:  # 想定外の止まり方は、窓を隠していても気づけるように窓を戻して知らせる
                    self.status(problem)
                    try:
                        self.show_window()
                    except Exception:
                        pass
                else:
                    self.status(f"{self.settings()['limit']}経ったので、自動で停止しました" if self.auto_stopped
                                else "文字起こしが途中で止まりました。録音は残っています。録音のファイルを画面へドラッグすると文字にできます" if r and r.asr_lost
                                else "停止しました。もう一度開始すると同じ会議に書き足します")
                # engram へは自動で送らない。「engramへ」ボタンを押したときだけ要点を作って送る
                # （伏せ忘れた発言が勝手に送られないように。残すかどうかも会議ごとに選べる）
                if self.finishing:
                    self.quit.set()

    def _timer(self):
        # 切り忘れ対策: 開始してから決めた時間が経ったら自動で止める。残り時間を画面へ流す
        warned = time.time()
        while True:
            time.sleep(0.5)
            r = self.recorder
            if not r:
                continue
            limit = mn.LIMITS.get(self.settings()["limit"], 0)
            left = limit - (time.time() - self.rec_started) if limit else None
            self.emit("tick", {"elapsed_ms": r.elapsed_ms(), "left_s": None if left is None else max(0, int(left))})
            if r.asr_lost and time.time() - warned > 30:  # 文字が増えないまま気づかないことがないように、出し続ける
                warned = time.time()
                self.status("文字起こしが止まっています。録音は続けています。あとで録音のファイルを画面へドラッグすると文字にできます")
            if limit and left <= 0 and not self.auto_stopped:
                self.auto_stopped = True
                r.stop()

    # 操作
    def open(self, mid):
        if not self.recorder and re.fullmatch(r"[^\\/:*?\"<>|]+", mid) and mid not in (".", ".."):  # 一覧の外へ出ない
            p = self.series / mid
            if (p / "meta.json").exists():
                self.meeting = mn.Meeting(p)
                self.sound_pending = None
        return payload(self.meeting)

    def set_series(self, path):
        """会議名を切り替える（"" は会議名なし）。記録中は切り替えない"""
        if self.recorder:
            self.status("記録中は会議名を変えられません。先に停止してください")
            return
        p = Path(path) if path else mn.DATA
        if p != mn.DATA and p not in mn.list_series():
            return
        self.series, self.meeting = p, None
        with self.settings_lock:
            mn.save_settings({**mn.load_settings(), "series_path": "" if p == mn.DATA else str(p)})

    def new(self):
        if not self.recorder:
            m = mn.Meeting(self.series / dt.datetime.now().strftime("%Y-%m-%d_%H%M%S"))
            if self.sound_pending:
                m.meta["sound"] = self.sound_pending
                self.sound_pending = None
            m.save_meta()
            self.meeting = m
        return payload(self.meeting)

    def import_begin(self, name):
        """画面へドラッグされた音声ファイルのために、新しい回を作る。記録中や取り込み中は断る"""
        if self.recorder or self.importing:
            self.status("記録中や文字起こしの最中は取り込めません。終わってからもう一度ドラッグしてください")
            return None
        self.importing = True
        m = mn.Meeting(self.series / dt.datetime.now().strftime("%Y-%m-%d_%H%M%S"))
        m.meta["title"] = Path(name).stem
        m.save_meta()
        self.meeting = m
        self.status("音声を取り込んでいます")
        return m

    def import_run(self, m, src, diarize):
        def run():
            try:
                self.status("文字起こししています。音声の長さの数分の1ほどかかります")
                mn.transcribe_file(src, m, diarize)
                self.status(f"文字起こしが終わりました（{len(m.lines())}件）。音声は {m.folder.name} に保存しました")
            except Exception as e:  # 失敗しても画面は動かし続ける
                self.status(f"文字起こしできませんでした: {e}")
            finally:
                self.importing = False
                try:
                    src.unlink()  # 受け取った元ファイルの写し。文字起こし用の audio.wav は会議のフォルダに残る
                except OSError:
                    pass
                self.emit("imported", payload(m))
        threading.Thread(target=run, daemon=True).start()

    def denoise(self, diarize):
        """雑音を除いて文字にし直す。元の文字起こしは残し、別の版を作る"""
        m = self.meeting
        if not m or self.recorder or self.importing:
            self.status("記録中や文字起こしの最中はできません。止めてから押してください")
            return
        self.importing = True

        def run():
            try:
                self.status("雑音を除いて文字にし直しています（録音の長さによっては時間がかかります）")
                recs = mn.denoise_meeting(m, diarize)
                self.status(f"雑音除去後の版を作りました（{len(recs)}件）。元の文字起こしはそのままです")
            except Exception as e:
                self.status(f"雑音を除けませんでした: {e}")
            finally:
                self.importing = False
                self.emit("imported", payload(m))
        threading.Thread(target=run, daemon=True).start()

    # 拾う音は回ごとに覚える。新しい回は既定の音から始まるので、Web 会議のあとに戻し忘れない
    def current_sound(self):
        m = self.meeting
        return (m.meta.get("sound") if m else self.sound_pending) or self.settings()["sound_default"]

    def set_sound(self, value):
        if self.recorder or value not in ("mic", "both", "pc"):
            return self.current_sound()
        if self.meeting:
            self.meeting.meta["sound"] = value
            self.meeting.save_meta()
        else:  # まだ回が無いときは、次にできる回のために取っておく
            self.sound_pending = value
        return value

    def start(self, mic, diarize):
        if self.recorder:
            return
        if self.importing:
            self.status("音声ファイルの文字起こしが終わるまで、記録は始められません")
            return
        try:
            if not self.meeting:
                self.new()
            self.recorder = mn.Recorder(mic, self.meeting, self.rec_events, diarize, self.current_sound())
        except Exception as e:  # 部品が無い、保存先に書けない、など。無反応にせず理由を出す
            self.status(f"記録を始められませんでした: {str(e)[:200]}")
            return
        self.rec_started, self.auto_stopped = time.time(), False
        self._rec_changed(True)
        self.emit("started", payload(self.meeting))
        self.status(f"保存先: {self.meeting.folder}")

    def stop(self):
        if self.recorder:
            self.recorder.stop()
            self.status("最後の発言を文字にしています")

    def ask(self, q, past, web=False, connectors=False):
        m = self.meeting
        if not m or not q:
            self.status("会議を選ぶか、開始してから質問してください")
            return
        if self.ask_job:  # 1度に1つ。答えを待つあいだ、画面の送信ボタンは中止ボタンになる
            self.status("前の質問の答えを待っています。止めるときは ■ を押してください")
            return
        job = self.ask_job = mn.llm.Job()

        def finish(rec):
            rec["at"] = dt.datetime.now().isoformat(timespec="seconds")
            m.add_chat(rec)
            m.write_readable()
            self.emit("answer", {"id": m.folder.name, "q": q, "a": rec["a"], "cancelled": bool(rec.get("cancelled"))})

        def run():
            rec = {"q": q}
            try:
                rec["a"] = mn.ask_claude(q, m, past, web, connectors, job)  # 接続先が Claude Code 以外なら web と connectors は使わない
            except mn.llm.Cancelled:
                return
            except Exception as e:  # 質問の失敗で記録を止めない
                # 失敗の印を付けて残す。印のあるやり取りは、次の質問や engram の要点づくりに渡さない
                rec["a"], rec["error"] = f"（質問に失敗しました: {e}）", True
            with self.job_lock:
                if job.cancelled.is_set():  # 外部の API は止められないので、遅れて届いた答えを捨てる
                    return
                self.ask_job = None
            finish(rec)

        self.ask_finish = lambda: finish({"q": q, "a": "（中止しました）", "error": True, "cancelled": True})
        threading.Thread(target=run, daemon=True).start()

    # 窓を隠して、画面の右端に小さなつまみだけを出す
    def hide_window(self):
        import edge_tab
        import screenshot
        windows = screenshot.sidenote_windows()
        if not windows:
            return False
        if not self.tab:
            self.tab = edge_tab.EdgeTab(self.show_window, self._tab_shot, self._tab_moved)
            self.tab.set_recording(bool(self.recorder))
        self.tab.show(edge_tab.work_area(windows[0]), mn.load_settings().get("tab_pos", 0.5))
        for h in windows:
            ctypes.windll.user32.ShowWindow(wintypes.HWND(h), 0)  # 0 = 隠す
        self.hidden = windows
        return True

    def show_window(self):
        """隠した窓を元に戻す。隠していなければ False"""
        windows, self.hidden = self.hidden, []
        for h in windows:
            ctypes.windll.user32.ShowWindow(wintypes.HWND(h), 5)  # 5 = 出す
            ctypes.windll.user32.SetForegroundWindow(wintypes.HWND(h))
        if self.tab:
            self.tab.hide()
        return bool(windows)

    def focus_window(self):
        """開いている窓を手前に出す。窓が増えると、別の窓で開いた回に操作が当たってしまうため。窓が無ければ False"""
        if self.show_window():
            return True
        now = time.time()
        if now - self.last_open < 3:  # 開いている途中の窓。続けて押されても増やさない
            return True
        try:
            import screenshot
            windows = screenshot.sidenote_windows()
        except Exception:
            return False
        for h in windows[:1]:
            ctypes.windll.user32.ShowWindow(wintypes.HWND(h), 9)  # 9 = 最小化を戻す
            ctypes.windll.user32.SetForegroundWindow(wintypes.HWND(h))
        if not windows:
            self.last_open = now
        return bool(windows)

    def _tab_shot(self):
        def run():
            r = self.shot()
            if r and r.get("ok"):
                self.tab.flash()
            if r and r.get("created"):
                self.emit("opened", payload(self.meeting))
        threading.Thread(target=run, daemon=True).start()

    def _tab_moved(self, ratio):
        with self.settings_lock:
            mn.save_settings({**mn.load_settings(), "tab_pos": round(ratio, 3)})

    def _rec_changed(self, on):
        self.on_rec(on)
        if self.tab:
            self.tab.set_recording(on)

    def shot(self):
        """Sidenote の窓を写さずに画面を撮り、この回のフォルダの「スクショ」に保存する。AI には渡さない"""
        with self.shot_lock:
            created = not self.meeting
            if created:
                if self.recorder:
                    return None
                self.new()
            m, r = self.meeting, self.recorder
            folder = m.folder / mn.SHOT_DIR
            now = dt.datetime.now().strftime("%Y%m%d_%H%M%S")
            at = ""
            if r:  # 記録中なら、録音の時刻をファイル名の頭に付ける（あとで発言と照らせる）
                s = r.elapsed_ms() // 1000
                at = f"{s // 3600:02d}-{s % 3600 // 60:02d}-{s % 60:02d}_"
            try:
                import screenshot  # Pillow が無くても、ほかの機能は動くように、使うときに読む
                folder.mkdir(parents=True, exist_ok=True)
                path, n = folder / f"{at}{now}.png", 1
                while path.exists():
                    n += 1
                    path = folder / f"{at}{now}_{n}.png"
                screenshot.capture(path)
            except Exception as e:
                self.status(f"スクリーンショットを撮れませんでした: {e}")
                return {"ok": False, "id": m.folder.name, "created": created}
            self.status(f"スクリーンショットを保存しました: {mn.SHOT_DIR}\\{path.name}（AI には渡しません）")
            return {"ok": True, "id": m.folder.name, "created": created, "name": path.name}

    def cancel_ask(self):
        """質問を中止する。動いている Claude Code やローカル LLM は止まる。外部の API の答えは届いても捨てる"""
        with self.job_lock:
            job, self.ask_job = self.ask_job, None
            if job:
                job.cancel()
        if job:
            self.ask_finish()
            if mn.llm.provider(mn.load_settings()) in ("anthropic", "openai", "gemini"):
                self.status("中止しました。送ってしまった頼みは取り消せないので、API の料金はかかることがあります")
            else:
                self.status("中止しました")

    def fix(self):
        m = self.meeting
        if not m or self.fix_job:
            return
        job = self.fix_job = mn.llm.Job()

        def run():
            done = 0
            try:
                while True:
                    job.check()
                    lines = m.lines()
                    todo = [i for i, r in enumerate(lines) if not r.get("checked")]
                    if not todo:
                        break
                    start, end = todo[0], min(len(lines), todo[0] + mn.FIX_BATCH)
                    fixes = mn.fix_batch(m, lines, start, end, job)
                    m.apply_fixes(fixes, range(start, end))
                    done += len(fixes)
                    self.emit("fixed", {"id": m.folder.name, "lines": m.lines(), "done": False,
                                        "msg": f"補正中… {end}/{len(lines)} 件まで見ました（直した発言 {done} 件）"})
                msg = f"補正しました（直した発言 {done} 件）" if done else "直すところは見つかりませんでした"
            except mn.llm.Cancelled:
                msg = f"補正を止めました（直した発言 {done} 件は残っています。「再補正」で続きから直せます）"
            except Exception as e:
                msg = f"補正の途中で止まりました: {e}"
            self.fix_job = None
            m.write_readable()
            self.emit("fixed", {"id": m.folder.name, "lines": m.lines(), "done": True, "msg": msg})

        threading.Thread(target=run, daemon=True).start()

    def engram(self):
        """「engramへ」ボタン: 要点を作って engram に送る。押したときだけ動く"""
        m = self.meeting
        if not m:
            return
        if self.recorder:
            self.status("記録中は engram に送れません。停止して、伏せたい発言を伏せてから押してください")
            return
        self.status("要点をまとめて engram に残しています（20秒ほど）")

        def run():
            try:
                eid, msg = mn.save_to_engram(m)
            except Exception as e:
                eid, msg = None, f"engram に残せませんでした: {e}"
            if eid:
                m.meta["engram_id"], m.meta["engram_summary"] = eid, msg
                m.save_meta()
                m.write_readable()
                self.status(f"engram に要点を残しました（{eid}）")
            else:
                self.status(msg)

        threading.Thread(target=run, daemon=True).start()

    def attach(self):
        if not self.meeting:
            self.new()
            self.emit("opened", payload(self.meeting))
        paths = pick_files() or []
        if not paths:
            return
        m = self.meeting
        self.status(f"資料を取り込んでいます（{len(paths)}件）")
        notes = []
        for p in paths:
            try:
                dst, n = m.add_doc(Path(p))
                notes.append(f"{dst.name}（{n}字）" if n else f"{dst.name}（画像など。接続先が Claude Code のときだけ読めます）")
            except Exception as e:
                notes.append(f"{Path(p).name}（取り込めませんでした: {e}）")
        m.write_readable()
        self.emit("docs", {"id": m.folder.name, "docs": [p.name for p in m.docs()]})
        self.status("資料を添付しました: " + "、".join(notes))

    def choose_data(self):
        if self.recorder:
            self.status("記録中は保存先を変えられません。先に停止してください")
            return
        d = pick_files(folder=True)
        if d:
            mn.DATA = Path(d)
            with self.settings_lock:
                mn.save_settings({**mn.load_settings(), "data_dir": str(mn.DATA), "series_path": ""})
            self.meeting, self.series = None, mn.DATA
            self.status("保存先を変えました。これまでの記録は元の場所に残っています")


def make_handler(hub: Hub, token: str):
    class H(http.server.BaseHTTPRequestHandler):
        def log_message(self, *a):
            pass

        def _ok(self, body=b"", ctype="application/json; charset=utf-8"):
            self.send_response(200)
            self.send_header("Content-Type", ctype)
            self.send_header("Cache-Control", "no-store")
            self.end_headers()
            self.wfile.write(body)

        def _json(self, obj):
            self._ok(json.dumps(obj, ensure_ascii=False).encode("utf-8"))

        def _authed(self):
            q = urllib.parse.parse_qs(urllib.parse.urlparse(self.path).query)
            return q.get("t", [""])[0] == token or self.headers.get("X-Token") == token

        def do_GET(self):
            path = urllib.parse.urlparse(self.path).path
            if path == "/ping":
                return self._ok(b"ok", "text/plain")
            if path == "/favicon.ico" and ICON_FILE.exists():
                return self._ok(ICON_FILE.read_bytes(), "image/x-icon")
            if not self._authed():
                return self.send_error(403)
            if path == "/":
                html = UI_HTML.replace("__TOKEN__", token)
                return self._ok(html.encode("utf-8"), "text/html; charset=utf-8")
            if path == "/events":
                return self._events()
            if path == "/audio":
                return self._audio()
            self.send_error(404)

        def _audio(self):
            """開いている回の録音を画面へ流す。止めて再開した回は1本につないで見せるので、発言の時刻と合う。
            このパソコンの中だけの通信で、録音の中身は変えない"""
            q = urllib.parse.parse_qs(urllib.parse.urlparse(self.path).query)
            m = hub.meeting
            if not m or hub.recorder or m.folder.name != q.get("id", [""])[0]:  # 記録中のファイルは流さない
                return self.send_error(404)
            clean = m.folder / mn.CLEAN_WAV
            paths = [clean] if q.get("ver", [""])[0] == "clean" and clean.exists() else mn.audio_parts(m.folder)
            head, segs = mn.wav_join_plan(paths)
            if not segs:
                return self.send_error(404)
            total = len(head) + sum(s[2] for s in segs)
            start, end, status = 0, total - 1, 200
            r = re.fullmatch(r"bytes=(\d*)-(\d*)", self.headers.get("Range") or "")
            if r and (r.group(1) or r.group(2)):
                if r.group(1):
                    start = int(r.group(1))
                    end = min(int(r.group(2)), total - 1) if r.group(2) else total - 1
                else:
                    start = max(0, total - int(r.group(2)))
                if start >= total or start > end:
                    self.send_response(416)
                    self.send_header("Content-Range", f"bytes */{total}")
                    self.end_headers()
                    return
                status = 206
            self.send_response(status)
            self.send_header("Content-Type", "audio/wav")
            self.send_header("Accept-Ranges", "bytes")
            self.send_header("Content-Length", str(end - start + 1))
            if status == 206:
                self.send_header("Content-Range", f"bytes {start}-{end}/{total}")
            self.send_header("Cache-Control", "no-store")
            self.end_headers()
            try:
                self._send_span(head, segs, start, end)
            except OSError:  # 画面が途中で読むのをやめた（早送りなど）
                pass

        def _send_span(self, head, segs, start, end):
            left, pos = end - start + 1, start
            if pos < len(head):
                part = head[pos:pos + left]
                self.wfile.write(part)
                left -= len(part)
                pos = 0
            else:
                pos -= len(head)
            for path, off, n in segs:
                if left <= 0:
                    break
                if pos >= n:
                    pos -= n
                    continue
                with open(path, "rb") as f:
                    f.seek(off + pos)
                    rest = min(n - pos, left)
                    while rest > 0:
                        chunk = f.read(min(rest, 1 << 16))
                        if not chunk:  # ファイルが想定より短い
                            return
                        self.wfile.write(chunk)
                        rest -= len(chunk)
                        left -= len(chunk)
                pos = 0

        def _events(self):
            self.send_response(200)
            self.send_header("Content-Type", "text/event-stream; charset=utf-8")
            self.send_header("Cache-Control", "no-store")
            self.end_headers()
            q = queue.Queue()
            with hub.clients_lock:
                hub.clients.append(q)
            hub.ever_connected = True
            try:
                while not hub.quit.is_set():
                    try:
                        kind, data = q.get(timeout=10)
                        msg = f"event: {kind}\ndata: {json.dumps(data, ensure_ascii=False)}\n\n"
                    except queue.Empty:
                        msg = ": keep\n\n"
                    self.wfile.write(msg.encode("utf-8"))
                    self.wfile.flush()
            except OSError:
                pass
            finally:
                with hub.clients_lock:
                    hub.clients.remove(q)
                hub.last_seen = time.time()

        def _import(self, n):
            """画面へドラッグされた音声ファイルを受け取り、文字起こしを始める。外へは何も送らない"""
            q = urllib.parse.parse_qs(urllib.parse.urlparse(self.path).query)
            name, ext = q.get("name", [""])[0], Path(q.get("name", [""])[0]).suffix.lower()
            self.close_connection = True  # 断るときは本文を読まないので、この接続は使い回さない
            if ext not in mn.AUDIO_EXTS:
                hub.status("音声か動画のファイルをドラッグしてください")
                return self._json(None)
            m = hub.import_begin(name)
            if not m:
                return self._json(None)
            src = m.folder / f"_source{ext}"
            try:
                with open(src, "wb") as f:
                    left = n
                    while left > 0:
                        chunk = self.rfile.read(min(left, 1 << 20))
                        if not chunk:
                            break
                        f.write(chunk)
                        left -= len(chunk)
            except OSError as e:
                hub.importing = False
                hub.status(f"音声を保存できませんでした（空き容量を確かめてください）: {e}")
                return self._json(None)
            hub.import_run(m, src, q.get("diarize", ["1"])[0] == "1")
            return self._json(payload(m))

        def do_POST(self):
            # 1つの操作の失敗で画面全体を止めない。画面には「失敗した」（null）を返し、理由を出す
            try:
                self._post()
            except (BrokenPipeError, ConnectionError):
                pass
            except Exception as e:
                hub.status(f"操作できませんでした: {str(e)[:200]}")
                try:
                    self._json(None)
                except OSError:
                    pass

        def _post(self):
            if not self._authed():
                return self.send_error(403)
            n = int(self.headers.get("Content-Length") or 0)
            op = urllib.parse.urlparse(self.path).path.removeprefix("/api/")
            if op == "import":  # 中身が音声そのものなので、JSON として読まずにファイルへ書く
                return self._import(n)
            try:
                body = json.loads(self.rfile.read(n) or b"{}")
            except ValueError:
                body = {}
            m = hub.meeting
            s = mn.load_settings()
            if op in ("ask", "fix", "engram") and not mn.ai_status()["ready"]:  # 使わない設定のときは外へ何も送らない
                return self._json(False)
            if op == "hide_window":  # 窓を隠して、右端のつまみだけにする
                try:
                    return self._json(hub.hide_window())
                except Exception as e:
                    hub.status(f"窓を隠せませんでした: {str(e)[:150]}")
                    return self._json(False)
            if op == "show_window":  # 隠した窓を戻す。開いている窓があれば手前に出す（起動し直されたときなど）
                return self._json(hub.focus_window())
            if op == "shot":  # 画面を撮る（Sidenote の窓は写さない）
                return self._json(hub.shot())
            if op == "cancel":  # ■ ボタン。答えを待っている質問か、動いている補正を止める
                if body.get("what") == "fix":
                    if hub.fix_job:
                        hub.fix_job.cancel()
                        hub.status("補正を止めています")
                else:
                    hub.cancel_ask()
                return self._json(True)
            if op == "ai_set":  # 接続先を選ぶ。API キーは資格情報マネージャーへ入れ、settings.json には書かない
                p = str(body.get("provider") or mn.llm.provider(s))
                if p not in mn.llm.PROVIDERS:
                    return self._json(None)
                error = ""
                with hub.settings_lock:  # 続けて変えても、前の変更を消さないように読み直して書く
                    s = mn.load_settings()
                    s["ai_provider"] = p
                    cfg = dict((s.get("ai") or {}).get(p) or {})
                    if "url" in body and mn.llm.PROVIDERS[p].get("url_edit"):
                        url = str(body["url"]).strip()
                        try:
                            if url:
                                mn.llm.check_url(url)
                            cfg["url"] = url
                        except mn.llm.AIError as e:
                            error = str(e)
                    if "model" in body:
                        cfg["model"] = str(body["model"]).strip()
                    s["ai"] = {**(s.get("ai") or {}), p: cfg}
                    mn.save_settings(s)
                if mn.llm.PROVIDERS[p].get("key"):
                    if str(body.get("key") or "").strip():
                        mn.llm.set_key(p, str(body["key"]).strip())
                    elif body.get("clear_key"):
                        mn.llm.del_key(p)
                return self._json({**mn.ai_status(), "error": error})
            if op == "ai_models":  # 選べるモデルの一覧を接続先に聞く
                try:
                    if body.get("provider") in mn.llm.PROVIDERS:  # 画面が選んでいる接続先の一覧を返す
                        s["ai_provider"] = body["provider"]
                    return self._json({"models": mn.llm.models(s)})
                except Exception as e:
                    return self._json({"models": [], "error": str(e)[:300]})
            if op == "ai_test":  # 短い文だけを送って、つながるか確かめる（会議の中身は送らない）
                try:
                    return self._json({"ok": True, "text": mn.llm.test(s)[:100]})
                except Exception as e:
                    return self._json({"ok": False, "error": str(e)[:300]})
            if op == "engram" and not s.get("engram_on", True):
                return self._json(False)
            if op == "state":
                return self._json(hub.state())
            if op == "open":
                return self._json(hub.open(str(body.get("id", ""))))
            if op == "new":
                return self._json(hub.new())
            if op == "sound":  # この回の拾う音を選ぶ
                return self._json(hub.set_sound(str(body.get("value", ""))))
            if op == "clear":
                if not hub.recorder:
                    hub.meeting = None
                    hub.sound_pending = None
                return self._json(None)
            if op == "start":
                hub.start(str(body.get("mic", "")), bool(body.get("diarize", True)))
                return self._json(True)
            if op == "stop":
                hub.stop()
                return self._json(True)
            if op == "title":
                if m:
                    m.meta["title"] = str(body.get("title", "")).strip()
                    m.save_meta()
                return self._json(True)
            if op == "ask":
                hub.ask(str(body.get("q", "")).strip(), bool(body.get("past")),
                        bool(body.get("web")), bool(body.get("connectors")))
                return self._json(True)
            if op == "hide":  # 発言を押してグレーにする（Claude と engram に渡さない）／もう一度押して戻す
                if not m or m.folder.name != body.get("id"):
                    return self._json(None)
                lines = m.set_hidden(int(body.get("i", -1)), bool(body.get("on")), body.get("at"))
                if lines is None:  # 画面と番号が食い違っている。別の発言を伏せない
                    return self._json(None)
                m.write_readable()
                return self._json(lines)
            if op == "speaker":  # 話者を手で直す
                if not m or m.folder.name != body.get("id"):
                    return self._json(None)
                lines = m.set_speaker(int(body.get("i", -1)), str(body.get("name", "")).strip()[:30], bool(body.get("all")))
                m.write_readable()
                return self._json(lines)
            if op == "speaker_rows":  # 話者の設定の表
                if not m:
                    return self._json(None)
                return self._json({"rows": m.speaker_rows()})
            if op == "speaker_names":  # この回の「話者A＝〇〇さん」を、保存のボタンでまとめて決める
                names = body.get("names")
                if not m or m.folder.name != body.get("id") or not isinstance(names, dict):
                    return self._json(None)
                lines = None
                for code, name in names.items():
                    lines = m.set_speaker_name(str(code), str(name).strip()[:30])
                m.write_readable()
                return self._json({"lines": m.lines() if lines is None else lines, "rows": m.speaker_rows()})
            if op == "pii":  # 個人情報らしい発言を、このPCの中だけで判定して伏せる
                if not m or hub.recorder:
                    hub.status("記録を止めてから押してください")
                    return self._json(None)
                lines, n, counts = m.mark_pii()
                m.write_readable()
                detail = "・".join(f"{k} {v}" for k, v in counts.items())
                hub.status(f"{n}件をグレーにしました（{detail}）。違うものは押すと戻せます" if n else
                           "新しく見つかった個人情報らしい発言はありません。「〜さん」などが付かない名前は見つけられません")
                return self._json(lines)
            if op == "denoise":
                if s.get("denoise_on") and mn.denoise_available():
                    hub.denoise(bool(s.get("diarize", True)))
                return self._json(True)
            if op == "fix":
                hub.fix()
                return self._json(True)
            if op == "engram":
                hub.engram()
                return self._json(True)
            if op == "save_chat":
                if not m or not m.chats():
                    hub.status("保存する会話がまだありません")
                    return self._json(None)
                p = m.write_chat_file()
                subprocess.Popen(["explorer", "/select,", str(p)], creationflags=NO_WINDOW)
                hub.status(f"会話を保存しました: {p.name}")
                return self._json(str(p))
            if op == "attach":
                threading.Thread(target=hub.attach, daemon=True).start()
                return self._json(True)
            if op == "series":  # 会議名を切り替える
                hub.set_series(str(body.get("path", "")))
                return self._json(hub.state())
            if op == "series_new":  # 保存先の直下に会議名のフォルダを作って、それに切り替える
                p = None if hub.recorder else mn.create_series(str(body.get("name", "")))
                if p:
                    hub.set_series(str(p))
                    hub.status(f"会議名「{p.name}」を作りました。新しい回はここに保存します: {p}")
                return self._json(hub.state())
            if op == "series_add":  # 好きな場所のフォルダを会議名として登録する
                d = None if hub.recorder else pick_files(folder=True, title="会議名にするフォルダを選ぶ")
                if d:
                    p = mn.register_series(Path(d))
                    hub.set_series(str(p))
                    hub.status(f"会議名「{p.name}」を登録しました。新しい回はここに保存します: {p}")
                return self._json(hub.state())
            if op == "choose_data":
                hub.choose_data()
                return self._json(hub.state())
            if op == "open_folder":
                which = body.get("which")
                if which == "shots":
                    p = m.folder / mn.SHOT_DIR if m else None
                    if p and p.exists():
                        os.startfile(p)
                    else:
                        hub.status("この回のスクリーンショットはまだありません")
                    return self._json(True)
                target = (m.doc_dir if which == "docs" and m else
                          m.folder if which == "meeting" and m else
                          hub.series if which == "series" else mn.DATA)
                if target.exists():
                    os.startfile(target)
                return self._json(True)
            if op == "pin":
                return self._json(set_topmost(bool(body.get("on"))))
            if op == "settings":
                with hub.settings_lock:  # 続けて変えても、前の変更を消さないように読み直して書く
                    s = mn.load_settings()
                    for k in ("theme", "limit", "diarize", "mic", "past", "pin", "split", "claude_on", "engram_on", "denoise_on", "rate", "chat_closed", "sound_default"):
                        if k in body:
                            s[k] = body[k]
                    mn.save_settings(s)
                return self._json(hub.settings())
            self.send_error(404)

    return H


def set_topmost(on):
    """題名が「Sidenote」の Edge の窓を、ほかの窓より手前に固定する（または外す）"""
    import ctypes
    from ctypes import wintypes
    user32 = ctypes.windll.user32
    found = []

    @ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.HWND, wintypes.LPARAM)
    def each(hwnd, _):
        buf = ctypes.create_unicode_buffer(256)
        user32.GetWindowTextW(hwnd, buf, 256)
        # InPrivate で開くと題名が「Sidenote - [InPrivate]」になる
        if buf.value.startswith("Sidenote") and user32.IsWindowVisible(hwnd):
            found.append(hwnd)
        return True

    user32.EnumWindows(each, 0)
    # HWND_TOPMOST=-1 / HWND_NOTOPMOST=-2。型を指定しないと -1 が32ビットの値で渡り、固定されなかった（実測）
    user32.SetWindowPos.argtypes = [wintypes.HWND, wintypes.HWND, ctypes.c_int, ctypes.c_int,
                                    ctypes.c_int, ctypes.c_int, ctypes.c_uint]
    for hwnd in found:  # 位置と大きさは変えず、手前に出すだけ
        user32.SetWindowPos(hwnd, wintypes.HWND(-1 if on else -2), 0, 0, 0, 0, 0x0001 | 0x0002 | 0x0010)
    return bool(found)


def open_window(url):
    """Chrome（無ければ Edge）の「アプリとして開く」で、画面の右端に縦長の窓を出す"""
    import ctypes
    from ctypes import wintypes
    rect = wintypes.RECT()
    ctypes.windll.user32.SystemParametersInfoW(0x30, 0, ctypes.byref(rect), 0)
    scale = ctypes.windll.user32.GetDpiForSystem() / 96
    w = 480
    h = int((rect.bottom - rect.top) / scale)
    x = int(rect.right / scale) - w
    # Sidenote の設定は Python 側（%APPDATA%\Sidenote\settings.json）に持つので、ブラウザには何も残らなくてよい
    exe, name = find_browser()
    if not exe:  # どちらも無いパソコン。窓の形は崩れるが、いつものブラウザで開ける
        webbrowser.open(url)
        return
    profile = PROFILE_ROOT / name  # いつものブラウザとは別の設定で開く。ログインや拡張機能に触らない
    profile.mkdir(parents=True, exist_ok=True)
    # Edge だけ InPrivate で開く。専用の設定フォルダを普通に使うと、Windows のアカウントで自動でサインインし、
    # 同期の案内が窓にかぶさった（実測）
    private = ["--inprivate"] if name == "edge" else []
    subprocess.Popen([exe, *private, f"--app={url}", f"--window-size={w},{h}", f"--window-position={x},0",
                      f"--user-data-dir={profile}", "--no-first-run", "--no-default-browser-check",
                      "--disable-sync", "--disable-features=Translate"], creationflags=NO_WINDOW)


def main():
    import ctypes
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except (OSError, AttributeError):
        pass
    # --background: Windows の起動時（スタートアップ）用。窓を出さずに常駐だけする
    background = "--background" in sys.argv[1:]
    # すでに動いていれば、その窓を開くだけにする（二重に記録しないため）
    def join_running():
        try:
            st = json.loads(STATE_FILE.read_text(encoding="utf-8"))
            LOCAL.open(f"http://127.0.0.1:{st['port']}/ping", timeout=1)
        except Exception:
            return False
        if not background:
            try:  # 開いている窓や隠している窓があれば、新しい窓を開かずに手前へ出す
                req = urllib.request.Request(f"http://127.0.0.1:{st['port']}/api/show_window?t={st['token']}", data=b"{}", method="POST")
                if json.loads(LOCAL.open(req, timeout=2).read()):
                    return True
            except Exception:
                pass
            open_window(f"http://127.0.0.1:{st['port']}/?t={st['token']}")
        return True

    if join_running():
        return
    # スタートアップと手動の起動が同時でも、1つだけが動く。設定の場所ごとに分ける
    name = "Local\\Sidenote_" + hashlib.sha1(str(STATE_FILE).lower().encode("utf-8")).hexdigest()[:16]
    mutex = ctypes.windll.kernel32.CreateMutexW(None, False, name)
    if ctypes.windll.kernel32.GetLastError() == 183:  # もう1つが起動の途中。準備ができたらそちらへ合流する
        for _ in range(20):
            time.sleep(0.5)
            if join_running():
                return
        return
    hub = Hub()
    token = secrets.token_urlsafe(16)
    srv = http.server.ThreadingHTTPServer(("127.0.0.1", 0), make_handler(hub, token))
    srv.daemon_threads = True
    port = srv.server_address[1]
    STATE_FILE.parent.mkdir(parents=True, exist_ok=True)
    STATE_FILE.write_text(json.dumps({"port": port, "token": token}), encoding="utf-8")
    threading.Thread(target=srv.serve_forever, daemon=True).start()
    url = f"http://127.0.0.1:{port}/?t={token}"
    if not background and not os.environ.get("MEETING_NOTES_NO_WINDOW"):
        open_window(url)

    # 常駐する。窓を閉じても終わらず、記録中なら裏で記録を続ける。
    # 終わるのは通知領域のアイコンの「終了」だけ。記録中なら停止して書き終えてから終わる（engram へは送らない）
    def quit_app():
        if hub.finishing:
            return
        hub.finishing = True
        if hub.recorder:
            hub.stop()
            threading.Timer(120, hub.quit.set).start()  # 部品が固まっても、終了できなくならないように
        else:
            hub.quit.set()

    from tray import Tray
    tray = Tray(ICON_FILE, "Sidenote", on_open=lambda: hub.focus_window() or open_window(url), on_quit=quit_app)
    hub.on_rec = tray.set_recording  # 記録中はアイコンを替えて、記録していることが分かるようにする
    tray.run(hub.quit)
    hub.show_window()  # 隠した窓を残したまま終わらない
    srv.shutdown()
    try:
        STATE_FILE.unlink()
    except OSError:
        pass


if __name__ == "__main__":
    main()
