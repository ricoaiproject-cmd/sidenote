"""パソコンから出ている音（Web 会議の相手の声など）を受け取る（Windows の WASAPI ループバック）。

ループバックとは、スピーカーやイヤホンへ出ていく音を、そのまま受け取る Windows の仕組みのこと。
仮想オーディオのドライバーを入れなくて済む。音が何も鳴っていないあいだはデータが届かないので、足りない分は無音で埋める。
受け取った音は ffmpeg で 16kHz・モノラルに変える。
"""
import subprocess
import threading
import time

CHUNK_RATE = 32000  # 16kHz・16ビット・モノラルの1秒分のバイト数
PRIME = CHUNK_RATE // 5  # 音が届き始めたら 0.2 秒ためてから使う（細かく途切れないように）
MAX_BUFFER = CHUNK_RATE  # 1秒を超えてたまったら、遅れないように古い分を捨てる
KEEP_AFTER_TRIM = CHUNK_RATE * 3 // 10
CHECK_DEVICE_SEC = 5  # 既定の出力先（イヤホンへの切り替えなど）を確かめる間隔


def available():
    try:
        import pyaudiowpatch  # noqa: F401
        return True
    except ImportError:
        return False


class PcAudio:
    def __init__(self, ffmpeg, on_error=lambda text: None):
        self.ffmpeg, self.on_error = ffmpeg, on_error
        self.buf = bytearray()
        self.lock = threading.Lock()
        self.primed = False
        self.stopping = threading.Event()
        self.received = 0

    def start(self):
        threading.Thread(target=self._run, daemon=True).start()

    def stop(self):
        self.stopping.set()

    def read(self, n):
        """n バイトを返す。音が届いていない分は無音"""
        with self.lock:
            if len(self.buf) > MAX_BUFFER:  # 録音の時計とずれてたまった分を捨てる
                del self.buf[:len(self.buf) - KEEP_AFTER_TRIM]
            if not self.primed and len(self.buf) >= PRIME:
                self.primed = True
            if not self.primed:
                return bytes(n)
            out = bytes(self.buf[:n])
            del self.buf[:n]
            if len(out) < n:
                self.primed = False
                out += bytes(n - len(out))
            return out

    def _default_device(self, pa):
        return pa.get_default_wasapi_loopback()

    def _open(self, pa, dev):
        """出力先 dev の音を受け取り始める。返り値: (音の入口, 変換の ffmpeg)"""
        rate, channels = int(dev["defaultSampleRate"]), int(dev["maxInputChannels"])
        ff = subprocess.Popen(
            [self.ffmpeg, "-hide_banner", "-loglevel", "error", "-fflags", "nobuffer", "-probesize", "32",
             "-analyzeduration", "0", "-f", "s16le", "-ar", str(rate), "-ac", str(channels), "-i", "-",
             "-ac", "1", "-ar", "16000", "-flush_packets", "1", "-f", "s16le", "-"],
            stdin=subprocess.PIPE, stdout=subprocess.PIPE, stderr=subprocess.DEVNULL, creationflags=subprocess.CREATE_NO_WINDOW)

        def collect():
            while True:
                b = ff.stdout.read(1600)  # 0.05 秒分
                if not b:
                    return
                with self.lock:
                    self.buf += b
                    self.received += len(b)

        threading.Thread(target=collect, daemon=True).start()
        import pyaudiowpatch as pa_mod

        def callback(data, frames, info, status):
            try:
                ff.stdin.write(data)
                ff.stdin.flush()
            except OSError:
                return (None, pa_mod.paComplete)
            return (None, pa_mod.paContinue)

        stream = pa.open(format=pa_mod.paInt16, channels=channels, rate=rate, input=True, input_device_index=dev["index"],
                         frames_per_buffer=rate // 20, stream_callback=callback)
        return stream, ff

    @staticmethod
    def _close(stream, ff):
        try:
            stream.stop_stream()
            stream.close()
        except OSError:
            pass
        try:
            ff.stdin.close()
        except OSError:
            pass
        ff.terminate()

    def _run(self):
        try:
            import pyaudiowpatch
        except ImportError:
            self.on_error("パソコンの音を拾う部品がありません。setup.bat をもう一度実行してください（マイクの音は録れています）")
            return
        stream = ff = None
        pa = pyaudiowpatch.PyAudio()
        try:
            dev = self._default_device(pa)
            stream, ff = self._open(pa, dev)
            name = dev["name"]
            while not self.stopping.wait(CHECK_DEVICE_SEC):
                # 会議の途中でイヤホンへ切り替えても追いかける。機器の一覧は作り直さないと新しくならない
                probe = pyaudiowpatch.PyAudio()
                try:
                    now = self._default_device(probe)
                finally:
                    probe.terminate()
                if now["name"] != name:
                    self._close(stream, ff)
                    pa.terminate()
                    pa = pyaudiowpatch.PyAudio()
                    dev = self._default_device(pa)
                    stream, ff = self._open(pa, dev)
                    name = dev["name"]
        except Exception as e:  # 拾えなくても、マイクの録音は止めない
            self.on_error(f"パソコンの音を拾えませんでした: {str(e)[:150]}")
        finally:
            if stream:
                self._close(stream, ff)
            pa.terminate()
            time.sleep(0)
