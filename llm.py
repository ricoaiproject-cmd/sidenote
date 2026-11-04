"""AI の接続先を切り替える。Claude Code 以外（API キー・ローカル LLM）への呼び出しをここにまとめる。

API キーは settings.json に書かず、Windows の資格情報マネージャーに保存する。画面へは「保存済みか」だけを返す。
"""
import http.client
import ipaddress
import json
import re
import socket
import threading
import urllib.error
import urllib.parse
import urllib.request

CRED_PREFIX = "Sidenote/"
ANTHROPIC_DEFAULT_MODEL = "claude-opus-5-5"
# 安全判定で断られたときに別のモデルへ回す指定（fallbacks）を受け付けるモデル
ANTHROPIC_FALLBACK_MODELS = {"claude-fable-5-1", "claude-opus-5-5", "claude-opus-5", "claude-sonnet-5-5"}
MAX_OUTPUT = 16000
DEFAULT_LOCAL_CTX = 32768
LOCAL_RESERVE = 4096  # 答えのために空けておく分。日本語は1字がおよそ1トークンとして見積もる
CUT_MARK = "\n\n（……長いので、ここから先の途中を省略しました……）\n\n"
CUT_NOTE = "\n\n（記録が長いので途中を省いて読みました。設定の local_max_ctx を増やすと、もっと読めます）"

# 文章を作らないモデルを、モデルの一覧から外す
NON_CHAT = re.compile(r"embed|whisper|tts|image|dall-e|audio|realtime|moderation|transcribe|sora|babbage|davinci|aqa|veo|imagen", re.I)

PROVIDERS = {
    "none": {"label": "使わない"},
    "claude_code": {"label": "Claude Code（ログインで使う）", "name": "Claude Code"},
    "anthropic": {"label": "Claude API（API キー）", "name": "Claude API", "key": True},
    "openai": {"label": "OpenAI API（API キー）", "name": "OpenAI API", "key": True, "url": "https://api.openai.com/v1"},
    "gemini": {"label": "Google Gemini API（API キー）", "name": "Gemini API", "key": True,
               "url": "https://generativelanguage.googleapis.com/v1beta/openai"},
    "custom": {"label": "その他の API（OpenAI 互換）", "name": "API", "key": "optional", "url": "", "url_edit": True},
    "ollama": {"label": "ローカル LLM：Ollama", "name": "Ollama", "url": "http://127.0.0.1:11434", "url_edit": True},
    "lmstudio": {"label": "ローカル LLM：LM Studio", "name": "LM Studio", "url": "http://127.0.0.1:1234/v1", "url_edit": True},
}


class AIError(RuntimeError):
    pass


class Cancelled(RuntimeError):
    """使う人が中止ボタンを押した"""


class Job:
    """中止できる頼みごと。中止すると、動いている Claude Code やローカル LLM への接続を切る。
    外部の API は、送ってしまった頼みを取り消せないので、答えを捨てるだけになる"""

    def __init__(self):
        self.cancelled = threading.Event()
        self._stop = None
        self._lock = threading.Lock()

    def on_cancel(self, stop):
        """中止されたときに呼ぶ処理を登録する。すでに中止されていれば、すぐ呼ぶ"""
        with self._lock:
            self._stop = stop
            done = self.cancelled.is_set()
        if done:
            self._run(stop)

    def cancel(self):
        with self._lock:
            self.cancelled.set()
            stop = self._stop
        if stop:
            self._run(stop)

    @staticmethod
    def _run(stop):
        try:
            stop()
        except OSError:
            pass

    def check(self):
        if self.cancelled.is_set():
            raise Cancelled("中止しました")


def provider(s):
    p = s.get("ai_provider")
    if p in PROVIDERS:
        return p
    return "none" if s.get("claude_on", True) is False else "claude_code"  # 接続先を選ぶ前の設定を引き継ぐ


def config(s, p=None):
    return (s.get("ai") or {}).get(p or provider(s)) or {}


def base_url(s, p=None):
    p = p or provider(s)
    info = PROVIDERS[p]
    url = config(s, p).get("url") if info.get("url_edit") else None
    url = (url or info.get("url") or "").strip().rstrip("/")
    return re.sub(r"/chat/completions$", "", url)


def is_loopback(url):
    host = urllib.parse.urlparse(url).hostname or ""
    if host == "localhost":
        return True
    try:
        return ipaddress.ip_address(host).is_loopback
    except ValueError:
        return False


def stays_on_pc(s):
    """答えを作る先がこのパソコンの中か"""
    p = provider(s)
    return p in ("ollama", "lmstudio", "custom") and is_loopback(base_url(s, p))


def check_url(url):
    u = urllib.parse.urlparse(url)
    if u.scheme not in ("http", "https") or not u.hostname:
        raise AIError("URL は http:// か https:// で始めてください")
    if u.scheme == "http" and not is_loopback(url):
        raise AIError("このパソコンの外へは https:// でしかつなげません（文字起こしと API キーが暗号化されずに流れるため）")


# ── API キー（Windows の資格情報マネージャー）────────────────────────
def _target(p):
    return CRED_PREFIX + p


def set_key(p, key):
    import win32cred
    win32cred.CredWrite({"Type": win32cred.CRED_TYPE_GENERIC, "TargetName": _target(p), "UserName": "Sidenote",
                         "CredentialBlob": key, "Persist": win32cred.CRED_PERSIST_LOCAL_MACHINE}, 0)


def get_key(p):
    try:
        import pywintypes
        import win32cred
    except ImportError:
        return ""
    try:
        c = win32cred.CredRead(_target(p), win32cred.CRED_TYPE_GENERIC, 0)
    except pywintypes.error:
        return ""
    return c["CredentialBlob"].decode("utf-16-le")


def del_key(p):
    import pywintypes
    import win32cred
    try:
        win32cred.CredDelete(_target(p), win32cred.CRED_TYPE_GENERIC, 0)
    except pywintypes.error:
        pass


# ── 画面に見せる状態 ───────────────────────────────────────────────
def status(s, claude_found):
    """接続先の状態。API キーの中身は決して入れない"""
    p = provider(s)
    info = PROVIDERS[p]
    has_key = bool(info.get("key")) and bool(get_key(p))
    model = config(s, p).get("model") or (ANTHROPIC_DEFAULT_MODEL if p == "anthropic" else "")
    url = base_url(s, p)
    need = ""
    if p == "claude_code" and not claude_found:
        need = "Claude Code が見つかりません。入れてログインし、窓を開き直すと使えます"
    elif info.get("key") is True and not has_key:
        need = "API キーを入れてください"
    elif "url" in info and not url:
        need = "URL を入れてください"
    elif p not in ("none", "claude_code") and not model:
        need = "モデルを選んでください"
    return {"provider": p, "name": info.get("name", ""), "ready": p != "none" and not need, "need": need,
            "model": model, "url": url, "has_key": has_key, "key": info.get("key", False),
            "url_edit": bool(info.get("url_edit")), "local": stays_on_pc(s),
            "providers": [{"id": k, "label": v["label"]} for k, v in PROVIDERS.items()]}


# ── 通信 ─────────────────────────────────────────────────────────
class _NoRedirect(urllib.request.HTTPRedirectHandler):
    def redirect_request(self, *a, **k):  # 転送先へ API キーを渡さない
        return None


def _http_local(url, data, headers, timeout, service, key, job):
    """このパソコンの中の窓口へ聞く。中止されたら接続を切る（ローカル LLM も計算をやめる）"""
    u = urllib.parse.urlparse(url)
    conn = http.client.HTTPConnection(u.hostname, u.port or 80, timeout=timeout)
    try:
        conn.connect()
        if job:
            job.on_cancel(lambda: conn.sock.shutdown(socket.SHUT_RDWR))
        conn.request("POST" if data is not None else "GET", u.path + (f"?{u.query}" if u.query else ""), data, headers)
        r = conn.getresponse()
        raw = r.read()
    except (OSError, http.client.HTTPException):
        if job:
            job.check()
        raise AIError(f"{service} に接続できません。起動しているか確かめてください") from None
    finally:
        conn.close()
    if job:
        job.check()
    if r.status != 200:
        raise AIError(_http_message(r.status, raw, key, service))
    try:
        return json.loads(raw.decode("utf-8"))
    except ValueError:
        raise AIError(f"{service} の答えの形が想定と違います。URL を確かめてください") from None


def _http(url, body=None, key="", timeout=30, service="接続先", job=None):
    headers = {"Content-Type": "application/json"}
    if key:
        headers["Authorization"] = f"Bearer {key}"
    data = json.dumps(body).encode("utf-8") if body is not None else None
    if is_loopback(url) and url.startswith("http://"):
        return _http_local(url, data, headers, timeout, service, key, job)
    req = urllib.request.Request(url, data=data, headers=headers)
    handlers = [_NoRedirect()]
    if is_loopback(url):  # このパソコンの中の通信を、プロキシへ回さない
        handlers.append(urllib.request.ProxyHandler({}))
    try:
        with urllib.request.build_opener(*handlers).open(req, timeout=timeout) as r:
            return json.loads(r.read().decode("utf-8"))
    except urllib.error.HTTPError as e:
        raise AIError(_http_message(e.code, e.read(), key, service)) from None
    except (urllib.error.URLError, TimeoutError, OSError) as e:
        if is_loopback(url):
            raise AIError(f"{service} に接続できません。起動しているか確かめてください") from None
        raise AIError(f"{service} に接続できません: {getattr(e, 'reason', e)}") from None
    except ValueError:
        raise AIError(f"{service} の答えの形が想定と違います。URL を確かめてください") from None


def _http_message(code, raw, key, service):
    if code in (401, 403):
        return "API キーが正しくないか、このモデルを使う権限がありません"
    if code == 429:
        return "利用上限か料金の枠に達しました。しばらく待つか、契約を確かめてください"
    try:
        detail = raw.decode("utf-8", "replace")
        o = json.loads(detail)
        err = o.get("error")
        detail = (err.get("message") if isinstance(err, dict) else err) or detail
    except (ValueError, AttributeError):
        detail = ""
    if key:
        detail = str(detail).replace(key, "（API キー）")
    return f"{service} がエラーを返しました（{code}）: {str(detail)[:200]}"


def _clean(text):
    """ローカル LLM の癖を直す。考えている途中の文を外し、<0xE3><0x80><0x80> のようなバイト表記を文字に戻す"""
    text = re.sub(r"<think>.*?</think>", "", text or "", flags=re.S)
    text = re.sub(r"(?:<0x[0-9A-Fa-f]{2}>)+",
                  lambda m: bytes.fromhex(m.group(0).replace("<0x", "").replace(">", "")).decode("utf-8", "replace"), text)
    return text.strip()


def _anthropic_client(p, timeout):
    try:
        import anthropic
    except ImportError:
        raise AIError("Claude API の部品が入っていません。setup.bat をもう一度実行してください") from None
    return anthropic, anthropic.Anthropic(api_key=get_key(p), timeout=timeout)


def _anthropic_error(anthropic, e):
    if isinstance(e, anthropic.AuthenticationError) or isinstance(e, anthropic.PermissionDeniedError):
        return "API キーが正しくないか、このモデルを使う権限がありません"
    if isinstance(e, anthropic.NotFoundError):
        return "モデルが見つかりません。設定でモデルを選び直してください"
    if isinstance(e, anthropic.RateLimitError):
        return "利用上限か料金の枠に達しました。しばらく待つか、契約を確かめてください"
    if isinstance(e, anthropic.APIStatusError):
        return f"Claude API がエラーを返しました（{e.status_code}）"
    return "Claude API に接続できません"


def _complete_anthropic(s, model, system, text, timeout):
    anthropic, client = _anthropic_client("anthropic", timeout)
    kw = dict(model=model, max_tokens=MAX_OUTPUT, system=system, messages=[{"role": "user", "content": text}])
    try:
        if model in ANTHROPIC_FALLBACK_MODELS:
            m = client.beta.messages.create(betas=["server-side-fallback-2026-07-01"], fallbacks="default", **kw)
        else:
            m = client.messages.create(**kw)
    except anthropic.APIError as e:
        raise AIError(_anthropic_error(anthropic, e)) from None
    if m.stop_reason == "refusal":
        raise AIError("Claude の安全判定で、この内容には答えられませんでした")
    out = "".join(b.text for b in m.content if b.type == "text").strip()
    if not out:
        raise AIError("答えが空でした")
    if m.stop_reason == "max_tokens":
        out += "\n\n（答えが長すぎて途中で切れました）"
    return out


def _complete_openai(s, p, model, system, text, timeout, job=None):
    url = base_url(s, p)
    check_url(url)
    o = _http(f"{url}/chat/completions",
              {"model": model, "messages": [{"role": "system", "content": system}, {"role": "user", "content": text}]},
              get_key(p), timeout, PROVIDERS[p]["name"], job)
    try:
        choice = o["choices"][0]
        out = _clean(choice["message"]["content"])
    except (KeyError, IndexError, TypeError):
        raise AIError(f"{PROVIDERS[p]['name']} の答えの形が想定と違います") from None
    if not out:
        raise AIError("答えが空でした")
    if choice.get("finish_reason") == "length":
        out += "\n\n（答えが長すぎて途中で切れました）"
    return out


def _complete_ollama(s, model, system, text, timeout, ctx, job=None):
    url = base_url(s, "ollama")
    check_url(url)
    # num_ctx を渡さないと、既定の小さな枠を超えた分が黙って捨てられる
    num_ctx = max(4096, min(ctx, len(system) + len(text) + LOCAL_RESERVE))
    o = _http(f"{url}/api/chat",
              {"model": model, "stream": False, "options": {"num_ctx": num_ctx},
               "messages": [{"role": "system", "content": system}, {"role": "user", "content": text}]},
              "", timeout, "Ollama", job)
    out = _clean((o.get("message") or {}).get("content"))
    if not out:
        raise AIError("答えが空でした")
    if o.get("done_reason") == "length":
        out += "\n\n（答えが長すぎて途中で切れました）"
    return out


def _fit(text, limit):
    """長すぎる入力の途中を省く。頭（会議の情報）と終わり（いま答える質問）は残す"""
    if len(text) <= limit:
        return text, False
    head = limit // 3
    return text[:head] + CUT_MARK + text[-(limit - head - len(CUT_MARK)):], True


def complete(s, system, text, timeout=None, may_cut=True, job=None):
    """Claude Code 以外の接続先に1回聞いて、答えの文を返す。失敗したら AIError を、中止されたら Cancelled を投げる"""
    p = provider(s)
    st_model = config(s, p).get("model") or (ANTHROPIC_DEFAULT_MODEL if p == "anthropic" else "")
    if not st_model:
        raise AIError("モデルを選んでください（右上の歯車）")
    if PROVIDERS[p].get("key") is True and not get_key(p):
        raise AIError("API キーが入っていません（右上の歯車）")
    local = p in ("ollama", "lmstudio")
    cut = False
    ctx = int(s.get("local_max_ctx") or DEFAULT_LOCAL_CTX)
    if local:  # ローカル LLM は読める長さが短く、計算にも時間がかかる
        limit = max(1000, ctx - LOCAL_RESERVE - len(system))
        if len(text) > limit and not may_cut:
            raise AIError("ローカル LLM には長すぎます。資料を減らすか、設定の local_max_ctx を増やしてください")
        text, cut = _fit(text, limit)
    timeout = timeout or (1800 if local else 300)
    if p == "anthropic":
        out = _complete_anthropic(s, st_model, system, text, timeout)
    elif p == "ollama":
        out = _complete_ollama(s, st_model, system, text, timeout, ctx, job)
    else:
        out = _complete_openai(s, p, st_model, system, text, timeout, job)
    if job:
        job.check()  # 外部の API は途中で止められないので、中止されていたら答えを捨てる
    return out + (CUT_NOTE if cut else "")


def models(s):
    """選べるモデルの名前を返す。失敗したら AIError を投げる（API キーやサーバーの確かめにもなる）"""
    p = provider(s)
    if p in ("none", "claude_code"):
        return []
    info = PROVIDERS[p]
    if info.get("key") is True and not get_key(p):
        raise AIError("API キーを入れると、モデルの一覧が出ます")
    if p == "anthropic":
        anthropic, client = _anthropic_client(p, 30)
        try:
            return [m.id for m in client.models.list(limit=100)]
        except anthropic.APIError as e:
            raise AIError(_anthropic_error(anthropic, e)) from None
    url = base_url(s, p)
    if not url:
        raise AIError("URL を入れてください")
    check_url(url)
    if p == "ollama":
        o = _http(f"{url}/api/tags", service="Ollama", timeout=10)
        return [m["name"] for m in o.get("models", []) if "embed" not in m["name"]]
    o = _http(f"{url}/models", key=get_key(p), service=info["name"], timeout=15)
    ids = [re.sub(r"^models/", "", m["id"]) for m in o.get("data", []) if m.get("id")]
    return sorted(i for i in ids if not NON_CHAT.search(i))


def test(s):
    """短い文だけを送って、つながるか確かめる。会議の中身は送らない"""
    return complete(s, "短く答えてください。", "「接続できました」とだけ答えてください。", timeout=None if provider(s) in ("ollama", "lmstudio") else 60)
