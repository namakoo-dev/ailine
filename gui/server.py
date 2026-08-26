"""ailine の GUI ── **薄い殻**。判定は 1 ビットも作らない。

★ 設計の縛り（これが守れなくなったら GUI ごと捨てる）:
  画面に出る ✓ / △ / ⚠ は、`ailine run --json` が返す `verdict` を**そのまま映す**。
  ここで advisories を数え直したり、postcondition から印を導いたりしない ──
  それは 2 つ目の実装で、この repo が 2026-08 の盲検で 4 回踏んだ欠陥
  （「検算の分母が、疑うべき対象と同じ所から作られる」）を自分で新造することになる。
  ★ 決めるのは 1 箇所（ailine 本体）・映すのは何箇所でも。

★ 依存を足さない: Python 標準ライブラリだけ（scripts/ci_parity.py の番人が守る）。
★ localhost のみに bind する（外に開かない）。
"""
from __future__ import annotations

import json
import subprocess
import sys
import threading
import webbrowser
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import parse_qs, urlparse

HERE = Path(__file__).resolve().parent
REPO = HERE.parent
PAGE = HERE / "index.html"
RUN_TIMEOUT = 300


def _ailine(args: list) -> tuple:
    """ailine を子プロセスで動かす。戻り値 (returncode, stdout, json or None)。

    ★ import して呼ばない: 画面の都合で本体の内部状態を触らないため
      （殻はあくまで外から使う人と同じ入口を叩く）。
    """
    cmd = [sys.executable, "-m", "ailine"] + args
    # ★★ 実測（2026-08-26）: これを付けないと、子プロセスは **site-packages に入っている
    #   古い版**を import する（cwd に `ailine/` が無いので `-m` が拾えない）。
    #   盲検 2 回目で検分者が古いタグを測ってしまったのと同じ形の事故を、
    #   今度は自分の画面が起こすところだった ── **殻は必ずこの repo の本体を叩く**。
    import os as _os
    env = dict(_os.environ)
    src = str(REPO / "src")
    env["PYTHONPATH"] = src + _os.pathsep + env.get("PYTHONPATH", "")
    env.setdefault("PYTHONUTF8", "1")
    proc = subprocess.run(cmd, cwd=str(REPO), capture_output=True, text=True,
                           encoding="utf-8", errors="replace", timeout=RUN_TIMEOUT, env=env)
    out = (proc.stdout or "") + (proc.stderr or "")
    payload = None
    for line in (proc.stdout or "").splitlines():
        s = line.strip()
        if s.startswith("{") and s.endswith("}"):
            try:
                payload = json.loads(s)
            except json.JSONDecodeError:
                pass
    return proc.returncode, out, payload


def _default_dir() -> Path:
    """最初に開くフォルダ。同梱のサンプルがあればそこ、無ければ repo 直下。"""
    for cand in (REPO / "demo_gui", REPO / "demo"):
        if cand.is_dir():
            return cand
    return REPO


class Handler(BaseHTTPRequestHandler):
    def log_message(self, *a):        # 端末を汚さない
        pass

    def _send(self, code: int, body: bytes, ctype: str) -> None:
        self.send_response(code)
        self.send_header("Content-Type", ctype)
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def _json(self, code: int, obj: dict) -> None:
        self._send(code, json.dumps(obj, ensure_ascii=False).encode("utf-8"),
                    "application/json; charset=utf-8")

    def do_GET(self):
        u = urlparse(self.path)
        if u.path in ("/", "/index.html"):
            self._send(200, PAGE.read_bytes(), "text/html; charset=utf-8")
            return
        if u.path == "/api/files":
            q = parse_qs(u.query)
            folder = Path((q.get("dir") or [str(_default_dir())])[0]).expanduser()
            try:
                folder = folder.resolve()
                items = sorted((p for p in folder.iterdir()
                                 if p.is_file() and p.suffix.lower() == ".xlsx"
                                 and not p.name.startswith("~$")),
                                key=lambda q: q.name)
            except OSError as e:
                self._json(200, {"dir": str(folder), "files": [], "error": str(e)})
                return
            # ★ 画面側でパスを**組み立てさせない**（区切り文字の二重化を実測で踏んだ）。
            #   完全なパスはここで作って、そのまま返す ── 組み立てが 1 箇所なら、ずれない。
            # ★ ailine 自身が作った下書き（.out.xlsx）は**隠さず、そう見せる**
            #   （一覧から消すと「無い」と読まれる ── 出ないことは信号でない）。
            self._json(200, {"dir": str(folder),
                              "files": [{"name": p.name, "path": str(p),
                                          "draft": p.name.endswith(".out.xlsx")}
                                         for p in items]})
            return
        self._json(404, {"error": "not found"})

    def do_POST(self):
        u = urlparse(self.path)
        try:
            length = int(self.headers.get("Content-Length") or 0)
            req = json.loads(self.rfile.read(length) or b"{}")
        except Exception as e:
            self._json(400, {"error": f"読めない要求: {e}"})
            return
        book = str(Path(req.get("book") or "").expanduser())
        if not book:
            self._json(400, {"error": "ファイルが選ばれていません"})
            return
        try:
            if u.path == "/api/run":
                args = ["run", book, req.get("task") or "", "--json"]
                if req.get("copy"):
                    args.append("--copy")
                rc, out, payload = _ailine(args)
            elif u.path == "/api/undo":
                rc, out, payload = _ailine(["undo", book])
            elif u.path == "/api/history":
                rc, out, payload = _ailine(["undo", book, "--list"])
            else:
                self._json(404, {"error": "not found"})
                return
        except subprocess.TimeoutExpired:
            self._json(200, {"rc": None, "text": "時間内に終わりませんでした", "json": None})
            return
        # ★ verdict は payload をそのまま渡す（ここで作らない・書き換えない）
        self._json(200, {"rc": rc, "text": out, "json": payload})


def main() -> int:
    port = 8760
    srv = ThreadingHTTPServer(("127.0.0.1", port), Handler)
    url = f"http://127.0.0.1:{port}/"
    print(f"ailine GUI: {url}  （終了は Ctrl+C）")
    threading.Timer(0.6, lambda: webbrowser.open(url)).start()
    try:
        srv.serve_forever()
    except KeyboardInterrupt:
        print("\n終了します")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
