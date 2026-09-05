"""local_only — 「外に 1 バイトも出ない」を、宣言でなく**機械**にする層。

★★ 出所（2026-09-05・盲検の査定）: この道具の売り文句はこう書いてある ──

    「外に 1 バイトも出ない。ollama（ローカル）と LibreOffice だけで完結します」

  だが実体はこうだった:

    OLLAMA = os.environ.get("OLLAMA_HOST", "http://localhost:11434")

  **既定がローカル**であって、外に出ないわけではない。`OLLAMA_HOST` を任意の URL に
  向ければ、依頼文・見出し・セルの実値がそこへ POST される。拒む assert も、
  それを確かめる試験も**無かった**。

★ 想定している使い手は「社外にデータを出せない人」だ。**その人が唯一気にする一点**が、
  この repo で唯一「宣言だけで機械が守っていない」契約になっていた ──
  「指示は意図、保証は機械」を、いちばん効かせるべき所で効かせていなかった。

★ 決裁（Namakoo 2026-09-05）: **既定で拒み、旗で明示許可**。
  社内の ollama サーバを使いたい人を締め出さないが、**黙って外へは出さない**。

★ 移植可能性（tests/test_line_budget.py が機械で守る）: ailine を import しない。
"""
from __future__ import annotations

from urllib.parse import urlsplit

#: この道具が「手元」とみなすホスト名（★ ここが唯一の名簿）
LOCAL_HOSTS = frozenset({"localhost", "127.0.0.1", "::1", "0.0.0.0"})

#: 明示許可の旗の名前（画面・文書・実装で同じ字を使う）
ALLOW_FLAG = "--allow-remote-model"


def host_of(url: str) -> str:
    """URL からホスト名だけを取り出す（ポートと角括弧は落とす）。

    ★ 素朴に文字列を含むかで見ると `http://localhost.evil.example/` を通してしまう。
      ホストの**完全一致**で見る。
    """
    if not url:
        return ""
    text = url if "//" in url else "//" + url
    try:
        host = urlsplit(text).hostname or ""
    except ValueError:
        return ""
    return host.strip("[]").lower()


def host_is_local(url: str) -> bool:
    """その URL が手元を指しているか。★ 判断はここ 1 箇所だけが持つ。"""
    return host_of(url) in LOCAL_HOSTS


def render_remote_refusal(url: str) -> list:
    """外を向いていた時に出す断り（★ 何が起きるかを具体的に言う）。"""
    return [
        f"？ モデルの宛先が手元ではありません: {url}",
        "  この道具は既定で外に出しません ── 依頼文・見出し・セルの中身が"
        "その宛先へ送られます。",
        f"  OLLAMA_HOST を手元（localhost）に戻すか、"
        f"承知のうえで送るなら {ALLOW_FLAG} を付けてください。",
    ]


def render_remote_notice(url: str) -> str:
    """明示許可された回に、**黙らずに**出す 1 行。"""
    return (f"★ 外部の ollama に送ります: {url}"
            f"（{ALLOW_FLAG} が指定されました。依頼文と表の中身がこの宛先へ出ます）")
