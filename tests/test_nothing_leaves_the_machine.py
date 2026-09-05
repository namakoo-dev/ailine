"""「外に 1 バイトも出ない」を、宣言でなく機械が守ること（2026-09-05）。

★★ 出所（盲検の査定・所見③）: この道具は README で 3 箇所こう言い切っている ──

    「外に 1 バイトも出ない」「製品の実行時に外部 API は 1 つも呼びません」
    「情報の流出 → 外部 API を呼ばない」

  実体は `OLLAMA = os.environ.get("OLLAMA_HOST", "http://localhost:11434")` で、
  **既定がローカル**なだけだった。外部 URL に向ければ依頼文・見出し・セルの中身が
  そこへ POST される。拒む assert も、それを確かめる試験も無かった。

  ★ 想定ユーザは「社外にデータを出せない人」── **その人が唯一気にする一点**が、
    この repo で唯一「宣言だけで機械が守っていない」契約になっていた。
    「指示は意図、保証は機械」を、いちばん効かせるべき所で効かせていなかった。

★ 決裁（Namakoo）: 既定で拒み、旗で明示許可。社内 ollama の使い手は締め出さない。
★ 叩く場所は 4 箇所（chat 2・doctor 2）── 各所に門を置けば必ず片方が漏れるので、
  **URL を作れるのは ollama_url ただ 1 本**にした。④ がそれを機械で縛る。
"""
import os
import re
import subprocess
import sys
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "src"))

import ailine  # noqa: E402
from ailine_core import local_only  # noqa: E402

ROOT = Path(ailine.__file__).resolve().parents[2]
NL = chr(10)


# --- ① 手元かどうかの判断（★ 素朴な部分一致で通る罠を含む）------------------

@pytest.mark.parametrize("url, local", [
    ("http://localhost:11434", True),
    ("http://127.0.0.1:11434", True),
    ("http://[::1]:11434", True),
    ("http://LOCALHOST:11434", True),          # ★ 大文字でも手元
    ("http://localhost.evil.example/", False),  # ★ 部分一致なら通ってしまう罠
    ("http://127.0.0.1.evil.example/", False),
    ("https://api.example.com", False),
    ("http://192.168.1.9:11434", False),        # ★ 社内でも「手元」ではない
    ("", False),
    ("not a url", False),
])
def test_only_this_machine_counts_as_local(url, local):
    assert local_only.host_is_local(url) is local


# --- ② 既定では外へ出さない（★ 実機・素の環境で確かめる）--------------------

def _run(args, env_extra):
    env = {**os.environ, "PYTHONPATH": str(ROOT / "src"), **env_extra}
    return subprocess.run([sys.executable, "-m", "ailine", *args],
                          capture_output=True, text=True, encoding="utf-8",
                          errors="replace", cwd=str(ROOT), env=env)


def test_a_remote_model_is_refused_by_default():
    got = _run(["doctor"], {"OLLAMA_HOST": "https://api.example.com"})
    assert got.returncode == ailine.EXIT_ENVIRONMENT, (got.returncode, got.stdout, got.stderr)
    said = got.stdout + got.stderr
    assert "手元ではありません" in said and "api.example.com" in said, said
    assert local_only.ALLOW_FLAG in said, "逃げ道を案内していない"


def test_the_flag_allows_it_but_does_not_stay_silent():
    """★ 許した回も黙らない ── 何がどこへ出るかを画面に出してから走る。"""
    got = _run(["doctor", local_only.ALLOW_FLAG], {"OLLAMA_HOST": "http://192.168.1.250:11434"})
    said = got.stdout + got.stderr
    assert "外部の ollama に送ります" in said and "192.168.1.250" in said, said


def test_the_default_host_is_this_machine():
    env = {k: v for k, v in os.environ.items() if k != "OLLAMA_HOST"}
    got = subprocess.run([sys.executable, "-c",
                          "import ailine,sys;print(ailine.OLLAMA)"],
                         capture_output=True, text=True, encoding="utf-8",
                         errors="replace", cwd=str(ROOT),
                         env={**env, "PYTHONPATH": str(ROOT / "src")})
    assert local_only.host_is_local(got.stdout.strip()), got.stdout


# --- ③ 門そのもの（純ロジック）----------------------------------------------

def test_the_gate_refuses_and_exits_with_the_environment_code(monkeypatch):
    monkeypatch.setattr(ailine, "OLLAMA", "https://api.example.com")
    monkeypatch.setattr(ailine, "ALLOW_REMOTE_MODEL", False)
    with pytest.raises(SystemExit) as e:
        ailine.ollama_url("/api/chat")
    assert e.value.code == ailine.EXIT_ENVIRONMENT


def test_the_gate_lets_the_local_host_through(monkeypatch):
    monkeypatch.setattr(ailine, "OLLAMA", "http://localhost:11434")
    monkeypatch.setattr(ailine, "ALLOW_REMOTE_MODEL", False)
    assert ailine.ollama_url("/api/tags") == "http://localhost:11434/api/tags"


def test_the_gate_obeys_the_explicit_permission(monkeypatch):
    monkeypatch.setattr(ailine, "OLLAMA", "http://192.168.1.250:11434")
    monkeypatch.setattr(ailine, "ALLOW_REMOTE_MODEL", True)
    assert ailine.ollama_url("/api/chat").startswith("http://192.168.1.250:11434")


# --- ④ 片配線の番人（★ 新しい呼び出しが門を迂回していないこと）--------------

def test_the_url_can_only_be_built_in_one_place():
    """★ 叩く場所は 4 箇所ある。門を迂回する 5 箇所目が生えたらここが赤くなる。"""
    src = Path(ailine.__file__).read_text(encoding="utf-8")
    # ★ 画面に宛先を出すだけの行（「ollama に繋がらない (…)」）は URL の組み立てではない。
    #   組み立ては「{OLLAMA} の直後に道が続く」形だけを数える。
    builds = [ln.strip() for ln in src.splitlines()
              if re.search(r"\{OLLAMA\}\s*(?:/|\{)", ln)]
    assert len(builds) == 1, f"OLLAMA から直接 URL を組んでいる箇所がある: {builds}"
    gate = src.split("def ollama_url")[1].split(NL + "def ")[0]
    assert "{OLLAMA}" in gate, "唯一の使用箇所が門の中に無い"
    assert "host_is_local" in gate and "ALLOW_REMOTE_MODEL" in gate


def test_every_command_that_talks_to_ollama_has_the_flag():
    """★ 旗の定義は 1 箇所（_add_allow_remote）── run と doctor の両方に届くこと。"""
    src = Path(ailine.__file__).read_text(encoding="utf-8")
    assert src.count("add_argument(local_only.ALLOW_FLAG") == 1, "旗の定義が 2 箇所ある"
    parser = src.split("def build_parser")[1].split(NL + "def ")[0]
    assert parser.count("_add_allow_remote(") == 2, "run / doctor の片方に旗が無い"
    for cmd in ("run", "doctor"):
        out = subprocess.run([sys.executable, "-m", "ailine", cmd, "--help"],
                             capture_output=True, text=True, encoding="utf-8",
                             errors="replace", cwd=str(ROOT),
                             env={**os.environ, "PYTHONPATH": str(ROOT / "src")})
        assert local_only.ALLOW_FLAG in out.stdout, (cmd, out.stdout[-300:])


# --- ⑤ 文書が実体より強く言っていないこと -----------------------------------

def test_the_readme_does_not_promise_more_than_the_machine_gives():
    """★ 査定の所見そのもの: 表紙ほど言い切り、深い所ほど正直だった（順序が逆）。"""
    text = (ROOT / "README.md").read_text(encoding="utf-8")
    assert "外に 1 バイトも出ない" not in text or local_only.ALLOW_FLAG in text, (
        "『外に 1 バイトも出ない』と言うなら、"
        f"それを覆せる旗（{local_only.ALLOW_FLAG}）の存在も同じ文書に書くこと")
    assert "OLLAMA_HOST" in text, "挙動を変える環境変数が文書に無い"


# --- ⑥ 動く下限が 1 箇所から来ていること -------------------------------------

def test_the_python_floor_matches_the_declaration():
    """★ doctor が『3.10+』と言い続けていた（宣言は 3.12）── 同じ形の食い違い。"""
    text = (ROOT / "pyproject.toml").read_text(encoding="utf-8")
    m = re.search(r'requires-python\s*=\s*"[^"]*?(\d+)\.(\d+)', text)
    assert m
    assert ailine.MIN_PYTHON == (int(m.group(1)), int(m.group(2))), (
        f"doctor の下限 {ailine.MIN_PYTHON} と pyproject の宣言が食い違っている")
