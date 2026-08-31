# GUI を **同じポートに 2 つ重ねない**番人（2026-08-31）。
#
# ★★ 実測した事故: 8/28 に起動したままの GUI と、今日起動した GUI が **同じ 8760 に
#   2 つ**載っていた。`http.server` は `allow_reuse_address = 1` を持つので、
#   Windows では後から来た方も bind に成功する（SO_REUSEADDR は乗っ取りを許す）。
#   どちらが応答するかは運で、実際 `/api/read` が **404** を返した ── 古い方（その
#   入口をまだ持っていない版）が answer していたため。
#   画面には「（読み取れませんでした）」としか出ないので、原因に辿り着けない。
#
# ★ この形は「番人が在るのに鳴らない」ではなく「**別人が返事をしている**」。
#   新しい方が正しく動いていても、利用者には壊れて見える。
#   ★ しかも **直したはずの不具合が再現する** ── 直しが効いていないように見える。
#
# 契約:
#   ① 応答が既に在るポートには載らない（起動を断る・終了コードは 0 でない）
#   ② 空いているポートでは今までどおり起動する（断りを広げすぎない）
#   ③ 断りの文面は、**今応答しているものの見つけ方**を言う（落ちるコマンドを勧めない）

import importlib.util
import socket
import sys
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent
SERVER_PY = REPO / "gui" / "server.py"


def _server_module():
    spec = importlib.util.spec_from_file_location("_ailine_gui_server", SERVER_PY)
    mod = importlib.util.module_from_spec(spec)
    sys.modules[spec.name] = mod
    spec.loader.exec_module(mod)
    return mod


def _free_port() -> int:
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        s.bind(("127.0.0.1", 0))
        return s.getsockname()[1]


def test_it_sees_a_port_that_is_already_answering():
    """① 誰かが listen していれば True。"""
    mod = _server_module()
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as srv:
        srv.bind(("127.0.0.1", 0))
        srv.listen(1)
        port = srv.getsockname()[1]
        assert mod.port_already_answers(port) is True


def test_it_does_not_cry_wolf_on_a_free_port():
    """② 空いていれば False（断りを広げすぎると、普通の起動が止まる）。

    ★ 陽性対照だけでは測れない ── 「いつも True」でも ① は通ってしまう。
    """
    mod = _server_module()
    assert mod.port_already_answers(_free_port()) is False


def test_main_refuses_before_binding_when_the_port_answers():
    """① を main が実際に使っていること（在るのに配線されていない、を作らない）。

    ★ 関数が在るだけでは意味がない。**本番の経路が通る**ことを見る。
    """
    src = SERVER_PY.read_text(encoding="utf-8")
    i = src.index("def main(")
    body = src[i:]
    call = body.index("port_already_answers(")
    bind = body.index("ThreadingHTTPServer(")
    assert call < bind, "確かめる前に bind している（重ねてから気づいても遅い）"


def test_the_refusal_says_how_to_find_who_is_answering():
    """③ 断りは、次にやることを言う（この repo の作法）。"""
    src = SERVER_PY.read_text(encoding="utf-8")
    i = src.index("def main(")
    body = src[i:]
    assert "Get-NetTCPConnection" in body, "今応答しているものの見つけ方を書く"
    assert "--port" in body, "並べたいときの逃げ道を書く"
    # ★ 名前でまとめて止めさせない（過去に taskkill /IM で無関係のプロセスを撃った）。
    assert "PID" in body
