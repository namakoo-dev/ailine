"""`ailine ops`（頼める操作の一覧）の番人。

★ なぜ在るか（2026-08-16 の盲検査定 2 本）: 独立した 2 人が揃って MISSING の筆頭に
「対応操作の一覧が無い」を挙げた。README 368 行の中で**何を頼めるか**が分からず、
語彙外の依頼で聞き返しループに入り「普通の購入検討者ならここで評価を終える」と書かれた。
★ 一覧は**登録簿から生成**する（手書きの表は必ずずれる）。この試験は
「操作を足したのに一覧に出ない」を機械で捕まえる。
"""
import subprocess
import sys
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO))

import ailine  # noqa: E402
from ailine_core.cli_render import render_ops_table  # noqa: E402


def _table():
    return render_ops_table(ailine.OP_META, ailine.OP_SCHEMA, ailine._CONFIRM_FIELDS)


def test_every_dsl_op_appears_in_the_listing():
    """★ 本命: OP_SCHEMA の全 op のラベルが一覧に出ること（足したのに出ない、を防ぐ）。"""
    text = "\n".join(_table())
    missing = [op for op in ailine.OP_SCHEMA if ailine.OP_LABELS[op] not in text]
    assert not missing, f"一覧に出ない操作がある: {missing}（OP_META への追記漏れ）"


def test_listing_declares_what_it_cannot_do():
    """★ 「できない」と明言する行があること。査定 A は語彙外の依頼を 4 回言い直して
    4 回とも質問返しになり、未対応だと分からないまま詰んだ。"""
    text = "\n".join(_table())
    # ★ freeform 最終決定 (2026-08-21): 旧文の凍結を新しい正直な文へ更新
    # （負の被覆: 旧い約束が復活しないことも見る）
    assert "頼める操作の一覧に照合できないため生成せず断ります" in text
    assert "今はできません" not in text
    assert "未対応" in text


def test_needed_info_uses_existing_japanese_labels_only():
    """必要な情報の日本語は確認行の登録簿から引く（新語を作らない＝二重帳簿にしない）。"""
    for op, slots in ailine.OP_SCHEMA.items():
        labels = {slot for _lab, slot, _f in ailine._CONFIRM_FIELDS.get(op, ())}
        for slot in slots:
            # 確認行に無い slot は英語のまま出る。それ自体は許すが、増えたら気づけるよう記録
            if slot not in labels:
                assert slot.isascii(), f"{op}: 確認行に無い slot が非 ASCII: {slot}"


def test_ops_subcommand_runs_and_exits_zero():
    r = subprocess.run([sys.executable, str(REPO / "ailine.py"), "ops"],
                       capture_output=True, text=True, encoding="utf-8")
    assert r.returncode == 0, r.stderr
    assert "ailine に頼めること" in r.stdout


def test_clarify_points_at_the_way_out():
    """★ 行き止まりに出口を置く。聞き返しは「言い方が悪い」と「未対応」を区別できないので、
    区別する手段を毎回そえる。"""
    src = (REPO / "ailine.py").read_text(encoding="utf-8")
    assert "（頼める操作の一覧: ailine ops）" in src
