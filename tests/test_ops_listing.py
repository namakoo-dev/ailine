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
sys.path.insert(0, str(REPO / "src"))

import ailine  # noqa: E402
from ailine_core.cli_render import render_ops_table  # noqa: E402
from _product_source import product_text  # noqa: E402 ── ★ 番人は本体決め打ちでなく製品コード全体を読む


def _table():
    return render_ops_table(ailine.OP_META, ailine.OP_SCHEMA, ailine._CONFIRM_FIELDS)


def test_every_dsl_op_appears_in_the_listing():
    """★ 本命: OP_SCHEMA の全 op のラベルが一覧に出ること（足したのに出ない、を防ぐ）。"""
    text = "\n".join(_table())
    missing = [op for op in ailine.OP_SCHEMA if ailine.OP_LABELS[op] not in text]
    assert not missing, f"一覧に出ない操作がある: {missing}（OP_META への追記漏れ）"


def test_the_readme_table_matches_the_register():
    """★ README の手書きの表が、登録簿とずれていないこと（2026-09-05・盲検の査定）。

    査定の指摘: README は「この表は登録簿から**自動生成される**ので、文書とずれません」
    と書いていたが、実際は手書きの Markdown で、突き合わせる機械も無かった。
    中身はたまたま合っていたので**間違っていたのは数字でなく「ずれない理由」**。
    ★ 次に op を足したとき静かに腐る形だったので、README の文言を実体に寄せ
      （「食い違ったらここが赤くなる」）、その約束をここで機械にする。
    """
    text = (REPO / "README.md").read_text(encoding="utf-8")
    table = [ln for ln in text.splitlines() if ln.startswith("| ") and " | " in ln]
    listed = set()
    for ln in table:
        for cell in ln.split("|"):
            for word in cell.replace("/", " ").split():
                if word in ailine.OP_SCHEMA:
                    listed.add(word)
    missing = sorted(set(ailine.OP_SCHEMA) - listed)
    assert not missing, (
        f"README の表に無い操作がある: {missing} ── "
        "README は「食い違ったらここが赤くなる」と書いているので、表に足すこと")
    unknown = sorted(listed - set(ailine.OP_SCHEMA))
    assert not unknown, f"README の表に、登録簿に無い操作が載っている: {unknown}"


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
    r = subprocess.run([sys.executable, "-m", "ailine", "ops"],
                       capture_output=True, text=True, encoding="utf-8")
    assert r.returncode == 0, r.stderr
    assert "ailine に頼めること" in r.stdout


def test_clarify_points_at_the_way_out():
    """★ 行き止まりに出口を置く。聞き返しは「言い方が悪い」と「未対応」を区別できないので、
    区別する手段を毎回そえる。"""
    assert "（頼める操作の一覧: ailine ops）" in product_text()
