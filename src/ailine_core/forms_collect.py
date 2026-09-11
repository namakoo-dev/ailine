"""forms_collect — 帳票の山から項目を集めて **1 冊 1 行**の一覧にする（2026-09-11）。

★★ 需要①（請求書 40 ファイルからの抜き出し）の出口。読む器官そのものは
  `form_read.read_book` が持ち、ここは**集めた結果を人が読める形に並べる**だけ。

★★ 凍結した判断をそのまま運ぶ（docs/DESIGN-20260910… §0f・Namakoo 決裁 2026-09-11）:

    確 / 単 → 値を出す        一覧のセルに書く
    割 / 無 → 値を出さない    セルは**空欄**。理由は必ず検分シートに出る

  ★ 空欄に理由が付かない経路を作らない ── `Record` が型で禁じているので、
    ここで «理由が無い空欄» を作ることは**できない**（後から足す形にしない）。

★ この層は区分を**作らない**。`field_record.grade()` に聞くだけ
  （`tests/test_a_grade_has_exactly_one_source.py` が AST で縛っている）。

★ ailine を import しない（ailine_core の作法）。
"""
from __future__ import annotations

from ailine_core.field_record import GRADES_WITH_VALUE, describe, grade, value
from ailine_core.form_read import FIELDS as _ORGAN_FIELDS
from ailine_core import forms_suspect

#: 書き手の印。★ stack.py の KIND_SIGNATURES に登録して初めて「自分の出力」と分かる
#:   （印だけでも列だけでも足りない ── 両方そろって自分の出力・fail closed）。
CREATOR_MARK = "ailine forms"

#: 一覧シートの名前と見出し。★ 署名でもあるので、変えると過去の出力が他人のものになる。
SHEET_NAME = "一覧"
HEADERS = ("元ファイル", "請求元", "宛先", "請求額(税込)", "請求日", "請求番号")

#: 一覧に載せる項目（見出しの 2 列目以降と 1 対 1）。★ 器官の一覧から導く（自前で持たない）。
FIELDS = tuple(f for f in ("請求元", "宛先", "請求額", "請求日", "請求番号") if f in _ORGAN_FIELDS)
assert set(FIELDS) == set(_ORGAN_FIELDS), "★ 一覧の項目と器官の項目がずれている"

#: 検分シートに出す見出し。
INSPECT_SHEET = "検分"
INSPECT_HEADERS = ("元ファイル", "項目", "区分", "なぜこうなったか")

#: 束の所見（1 冊ずつは正常でも、束で見ると怪しいもの）を出すシート。
SUSPECT_SHEET = "束の所見"
SUSPECT_HEADERS = ("種類", "関わる冊", "なぜ怪しいか")


def row_for(name: str, records: dict) -> list:
    """1 冊ぶんの一覧の行。★ 何を出すかは `value()` **だけ**が決める。

    ★ 初版はここでも `grade(rec) in GRADES_WITH_VALUE` を見ていたが、`value()` は既に
      割/無 で None を返す ── **守りが二重**になり、どちらを壊しても変異試験が緑になった
      （2026-09-11・同じ日に敬称の処理でも踏んだ形）。1 箇所に畳む。
      値を出す/出さないの契約は `field_record` の試験が縛っている。
    """
    return [name] + [value(records.get(field)) if records.get(field) is not None else None
                     for field in FIELDS]


def findings_for(name: str, records: dict) -> list:
    """検分シートの行 ── **人が手を動かすものだけ**（空欄＝割 / 無）。

    ★★ 初版は「単（裏が取れていない）」も 1 行ずつ出していた。実測（検体 87 冊）:

        単 201 行 ／ 割 12 行 ／ 無 15 行   ── 88% が雑音で、**見てほしい 27 行を埋めた**

      ★ `単` は**普通の状態**だ（口が 1 つしかない項目がほとんど）。普通を所見として
        並べるのは、この repo が一度直した「オオカミ少年防止を謳う道具が自分で
        オオカミ少年になる」形そのもの。
      → 検分に載せるのは **空欄だけ**。裏が取れているかの内訳は要約行と `--json` に出す
        （消すのではなく、**行動が要る所と、質の統計を分ける**）。
    """
    rows = []
    for field in FIELDS:
        rec = records.get(field)
        if rec is None or grade(rec) in GRADES_WITH_VALUE:
            continue
        rows.append([name, field, grade(rec), describe(rec)])
    return rows


def grades_per_file(all_records: list) -> list:
    """機械可読の側へ出す、1 冊 1 項目ごとの区分。★ 人向けの画面には出さない。

    ★ 検分から `単` を外したぶん、「どの値が裏取り済みか」を知りたい自動化のために
      ここに残す ── **消したのではなく、出す先を分けた**。
    """
    return [{"file": name, "field": field, "grade": grade(records[field])}
            for name, records in all_records for field in FIELDS if records.get(field) is not None]


def tally(all_records: list) -> dict:
    """区分ごとの件数（分母つきの報告に使う）。all_records: [(名前, records), …]"""
    counts: dict = {}
    for _name, records in all_records:
        for field in FIELDS:
            rec = records.get(field)
            if rec is None:
                continue
            counts[grade(rec)] = counts.get(grade(rec), 0) + 1
    return counts


def blanks_have_reasons(all_records: list) -> tuple:
    """★ 事後条件: 空欄の数 == 理由の数（G2' の必達をこの経路でも機械で確かめる）。

    戻り値: (空欄の数, 理由のある空欄の数, 理由の無い項目の名前)
    ★ `Record` が型で禁じているので普通は起きない。それでも**ここで数える** ──
      「型が守っているはず」は検算ではない（この repo が何度も踏んだ形）。
    """
    blanks, with_reason, bad = 0, 0, []
    for name, records in all_records:
        for field in FIELDS:
            rec = records.get(field)
            if rec is None:
                # ★ 記録そのものが無い ＝ 理由の無い空欄（2026-09-11 に実際に生まれた形）。
                #   初版はここを continue で飛ばしていて、この穴を**数えていなかった**。
                blanks += 1
                bad.append(f"{name}/{field}（記録が無い）")
                continue
            if grade(rec) in GRADES_WITH_VALUE:
                continue
            blanks += 1
            if rec.blank_reason.strip():
                with_reason += 1
            else:
                bad.append(f"{name}/{field}")
    return blanks, with_reason, bad


def suspicions_for(all_records: list) -> list:
    """束の所見。all_records: [(名前, records), …] → [{"種類", "冊", "理由"}, …]

    ★ 疑う材料は `value()` が出した値**だけ**（割/無 は None ── 読めなかったものを根拠にしない）。
      値を出す/出さないの線はここでも `field_record` のもの（書き写さない）。
    """
    table = {name: {field: value(records[field]) for field in FIELDS if records.get(field) is not None}
             for name, records in all_records}
    return forms_suspect.suspect(table)


def suspicion_rows(suspicions: list) -> list:
    """束の所見シートの行（見出しは SUSPECT_HEADERS）。"""
    return [[s["種類"], "／".join(s["冊"]), s["理由"]] for s in suspicions]
