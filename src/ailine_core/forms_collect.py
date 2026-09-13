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

#: 一覧の中で**金額**の列（3 桁区切りで見せる ── C8）。
#: ★ 列番号は手書きしない（`money_columns()` が `HEADERS` から導く ── 並びを変えても追う）。
MONEY_HEADERS = ("請求額(税込)",)


def money_column_indexes() -> tuple:
    """金額の列の 1 始まりの番号（`HEADERS` の並びから導く）。

    ★ `inspection.money_columns`（書式を**付ける**器）と名前を分ける ── 同名だと
      「同名だが実装が違う」台帳に載る（`tests/test_duplicate_definitions_ledger.py`）。
    """
    return tuple(HEADERS.index(h) + 1 for h in MONEY_HEADERS)

#: 束の所見（1 冊ずつは正常でも、束で見ると怪しいもの）を出すシート。
SUSPECT_SHEET = "束の所見"
SUSPECT_HEADERS = ("種類", "関わる冊", "なぜ怪しいか")

#: 束の要約（画面に出た事実を、ブックだけ開く人のために残す）シート。
#: ★★ 2026-09-13（2 回目の買い手役・経理）: 「月末、黒い画面をずっと見ているわけではない（翌朝ブック
#:   だけ開く）」── 読めなかった冊・載せなかった冊・月の混在が画面にしか無く、ブックに 1 セルも残らなかった。
SUMMARY_SHEET = "束の要約"
SUMMARY_HEADERS = ("項目", "冊", "内容")


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


def nothing_found(all_records: list) -> list:
    """**1 項目も値が出なかった冊**の名前（★ 請求書でない冊が混ざっている徴候）。

    ★★ なぜ在るか（2026-09-13・買い手の初見 B7）: 送付状・稟議書のような請求書でない冊が
      受領フォルダに混ざると、一覧に**全列が空の行**として並び、画面には何も出なかった。
      買い手には「請求書なのに読めなかった」と区別が付かない ── 次の一手が正反対になる
      （読めない冊は直す・請求書でない冊は放っておく）。
    ★ PDF 側には同じ線が先に在った（設計 D7「読めなかった」と「請求書ではなかった」を
      混ぜない）── xlsx 側が未配線だった（この repo の片配線の形・開発手法 §13）。
    ★ ここでは**一覧から外さない**（分母を動かさない）── 名指しして人に言うだけ。
      外すかどうかは実物の分布を見てから決める（設計 §10 の発火条件つき保留）。
    """
    out = []
    for name, records in all_records:
        got = [field for field in FIELDS
               if records.get(field) is not None and value(records[field]) is not None]
        if not got:
            out.append(name)
    return out


def summary_rows(result: dict, all_records: list) -> list:
    """束の要約シートの行（見出しは SUMMARY_HEADERS）── 画面の報告と同じ事実を同じ順で。"""
    rows = [["読んだ冊", "", f"{result.get('denominator', 0)} ファイル中 {result.get('collected', 0)} 冊"]]
    for u in result.get("unreadable") or ():
        rows.append(["読めなかった冊", u.get("name", ""), u.get("reason", "")])
    for name in result.get("nothing_found") or ():
        rows.append(["一覧に載せていない冊", name, "項目が 1 つも取れませんでした（請求書でない冊の可能性）── 理由は『検分』に"])
    for name in result.get("self_excluded") or ():
        rows.append(["入力に数えなかった冊", name, "ailine の前回の出力です"])
    by_month: dict = {}
    for name, records in all_records:
        rec = records.get("請求日")
        d = value(rec) if rec is not None else None
        key = f"{d.year:04d}年{d.month}月" if hasattr(d, "year") and hasattr(d, "month") else "請求日なし"
        by_month.setdefault(key, []).append(name)
    if len([k for k in by_month if k != "請求日なし"]) >= 2:
        for key in sorted(by_month):
            rows.append(["請求日の月", "／".join(by_month[key]), f"{key}: {len(by_month[key])} 冊"])
    sus = result.get("suspicions") or ()
    rows.append(["束の所見", "", f"{len(sus)} 件（『束の所見』シート）"])
    return rows


def months_of(all_records: list) -> dict:
    """請求日の**月ごとの冊数**（`"2026-09"` → 15）。日付の無い冊は `""` に数える。

    ★★ 2026-09-13（買い手役の初見・経理）: 「9月受領分」のフォルダに 5〜8 月の請求が 4 冊
      （142,000 円）黙って混ざっていた。道具は月を理解している（束の所見に「どちらも 2026年9月」）
      のに、期間外を一言も言わなかった。★ 疑いにはしない ── 受領フォルダに前月分が遅れて混ざるのは
      実務で普通（検体の設計 §9.2 の ASSUMED）。**内訳を 1 行言う**だけ。数えるのは値の出た請求日。
    """
    out: dict = {}
    for _name, records in all_records:
        rec = records.get("請求日")
        d = value(rec) if rec is not None else None
        key = f"{d.year:04d}-{d.month:02d}" if hasattr(d, "year") and hasattr(d, "month") else ""
        out[key] = out.get(key, 0) + 1
    return out


def grades_per_file(all_records: list) -> list:
    """機械可読の側へ出す、1 冊 1 項目ごとの区分。★ 人向けの画面には出さない。

    ★ 検分から `単` を外したぶん、「どの値が裏取り済みか」を知りたい自動化のために
      ここに残す ── **消したのではなく、出す先を分けた**。
    """
    return [{"file": name, "field": field, "grade": grade(records[field])}
            for name, records in all_records for field in FIELDS if records.get(field) is not None]


def tally_by_field(all_records: list) -> dict:
    """項目ごとの区分の件数 {項目: {区分: n}}（★ 集計の 1 行に溶かさない）。

    ★★ 2026-09-13（2 回目の買い手役・経理）: 「確 20／単 88／無 12」の 1 行では、確 20 が**全部
      請求額**で「Excel 20 冊は金額の裏が取れ・PDF 3 冊は根拠 1 つ」という一番知りたい事実が
      消えていた。項目ごとに言えば、買い手は PDF 3 枚だけ電卓を叩けば済む。
    """
    out: dict = {}
    for _name, records in all_records:
        for field in FIELDS:
            rec = records.get(field)
            if rec is None:
                continue
            g = grade(rec)
            out.setdefault(field, {})[g] = out.setdefault(field, {}).get(g, 0) + 1
    return out


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
