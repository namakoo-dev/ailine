# -*- coding: utf-8 -*-
"""検体「経費の勘定科目を先例から引く」の宣言（2026-09-13・設計レビュー後に作り直し）。

★ この検体は `src/ailine_core/form_read.py`・`field_record.py`・`forms_collect.py`・
  `split_people.py`・`verify_*.py`・`tests/` 配下を一切読まずに書いた。実装がどう引くかでは
  なく、設計文書と採点器の契約（`bench/accounts/score_accounts.py` docstring・レビュー後版）
  だけを根拠にしている。読んで書くと実装の説明に当たる検体になり、穴を突けなくなる。

読んだのは設計文書 `docs/DESIGN-20260913-経費の勘定科目を先例から引く.md` と
`bench/accounts/score_accounts.py`（答えの JSON の形の契約）だけ。宣言と生成器を分ける
書き方は `bench/split_people/`（読んでよい範囲・フォーム読み取りの実装ではない）を真似た。

## レビュー後の区分の導出（採点器 docstring そのまま・★ 一番大事な変更）
鍵は 4 本 ── 借方取引先／貸方取引先／借方補助科目／摘要（norm で畳んだ完全一致）。
**同じ鍵の過去 N 件は 1 出所**（12 件あっても裏は 1 つ）。今回の 1 行につき、4 本の鍵のうち
値がある鍵だけを過去に当て、それぞれ「1 科目に決まる／2 科目以上で割れる／先例なし」を見る。

    確 … 2 本以上の鍵が当たり、それぞれ 1 科目に決まり、全部同じ科目（食い違う鍵が無い）
    単 … 当たった鍵が 1 本だけで、その鍵が 1 科目に決まる
    割 … どれかの鍵の内訳が 2 科目以上、または鍵どうしが違う科目を指す
    無 … どの鍵にも先例が無い

`mk_accounts.py` はこの規則をそのまま実装した `grade_row()` で機械的に区分を出す
（specimens には区分や科目を一切書かない・過去の事実（鍵・値・科目）と今回の行が
 どの鍵を持つかだけを書く）。

## 検体の単位
1 ケース ＝ 過去の事実（`past_facts`／`extra_past_facts` ── (鍵, 値, 科目) の並び）
＋ 今回のシナリオ（`scenarios[].today_count` 件ぶんの行を機械的に作る）。

シナリオの主な項目:
  vendor           今回の行の取引先文字列（None なら書かない。弥生には取引先の列が無いので使わない）
  credit_side      True なら vendor は「貸方取引先」に書く（借方取引先は空 ── カード払いの形）
  subsidiary       今回の行の「借方補助科目」文字列（None なら空。弥生の支払先はここに書く）
  desc_base/desc_fixed  摘要の元テキスト。desc_fixed=True ならそのまま毎回同じ文字列で書く
                        （＝先例と完全一致しうる）。False なら「{desc_base} {月}月分」を
                        毎回の日付から作る（＝月が変わって完全一致が効かない・実務そのまま）
  desc_overrides   1 行ごとに違う摘要を明示したい時だけ（月ゆれの実例など）
  past_facts       この鍵で過去に何科目・何件あったか（(鍵, 値, 科目) のタプル列）
  untouchable      None（通常の対象行）／"filled"（既に借方勘定科目が埋まっている）／
                   "continuation"（複合仕訳の貸方だけの継続行）／"total"（合計行）
  filled_account   untouchable="filled" の時だけ、既に入っている科目

★ 摘要は実務のまま書く（月・番号・カード会社名・手入力の揺れを含める）。「先例で引ける形」に
  寄せる細工はしない ── 完全一致では無になる行が自然に出る（月ゆれの C06 がその実例）。

## 規模・床（採点器 distribution_floor が要求するもの）
9〜12 ケース・today の合計行 150〜200・区分 4 値それぞれ 8 行以上・
測る対象のソフト（freee を除く）2 種以上・複合仕訳の冊 ≥1・貸方だけ取引先の冊 ≥1・
文字コード 2 種以上・怪しくない冊 ≥1・freee は 1 冊だけ（測る対象から外す）。

## ケース一覧
  C01 mf/csv    アルファ商事   確/単/割/無 の基本形
  C02 mf/csv    ベータ物流     怪しくない（確/単のみ）
  C03 mf/csv    ガンマ興業     複合仕訳（継続行）・既に埋まっている行・合計行
  C04 xlsx      デルタ商店     貸方だけ取引先（カード払い）
  C05 xlsx      イプシロン興業 表記ゆれ（㈱ vs 株式会社）
  C06 弥生/csv  ゼータ工房     月ゆれ（3月分/4月分 電話代・自然に無になる）
  C07 弥生/csv  イータ電機     怪しくない（確/単のみ）
  C08 弥生/csv  シータ食品     鍵どうしの食い違い（割）多め
  C09 弥生/csv  カッパ興産     基本形（和暦混在）
  C10 freee/csv ラムダ工業     freee は 1 冊だけ・測る対象から外す
"""

CASES = [
    # ═══ C01: mf/csv ── 確/単/割/無 の基本形 ══════════════════════════
    dict(
        id="C01", software="mf", company="アルファ商事",
        today_file="MF_仕訳_アルファ商事_2026-07.csv",
        past_files=["MF_過去_アルファ商事.csv"],
        scenarios=[
            dict(label="確:ベータ通信(取引先+摘要が一致)", vendor="ベータ通信サービス",
                 desc_base="電話料金", desc_fixed=True,
                 past_facts=[("借方取引先", "ベータ通信サービス", "通信費"),
                             ("摘要", "電話料金", "通信費")],
                 today_count=5),
            dict(label="単:ガンマ運輸(取引先のみ)", vendor="ガンマ運輸",
                 desc_base="出張旅費", desc_fixed=False,
                 past_facts=[("借方取引先", "ガンマ運輸", "旅費交通費")],
                 today_count=3),
            dict(label="単:保守契約料(摘要のみ・取引先無し)", vendor=None,
                 desc_base="保守契約料一式", desc_fixed=True,
                 past_facts=[("摘要", "保守契約料一式", "支払手数料")],
                 today_count=2),
            dict(label="割:デルタ文具店(鍵内部が2科目)", vendor="デルタ文具店",
                 desc_base="文具購入", desc_fixed=False,
                 past_facts=[("借方取引先", "デルタ文具店", "消耗品費"),
                             ("借方取引先", "デルタ文具店", "雑費")],
                 today_count=2),
            dict(label="割:イプシロン興業(取引先と摘要が違う科目)", vendor="イプシロン興業",
                 desc_base="業務委託費一式", desc_fixed=True,
                 past_facts=[("借方取引先", "イプシロン興業", "雑費"),
                             ("摘要", "業務委託費一式", "会議費")],
                 today_count=2),
            dict(label="無:ゼータ新装(先例なし)", vendor="ゼータ新装",
                 desc_base="内装工事費", desc_fixed=False,
                 past_facts=[], today_count=2),
            dict(label="単:シータ電力", vendor="シータ電力",
                 desc_base="電気料金", desc_fixed=False,
                 past_facts=[("借方取引先", "シータ電力", "水道光熱費")],
                 today_count=2),
            dict(label="単:カッパ興産", vendor="カッパ興産",
                 desc_base="清掃委託費", desc_fixed=False,
                 past_facts=[("借方取引先", "カッパ興産", "雑費")],
                 today_count=2),
        ],
    ),

    # ═══ C02: mf/csv ── 怪しくない（確/単のみ） ══════════════════════
    dict(
        id="C02", software="mf", company="ベータ物流",
        today_file="MF_仕訳_ベータ物流_2026-07.csv",
        past_files=["MF_過去_ベータ物流.csv"],
        怪しくない=True,
        scenarios=[
            dict(label="確:ノルド運輸", vendor="ノルド運輸",
                 desc_base="定期便利用料", desc_fixed=True,
                 past_facts=[("借方取引先", "ノルド運輸", "旅費交通費"),
                             ("摘要", "定期便利用料", "旅費交通費")],
                 today_count=3),
            dict(label="単:サウス倉庫", vendor="サウス倉庫",
                 desc_base="保管料", desc_fixed=False,
                 past_facts=[("借方取引先", "サウス倉庫", "地代家賃")],
                 today_count=3),
            dict(label="単:振込手数料(摘要のみ)", vendor=None,
                 desc_base="振込手数料一式", desc_fixed=True,
                 past_facts=[("摘要", "振込手数料一式", "支払手数料")],
                 today_count=2),
            dict(label="単:セントラル通信", vendor="セントラル通信",
                 desc_base="通信費一式", desc_fixed=False,
                 past_facts=[("借方取引先", "セントラル通信", "通信費")],
                 today_count=3),
            dict(label="確:パシフィック電力", vendor="パシフィック電力",
                 desc_base="電気基本料金", desc_fixed=True,
                 past_facts=[("借方取引先", "パシフィック電力", "水道光熱費"),
                             ("摘要", "電気基本料金", "水道光熱費")],
                 today_count=2),
            dict(label="単:オリエント警備", vendor="オリエント警備",
                 desc_base="警備委託費", desc_fixed=False,
                 past_facts=[("借方取引先", "オリエント警備", "支払手数料")],
                 today_count=2),
        ],
    ),

    # ═══ C03: mf/csv ── 複合仕訳・既に埋まっている行・合計行 ═════════
    dict(
        id="C03", software="mf", company="ガンマ興業",
        today_file="MF_仕訳_ガンマ興業_2026-08.csv",
        past_files=["MF_過去_ガンマ興業_1.csv", "MF_過去_ガンマ興業_2.csv"],
        複合仕訳=True,
        scenarios=[
            dict(label="単:タウ通信サービス", vendor="タウ通信サービス",
                 desc_base="通信費一式", desc_fixed=False,
                 past_facts=[("借方取引先", "タウ通信サービス", "通信費")],
                 today_count=4),
            dict(label="割:ユプシロン運輸(鍵内部が2科目)", vendor="ユプシロン運輸",
                 desc_base="配送料", desc_fixed=False,
                 past_facts=[("借方取引先", "ユプシロン運輸", "旅費交通費"),
                             ("借方取引先", "ユプシロン運輸", "雑費")],
                 today_count=2),
            dict(label="無:ファイ興業(先例なし)", vendor="ファイ興業",
                 desc_base="設備更新費", desc_fixed=False,
                 past_facts=[], today_count=2),
            dict(label="確:カイ電力", vendor="カイ電力",
                 desc_base="電気料金一式", desc_fixed=True,
                 past_facts=[("借方取引先", "カイ電力", "水道光熱費"),
                             ("摘要", "電気料金一式", "水道光熱費")],
                 today_count=2),
            dict(label="単:顧問料(摘要のみ)", vendor=None,
                 desc_base="顧問料", desc_fixed=True,
                 past_facts=[("摘要", "顧問料", "支払手数料")],
                 today_count=2),
            # ── 複合仕訳: 先頭行（対象・単）＋ 継続行（触らない行）───
            dict(label="単:ロー興業(複合仕訳の先頭行)", vendor="ロー興業",
                 desc_base="工事費用", desc_fixed=False,
                 past_facts=[("借方取引先", "ロー興業", "雑費")],
                 today_count=1),
            dict(label="複合仕訳の継続行(貸方だけ・借方は正当に空)",
                 vendor=None, desc_base="（分割・カード引落分）", desc_fixed=True,
                 past_facts=[], today_count=1,
                 untouchable="continuation"),
            # ── 既に借方勘定科目が埋まっている行（触らない行）───
            dict(label="既に埋まっている行1", vendor="オミクロン文具",
                 desc_base="文具購入", desc_fixed=False,
                 past_facts=[], today_count=1,
                 untouchable="filled", filled_account="消耗品費"),
            dict(label="既に埋まっている行2", vendor="パイ運輸",
                 desc_base="出張費", desc_fixed=False,
                 past_facts=[], today_count=1,
                 untouchable="filled", filled_account="旅費交通費"),
            # ── 合計行（触らない行）───
            dict(label="合計行", vendor=None, desc_base="合計", desc_fixed=True,
                 past_facts=[], today_count=1, untouchable="total"),
        ],
    ),

    # ═══ C04: xlsx ── 貸方だけ取引先（カード払い） ═══════════════════
    dict(
        id="C04", software="xlsx", company="デルタ商店",
        today_file="MF_仕訳_デルタ商店_2026-08.xlsx",
        past_files=["MF_過去_デルタ商店.xlsx"],
        貸方だけ取引先=True,
        scenarios=[
            dict(label="確:スミレカード(貸方取引先+摘要が一致)", vendor="スミレカード",
                 credit_side=True, desc_base="カード利用分一括", desc_fixed=True,
                 past_facts=[("貸方取引先", "スミレカード", "消耗品費"),
                             ("摘要", "カード利用分一括", "消耗品費")],
                 today_count=3),
            dict(label="単:ボタンカード(貸方取引先のみ)", vendor="ボタンカード",
                 credit_side=True, desc_base="カード利用分", desc_fixed=False,
                 past_facts=[("貸方取引先", "ボタンカード", "雑費")],
                 today_count=3),
            dict(label="割:ヒマワリカード(鍵内部が2科目)", vendor="ヒマワリカード",
                 credit_side=True, desc_base="カード利用分", desc_fixed=False,
                 past_facts=[("貸方取引先", "ヒマワリカード", "会議費"),
                             ("貸方取引先", "ヒマワリカード", "接待交際費")],
                 today_count=2),
            dict(label="無:アヤメカード(先例なし)", vendor="アヤメカード",
                 credit_side=True, desc_base="新規カード利用分", desc_fixed=False,
                 past_facts=[], today_count=2),
            dict(label="単:振込手数料(摘要のみ・現金払い)", vendor=None,
                 desc_base="振込手数料一式", desc_fixed=True,
                 past_facts=[("摘要", "振込手数料一式", "支払手数料")],
                 today_count=2),
            dict(label="単:ツバキ工業(通常の借方取引先・非カード)", vendor="ツバキ工業",
                 desc_base="部材購入費", desc_fixed=False,
                 past_facts=[("借方取引先", "ツバキ工業", "消耗品費")],
                 today_count=2),
        ],
    ),

    # ═══ C05: xlsx ── 表記ゆれ（㈱ vs 株式会社） ═════════════════════
    dict(
        id="C05", software="xlsx", company="イプシロン興業",
        today_file="MF_仕訳_イプシロン興業_2026-08.xlsx",
        past_files=["MF_過去_イプシロン興業.xlsx"],
        extra_past_facts=[("借方取引先", "㈱イプシロン興業", "接待交際費")],
        scenarios=[
            dict(label="無:表記ゆれ(先例は㈱イプシロン興業・別の鍵)",
                 vendor="株式会社イプシロン興業",
                 desc_base="懇親会費", desc_fixed=False,
                 past_facts=[], today_count=2),
            dict(label="確:ゼータ運送", vendor="ゼータ運送",
                 desc_base="定期輸送料一式", desc_fixed=True,
                 past_facts=[("借方取引先", "ゼータ運送", "旅費交通費"),
                             ("摘要", "定期輸送料一式", "旅費交通費")],
                 today_count=3),
            dict(label="単:イータ設備", vendor="イータ設備",
                 desc_base="設備保守費", desc_fixed=False,
                 past_facts=[("借方取引先", "イータ設備", "支払手数料")],
                 today_count=3),
            dict(label="割:シータ倉庫(鍵内部が2科目)", vendor="シータ倉庫",
                 desc_base="保管費", desc_fixed=False,
                 past_facts=[("借方取引先", "シータ倉庫", "地代家賃"),
                             ("借方取引先", "シータ倉庫", "雑費")],
                 today_count=2),
            dict(label="単:新聞購読料(摘要のみ)", vendor=None,
                 desc_base="新聞購読料", desc_fixed=True,
                 past_facts=[("摘要", "新聞購読料", "新聞図書費")],
                 today_count=4),
        ],
    ),

    # ═══ C06: 弥生/csv ── 月ゆれ（3月分/4月分 電話代） ═══════════════
    dict(
        id="C06", software="yayoi", company="ゼータ工房",
        today_file="弥生_仕訳_ゼータ工房_202608.csv",
        past_files=["弥生_過去_ゼータ工房_1.csv", "弥生_過去_ゼータ工房_2.csv"],
        scenarios=[
            dict(label="確:駐車場代(補助科目+摘要が一致)", subsidiary="オメガ商店",
                 desc_base="駐車場代", desc_fixed=True,
                 past_facts=[("借方補助科目", "オメガ商店", "地代家賃"),
                             ("摘要", "駐車場代", "地代家賃")],
                 today_count=3),
            dict(label="単:消耗品購入(摘要のみ)", subsidiary=None,
                 desc_base="消耗品購入", desc_fixed=True,
                 past_facts=[("摘要", "消耗品購入", "消耗品費")],
                 today_count=5),
            dict(label="単:カッパ興業(補助科目のみ)", subsidiary="カッパ興業",
                 desc_base="業務委託費", desc_fixed=False,
                 past_facts=[("借方補助科目", "カッパ興業", "雑費")],
                 today_count=2),
            dict(label="割:運送費一式(鍵内部が2科目)", subsidiary=None,
                 desc_base="運送費一式", desc_fixed=True,
                 past_facts=[("摘要", "運送費一式", "旅費交通費"),
                             ("摘要", "運送費一式", "雑費")],
                 today_count=2),
            dict(label="割:ラムダ工業(鍵どうしが違う科目)", subsidiary="ラムダ工業",
                 desc_base="打合せ費用一式", desc_fixed=True,
                 past_facts=[("借方補助科目", "ラムダ工業", "会議費"),
                             ("摘要", "打合せ費用一式", "接待交際費")],
                 today_count=2),
            dict(label="無:電話代の月ゆれ(先例は「電話代」・完全一致せず無)",
                 subsidiary=None,
                 desc_overrides=["3月分 電話代", "4月分 電話代"],
                 past_facts=[("摘要", "電話代", "通信費")],
                 today_count=2),
            dict(label="単:シータ興産(補助科目のみ)", subsidiary="シータ興産",
                 desc_base="新聞購読料", desc_fixed=False,
                 past_facts=[("借方補助科目", "シータ興産", "新聞図書費")],
                 today_count=1),
        ],
    ),

    # ═══ C07: 弥生/csv ── 怪しくない（確/単のみ） ══════════════════
    dict(
        id="C07", software="yayoi", company="イータ電機",
        today_file="弥生_仕訳_イータ電機_202607.csv",
        past_files=["弥生_過去_イータ電機.csv"],
        怪しくない=True,
        scenarios=[
            dict(label="確:宅配便定期便", subsidiary="ミュー商会",
                 desc_base="宅配便定期便", desc_fixed=True,
                 past_facts=[("借方補助科目", "ミュー商会", "旅費交通費"),
                             ("摘要", "宅配便定期便", "旅費交通費")],
                 today_count=3),
            dict(label="単:事務用品代(摘要のみ)", subsidiary=None,
                 desc_base="事務用品代一式", desc_fixed=True,
                 past_facts=[("摘要", "事務用品代一式", "消耗品費")],
                 today_count=3),
            dict(label="単:ニュー電設(補助科目のみ)", subsidiary="ニュー電設",
                 desc_base="電気工事費", desc_fixed=False,
                 past_facts=[("借方補助科目", "ニュー電設", "水道光熱費")],
                 today_count=2),
            dict(label="確:新聞購読料一式", subsidiary="クサイ商店",
                 desc_base="新聞購読料一式", desc_fixed=True,
                 past_facts=[("借方補助科目", "クサイ商店", "新聞図書費"),
                             ("摘要", "新聞購読料一式", "新聞図書費")],
                 today_count=2),
            dict(label="単:来客対応費(摘要のみ)", subsidiary=None,
                 desc_base="来客対応費一式", desc_fixed=True,
                 past_facts=[("摘要", "来客対応費一式", "接待交際費")],
                 today_count=2),
            dict(label="単:オミクロン設備(補助科目のみ)", subsidiary="オミクロン設備",
                 desc_base="賃借料", desc_fixed=False,
                 past_facts=[("借方補助科目", "オミクロン設備", "地代家賃")],
                 today_count=2),
        ],
    ),

    # ═══ C08: 弥生/csv ── 鍵どうしの食い違い（割）多め ═══════════════
    dict(
        id="C08", software="yayoi", company="シータ食品",
        today_file="弥生_仕訳_シータ食品_202608.csv",
        past_files=["弥生_過去_シータ食品.csv"],
        scenarios=[
            dict(label="単:パイ運輸(補助科目のみ)", subsidiary="パイ運輸",
                 desc_base="配送費", desc_fixed=False,
                 past_facts=[("借方補助科目", "パイ運輸", "旅費交通費")],
                 today_count=5),
            dict(label="割:ロー興業(鍵どうしが違う科目)", subsidiary="ロー興業",
                 desc_base="業務委託費一式", desc_fixed=True,
                 past_facts=[("借方補助科目", "ロー興業", "雑費"),
                             ("摘要", "業務委託費一式", "会議費")],
                 today_count=2),
            dict(label="割:修繕費一式(鍵内部が2科目)", subsidiary=None,
                 desc_base="修繕費一式", desc_fixed=True,
                 past_facts=[("摘要", "修繕費一式", "消耗品費"),
                             ("摘要", "修繕費一式", "雑費")],
                 today_count=2),
            dict(label="無:シグマ興業(先例なし)", subsidiary="シグマ興業",
                 desc_base="新規委託費", desc_fixed=False,
                 past_facts=[], today_count=2),
            dict(label="確:保守契約料一式", subsidiary="タウ商会",
                 desc_base="保守契約料一式", desc_fixed=True,
                 past_facts=[("借方補助科目", "タウ商会", "支払手数料"),
                             ("摘要", "保守契約料一式", "支払手数料")],
                 today_count=3),
            dict(label="単:水道料金一式(摘要のみ)", subsidiary=None,
                 desc_base="水道料金一式", desc_fixed=True,
                 past_facts=[("摘要", "水道料金一式", "水道光熱費")],
                 today_count=2),
        ],
    ),

    # ═══ C09: 弥生/csv ── 基本形（和暦混在） ═══════════════════════
    dict(
        id="C09", software="yayoi", company="カッパ興産",
        today_file="弥生_仕訳_カッパ興産_202608.csv",
        past_files=["弥生_過去_カッパ興産.csv"],
        scenarios=[
            dict(label="確:印刷委託費一式", subsidiary="ユプシロン印刷",
                 desc_base="印刷委託費一式", desc_fixed=True,
                 past_facts=[("借方補助科目", "ユプシロン印刷", "消耗品費"),
                             ("摘要", "印刷委託費一式", "消耗品費")],
                 today_count=3),
            dict(label="単:ファイ運送(補助科目のみ)", subsidiary="ファイ運送",
                 desc_base="配送費", desc_fixed=False,
                 past_facts=[("借方補助科目", "ファイ運送", "旅費交通費")],
                 today_count=5),
            dict(label="単:振込手数料一式(摘要のみ)", subsidiary=None,
                 desc_base="振込手数料一式", desc_fixed=True,
                 past_facts=[("摘要", "振込手数料一式", "支払手数料")],
                 today_count=2),
            dict(label="割:カイ商店(鍵内部が2科目)", subsidiary="カイ商店",
                 desc_base="消耗品費用", desc_fixed=False,
                 past_facts=[("借方補助科目", "カイ商店", "消耗品費"),
                             ("借方補助科目", "カイ商店", "雑費")],
                 today_count=2),
            dict(label="無:オメガ工業(先例なし)", subsidiary="オメガ工業",
                 desc_base="新規発注費", desc_fixed=False,
                 past_facts=[], today_count=2),
            dict(label="単:ニュー倉庫(補助科目のみ)", subsidiary="ニュー倉庫",
                 desc_base="保管料", desc_fixed=False,
                 past_facts=[("借方補助科目", "ニュー倉庫", "地代家賃")],
                 today_count=2),
        ],
    ),

    # ═══ C10: freee/csv ── 1 冊だけ・測る対象から外す ═══════════════
    dict(
        id="C10", software="freee", company="ラムダ工業",
        today_file="freee_仕訳帳_ラムダ工業_2026-08.csv",
        past_files=["freee_過去_ラムダ工業.csv"],
        scenarios=[
            dict(label="確:オメガ食品商事(取引先+摘要が一致)", vendor="オメガ食品商事",
                 desc_base="定期仕入諸経費", desc_fixed=True,
                 past_facts=[("借方取引先", "オメガ食品商事", "消耗品費"),
                             ("摘要", "定期仕入諸経費", "消耗品費")],
                 today_count=3),
            dict(label="単:ファイ運輸(取引先のみ)", vendor="ファイ運輸",
                 desc_base="配送費", desc_fixed=False,
                 past_facts=[("借方取引先", "ファイ運輸", "旅費交通費")],
                 today_count=5),
            dict(label="単:振込手数料一式(摘要のみ)", vendor=None,
                 desc_base="振込手数料一式", desc_fixed=True,
                 past_facts=[("摘要", "振込手数料一式", "支払手数料")],
                 today_count=2),
            dict(label="割:プサイ興業(鍵内部が2科目)", vendor="プサイ興業",
                 desc_base="業務委託費", desc_fixed=False,
                 past_facts=[("借方取引先", "プサイ興業", "雑費"),
                             ("借方取引先", "プサイ興業", "会議費")],
                 today_count=2),
            dict(label="無:シグマ興産(先例なし)", vendor="シグマ興産",
                 desc_base="新規委託費", desc_fixed=False,
                 past_facts=[], today_count=2),
            dict(label="単:タウ通信サービス(取引先のみ)", vendor="タウ通信サービス",
                 desc_base="通信費一式", desc_fixed=False,
                 past_facts=[("借方取引先", "タウ通信サービス", "通信費")],
                 today_count=2),
        ],
    ),
]

__all__ = ["CASES"]
