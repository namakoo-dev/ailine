# -*- coding: utf-8 -*-
"""骨（実物の雛形）の番地表。★ 全部 dump.py で開いて目で確かめた値（2026-09-11）。

★ 別の雛形を足すときは、必ずその雛形を開いて番地を確かめてから足すこと。
★ ここに嘘の番地を書くと、答えが嘘になる ── 読み戻し検査で落ちるので黙って通ることは無い。
"""
from pathlib import Path

SP = Path(__file__).resolve().parent   # ★ 雛形はこの隣（既存 mk_received.py と同じ規約）
#   ★ 2026-09-11: 初版は .parent.parent（作業場では雛形が親に在ったため）。
#     repo に置いた瞬間に bench/ を見て落ちた ── 規約が置き場に依存していた。

MISOCA_COMMON = dict(kind="請求書", vendor="misoca", tax_style="single_rate")

SKELETONS = {
    # ── misoca 4 変種（同じベンダ・幾何が全部違う）─────────────────────
    "misoca256": dict(MISOCA_COMMON,
        src=SP / "real14" / "a0cf7302-blackline.xlsx", sheet="misoca_invoice",
        issuer=dict(name="G5", zip="G6", addr=["G7", "G8"], tel="G9", fax="G10",
                    mail="G11", tanto="G12"),
        to=dict(name="B3", attn="B4"),
        detail=dict(first=16, last=36, name="B", qty="E", unitname="F", price="G",
                    amount="H", clear=list("BCDEFG")),
        sums=dict(小計="H37", 消費税="H38", 合計="H39", 請求額="C11"),
        rate_cell="G38", note=["B41", "B42"], furikomi="B45"),
    "misoca12": dict(MISOCA_COMMON,
        src=SP / "real14" / "6d186d0c-simpleblack.xlsx", sheet="misoca_invoice",
        issuer=dict(name="G7", zip="G8", addr=["G9", "G10"], tel="G11", fax="G12",
                    mail="G13", tanto="G14"),
        to=dict(name="B4", attn="B6"),
        detail=dict(first=17, last=37, name="B", qty="F", unitname="G", price="H",
                    amount="I", clear=list("BCDEFGH")),
        sums=dict(小計="I38", 消費税="I39", 合計="I40", 請求額="D13"),
        rate_cell="H39", note=["B42", "B43"], furikomi=None),
    "misoca13": dict(MISOCA_COMMON,
        src=SP / "real14" / "501b2d6e-simpleline.xlsx", sheet="misoca_invoice",
        issuer=dict(name="H5", zip="H6", addr=["H7", "H8"], tel="H9", fax="H10",
                    mail="H11", tanto="H12"),
        to=dict(name="C7", attn="C8"),
        detail=dict(first=17, last=37, name="C", qty="F", unitname="G", price="H",
                    amount="I", clear=list("CDEFGH")),
        sums=dict(小計="I39", 消費税="I40", 合計="I41", 請求額="D13"),
        rate_cell="H40", note=["C44", "C45"], furikomi=None,
        quirk="B1 に『見積書 ESTIMATE』が雛形のまま残っている（実物）"),
    "misoca16": dict(MISOCA_COMMON,
        src=SP / "real14" / "7e1a5f1d-simple.xlsx", sheet="misoca_invoice",
        issuer=dict(name="K6", zip="K7", addr=["K8", "K9"], tel="L10", fax="L11",
                    mail="L12", tanto="L13"),
        to=dict(name="B3", attn="B5"),
        detail=dict(first=16, last=33, name="B", qty="J", unitname="K", price="L",
                    amount="N", clear=list("BCDEFGHIJKLM")),
        sums=dict(小計="L34", 消費税="L35", 合計="L36", 請求額="E12"),
        rate_cell="K35", note=["B39"], furikomi=None),

    # ── マネーフォワード 2 種 ────────────────────────────────────────
    "constr": dict(kind="請求書", vendor="MF", tax_style="mark_column",
        src=SP / "construction_bill.xlsx", sheet="適格請求書（インボイス）",
        issuer=dict(name="F2", zip="F3", addr=["F4"], tel="F5", regno="G6"),
        to=dict(zip="B4", addr="B5", name="B7", attn=None),
        detail=dict(first=17, last=24, date="B", name="C", mark="D", price="E",
                    qty="F", amount="G", clear=list("BCDEF")),
        sums=dict(対象10="E26", 税10="G26", 対象8="E27", 税8="G27",
                  小計="E28", 消費税="G28", 請求額="E14"),
        note=[], furikomi="B34",
        quirk="1 枚目のシートが『説明』／軽減税率は D 列の『※』／"
              "雛形の D22:D24 に全角スペースが入っている"),
    "inv21": dict(kind="請求書", vendor="MF", tax_style="rate_column",
        src=SP / "inv21.xlsx", sheet="インボイス対応請求書",
        issuer=dict(zip="H8", addr=["H9", "H10"], name="H11", regno="H12",
                    tel="H13", mail="H14", tanto="H15"),
        to=dict(zip="B8", addr="B9", addr2="B10", name="B11", attn="B13"),
        detail=dict(first=33, last=45, date="B", code="C", name="D", mark="E",
                    color="F", spec="G", price="H", qty="I", rate="J", amount="K",
                    clear=list("BCDEFGHIJ")),
        sums=dict(小計="K49", 消費税="K50", 合計="K51", 請求額="D22",
                  税10="C51", 対象10="D51", 税8="C52", 対象8="D52"),
        invno="H19", note=["B55", "B56", "B57"], furikomi="H22",
        quirk="請求番号 H19 が裸の大きな数字／消費税が ROUND(x,1)＝小数第 1 位まで／"
              "宛先ブロックと発行者ブロックが同じ形（〒・住所・社名）"),

    # ── spreadoffice 4 種 ───────────────────────────────────────────
    "spread3": dict(kind="請求書", vendor="spreadoffice", tax_style="fixed10",
        src=SP / "vendors" / "spread_3.xlsx", sheet="請求書",
        issuer=dict(name="P3", zip="P5", addr=["P6", "P7", "P8", "P9"]),
        to=dict(zip="C2", addr="C3", addr2="C4", addr3="C5", name="C6"),
        detail=dict(first=22, last=39, code="A", name="D", qty="O", unitname="Q",
                    price="S", amount="V", note="Y", clear=list("ADOQSY")),
        sums=dict(小計="V40", 消費税="V42", 合計="V44", 請求額="T19"),
        kenmei="D19", note=[], furikomi="D12",
        quirk="発行者名と登録番号が P3 の 1 セルに改行同居（実物）"),
    "spread5": dict(kind="請求書", vendor="spreadoffice", tax_style="fixed10",
        src=SP / "vendors" / "spread_5.xlsx", sheet="請求書",
        issuer=dict(name="P3", zip="P5", addr=["P6", "P7", "P8"]),
        to=dict(zip="C2", addr="C3", addr2="C4", name="C5"),
        detail=dict(first=22, last=39, code="A", name="D", qty="O", unitname="Q",
                    price="S", amount="V", note="Y", clear=list("ADOQSY")),
        sums=dict(小計="V40", 消費税="V42", 合計="V44", 請求額="Q19",
                  前回="A19", 入金="F19", 繰越="K19", 今回請求額="W19"),
        note=[], furikomi="D10",
        quirk="繰越請求（前回請求額／入金額／繰越金額／今回請求額）が在る実物"),
    "nouhin": dict(kind="納品書", vendor="spreadoffice", tax_style="fixed10",
        src=SP / "vendors" / "spread_1.xlsx", sheet="納品書",
        issuer=dict(name="P3", regno="P5", zip="P6", addr=["P7", "P8", "P9", "P10"]),
        to=dict(name="C5"),
        detail=dict(first=22, last=39, code="A", name="D", qty="O", unitname="Q",
                    price="S", amount="V", note="Y", clear=list("ADOQSY")),
        sums=dict(小計="V40", 消費税="V42", 合計="V44", 請求額="T19"),
        kenmei="D19", note=[], furikomi=None,
        quirk="★ 請求書ではない（納品書）"),
    "irai": dict(kind="請求依頼書", vendor="spreadoffice", tax_style="fixed10",
        src=SP / "vendors" / "spread_2.xlsx", sheet="請求依頼書",
        issuer=dict(name="AC3", zip="AC5", addr=["AC6", "AC7", "AC8", "AC9"]),
        to=dict(zip="B3", addr="B4", addr2="B5", name="B6", attn="B7"),
        detail=dict(first=18, last=27, code="A", name="E", qty="S", unitname="V",
                    price="Y", amount="AC", clear=list("AESVY")),
        sums=dict(小計="AC28", 消費税="AC29", 合計="AC30", 請求額="F14"),
        kenmei="D8", note=[], furikomi="S11",
        quirk="★ 請求書ではない（請求依頼書）／宛先は『様』で御中が 1 つも無い"),
}
