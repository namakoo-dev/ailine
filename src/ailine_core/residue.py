"""residue — W10 便C2 S5: もしかして提案の残差検出（部分対応の罠への機械の防壁）。

★ なぜ在るか（test_suggest_flow.py 冒頭の決裁参照）: 判定器(judge_ops_via_llm)の
自己申告に頼ると「対応外の部分を黙ったまま提案する」害が起きる。7B に「一部だけ対応で
残りは対応外」と自己申告させる実験は 5/6 で素通りした（部分対応の罠に自分では気づかない）。
指示は意図、保証は機械 ── 提案する側が「この操作に反映される部分」を機械で確定し、
反映されない残りを名指しする。この判定は LLM を一切使わない（+0ms/+0依存）。

★ 判定方法（形態素解析はしない・置換による span 除去）: 依頼文から
  ①数字 ②解決済み args の文字列値（列名・ラベル・引用値など） ③op の照合語彙
  (pool_phrases・label/synonyms/match_phrases) を文字列として取り除き、残った文字列から
  「内容語らしい連続語」（漢字/カタカナ/半角英数字の2文字以上の連続）を拾う。

★ ひらがなは対象にしない ── ailine.py の `_raw_target_not_embedded_in_task`
（単位B・列名照合の断片ガード）が「ひらがな/カタカナは日本語の語境界」として扱うのと
同じ観察を裏返しに使う: 助詞・活用語尾はひらがなで書かれるため、内容語の候補にそもそも
含めなければ、専用の助詞リストを持たなくても「消費」したのと同じ効果になる。

★ オオカミ少年回避が最優先（きれいな依頼に誤って残差行を出すと信頼を失う ──
Namakoo 決裁「迷ったら出さない側に倒す」）: 呼び出し側(ailine.py)は pool_phrases に
op の label/synonyms/match_phrases 全部（suggest_ops の照合プールと同じもの）を渡すこと。
広く消費させるほど、残差行の誤発火（きれいな依頼への誤爆）は減る。

★ 置き場所: ailine_core/（sum_identity.py と同じ理由）。ailine を import しない
（移植可能性の番人 test_line_budget.py が機械で守る）。
"""
from __future__ import annotations

import re

# 内容語らしい連続語: 漢字(2文字以上) / カタカナ(長音符込み・2文字以上) / 半角英数字(2文字以上)。
# ひらがなは含めない（docstring 参照 ── 助詞・活用語尾は拾わないことで「消費」を兼ねる）。
_CONTENT_RUN_RE = re.compile(
    r"[㐀-䶿一-鿿豈-﫿]{2,}|[ァ-ヶー]{2,}|[A-Za-z]{2,}")
_NUMBER_RE = re.compile(r"\d+(?:\.\d+)?")


_COND_RE_CACHE: dict = {}


def stated_as_condition(task: str, column: str, value: str) -> bool:
    """依頼文が「<列名> が <値>」の形で**条件として書いている**か。

    ★★ 2026-09-06（自作 review が致命 2 件・重大 1 件として拾った）: 2 組目の条件を
      「残差に残った語 ∩ 他の列の実在値」だけで決めていたため、**条件として書かれて
      いない語**まで条件に採っていた:

          「田中さんのように、所属が営業の行のメモに○を付けて」
            → 『田中』が氏名列の実在値なので「かつ氏名が田中」が勝手に付き、
              3 行に当たるはずが **1 行**に縮んだ（警告なし・✓ のまま）

    ★ 今日の否定の直しと**同じ形**だ ── 文に在るかではなく、**何に付いているか**を見る。
      （[[negation.reading]] が値／列名の別で 3 通りに割ったのと対）。

    ★ 括弧は 1 つだけ挟めるようにする（「担当が『主任』の行」も条件として書かれている）。
    """
    text, col, val = task or "", str(column or ""), str(value or "")
    if not text or not col or not val:
        return False
    key = (col, val)
    rx = _COND_RE_CACHE.get(key)
    if rx is None:
        rx = re.compile(re.escape(col) + r"[がはのを]?[「『\"']?" + re.escape(val))
        _COND_RE_CACHE[key] = rx
    return bool(rx.search(text))


#: 見出し名の後ろに付く**構造の語**（★ 開いた接頭辞一致にはしない ── `原価率` のような
#: 別の語まで拾うと、誤爆の側へ倒れる）。
_STRUCTURE_SUFFIXES = ("列", "行", "欄", "の列", "の行", "の欄")


def _names_a_header(word: str, headers) -> str | None:
    """その語が見出しを指しているなら、**見出しの側**の名前を返す（無ければ None）。

    ★★ 2026-09-07（UX 検品が拾った・昨日入れた関所が黙っていた真因）:
      残差は語を**最長で**切り出すので、「原価**列**の右隣に」からは『原価列』が出る。
      見出しは『原価』なので完全一致に失敗し、**関所が鳴らなかった**。
      実際の事故: 「原価列の右隣に備考列を追加して」→ 備考が**末尾**に入り、
      解釈行は「依頼文に位置の指定が無いため」と**嘘をついて** `✓` を出していた。
    ★ 「原価列」と書くのは利用者にとってごく自然な言い方 ── そこに穴が開いていた。
    """
    if word in headers:
        return word
    for h in headers:
        if word.startswith(h) and word[len(h):] in _STRUCTURE_SUFFIXES:
            return h
    return None


def unaccounted_request_words(task: str, declaration: str, pool_phrases, headers) -> list:
    """依頼に在って**実行した解釈のどこにも出ていない列名**を返す（無ければ空）。

    ★★ なぜ在るか（2026-09-06・実測で決めた）: 事後条件が見るのは「宣言 vs 実体」だけで、
      **依頼から落ちた分は原理的に見えない**。実際 3179 件の本物の走行を調べたところ、
      条件つき書換を *行追加* と誤読した 28 件は**事後条件が pass**と言っていた。
      同じ依頼の前後で自然実験になっており、直った後は黙る:

          09-04 06:26〜09-05 03:34  op=ADD_ROW    解釈に『所属』が出ない → 鳴る
          09-05 06:01 以降          op=SET_WHERE  解釈に『所属』が出る   → 黙る

    ★ 3179 件で鳴ったのは 28 件（0.88%）で、**そのすべてが本物の欠陥**（誤爆 0）。
      再現は `bench/residue_gate_on_history.py`。

    ★ 宣言は**解釈行**（人に見せる方）を渡すこと ── 引数の dict ではない。
      「ナットの行を削除して」は引数だと `row=3` に化けて『ナット』が消えるが、
      解釈行には「位置の根拠:『ナット』の行＝3行目」として残る。引数で測った版は
      誤爆 21% だった（`bench/residue_gate_probe.py --values`）。

    ★ 絞りを **実表の列名**に限るのが要（この絞りが無いと履歴で 49% が鳴る）。

    ★ 捕まえない形（測って分かっている・広げようとして却下した）:
      ・落ちたのが**値だけ**の時（列名は宣言に在る）
      ・**否定の反転**（「営業以外」を「営業」と読む ── 語は 1 つも落ちない）
      ・**宣言の語に部分一致して消える**形（`w not in decl` は部分文字列で見るため、
        依頼の『利益』は宣言の『利益率』に食われる ── 2026-09-08 に実測）。
        新しい列の名前が依頼の語を含む回は、この関所は原理的に黙る。
        ★ 式の食い違いは `ailine_core/arith.py` が隣で受け持つ。
      どれも同じ履歴に実例が在り、実際に黙っていた。**開示して持つ**。
    """
    left = find_unconsumed_words(task, {}, pool_phrases)
    decl = declaration or ""
    heads = {str(h) for h in (headers or ()) if h}
    out = []
    for w in left:
        name = _names_a_header(w, heads)
        if name and name not in decl and w not in decl:
            out.append(name)
    return list(dict.fromkeys(out))


def find_unconsumed_words(task: str, resolved_args: dict, pool_phrases) -> list:
    """依頼文 task のうち、resolved_args の文字列値・pool_phrases・数字のどれにも
       消費されなかった内容語を、出現順・重複除去で返す（無ければ空リスト）。
       ★ span 除去方式（文字列としてこの文中に現れるかだけを見る・形態素解析はしない）。
       消費の候補が広いほど安全（オオカミ少年を避ける側に倒れる）ので、resolved_args は
       検証済みの解決値（列名・ラベル等）を、pool_phrases は op の照合語彙を広く渡すこと。"""
    if not task:
        return []
    remaining = _NUMBER_RE.sub(" ", task)
    # ★ 第二波 ④（本家 bug_008）: dict の反復順（＝呼び出し側が args を組んだ key の順）で
    #   はなく、値の**長さ降順**で消費する。「商品」「商品コード」のように片方がもう片方を
    #   部分文字列として含む場合、短い方を先に消費すると長い方の残骸（「コード」）が
    #   偽の残差として漏れる（pool_phrases 側は元々この順でやっていた・args 側に同じ規律を
    #   足すだけ＝pool と対称）。
    # ★★ 2026-09-06（自作 review が致命として拾った）: **リストの中身も消費する**。
    #   否定（cmp=nin）では `cond_value` が `['営業']` のようなリストになるが、
    #   ここが文字列しか見ていなかったため、**否定した値が未消費のまま残差に残り**、
    #   他の列に同じ値が在ると「頼んでいない 2 組目の条件」として採られていた
    #   （実測: 「所属が営業でない行のメモに『済』」→ 担当=営業 が勝手に付き 2 行 → 1 行）。
    #   ★ この欠陥は同じ日に**測定器の中で見つけて直していた**（bench/residue_gate_probe.py）。
    #     測定器だけ直して**本体に持ち帰らなかった** ── 直す時は必ず両方を見る。
    #   ★ 呼び出し側で平らにしない（4 箇所ある）── 消費の意味はここ 1 つが持つ。
    def _flat(v):
        if isinstance(v, str):
            return [v] if v else []
        if isinstance(v, (list, tuple, set)):
            return [x for x in v if isinstance(x, str) and x]
        return []

    for v in sorted({s for v in (resolved_args or {}).values() for s in _flat(v)},
                    key=len, reverse=True):
        if v in remaining:
            remaining = remaining.replace(v, " ")
    for phrase in sorted({p for p in (pool_phrases or ()) if p}, key=len, reverse=True):
        if phrase in remaining:
            remaining = remaining.replace(phrase, " ")
    seen = []
    for m in _CONTENT_RUN_RE.finditer(remaining):
        w = m.group(0)
        if w not in seen:
            seen.append(w)
    return seen


def render_residue_note(words: list) -> str | None:
    """残差語があれば注記1行、無ければ None（呼び出し側は None なら何も印字しない
       ── きれいな依頼に沈黙させるのが既定・S5 の対照検体）。"""
    if not words:
        return None
    joined = "・".join(words)
    return f"（『{joined}』などの部分はこの操作に反映されません）"
