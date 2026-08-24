"""cli_render — C8: ailine.py に散った `print` のうち、複数経路で同じ形を手書きしていた
   ものを「文字列を組み立てる純関数」に寄せる。呼び側は組み立てた文字列/行リストを
   `print()` するだけにする（claim.py が『✓ 機械検証済み』でやった流儀を広げる）。

★ なぜ（C8 ブリーフより）: queue してある挙動変更2件（--dry プレビューの verified=True
誤認・差分番人と ✓ の矛盾警告）はどちらも表示に触る。表示が ailine.py 内に散っている
うちに直すと同じ場所を二度触ることになるため、先に層を分ける。

★ ここに移した範囲（と、あえて移さなかった範囲）: 実測で **複数の呼び出し箇所が同じ形を
手書きしていた** ものだけを対象にした（生成 .bas のコード表示ブロック3箇所・「続けるには
以下のいずれかを指定して」の再試行案内3箇所・「× 中止した」3箇所・restore/undo の
バックアップ一覧2箇所・run 見出し行3箇所・vocab の一覧/追加結果）。cmd_run_dsl や
cmd_run_plan の中に残る他の print（`？ {質問}` `× {エラー}` 等）は、その場限りの
1回書き（他に同形の呼び出し箇所が無い）で、関数に括り出しても呼ばれる場所が1つのままの
薄いラッパーにしかならないため、今回は動かしていない（「呼ばれていない/1箇所しか呼ばない
関数を作らない」という C8 ブリーフの縛り）。

★★ 出力は1バイトも変えない（純リファクタ）。retry-options のフラグ列は元の手書き文言が
グループごとに右端を揃えていた実測の桁（fidelity ゲート15桁・overwrite ゲート13桁・
freeform ゲート18桁）を、`max(len(flag)) + 2` で再現する形にした（ハードコードではなく
計算にした理由: 3ゲートで桁が違う＝手書きの揃え幅そのものに規則性があったため、その規則を
関数に持たせるほうが「桁だけ後で増える4つ目のゲート」が来ても自動で揃う）。
"""
from __future__ import annotations

from pathlib import Path

# --- 生成 .bas コード表示ブロック（単発 DSL / 自由生成の試行ループ / 複合計画の語彙外段） --

_CODE_BLOCK_FOOTER = "──────────────────────────────────────────"


def render_code_block(header: str, code: str, step_prefix: str = "") -> list:
    """生成コードの表示ブロック（見出し行・コード本体・区切り線）を行のリストにする。
       header は呼び出し側が組んだ見出し行そのもの（先頭の `\\n` の有無や文言は経路ごとに
       違う＝呼び出し側の責務のまま）。footer の区切り線だけが3経路で共通（実測で確認済み）。"""
    return [header, code, f"{step_prefix}{_CODE_BLOCK_FOOTER}"]


# --- 「続けるには以下のいずれかを指定して」再試行案内（忠実度/上書き/自由生成の3ゲート） --

def render_retry_options(step_prefix: str, options: list) -> list:
    """options: [(flag, 説明), ...]。フラグ列は同ブロック内の最長 flag + 2桁の空白幅で
       右側の説明を揃える（3ゲートとも手書きの揃え幅がこの計算と一致することを実測で
       確認済み — docs/behavior-corpus 側の挙動そのものは変えていない）。"""
    width = max(len(flag) for flag, _desc in options) + 2
    lines = [f"{step_prefix}この処理を続けるには、以下のいずれかを指定して再実行してください:"]
    for flag, desc in options:
        lines.append(f"{step_prefix}  {flag:<{width}}{desc}")
    return lines


def render_aborted(step_prefix: str = "") -> str:
    """対話で拒否された/非対話で確認できなかった時の中止行（3ゲートで同一文言）。"""
    return f"{step_prefix}× 中止した"


# --- run 見出し行（単発 DSL / 自由生成 / 複合計画で共通の骨格） ------------------------

def render_run_header(label: str, model: str, book_name: str) -> str:
    """「■ ailine（〜）  model=...  book=...」の骨格。label だけが経路ごとに違う
       （例: "DSL 経路" / "AI が直接作成・機械保証なし" / "複合計画・N 段"）。"""
    return f"■ ailine（{label}）  model={model}  book={book_name}"


# --- K-1: 語彙外に落ちる瞬間の通知（生成に入る前・理由/費用/次の手を1ブロックで言う） -------
#   ★ 通知だけ。同意の門(y/N)は作らない（K-2として意図的に保留 — operator の盲検査定
#   2026-08-19: 非対応と文書に明記された操作を頼んだら、説明なしに63秒の自由生成が3回走り、
#   最後に依頼と無関係なコードを見せられて止まった。橋の実験の根拠2点を踏まえる:
#   コンパイラ構成＝照合が先・実行は後（照合の結果を即言う）／行政法＝断りは理由の提示を
#   伴う正式な手続き（黙って省略しない）。門は別（発火頻度が高い門は無視が習慣化して死ぬ、
#   という設計判断で K-2 は意図的に保留）。

def freeform_notice_reason(op: str, about: str = "") -> str:
    """通知の1行目（理由）。経路で言い分ける:
       - OUT_OF_VOCAB: モデルが照合した結果「この語彙には無い」と明示的に答えた経路
         （about があれば「（何についての依頼か）」を添える）。
       - それ以外(FREEFORM＝語彙外の op・必須 slot 欠落・JSON 不正・API 不通などの
         退避先すべて): 翻訳がそもそも DSL の形（op+args）にならなかった経路。"""
    if op == "OUT_OF_VOCAB":
        suffix = f"（{about}）" if about else ""
        return f"この依頼{suffix}は、頼める操作の一覧に照合できませんでした。"
    return "この依頼は、翻訳が頼める操作の形になりませんでした。"


# ★ freeform 最終決定（DESIGN-20260821-multifile.md「freeform 最終決定」節・
#   Namakoo 2026-08-21 19:37「廃止しよう」で確定）: 単発の語彙外は生成に入らず即座に断る。
#   旧 render_freeform_notice（K-1・生成へ進む前提の通知）はここで廃止 ── 使っていた
#   cmd_run_freeform 自体が消えたため（git が墓場・復活条件は設計書に凍結）。
#   複合計画側の render_freeform_notice_compact（下）は生成が残る経路なので変えない。

def render_vocab_miss_refusal(about: str = "", sunset_notice: bool = False,
                               translate_error: bool = False) -> list:
    """単発の語彙外（FREEFORM/OUT_OF_VOCAB）の断り。既存の CLARIFY 系（`？` 接頭・
       「（頼める操作の一覧: ailine ops）」の1文）に文体をそろえる。3要素は必須:
       理由・vocab_miss を記録する開示・次の手（ops/言い換え/照合への導線）。
       about があれば OUT_OF_VOCAB が名指しした対象を理由行に添える（FREEFORM は空）。
       sunset_notice: --allow-freeform を受け取った場合だけ廃止告知を1行足す
       （自由生成そのものは受理しない ── 断りの中身は変えない）。
       ★ W10 前提工事①: translate_error=True（ollama 不通/JSON 不正/空応答で
       translate_task 自体が退避した経路）は理由行を差し替える ──「頼める操作の
       一覧に照合できませんでした」は**照合を試みた**ことを含意し、翻訳が
       そもそも走らなかった/形にならなかったこの経路では嘘になるため。"""
    suffix = f"（{about}）" if about else ""
    if translate_error:
        lines = ["？ 翻訳に失敗した（ollama に接続できない等）。要望として記録します。"]
    else:
        lines = [f"？ この依頼{suffix}は、頼める操作の一覧に照合できませんでした。要望として記録します。"]
    if sunset_notice:
        lines.append("自由生成は廃止しました（理由: 機械検証できない操作は行わない方針）。")
    lines.append("  （頼める操作の一覧: ailine ops）")
    lines.append("  言い換えるか、2 冊の突き合わせなら: ailine run 入金.xlsx 請求.xlsx \"…\"")
    return lines


def render_freeform_notice_compact(reason: str, step_prefix: str = "") -> str:
    """複合計画の語彙外段（run_freeform_plan_step）向け。段の文脈に合わせて1行に畳む
       （3要素は保つ: 理由・機械保証なし・次の手。★ ✓ は使わない・上記参照）。"""
    return (f"{step_prefix}{reason} AI が直接生成します（機械保証なし・適用の確認は出ません。"
            "次の手: `ailine ops` / 依頼を言い換える）。")


# --- restore/undo のバックアップ一覧・復元結果（cmd_restore と cmd_undo が手書きで重複） --

def render_backup_list(book_name: str, backups: list, shelved: int = 0) -> list:
    """`ailine restore --list` / `ailine undo --list` の一覧表示。backups は新しい順。
       ★ W11: shelved は「undo が取った復元前の退避」の件数。0 でなければ 1 行だけ添える
       （遡りには数えないが**捨ててはいない**ので、undo をやり直したい人に在り処を示す）。"""
    if not backups:
        return [f"{book_name} のバックアップは無い"]
    lines = [f"{book_name} のバックアップ（{len(backups)} 世代・新しい順）:"]
    lines.extend(p.name for p in backups)
    if shelved:
        lines.append(f"（このほかに undo が取った復元前の退避が {shelved} 件"
                     f"・backups/<名前空間>/undo/ 内・遡りには数えない）")
    return lines


def render_restore_done(book_name: str, used_name: str, remaining: int | None = None) -> str:
    """復元完了行。remaining（undo のみ、まだ戻せる回数）が渡されたときだけ注記を足す
       （restore は remaining=None のまま＝従来どおり注記なし）。"""
    suffix = f"（あと {remaining} 回戻せます）" if remaining is not None else ""
    return f"✓ {book_name} を {used_name} から復元した{suffix}"


# --- 用語集（vocab）コマンドの表示 -----------------------------------------------------

def render_vocab_add_result(ok: bool, msg: str) -> str:
    return ("✓ " if ok else "× ") + msg


def render_vocab_listing(vocab: dict, vocab_file: Path) -> list:
    """`ailine vocab list`。空なら登録方法の案内1行だけ返す。"""
    if not vocab:
        return [f"（用語集は空。{vocab_file} に登録するか `ailine vocab add <語> <値>` で追加）"]
    lines = [f"用語集（{vocab_file}・{len(vocab)}件）:"]
    lines.extend(f"  {term} = {vocab[term]:g}" for term in sorted(vocab))
    return lines


# --- W10 便A: 別名ストア（言い回し → op 名）の一覧表示 --------------------------------

def render_alias_listing(aliases: dict, order: list, aliases_file: Path) -> list:
    """`ailine alias list`。空なら登録方法の案内1行だけ返す。★ vocab の一覧はキーの
       アルファベット順（値の更新しかない=順序に意味が無い）だが、こちらは登録順
       （undo が「直近の登録」を取り消す対象を人が確認できるように、末尾が最新）。"""
    if not aliases:
        return [f"（別名ストアは空。{aliases_file} に登録するか"
                f" `ailine alias add <言い回し> <OP>` で追加）"]
    lines = [f"別名ストア（{aliases_file}・{len(aliases)}件・登録順）:"]
    lines.extend(f"  {phrase} → {aliases[phrase]}" for phrase in order if phrase in aliases)
    return lines


# --- 対応操作の一覧（★ 査定 2 本が独立に「無い」と指摘した唯一のもの） ------------------

def render_ops_table(op_meta: dict, op_schema: dict, confirm_fields: dict) -> list:
    """「こう頼めばこれができる」の一覧を**登録簿から生成**する。

    ★ なぜ生成か（2026-08-16 の盲検査定 2 本より）: 二人とも「対応操作の一覧が無い」を
    MISSING の筆頭に挙げた。README 368 行の中で**何を頼めるのかが分からない**ため、
    語彙外の依頼で質問ループに入り「普通の購入検討者ならここで評価を終える」と書かれた。
    ★ 手書きの表は必ずずれる（この repo は索引のずれを何度も踏んでいる）。
    op_meta / op_schema / confirm_fields から作れば、**操作を足した日に表も増える**。

    引数で登録簿を受け取るのは ailine_core → ailine の逆流を避けるため（移植可能性の番人）。
    """
    order, seen = [], set()
    for meta in op_meta.values():          # 宣言順をそのまま使う（並べ替えない＝出力が安定）
        if meta["category"] not in seen:
            seen.add(meta["category"])
            order.append(meta["category"])
    lines = ["ailine に頼めること（この表は登録簿から自動生成しています）", ""]
    for category in order:
        lines.append(f"■ {category}")
        for op, meta in op_meta.items():
            if meta["category"] != category:
                continue
            says = "／".join(meta["synonyms"])
            need = _needed_info(op, op_schema, confirm_fields)
            # ★ 第二波 ⑤: 行の左端に英字 op 名を出す（README「対応する操作名は
            #   ailine ops の左端の英字」・alias --help「op 名（例: SORT）」の誘導先を
            #   実在させる。手書きでなく登録簿の op（op_meta のキー）そのもの）。
            lines.append(f"  {op}  {meta['label']}    こう書く: {says}")
            if need:
                lines.append(f"      必要な情報: {need}")
        lines.append("")
    # ★ freeform 最終決定 (DESIGN-20260821-multifile.md・2026-08-21): K-1 の約束文
    #   「AI の直接生成を試します」は単発経路ではもう嘘になった（生成せず断る）。
    #   約束は実装に合わせる（同じ理由で K-1 が旧文を直したときと同じ原則）。
    #   複合計画は語彙外の段だけ生成が残る（run_freeform_plan_step は今回変えていない）ので
    #   その例外を1文添える。
    lines.append("※ ここに無い依頼は、頼める操作の一覧に照合できないため生成せず断ります（要望として記録します）。")
    lines.append("※ 複合的な依頼の一部だけが語彙外なら、その段だけ AI が直接生成することがあります（機械保証なし）。")
    lines.append("※ 一覧に無い依頼は聞き返します。言い換えても通らないときは未対応です。")
    return lines


def _needed_info(op: str, op_schema: dict, confirm_fields: dict) -> str:
    """必須 slot を日本語ラベルに直す。ラベルは確認行の登録簿から引く（新しい語を作らない）。"""
    labels = {slot: label for label, slot, _fmt in confirm_fields.get(op, ())}
    return "・".join(labels.get(slot, slot) for slot in op_schema.get(op, ()))


# --- `ailine scan` の人間向け出力（M1読み・DESIGN-20260821-multifile.md §2骨） ------------

def render_scan_report(folder_label: str, result: dict) -> list:
    """分母つき報告（「N ファイル中 M 照合できた」）・失敗は名指し+理由・並べ替えは開示する。
       ★ operator 盲検9回目 CONFUSING①: 全員照合できた場合、従来は分母の1行だけで
       「どのファイルが」照合できたのかが見えなかった（README の「列は揃っているかを
       分母つきで報告」の約束に対し、分子の中身が不透明）。★ ファイルごとに必ず1行出す形に
       直した ── ⚠ の連打（憲法の禁）は避けたまま、取れたファイルも「{名前}: 取れた」で
       名指しする（並べ替え/シートのフォールバックはその1行に畳んで足す）。
       ★ データ行数は含めない ── evaluate_file の戻り値に無い情報を、この報告のためだけに
       evaluate 側へ足すのは今回の範囲外（挙動の本体は変えない・表示側だけの修正）。"""
    files = result["files"]
    matched = sum(1 for f in files if f["status"] == "取れた")
    lines = [f"■ ailine scan  folder={folder_label}"]
    lines.append(f"基準: {result['base']}" if result["base"] else "基準: 見つかりません（読める .xlsx が無い）")
    lines.append(f"{result['denominator']} ファイル中 {matched} 照合できた")
    excluded = result["excluded"]
    lines.extend(render_excluded_lines(excluded))
    for f in files:
        lines.append(_render_scan_file_line(f))
    return lines


def _render_scan_file_line(f: dict) -> str:
    """scan のファイル1件分の行（★ 必ず1行・情報が複数あれば同じ行に畳む）。"""
    if f["status"] == "取れなかった":
        return f"  ⚠ {f['name']}: 取れなかった（{f['reason']}）"
    notes = []
    if f.get("reordered"):
        notes.append("並べ替え")
    fb = f.get("sheet_fallback")
    if fb:
        notes.append(f"シート『{fb['wanted']}』が無いので1枚目『{fb['used']}』を使用")
    suffix = f"（{'・'.join(notes)}）" if notes else ""
    return f"  {f['name']}: 取れた{suffix}"


# --- `ailine stack` / `ailine verify` の人間向け出力 -------------------------------------
# ★ ここも他の render_* と同じ流儀: 渡すのはプリミティブ型だけの dict（dataclass をそのまま
#   渡さない）── cli_render.py は他の ailine_core module を import しない、という既存の
#   自己完結を保つ（呼び出し側の ailine.py が dataclass → dict へ整形してから渡す）。

def _fmt_num(v) -> str:
    """数値表示: 整数値は小数点なしで（650.0 でなく 650）。"""
    try:
        f = float(v)
    except (TypeError, ValueError):
        return str(v)
    return str(int(f)) if f.is_integer() else str(f)


def render_excluded_lines(excluded: dict) -> list:
    """フォルダから外したものの開示。**scan / stack / run<folder> が同じ文を使う。**

    ★ なぜ切り出したか（2026-08-24）: 同じ規則を 3 箇所が別々に書いていたので、
      scan だけが .csv を開示し、`run <folder>` は除外を丸ごと捨て、stack は一部だけ、
      という**片配線**になっていた（盲検で 2 者が独立に指摘）。
      分母は入力側から作る ── そして**開示も 1 箇所から出す**。
    """
    excluded = excluded or {}
    lines = []
    if excluded.get("temp"):
        lines.append(f"対象外: 一時ファイル {excluded['temp']} 件（~$ で除外）")
    if excluded.get("subdirs"):
        lines.append(f"対象外: サブフォルダ {excluded['subdirs']} 件（中は見ていません）")
    if excluded.get("csv"):
        lines.append(f"対象外: .csv {excluded['csv']} 件（1本ずつなら `ailine csv` で扱ってください）")
    if excluded.get("other_format"):
        names = list(excluded.get("other_format_names") or [])
        head = "・".join(names[:3])
        more = f" ほか {len(names) - 3} 件" if len(names) > 3 else ""
        detail = f"（{head}{more}）" if head else ""
        lines.append(f"対象外: 読めない形式 {excluded['other_format']} 件{detail}"
                     f" ── .xlsx に保存し直すと扱えます")
    return lines

def render_stack_report(folder_label: str, out_label: str, result: dict) -> list:
    """M1書き `ailine stack` の人間向け報告。分母つき + 除外の名指し + Σ の両側表示。
       ★ ⚠ は異常のあるファイルだけ（合計行の不一致・A列/used range の食い違い）。"""
    lines = [f"■ ailine stack  folder={folder_label}  out={out_label}"]
    self_excluded = result.get("self_excluded")
    if self_excluded:
        # ★ architect 致命2: 複数ファイルがありうる（V6 が out 一致に限らず広がったため）。
        names = "、".join(f"『{n}』" for n in self_excluded)
        lines.append(f"（自分の出力 {names} を入力から除外しました）")
    if result.get("collision_notice"):
        lines.append(f"（{result['collision_notice']}）")
    lines.extend(render_excluded_lines(result.get("excluded")))
    lines.append(f"{result['denominator']} ファイル中 {result['stacked_files']} 積んだ")
    for f in result.get("skipped", ()):
        lines.append(f"  ⚠ {f['name']}: 積めなかった（{f['reason']}）")
    for f in result.get("files", ()):
        if f.get("reordered"):
            lines.append(f"  {f['name']}: 取れた（並べ替え）")
    for f in result.get("sheet_fallbacks", ()):
        lines.append(f"  {f['name']}: シート『{f['wanted']}』が無いので1枚目『{f['used']}』を使いました")
    for entry in result.get("excluded_detail", ()):
        rows_txt = "、".join(f"{r['row']}行目" for r in entry["rows"])
        lines.append(f"  {entry['name']}: 合計行を{len(entry['rows'])}件除外（{rows_txt}）")
    for entry in result.get("mismatches", ()):
        for m in entry["rows"]:
            lines.append(f"  ⚠ {entry['name']}: 合計行({m['row']}行目) の値 "
                         f"{_fmt_num(m['excluded_value'])} ≠ 明細の和 {_fmt_num(m['adopted_sum'])}")
    if result.get("unverified_cols"):
        # ★ 2026-08-24: 「Σ の行が出ない」で伝えるのをやめ、**出ない理由を書く**。
        #   出ないものは読めない（今日ずっと出ている「無いことを信号にする」家系）。
        names = "・".join(f"『{h}』" for h in result["unverified_cols"])
        lines.append(f"  ★ 疑わしい: {names} は検算していません"
                     "（数式のまま値が入っていないため Σ で確かめられません"
                     "── 一度 Excel/LibreOffice で開いて保存すると値が入ります）")
    if result.get("header_drop_warning"):
        # ★ 2026-08-24: 見出し行の選び方で列が丸ごと落ちる時の名指し。
        #   col_a_warnings とは意味が違うので混ぜない（別の警告を同じ器に入れない）。
        lines.append(f"  {result['header_drop_warning']}")
    for w in result.get("col_a_warnings", ()):
        # ★ 2026-08-24: 文言を買い手の言葉に直した。盲検の査定者に
        #   「『根1』は開発者用語で読めない」「正常系で 100% 鳴る」と名指しされた
        #   （原因は自分で正しく除外した合計行）。何が起きたかと、どうすればよいかを言う。
        lines.append(f"  ⚠ {w['name']}: 1列目から数えると {w['col_a']} 行ですが、"
                     f"表の範囲は {w['used_range']} 行あります"
                     "（1列目に空欄があるか、合計行を除いた差です）")
    # ★ 第二の独立検出器（operator 盲検7度目 修正2）: 列解決に依存しない語のトリップワイヤ。
    #   除外はしない（検出のみ）── 検出器1が沈黙しても黙って倍額にはならない、の開示。
    for w in result.get("total_word_warnings", ()):
        lines.append(f"  ⚠ {w['file']} の{w['row']}行目に合計語『{w['word']}』を含む行が"
                     "積まれています（除外していません・確認してください）")
    lines.append(f"出力データ行数: {result['rows_written']}")
    for col, both in result.get("sums", {}).items():
        lines.append(f"Σ{col}: 元 {_fmt_num(both['source'])} / 出力 {_fmt_num(both['output'])}")
    if result.get("rebuilt_own_output"):
        lines.append(f"（前回の縦積み出力『{out_label}』を作り直しました）")
    return lines


def render_verify_report(out_label: str, folder_label: str, result: dict) -> list:
    """`ailine verify` の人間向け報告。合格: 両側の数字を並べる／不合格: 列名 + 両側の数字。
       ★ デモ撮影のリハで発覚（2026-08-21）: 最初の不一致（Σ）で打ち切ると、帰属検算が
       名指しできるはずの「どの行（元どのファイルの何行目）がいくつ→いくつ」を報告し
       損なう ── 憲法1（誘導）+ 一括検出。`result["mismatches"]`（複数・verify.py が
       行数を除く全種を集めて返す）を全部並べる。行数不一致だけは別（それ単独で致命
       なので従来どおり単独の1行で止める・`mismatch`（単数）で後方互換）。"""
    lines = [f"■ ailine verify  out={out_label}  folder={folder_label}"]
    mismatch = result.get("mismatch")
    if mismatch and mismatch["kind"] == "row_count":
        lines.append(f"⚠ 行数が一致しません: 元 {mismatch['source']} / 出力 {mismatch['output']}")
        return lines
    lines.append(f"行数: 元 {result['row_count']['source']} / 出力 {result['row_count']['output']}")
    for col, both in result.get("sums", {}).items():
        lines.append(f"Σ{col}: 元 {_fmt_num(both['source'])} / 出力 {_fmt_num(both['output'])}")
    mismatches = result.get("mismatches")
    if mismatches is None:   # ★ 後方互換: 複数形を持たない呼び出し元（無いはずだが fail closed）
        mismatches = [mismatch] if mismatch else []
    for m in mismatches:
        kind = m["kind"]
        if kind == "sum":
            lines.append(f"⚠ {m['column']} の合計が一致しません: "
                         f"元 {_fmt_num(m['source'])} / 出力 {_fmt_num(m['output'])}")
        elif kind == "attribution":
            # ★ review3#3: 集計は合っていても帰属（どの行がどのファイルの何行目か）が嘘。
            lines.append(f"⚠ 帰属が一致しません: {m['file']} の {m['src_row']}行目 "
                         f"列『{m['column']}』 元 {_fmt_num(m['source'])} / "
                         f"出力 {_fmt_num(m['output'])}")
        elif kind == "missing_source":
            # ★ 2026-08-24: 元フォルダに在るのに出力の出所列に一度も現れない冊。
            #   旧版は元側の一覧を**出力自身の出所列**から作っていたので、この不一致は
            #   原理的に発生しえなかった（＝冊が丸ごと落ちても exit 0 だった）。
            lines.append(f"⚠ {m.get('name')} の行が出力に 1 行もありません"
                         "（フォルダには在るのに積まれていない ── 見出しの綴りや"
                         "列の欠けを確認してください）")
        elif kind == "row_count":
            # ★ 2026-08-24 第三波: 分母を入力側から作るようにした時に生まれた kind。
            #   枝の追加を忘れていて、生の dict がそのまま人に出ていた（下の
            #   フォールバックが設計どおり拾った ── 黙る不合格にはならなかった）。
            #   ★ 原因（missing_source）が既に名指し済みなら、それが差の説明だと繋ぐ。
            gap = abs((m.get("source") or 0) - (m.get("output") or 0))
            named = [x.get("name") for x in mismatches
                     if x.get("kind") == "missing_source" and x.get("name")]
            if named:
                lines.append(f"→ 行数の差 {gap} 行は、上で名指しした {len(named)} 冊が"
                             "積まれていないためです")
            else:
                lines.append(f"⚠ 行数が合いません: 元 {m.get('source')} / "
                             f"出力 {m.get('output')}（差 {gap} 行・原因は特定できていません）")
        elif kind == "total_word":
            # ★ operator 盲検7度目 修正2（第二の独立検出器）: 再演検分の直し ── この分岐が
            #   無いと exit 5 なのに理由が1行も出ない「黙る不合格」になっていた（憲法1違反）。
            file_label = m.get("file") or "(ファイル不明)"
            lines.append(f"⚠ {file_label} の{m['row']}行目に合計語『{m['word']}』を含む行が"
                         "あります（除外していません ── 合計行なら元を確認してください）")
        else:
            # ★ 型で塞ぐフォールバック: 将来 kind が増えた時に、この分岐の追加漏れで
            #   「mismatches は非空なのに ⚠ が1行も出ない」黙る不合格を起こさない
            #   （生データを出すので、少なくとも exit 非0 の理由がゼロにはならない）。
            lines.append(f"⚠ 不明な種類の不一致（kind={kind}）: {m}")
    return lines


def render_verify_match_report(out_label: str, a_label: str, b_label: str, result: dict) -> list:
    """`ailine verify <出力> <元A> <元B>`（M3 照合出力）の人間向け報告。
       合格: 両側の Σ を並べるだけ／不合格: キーごとに何がいくつ→いくつ、を全所見並べる
       （一括検出・最初の1件で止めない・憲法1: 修正箇所への誘導）。"""
    lines = [f"■ ailine verify（照合）  out={out_label}  A={a_label}  B={b_label}"]
    sums = result.get("sums") or {}
    if sums:
        lines.append(f"Σ A: {_fmt_num(sums.get('A', 0))} / Σ B: {_fmt_num(sums.get('B', 0))}")
    for m in result.get("mismatches", ()):
        kind = m["kind"]
        if kind == "count":
            lines.append(f"⚠ {m['key']}: {m['side']}側件数が一致しません"
                         f"（独立再集計 {m['expected']} / 出力 {m['written']}）")
        elif kind == "sum":
            lines.append(f"⚠ {m['key']}: {m['side']}側合計が一致しません"
                         f"（独立再集計 {_fmt_num(m['expected'])} / 出力 {_fmt_num(m['written'])}）")
        elif kind == "diff":
            lines.append(f"⚠ {m['key']}: 差額が一致しません"
                         f"（算出 {_fmt_num(m['expected'])} / 出力 {_fmt_num(m['written'])}）")
        elif kind == "missing_key":
            lines.append(f"⚠ {m['key']}: 元帳にあるキーが照合の出力に見当たりません")
        elif kind == "extra_key":
            lines.append(f"⚠ {m['key']}: 出力にあるキーが元帳の独立再集計に見当たりません"
                         "（捏造の可能性）")
    return lines
