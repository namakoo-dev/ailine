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

from ailine_core import field_record
from ailine_core import route   # ★ 断る前に「別のコマンドで出来る」を見る

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
                               translate_error: bool = False, task: str = "") -> list:
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
    # ★ 2026-08-24（初回体験の盲検・致命①）: 翻訳がそもそも走らなかった経路
    #   （ollama 不通・モデル未取得）に、語彙外と同じ「言い換えるか」を出していた。
    #   直し方は言い回しではないので、**間違った方向へ人を送っていた**。
    #   原因が環境なら、環境の直し方だけを言う（ops も言い換えも出さない）。
    if translate_error:
        return [
            "× AI エンジン（ollama）に依頼を渡せませんでした。",
            "  → `ailine doctor` で、何が足りないかを名指しします",
            "     （よくある原因: ollama が起動していない / モデルが未取得）",
        ]
    # ★★ 2026-09-08（盲検の検品が挙げた摩擦）: `run` では出来なくても、**別のコマンドで
    #   出来る**ものが在る（「PDF にして」→ `ailine export-pdf` は実在する）。
    #   その回に「要望として記録します」と言うのは、持っている物を持っていないと言う形。
    #   ★ 判定は ailine_core/route.py に 1 つだけ置き、ここは材料を渡すだけ。
    if (_go := route.command_that_can_do_this(task)):
        _cmd, _note = _go
        return [f"？ この依頼{suffix}は `ailine run` では受け取れません。",
                f"  → `{_cmd}` が同じことをします（{_note}）",
                "  （頼める操作の一覧: ailine ops）"]
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

def render_legacy_note(legacy: list) -> list:
    """旧フラット領域の同名世代を**開示する** 1 行（混ぜないが、隠しもしない）。

    ★ 2026-08-25（復元の致命①）: 直下のファイルはフォルダ情報を持たないので、
      同名なら別フォルダの他人のものでも遡り履歴に入っていた。混ぜるのをやめた以上、
      在ることは言う ── **実パスを出す**（手で戻せる形で渡す）。
    ★ ⚠ で始めない: これは「今の操作が確かめられていない」ではなく、
      「使わなかったものが在る」という案内。✓ を降ろす理由にはしない。
    """
    if not legacy:
        return []
    return [f"（旧領域に同名の世代が {len(legacy)} 件ありますが、どのフォルダのものか"
            f"分からないため遡りには混ぜません: {legacy[0].parent}）"]


def render_backup_list(book_name: str, backups: list, shelved: int = 0,
                        shelf_dir=None, legacy: list | None = None) -> list:
    """`ailine restore --list` / `ailine undo --list` の一覧表示。backups は新しい順。
       ★ W11: shelved は「undo が取った復元前の退避」の件数。0 でなければ 1 行だけ添える
       （遡りには数えないが**捨ててはいない**ので、undo をやり直したい人に在り処を示す）。
       ★ 2026-08-25（復元の重大⑤）: 在り処は `backups/<名前空間>/undo/` という
       **プレースホルダをそのまま印字**していた（実体は sha1 先頭 8 桁）。
       案内は打てる形で出す ── shelf_dir をもらって実パスを書く。"""
    lines = []
    if not backups:
        lines.append(f"{book_name} のバックアップは無い")
    else:
        lines.append(f"{book_name} のバックアップ（{len(backups)} 世代・新しい順）:")
        lines.extend(p.name for p in backups)
    if shelved:
        where = str(shelf_dir) if shelf_dir else "backups/<名前空間>/undo/"
        lines.append(f"（このほかに undo が取った復元前の退避が {shelved} 件"
                     f"・{where} 内・遡りには数えない）")
    lines.extend(render_legacy_note(legacy))
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


#: 「後から確かめる呼び方」を出す出力の種類（鍵は書き手の印 creator）。
#: ★★ 2026-09-13: `ailine verify` が stack/extract/forms/split を支えているのに、
#:   **出力を作った画面には呼び方が一言も出ていなかった**（grep 0 件）── ① と同じ形で、
#:   しかも独立の検算を足した当日に同じ穴を開けた。**知られない検算は無い検算**。
#: ★ csv は独立の検算が無いので**書かない**（案内すると嘘になる）。
#: ★ どれも引数の並びは同じ（出力 → 元）なので、文言は 1 本で足りる。
#: ★ 2026-09-13（需要③）: `ailine accounts` は独立の検算（`verify_accounts`）を持って
#:   出たので足す ── 持っていないものを足すと案内が嘘になり、持っているのに足さないと
#:   「知られない検算は無い検算」になる。
VERIFY_HINT_KINDS = ("ailine stack", "ailine extract", "ailine forms", "ailine split",
                     "ailine accounts")


def _quoted(label: str) -> str:
    """空白を含むパスは引用する（そのまま貼って動く形にする）。"""
    return f'"{label}"' if " " in label or "\u3000" in label else label


def verify_hint(creator: str, out_label: str, source_label: str, extra: str = "") -> list:
    """後から確かめる呼び方を 1 行で返す。支えていない種類には**書かない**。

    ★ 文言はここだけ ── 報告の側で書き写すと、次に種類が増えた時に片配線になる。
    ★ `extra` は**使い手が渡した条件をそのまま運ぶ**ためのもの（split の `--amount`）。
      これが無いと、案内した検算が**本人が実行した検算より弱く**なる ── 勧める側が
      黙って手を抜くことになるので、渡された条件は案内にも出す。
    """
    if creator not in VERIFY_HINT_KINDS:
        return []
    tail = f" {extra}" if extra else ""
    return [f"あとから確かめる: ailine verify {_quoted(out_label)} "
            f"{_quoted(source_label)}{tail}"]


# --- 対応操作の一覧（★ 査定 2 本が独立に「無い」と指摘した唯一のもの） ------------------

def render_folder_routes(subcommands, multi) -> list:
    """★ 2026-08-24（第三波 S6）: 複数ファイルの入口（棚卸し・縦積み・2冊照合）に
    **たどり着く道が無かった** ── ops の表にも README にも出ておらず、知らなければ
    一生使われない機能だった。表と同じ作法で **argparse の登録簿から生成**する
    （手書きの一覧は必ずずれる）。

    subcommands: [(名前, help 文字列, 引数の形), ...]。
    multi: 複数ファイルの入口の名前（★ 正は呼び出し側の宣言 `ailine.ROUTE_KIND`）。

    ★★ 2026-09-12: ここは `wanted = ("scan", "stack", "verify")` という**手書きの白名簿**
      だった。すぐ上に「手書きの一覧は必ずずれる」と書いてあるのに、`forms`（請求書から
      項目を集める）と `split`（担当者ごとに分けて配る）を出荷してもここに入らず、
      **ops の一覧にも README にも 1 度も出なかった** ── 完成しているのに買い手が
      見つけられない。宣言を 1 箇所に集め、登録簿との**等号**を番人が縛る形にした。
    """
    # ★ この 2 行だけは生成できない（run は位置引数の数で分岐するので argparse の
    #   登録簿には「フォルダ 1 個」「ブック 2 冊」の区別が無い）。手書きだと明示する。
    spelled_out = ['  ailine run <folder> "<依頼>"   '
                   "フォルダ内の全ブックから条件で抜き出す",
                   '  ailine run <a.xlsx> <b.xlsx> "<依頼>"   '
                   "2 冊をキーで突き合わせて差額を出す"]
    # ★ 手書き行が覆っている名前は、生成側から外す ── **その名前は手書き行自身から取る**
    #   （二つ目の白名簿を作らない）。これを怠ると run が 3 行に増え、生成された
    #   `<book> <task>`（1 冊の形）が「複数ファイルの入口」として並んでしまう。
    covered = {line.split()[1] for line in spelled_out}
    rows = [(n, h, u) for n, h, u in subcommands if n in multi and n not in covered]
    if not rows and not spelled_out:
        return []
    # ★ 引数の形も argparse に言わせる ── ここを雛形（"<フォルダ>" 決め打ち）で書いたら
    #   即座にずれた（verify はフォルダを取らない）。自分で「手書きはずれる」と書いた
    #   直後にずらしたので、実物として残す。
    lines = ["", "── 複数のファイルをまとめて扱う ──"]
    for name, help_text, usage in rows:
        lines.append(f"  ailine {name} {usage}".rstrip() + f"   {help_text}")
    return lines + spelled_out


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
    # ★ 第三波 S1: 自分の出力を外したことを、stack と同じ文言で言う（黙って減らさない）。
    self_excluded = result.get("self_excluded") or []
    if self_excluded:
        names = "、".join(f"『{n}』" for n in self_excluded)
        lines.append(f"（自分の出力 {names} を入力から除外しました）")
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
    # ★ 基準と違う行を見出しとして読んだら言う ── 黙って別の行を読まない（2026-09-11）。
    for f in result.get("header_row_fallbacks", ()):
        lines.append(f"  {f['name']}: 見出しは {f['used_row']} 行目にありました"
                     f"（基準は {f['base_row']} 行目）")
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
    # ★ 2026-08-24（第三波 S2）: 見出しに out=<path> と**宣言**しておいて、積むものが
    #   0 件だとファイルを作らずに exit 0 で終わっていた（宣言と実体の食い違い ──
    #   今日の「判定には三項が要る」の穴）。作らなかったなら、そう言う。
    # ★ 2026-08-24（飾りの生存表の実測）: stack は**新しいブックを作る** op なので、
    #   元の飾り（罫線・塗り・条件付き書式・数式・図形…）は 1 つも持ち越さない。
    #   仕様どおりだが、どこにも書いていなかった ── 12 冊の請求書を積んで素の表が
    #   出てきた買い手は「壊れた」と読む。作った時に言う。
    # ★ 2026-08-24 の訂正: 初版は「元の書式…は持ち越しません」と書いたが、**嘘だった**
    #   ── 数値書式（#,##0 も yyyy-mm-dd も）は運んでいる（設計文書にも「日付書式の
    #   引き継ぎ」と書いてあった）。測らずに一般化した。運ぶものと運ばないものを正確に言う。
    if result.get("rows_written"):
        lines.append("（縦積みは新しいブックを作ります ── 値と数値書式は運びますが、"
                     "罫線・塗り・数式・図形・条件付き書式は持ち越しません）")
        # ★ コメントとハイパーリンクは**飾りでなく中身**（人が打った情報）。
        #   持ち越さないのは同じでも、消えたことを言わないと気づけない。
        dropped = result.get("dropped_notes") or {}
        if dropped:
            parts = "・".join(f"{k} {n} 件" for k, n in sorted(dropped.items()))
            lines.append(f"  ⚠ 元のセルにあった{parts}は運んでいません"
                         "（値でないため）。元のブックで確認してください")
    if not result.get("file_written", True):
        lines.append("（ファイルは作っていません ── 積む対象が 0 件のため。"
                     "上の除外の内訳を確認してください）")
    n_excluded = sum(len(e.get("rows", ())) for e in result.get("excluded_detail", ()))
    for col, both in result.get("sums", {}).items():
        lines.append(sum_line(col, both, n_excluded))
    if result.get("rebuilt_own_output"):
        lines.append(f"（前回の縦積み出力『{out_label}』を作り直しました）")
    if result.get("file_written", True):
        lines += verify_hint("ailine stack", out_label, folder_label)
    return lines


def sum_line(col: str, both: dict, excluded_rows: int = 0) -> str:
    """Σ の 1 行。★ 2026-08-24（第三波 S4）: 『元』が**合計行を除いた後**の値なのに、
    そう読めなかった（元 100 / 出力 100 が一致していても、除外が間違っていれば
    間違い同士で合う ── 数字は合っているのに事実は間違っている、という形）。
    意味は変えず、何を『元』と呼んだかを書く。★ 実装は 1 つ（呼ぶのは 2 経路）。"""
    line = f"Σ{col}: 元 {_fmt_num(both['source'])} / 出力 {_fmt_num(both['output'])}"
    if excluded_rows:
        line += f"（『元』は合計行 {excluded_rows} 件を除いた後の値です）"
    return line


def render_independent_verify_report(label: str, out_label: str, source_label: str,
                                    result: dict) -> list:
    """後からの独立検算の報告（③・2026-09-12）── 分けた冊と帳票の一覧で**同じ器**を使う。

    ★ 2 つ書くと片配線になる（この repo の実測: ⚠ を出す経路が 2 本あって片方だけ直した）。

    ★ 出すのは**分母つきの事実と名指し**だけ ── 成績のバーは置かない。
    ★ 破れは件数で終わらせず「どの行のどれが」まで出す（件数だけでは人は動けない）。
    ★ 金額を測っていない回は、測っていないと**書く**（黙って合格に混ぜない）。
    """
    lines = [f"■ ailine verify（{label}）  out={out_label}  元={source_label}"]
    for name, value in (result.get("facts") or {}).items():
        lines.append(f"  {name}: {value}")
    breaks = result.get("breaks") or []
    if not breaks:
        lines.append(f"✓ 破れはありません（上の分母で測りました）── {label}")
        return lines
    lines.append(f"⚠ 破れ {len(breaks)} 件:")
    for kind, detail in breaks:
        lines.append(f"   {kind}: {detail}")
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
        lines.append(sum_line(col, both, result.get("excluded_rows", 0)))
    # ★★ 2026-08-26（複数ファイルの盲検・致命①⑨）: 検算は書き手と**同じ関数**で
    #   「合計行」を落としている。両方が同じ間違いをすれば一致してしまう（恒真）。
    #   実測: 区切りの空行がある表で、3 列すべて埋まった売上 1,000 円が消えて exit 0。
    #   ★ 判定は変えない（正しい合計行を持つ表を全部不合格にしないため）。
    #     代わりに**落とした行を名指しする** ── 人が「それは売上だ」と気づける形にする。
    for u in result.get("unbacked_exclusions") or []:
        lines.append(f"⚠ {u['file']} の {u['row']}行目を『合計行』として除外しています"
                     "（検算はこの判断を裏取りしていません ── 本物のデータ行なら、"
                     "その金額は上の合計に入っていません）")
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
        elif kind == "unreadable_source":
            # ★ 2026-08-26（致命⑥）: 壊れた .xlsx を生の traceback で落とさず、名指しする。
            lines.append(f"⚠ {m['name']} を読めませんでした（{m['source']}）"
                         " ── この冊の行は検算の分母に入っていません")
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

def render_forms_report(folder_label: str, out_label: str, result: dict) -> list:
    """`ailine forms` の人向け報告（★ 分母つき・名指し・区分の内訳）。

    ★ 買い手の信用条件をそのまま運ぶ: 何冊中何冊を読めたか（分母）、読めなかった冊の
      名指し、そして**空欄がいくつでその理由がいくつか**。
    ★ 成績のバーは置かない（置くと『割』を『単』へ格下げする方へ手が動く・§0g の決裁）。
    """
    lines = [f"■ ailine forms（帳票の一覧）  folder={folder_label}"]
    if result.get("file_written"):
        lines.append(f"出力先: {out_label}")
    exc = result.get("excluded") or {}
    if isinstance(exc, dict):
        # ★ 2026-09-12: 初版は dict の**鍵**を「（対象外）temp」と並べていた（毎回 5 行の雑音）。
        #   件数が 0 でないものだけ、扱えなかった形式は**名前で**出す。
        for key, label in (("temp", "一時ファイル"), ("subdirs", "サブフォルダ"), ("csv", "CSV")):
            if exc.get(key):
                lines.append(f"  （対象外）{label} {exc[key]} 件")
        for name in exc.get("other_format_names") or ():
            lines.append(f"  （対象外・扱えない形式）{name}")
    else:
        for f in exc:
            lines.append(f"  （対象外）{f}")
    for n in result.get("self_excluded", ()) or ():
        lines.append(f"  （自分の出力 『{n}』 を入力から除外しました）")
    lines.append(f"{result['denominator']} ファイル中 {result['collected']} 冊を読みました")
    for f in result.get("unreadable", ()) or ():
        lines.append(f"  ⚠ {f['name']}: {f['reason']}")
    grades = result.get("grades") or {}
    if grades:
        # ★ 区分の語も意味も field_record が持つ（ここで書き写さない・AST の番人が縛る）。
        got = "／".join(f"{g} {grades[g]}" for g in field_record.GRADE_ORDER if g in grades)
        lines.append(f"項目の区分: {got}（{field_record.grade_legend()}）")
    blanks = result.get("blanks", 0)
    if blanks:
        lines.append(f"空欄 {blanks} 件（理由つき {result.get('blanks_with_reason', 0)} 件）"
                     "── 理由は『検分』シートに 1 件ずつ出しています")
    sus = result.get("suspicions") or []
    if sus:
        kinds: dict = {}
        for s in sus:
            kinds[s["種類"]] = kinds.get(s["種類"], 0) + 1
        lines.append(f"束で見て怪しいもの {len(sus)} 件（"
                     + "／".join(f"{k} {n}" for k, n in kinds.items())
                     + "）── 『束の所見』シートに 1 件ずつ出しています")
        for s in sus[:8]:
            lines.append(f"  ⚠ {s['種類']}: {'／'.join(s['冊'])}")
    if result.get("file_written"):
        lines.append("（一覧は新しいブックです ── 元の請求書は 1 バイトも変えていません）")
        lines += verify_hint("ailine forms", out_label, folder_label)
    return lines


def render_split_report(book_label: str, out_label: str, result: dict) -> list:
    """`ailine split`（担当者別に分けて配る）の人向け報告（2026-09-12・需要⑤）。

    ★ 分母つき（全体の行が部分にどう散ったか）・名指し（空欄・ゆれ・複数担当・分けない行は
      行番号で言う）・そして**証明**（部分の和＝全体を、書いた出力から読み戻して確かめた）。
    ★ 成績のバーは置かない（forms と同じ線 ── 置くと「割合を上げる」方へ手が動く）。
    ★ ✓ は証明が通った時だけ ── ここで手書きの ✓ を作らない（`proof.ok` が唯一の根拠）。
    """
    lines = [f"■ ailine split（担当者別に分けて配る）  book={book_label}  out={out_label}"]
    if result.get("sheet"):
        lines.append(f"シート『{result['sheet']}』の {result.get('header_row')} 行目を"
                     "見出しとして読みました")
    for name in result.get("other_sheets", ()) or ():
        lines.append(f"  （見たのはこのシートだけです ── 同じ冊に『{name}』もあります）")
    if result.get("refused"):
        lines.append(f"× 分けていません: {result['refused']}")
        lines.append("（1 冊も作っていません ── 表から決まらないことは、こちらで決めません）")
        return lines

    parts = result.get("parts") or {}
    proof = result.get("proof") or {}
    rows = proof.get("rows") or {}
    lines.append(f"{rows.get('whole')} 行のうち {rows.get('parts')} 行を {len(parts)} 冊に"
                 f"配りました（空欄 {rows.get('blank')}／複数担当 {rows.get('multi')}／"
                 f"分けない行 {rows.get('excluded')}）")
    for value, part in parts.items():
        amount = part.get("amount")
        tail = f"／{_fmt_num(amount)}" if amount is not None else ""
        lines.append(f"  ・{value}: {len(part.get('rows') or ())} 行{tail}"
                     f"  → {part.get('file')}")
    if proof.get("ok"):
        amount = proof.get("amount") or {}
        money = ""
        if amount.get("counted"):
            money = (f"／金額 {_fmt_num(amount.get('whole'))} ＝ "
                     f"{_fmt_num(amount.get('parts'))}＋{_fmt_num(amount.get('blank'))}"
                     f"＋{_fmt_num(amount.get('multi'))}")
        lines.append(f"✓ 部分の和 ＝ 全体（行 {rows.get('whole')} ＝ {rows.get('parts')}＋"
                     f"{rows.get('blank')}＋{rows.get('multi')}＋{rows.get('excluded')}"
                     f"{money}）── 配った冊を開き直して数えた結果です")
    for line in proof.get("broken", ()) or ():
        lines.append(f"× 証明が破れました: {line}")
    if not result.get("amount"):
        lines.append("金額は数えていません（`--amount <金額の見出し>` を付けると、"
                     "行数と同じやり方で金額の和も証明します）")
    blank = result.get("blank") or []
    if blank:
        lines.append(f"⚠ 空欄 {len(blank)} 行（{_rows_label(blank)}）"
                     "── どの冊にも入れていません（空欄は誤配より安いので、こちらで決めません）")
    for pair in result.get("lookalike", ()) or ():
        lines.append(f"⚠ 表記ゆれ: 『{pair[0]}』／『{pair[1]}』── 空白や中黒を無視すると"
                     "同じ文字です。別の冊のままにしています（同じ人だと決めるのは人の仕事です）")
    for row_num, value in result.get("multi", ()) or ():
        lines.append(f"⚠ 複数担当: {row_num} 行目『{value}』── どちらの冊に入れるかは"
                     "表からは決まらないので、どの冊にも入れていません")
    unparsed = result.get("unparsed") or []
    if unparsed:
        lines.append(f"⚠ 金額が文字の行 {len(unparsed)} 行（{_rows_label(unparsed)}）"
                     "── 和に数えていません（読み替えていません）")
    excluded = result.get("excluded") or []
    if excluded:
        lines.append(f"（分けない行 {len(excluded)} 行: {_rows_label(excluded)}"
                     " ── 合計・小計・空の行。誰の冊にも入れていません）")
    if result.get("files_written"):
        lines.append(f"（配った冊 {len(parts)} 件 ＋ 検分 1 件は新しいブックです"
                     " ── 元の表は 1 バイトも変えていません）")
        amount = result.get("amount")
        lines += verify_hint("ailine split", out_label, book_label,
                             f"--amount {_quoted(str(amount))}" if amount else "")
    return lines


def _rows_label(rows, show: int = 8) -> str:
    """行番号の並びを「4, 7, 9 行目」の形に（多い時は件数で締める）。"""
    nums = [str(r) for r in rows]
    if len(nums) > show:
        return f"{', '.join(nums[:show])} 行目 ほか {len(nums) - show} 行"
    return f"{', '.join(nums)} 行目"


def render_accounts_report(today_label: str, out_label: str, result: dict) -> list:
    """`ailine accounts`（経費の勘定科目を先例から引く）の人向け報告（需要③・2026-09-13）。

    ★ 分母つき（候補を出す行が何行で、そのうち何行に候補が出たか）・名指し（空欄の理由・
      表記ゆれ・触らない行は行番号で言う）。
    ★ 成績のバーは置かない（forms / split と同じ線 ── 置くと「割合を上げる」方へ手が動き、
      割 を 単 へ格下げしたくなる）。
    ★ 区分の語と意味は `field_record` から引く（ここで書き写さない ── AST の番人が縛る）。
    """
    lines = [f"■ ailine accounts（経費の勘定科目を先例から引く）  今回={today_label}"]
    for name in result.get("past") or ():
        lines.append(f"  過去の仕訳: {name}")
    for f in result.get("unreadable") or ():
        lines.append(f"  ⚠ {f}")
    if result.get("header_row"):
        lines.append(f"見出しは {result['header_row']} 行目として読みました")
    elif result.get("refused") is None:
        lines.append("見出しの行はありません（列の位置で読みました ── 弥生の形）")
    if result.get("encoding"):
        tail = ("・★ UTF-8 でも cp932 でも復号できる冊です（UTF-8 として読みました ── "
                "文字化けが見えたら元のソフトの書き出し設定を確かめてください）"
                if result.get("ambiguous") else "")
        lines.append(f"文字コード: {result['encoding']}{tail}")
    if result.get("refused"):
        lines.append(f"× 候補を出していません: {result['refused']}")
        lines.append("（冊は作っていません ── 表から決まらないことは、こちらで決めません）")
        return lines

    rows = result.get("rows") or {}
    valued = [r for r, v in rows.items() if (v or {}).get("account")]
    lines.append(f"候補を出す行 {len(rows)} 行のうち {len(valued)} 行に科目の候補が出ました"
                 f"（触らない行 {len(result.get('untouched') or ())}／"
                 f"過去の行 {result.get('past_rows', 0)}・うち借方が埋まった行 "
                 f"{result.get('past_precedents', 0)}）")
    grades = result.get("grades") or {}
    if grades:
        got = "／".join(f"{g} {grades[g]}" for g in field_record.GRADE_ORDER if g in grades)
        lines.append(f"区分: {got}（{field_record.grade_legend()}）")
    blanks = [r for r, v in rows.items() if not (v or {}).get("account")]
    if blanks:
        lines.append(f"⚠ 空欄 {len(blanks)} 行（{_rows_label(sorted(blanks, key=int))}）"
                     "── 理由は『検分』シートに 1 行ずつ出しています"
                     "（空欄は誤値より安いので、こちらで決めません）")
    for key, first, second in result.get("lookalike") or ():
        lines.append(f"⚠ 表記ゆれ: {key} 『{first}』／『{second}』── 畳むと同じ文字です。"
                     "別の鍵のままにしています（同じものだと決めるのは人の仕事です）")
    for name in result.get("ambiguous_books") or ():
        lines.append(f"⚠ {name}: UTF-8 でも cp932 でも復号できる冊です"
                     "（UTF-8 として読みました ── 文字化けが見えたら書き出し設定を確かめて）")
    for note in result.get("notes") or ():
        lines.append(f"（{note}）")
    if result.get("原本が変わった"):
        # ★ 読むだけの約束が破れた（入力の指紋が前後で違う）── 黙って合格に混ぜない。
        lines.append("⚠ 入力の指紋が前後で違います: "
                     + "／".join(result.get("changed_inputs") or ())
                     + "（読むだけのはずの入力が変わりました ── 出した候補は信じないでください）")
    if result.get("file_written"):
        lines.append(f"出力先: {out_label}")
        lines.append("（候補の冊は新しいブックです ── 今回の仕訳も過去の仕訳も "
                     "1 バイトも変えていません）")
        past = " ".join(_quoted(str(p)) for p in result.get("past_paths") or ())
        lines += verify_hint("ailine accounts", out_label, today_label, past)
    else:
        lines.append("（冊は作っていません）")
    return lines
