# 型の門（`tests/pyright_gate.json`）── 何を入れ、何を入れなかったか

★ pyright の設定ファイルは JSON で、**注記を書けない**（書くと毎回「不明な設定」と言われる）。
  だから理由はここに書く。設定を変える人は、必ずこの文書も一緒に変えること。

## 約束

**ここに在る規則は、入れた時点ですべて 0 件。**
0 でない規則は入れない ── 初日から赤い門は「在っても鳴らない」になり、そのうち誰も見なくなる。

走らせ方（版を固定する。固定しないと、道具が上がった日に門が勝手に変わる）:

    uvx pyright@1.1.414 -p tests/pyright_gate.json

`src/` だけを見る。`tests/` は見ない（試験は製品ではない）。

## 入れた規則

| 規則 | 何を捕まえるか |
|---|---|
| `reportPossiblyUnbound` | その経路では代入されていないかもしれない名前 |
| `reportAttributeAccessIssue` | 無い属性・無いモジュール属性 |
| `reportUndefinedVariable` | 知らない名前（ruff の F821 と**重ねて**持つ ── 片方が黙る日のため） |
| `reportInvalidStringEscapeSequence` | 文字列の中の無効なエスケープ |
| `reportSelfClsParameterName` | `self` / `cls` の書き間違い |

★ 入れるために 2026-09-22 に直したもの（どれも実体が在った）:

- `exit_environment` に `NoReturn` を宣言 ── **戻らないこと**を言わないと、型検査器は
  この先を「未定義かもしれない」と数える。**誤報 67 件 → 0 件**。
- `accounts_core.py` の `headers` を `head_row` と対で初期化 ── 1 つの事実を 2 変数が
  別々に持っていた（片配線の予約）。
- `urllib.error` を名指しで輸入 ── `urllib.request` が内側で輸入するのに**頼っていた**。
- `_path_key` の docstring を raw にする ── **パスの取り違えを書いた文章そのもの**が
  無効なエスケープを含んでいた（Python 3.12 は `SyntaxWarning` を出す）。
- `_CsvEvaluation` の `encoding` / `parsed` を `object` から本当の型へ ──
  注釈が「何でも入る」と言っていたので、**番人の目を注釈が塞いでいた**。
- `render_fn` の分岐 ── 同じ判断（`warning_count > 0`）を 2 箇所に書いていたのを 1 回に畳んだ。

## 入れなかった規則（2026-09-22 の実測。数は `src/` のみ）

| 規則 | 件数 | 入れない理由 |
|---|---|---|
| `reportOptionalMemberAccess` | 80 | openpyxl の値が `Optional` なことに由来。潰すには 20,000 行への注釈工事が要る |
| `reportCallIssue` | 16 | 同上（`ws.cell(...)` などの多重定義） |
| `reportOptionalSubscript` ほか Optional 系 | 11 | 同上 |
| `reportMissingImports` | 4 | すべて `pdfplumber`。**わざと入れていない**任意の依存で、素の環境では必ず解決しない |
| `reportMissingModuleSource` | 31 | 型スタブだけ在る外部パッケージ。製品の欠陥ではない |

★ 注釈工事に踏み込まないのは値付けの判断であって、「安全だから」ではない。
  ここは**測って見送った**ものの台帳で、実需が来たら再測定の候補になる。

## 名指しで外した 3 箇所（`# pyright: ignore[...]`）

道具の型情報の限界で、文脈上は起きえないもの。数を緩めるためではないので、
**増やすときは必ず理由を行に書く**こと。

- `sys.stdout.reconfigure` … スタブは `TextIO` と言うが実体は `TextIOWrapper`。既に `try` で囲ってある
- `csv_quarantine.py` の `cell.value =` 2 箇所 … 自分で作った新規シートなので結合セルは在りえない
