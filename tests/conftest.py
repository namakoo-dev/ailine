"""全テスト共通の autouse fixture。

★ CI 落ち対応（W10c 追加項目）: normalize_book は basrun.py（sibling repo）+実機
LibreOffice を要する（空マクロ/StructDump で一度開いて保存する正規化パス）。
tests/test_ailine.py の関所の統合テストの一部が、これを個別に monkeypatch し忘れて
いた（実測: 8本）。開発機に basrun が隣接している（例: C:/Dev/basrun）ため気づかず緑の
ままだったが、CI（ailine だけを checkout・basrun.py 無し）では
`SystemExit: basrun.py が見つからない` で落ちる。

★ 個々のテストへ「_find_basrun_path をダミーで存在させる」workaround を1本ずつ足すと、
新しく足すテストが同じ地雷を踏み続ける（今回の再発そのもの）。ここで一箇所に寄せ、
normalize_book の既定を「実機に触らずコピーを返すだけ」にする。個々のテストが自分で
normalize_book/basrun_apply を monkeypatch すれば、そちらが後勝ちでそのまま優先される
（monkeypatch は setattr の順番どおりに効く・teardown は自動）。

★ @pytest.mark.local（例: test_bold_local.py）は実機の basrun/LibreOffice往復を見る
ことが目的のテストなので、この既定の対象から外す（本物の normalize_book を使わせる）。

★ 実運用（本番の ailine.py 実行）の既定値・呼び出し経路は一切変えない（テスト専用の
既定であり、production の normalize_book 定義そのものは無傷）。
"""
import os
import sys
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "src"))

# ★ wheel 化（2026-08-23）: subprocess から `python -m ailine` を叩く検体は子プロセスなので
#   上の sys.path 挿入が届かない。PYTHONPATH で src を渡す ── 実運用では wheel を install
#   すれば不要な、テスト harness 側だけの橋渡し。
_SRC = str(Path(__file__).resolve().parent.parent / "src")
os.environ["PYTHONPATH"] = _SRC + os.pathsep + os.environ.get("PYTHONPATH", "")
import ailine  # noqa: E402

# ★★ 2026-09-23: 本物の ~/.ailine/history.jsonl に、pytest の一時フォルダの冊が 126 行（全件 18 回分・
#   毎回 7 行）書かれていた。出所は 1 か所 ── `tests/test_the_sheets_are_readable.py` の
#   `scope="module"` の fixture。★ 下の切り離しはどれも**関数ごと**の autouse なので、
#   module / session の fixture は**それより先に**走り、本物のホームに書く（切り離しの外）。
#   全件の後に本物のホームの sha256 を照らし合わせて初めて見えた（試験は全部緑だった）。
#   ★ 処方は 2 つ: ① セッションの最初に全体を切り離す（module の fixture もその内側で走る）
#                   ② 終わった時に本物のホームが変わっていたら、全件を赤にする（在っても鳴らない、にしない）
from _home_isolation import home_bound_paths  # noqa: E402

#: 切り離す前の本物のホーム（AILINE_HOME を利用者が指定していればそれ）
_REAL_HOME = Path(ailine.HISTORY_DIR)


def _home_state(root: Path) -> dict:
    """本物のホームの下のファイル（大きさ・更新時刻）。★ 中身は読まない（安い・機密を見ない）。"""
    out = {}
    if root.exists():
        for p in root.rglob("*"):
            try:
                if p.is_file():
                    st = p.stat()
                    out[str(p)] = (st.st_size, st.st_mtime_ns)
            except OSError:
                continue
    return out


#: ★ 起点は conftest を読んだ瞬間に取る（どの試験・fixture よりも前）。
#:   初版は pytest_sessionstart で取っていたが、tests/ の conftest からは呼ばれず、起点が無いまま
#:   「何もしない」で抜けていた ── 変異（切り離しを外す）で本物のホームに 7 行書かれても鳴らなかった。
_REAL_HOME_BEFORE = _home_state(_REAL_HOME)


def _check_the_real_home(session):
    """終わった時に本物のホームが変わっていたら全件を赤にする（下の pytest_sessionfinish から呼ぶ）。

    ★ 初版は自分で `pytest_sessionfinish` を名乗り、**同じファイルの下にある同名の hook に黙って
      上書きされていた**（Python の再定義）── 変異で本物のホームに 7 行書かれても鳴らなかった。
      hook は 1 つに畳み、そこから呼ぶ。
    """
    before = _REAL_HOME_BEFORE
    after = _home_state(_REAL_HOME)
    changed = sorted(k for k in set(before) | set(after) if before.get(k) != after.get(k))
    if changed:
        tr = session.config.pluginmanager.get_plugin("terminalreporter")
        msg = ("✗ 試験が本物のホーム（" + str(_REAL_HOME) + "）を変えた ── 切り離しの外で製品が走った:"
               + "".join(chr(10) + "    " + c for c in changed[:10])
               + chr(10) + "  ★ 同じ時間に ailine を手で使っていたなら、それが出所（その時は流し直す）")
        if tr is not None:
            tr.write_line(msg, red=True)
        else:
            print(msg)
        session.exitstatus = 1


@pytest.fixture(scope="session", autouse=True)
def _isolate_the_whole_session(tmp_path_factory):
    """★ セッションの最初に保存先を全部一時フォルダへ寄せる（module / session の fixture もこの内側）。
    関数ごとの切り離し（下の _guard_real_home_writes）は、その上に後勝ちで重なる。"""
    # ★ ホームの中の**構造を保って**寄せる（isolated_home と同じ）。初版は `root/_session_<名前>` と
    #   ばらばらに置いたので、どれも HISTORY_DIR の下でなくなり、歩き手の isolated_home が
    #   「ホームの下の保存先」を HISTORY_DIR 1 つしか見つけられず、残りをセッションで共用した
    #   ── 前の試験の状態が次の歩きに漏れ、全件の中でだけ導線の歩きが落ちた（単独では緑）。
    from _home_isolation import isolated_home
    root = tmp_path_factory.mktemp("_session_ailine_home")
    with isolated_home(ailine, root):
        yield


@pytest.fixture(autouse=True)
def _default_normalize_book_is_passthrough(request, monkeypatch):
    if request.node.get_closest_marker("local"):
        return   # 実機往復そのものを見たいテストは対象外
    if request.node.get_closest_marker("real_normalize_book"):
        return   # normalize_book 自身の挙動を検証するテストは対象外（pytest.ini 参照）
    monkeypatch.setattr(ailine, "normalize_book", lambda book, workdir, timeout=None: book)


# ★ W10 前提工事②（architect レビュー致命5-3）: 実 ~/.ailine/history.jsonl 等へ
# テストが書いてしまう穴を塞ぐ番人。
#
# 実測（2026-08-22）: 実 history.jsonl 641 行中、failure_kind=語彙外(vocab_miss) 93 件
# のうち 92 件が pytest 由来（book=b.xlsx 等・real 実利用は 1 件のみ）。原因は
# tests/test_golden_transcripts.py の _isolate（HISTORY_FILE/VOCAB_FILE/BACKUP_DIR/
# RUN_LOCK_FILE を tmp_path に寄せる小道具）を、個々のテストが monkeypatch し忘れる
# 「適用漏れ」（例: tests/test_freeform_out_only.py の一部は HISTORY_FILE を一切
# monkeypatch しない）。_isolate は各テストが「思い出して呼ぶ」もので、思い出さなければ
# 実ホームに書く既定のまま――同じ地雷を新しいテストが踏み続ける。
#
# ここで一箇所に寄せ、HISTORY_FILE/VOCAB_FILE/MISCLASS_FILE の既定を autouse で
# tmp_path に強制する。個々のテストが自分の tmp_path へ monkeypatch する分は、
# こちらより後に setattr されるのでそのまま優先される（後勝ち＝無害・_isolate や
# 個別 monkeypatch を削る必要は無い）。
#
# ★ 対象は history/vocab/misclass/aliases の4つ（BACKUP_DIR/RUN_LOCK_FILE はこの番人の
# setattr スコープ外だったが、下の AILINE_HOME 環境変数がそれも含めて根治する）。
# ★ 既存の実 ~/.ailine/history.jsonl 等の掃除はしない（本番データ・触るのは
# Namakoo 決裁）。
#
# ★ 第二波 ①（SEALED-20260823-jisaku-ultra.md 所見⑦の根治）: setattr は同一プロセスにしか
# 効かないため、`ailine.py` を subprocess で別プロセス起動するテスト（14 ファイル）は
# この番人の setattr をすり抜けて実 home に書いていた。env 経由なら subprocess.run が
# 明示 env= を渡さない限り os.environ をそのまま継承するので、そちら側も一箇所で塞げる
# （resolve_home_dir() が呼び出しのたび環境変数を読むため、import 順を問わず効く）。
@pytest.fixture(autouse=True)
def _guard_real_home_writes(monkeypatch, tmp_path):
    monkeypatch.setenv("AILINE_HOME", str(tmp_path / "_guard_ailine_home"))
    monkeypatch.setattr(ailine, "HISTORY_FILE", tmp_path / "_guard_history.jsonl")
    monkeypatch.setattr(ailine, "VOCAB_FILE", tmp_path / "_guard_vocab.json")
    monkeypatch.setattr(ailine, "MISCLASS_FILE", tmp_path / "_guard_misclass.jsonl")
    monkeypatch.setattr(ailine, "ALIASES_FILE", tmp_path / "_guard_aliases.json")
    # ★ 2026-08-23 の追補: AILINE_HOME の setenv は子プロセスにしか効かない
    #   （module 変数は import 時に実 home で束縛済み）。同一プロセスの run.lock /
    #   backups も明示的に tmp へ寄せる ── 実測で並行 pytest の相互妨害が続いていた。
    monkeypatch.setattr(ailine, "RUN_LOCK_FILE", tmp_path / "_guard_run.lock", raising=False)
    monkeypatch.setattr(ailine, "BACKUP_DIR", tmp_path / "_guard_backups", raising=False)
    # ★★ 2026-09-23: 上の手書きの名簿には ATTRIBUTES_FILE と HISTORY_DIR が抜けていた（片配線）。
    #   残りは**モジュールから導いて**寄せる ── 保存先が増えても自動で守られる。
    from _home_isolation import home_bound_paths
    for name, _p in home_bound_paths(ailine).items():
        monkeypatch.setattr(ailine, name, tmp_path / f"_guard_{name.lower()}")


@pytest.fixture(autouse=True)
def _no_real_ollama(request, monkeypatch):
    """★ W10 便C2 検分（2026-08-22 夜）: 「CI には ollama が居ない」をローカルで再現する番人。
       mock されていない経路が実 ollama を呼ぶと、ローカルでは黙って緑（ollama が答える）・
       CI では赤/非決定になる ── 「居るから見えない」を居ない側に倒して全穴をその場で鳴らす。
       実機を使う検体は @pytest.mark.local で免除（従来どおり -m local で別走）。"""
    if "local" in request.keywords or "ollama_internals" in request.keywords:
        yield
        return
    def _boom(*a, **k):
        raise AssertionError(
            "実 ollama を呼んだ（mock されていない経路 ── CI には存在しない・conftest._no_real_ollama）")
    monkeypatch.setattr(ailine, "ollama_generate_json", _boom, raising=False)
    monkeypatch.setattr(ailine, "ollama_generate", _boom, raising=False)
    yield


@pytest.fixture(autouse=True)
def _release_run_lock_after_each_test():
    """★ 2026-08-24: 実行ロックを OS の排他ロックに移した。持ち主はプロセス単位なので、
    ある検体が解放し忘れると**次の検体が壊れる**（実測: 単独では通るのに並べると落ちた）。
    後始末を検体の善意に任せず、ここで必ず外す。
    """
    yield
    import ailine as _al
    handle = getattr(_al, '_RUN_LOCK_HANDLE', None)
    if handle is not None:
        try:
            _al.release_run_lock(handle[1])
        except Exception:
            _al._RUN_LOCK_HANDLE = None


# ★ 実機を起こした走行は、終わりに LibreOffice を落とす -----------------------

_LOCAL_RAN = []


@pytest.fixture(autouse=True)
def _remember_if_the_real_machine_was_used(request):
    """この走行で実機（@pytest.mark.local）が 1 本でも走ったかを覚える。"""
    if request.node.get_closest_marker("local"):
        _LOCAL_RAN.append(1)
    yield


def pytest_sessionfinish(session, exitstatus):
    """★★ 2026-09-06: 実機を起こした走行は、終わりに待ち受けを落とす。

    ★ 実測: `pytest -m local` が終わっても `soffice` が 2 つ（約 270MB）残り続ける。
      成功した走行でも毎回残る ── 押すたびに増える。手で PID を見て落としていた。
    ★ 落とし方は `basrun stop` に委譲する ── UNO で**接続先だけ** terminate するので、
      人が GUI で開いている LibreOffice は巻き込まない（別プロファイル・別ポート）。
      名前一括の kill は使わない（[[feedback_taskkill_kills_mcp]] の教訓）。
    ★ 実機を 1 本も走らせていない回は**何もしない**（起こしてもいないものを止めない）。
    ★ ここでの失敗は無視する ── 後始末が走行の合否を変えてはいけない。
    ★★ 2026-09-23: 本物のホームの検算（_check_the_real_home）もここから呼ぶ ── hook は 1 つだけ
      （2 つ書くと後の方が前の方を黙って消す。実際に消していた）。

    ★ 断り書き: 居残りが「実機テストの大量失敗」を起こしたという証拠は**無い**
      （2026-09-06 に再現を試みて失敗した ── 切られた後の待ち受けは健康だった）。
      これは**資源の後始末**として正しいから入れる。原因不明の失敗への処方ではない。
    """
    _check_the_real_home(session)
    if not _LOCAL_RAN:
        return
    try:
        import subprocess
        import sys as _sys
        import ailine as _al
        subprocess.run([_sys.executable, str(_al.basrun_path()), "stop"],
                       capture_output=True, timeout=60)
    except Exception:
        pass
