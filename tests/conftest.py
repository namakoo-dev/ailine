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
import sys
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
import ailine  # noqa: E402


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
# ★ 対象は history/vocab/misclass の3つだけ（BACKUP_DIR/RUN_LOCK_FILE はこの番人の
# スコープ外 ── 実測で漏れが確認されたのはこの3つ）。
# ★ 既存の実 ~/.ailine/history.jsonl 等の掃除はしない（本番データ・触るのは
# Namakoo 決裁）。
@pytest.fixture(autouse=True)
def _guard_real_home_writes(monkeypatch, tmp_path):
    monkeypatch.setattr(ailine, "HISTORY_FILE", tmp_path / "_guard_history.jsonl")
    monkeypatch.setattr(ailine, "VOCAB_FILE", tmp_path / "_guard_vocab.json")
    monkeypatch.setattr(ailine, "MISCLASS_FILE", tmp_path / "_guard_misclass.jsonl")


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
