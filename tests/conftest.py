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
