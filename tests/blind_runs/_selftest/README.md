# _selftest ── 道具の自己検査であって、盲検ではない

★★ これは `scripts/blind_run.py` が**動くこと**を確かめるために 2026-09-18 に
作ったミニの走行（冊 1 つ・コマンド 4 つ）。**買い手役の盲検ではない。**

★ 名前を `000` にしなかった理由: `blind_runs/` に数字で並ぶと「1 回目の盲検」と
読まれる。本物の回は **003（3 体目）から**始める ── 1・2 体目は当時 trace を
取っていないので、この形では凍らせられない（凍結文書の「現状の穴」を参照）。

    python scripts/blind_run.py replay _selftest
    python scripts/blind_run.py floor  _selftest
    python scripts/blind_run.py sheet  _selftest

★ `_replay/` は走らせるたびに作り直される作業物なので commit しない。
