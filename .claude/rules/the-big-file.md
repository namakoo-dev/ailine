---
paths:
  - "src/ailine/__init__.py"
---

# この冊を触ったら、記録も動く

`src/ailine/__init__.py` は **公開面が凍結**され、**行数が README に記録**されている。

編集したあと commit する前に:

1. 公開面が変わったなら作り直して **diff を人が読む**
   `AILINE_REGEN_SURFACE=1 python -m pytest tests/test_public_surface_is_frozen.py`
   ★ `from X import Y` を足すと **Y が ailine の公開面に生える**（実際に `NoReturn` が生えた）。
     注釈のためだけなら `import X` にする（モジュールは数えられない）。
2. 行数・図・試験数を揃える
   `python scripts/refresh_records.py --write`

★ どちらも pre-commit が止めるので、忘れても事故にはならない。ここに書いてあるのは
  **止められてから気づくより、書く前に知っている方が速い**から。
