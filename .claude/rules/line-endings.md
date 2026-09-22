---
paths:
  - "src/ailine/__init__.py"
  - "gui/**/*.py"
  - "src/ailine_core/target_sheet.py"
  - "README.md"
  - "docs/ENGINEERING.md"
---

# この冊は CRLF

**書く前にバイトで数える。書いた後ももう一度数える。**

```python
import io
b = io.open(path, "rb").read()
print("CRLF", b.count(b"\r\n"), "LF", b.count(b"\n") - b.count(b"\r\n"))
```

- 読むときは `io.open(..., newline="")`。`Path.read_text()` は**読んだ時点で CRLF を LF に潰す**ので、
  そのあと `write_bytes` で書いても手遅れ。
- ヒアドキュメントで追記しない（LF が混ざる）。
- 同じ repo で `src/ailine_core/*.py` と `docs/*.md` は **LF**。混在させると番人が止める。

★ なぜ在るか: この記憶は「書く前にバイトで数える」と書いてあり、セッション冒頭で読んだ。
  **そのうえで 1 日に 5 回壊した。** 4 回目までは pre-push の 17 分を使って気づいた。
  規律では効かなかったので、**そのファイルを読んだ瞬間**に出る場所へ移した。
