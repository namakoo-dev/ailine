---
paths:
  - "bench/*.json"
  - "tests/golden/**"
  - "tests/*_register.json"
  - "tests/battery_recorded.json"
---

# これは凍結した検体・記録

**測定の結果に合わせて検体を書き換えない。**

- `bench/*.json` は**実行前に凍結**してある（`_meta.frozen` に日時がある）。
  合わなかったら直すのは**実装か仮説**であって、検体ではない。
- `tests/golden/**` は画面の写真。意図して文言を変えたときだけ作り直し、
  **必ず `git diff` で中身を人が読む**（増えたぶんは意図した追加か）。
- `tests/*_register.json` の免除には**理由と unlock 条件**が要る。数を合わせるために足さない。
- `tests/battery_recorded.json` の MATRIX / BATTERY / ACCURACY は**実機を回して人が書く**。
  `refresh_records.py` は触らない（自動で揃えたら記録は常に一致し、二度と警告しなくなる＝恒真）。

★ 実際に踏んだ形: 検体が実事故を誤再現したまま処方を書いた／検体の発明した引数に
  実装が受け口を作りかけた。**直すのは常に検体でなく治具の側か、仮説の側。**
