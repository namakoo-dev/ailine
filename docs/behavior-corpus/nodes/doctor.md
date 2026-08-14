# doctor — セットアップ診断

**分類: う（環境検出に結合: ollama HTTP・LibreOffice 検出・ファイル配置）**

## 挙動

`ailine doctor` が 7 項目を検査し ✓/× で列挙、×には直し方を 1 行添える。全✓で exit 0:
①python 3.10+ ②openpyxl ③ollama 到達 (GET /api/tags) ④既定モデルの有無
（無ければ `ollama pull <モデル>` を案内・タグは完全一致で照合） ⑤LibreOffice 検出
（basrun の検出ロジックと同経路） ⑥basrun.py の所在 ⑦demo/ の有無。

## なぜ

冷間監査の教訓: セットアップは「全部揃っていれば」通るが、揃っていない環境で
どこが欠けているかをユーザーが自力診断するのは難しい（ollama 404 に「serve を確認」と
誤案内した前科もある — 404 と接続拒否は別の病気）。買った人の最初の 5 分を守る器官。

- 出典: commit 688eab7・冷間監査 CONFUSING 所見（誤ヒント）
