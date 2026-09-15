#!/bin/sh
# 開発者用: git hooks を入れる（clone 直後に一度）。
#   sh scripts/install-hooks.sh
root=$(git rev-parse --show-toplevel) || exit 1
printf '#!/bin/sh\nexec "$(git rev-parse --show-toplevel)/scripts/pre-push.sh" "$@"\n' > "$root/.git/hooks/pre-push"
chmod +x "$root/.git/hooks/pre-push"
echo "✓ pre-push を入れました（CI で走らない -m local を押す前に走らせます）"
# ★ 2026-09-15: 速い番人（改行コード・制御文字 2.4 秒）を commit の瞬間に走らせる。
#   番人は元から在って毎回鳴っていたが、鳴る場所が全件テストの末尾だった ── 直すのはループの長さ。
#   ★ global の ~/.githooks/pre-commit（機密語）が、この repo 固有フックを自分の後ろで呼ぶ。
printf '#!/bin/sh\nexec "$(git rev-parse --show-toplevel)/scripts/pre-commit.sh" "$@"\n' > "$root/.git/hooks/pre-commit"
chmod +x "$root/.git/hooks/pre-commit"
echo "✓ pre-commit を入れました（改行コード/制御文字・2.4 秒）"
