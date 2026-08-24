#!/bin/sh
# 開発者用: git hooks を入れる（clone 直後に一度）。
#   sh scripts/install-hooks.sh
root=$(git rev-parse --show-toplevel) || exit 1
printf '#!/bin/sh\nexec "$(git rev-parse --show-toplevel)/scripts/pre-push.sh" "$@"\n' > "$root/.git/hooks/pre-push"
chmod +x "$root/.git/hooks/pre-push"
echo "✓ pre-push を入れました（CI で走らない -m local を押す前に走らせます）"
