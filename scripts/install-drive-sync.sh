#!/bin/sh
# 컴퓨터마다 한 번만 실행: 커밋/풀(pull) 직후 내작업대_젠트리.html을 구글 드라이브 작업대 폴더로 자동 복사하는 git hook 설치.
# 사용: 저장소 폴더에서  sh scripts/install-drive-sync.sh
# 드라이브 경로가 다르면:  DRIVE_DIR="/g/내 드라이브/작업대" sh scripts/install-drive-sync.sh

DRIVE_DIR="${DRIVE_DIR:-/g/내 드라이브/작업대}"
ROOT="$(git rev-parse --show-toplevel)" || exit 1
HOOKS="$ROOT/.git/hooks"

for name in post-commit post-merge; do
  cat > "$HOOKS/$name" <<EOF
#!/bin/sh
SRC="\$(git rev-parse --show-toplevel)/내작업대_젠트리.html"
DEST="$DRIVE_DIR/내작업대_젠트리.html"
if [ -f "\$SRC" ] && [ -d "$DRIVE_DIR" ]; then
  cp -f "\$SRC" "\$DEST" && echo "[drive-sync] 내작업대_젠트리.html -> \$DEST"
fi
exit 0
EOF
  chmod +x "$HOOKS/$name"
done

echo "설치 완료: $HOOKS/post-commit, post-merge  (복사 대상: $DRIVE_DIR)"
