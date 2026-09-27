#!/usr/bin/env bash
# Workspace GAS を clasp で本番へ送る。
# 認証: ~/.clasprc.json か環境変数 CLASPRC_JSON（clasp login 済みの中身）
set -euo pipefail

ROOT="$(cd "$(dirname "$0")/.." && pwd)"
cd "$ROOT"

DEPLOY_ID="${GAS_DEPLOY_ID:-AKfycbzMELimQThNdPUShwo2_KBzJd8kGy9BNdRyOYNgu_sg41t2SleVRiWXFztZJ48e2l9L}"
WEBAPP="${GAS_WEB_APP_URL:-https://script.google.com/macros/s/${DEPLOY_ID}/exec}"

if [[ -n "${CLASPRC_JSON:-}" ]]; then
  umask 077
  printf '%s\n' "$CLASPRC_JSON" > "$HOME/.clasprc.json"
fi

if [[ ! -s "$HOME/.clasprc.json" ]]; then
  echo "clasp のログインがありません。手元で clasp login した ~/.clasprc.json を GitHub secret CLASPRC_JSON に入れてください。" >&2
  exit 2
fi

if ! command -v clasp >/dev/null 2>&1; then
  mkdir -p "$HOME/.local"
  npm install --prefix "$HOME/.local" @google/clasp@3.4.1
  export PATH="$HOME/.local/bin:$HOME/.local/node_modules/.bin:$PATH"
fi
if ! command -v clasp >/dev/null 2>&1; then
  echo "npx で clasp を実行します"
  clasp() { npx --yes @google/clasp@3.4.1 "$@"; }
fi

echo "== clasp push (1/2) =="
clasp push --force
echo "== clasp push (2/2) =="
clasp push --force

echo "== clasp deploy =="
clasp deploy -i "$DEPLOY_ID" -d "workspace-hub"

if [[ "${APPLY_HUB_HOME:-1}" == "1" ]]; then
  echo "== wait then setupHubHome =="
  sleep 20
  curl -fsSL "${WEBAPP}?api=status" || true
  echo
  curl -fsSL "${WEBAPP}?api=setupHubHome" || true
  echo
fi

echo "done"
