#!/bin/sh
set -eu
SCRIPT_DIR=$(CDPATH= cd -- "$(dirname -- "$0")" && pwd)
BUNDLED_PYTHON="$HOME/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/bin/python3"
if [ -n "${VOCAB_PYTHON:-}" ]; then
  PY="$VOCAB_PYTHON"
elif [ -x "$BUNDLED_PYTHON" ]; then
  PY="$BUNDLED_PYTHON"
else
  PY=python3
fi
if [ "${1:-}" = test ]; then
  exec "$PY" "$SCRIPT_DIR/test_vocab.py"
fi
exec "$PY" "$SCRIPT_DIR/vocab.py" "$@"
