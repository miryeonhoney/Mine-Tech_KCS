#!/bin/zsh
# Jarvis STT 사이드카 실행 — 회의실 음성 인식
# 사용: ./jarvis_stt.sh   (기본 포트 8765)
DIR="$(cd "$(dirname "$0")" && pwd)"
JARVIS="${JARVIS_DIR:-$DIR/../jarvis}"
PY="$JARVIS/stt-venv/bin/python"
if [ ! -x "$PY" ]; then
  echo "❌ jarvis stt-venv를 찾을 수 없습니다: $PY"
  echo "   jarvis 폴더 위치가 다르면: JARVIS_DIR=/path/to/jarvis ./jarvis_stt.sh"
  exit 1
fi
export JARVIS_DIR="$JARVIS"
export WHISPER_MODEL="${WHISPER_MODEL:-small}"   # 한국어 정확도 (base보다 우수)
exec "$PY" "$DIR/jarvis_stt_server.py"
