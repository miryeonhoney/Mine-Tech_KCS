# -*- coding: utf-8 -*-
"""Jarvis STT 사이드카 — 회의실 음성 인식 서버.

jarvis/ 폴더의 Whisper 엔진을 HTTP로 노출한다 (모델 1회 로드, 이후 즉시 응답).
실행:  ./jarvis_stt.sh   (jarvis의 stt-venv 파이썬으로 구동)
포트:  8765  (dashboard_app이 /api/conference/stt 로 프록시)

API:
  GET  /health → {"ok": true}
  POST /stt    → body: WAV(PCM 16-bit mono) → {"ok": true, "text": "..."}
"""
import io
import json
import os
import sys
import wave
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer

import numpy as np

# jarvis 폴더를 경로에 추가 (JARVIS_DIR 환경변수 또는 ../jarvis)
JARVIS_DIR = os.environ.get(
    "JARVIS_DIR",
    os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "jarvis"),
)
sys.path.insert(0, JARVIS_DIR)

from core.stt import WhisperSTT  # noqa: E402  (jarvis 모듈)

MODEL = os.environ.get("WHISPER_MODEL", "base")
PORT = int(os.environ.get("JARVIS_STT_PORT", "8765"))

print(f"[jarvis-stt] Whisper '{MODEL}' 모델 로딩 중… (jarvis: {JARVIS_DIR})")
stt = WhisperSTT(MODEL)
print(f"[jarvis-stt] 준비 완료 → http://127.0.0.1:{PORT}/stt")


def wav_to_float32_16k(body: bytes):
    """WAV 바이트 → 16kHz mono float32 (Whisper 입력 형식)."""
    with wave.open(io.BytesIO(body), "rb") as w:
        rate = w.getframerate()
        ch = w.getnchannels()
        width = w.getsampwidth()
        frames = w.readframes(w.getnframes())
    if width != 2:
        raise ValueError(f"16-bit PCM만 지원 (받은 폭: {width * 8}bit)")
    audio = np.frombuffer(frames, dtype=np.int16).astype(np.float32) / 32768.0
    if ch > 1:
        audio = audio.reshape(-1, ch).mean(axis=1)
    if rate != 16000:
        n = int(len(audio) * 16000 / rate)
        audio = np.interp(
            np.linspace(0, len(audio) - 1, n), np.arange(len(audio)), audio
        ).astype(np.float32)
    return audio


class Handler(BaseHTTPRequestHandler):
    def _json(self, code, obj):
        body = json.dumps(obj, ensure_ascii=False).encode("utf-8")
        self.send_response(code)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def do_GET(self):
        if self.path == "/health":
            return self._json(200, {"ok": True, "model": MODEL})
        return self._json(404, {"ok": False})

    def do_POST(self):
        if self.path != "/stt":
            return self._json(404, {"ok": False})
        try:
            length = int(self.headers.get("Content-Length", 0))
            if not (44 < length < 32 * 1024 * 1024):
                return self._json(400, {"ok": False, "error": "잘못된 오디오 크기"})
            body = self.rfile.read(length)
            audio = wav_to_float32_16k(body)
            if len(audio) < 16000 * 0.3:  # 0.3초 미만은 무시
                return self._json(200, {"ok": True, "text": ""})
            r = stt.model.transcribe(
                audio, language="ko", fp16=False,
                initial_prompt=("핵심광물 전문가 회의. 리튬, 코발트, 니켈, 희토류, 텅스텐, "
                                "몰리브덴, 망간, 흑연, 갈륨, 게르마늄, 우라늄, 유연탄, "
                                "공급망, 수급, 매장량, 수출통제, K-RISK, 비축."))
            text = (r.get("text") or "")
            print(f"[jarvis-stt] {len(audio)/16000:.1f}s → {text!r}")
            return self._json(200, {"ok": True, "text": text.strip()})
        except Exception as e:
            print("[jarvis-stt] 오류:", e)
            return self._json(500, {"ok": False, "error": str(e)})

    def log_message(self, *a):  # 기본 액세스 로그 끄기
        pass


if __name__ == "__main__":
    ThreadingHTTPServer(("127.0.0.1", PORT), Handler).serve_forever()
