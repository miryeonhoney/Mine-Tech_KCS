"""
핵심광물 위기 대시보드 v2
=====================================================================
데이터 소스:
   1) 로컬 파일   → mineral_collector_all.py 가 수집한 JSON
   2) KOMIR API  → 한국광해광업공단 국가별 광종 수출입 현황 (대시보드용)
   3) 네이버 뉴스 → 핵심광물 관련 뉴스 검색
   4) USGS       → 미국지질조사국 광물 매장량/생산량 (2025 기준)

실행:
   pip install flask pandas requests openpyxl anthropic
   python dashboard_app.py
   → http://127.0.0.1:8080
=====================================================================
"""

import os, re, json, glob, time, smtplib, hmac, hashlib, xml.etree.ElementTree as ET
import pandas as pd
from datetime import datetime
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import requests
import anthropic
from openai import OpenAI
from flask import Flask, request, jsonify, Response, stream_with_context, session, redirect

# ═══════════════════════════════════════════════════════════════
#  ① 설정
# ═══════════════════════════════════════════════════════════════
# ── .env 로더 (외부 의존성 없이 .env 파일을 읽어 환경변수로) ──
def _load_dotenv(path=".env"):
    p = os.path.join(os.path.dirname(__file__), path)
    if not os.path.exists(p):
        return
    with open(p, encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if not line or line.startswith("#") or "=" not in line:
                continue
            k, v = line.split("=", 1)
            os.environ.setdefault(k.strip(), v.strip().strip('"').strip("'"))
_load_dotenv()

# 비밀키는 환경변수(.env 또는 호스팅 환경변수)에서 읽습니다. 코드에 하드코딩하지 마세요.
PUBLIC_DATA_KEY     = os.environ.get("PUBLIC_DATA_KEY", "")
NAVER_CLIENT_ID     = os.environ.get("NAVER_CLIENT_ID", "")
NAVER_CLIENT_SECRET = os.environ.get("NAVER_CLIENT_SECRET", "")
EMAIL_ADDRESS       = os.environ.get("EMAIL_ADDRESS", "")
EMAIL_PASSWORD      = os.environ.get("EMAIL_PASSWORD", "")
DATA_DIR            = "mineral_data"
SUBSCRIBERS_FILE    = os.environ.get("SUBSCRIBERS_FILE", "subscribers.json")
CACHE_TTL           = 300
ANTHROPIC_API_KEY   = os.environ.get("ANTHROPIC_API_KEY", "")

# ── AI 전문가 회의실 (OpenAI) 설정 ──────────────────────────────
# 전역 공용 키: 전문가별 "api_key"가 비어 있으면 이 값을 사용합니다.
OPENAI_API_KEY       = os.environ.get("OPENAI_API_KEY", "")
DEFAULT_OPENAI_MODEL = os.environ.get("OPENAI_MODEL", "gpt-5.2-chat-latest")  # GPT-5 계열. 플래그십은 "gpt-5.5"

# ── 공개 배포용 보안: 회의실 비밀번호 게이트 ──
# CONFERENCE_PASSWORD 가 비어 있으면(로컬 개발) 게이트가 꺼집니다.
CONFERENCE_PASSWORD = os.environ.get("CONFERENCE_PASSWORD", "")
SECRET_KEY          = os.environ.get("SECRET_KEY", "dev-only-change-me")

# ── 구독: 저장(DB)·발송(SMTP)·자동발송(크론) ──────────────────
DATABASE_URL = os.environ.get("DATABASE_URL", "")          # 있으면 Postgres, 없으면 JSON 파일
SMTP_HOST    = os.environ.get("SMTP_HOST", "smtp.gmail.com")
SMTP_PORT    = int(os.environ.get("SMTP_PORT", "587") or 587)
SMTP_USER    = os.environ.get("SMTP_USER", "") or EMAIL_ADDRESS
SMTP_PASS    = os.environ.get("SMTP_PASS", "") or EMAIL_PASSWORD
MAIL_FROM    = os.environ.get("MAIL_FROM", "") or EMAIL_ADDRESS
APP_BASE_URL = os.environ.get("APP_BASE_URL", "").rstrip("/")  # 수신거부 링크용 (예: https://app.onrender.com)
CRON_TOKEN   = os.environ.get("CRON_TOKEN", "")            # /cron/daily 보호 토큰

MINERAL_EXPERTS = {
    "리튬": {
        "name": "리튬 박사",
        "title": "배터리·전기차 공급망 전문가",
        "avatar": "🔋",
        "color": "#4a9eff",
        "system": """당신은 '리튬 박사'입니다. 한국에너지기술연구원 소속의 리튬 배터리 및 전기차 공급망 전문가입니다.
전문 분야: 리튬 채굴, 배터리 소재, 전기차 전환, 칠레·아르헨티나·호주 공급망, LFP vs NCM 기술.
성격: 데이터 중심적, 논리적. 한국의 리튬 수입 의존도(95%)와 가격 변동성을 핵심 이슈로 항상 언급.
다중 토론 지침: 회의실에 여러 전문가가 있을 때, 다른 박사의 발언을 직접 인용하며 동의/반박하세요. 200자 내외로 핵심만."""
    },
    "코발트": {
        "name": "코발트 박사",
        "title": "아프리카 자원 리스크 분석가",
        "avatar": "⚠️",
        "color": "#e8281a",
        "system": """당신은 '코발트 박사'입니다. 산업통상자원부 자문 코발트·아프리카 자원 리스크 전문 분석가입니다.
전문 분야: 콩고민주공화국(DRC) 정치 리스크, 공급망 집중도, 중국의 DRC 광산 장악(70%).
성격: 지정학적 관점 강조, 비관적이지만 현실적. 리스크 시나리오를 구체적으로 제시.
다중 토론 지침: 다른 박사 발언에 "잠깐, ○○ 박사—" 처럼 끼어드는 스타일. 최악의 시나리오를 항상 경고. 200자 내외."""
    },
    "니켈": {
        "name": "니켈 박사",
        "title": "인도네시아·필리핀 광물 시장 전문가",
        "avatar": "🌏",
        "color": "#39c96e",
        "system": """당신은 '니켈 박사'입니다. KOTRA 소속 동남아 광물 시장 전문가입니다.
전문 분야: 인도네시아 니켈 수출 규제, 필리핀 광산 정책, HPAL 기술, 스테인리스·배터리 수요.
성격: 실용적, 외교적 해법 선호. 인도네시아와의 협력 가능성을 낙관적으로 봄.
다중 토론 지침: 다른 박사들이 위기를 강조할 때 "그렇지만 기회도 있습니다—"로 균형을 잡음. 200자 내외."""
    },
    "희토류": {
        "name": "희토류 박사",
        "title": "중국 자원 외교 및 희토류 정책 교수",
        "avatar": "🇨🇳",
        "color": "#ffc400",
        "system": """당신은 '희토류 박사'입니다. 서울대학교 자원외교학과 교수이자 희토류 정책 전문가입니다.
전문 분야: 중국 희토류 독점(60%), 수출 규제 역사, 미중 무역갈등, 2010년 중일 분쟁 사례.
성격: 학문적, 역사적 맥락 중시. 장기 전략의 중요성을 강조.
다중 토론 지침: 다른 박사 논의에 역사적 사례로 무게를 더함. "역사를 보면..."으로 시작하는 발언 자주 함. 200자 내외."""
    },
    "텅스텐": {
        "name": "텅스텐 박사",
        "title": "방산·산업소재 공급망 리스크 전문가",
        "avatar": "⚙️",
        "color": "#a78bfa",
        "system": """당신은 '텅스텐 박사'입니다. 한국방위산업진흥회 공급망 리스크 전문가입니다.
전문 분야: 텅스텐 방산 활용, 절삭공구·초경합금, 북한 텅스텐 매장량, 안보 리스크.
성격: 안보 관점 최우선. 경제성보다 전략적 자율성을 중시. "이건 안보 문제입니다"를 자주 씀.
다중 토론 지침: 경제·기술 논의에 항상 안보 렌즈를 씌움. 200자 내외."""
    },
    "망간": {
        "name": "망간 박사",
        "title": "철강·차세대 배터리 소재 연구원",
        "avatar": "🔩",
        "color": "#f97316",
        "system": """당신은 '망간 박사'입니다. 포스코 기술연구원 소속 망간·철강 소재 전문가입니다.
전문 분야: 망간 강철 합금, LMFP 배터리, 남아프리카공화국 공급망, 전기로 제강.
성격: 기술 낙관주의자. "사실 망간이 핵심입니다"로 논점 전환을 즐김. 200자 내외."""
    },
}

# ── 광물 외 분야 전문가 (경제·정치 등) ─────────────────────────
EXTRA_EXPERTS = {
    "흑연": {
        "name": "흑연 박사",
        "title": "음극재·배터리 소재 전문가",
        "avatar": "⚫",
        "color": "#94a3b8",
        "category": "광물",
        "system": """당신은 '흑연 박사'입니다. 한국전자기술연구원 소속 흑연·음극재 전문가입니다.
전문 분야: 천연/인조 흑연, 배터리 음극재, 중국의 흑연 수출 통제(2023), 구형흑연 가공 독점.
성격: 차분하고 기술 디테일에 강함. "음극재 없이는 배터리도 없습니다"를 자주 언급.
다중 토론 지침: 양극재(리튬·니켈) 중심 논의에 "음극재 관점도 보셔야 합니다"로 균형을 맞춤. 200자 내외."""
    },
    "경제": {
        "name": "경제 박사",
        "title": "거시경제·자원가격 전문가",
        "avatar": "📈",
        "color": "#22d3ee",
        "category": "경제",
        "system": """당신은 '경제 박사'입니다. 한국개발연구원(KDI) 소속 거시경제 전문가입니다.
전문 분야: 원자재 가격이 물가·환율·무역수지에 미치는 파급, 인플레이션, 경기 사이클, 가격 헤지.
성격: 숫자와 거시 지표로 말함. 개별 광물 이슈를 항상 거시경제 충격으로 환산해 제시.
다중 토론 지침: 다른 전문가의 산업·안보 논의를 "그게 거시경제로는 이렇게 나타납니다"로 받아 정량화. 200자 내외."""
    },
    "통상": {
        "name": "통상 박사",
        "title": "무역·통상정책 전문가",
        "avatar": "🤝",
        "color": "#2dd4bf",
        "category": "경제",
        "system": """당신은 '통상 박사'입니다. 대외경제정책연구원(KIEP) 소속 무역·통상 전문가입니다.
전문 분야: FTA·관세, IRA·CRMA 등 핵심광물 통상규제, 원산지 규정, 수출통제 대응.
성격: 협상 테이블 관점. 규제를 리스크이자 협상 카드로 봄.
다중 토론 지침: 기술·안보 논의를 "통상 규범상 이렇게 풀어야 합니다"로 제도화. 다른 박사 의견을 통상 조항에 연결. 200자 내외."""
    },
    "지정학": {
        "name": "지정학 박사",
        "title": "자원안보·국제정치 전문가",
        "avatar": "🌐",
        "color": "#f472b6",
        "category": "정치",
        "system": """당신은 '지정학 박사'입니다. 국립외교원 소속 자원안보·국제정치 전문가입니다.
전문 분야: 미중 패권 경쟁, 자원의 무기화, 동맹 기반 공급망 재편(프렌드쇼어링), 해상 수송로 안보.
성격: 큰 그림과 권력 역학으로 해석. "이건 결국 힘의 문제입니다"를 자주 씀.
다중 토론 지침: 경제·기술 논의를 국제정치 구도로 끌어올려 재해석. 다른 박사 발언의 지정학적 함의를 짚음. 200자 내외."""
    },
    "정책": {
        "name": "정책 박사",
        "title": "산업정책·자원전략 전문가",
        "avatar": "🏛️",
        "color": "#a3e635",
        "category": "정치",
        "system": """당신은 '정책 박사'입니다. 산업연구원(KIET) 소속 산업정책·자원전략 전문가입니다.
전문 분야: 비축, 국산화·재자원화(리사이클), 보조금·세제, 해외 자원개발, 컨트롤타워.
성격: 실행 가능한 정책 대안 제시에 집중. "그래서 정부는 무엇을 해야 하나"로 토론을 수렴.
다중 토론 지침: 다른 전문가들이 진단한 문제를 받아 "그렇다면 정책 처방은—"으로 구체적 대안을 묶어냄. 200자 내외."""
    },
}
MINERAL_EXPERTS.update(EXTRA_EXPERTS)

# 모든 전문가에 기본 필드 채우기 (category / model / api_key)
for _v in MINERAL_EXPERTS.values():
    _v.setdefault("category", "광물")
    _v.setdefault("model", DEFAULT_OPENAI_MODEL)
    _v.setdefault("api_key", "")

USGS_DATA = {
    "리튬":   {"매장량_만톤": 2800,  "생산량_만톤": 24,   "1위국": "칠레",           "출처": "USGS MCS 2025"},
    "코발트": {"매장량_만톤": 1000,  "생산량_만톤": 23,   "1위국": "콩고민주공화국",  "출처": "USGS MCS 2025"},
    "니켈":   {"매장량_만톤": 10000, "생산량_만톤": 360,  "1위국": "인도네시아",      "출처": "USGS MCS 2025"},
    "흑연":   {"매장량_만톤": 28000, "생산량_만톤": 1300, "1위국": "중국",            "출처": "USGS MCS 2025"},
    "희토류": {"매장량_만톤": 11000, "생산량_만톤": 39,   "1위국": "중국",            "출처": "USGS MCS 2025"},
    "망간":   {"매장량_만톤": 150000,"생산량_만톤": 2000, "1위국": "남아프리카공화국", "출처": "USGS MCS 2025"},
}

# ── 국가별 주요 항구/중심 좌표 (뱃길 시각화용 · 항로 보정판) ─────
COUNTRY_COORDS = {
    "호주":           [-20.3,  118.6],  "호 주":          [-20.3,  118.6],  # Port Hedland
    "중국":           [31.2,   121.5],  "중 국":          [31.2,   121.5],
    "인도네시아":     [-6.2,   106.8],
    "칠레":           [-23.6,  -70.4],
    "캐나다":         [49.3,  -123.1],
    "남아프리카공화국":[-29.9,   31.0],
    "러시아":         [43.1,   131.9],  "러시아연방":     [43.1,   131.9],
    "콩고민주공화국": [-6.0,    12.2],  # 마타디/바나나 (콩고강 하구)
    "필리핀":         [14.6,   121.0],  "필 리 핀":       [14.6,   121.0],
    "미국":           [33.7,  -118.2],  "미 국":          [33.7,  -118.2],
    "인도":           [18.9,    72.8],  "인 도":          [18.9,    72.8],
    "브라질":         [-23.9,  -46.3],
    "카자흐스탄":     [51.2,    71.4],
    "일본":           [35.4,   139.6],  "일 본":          [35.4,   139.6],
    "페루":           [-12.1,  -77.1],
    "잠비아":         [-15.4,   28.3],
    "짐바브웨":       [-17.8,   31.1],
    "모잠비크":       [-25.9,   32.6],
    "마다가스카르":   [-18.2,   49.4],
    "탄자니아":       [-6.8,    39.3],
    "가봉":           [0.4,     9.5],
    "베트남":         [10.8,   106.7],  "베 트 남":       [10.8,   106.7],
    "말레이시아":     [3.1,    101.7],
    "태국":           [13.7,   100.5],  "태 국":          [13.7,   100.5],
    "싱가포르":       [1.3,    103.8],
    "미얀마":         [16.9,    96.2],
    "뉴칼레도니아":   [-22.3,   166.5],
    "나미비아":       [-22.9,   14.5],
    "사우디아라비아": [27.0,    49.9],  # 주바일/라스타누라 (걸프 연안)
    "아랍에미리트":   [25.2,    55.3],
    "카메룬":         [4.0,     9.7],
    "콩고":           [-4.8,    11.8],  # 푸앵트누아르
    "쿠바":           [23.1,   -82.4],
    "우크라이나":     [46.5,    30.7],
    "독일":           [53.5,     9.9],  "독 일":          [53.5,     9.9],
    "영국":           [51.5,    -0.1],  "영 국":          [51.5,    -0.1],
    "네덜란드":       [51.9,     4.5],
    "벨기에":         [51.2,     4.4],  "벨 기 에":       [51.2,     4.4],
    "프랑스":         [49.5,     0.1],  "프 랑 스":       [49.5,     0.1],  # 르아브르
    "스페인":         [36.7,    -6.4],  "스 페 인":       [36.7,    -6.4],
    "이탈리아":       [40.6,    14.3],
    "터키":           [41.0,    29.0],  "튀르키예":       [41.0,    29.0],
    "튀르키예공화국": [41.0,    29.0],
    "아르헨티나":     [-34.6,  -58.4],
    "멕시코":         [19.0,  -104.3],  # 만사니요 (태평양측)
    "콜롬비아":       [10.4,   -75.5],
    "볼리비아":       [-18.5,  -70.4],  # 아리카항 경유 (칠레)
    "대만":           [25.1,   121.5],  "대 만":          [25.1,   121.5],
    "홍콩":           [22.3,   114.2],
    "몽골":           [47.9,   106.9],  "몽 골":          [47.9,   106.9],
    "파푸아뉴기니":   [-9.4,   147.2],
    "뉴질랜드":       [-36.8,   174.7],
    "카타르":         [25.3,    51.5],
    "코트디부아르":   [5.4,     -4.0],
    "가나":           [5.6,     -0.2],
    "나이지리아":     [6.5,     3.4],
    "에티오피아":     [11.6,   43.1],  # 지부티항 경유
    "케냐":           [-4.1,    39.7],
    "모로코":         [33.6,    -7.6],  "모 로 코":       [33.6,    -7.6],
    "알제리":         [36.7,     3.1],
    "이집트":         [31.2,    29.9],
    "파키스탄":       [24.9,    67.0],
    "방글라데시":     [22.3,    91.8],
    "스리랑카":       [6.9,    79.9],
    "이란":           [27.2,    56.3],  "이 란":          [27.2,    56.3],
    "이라크":         [29.4,    48.0],
    "캄보디아":       [10.6,   103.5],
    "라오스":         [17.9,   102.6],
    "우즈베키스탄":   [41.3,    69.2],
    "카자흐스탄":     [51.2,    71.4],
    "그리스":         [37.9,    23.7],  "그 리 스":       [37.9,    23.7],
    "노르웨이":       [59.9,    10.7],
    "스웨덴":         [57.7,    11.9],  "스 웨 덴":       [57.7,    11.9],
    "핀란드":         [60.2,    25.0],
    "폴란드":         [54.4,    18.6],
    "포르투갈":       [38.7,    -9.1],  "포루투갈":       [38.7,    -9.1],
}


HS_CODES = {
    "2825":"리튬화합물","2530":"리튬광석","2605":"코발트광석",
    "2604":"니켈광석","2504":"천연흑연","2615":"희토류광석",
    "2846":"희토류화합물","8105":"코발트가공품","7501":"니켈가공품",
}

NEWS_KEYWORDS = ["핵심광물","리튬 광물","코발트 광물","니켈 광물","희토류","광물 공급망"]

app    = Flask(__name__)
app.secret_key = SECRET_KEY

def _conf_authed():
    """회의실 접근 허용 여부. 비밀번호 미설정(로컬)이면 항상 허용."""
    return (not CONFERENCE_PASSWORD) or (session.get("conf_ok") is True)

_cache = {}

def cache_get(k):
    it = _cache.get(k)
    if not it: return None
    ttl = it.get("ttl") or CACHE_TTL
    return it["d"] if time.time()-it["t"] < ttl else None

def cache_set(k, d, ttl=None):
    _cache[k] = {"t": time.time(), "d": d, "ttl": ttl}

def latest_json(prefix):
    fs = sorted(glob.glob(os.path.join(DATA_DIR,"**",f"{prefix}_*.json"), recursive=True))
    return fs[-1] if fs else None

def load_json(path):
    if not path or not os.path.exists(path): return []
    try:
        with open(path, encoding="utf-8") as f: return json.load(f)
    except: return []

def local_news():  return load_json(latest_json("news"))
def local_mines(): return load_json(latest_json("komir_mines"))

def local_customs():
    rows = load_json(latest_json("customs_trade"))
    if rows: return rows
    csv_path = os.path.join(os.path.dirname(__file__),
                            "한국광해광업공단_국가별 광종 수출입 현황_20250328.csv")
    if not os.path.exists(csv_path): return []
    try:
        df = pd.read_csv(csv_path, encoding="cp949")
        # 최신 연도만 사용
        if "기간" in df.columns:
            latest_year = df["기간"].dropna().astype(int).max()
            df = df[df["기간"].astype(int) == latest_year]
        result = []
        for _, row in df.iterrows():
            imp = exp = 0
            imp_ton = exp_ton = 0
            try:
                raw = row.get("수입금액(천불)", 0)
                imp = 0 if pd.isna(raw) else float(str(raw).replace(",","")) * 1000
            except: pass
            try:
                raw = row.get("수출금액(천불)", 0)
                exp = 0 if pd.isna(raw) else float(str(raw).replace(",","")) * 1000
            except: pass
            try:
                raw = row.get("수입중량(톤)", 0)
                imp_ton = 0 if pd.isna(raw) else float(str(raw).replace(",",""))
            except: pass
            try:
                raw = row.get("수출중량(톤)", 0)
                exp_ton = 0 if pd.isna(raw) else float(str(raw).replace(",",""))
            except: pass
            result.append({
                "광물명":        str(row.get("품목명","") or "").strip(),
                "국가명":        str(row.get("국가명","") or "").strip(),
                "수입금액(달러)": imp,
                "수출금액(달러)": exp,
                "수입중량(톤)":  imp_ton,
                "수출중량(톤)":  exp_ton,
                "무역수지(달러)": exp - imp,
                "출처":         "KOMIR CSV",
            })
        return result
    except Exception as e:
        print(f"[CSV 로드 오류] {e}")
        return []

def prev_month():
    n = datetime.now()
    return f"{n.year}{n.month-1:02d}" if n.month > 1 else f"{n.year-1}12"

def fetch_customs():
    c = cache_get("customs")
    if c is not None: return c
    if PUBLIC_DATA_KEY.startswith("여기에"): return local_customs()
    url      = "https://api.odcloud.kr/api/3070183/v1/uddi:8e13f741-2a4f-4e7a-8e60-a3a9de3a9b50"
    all_rows = []
    page     = 1
    per_page = 100
    while True:
        try:
            r = requests.get(url, params={
                "serviceKey": PUBLIC_DATA_KEY,
                "page": page, "perPage": per_page,
            }, timeout=15)
            if r.status_code == 401:
                print("[KOMIR API] 인증 오류")
                break
            if r.status_code != 200:
                print(f"[KOMIR API] HTTP {r.status_code}")
                break
            raw   = r.json()
            items = raw.get("data", [])
            if not items: break
            for it in items:
                mineral = (it.get("광종명") or it.get("광종") or it.get("mineral") or "")
                country = (it.get("국가명") or it.get("국가") or it.get("country") or "")
                imp_amt = exp_amt = 0
                for k, v in it.items():
                    if "수입" in k and ("액" in k or "금액" in k):
                        try: imp_amt = float(str(v).replace(",","") or 0)
                        except: pass
                    if "수출" in k and ("액" in k or "금액" in k):
                        try: exp_amt = float(str(v).replace(",","") or 0)
                        except: pass
                all_rows.append({
                    "광물명": mineral, "국가명": country,
                    "수입금액(달러)": imp_amt, "수출금액(달러)": exp_amt,
                    "수입중량(kg)": 0, "무역수지(달러)": exp_amt - imp_amt,
                    "출처": "KOMIR(광해광업공단)",
                })
            total = raw.get("totalCount", 0)
            print(f"[KOMIR API] 페이지 {page}: {len(items)}건 (전체 {total}건)")
            if page * per_page >= total: break
            page += 1
            time.sleep(0.3)
        except Exception as e:
            print(f"[KOMIR API] 오류: {e}")
            break
    result = all_rows if all_rows else local_customs()
    cache_set("customs", result)
    return result

def clean(t): return re.sub(r"<[^>]+>","",str(t)).strip()

def fetch_news():
    c = cache_get("news")
    if c is not None: return c
    if NAVER_CLIENT_ID.startswith("여기에"): return local_news()
    hdrs = {"X-Naver-Client-Id": NAVER_CLIENT_ID, "X-Naver-Client-Secret": NAVER_CLIENT_SECRET}
    all_news, seen = [], set()
    for kw in NEWS_KEYWORDS:
        try:
            r = requests.get("https://openapi.naver.com/v1/search/news.json",
                headers=hdrs, params={"query":kw,"display":5,"sort":"date"}, timeout=8)
            if r.status_code != 200: continue
            for it in r.json().get("items",[]):
                lnk = it.get("originallink","")
                if lnk in seen: continue
                seen.add(lnk)
                try: dt = datetime.strptime(it.get("pubDate",""),"%a, %d %b %Y %H:%M:%S +0900").strftime("%Y-%m-%d %H:%M")
                except: dt = it.get("pubDate","")
                all_news.append({
                    "제목": clean(it.get("title","")),
                    "요약": clean(it.get("description",""))[:80],
                    "언론사링크": lnk,
                    "발행일시": dt,
                    "검색키워드": kw,
                })
        except: continue
        time.sleep(0.15)
    result = all_news if all_news else local_news()
    cache_set("news", result)
    return result

def by_mineral(rows):
    s = {}
    for r in rows:
        nm = r.get("광물명","기타")
        try: v = float(str(r.get("수입금액(달러)",0)).replace(",","") or 0)
        except: v = 0
        s[nm] = s.get(nm,0) + v
    return sorted(s.items(), key=lambda x: x[1], reverse=True)

def by_country(rows):
    """국가별 수입 집계 (전체 합산) — 톤수 우선, 없으면 USD. 차트용."""
    tons, usd = {}, {}
    for r in rows:
        cn = (r.get("국가명","") or "").strip()
        if not cn or cn == "-": continue
        try: t = float(str(r.get("수입중량(톤)", 0) or 0).replace(",",""))
        except: t = 0
        try: v = float(str(r.get("수입금액(달러)", 0) or 0).replace(",",""))
        except: v = 0
        tons[cn] = tons.get(cn, 0) + t
        usd[cn]  = usd.get(cn, 0)  + v
    has_tons = any(v > 0 for v in tons.values())
    if has_tons:
        return sorted(tons.items(), key=lambda x: x[1], reverse=True)[:10]
    return sorted(usd.items(), key=lambda x: x[1], reverse=True)[:10]

def by_country_unit(rows):
    has_tons = any(
        float(str(r.get("수입중량(톤)", 0) or 0).replace(",","")) > 0
        for r in rows
    )
    return "톤" if has_tons else "USD"

_MINERAL_ALIAS = {
    "인상흑연": "흑연", "토상흑연": "흑연",  # 흑연 두 종류를 합산
}

def by_mineral_country(rows):
    """광물별 × 국가별 수입량(톤) 중첩 딕셔너리 반환.
    구조: {광물명: {국가명: 톤수}}  — 패널의 선택 광물 필터에 사용."""
    result = {}
    has_any_ton = False
    for r in rows:
        mn = (r.get("광물명","") or "").strip()
        cn = (r.get("국가명","") or "").strip()
        if not mn or not cn or cn == "-": continue
        mn = _MINERAL_ALIAS.get(mn, mn)  # 별칭 정규화
        try: t = float(str(r.get("수입중량(톤)", 0) or 0).replace(",",""))
        except: t = 0
        try: v = float(str(r.get("수입금액(달러)", 0) or 0).replace(",",""))
        except: v = 0
        if t > 0: has_any_ton = True
        val = t if t > 0 else v
        result.setdefault(mn, {})
        result[mn][cn] = result[mn].get(cn, 0) + val
    # 각 광물별로 상위 7개 국가만 유지
    trimmed = {
        mn: dict(sorted(cv.items(), key=lambda x: x[1], reverse=True)[:7])
        for mn, cv in result.items()
    }
    return trimmed, ("톤" if has_any_ton else "USD")

# ── 구독자 저장: DATABASE_URL 있으면 Postgres, 없으면 JSON 파일 ──
_db_ready = False

def _db_conn():
    import psycopg
    return psycopg.connect(DATABASE_URL)

def _ensure_db():
    global _db_ready
    if _db_ready:
        return
    with _db_conn() as c:
        c.execute("CREATE TABLE IF NOT EXISTS subscribers (email TEXT PRIMARY KEY, created_at TIMESTAMP DEFAULT now())")
        c.commit()
    _db_ready = True

def save_subs(s):  # JSON 폴백 전용
    with open(SUBSCRIBERS_FILE, "w", encoding="utf-8") as f:
        json.dump(s, f, ensure_ascii=False, indent=2)

def load_subs():
    if DATABASE_URL:
        try:
            _ensure_db()
            with _db_conn() as c:
                rows = c.execute("SELECT email FROM subscribers ORDER BY created_at").fetchall()
            return [r[0] for r in rows]
        except Exception as e:
            print("[DB] load_subs:", e); return []
    try:
        with open(SUBSCRIBERS_FILE, encoding="utf-8") as f: return json.load(f)
    except: return []

def add_sub(email):
    """추가되면 True, 이미 있으면 False."""
    if DATABASE_URL:
        _ensure_db()
        with _db_conn() as c:
            cur = c.execute("INSERT INTO subscribers(email) VALUES(%s) ON CONFLICT (email) DO NOTHING", (email,))
            c.commit()
            return cur.rowcount > 0
    subs = load_subs()
    if email in subs: return False
    subs.append(email); save_subs(subs); return True

def remove_sub(email):
    if DATABASE_URL:
        _ensure_db()
        with _db_conn() as c:
            c.execute("DELETE FROM subscribers WHERE email=%s", (email,)); c.commit()
        return True
    subs = load_subs()
    if email in subs:
        subs.remove(email); save_subs(subs)
    return True

def unsub_token(email):
    return hmac.new(SECRET_KEY.encode(), email.encode(), hashlib.sha256).hexdigest()[:24]

def unsub_link(email):
    if not APP_BASE_URL:
        return ""
    from urllib.parse import quote
    return f"{APP_BASE_URL}/unsubscribe?email={quote(email)}&t={unsub_token(email)}"

def valid_email(e): return bool(re.match(r"^[^@\s]+@[^@\s]+\.[^@\s]+$", e or ""))

def send_mail(to, subj, html):
    if not SMTP_USER or not SMTP_PASS:
        return False, "메일 발송 설정(SMTP_USER/SMTP_PASS 또는 EMAIL_ADDRESS/PASSWORD)이 필요합니다."
    try:
        msg = MIMEMultipart("alternative")
        msg["Subject"] = subj; msg["From"] = MAIL_FROM; msg["To"] = to
        msg.attach(MIMEText(html, "html", "utf-8"))
        with smtplib.SMTP(SMTP_HOST, SMTP_PORT) as s:
            s.starttls(); s.login(SMTP_USER, SMTP_PASS)
            s.sendmail(MAIL_FROM, to, msg.as_string())
        return True, "OK"
    except Exception as e: return False, str(e)

def build_newsletter(to=None):
    customs = fetch_customs(); news = fetch_news()[:8]
    summary = by_mineral(customs)[:5]
    rows = "".join(
        f'<tr><td style="padding:8px;border-bottom:1px solid #eee;">{nm}</td>'
        f'<td style="padding:8px;border-bottom:1px solid #eee;text-align:right;color:#c0531a;font-weight:700;">${v:,.0f}</td></tr>'
        for nm, v in summary
    ) or '<tr><td colspan="2" style="color:#999;padding:8px;">데이터 없음</td></tr>'
    news_html = "".join(
        f'<div style="padding:10px 0;border-bottom:1px solid #f0f0f0;">'
        f'<a href="{n.get("언론사링크","#")}" style="color:#1a3a52;text-decoration:none;font-weight:600;font-size:14px;">{n.get("제목","")}</a>'
        f'<div style="color:#999;font-size:12px;margin-top:3px;">{n.get("검색키워드","")} · {n.get("발행일시","")}</div></div>'
        for n in news
    ) or '<div style="color:#999;">뉴스 없음</div>'
    _date  = datetime.now().strftime("%Y년 %m월 %d일")
    _email = EMAIL_ADDRESS or "(메일 미설정)"
    _link  = unsub_link(to) if to else ""
    _unsub = (f' · <a href="{_link}" style="color:#bbb;">수신거부</a>' if _link else "")
    return (
        '<div style="max-width:620px;margin:0 auto;font-family:Malgun Gothic,sans-serif;">'
        '<div style="background:linear-gradient(135deg,#0b1a27,#1a3a52);padding:28px 32px;">'
        '<h1 style="color:#fff;margin:0;font-size:22px;">핵심광물 동향 리포트</h1>'
        f'<p style="color:#7ab3cc;margin:6px 0 0;font-size:13px;">{_date}</p></div>'
        '<div style="padding:24px 32px;background:#fff;border:1px solid #e8e8e8;">'
        '<h2 style="font-size:15px;color:#0b1a27;margin:0 0 12px;">광물별 수입액</h2>'
        f'<table style="width:100%;border-collapse:collapse;font-size:14px;">{rows}</table>'
        '<h2 style="font-size:15px;color:#0b1a27;margin:12px 0 12px;">최신 뉴스</h2>'
        f'{news_html}</div>'
        f'<div style="padding:16px;text-align:center;color:#aaa;font-size:12px;">문의: {_email}{_unsub}</div></div>'
    )


# ═══════════════════════════════════════════════════════════════
#  ② 대시보드 HTML 렌더링
# ═══════════════════════════════════════════════════════════════
def render_dashboard():
    customs  = fetch_customs()
    news     = fetch_news()
    mines    = local_mines()
    subs     = load_subs()
    bm       = by_mineral(customs)
    now      = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    total    = sum(v for _, v in bm)
    top_min  = bm[0][0] if bm else "—"

    cl  = json.dumps([n for n, _ in bm[:7]],  ensure_ascii=False)
    cd  = json.dumps([round(v) for _, v in bm[:7]])
    mineral_imports, imports_unit = by_mineral_country(customs)
    # 국가별 차트: 리튬 수입량 기준 (첫 번째 광물 없으면 전체 합산)
    first_min = next(iter(mineral_imports), None)
    if first_min:
        bc_chart = sorted(mineral_imports[first_min].items(), key=lambda x: x[1], reverse=True)[:7]
    else:
        bc_chart = by_country(customs)[:7]
    top_cntry = bc_chart[0][0] if bc_chart else "—"
    cl2 = json.dumps([n for n, _ in bc_chart],  ensure_ascii=False)
    cd2 = json.dumps([round(v) for _, v in bc_chart])
    korea_imports_js = json.dumps(
        {mn: {cn: round(v) for cn, v in cv.items()} for mn, cv in mineral_imports.items()},
        ensure_ascii=False)
    imports_unit_js = json.dumps(imports_unit)

    usgs_html = "".join(f"""
    <div class="uc">
      <div class="uc-nm">{mn}</div>
      <div class="uc-row"><span class="uc-lb">매장량</span><span class="uc-vl">{info["매장량_만톤"]:,}만톤</span></div>
      <div class="uc-row"><span class="uc-lb">생산량</span><span class="uc-vl">{info["생산량_만톤"]:,}만톤/년</span></div>
      <div class="uc-row"><span class="uc-lb">1위 생산국</span><span class="uc-vl hi">{info["1위국"]}</span></div>
      <div class="uc-src">{info["출처"]}</div>
    </div>""" for mn, info in USGS_DATA.items())

    trade_rows = "".join(f"""<tr>
      <td class="t-nm">{nm}</td>
      <td class="t-num">${v:,.0f}</td>
      <td class="t-bar">
        <div class="bw"><div class="bf" style="width:{min(v/total*100,100) if total else 0:.1f}%"></div></div>
        <span class="bp">{v/total*100:.1f}%</span>
      </td></tr>""" for nm, v in bm[:10]) if bm else \
      '<tr><td colspan="3" class="empty">데이터 없음 — mineral_collector_all.py 를 먼저 실행하거나 API 키를 설정하세요.</td></tr>'

    news_html = "".join(f"""
    <a href="{n.get('언론사링크','#')}" target="_blank" class="nc">
      <span class="nc-kw">{n.get('검색키워드','')}</span>
      <div class="nc-ti">{n.get('제목','')}</div>
      <div class="nc-sm">{n.get('요약','')}</div>
      <div class="nc-dt">{n.get('발행일시','')}</div>
    </a>""" for n in news[:12]) if news else \
    '<div class="empty">뉴스 없음 — API 키를 설정하거나 수집기를 먼저 실행하세요.</div>'

    komir_rows = "".join(f"""<tr>
      <td class="t-nm">{r.get('광물명','')}</td>
      <td class="t-nm">{r.get('국가명','')}</td>
      <td class="t-num">${float(str(r.get('수입금액(달러)',0)).replace(',','') or 0):,.0f}</td>
      <td class="t-num">${float(str(r.get('수출금액(달러)',0)).replace(',','') or 0):,.0f}</td>
    </tr>""" for r in customs[:30]) if customs else \
    '<tr><td colspan="4" class="empty">KOMIR 데이터 없음</td></tr>'

    DASH_OVERRIDE = r"""
/* === K-MINERAL AI 디자인 시스템 리스킨 (대시보드) === */
:root{
  --bg:#131315;--bg2:#1b1b1d;--bg3:#26262a;
  --border:#2a2c2f;--border2:#45464d;
  --text:#e4e2e4;--muted:#c6c6cd;--muted2:#909097;
  --red:#ffb4ab;--red-dim:#2a1512;--red-bright:#ff8a7e;
  --accent:#e9c349;--accent2:#ffe088;
  --blue:#bec6e0;--cyan:#bec6e0;--green:#7ee0a8;
  --mono:'JetBrains Mono','IBM Plex Mono',monospace;
  --sans:'Inter','Noto Sans KR',sans-serif;
}
body{background:var(--bg);padding-left:256px;}
.stat-card,.chart-box,.section,.sub-box,.uc{
  background:rgba(31,31,33,.72)!important;border:1px solid #2f3033!important;border-radius:12px!important;
}
.ticker{background:linear-gradient(90deg,#1b1b1d,#23200f 50%,#1b1b1d)!important;border-bottom:1px solid rgba(233,195,73,.22)!important;}
.ticker-inner{color:#e9c349!important;text-shadow:0 0 8px rgba(233,195,73,.3)!important;}
.nav{position:fixed!important;left:0;top:0;width:256px;height:100vh!important;flex-direction:column;align-items:stretch;
  gap:3px;padding:20px 14px!important;background:var(--bg2)!important;border-right:1px solid var(--border)!important;
  border-bottom:none!important;overflow-y:auto;z-index:200;}
.nav-brand{font-size:14px!important;color:var(--accent)!important;margin:0 0 22px 4px!important;text-shadow:none!important;letter-spacing:.08em!important;}
.nav-brand .sys-dot{background:var(--accent)!important;box-shadow:0 0 8px var(--accent)!important;}
.nav a[data-tab]{display:flex!important;align-items:center;width:100%;padding:11px 14px!important;border-radius:8px!important;
  font-size:13px!important;font-weight:600;color:var(--muted)!important;border:1px solid transparent!important;font-family:var(--mono);}
.nav a[data-tab]:hover{background:rgba(255,255,255,.05)!important;color:var(--text)!important;}
.nav a[data-tab].active{background:#0f172a!important;color:#e9c349!important;text-shadow:none!important;border-color:transparent!important;}
.nav-right{margin-top:auto!important;margin-left:0!important;flex-direction:column;align-items:stretch;gap:10px;padding-top:14px;border-top:1px solid var(--border);}
.nav-time{color:var(--muted)!important;font-size:10px!important;text-align:center;}
.nav-conf{text-align:center;color:#241a00!important;background:var(--accent)!important;border:none!important;font-weight:700!important;padding:10px!important;border-radius:8px!important;}
.nav-conf:hover{opacity:.9;background:var(--accent)!important;}
.sidebar{background:var(--bg2)!important;border-right:1px solid var(--border)!important;}
.stat-card.red{border-color:rgba(255,180,171,.45)!important;background:rgba(147,0,10,.12)!important;}
.sc-val.red{color:var(--red)!important;}
.t-num,.kp-amount{color:var(--accent)!important;}
.bf{background:var(--accent)!important;}
.mode-btn.active,.mineral-btn.active{background:var(--accent)!important;color:#241a00!important;border-color:var(--accent)!important;}
.mineral-btn:hover,.mode-btn:hover{border-color:var(--accent)!important;color:var(--text)!important;}
.nc:hover{border-color:var(--accent)!important;}
.nc-kw,.uc-nm{color:var(--accent)!important;}
.sub-btn{background:var(--accent)!important;color:#241a00!important;}
.sub-btn:hover{background:var(--accent2)!important;}
.sub-input:focus{border-color:var(--accent)!important;}
.kp-title{text-shadow:none!important;color:var(--accent)!important;}
"""
    return f"""<!DOCTYPE html>
<html lang="ko">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>핵심광물 위기 현황 — MINERAL CRISIS DESK</title>
<link href="https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800;900&family=Noto+Sans+KR:wght@400;500;700;900&family=IBM+Plex+Mono:wght@400;600&family=JetBrains+Mono:wght@400;500&display=swap" rel="stylesheet">
<script src="https://cdnjs.cloudflare.com/ajax/libs/Chart.js/4.4.1/chart.umd.min.js"></script>
<link rel="stylesheet" href="https://unpkg.com/leaflet@1.9.4/dist/leaflet.css">
<script src="https://unpkg.com/leaflet@1.9.4/dist/leaflet.js"></script>
<style>
:root{{
  --bg:#050d14;--bg2:#091520;--bg3:#0e1f2e;
  --border:#13283a;--border2:#1c3a52;
  --text:#d8eef7;--muted:#6b8a9c;--muted2:#3e5a6c;
  --red:#ff2200;--red-dim:#2a0703;--red-bright:#ff4422;
  --accent:#ff8800;--accent2:#ffaa33;
  --blue:#00e5ff;--cyan:#00e5ff;--green:#00e676;
  --sans:'Inter','Noto Sans KR',sans-serif;
  --mono:'IBM Plex Mono',monospace;
}}
*{{box-sizing:border-box;margin:0;padding:0}}
body{{background:var(--bg);color:var(--text);font-family:var(--sans);line-height:1.5;height:100vh;display:flex;flex-direction:column;overflow:hidden;}}

/* TICKER */
.ticker{{background:linear-gradient(90deg,#1c0400,#330900 50%,#1c0400);border-bottom:1px solid rgba(255,34,0,.45);height:30px;display:flex;align-items:center;overflow:hidden;flex-shrink:0;}}
.ticker-inner{{white-space:nowrap;animation:ticker 40s linear infinite;font-size:11px;font-weight:600;letter-spacing:.08em;color:#ff5533;font-family:var(--mono);padding-left:100%;text-shadow:0 0 8px rgba(255,34,0,.5);}}
@keyframes ticker{{from{{transform:translateX(0)}}to{{transform:translateX(-50%)}}}}

/* NAV */
.nav{{background:var(--bg2);border-bottom:1px solid var(--border);padding:0 20px;display:flex;align-items:center;gap:4px;flex-shrink:0;height:48px;}}
.nav-brand{{display:flex;align-items:center;gap:8px;font-size:13px;font-weight:800;color:var(--cyan);letter-spacing:.12em;text-transform:uppercase;font-family:var(--mono);margin-right:20px;text-shadow:0 0 12px rgba(0,229,255,.45);}}
.nav-brand .sys-dot{{width:8px;height:8px;border-radius:50%;background:var(--red);box-shadow:0 0 8px var(--red);animation:sys-blink 1.2s steps(2,start) infinite;}}
@keyframes sys-blink{{to{{opacity:.25}}}}
.nav a{{color:var(--muted);text-decoration:none;font-size:12px;font-weight:500;padding:6px 12px;border-radius:3px;transition:.2s;cursor:pointer;border:1px solid transparent;}}
.nav a:hover{{color:var(--cyan);background:var(--bg3);}}
.nav a.active{{color:var(--cyan);background:var(--bg3);border-color:rgba(0,229,255,.35);text-shadow:0 0 8px rgba(0,229,255,.4);}}
.nav-right{{margin-left:auto;display:flex;align-items:center;gap:12px;}}
.nav-time{{font-size:11px;color:var(--cyan);font-family:var(--mono);letter-spacing:.08em;opacity:.85;}}
.nav-conf{{font-size:11px;color:var(--accent);text-decoration:none;font-weight:600;border:1px solid rgba(255,136,0,.35);padding:4px 10px;border-radius:3px;}}
.nav-conf:hover{{background:rgba(255,136,0,.12);}}

/* TAB PANELS */
.tab-panel{{display:none;flex:1;overflow:hidden;}}
.tab-panel.active{{display:flex;}}

/* ── SUPPLY TAB ── */
#tab-supply{{flex-direction:row;overflow:hidden;}}

/* SIDEBAR */
.sidebar{{width:200px;background:var(--bg2);border-right:1px solid var(--border);padding:16px 12px;overflow-y:auto;flex-shrink:0;}}
.sb-section{{margin-bottom:20px;}}
.sb-title{{font-size:9px;font-weight:700;color:var(--muted2);letter-spacing:.15em;text-transform:uppercase;font-family:var(--mono);margin-bottom:8px;}}
.sb-stat{{display:flex;justify-content:space-between;align-items:center;padding:5px 0;border-bottom:1px solid var(--border);}}
.sb-stat-name{{font-size:11px;color:var(--muted);}}
.sb-stat-val{{font-size:11px;font-weight:600;font-family:var(--mono);color:var(--text);}}
.sb-stat-val.amber{{color:var(--accent);}}

/* MAIN CONTENT */
.main{{flex:1;overflow-y:auto;padding:16px;}}
.stat-row{{display:flex;gap:12px;margin-bottom:16px;flex-wrap:wrap;}}
.stat-card{{flex:1;min-width:140px;background:var(--bg2);border:1px solid var(--border);border-radius:6px;padding:12px 16px;}}
.stat-card.red{{border-color:var(--red);background:var(--red-dim);}}
.sc-label{{font-size:10px;color:var(--muted);text-transform:uppercase;letter-spacing:.1em;font-family:var(--mono);}}
.sc-val{{font-size:22px;font-weight:800;color:var(--text);margin:4px 0;}}
.sc-val.red{{color:var(--red-bright);}}
.sc-sub{{font-size:11px;color:var(--muted);}}

/* TABLE */
.section{{background:var(--bg2);border:1px solid var(--border);border-radius:6px;margin-bottom:16px;}}
.sec-head{{padding:10px 16px;border-bottom:1px solid var(--border);font-size:11px;font-weight:700;color:var(--muted);letter-spacing:.1em;text-transform:uppercase;font-family:var(--mono);}}
table{{width:100%;border-collapse:collapse;}}
.t-nm{{padding:8px 16px;font-size:13px;color:var(--text);}}
.t-num{{padding:8px 16px;font-size:12px;font-family:var(--mono);color:var(--accent);text-align:right;}}
.t-bar{{padding:8px 16px;}}
.bw{{background:var(--bg3);height:6px;border-radius:3px;flex:1;margin-bottom:2px;}}
.bf{{background:var(--red);height:6px;border-radius:3px;}}
.bp{{font-size:10px;color:var(--muted);font-family:var(--mono);}}
tr:hover td{{background:var(--bg3);}}
.empty{{padding:24px;color:var(--muted);font-size:13px;text-align:center;}}

/* CHARTS */
.charts-row{{display:flex;gap:12px;margin-bottom:16px;height:200px;flex-shrink:0;}}
.chart-box{{flex:1;background:var(--bg2);border:1px solid var(--border);border-radius:6px;padding:12px;min-width:0;overflow:hidden;display:flex;flex-direction:column;}}
.chart-title{{font-size:10px;color:var(--muted);font-family:var(--mono);text-transform:uppercase;letter-spacing:.1em;margin-bottom:8px;flex-shrink:0;}}

/* ── MAP TAB ── */
#tab-map{{flex-direction:column;overflow:hidden;}}
.map-page{{display:flex;flex-direction:column;height:100%;}}
.map-ctrl{{background:var(--bg2);border-bottom:1px solid var(--border);padding:12px 20px;display:flex;align-items:center;gap:10px;flex-shrink:0;flex-wrap:wrap;}}
.map-ctrl-label{{font-size:10px;color:var(--muted);font-family:var(--mono);text-transform:uppercase;letter-spacing:.1em;}}
.mineral-btn{{padding:5px 14px;border:1px solid var(--border2);background:var(--bg3);color:var(--muted);font-size:12px;font-family:var(--sans);border-radius:20px;cursor:pointer;transition:.2s;}}
.mineral-btn:hover{{color:var(--text);border-color:var(--accent);}}
.mode-btn{{padding:5px 14px;border:1px solid var(--border2);background:var(--bg3);color:var(--muted);font-size:12px;font-family:var(--sans);border-radius:4px;cursor:pointer;transition:.2s;}}
.mode-btn:hover{{color:var(--text);border-color:var(--blue);}}
.mode-btn.active{{background:var(--blue);color:#fff;border-color:var(--blue);font-weight:700;}}
.mineral-btn.active{{background:var(--red);color:#fff;border-color:var(--red);font-weight:700;}}
.map-body{{display:flex;overflow:hidden;align-items:flex-start;}}
#mineral-map{{flex:1;min-width:0;aspect-ratio:16/9;}}
.leaflet-container{{outline:0 !important;}}
.choke-tip{{background:#111;border:1px solid #ff4444;color:#fff;font-size:11px;padding:4px 7px;border-radius:4px;}}
.leaflet-interactive{{outline:0 !important;}}
.leaflet-grab{{outline:0 !important;}}
.route-line {{ stroke-linecap: round; stroke-linejoin: round; }}
/* 루트 라인 — 점선 흐름 애니메이션 (공급국 → 부산 방향) */
.route-flow {{ stroke-dasharray: 7 11; animation: route-dash 1.1s linear infinite; }}
@keyframes route-dash {{ to {{ stroke-dashoffset: -18; }} }}
.leaflet-tooltip.map-tip {{ background:rgba(5,13,20,.95); border:1px solid var(--border2); color:var(--text); font-size:12px; padding:6px 10px; font-family:var(--mono); }}
.leaflet-tooltip.map-tip::before {{ border-right-color:var(--border2); }}

/* 지도 격자 오버레이 (위경도 눈금 느낌) */
.map-grid{{position:absolute;inset:0;z-index:450;pointer-events:none;
  background:
    repeating-linear-gradient(0deg,  transparent 0 79px, rgba(0,229,255,.06) 79px 80px),
    repeating-linear-gradient(90deg, transparent 0 79px, rgba(0,229,255,.06) 79px 80px);
  box-shadow:inset 0 0 120px rgba(0,229,255,.05);}}
.map-grid::after{{content:'';position:absolute;inset:0;
  background:linear-gradient(180deg,transparent 0%,rgba(0,229,255,.025) 50%,transparent 100%);
  background-size:100% 240px;animation:grid-scan 7s linear infinite;}}
@keyframes grid-scan{{from{{background-position:0 -240px}}to{{background-position:0 100vh}}}}

/* 초크포인트 레이더 ping 마커 */
.cp-wrap{{position:relative;cursor:pointer;}}
.cp-core{{position:absolute;border-radius:50%;border:1.5px solid #fff;
  box-shadow:0 0 10px currentColor,0 0 4px currentColor;}}
.cp-core.cp-crit{{animation:cp-blink .85s steps(2,start) infinite;}}
@keyframes cp-blink{{to{{filter:brightness(.45)}}}}
.cp-ring{{position:absolute;border-radius:50%;border:2px solid currentColor;
  animation:cp-ping 1.8s cubic-bezier(0,.5,.4,1) infinite;opacity:0;}}
.cp-ring.r2{{animation-delay:.9s;}}
@keyframes cp-ping{{0%{{transform:scale(.4);opacity:.9}}100%{{transform:scale(2.8);opacity:0}}}}

/* SUPPLY INTEL — HUD 사이드 패널 */
.map-korea-panel{{width:250px;background:linear-gradient(180deg,var(--bg2),#06101a);border-left:1px solid var(--border2);padding:16px;overflow-y:auto;flex-shrink:0;position:relative;}}
.map-korea-panel::before{{content:'';position:absolute;top:0;left:0;right:0;height:2px;background:linear-gradient(90deg,transparent,var(--cyan),transparent);opacity:.7;}}
.kp-title{{font-size:11px;font-weight:700;color:var(--cyan);letter-spacing:.22em;text-transform:uppercase;font-family:var(--mono);margin-bottom:4px;text-shadow:0 0 10px rgba(0,229,255,.4);}}
.kp-sub{{font-size:9px;color:var(--muted2);font-family:var(--mono);letter-spacing:.12em;margin-bottom:12px;}}
.kp-flag{{font-size:22px;margin-bottom:6px;}}
.kp-desc{{font-size:12px;color:var(--muted);line-height:1.6;margin-bottom:12px;}}
.kp-row{{padding:6px 0 7px;border-bottom:1px solid var(--border);opacity:0;transform:translateX(-8px);animation:kp-scan .35s ease forwards;}}
@keyframes kp-scan{{to{{opacity:1;transform:translateX(0)}}}}
.kp-line{{display:flex;justify-content:space-between;align-items:baseline;}}
.kp-country{{font-size:12px;color:var(--text);font-family:var(--mono);}}
.kp-amount{{font-size:11px;font-family:var(--mono);color:var(--accent);}}
.kp-bar{{margin-top:4px;height:4px;background:rgba(0,229,255,.08);border-radius:2px;overflow:hidden;}}
.kp-bar i{{display:block;height:100%;background:linear-gradient(90deg,rgba(0,229,255,.5),var(--cyan));box-shadow:0 0 6px rgba(0,229,255,.6);border-radius:2px;}}
.kp-bar.warn i{{background:linear-gradient(90deg,rgba(255,136,0,.5),var(--accent));box-shadow:0 0 6px rgba(255,136,0,.6);}}
.kp-bar.crit i{{background:linear-gradient(90deg,rgba(255,34,0,.5),var(--red));box-shadow:0 0 6px rgba(255,34,0,.6);}}

/* RISK 배지 */
.risk-badge{{margin-left:auto;font-family:var(--mono);font-size:11px;font-weight:700;letter-spacing:.14em;
  padding:5px 12px;border-radius:3px;border:1px solid;display:inline-flex;align-items:center;gap:7px;}}
.risk-badge .rb-dot{{width:7px;height:7px;border-radius:50%;background:currentColor;box-shadow:0 0 7px currentColor;}}
.risk-badge.high{{color:var(--red-bright);border-color:rgba(255,34,0,.5);background:rgba(255,34,0,.08);}}
.risk-badge.high .rb-dot{{animation:cp-blink .8s steps(2,start) infinite;}}
.risk-badge.medium{{color:var(--accent);border-color:rgba(255,136,0,.5);background:rgba(255,136,0,.08);}}
.risk-badge.low{{color:var(--green);border-color:rgba(0,230,118,.5);background:rgba(0,230,118,.08);}}
.map-legend{{position:absolute;bottom:20px;left:20px;background:rgba(15,15,15,.92);border:1px solid var(--border);border-radius:6px;padding:10px 14px;z-index:1000;font-size:11px;}}
.legend-title{{color:var(--muted);font-family:var(--mono);font-size:9px;text-transform:uppercase;letter-spacing:.1em;margin-bottom:6px;}}
.legend-item{{display:flex;align-items:center;gap:6px;margin-bottom:3px;color:var(--text);}}
.legend-color{{width:16px;height:10px;border-radius:2px;}}

/* ── NEWS TAB ── */
#tab-news{{flex-direction:column;overflow-y:auto;padding:16px;}}
.news-grid{{display:grid;grid-template-columns:repeat(auto-fill,minmax(280px,1fr));gap:12px;}}
.nc{{display:block;background:var(--bg2);border:1px solid var(--border);border-radius:6px;padding:14px;text-decoration:none;transition:.2s;}}
.nc:hover{{border-color:var(--accent);background:var(--bg3);}}
.nc-kw{{font-size:10px;color:var(--accent);font-family:var(--mono);font-weight:600;}}
.nc-ti{{font-size:13px;color:var(--text);font-weight:600;margin:4px 0;line-height:1.4;}}
.nc-sm{{font-size:12px;color:var(--muted);line-height:1.5;}}
.nc-dt{{font-size:10px;color:var(--muted2);font-family:var(--mono);margin-top:6px;}}

/* ── SUBSCRIBE TAB ── */
#tab-subscribe{{flex-direction:column;align-items:center;justify-content:center;padding:40px 20px;}}
.sub-box{{background:var(--bg2);border:1px solid var(--border);border-radius:8px;padding:32px;max-width:480px;width:100%;}}
.sub-title{{font-size:18px;font-weight:700;margin-bottom:8px;}}
.sub-desc{{font-size:13px;color:var(--muted);margin-bottom:20px;line-height:1.6;}}
.sub-input{{width:100%;background:var(--bg3);border:1px solid var(--border2);color:var(--text);padding:10px 14px;border-radius:4px;font-size:14px;outline:none;margin-bottom:10px;}}
.sub-input:focus{{border-color:var(--accent);}}
.sub-btn{{width:100%;background:var(--red);color:#fff;border:none;padding:11px;font-size:14px;font-weight:600;border-radius:4px;cursor:pointer;margin-bottom:8px;transition:.2s;}}
.sub-btn:hover{{background:var(--red-bright);}}
.sub-btn2{{width:100%;background:var(--bg3);color:var(--muted);border:1px solid var(--border2);padding:10px;font-size:13px;border-radius:4px;cursor:pointer;transition:.2s;}}
.sub-btn2:hover{{color:var(--text);border-color:var(--border);}}
.sub-msg{{font-size:13px;margin-top:10px;text-align:center;}}

/* ── KOMIR TAB ── */
#tab-komir{{flex-direction:column;overflow-y:auto;padding:16px;}}
#tab-komir .section{{margin-bottom:16px;}}

/* ── USGS TAB ── */
#tab-usgs{{flex-direction:column;overflow-y:auto;padding:16px;}}
.usgs-grid{{display:grid;grid-template-columns:repeat(auto-fill,minmax(220px,1fr));gap:12px;}}
.uc{{background:var(--bg2);border:1px solid var(--border);border-radius:6px;padding:16px;}}
.uc-nm{{font-size:16px;font-weight:700;color:var(--accent);margin-bottom:10px;}}
.uc-row{{display:flex;justify-content:space-between;padding:5px 0;border-bottom:1px solid var(--border);}}
.uc-lb{{font-size:11px;color:var(--muted);}}
.uc-vl{{font-size:12px;font-weight:600;font-family:var(--mono);color:var(--text);}}
.uc-vl.hi{{color:var(--red-bright);}}
.uc-src{{font-size:10px;color:var(--muted2);margin-top:8px;font-family:var(--mono);}}

/* 공통 유틸 */
.hi{{color:var(--red-bright);}}
.page-title{{font-size:14px;font-weight:700;color:var(--muted);margin-bottom:16px;font-family:var(--mono);text-transform:uppercase;letter-spacing:.1em;}}
</style>
<style>{DASH_OVERRIDE}</style>
</head>
<body>

<!-- TICKER -->
<div class="ticker">
  <div class="ticker-inner">
    ⚠ 핵심광물 공급망 위기 모니터링 &nbsp;|&nbsp;
    리튬 수입 의존도 95% &nbsp;|&nbsp;
    코발트 콩고 집중도 70% &nbsp;|&nbsp;
    희토류 중국 생산 점유율 60% &nbsp;|&nbsp;
    니켈 인도네시아 수출 규제 강화 &nbsp;|&nbsp;
    USGS Mineral Commodity Summaries 2025 기준 &nbsp;|&nbsp;
    ⚠ 핵심광물 공급망 위기 모니터링 &nbsp;|&nbsp;
    리튬 수입 의존도 95% &nbsp;|&nbsp;
    코발트 콩고 집중도 70% &nbsp;|&nbsp;
    희토류 중국 생산 점유율 60% &nbsp;|&nbsp;
    니켈 인도네시아 수출 규제 강화 &nbsp;|&nbsp;
    USGS Mineral Commodity Summaries 2025 기준 &nbsp;|&nbsp;
  </div>
</div>

<!-- NAV -->
<nav class="nav">
  <span class="nav-brand"><span class="sys-dot"></span>MINERAL SUPPLY RISK MONITOR</span>
  <a href="#" class="active" data-tab="supply"    onclick="switchTab('supply',this);return false;">수급 현황</a>
  <a href="#" data-tab="map"       onclick="switchTab('map',this);return false;">글로벌 매장량</a>
  <a href="#" data-tab="news"      onclick="switchTab('news',this);return false;">뉴스 피드</a>
  <a href="#" data-tab="subscribe" onclick="switchTab('subscribe',this);return false;">리포트 구독</a>
  <a href="#" data-tab="komir"     onclick="switchTab('komir',this);return false;">KOMIR</a>
  <a href="#" data-tab="usgs"      onclick="switchTab('usgs',this);return false;">USGS 2025</a>
  <div class="nav-right">
    <span class="nav-time" id="nav-clock">{now}</span>
    <a href="/conference" class="nav-conf">AI 전문가 회의실 →</a>
  </div>
</nav>

<!-- ============================
     TAB: 수급 현황
     ============================ -->
<div id="tab-supply" class="tab-panel active">
  <!-- 사이드바 -->
  <aside class="sidebar">
    <div class="sb-section">
      <div class="sb-title">USGS 2025 매장량</div>
      {''.join(f'<div class="sb-stat"><span class="sb-stat-name">{mn}</span><span class="sb-stat-val amber">{info["매장량_만톤"]:,}만t</span></div>' for mn,info in USGS_DATA.items())}
    </div>
    <div class="sb-section">
      <div class="sb-title">1위 생산국</div>
      {''.join(f'<div class="sb-stat"><span class="sb-stat-name">{mn}</span><span class="sb-stat-val" style="font-size:11px;color:#888">{info["1위국"]}</span></div>' for mn,info in USGS_DATA.items())}
    </div>
    <div class="sb-section">
      <div class="sb-title">구독자</div>
      <div class="sb-stat"><span class="sb-stat-name">등록</span><span class="sb-stat-val">{len(subs)}명</span></div>
    </div>
  </aside>

  <!-- 메인 -->
  <div class="main">
    <!-- 상단 통계 카드 -->
    <div class="stat-row">
      <div class="stat-card red">
        <div class="sc-label">총 수입액</div>
        <div class="sc-val red">${total:,.0f}</div>
        <div class="sc-sub">KOMIR 기준</div>
      </div>
      <div class="stat-card">
        <div class="sc-label">최대 수입 광물</div>
        <div class="sc-val">{top_min}</div>
        <div class="sc-sub">수입액 1위</div>
      </div>
      <div class="stat-card">
        <div class="sc-label">최대 수입국</div>
        <div class="sc-val">{top_cntry}</div>
        <div class="sc-sub">국가별 1위</div>
      </div>
      <div class="stat-card">
        <div class="sc-label">뉴스</div>
        <div class="sc-val">{len(news)}</div>
        <div class="sc-sub">수집된 기사</div>
      </div>
    </div>

    <!-- 차트 -->
    <div class="charts-row">
      <div class="chart-box">
        <div class="chart-title">광물별 수입액 (상위 7)</div>
        <div style="flex:1;position:relative;min-height:0;"><canvas id="chartMin"></canvas></div>
      </div>
      <div class="chart-box">
        <div class="chart-title">국가별 수입액 (상위 7)</div>
        <div style="flex:1;position:relative;min-height:0;"><canvas id="chartCnt"></canvas></div>
      </div>
    </div>

    <!-- 광물별 수입 테이블 -->
    <div class="section">
      <div class="sec-head">광물별 수입 현황</div>
      <table>
        <tbody>{trade_rows}</tbody>
      </table>
    </div>
  </div>
</div>

<!-- ============================
     TAB: 글로벌 매장량
     ============================ -->
<div id="tab-map" class="tab-panel">
  <div class="map-page">
    <div class="map-ctrl">
      <span class="map-ctrl-label">모드:</span>
      <button class="mode-btn active" id="modeReserves" onclick="setMode('reserves',this)">🌍 매장량</button>
      <button class="mode-btn" id="modeRoutes" onclick="setMode('routes',this)">🚢 수입 루트</button>
      <span class="map-ctrl-label" style="margin-left:12px;">광물:</span>
      <button class="mineral-btn active" onclick="selectMineral('리튬',this)">리튬</button>
      <button class="mineral-btn" onclick="selectMineral('코발트',this)">코발트</button>
      <button class="mineral-btn" onclick="selectMineral('니켈',this)">니켈</button>
      <button class="mineral-btn" onclick="selectMineral('흑연',this)">흑연</button>
      <button class="mineral-btn" onclick="selectMineral('희토류',this)">희토류</button>
      <button class="mineral-btn" onclick="selectMineral('망간',this)">망간</button>
      <button class="mineral-btn" onclick="selectMineral('전체',this)">전체</button>
      <span class="risk-badge low" id="risk-badge"><span class="rb-dot"></span><span id="risk-badge-txt">RISK: —</span></span>
    </div>
    <div class="map-body">
      <div id="mineral-map" style="position:relative;">
        <div class="map-grid"></div>
        <div class="map-legend" id="map-legend">
          <div class="legend-title">매장량 규모</div>
          <div class="legend-item"><div class="legend-color" style="background:#ff2200"></div> 초대형 (30%+)</div>
          <div class="legend-item"><div class="legend-color" style="background:#ff6600"></div> 대형 (15-30%)</div>
          <div class="legend-item"><div class="legend-color" style="background:#ffaa00"></div> 중형 (5-15%)</div>
          <div class="legend-item"><div class="legend-color" style="background:#ffdd44"></div> 소형 (1-5%)</div>
          <div class="legend-item"><div class="legend-color" style="background:#88bb44"></div> 미량 (&lt;1%)</div>
          <div class="legend-item"><div class="legend-color" style="background:#00ccff"></div> 🇰🇷 한국 (수입의존)</div>
        </div>
        <!-- 초크포인트 뉴스 패널 -->
        <div id="choke-panel" style="
          display:none;position:absolute;top:10px;right:10px;z-index:1500;
          width:320px;max-height:70vh;overflow-y:auto;
          background:rgba(10,10,18,0.96);border:1px solid #333;
          border-radius:8px;padding:14px;
          box-shadow:0 4px 24px rgba(0,0,0,0.7);
          scrollbar-width:thin;scrollbar-color:#333 transparent;
        "></div>
      </div><!-- /#mineral-map -->
      <div class="map-korea-panel">
        <div class="kp-title">▮ SUPPLY INTEL</div>
        <div class="kp-sub">KR-BUSAN // IMPORT FEED</div>
        <div class="kp-flag">🇰🇷</div>
        <div class="kp-desc" id="kp-desc">리튬을 선택하면 한국의 주요 수입국 정보가 표시됩니다.</div>
        <div id="kp-rows"></div>
      </div>
    </div>
  </div>
</div>

<!-- ============================
     TAB: 뉴스 피드
     ============================ -->
<div id="tab-news" class="tab-panel">
  <div class="page-title">최신 뉴스 피드</div>
  <div class="news-grid">
    {news_html}
  </div>
</div>

<!-- ============================
     TAB: 리포트 구독
     ============================ -->
<div id="tab-subscribe" class="tab-panel">
  <div class="sub-box">
    <div class="sub-title">핵심광물 동향 리포트 구독</div>
    <div class="sub-desc">
      매일 최신 광물 수급 현황과 글로벌 이슈를 이메일로 받아보세요.<br>
      현재 <strong>{len(subs)}명</strong>이 구독 중입니다.
    </div>
    <input id="sub-email" class="sub-input" type="email" placeholder="이메일 주소 입력">
    <button class="sub-btn" onclick="doSubscribe()">구독 신청</button>
    <button class="sub-btn2" onclick="doSendNow()">지금 바로 받기 (1회)</button>
    <div class="sub-msg" id="sub-msg"></div>
  </div>
</div>

<!-- ============================
     TAB: KOMIR
     ============================ -->
<div id="tab-komir" class="tab-panel">
  <div class="page-title">KOMIR — 광종별 국가별 수출입 현황</div>
  <div class="section">
    <div class="sec-head">수출입 데이터 (최근 30건)</div>
    <table>
      <thead>
        <tr>
          <td class="t-nm" style="color:#888;font-size:11px;">광물명</td>
          <td class="t-nm" style="color:#888;font-size:11px;">국가</td>
          <td class="t-num" style="color:#888;font-size:11px;">수입액(USD)</td>
          <td class="t-num" style="color:#888;font-size:11px;">수출액(USD)</td>
        </tr>
      </thead>
      <tbody>{komir_rows}</tbody>
    </table>
  </div>
</div>

<!-- ============================
     TAB: USGS 2025
     ============================ -->
<div id="tab-usgs" class="tab-panel">
  <div class="page-title">USGS Mineral Commodity Summaries 2025</div>
  <div class="usgs-grid">
    {usgs_html}
  </div>
</div>

<script>
// ── 탭 전환 ──────────────────────────────────────────────────
function switchTab(name, el) {{
  document.querySelectorAll('.tab-panel').forEach(p => p.classList.remove('active'));
  document.querySelectorAll('.nav a[data-tab]').forEach(a => a.classList.remove('active'));
  document.getElementById('tab-' + name).classList.add('active');
  if (el) el.classList.add('active');
  if (name === 'map' && !window._mapInited) initMap();
}}

// 다른 페이지(회의실 등)에서 #map / #news 등으로 들어오면 해당 탭으로 이동
(function(){{
  var h = (location.hash || '').replace('#','');
  var valid = ['supply','map','news','subscribe','komir','usgs'];
  if (valid.indexOf(h) >= 0) {{
    switchTab(h, document.querySelector('.nav a[data-tab="' + h + '"]'));
  }}
}})();

// ── 실시간 시스템 시계 ───────────────────────────────────────
setInterval(() => {{
  const el = document.getElementById('nav-clock');
  if (!el) return;
  const d = new Date();
  const p = n => String(n).padStart(2, '0');
  el.textContent = `${{d.getFullYear()}}-${{p(d.getMonth()+1)}}-${{p(d.getDate())}} ${{p(d.getHours())}}:${{p(d.getMinutes())}}:${{p(d.getSeconds())}} KST ● LIVE`;
}}, 1000);

// ── Chart.js 차트 ────────────────────────────────────────────
const CHART_OPTS = {{
  responsive: true, maintainAspectRatio: false,
  plugins: {{ legend: {{ display: false }}, tooltip: {{ callbacks: {{ label: ctx => ' $' + ctx.raw.toLocaleString() }} }} }},
  scales: {{
    x: {{ ticks: {{ color: '#888', font: {{ size: 10 }} }}, grid: {{ color: '#13283a' }} }},
    y: {{ ticks: {{ color: '#888', font: {{ size: 10 }}, callback: v => '$' + (v/1e6).toFixed(1)+'M' }}, grid: {{ color: '#13283a' }} }}
  }}
}};
const isCntTon = {imports_unit_js} === '톤';
const CHART_CNT_OPTS = {{
  responsive: true, maintainAspectRatio: false,
  plugins: {{ legend: {{ display: false }}, tooltip: {{ callbacks: {{ label: ctx => ' ' + ctx.raw.toLocaleString() + (isCntTon ? ' 톤' : '') }} }} }},
  scales: {{
    x: {{ ticks: {{ color: '#888', font: {{ size: 10 }} }}, grid: {{ color: '#13283a' }} }},
    y: {{ ticks: {{ color: '#888', font: {{ size: 10 }}, callback: v => isCntTon ? (v/1e3).toFixed(0)+'K톤' : '$'+(v/1e6).toFixed(1)+'M' }}, grid: {{ color: '#13283a' }} }}
  }}
}};
new Chart(document.getElementById('chartMin'), {{
  type: 'bar',
  data: {{ labels: {cl}, datasets: [{{ data: {cd}, backgroundColor: '#ff2200', borderRadius: 3 }}] }},
  options: CHART_OPTS
}});
new Chart(document.getElementById('chartCnt'), {{
  type: 'bar',
  data: {{ labels: {cl2}, datasets: [{{ data: {cd2}, backgroundColor: '#00e5ff', borderRadius: 3 }}] }},
  options: CHART_CNT_OPTS
}});

// ── 구독 ─────────────────────────────────────────────────────
function doSubscribe() {{
  const email = document.getElementById('sub-email').value;
  fetch('/subscribe', {{ method:'POST', headers:{{'Content-Type':'application/json'}}, body:JSON.stringify({{email}}) }})
    .then(r=>r.json()).then(d=>{{
      const m = document.getElementById('sub-msg');
      m.style.color = d.ok ? '#39c96e' : '#ff4444';
      m.textContent = d.message;
    }});
}}
function doSendNow() {{
  const email = document.getElementById('sub-email').value;
  fetch('/send_now', {{ method:'POST', headers:{{'Content-Type':'application/json'}}, body:JSON.stringify({{email}}) }})
    .then(r=>r.json()).then(d=>{{
      const m = document.getElementById('sub-msg');
      m.style.color = d.ok ? '#39c96e' : '#ff4444';
      m.textContent = d.message;
    }});
}}

// ── 세계 지도 ─────────────────────────────────────────────────
const WORLD_RESERVES = {{
  '리튬': {{
    '칠레':          {{iso:'CHL', reserves:930,  share:33, unit:'만톤'}},
    '호주':          {{iso:'AUS', reserves:570,  share:20, unit:'만톤'}},
    '아르헨티나':    {{iso:'ARG', reserves:220,  share:8,  unit:'만톤'}},
    '중국':          {{iso:'CHN', reserves:150,  share:5,  unit:'만톤'}},
    '캐나다':        {{iso:'CAN', reserves:93,   share:3,  unit:'만톤'}},
    '브라질':        {{iso:'BRA', reserves:55,   share:2,  unit:'만톤'}},
    '짐바브웨':      {{iso:'ZWE', reserves:23,   share:1,  unit:'만톤'}},
  }},
  '코발트': {{
    '콩고민주공화국':{{iso:'COD', reserves:400,  share:51, unit:'만톤'}},
    '호주':          {{iso:'AUS', reserves:150,  share:19, unit:'만톤'}},
    '필리핀':        {{iso:'PHL', reserves:26,   share:3,  unit:'만톤'}},
    '쿠바':          {{iso:'CUB', reserves:50,   share:6,  unit:'만톤'}},
    '카메룬':        {{iso:'CMR', reserves:29,   share:4,  unit:'만톤'}},
    '러시아':        {{iso:'RUS', reserves:25,   share:3,  unit:'만톤'}},
    '잠비아':        {{iso:'ZMB', reserves:27,   share:3,  unit:'만톤'}},
  }},
  '니켈': {{
    '인도네시아':    {{iso:'IDN', reserves:2100, share:42, unit:'만톤'}},
    '필리핀':        {{iso:'PHL', reserves:480,  share:10, unit:'만톤'}},
    '러시아':        {{iso:'RUS', reserves:750,  share:15, unit:'만톤'}},
    '호주':          {{iso:'AUS', reserves:210,  share:4,  unit:'만톤'}},
    '캐나다':        {{iso:'CAN', reserves:200,  share:4,  unit:'만톤'}},
    '중국':          {{iso:'CHN', reserves:280,  share:6,  unit:'만톤'}},
    '뉴칼레도니아':  {{iso:'NCL', reserves:370,  share:7,  unit:'만톤'}},
  }},
  '흑연': {{
    '중국':          {{iso:'CHN', reserves:5200, share:35, unit:'만톤'}},
    '브라질':        {{iso:'BRA', reserves:700,  share:5,  unit:'만톤'}},
    '탄자니아':      {{iso:'TZA', reserves:800,  share:5,  unit:'만톤'}},
    '마다가스카르':  {{iso:'MDG', reserves:150,  share:1,  unit:'만톤'}},
    '모잠비크':      {{iso:'MOZ', reserves:700,  share:5,  unit:'만톤'}},
    '인도':          {{iso:'IND', reserves:800,  share:5,  unit:'만톤'}},
    '러시아':        {{iso:'RUS', reserves:1000, share:7,  unit:'만톤'}},
  }},
  '희토류': {{
    '중국':          {{iso:'CHN', reserves:4400, share:38, unit:'만톤'}},
    '베트남':        {{iso:'VNM', reserves:2200, share:19, unit:'만톤'}},
    '브라질':        {{iso:'BRA', reserves:2100, share:18, unit:'만톤'}},
    '러시아':        {{iso:'RUS', reserves:210,  share:2,  unit:'만톤'}},
    '인도':          {{iso:'IND', reserves:69,   share:1,  unit:'만톤'}},
    '호주':          {{iso:'AUS', reserves:480,  share:4,  unit:'만톤'}},
    '미국':          {{iso:'USA', reserves:180,  share:2,  unit:'만톤'}},
  }},
  '망간': {{
    '남아프리카공화국':{{iso:'ZAF', reserves:40000,share:35,unit:'만톤'}},
    '우크라이나':    {{iso:'UKR', reserves:14000,share:12, unit:'만톤'}},
    '호주':          {{iso:'AUS', reserves:25000,share:22, unit:'만톤'}},
    '브라질':        {{iso:'BRA', reserves:2700, share:2,  unit:'만톤'}},
    '인도':          {{iso:'IND', reserves:5900, share:5,  unit:'만톤'}},
    '중국':          {{iso:'CHN', reserves:4400, share:4,  unit:'만톤'}},
    '가봉':          {{iso:'GAB', reserves:2500, share:2,  unit:'만톤'}},
  }},
}};
const KOREA_IMPORTS = {korea_imports_js};
const KOREA_IMPORTS_UNIT = {imports_unit_js};

let _map = null;
let _geojson = null;
let _currentMineral = '리튬';
let _layer = null;

function initMap() {{
  window._mapInited = true;
  _map = L.map('mineral-map', {{
    center: [30, 15], zoom: 2, minZoom: 2, maxZoom: 6,
    zoomControl: true, attributionControl: false,
    worldCopyJump: false,
    maxBounds: [[-62, -180], [80, 180]],
    maxBoundsViscosity: 1.0,
  }});
  L.tileLayer('https://{{s}}.basemaps.cartocdn.com/dark_all/{{z}}/{{x}}/{{y}}{{r}}.png', {{
    subdomains: 'abcd', maxZoom: 19, noWrap: true,
  }}).addTo(_map);

  fetch('/api/geojson')
    .then(r => r.json())
    .then(data => {{
      if (!data || data.type === 'Topology' || data.error) {{
        // server returned bad data — retry direct fetch
        return fetch('https://raw.githubusercontent.com/datasets/geo-countries/master/data/countries.geojson').then(r=>r.json());
      }}
      return data;
    }})
    .then(data => {{
      if (!data || !data.features) throw new Error('유효한 GeoJSON 없음');
      _geojson = data;
      renderMineralLayer(_currentMineral);
    }})
    .catch(e => {{
      document.getElementById('kp-desc').textContent = 'GeoJSON 로드 실패: ' + e.message;
    }});
}}

function getColor(share) {{
  if (share >= 30) return '#ff2200';
  if (share >= 15) return '#ff6600';
  if (share >=  5) return '#ffaa00';
  if (share >=  1) return '#ffdd44';
  return '#88bb44';
}}

function renderMineralLayer(mineral) {{
  if (_layer) {{ _map.removeLayer(_layer); _layer = null; }}
  const data = WORLD_RESERVES[mineral] || {{}};
  const isoMap = {{}};
  for (const [country, info] of Object.entries(data)) {{
    isoMap[info.iso] = {{ country, ...info }};
  }}

  _layer = L.geoJSON(_geojson, {{
    style: feature => {{
      const p = feature.properties;
      const iso = p['ISO3166-1-Alpha-3'] || p.ISO_A3 || p.ADM0_A3 || p.iso_a3 || '';
      if (iso === 'KOR') return {{ fillColor:'#00ccff', fillOpacity:.9, color:'#fff', weight:2 }};
      const d = isoMap[iso];
      if (d) return {{ fillColor: getColor(d.share), fillOpacity:.85, color:'#222', weight:.5 }};
      return {{ fillColor:'#2a2a2a', fillOpacity:.5, color:'#333', weight:.3 }};
    }},
    onEachFeature: (feature, layer) => {{
      const p = feature.properties;
      const iso = p['ISO3166-1-Alpha-3'] || p.ISO_A3 || p.ADM0_A3 || p.iso_a3 || '';
      const name = p.name || p.ADMIN || p.NAME || iso;
      if (iso === 'KOR') {{
        layer.bindTooltip(`<b>🇰🇷 대한민국</b><br>수입 의존국 — 국내 매장량 없음`, {{className:'map-tip'}});
      }} else if (isoMap[iso]) {{
        const d = isoMap[iso];
        layer.bindTooltip(`<b>${{d.country}}</b><br>매장량: ${{d.reserves.toLocaleString()}}${{d.unit}}<br>점유율: ${{d.share}}%`, {{className:'map-tip'}});
      }} else {{
        layer.bindTooltip(name, {{className:'map-tip'}});
      }}
    }}
  }}).addTo(_map);

  updateKoreaPanel(mineral);
}}

// 공급 집중도 기반 RISK 배지 — 상위 1개국 점유율 30%+ HIGH / 15~30% MEDIUM / 미만 LOW
function setRiskBadge(mineral) {{
  const badge = document.getElementById('risk-badge');
  const txt   = document.getElementById('risk-badge-txt');
  if (!badge) return;
  const data = WORLD_RESERVES[mineral] || {{}};
  const shares = Object.values(data).map(d => d.share);
  if (!shares.length) {{
    badge.className = 'risk-badge low';
    txt.textContent = 'RISK: —';
    return;
  }}
  const top = Math.max(...shares);
  const lvl = top >= 30 ? 'high' : top >= 15 ? 'medium' : 'low';
  badge.className = 'risk-badge ' + lvl;
  txt.textContent = `RISK: ${{lvl.toUpperCase()}} · TOP ${{top}}%`;
}}

// HUD 스타일 행 — 스캔 등장 애니메이션 + 퍼센트 바
function kpRow(i, name, val, pct) {{
  const cls = pct >= 30 ? 'crit' : pct >= 15 ? 'warn' : '';
  const bar = pct != null ? `<div class="kp-bar ${{cls}}"><i style="width:${{Math.min(pct,100)}}%"></i></div>` : '';
  return `<div class="kp-row" style="animation-delay:${{i*70}}ms">
    <div class="kp-line"><span class="kp-country">${{name}}</span><span class="kp-amount">${{val}}</span></div>
    ${{bar}}
  </div>`;
}}

function updateKoreaPanel(mineral) {{
  const desc = document.getElementById('kp-desc');
  const rows = document.getElementById('kp-rows');
  const data = WORLD_RESERVES[mineral] || {{}};
  const topCountries = Object.entries(data).sort((a,b) => b[1].share - a[1].share).slice(0,5);

  desc.textContent = `${{mineral}} 글로벌 매장량 현황 — 한국은 전량 수입에 의존합니다.`;
  setRiskBadge(mineral);

  let html = topCountries.map(([c, d], i) =>
    kpRow(i, c, `${{d.share}}% (${{d.reserves.toLocaleString()}}만t)`, d.share)).join('');

  // 선택 광물의 국가별 수입량 표시 (톤 or USD)
  const mineralImports = KOREA_IMPORTS[mineral] || {{}};
  const importEntries = Object.entries(mineralImports).sort((a,b) => b[1]-a[1]);
  if (importEntries.length > 0) {{
    const isTon = KOREA_IMPORTS_UNIT === '톤';
    const maxV = Math.max(...importEntries.map(([,v]) => v));
    html += `<div style="margin-top:12px;margin-bottom:2px;font-size:9px;color:var(--cyan);opacity:.7;font-family:var(--mono);text-transform:uppercase;letter-spacing:.18em;">▮ KR IMPORT ${{isTon ? 'VOLUME' : 'VALUE'}} — ${{mineral}}</div>`;
    html += importEntries.map(([c, v], i) =>
      kpRow(topCountries.length + i, c,
        isTon ? v.toLocaleString() + ' 톤' : '$' + v.toLocaleString(),
        Math.round(v / maxV * 100))).join('');
  }}
  rows.innerHTML = html;
}}


let _mapMode = 'reserves';
let _routeLayers = [];

function setMode(mode, btn) {{
  _mapMode = mode;
  document.querySelectorAll('.mode-btn').forEach(b => b.classList.remove('active'));
  btn.classList.add('active');
  if (mode === 'reserves') {{
    clearRoutes();
    if (_geojson) renderMineralLayer(_currentMineral);
    document.getElementById('map-legend').style.display = 'block';
  }} else {{
    clearLayer();
    document.getElementById('map-legend').style.display = 'none';
    renderRoutes(_currentMineral);
  }}
}}

function clearRoutes() {{
  _routeLayers.forEach(l => _map.removeLayer(l));
  _routeLayers = [];
  _chokeLayers.forEach(l => _map.removeLayer(l));
  _chokeLayers = [];
  const old = document.getElementById('route-svg-overlay');
  if (old) old.remove();
}}

function clearLayer() {{
  if (_layer) {{ _map.removeLayer(_layer); _layer = null; }}
  clearRoutes();
}}

// ── 실제 해상 항로 경유지 기반 라우팅 ──────────────────────────
// ── 주요 해상 초크포인트 (정확한 좌표) ──────────────────────────────────────
const WP = {{
  BUSAN:      [35.1,   129.1],   // 부산항
  // ── 동북아 ──
  KOREA_STR:  [34.5,   129.5],   // 대한해협
  TSUSHIMA:   [34.2,   129.3],   // 쓰시마 해협
  E_CHINA_SEA:[30.0,   125.0],   // 동중국해
  YELLOW_SEA: [33.5,   124.0],   // 황해 남부
  JEJU_S:     [33.0,   127.0],   // 제주 남방
  TSUGARU:    [41.5,   140.7],   // 쓰가루 해협
  TAIWAN_STR: [24.0,   119.0],   // 대만해협 (해협 중앙 수역)
  TAIWAN_E:   [24.5,   122.7],   // 대만 동측 수역
  LUZON_STR:  [20.8,   121.8],   // 루손 해협 (바시 채널)
  // ── 남중국해 ──
  SCS_N:      [19.0,   116.0],   // 남중국해 북부
  SCS_MID:    [13.0,   112.0],   // 남중국해 중부
  SCS_S:      [6.0,    109.5],   // 남중국해 남부 (나투나 동측)
  MALACCA:    [1.5,    104.0],   // 말라카 해협 ★ (싱가포르 해협)
  MAL_MID:    [3.0,    100.5],   // 말라카 해협 중앙
  MAL_NW:     [5.5,     96.0],   // 말라카 북서 입구 (아체 북방)
  KARIMATA:   [-1.5,   108.8],   // 카리마타 해협
  LOMBOK:     [-8.7,   115.8],   // 롬복 해협 ★
  MAKASSAR:   [-2.0,   117.8],   // 마카사르 해협
  CELEBES:    [4.2,    120.5],   // 셀레베스해
  SULU:       [8.5,    119.8],   // 술루해
  MINDORO:    [12.8,   120.1],   // 민도로 해협
  TORRES:     [-10.5,  142.0],   // 토레스 해협 ★
  ARAFURA:    [-9.8,   135.5],   // 아라푸라해
  BANDA:      [-5.8,   128.5],   // 반다해
  MOLUCCA:    [0.5,    126.0],   // 몰루카해
  PHIL_E_S:   [7.0,    127.8],   // 필리핀해 남부 (민다나오 동측)
  PHIL_E_M:   [14.0,   127.0],   // 필리핀해 중부
  // ── 인도양 ──
  DONDRA:     [5.5,     80.6],   // 스리랑카 남방 (돈드라곶)
  IND_MID:    [-5.0,    75.0],   // 인도양 중부
  IND_SW:     [-15.0,   65.0],   // 인도양 남서
  ARABIAN:    [9.0,     63.0],   // 아라비아해
  // ── 중동·홍해 ──
  HORMUZ:     [26.5,    56.3],   // 호르무즈 해협 ★
  OMAN_G:     [24.0,    59.5],   // 오만만 (호르무즈 출구)
  ADEN:       [12.3,    46.5],   // 아덴만
  ADEN_E:     [13.3,    51.3],   // 아덴만 동측 출구 (소코트라 북방)
  BAB:        [12.8,    43.3],   // 밥엘만데브 ★
  RED_S:      [15.0,    41.5],   // 홍해 남부
  RED_MID:    [20.0,    38.5],   // 홍해 중부
  RED_N:      [27.5,    34.0],   // 홍해 북부
  SUEZ_S:     [29.9,    32.6],   // 수에즈 남단 ★
  SUEZ_N:     [31.3,    32.3],   // 수에즈 북단 (포트사이드)
  // ── 지중해·유럽 ──
  MED_E:      [34.5,    27.0],   // 동지중해 (크레타 남방)
  MED_IONIAN: [35.5,    21.0],   // 이오니아해 남부 (펠로폰네소스 남방)
  MED_C:      [36.2,    13.8],   // 중지중해 (시칠리아-튀니지 사이)
  MED_TUNIS:  [37.5,    10.8],   // 튀니지 북방 수역
  MED_W:      [37.0,     0.5],   // 서지중해 (알제리 북방)
  GIBRALTAR:  [35.9,    -5.3],   // 지브롤터 ★
  CADIZ_OFF:  [36.2,    -7.3],   // 카디스만 외해
  LISBON_OFF: [38.4,    -9.9],   // 리스본 외해
  FINISTERRE: [43.4,   -10.0],   // 피니스테레곶 외해
  BISCAY:     [45.5,    -6.0],   // 비스케이만
  CHANNEL_W:  [49.6,    -3.5],   // 영불해협 서측
  DOVER:      [51.1,     1.5],   // 도버 해협
  NORTH_SEA:  [55.0,     4.0],   // 북해
  SKAGERRAK:  [57.8,     9.0],   // 스카게라크 해협
  KATTEGAT:   [56.8,    11.5],   // 카테가트
  ORESUND:    [55.7,    12.7],   // 외레순 해협
  BALTIC_S:   [55.4,    15.0],   // 발트해 남부 (보른홀름 북방)
  BOSPHORUS:  [41.15,   29.05],  // 보스포루스
  MARMARA:    [40.75,   28.0],   // 마르마라해
  DARDANELLES:[40.05,   26.2],   // 다르다넬스
  AEGEAN_S:   [38.0,    25.0],   // 에게해 중남부
  // ── 아프리카 연안 ──
  CAPE:       [-34.3,   18.5],   // 희망봉 ★
  GUINEA_G:   [2.0,      5.0],   // 기니만
  ANGOLA_OFF: [-12.0,    8.5],   // 앙골라 외해
  NAMIBIA_OFF:[-25.0,   11.5],   // 나미비아 외해
  // ── 대서양·아메리카 ──
  S_ATL:      [-36.0,  -10.0],   // 남대서양 (희망봉 서방)
  N_ATL_MID:  [38.0,   -28.0],   // 북대서양 (아조레스)
  CARIB:      [14.0,   -68.0],   // 카리브해
  YUCATAN:    [21.8,   -85.5],   // 유카탄 해협
  PANAMA_P:   [8.9,    -79.5],   // 파나마 태평양 측 ★
  PANAMA_A:   [9.4,    -79.9],   // 파나마 대서양 측
  CAPE_HORN:  [-55.8,  -67.2],   // 케이프혼 ★
}};

// ── 초크포인트 목록 (지도에 마커로 표시) ───────────────────────────────────
const CHOKEPOINTS = [
  {{ key:'MALACCA',  name:'말라카 해협',   pos:WP.MALACCA,   color:'#ff3333', risk:'critical',
     reason:'전 세계 해상 물동량 약 25% 통과. 봉쇄 시 한국 에너지·광물 수입 절반 이상 차질' }},
  {{ key:'HORMUZ',   name:'호르무즈 해협', pos:WP.HORMUZ,    color:'#ff3333', risk:'critical',
     reason:'중동산 광물·원유의 유일한 출구. 이란과 서방 갈등 시 수시로 봉쇄 위협 발생' }},
  {{ key:'BAB',      name:'밥엘만데브',    pos:WP.BAB,       color:'#ff3333', risk:'critical',
     reason:'2024~25년 후티 반군 상선 공격으로 주요 선사 우회 운항 중. 수에즈 루트 전체 위협' }},
  {{ key:'SUEZ_S',   name:'수에즈 운하',   pos:[30.4, 32.4], color:'#ff8800', risk:'high',
     reason:'유럽·아프리카~아시아 최단 경로. 봉쇄 시 희망봉 우회로 운임 2~3배 급등' }},
  {{ key:'GIBRALTAR',name:'지브롤터 해협', pos:WP.GIBRALTAR, color:'#ff8800', risk:'high',
     reason:'대서양~지중해 관문. 유럽발 광물 수입 루트의 필수 경유지. 영국-스페인 영유권 분쟁 잠재' }},
  {{ key:'CAPE',     name:'희망봉',        pos:WP.CAPE,      color:'#ffaa00', risk:'medium',
     reason:'수에즈 봉쇄 대안 경로. 우회 시 운항 기간 10~14일 추가. 강풍·高파도 위험' }},
  {{ key:'PANAMA_P', name:'파나마 운하',   pos:WP.PANAMA_P,  color:'#ffaa00', risk:'medium',
     reason:'2023~24년 엘니뇨 가뭄으로 통항 40% 감소. 기후변화로 반복 위험. 미국 영향력 확대 논란' }},
  {{ key:'LOMBOK',   name:'롬복 해협',     pos:WP.LOMBOK,    color:'#ffcc44', risk:'low',
     reason:'말라카 우회 대안. 수심 깊어 대형 선박 통과 가능. 인도네시아 정세에 의존' }},
  {{ key:'TORRES',   name:'토레스 해협',   pos:WP.TORRES,    color:'#ffcc44', risk:'low',
     reason:'호주 동부~아시아 경로. 수심 얕고 암초 많아 항법 주의. 호주산 니켈·코발트 수입에 활용' }},
  {{ key:'CAPE_HORN',name:'케이프혼',      pos:WP.CAPE_HORN, color:'#ffcc44', risk:'low',
     reason:'칠레산 리튬·구리의 아시아행 경로. 강풍·너울로 운항 위험. 파나마 막힐 경우 대안 경로' }},
];

let _chokeLayers = [];
function renderChokepoints() {{
  _chokeLayers.forEach(l => _map.removeLayer(l));
  _chokeLayers = [];
  CHOKEPOINTS.forEach(cp => {{
    const RISK_SIZE = {{critical:14, high:11, medium:9, low:7}};
    const sz = RISK_SIZE[cp.risk] || 10;
    const box = sz * 3;                      // ping 링 확장 공간
    const off = (box - sz) / 2;
    const crit = cp.risk === 'critical' ? ' cp-crit' : '';
    // 레이더 ping 애니메이션 마커 (critical은 점멸)
    const icon = L.divIcon({{
      className: '',
      html: `<div class="cp-wrap" style="width:${{box}}px;height:${{box}}px;color:${{cp.color}};">
        <span class="cp-ring"    style="left:${{off}}px;top:${{off}}px;width:${{sz}}px;height:${{sz}}px;"></span>
        <span class="cp-ring r2" style="left:${{off}}px;top:${{off}}px;width:${{sz}}px;height:${{sz}}px;"></span>
        <span class="cp-core${{crit}}" style="left:${{off}}px;top:${{off}}px;width:${{sz}}px;height:${{sz}}px;background:${{cp.color}};"></span>
      </div>`,
      iconSize:[box,box], iconAnchor:[box/2,box/2],
    }});
    const m = L.marker(cp.pos, {{icon, zIndexOffset:1000}}).addTo(_map);
    m.bindTooltip(
      `<b style="color:${{cp.color}}">⚠ ${{cp.name}}</b><br><small style="color:#aaa">클릭하면 관련 뉴스 보기</small>`,
      {{permanent:false, direction:'top', className:'choke-tip'}}
    );
    m.on('click', () => openChokeNews(cp));
    _chokeLayers.push(m);
    const lbl = L.divIcon({{
      className: '',
      html: `<span style="color:${{cp.color}};font-size:10px;font-weight:bold;
             white-space:nowrap;text-shadow:0 0 4px #000,0 0 2px #000,1px 1px 0 #000;">${{cp.name}}</span>`,
      iconSize:[90,16], iconAnchor:[-4, 8],
    }});
    const lm = L.marker(cp.pos, {{icon:lbl, interactive:false, zIndexOffset:999}}).addTo(_map);
    _chokeLayers.push(lm);
  }});
}}

function openChokeNews(cp) {{
  const panel = document.getElementById('choke-panel');
  const RISK_KO = {{critical:'🔴 위험 (Critical)', high:'🟠 높음 (High)', medium:'🟡 보통 (Medium)', low:'🟢 낮음 (Low)'}};
  panel.innerHTML = `
    <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:10px;">
      <span style="font-size:15px;font-weight:bold;color:#00e5ff;font-family:var(--mono);">⚓ ${{cp.name}}</span>
      <button onclick="document.getElementById('choke-panel').style.display='none'"
        style="background:none;border:none;color:#888;font-size:18px;cursor:pointer;line-height:1;">✕</button>
    </div>
    <div style="font-size:11px;color:${{cp.color}};margin-bottom:6px;font-weight:bold;">${{RISK_KO[cp.risk] || cp.risk}}</div>
    <div style="font-size:11px;color:#ccc;margin-bottom:12px;line-height:1.5;">${{cp.reason}}</div>
    <div id="choke-news-list" style="font-size:11px;color:#aaa;">뉴스 로딩 중...</div>
  `;
  panel.style.display = 'block';

  fetch('/api/chokepoint-news?key=' + cp.key)
    .then(r => r.json())
    .then(data => {{
      const list = document.getElementById('choke-news-list');
      if (!data.articles || data.articles.length === 0) {{
        list.innerHTML = '<div style="color:#666;padding:8px 0;">관련 뉴스가 없습니다</div>';
        return;
      }}
      list.innerHTML = '<div style="color:#888;margin-bottom:6px;border-bottom:1px solid #333;padding-bottom:4px;">📰 관련 뉴스</div>' +
        data.articles.map(a => `
          <div style="margin-bottom:8px;padding:6px;background:#1a1a1a;border-radius:4px;border-left:2px solid ${{cp.color}};">
            <div style="margin-bottom:2px;">
              <a href="${{a.link}}" target="_blank"
                style="color:#ddd;text-decoration:none;font-size:11px;line-height:1.4;"
                onmouseover="this.style.color='#00ccff'" onmouseout="this.style.color='#ddd'">
                ${{a.title}}
              </a>
            </div>
            ${{a.desc ? `<div style="color:#666;font-size:10px;margin-top:2px;line-height:1.3;">${{a.desc}}</div>` : ''}}
            <div style="color:#555;font-size:10px;margin-top:3px;">${{a.date}} · ${{a.kw}}</div>
          </div>
        `).join('');
    }})
    .catch(() => {{
      const list = document.getElementById('choke-news-list');
      list.innerHTML = '<div style="color:#666;">뉴스를 불러올 수 없습니다</div>';
    }});
}}

// ── 공통 구간 헬퍼 ──────────────────────────────────────────────
// 말라카(싱가포르 해협) → 남중국해 북상 → 대만해협 → 부산
function _malaccaToKorea() {{ return [WP.MALACCA, WP.SCS_S, WP.SCS_MID, WP.SCS_N, WP.TAIWAN_STR, WP.E_CHINA_SEA, WP.BUSAN]; }}
// 인도양 → 말라카 북서 입구 → 해협 종주 → 부산
function _indianToMalacca() {{ return [WP.MAL_NW, WP.MAL_MID, ..._malaccaToKorea()]; }}
// 수에즈 남단 → 홍해 종주 → 밥엘만데브 → 아덴만 → 인도양 → 말라카
function _suezToKorea() {{ return [WP.SUEZ_S, WP.RED_N, WP.RED_MID, WP.RED_S, WP.BAB, WP.ADEN, WP.ADEN_E, WP.ARABIAN, WP.DONDRA, ..._indianToMalacca()]; }}
// 희망봉 → 마다가스카르 남방 → 인도양 횡단 → 말라카
function _capeToKorea() {{ return [WP.CAPE, [-36.5, 25], [-30, 45], [-18, 60], [-6, 76], ..._indianToMalacca()]; }}
// 지브롤터 → 지중해 종주 → 포트사이드
function _gibToSuez() {{ return [WP.GIBRALTAR, WP.MED_W, WP.MED_TUNIS, WP.MED_C, WP.MED_IONIAN, WP.MED_E, WP.SUEZ_N]; }}
// 도버 → 영불해협 → 비스케이 → 이베리아 연안 (지브롤터 직전까지)
function _channelToGib() {{ return [WP.DOVER, WP.CHANNEL_W, WP.BISCAY, WP.FINISTERRE, WP.LISBON_OFF, WP.CADIZ_OFF]; }}
// 서태평양 (필리핀해) → 루손해협 → 대만 동측 → 부산
function _philSeaToKorea() {{ return [[10, 141], [16.5, 130], WP.LUZON_STR, WP.TAIWAN_E, WP.E_CHINA_SEA, WP.BUSAN]; }}

// ── 지역별 정확한 뱃길 ────────────────────────────────────────────────────────
// 반환값: 항로 세그먼트 배열. 대부분 1개, 태평양 횡단(날짜변경선 통과)은 2개.
function getSeaRoute(lat, lng) {{
  const B = WP.BUSAN, o = [lat, lng];

  // ── 동북아 근거리 ──────────────────────────────────────────────────────────

  // 러시아 극동 (블라디보스토크·나홋카) → 동해 남하 직항
  if (lng > 127 && lng < 145 && lat > 42)
    return [[o, [40, 132], [37, 130.8], WP.KOREA_STR, B]];

  // 일본 홋카이도·도호쿠 동해측 → 동해 남하
  if (lng >= 137 && lat > 38 && lat <= 46)
    return [[o, [41.5, 138.5], [39, 134], [36.5, 130.8], WP.KOREA_STR, B]];

  // 일본 혼슈 태평양측 (도쿄·나고야) → 기이반도·시코쿠 남방 → 오스미 해협 → 쓰시마
  if (lng > 135 && lng < 143 && lat > 32 && lat <= 38)
    return [[o, [34.2, 140.0], [33.0, 136.5], [32.3, 132.6], [30.9, 130.7],
             [31.8, 128.9], [33.4, 128.3], WP.TSUSHIMA, B]];

  // 일본 규슈·세토내해 서부 → 쓰시마 직항
  if (lng > 129 && lng <= 135 && lat > 31 && lat <= 36)
    return [[o, WP.TSUSHIMA, B]];

  // 대만 → 동중국해 직항 (대만해협 우회 불필요)
  if (lng > 119.5 && lng < 122.5 && lat > 21 && lat <= 26)
    return [[o, [26.5, 122.3], WP.E_CHINA_SEA, B]];

  // 중국 동부 (상하이·칭다오) → 동중국해 횡단 → 제주 남방
  if (lng > 117 && lng < 127 && lat > 28)
    return [[o, [32, 124.5], WP.JEJU_S, B]];

  // 중국 남부 (홍콩·광저우·샤먼) → 대만해협 북상
  if (lng > 105 && lng < 122 && lat > 18 && lat <= 28)
    return [[o, [21.5, 116.5], WP.TAIWAN_STR, WP.E_CHINA_SEA, B]];

  // ── 동남아 ─────────────────────────────────────────────────────────────────

  // 필리핀 (마닐라) → 루손 서측 연안 → 루손해협 → 대만 동측
  if (lng >= 117 && lng < 127 && lat > 5 && lat <= 19)
    return [[o, [15.8, 119.3], [19.5, 120.2], WP.LUZON_STR, WP.TAIWAN_E, WP.E_CHINA_SEA, B]];

  // 미얀마 (양곤) → 안다만해 남하 → 말라카
  if (lng > 92 && lng < 99.5 && lat > 9 && lat < 23)
    return [[o, [13.5, 95.5], ..._indianToMalacca()]];

  // 베트남·캄보디아·태국 → 남중국해 직행 북상 (말라카 경유 없음)
  if (lng >= 99.5 && lng < 110 && lat > 5 && lat < 23)
    return [[o, ...(lng < 105 ? [[10.5, 101.5], [8.0, 104.5]] : []),
             [9.5, 109], WP.SCS_MID, WP.SCS_N, WP.TAIWAN_STR, WP.E_CHINA_SEA, B]];

  // 말레이시아·싱가포르·수마트라 → 말라카 해협 → 남중국해
  if (lng > 95 && lng < 112 && lat > -6 && lat < 7)
    return [[o, ..._malaccaToKorea()]];

  // 인도네시아 자바 (자카르타) → 자바해 → 카리마타 해협 → 남중국해
  if (lng >= 105 && lng < 117 && lat >= -9 && lat < 0)
    return [[o, [-5.7, 107.6], WP.KARIMATA, WP.SCS_S, WP.SCS_MID, WP.SCS_N,
             WP.TAIWAN_STR, WP.E_CHINA_SEA, B]];

  // 인도네시아 동부 (술라웨시·칼리만탄 동부) → 마카사르 → 술루해 → 민도로
  if (lng >= 117 && lng < 132 && lat >= -10 && lat <= 5)
    return [[o, WP.MAKASSAR, WP.CELEBES, WP.SULU, WP.MINDORO, WP.SCS_N,
             WP.TAIWAN_STR, WP.E_CHINA_SEA, B]];

  // 호주 서부 (포트헤들랜드 — 리튬·철광석) → 롬복 → 마카사르 → 술루 → 북상
  if (lng >= 105 && lng <= 132 && lat < 0)
    return [[o, [-13.5, 116.5], WP.LOMBOK, WP.MAKASSAR, WP.CELEBES, WP.SULU,
             WP.MINDORO, WP.SCS_N, WP.TAIWAN_STR, WP.E_CHINA_SEA, B]];

  // 호주 동부·파푸아뉴기니 → 토레스 해협 → 반다해 → 필리핀 동측 북상
  if (lng > 132 && lng < 155 && lat < 0)
    return [[o, WP.TORRES, WP.ARAFURA, WP.BANDA, WP.MOLUCCA, WP.PHIL_E_S,
             WP.PHIL_E_M, WP.LUZON_STR, WP.TAIWAN_E, WP.E_CHINA_SEA, B]];

  // 뉴칼레도니아·뉴질랜드 → 산호해 동측 북상 → 필리핀해
  if (lng >= 155 && lat < 0)
    return [[o, [-15, 162], [-5, 158], [3, 150], ..._philSeaToKorea()]];

  // ── 인도양 ─────────────────────────────────────────────────────────────────

  // 스리랑카 → 돈드라곶 → 말라카
  if (lng > 78 && lng < 83 && lat > 4 && lat < 11)
    return [[o, WP.DONDRA, ..._indianToMalacca()]];

  // 방글라데시·인도 동안 → 벵골만 남하 (안다만 동측)
  if (lng >= 80 && lng < 98 && lat > 5 && lat < 25)
    return [[o, [15, 88], [10, 94.5], ..._indianToMalacca()]];

  // 인도 서안·파키스탄 → 연안 남하 → 돈드라곶
  if (lng > 60 && lng < 80 && lat > 4 && lat < 28)
    return [[o, [14, 71.5], [7.8, 76.2], WP.DONDRA, ..._indianToMalacca()]];

  // ── 중동 ───────────────────────────────────────────────────────────────────

  // 페르시아만 (사우디·이라크·쿠웨이트·카타르·UAE·이란) → 호르무즈 필수 통과
  if (lng > 44 && lng < 60 && lat >= 22 && lat < 32)
    return [[o, WP.HORMUZ, WP.OMAN_G, WP.ARABIAN, WP.DONDRA, ..._indianToMalacca()]];

  // 아덴만 연안 (예멘·지부티·에티오피아) → 아라비아해 동진
  if (lng >= 41 && lng < 55 && lat >= 10 && lat < 19)
    return [[o, WP.ADEN, WP.ADEN_E, WP.ARABIAN, WP.DONDRA, ..._indianToMalacca()]];

  // ── 아프리카 ────────────────────────────────────────────────────────────────

  // 동아프리카 북부 (케냐·탄자니아) → 마다가스카르 북방 → 인도양 횡단
  if (lng > 33 && lng < 46 && lat > -12 && lat < 8)
    return [[o, [-7, 48], [-4, 62], WP.DONDRA, ..._indianToMalacca()]];

  // 마다가스카르 동안 → 인도양 직행
  if (lng >= 44 && lng < 52 && lat >= -28 && lat <= -12)
    return [[o, [-15, 55], [-8, 68], [-2, 80], ..._indianToMalacca()]];

  // 모잠비크 남부 → 마다가스카르 남방 우회 → 인도양
  if (lng > 30 && lng < 44 && lat >= -28 && lat <= -12)
    return [[o, [-27, 40], [-26.5, 47.5], [-15, 62], [-5, 77], ..._indianToMalacca()]];

  // 남아공 (더반) → 인도양 동진 (희망봉 경유 불필요 — 동안 출항)
  if (lat <= -24 && lng > 14 && lng < 37)
    return [[o, [-33, 33], [-28, 48], [-15, 62], [-5, 77], ..._indianToMalacca()]];

  // 서아프리카 기니만 (나이지리아·가나·코트디부아르·카메룬) → 남하 → 희망봉 ★
  if (lng >= -20 && lng < 16 && lat >= 1 && lat < 12)
    return [[o, WP.GUINEA_G, WP.ANGOLA_OFF, WP.NAMIBIA_OFF, ..._capeToKorea()]];

  // 중서부 아프리카 (콩고·가봉·앙골라·나미비아) → 연안 남하 → 희망봉 ★
  if (lng >= 5 && lng < 20 && lat >= -24 && lat < 1.5)
    return [[o, WP.ANGOLA_OFF, WP.NAMIBIA_OFF, ..._capeToKorea()]];

  // ── 유럽·지중해 ─────────────────────────────────────────────────────────────

  // 흑해 (우크라이나) → 보스포루스 → 에게해 → 수에즈
  if (lng >= 27 && lng <= 42 && lat > 41.3 && lat < 48)
    return [[o, [43.5, 30.3], WP.BOSPHORUS, WP.MARMARA, WP.DARDANELLES,
             WP.AEGEAN_S, WP.MED_E, WP.SUEZ_N, ..._suezToKorea()]];

  // 튀르키예 (이스탄불·이즈미르) → 마르마라 → 에게해 → 수에즈
  if (lng > 26 && lng < 45 && lat >= 36 && lat <= 41.3)
    return [[o, WP.MARMARA, WP.DARDANELLES, WP.AEGEAN_S, WP.MED_E, WP.SUEZ_N, ..._suezToKorea()]];

  // 그리스 → 키티라 해협 → 크레타 남방 → 수에즈
  if (lng > 19 && lng <= 26 && lat > 34 && lat < 42)
    return [[o, [36.2, 23.2], [34.8, 25.0], WP.MED_E, WP.SUEZ_N, ..._suezToKorea()]];

  // 이탈리아·아드리아 → 시칠리아 남동 우회 → 이오니아해 → 수에즈
  if (lng > 8 && lng <= 19 && lat > 36 && lat < 46)
    return [[o, [38.8, 14.0], [36.3, 15.6], WP.MED_IONIAN, WP.MED_E, WP.SUEZ_N, ..._suezToKorea()]];

  // 이집트·레반트 연안 → 포트사이드 직행
  if (lng > 25 && lng < 36 && lat >= 30 && lat < 36)
    return [[o, WP.SUEZ_N, ..._suezToKorea()]];

  // 알제리·튀니지 연안 → 동지중해 동진
  if (lng >= 2 && lng <= 12 && lat > 33 && lat < 38.5)
    return [[o, [37.8, 7.5], WP.MED_TUNIS, WP.MED_C, WP.MED_IONIAN, WP.MED_E,
             WP.SUEZ_N, ..._suezToKorea()]];

  // 이베리아·모로코 (스페인·포르투갈) → 지브롤터 ★ → 지중해 → 수에즈
  if (lng >= -12 && lng < 2 && lat >= 30 && lat < 44)
    return [[o, WP.CADIZ_OFF, ..._gibToSuez(), ..._suezToKorea()]];

  // 발트해 (폴란드·핀란드) → 외레순 → 스카게라크 → 도버 → 지브롤터 → 수에즈
  if (lng > 13 && lng < 31 && lat > 53)
    return [[o, ...(lat > 57 ? [[58.8, 21.5], [56.3, 18.8]] : [[55.6, 16.5]]),
             WP.BALTIC_S, WP.ORESUND, WP.KATTEGAT, WP.SKAGERRAK, WP.NORTH_SEA,
             ..._channelToGib(), ..._gibToSuez(), ..._suezToKorea()]];

  // 스칸디나비아 (노르웨이·스웨덴 서안) → 스카게라크 → 북해
  if (lng >= 4 && lng <= 13 && lat > 55.5)
    return [[o, [58.2, 10.3], WP.SKAGERRAK, WP.NORTH_SEA,
             ..._channelToGib(), ..._gibToSuez(), ..._suezToKorea()]];

  // 서유럽 (영국·독일·네덜란드·벨기에·프랑스) → 도버 → 비스케이 → 지브롤터 ★
  if (lng > -12 && lng < 14 && lat >= 44)
    return [[o, ...(lat > 52.5 && lng > 5 ? [[54.6, 6.8]] : []),
             ...(lat > 50.2 ? [WP.DOVER] : []),
             WP.CHANNEL_W, WP.BISCAY, WP.FINISTERRE, WP.LISBON_OFF, WP.CADIZ_OFF,
             ..._gibToSuez(), ..._suezToKorea()]];

  // ── 아메리카 ────────────────────────────────────────────────────────────────

  // 북미 서해안 (밴쿠버·LA) / 멕시코 태평양측 → 북태평양 횡단 (날짜변경선 분할)
  if (lng <= -100 && lat > 5) {{
    if (lat > 40)   // 캐나다 BC
      return [[o, [51, -140], [52, -160], [51, -180]],
              [[51, 180], [48, 168], [44, 152], WP.TSUGARU, [40, 135.5], [36.8, 130.6], B]];
    if (lat > 25)   // 미국 서부
      return [[o, [38, -135], [45, -157], [47, -180]],
              [[47, 180], [45, 165], [42.5, 150], WP.TSUGARU, [40, 135.5], [36.8, 130.6], B]];
    // 멕시코 (만사니요) — 일본 남방 항로
    return [[o, [24, -125], [30, -150], [33, -180]],
            [[33, 180], [33.5, 160], [32, 140], [31.8, 133], [30.9, 130.7],
             [31.8, 128.9], [33.4, 128.3], WP.TSUSHIMA, B]];
  }}

  // 남미 서해안 (칠레·페루) → 남태평양 횡단 (날짜변경선 분할)
  if (lng < -65 && lat < 5)
    return [[o, [-22, -95], [-26, -125], [-29, -155], [-30, -180]],
            [[-30, 180], [-26, 169], [-15, 162], [-5, 158], [3, 150], ..._philSeaToKorea()]];

  // 카리브·멕시코만·남미 북안 (쿠바·콜롬비아) → 파나마 운하 ★ → 태평양 횡단
  if (lng < -30 && lat >= 5)
    return [[o, ...(lat > 17 ? [WP.YUCATAN, [17.5, -81.5]] : [[12, -78.5]]),
             WP.PANAMA_A, WP.PANAMA_P, [5, -95], [2, -125], [0, -155], [-1, -180]],
            [[-1, 180], [1, 162], [5, 150], ..._philSeaToKorea()]];

  // 남미 동해안 (브라질·아르헨티나) → 남대서양 → 희망봉 ★
  if (lng < -30 && lat < 5)
    return [[o, [-30, -38], WP.S_ATL, ..._capeToKorea()]];

  // 중앙아시아·몽골 (내륙국 — 중국 횡단 철송 후 보하이만 출항)
  if (lng > 55 && lng < 125 && lat > 33)
    return [[o, [38.5, 119.5], [37.8, 122.8], [35.5, 124.3], [33.9, 126.4], B]];

  // 기본 (인도양 경유)
  return [[o, WP.IND_MID, ..._indianToMalacca()]];
}}

// Cardinal spline 보간 (텐션 낮춰 곡선 오버슈트로 인한 육지 침범 방지)
function smoothRoute(pts, steps) {{
  if (pts.length < 2) return pts;
  const K = 0.55;   // 접선 스케일 (1.0 = Catmull-Rom, 작을수록 직선에 가까움)
  const out = [];
  for (let i=0; i<pts.length-1; i++) {{
    const p0 = pts[Math.max(0, i-1)];
    const p1 = pts[i];
    const p2 = pts[i+1];
    const p3 = pts[Math.min(pts.length-1, i+2)];
    const m1lat = K*(p2[0]-p0[0])/2, m1lng = K*(p2[1]-p0[1])/2;
    const m2lat = K*(p3[0]-p1[0])/2, m2lng = K*(p3[1]-p1[1])/2;
    for (let t=0; t<steps; t++) {{
      const tt = t/steps, tt2 = tt*tt, tt3 = tt2*tt;
      const h00 =  2*tt3 - 3*tt2 + 1;
      const h10 =      tt3 - 2*tt2 + tt;
      const h01 = -2*tt3 + 3*tt2;
      const h11 =      tt3 -    tt2;
      out.push([
        h00*p1[0] + h10*m1lat + h01*p2[0] + h11*m2lat,
        h00*p1[1] + h10*m1lng + h01*p2[1] + h11*m2lng,
      ]);
    }}
  }}
  out.push(pts[pts.length-1]);
  return out;
}}

function renderRoutes(mineral) {{
  clearRoutes();
  renderChokepoints();
  const apiMineral = (mineral === '리튬'||mineral === '코발트'||mineral === '니켈'||
                      mineral === '흑연'||mineral === '희토류'||mineral === '망간') ? mineral : '';
  const kp = document.getElementById('kp-desc');
  kp.textContent = '수입 루트 로딩 중...';

  fetch('/api/trade-map?mineral=' + encodeURIComponent(apiMineral))
    .then(r => r.json())
    .then(data => {{
      const routes = data.routes;
      if (!routes || !routes.length) {{
        kp.textContent = '수입 데이터 없음';
        return;
      }}
      const isTon = data.unit === '톤';
      const BUSAN = [35.1, 129.07];
      const maxAmt = Math.max(...routes.map(r => r.amount));

      routes.forEach(r => {{
        const from = [r.lat, r.lng];
        const segments = getSeaRoute(r.lat, r.lng);   // 세그먼트 배열 (날짜변경선 분할 지원)
        const width = Math.max(1.5, (r.amount / maxAmt) * 8);
        const opacity = 0.45 + (r.amount / maxAmt) * 0.55;
        const amtStr = isTon ? r.amount.toLocaleString() + ' 톤' : '$' + r.amount.toLocaleString();

        segments.forEach(seg => {{
          const pts = smoothRoute(seg, 14);

          // Glow line (thicker, transparent)
          const glow = L.polyline(pts, {{
            color: '#00e5ff',
            weight: width * 2.6,
            opacity: 0.10,
            smoothFactor: 1,
          }}).addTo(_map);
          _routeLayers.push(glow);

          // Main route line — 점선 흐름 애니메이션 (CSS stroke-dashoffset)
          const line = L.polyline(pts, {{
            color: '#00e5ff',
            weight: width,
            opacity: opacity,
            smoothFactor: 1,
            className: 'route-line route-flow',
          }}).addTo(_map);
          line.bindTooltip(
            `<b>${{r.country}}</b><br>수입${{isTon?'량':'액'}}: ${{amtStr}}<br>비중: ${{r.share}}%`,
            {{sticky: true, className: 'map-tip'}}
          );
          _routeLayers.push(line);
        }});

        // Origin marker
        const dot = L.circleMarker(from, {{
          radius: Math.max(4, width * 1.2),
          fillColor: '#ff8800',
          color: '#ffb347',
          weight: 1.5,
          fillOpacity: 0.85,
        }}).addTo(_map);
        dot.bindTooltip(`<b>${{r.country}}</b><br>${{isTon ? r.amount.toLocaleString()+' 톤' : '$'+r.amount.toLocaleString()}}`, {{className:'map-tip'}});
        _routeLayers.push(dot);
      }});

      // Korea marker
      const korea = L.circleMarker(BUSAN, {{
        radius: 10,
        fillColor: '#00e5ff',
        color: '#fff',
        weight: 2,
        fillOpacity: 1,
      }}).addTo(_map);
      korea.bindTooltip('🇰🇷 부산항 (수입 거점)', {{permanent: false, className:'map-tip'}});
      _routeLayers.push(korea);

      // Update side panel
      updateRoutePanel(routes, mineral, isTon);
    }})
    .catch(e => {{ kp.textContent = '로드 실패: ' + e.message; }});
}}

function updateRoutePanel(routes, mineral, isTon) {{
  const desc = document.getElementById('kp-desc');
  const rowsEl = document.getElementById('kp-rows');
  const label = mineral && mineral !== '전체' ? mineral : '전체 광물';
  desc.textContent = label + ' 수입 루트 — 공급국 → 부산항';
  setRiskBadge(mineral);
  const top = routes.slice(0, 10);
  rowsEl.innerHTML = top.map((r, i) => {{
    const amt = isTon
      ? (r.amount >= 1e6 ? (r.amount/1e6).toFixed(1)+'M 톤' : r.amount.toLocaleString()+' 톤')
      : (r.amount >= 1e9 ? (r.amount/1e9).toFixed(1)+'B USD' : '$'+r.amount.toLocaleString());
    return kpRow(i, r.country, amt, r.share);
  }}).join('');
}}

function selectMineral(mineral, btn) {{
  _currentMineral = mineral;
  document.querySelectorAll('.mineral-btn').forEach(b => b.classList.remove('active'));
  if (btn) btn.classList.add('active');
  if (_mapMode === 'routes') {{
    renderRoutes(mineral);
  }} else {{
    if (_geojson) renderMineralLayer(mineral);
    else updateKoreaPanel(mineral);
  }}
}}

</script>
</body>
</html>"""


# ═══════════════════════════════════════════════════════════════
#  ③ 라우트
# ═══════════════════════════════════════════════════════════════
@app.route("/")
def index(): return Response(render_dashboard(), mimetype="text/html")

@app.route("/api/geojson")
def api_geojson():
    """GeoJSON 프록시 — 브라우저 CORS 우회"""
    c = cache_get("geojson")
    if c: return Response(c, mimetype="application/json")
    # GeoJSON-only sources (NOT TopoJSON — browser layer renders GeoJSON only)
    urls = [
        "https://raw.githubusercontent.com/datasets/geo-countries/master/data/countries.geojson",
        "https://raw.githubusercontent.com/holtzy/D3-graph-gallery/master/DATA/world.geojson",
    ]
    for url in urls:
        try:
            r = requests.get(url, timeout=30)
            if r.status_code == 200 and '"type"' in r.text and 'Feature' in r.text:
                cache_set("geojson", r.text, ttl=86400)  # cache 24h — rarely changes
                return Response(r.text, mimetype="application/json")
        except: continue
    return jsonify({"error": "GeoJSON 로드 실패"}), 500


@app.route("/api/chokepoint-news")
def api_chokepoint_news():
    """초크포인트별 관련 뉴스 반환"""
    key = request.args.get("key", "")
    CHOKE_INFO = {
        "MALACCA":  {"name":"말라카 해협",
                     "kw":["말라카 해협 봉쇄", "말라카 해협 해운 위기", "남중국해 해상 분쟁"],
                     "risk":"critical",
                     "reason":"전 세계 해상 물동량의 약 25%가 통과. 봉쇄 시 한국 에너지·광물 수입의 절반 이상에 영향"},
        "HORMUZ":   {"name":"호르무즈 해협",
                     "kw":["호르무즈 해협 봉쇄", "이란 해협 통항 위협", "페르시아만 선박 억류"],
                     "risk":"critical",
                     "reason":"중동산 원유·광물의 유일한 출구. 이란과 서방 갈등 시 수시로 봉쇄 위협 발생"},
        "BAB":      {"name":"밥엘만데브 해협",
                     "kw":["홍해 후티 선박 공격", "홍해 해운 운항 중단", "예멘 후티 반군 해상"],
                     "risk":"critical",
                     "reason":"2024년 후티 반군의 상선 공격으로 세계 주요 선사들이 우회 운항 중. 수에즈 루트 전체 위협"},
        "SUEZ_S":   {"name":"수에즈 운하",
                     "kw":["수에즈 운하 통항 차질", "수에즈 운하 봉쇄 우회", "홍해 수에즈 해운"],
                     "risk":"high",
                     "reason":"유럽·북아프리카~아시아 최단 경로. 봉쇄 시 희망봉 우회로 운임 2~3배 상승"},
        "GIBRALTAR":{"name":"지브롤터 해협",
                     "kw":["지브롤터 해협 선박", "지중해 해운 통항", "지브롤터 분쟁"],
                     "risk":"medium",
                     "reason":"대서양~지중해 연결. 유럽발 광물 수입 루트의 관문. 분쟁 가능성은 낮으나 전략적 요충"},
        "CAPE":     {"name":"희망봉",
                     "kw":["희망봉 우회 항로 해운", "수에즈 대체 케이프 항로", "남아프리카 해상 운임"],
                     "risk":"high",
                     "reason":"수에즈 봉쇄 시 필수 대안 경로. 거리·시간·비용 증가. 남아프리카 정세 안정적이나 날씨 위험"},
        "PANAMA_P": {"name":"파나마 운하",
                     "kw":["파나마 운하 통항 제한", "파나마 운하 가뭄 수위", "파나마 운하 해운"],
                     "risk":"high",
                     "reason":"2023~24년 엘니뇨로 수위 저하 → 통항 선박 수 40% 감소. 기후변화로 반복 위험"},
        "LOMBOK":   {"name":"롬복 해협",
                     "kw":["롬복 해협 선박 통항", "인도네시아 해협 해운", "말라카 대체 항로"],
                     "risk":"medium",
                     "reason":"말라카 우회 대안 경로. 수심이 깊어 대형 선박 통과 가능. 인도네시아 정세 의존"},
        "TORRES":   {"name":"토레스 해협",
                     "kw":["토레스 해협 선박 항법", "호주 광물 해상 수출 항로", "파푸아뉴기니 해역 해운"],
                     "risk":"low",
                     "reason":"호주 동부~아시아 경로. 수심 얕고 암초 多. 주로 호주 광물(니켈·코발트) 수입에 활용"},
        "CAPE_HORN":{"name":"케이프혼",
                     "kw":["케이프혼 항로 선박", "남미 칠레 해상 수출", "케이프혼 기상 운항"],
                     "risk":"medium",
                     "reason":"남미 서해안(칠레산 리튬·구리)의 아시아행 주요 경로. 극단적 기상으로 운항 위험 높음"},
    }
    info = CHOKE_INFO.get(key)
    if not info:
        return jsonify({"error": "unknown key"}), 400

    cache_key = f"choke_news_{key}"
    cached = cache_get(cache_key)
    if cached:
        return jsonify(cached)

    # 해운/지정학과 무관한 노이즈 기사 필터 키워드
    NOISE_WORDS = ["팝업", "브랜드", "쇼핑", "맛집", "패션", "뷰티", "아이돌", "콘서트",
                   "드라마", "영화", "게임", "인테리어", "부동산", "주식", "코인", "NFT",
                   "롯데월드", "에버랜드", "면세점", "카페", "레스토랑"]
    # 해운 관련 키워드가 하나라도 있으면 통과
    MARITIME_WORDS = ["해협", "운하", "선박", "해운", "항로", "봉쇄", "통항", "수출", "수입",
                      "화물", "항만", "후티", "이란", "분쟁", "위기", "우회", "가뭄", "제재",
                      "해상", "광물", "원유", "LNG", "컨테이너", "벌크선", "해적"]

    def is_maritime_relevant(title, desc):
        text = (title + " " + desc).lower()
        if any(w in text for w in NOISE_WORDS): return False
        return any(w in text for w in MARITIME_WORDS)

    articles = []
    if not NAVER_CLIENT_ID.startswith("여기에"):
        hdrs = {"X-Naver-Client-Id": NAVER_CLIENT_ID, "X-Naver-Client-Secret": NAVER_CLIENT_SECRET}
        seen = set()
        for kw in info["kw"]:
            try:
                r = requests.get("https://openapi.naver.com/v1/search/news.json",
                    headers=hdrs, params={"query": kw, "display": 6, "sort": "date"}, timeout=8)
                if r.status_code != 200: continue
                for it in r.json().get("items", []):
                    lnk = it.get("originallink", "") or it.get("link", "")
                    if lnk in seen: continue
                    seen.add(lnk)
                    title = clean(it.get("title", ""))
                    desc  = clean(it.get("description", ""))
                    if not is_maritime_relevant(title, desc): continue
                    try: dt = datetime.strptime(it.get("pubDate",""), "%a, %d %b %Y %H:%M:%S +0900").strftime("%m/%d %H:%M")
                    except: dt = it.get("pubDate","")[:10]
                    articles.append({
                        "title": title,
                        "desc":  desc[:120],
                        "link":  lnk,
                        "date":  dt,
                        "kw":    kw,
                    })
                    if len(articles) >= 8: break
            except: continue
            time.sleep(0.1)
            if len(articles) >= 8: break

    result = {"name": info["name"], "risk": info["risk"], "reason": info["reason"], "articles": articles[:8]}
    cache_set(cache_key, result, ttl=300)
    return jsonify(result)

@app.route("/api/trade-map")
def api_trade_map():
    """광물별 수입 루트 데이터 반환"""
    mineral = request.args.get("mineral", "")
    customs = fetch_customs()
    from collections import defaultdict
    mc = defaultdict(dict)
    has_tons = False
    for r in customs:
        mn = _MINERAL_ALIAS.get(r.get("광물명","").strip(), r.get("광물명","").strip())
        cn = r.get("국가명","").strip()
        try: t = float(str(r.get("수입중량(톤)", 0) or 0).replace(",",""))
        except: t = 0
        try: v = float(str(r.get("수입금액(달러)", 0) or 0).replace(",",""))
        except: v = 0
        if t > 0: has_tons = True
        val = t if t > 0 else v
        if mn and cn and cn != "-" and val > 0:
            mc[mn][cn] = mc[mn].get(cn, 0) + val
    unit = "톤" if has_tons else "USD"

    if mineral and mineral in mc:
        target = {mineral: mc[mineral]}
    elif mineral == "전체" or not mineral:
        # 전체 합산
        total = {}
        for mn, data in mc.items():
            for cn, v in data.items():
                total[cn] = total.get(cn, 0) + v
        target = {"전체": total}
    else:
        target = {"전체": {}}

    routes = []
    for mn, data in target.items():
        max_v = max(data.values()) if data else 1
        for cn, v in sorted(data.items(), key=lambda x: -x[1])[:20]:
            coords = COUNTRY_COORDS.get(cn)
            if not coords:
                continue
            routes.append({
                "mineral": mn,
                "country": cn,
                "lat": coords[0],
                "lng": coords[1],
                "amount": round(v),
                "share": round(v / max_v * 100, 1),
            })

    return jsonify({"mineral": mineral or "전체", "unit": unit, "routes": routes})

@app.route("/api/summary")
def api_summary():
    c = fetch_customs(); n = fetch_news()
    return jsonify({"updated": datetime.now().isoformat(),
        "by_mineral": by_mineral(c)[:10], "by_country": by_country(c),
        "news_count": len(n), "latest_news": n[:5], "usgs": USGS_DATA})

@app.route("/subscribe", methods=["POST"])
def subscribe():
    email = ((request.get_json(silent=True) or {}).get("email") or "").strip().lower()
    if not valid_email(email): return jsonify(ok=False, message="올바른 이메일 형식이 아닙니다.")
    if not add_sub(email):
        return jsonify(ok=False, message="이미 구독 중인 이메일입니다.")
    return jsonify(ok=True, message=f"구독 완료! 현재 {len(load_subs())}명이 구독 중입니다.")

@app.route("/send_now", methods=["POST"])
def send_now():
    email = ((request.get_json(silent=True) or {}).get("email") or "").strip().lower()
    if not valid_email(email): return jsonify(ok=False, message="올바른 이메일 형식이 아닙니다.")
    if not SMTP_USER: return jsonify(ok=False, message="메일 발송 설정이 필요합니다.")
    subj = f"[핵심광물] {datetime.now().strftime('%m/%d')} 동향 리포트"
    ok, info = send_mail(email, subj, build_newsletter(email))
    return jsonify(ok=ok, message="리포트를 발송했습니다!" if ok else f"발송 실패: {info}")

@app.route("/send_all", methods=["POST"])
def send_all():
    subs = load_subs()
    if not subs: return jsonify(ok=False, message="구독자가 없습니다.")
    subj = f"[핵심광물] {datetime.now().strftime('%m/%d')} 동향 리포트"
    sent = failed = 0
    for e in subs:
        ok, _ = send_mail(e, subj, build_newsletter(e)); sent += ok; failed += (not ok)
    return jsonify(ok=True, message=f"발송 완료 — 성공 {sent}명 / 실패 {failed}명")

@app.route("/unsubscribe")
def unsubscribe():
    email = (request.args.get("email") or "").strip().lower()
    token = request.args.get("t") or ""
    page  = "<div style='font-family:sans-serif;max-width:480px;margin:60px auto;text-align:center;color:#333'>{}</div>"
    if not email or not hmac.compare_digest(token, unsub_token(email)):
        return Response(page.format("<h2>잘못된 수신거부 링크입니다.</h2>"), mimetype="text/html")
    remove_sub(email)
    return Response(page.format(f"<h2>수신거부 완료</h2><p>{email} 님은 더 이상 리포트를 받지 않습니다.</p>"), mimetype="text/html")

@app.route("/cron/daily")
def cron_daily():
    # 외부 크론(cron-job.org 등)이 매일 호출. CRON_TOKEN 으로 보호.
    if not CRON_TOKEN or request.args.get("token") != CRON_TOKEN:
        return jsonify(ok=False, message="forbidden"), 403
    subs = load_subs()
    subj = f"[핵심광물] {datetime.now().strftime('%m/%d')} 동향 리포트"
    sent = failed = 0
    for e in subs:
        ok, _ = send_mail(e, subj, build_newsletter(e)); sent += ok; failed += (not ok)
    return jsonify(ok=True, sent=sent, failed=failed, total=len(subs))

@app.route("/conference")
def conference():
    if not _conf_authed():
        return redirect("/conference/login")
    return Response(render_conference(), mimetype="text/html")

@app.route("/conference/login", methods=["GET", "POST"])
def conference_login():
    if not CONFERENCE_PASSWORD:          # 게이트 비활성(로컬)
        return redirect("/conference")
    err = ""
    if request.method == "POST":
        if request.form.get("password", "") == CONFERENCE_PASSWORD:
            session["conf_ok"] = True
            return redirect("/conference")
        err = "비밀번호가 올바르지 않습니다."
    return Response(render_login(err), mimetype="text/html")

@app.route("/api/conference/chat", methods=["POST"])
def conference_chat():
    """턴제: 지정된 전문가 1명만 발언한다. 다음 발언자/내 발언은 프론트가 제어."""
    if not _conf_authed():
        return jsonify(ok=False, message="인증이 필요합니다. 다시 로그인하세요."), 401
    data    = request.get_json(silent=True) or {}
    speaker = data.get("speaker")
    history = data.get("history", [])
    if not speaker or speaker not in MINERAL_EXPERTS:
        return jsonify(ok=False, message="발언할 전문가가 지정되지 않았습니다."), 400
    expert = MINERAL_EXPERTS[speaker]

    # 회의 주제 = 가장 처음의 사용자(진행자) 발언
    topic = ""
    for h in history:
        if h.get("role") == "user":
            topic = h.get("content", "")
            break

    def sse(obj):
        return f"data: {json.dumps(obj, ensure_ascii=False)}\n\n"

    def transcript_text(turns):
        lines = []
        for h in turns:
            if h.get("role") == "user":
                lines.append(f"[진행자] {h.get('content','')}")
            else:
                lines.append(f"[{h.get('name','전문가')}] {h.get('content','')}")
        return "\n".join(lines)

    def generate():
        yield sse({'speaker_start': speaker, 'name': expert['name'],
                   'avatar': expert['avatar'], 'color': expert['color']})

        api_key = expert.get("api_key") or OPENAI_API_KEY
        if not api_key:
            yield sse({'text': '⚠️ OpenAI API 키가 설정되지 않았습니다. dashboard_app.py의 OPENAI_API_KEY(또는 전문가별 api_key)를 채워주세요.'})
            yield sse({'speaker_end': speaker})
            yield "data: [DONE]\n\n"
            return

        # 이 전문가가 이미 한 발언들 / 직전 발언자 여부
        own_prior = [h.get("content", "") for h in history
                     if h.get("role") == "assistant" and h.get("name") == expert["name"]]
        last_turn = history[-1] if history else None
        is_consecutive = bool(last_turn and last_turn.get("role") == "assistant"
                              and last_turn.get("name") == expert["name"])

        repeat_guard = ""
        if own_prior:
            said = " // ".join(s.strip() for s in own_prior[-3:] if s.strip())
            repeat_guard = (
                f"\n\n[중요·반복 금지] 당신({expert['name']})은 이 회의에서 이미 발언했습니다. "
                f"당신이 앞서 한 말 → \"{said}\". "
                "위 내용을 절대 반복하거나 바꿔 말하지 마세요. 대신 ① 이전 발언을 더 깊이 보충하거나 "
                "② 아직 다루지 않은 새로운 논점을 제시하세요. 구체적 수치·사례·실행 단계·반론 등 "
                "직전에 없던 정보를 반드시 더해 논의를 한 발짝 진전시키세요."
            )
            if is_consecutive:
                repeat_guard += (
                    " 지금은 방금 당신의 발언에 곧바로 이어지는 추가 발언입니다. "
                    "\"앞서 말씀드린 데 더해—\" 같은 식으로 자연스럽게 이어, 한 단계 더 들어가세요."
                )

        sys_prompt = expert["system"] + (
            "\n\n[회의 형식] 이것은 여러 전문가와 진행자가 함께하는 실시간 회의입니다. "
            "아래 회의록을 읽고, 다른 전문가나 진행자의 발언을 직접 인용하며 동의하거나 반박한 뒤 "
            "자신의 핵심 의견을 200자 내외로 말하세요. 이미 나온 말을 반복하지 말고 논의를 진전시키세요. "
            "발언 앞에 자신의 이름이나 '[이름]' 같은 라벨을 붙이지 말고, 바로 본문부터 말하세요."
        ) + repeat_guard
        convo = transcript_text(history) or f"회의 주제: {topic}"
        user_prompt = (
            f"[회의 주제]\n{topic}\n\n"
            f"[지금까지의 회의록]\n{convo}\n\n"
            f"이제 {expert['name']}으로서 발언하세요."
            + ("  (반드시 앞서 당신이 한 말과 다른, 새로운 내용을 더하세요.)" if own_prior else "")
        )

        try:
            stream = OpenAI(api_key=api_key).chat.completions.create(
                model=expert.get("model", DEFAULT_OPENAI_MODEL),
                max_completion_tokens=400,
                messages=[
                    {"role": "system", "content": sys_prompt},
                    {"role": "user",   "content": user_prompt},
                ],
                stream=True,
            )
            for chunk in stream:
                if not chunk.choices:
                    continue
                delta = chunk.choices[0].delta.content
                if delta:
                    yield sse({'text': delta})
        except Exception as e:
            yield sse({'error': str(e), 'mineral': speaker})

        yield sse({'speaker_end': speaker})
        yield "data: [DONE]\n\n"

    return Response(stream_with_context(generate()), mimetype="text/event-stream",
                    headers={"Cache-Control": "no-cache", "X-Accel-Buffering": "no"})


# ═══════════════════════════════════════════════════════════════
#  ④ AI 전문가 회의실 페이지
# ═══════════════════════════════════════════════════════════════
def render_login(err=""):
    PAGE = r"""<!DOCTYPE html>
<html class="dark" lang="ko">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>로그인 — AI 전문가 회의실</title>
<script src="https://cdn.tailwindcss.com"></script>
<link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;600;700;900&family=Noto+Sans+KR:wght@400;700;900&display=swap" rel="stylesheet">
<style>body{background:#131315;color:#e4e2e4;font-family:'Inter','Noto Sans KR',sans-serif;}</style>
</head>
<body class="min-h-screen flex items-center justify-center p-6">
  <form method="POST" action="/conference/login" class="w-full max-w-sm bg-[#1f1f21] border border-[#2a2c2f] rounded-2xl p-8">
    <div class="flex items-center gap-3 mb-6">
      <div class="w-10 h-10 rounded-lg flex items-center justify-center font-black text-xl" style="background:#e9c349;color:#0f172a">K</div>
      <div>
        <div class="font-black text-[#e9c349] tracking-wider">K-MINERAL AI</div>
        <div class="text-[10px] uppercase tracking-widest text-[#909097]">AI 전문가 회의실</div>
      </div>
    </div>
    <p class="text-sm text-[#c6c6cd] mb-5">이 회의실은 비밀번호로 보호되어 있습니다.</p>
    <input type="password" name="password" autofocus placeholder="비밀번호" class="w-full bg-[#0e0e10] border border-[#45464d] rounded-lg px-4 py-3 text-sm outline-none focus:border-[#e9c349] mb-3">
    <div class="text-[#ffb4ab] text-xs mb-3" style="min-height:16px">__ERR__</div>
    <button type="submit" class="w-full font-bold py-3 rounded-lg" style="background:#e9c349;color:#241a00">입장하기</button>
    <a href="/" class="block text-center text-xs text-[#909097] mt-4 hover:text-[#e9c349]">← 대시보드로</a>
  </form>
</body>
</html>"""
    return PAGE.replace("__ERR__", err)


def render_conference():
    experts_json = json.dumps(
        {k: {kk: vv for kk, vv in v.items() if kk not in ("system", "api_key")} for k, v in MINERAL_EXPERTS.items()},
        ensure_ascii=False
    )
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    PAGE = r"""<!DOCTYPE html>
<html class="dark" lang="ko">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>AI 전문가 회의실 — K-Mineral AI Insight</title>
<script src="https://cdn.tailwindcss.com?plugins=forms,container-queries"></script>
<link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800;900&family=JetBrains+Mono:wght@400;500&family=Noto+Sans+KR:wght@400;500;700;900&display=swap" rel="stylesheet">
<link href="https://fonts.googleapis.com/css2?family=Material+Symbols+Outlined:wght,FILL@100..700,0..1&display=swap" rel="stylesheet">
<script>
tailwind.config = {
  darkMode: "class",
  theme: { extend: {
    colors: {
      "surface-variant":"#353436","outline-variant":"#45464d","surface-container-low":"#1b1b1d",
      "surface-container-lowest":"#0e0e10","on-surface":"#e4e2e4","surface-container-high":"#2a2a2b",
      "background":"#131315","surface-container-highest":"#353436","primary":"#bec6e0",
      "surface-container":"#1f1f21","on-secondary":"#3c2f00","outline":"#909097",
      "primary-container":"#0f172a","on-primary-container":"#798098","on-surface-variant":"#c6c6cd",
      "secondary":"#e9c349","on-secondary-fixed":"#241a00","surface":"#131315","error":"#ffb4ab",
      "tertiary":"#b9c8de"
    },
    fontFamily: { "data-tabular":["JetBrains Mono","monospace"], "sans":["Inter","Noto Sans KR","sans-serif"] },
    fontSize: {
      "headline-lg":["32px",{"lineHeight":"1.3","fontWeight":"700"}],
      "headline-md":["22px",{"lineHeight":"1.4","fontWeight":"700"}],
      "label-md":["14px",{"lineHeight":"1.0","fontWeight":"500"}]
    }
  }}
}
</script>
<style>
  body{background:#131315;color:#e4e2e4;font-family:'Inter','Noto Sans KR',sans-serif;}
  .material-symbols-outlined{font-variation-settings:'FILL' 0,'wght' 400,'GRAD' 0,'opsz' 24;}
  .glass-panel{background:rgba(30,41,59,.55);backdrop-filter:blur(12px);border:1px solid #334155;}
  .custom-scrollbar::-webkit-scrollbar{width:6px;}
  .custom-scrollbar::-webkit-scrollbar-track{background:transparent;}
  .custom-scrollbar::-webkit-scrollbar-thumb{background:#45464d;border-radius:10px;}
  .expert-card.selected{border-color:#e9c349 !important;box-shadow:0 0 0 1px #e9c349,0 0 16px rgba(233,195,73,.18);}
  .expert-card.selected .ec-check{opacity:1 !important;}
  .tc-suggested{box-shadow:0 0 0 1px #e9c349,0 0 10px rgba(233,195,73,.35);}
  .lobby-screen,#roomScreen{display:none;}
</style>
</head>
<body class="flex min-h-screen bg-background">

<!-- Sidebar -->
<aside class="fixed left-0 top-0 h-screen w-64 bg-surface-container-low border-r border-outline-variant/20 flex flex-col p-4 z-50">
  <div class="flex items-center gap-3 px-2 mb-8">
    <div class="w-10 h-10 bg-secondary rounded-lg flex items-center justify-center text-primary-container font-black text-xl">K</div>
    <div>
      <h1 class="text-base font-black text-secondary uppercase tracking-wider leading-none">K-Mineral AI</h1>
      <p class="text-on-surface-variant text-[10px] uppercase tracking-widest mt-1">Strategic Insight</p>
    </div>
  </div>
  <nav class="flex-1 space-y-1 font-data-tabular">
    <a href="/" class="flex items-center gap-3 px-4 py-3 text-on-surface-variant hover:bg-surface-variant/50 rounded-lg transition"><span class="material-symbols-outlined">dashboard</span><span class="text-sm">Dashboard</span></a>
    <a href="/#map" class="flex items-center gap-3 px-4 py-3 text-on-surface-variant hover:bg-surface-variant/50 rounded-lg transition"><span class="material-symbols-outlined">public</span><span class="text-sm">Supply Chain Map</span></a>
    <a href="/#news" class="flex items-center gap-3 px-4 py-3 text-on-surface-variant hover:bg-surface-variant/50 rounded-lg transition"><span class="material-symbols-outlined">rss_feed</span><span class="text-sm">News Feed</span></a>
    <a href="/conference" class="flex items-center gap-3 px-4 py-3 bg-primary-container text-on-primary-container rounded-lg font-bold"><span class="material-symbols-outlined">forum</span><span class="text-sm">AI Meeting Room</span></a>
    <a href="/#komir" class="flex items-center gap-3 px-4 py-3 text-on-surface-variant hover:bg-surface-variant/50 rounded-lg transition"><span class="material-symbols-outlined">database</span><span class="text-sm">Raw Data</span></a>
  </nav>
  <div class="mt-auto pt-4 border-t border-outline-variant/20 space-y-2 px-2">
    <div class="flex items-center gap-2 text-on-surface-variant text-[11px]"><span class="material-symbols-outlined text-sm text-secondary" style="font-variation-settings:'FILL' 1">schedule</span>System Status: Live</div>
    <div class="flex items-center gap-2 text-on-surface-variant text-[11px]"><span class="material-symbols-outlined text-sm text-secondary" style="font-variation-settings:'FILL' 1">check_circle</span>API: Stable</div>
  </div>
</aside>

<!-- Main -->
<main class="ml-64 flex-1 h-screen flex flex-col bg-background overflow-hidden">
  <header class="h-16 shrink-0 flex items-center justify-between px-8 border-b border-outline-variant/30 bg-surface/70 backdrop-blur-xl">
    <div class="flex items-center gap-3">
      <span class="material-symbols-outlined text-secondary">forum</span>
      <h2 class="text-headline-md text-secondary">AI 전문가 회의실</h2>
    </div>
    <span id="confClock" class="font-data-tabular text-xs text-on-surface-variant">__NOW__ KST ● LIVE</span>
  </header>

  <div class="flex-1 min-h-0 relative">

    <!-- STEP 1 -->
    <div id="step1Screen" class="lobby-screen absolute inset-0 flex-col items-center overflow-y-auto p-8 custom-scrollbar" style="display:flex">
      <div class="w-full max-w-3xl mx-auto">
        <h1 class="text-headline-lg text-on-surface mb-2">핵심광물 AI 전문가 회의실</h1>
        <p class="text-on-surface-variant text-sm mb-8"><span class="text-secondary font-bold">STEP 1.</span> 회의에 데려갈 전문가를 선택하세요. 광물·경제·정치 분야 전문가를 자유롭게 조합할 수 있습니다.</p>
        <div id="expertGrid" class="space-y-6"></div>
        <div class="flex items-center justify-between mt-8">
          <span id="selCount" class="font-data-tabular text-xs text-on-surface-variant">0명 선택됨</span>
          <button id="toStep2Btn" onclick="goToStep2()" class="bg-secondary text-on-secondary-fixed font-bold py-3 px-6 rounded-lg hover:opacity-90 transition disabled:opacity-40 disabled:cursor-not-allowed">다음 → 안건 설정</button>
        </div>
      </div>
    </div>

    <!-- STEP 2 -->
    <div id="step2Screen" class="lobby-screen absolute inset-0 flex-col items-center overflow-y-auto p-8 custom-scrollbar">
      <div class="w-full max-w-3xl mx-auto">
        <h1 class="text-headline-lg text-on-surface mb-2">회의 안건 설정</h1>
        <p class="text-on-surface-variant text-sm mb-8"><span class="text-secondary font-bold">STEP 2.</span> 선택한 전문가들에게 던질 회의 안건(질문)을 입력하세요.</p>
        <div class="text-[10px] font-bold text-outline uppercase tracking-widest mb-3 font-data-tabular">회의에 참여할 전문가</div>
        <div id="teamSummary" class="flex flex-wrap gap-2 mb-6 min-h-[28px]"></div>
        <div class="text-[10px] font-bold text-outline uppercase tracking-widest mb-3 font-data-tabular">질문 입력</div>
        <textarea id="questionInput" rows="3" class="w-full bg-surface-container-lowest border border-outline-variant/30 rounded-lg p-4 text-sm text-on-surface focus:ring-1 focus:ring-secondary outline-none resize-none mb-3" placeholder="예: 중국의 희토류 수출 규제가 한국 배터리 산업에 미치는 영향은?"></textarea>
        <p class="text-[11px] text-on-surface-variant mb-8">회의가 시작되면 한 명씩 발언합니다. 발언이 끝날 때마다 <b class="text-secondary">다음 발언자</b>를 직접 고르거나, 직접 발언할 수 있어요.</p>
        <div class="flex items-center justify-between">
          <button onclick="backToStep1()" class="text-sm text-on-surface-variant border border-outline-variant/40 rounded-lg px-4 py-2.5 hover:border-secondary hover:text-secondary transition">← 전문가 다시 선택</button>
          <button id="startBtn" onclick="startSession()" class="bg-secondary text-on-secondary-fixed font-bold py-3 px-6 rounded-lg hover:opacity-90 transition">회의 시작 →</button>
        </div>
      </div>
    </div>

    <!-- ROOM -->
    <div id="roomScreen" class="absolute inset-0 flex-col">
      <div class="px-6 py-3 border-b border-outline-variant/20 flex items-center gap-3 bg-surface-container-low/40 shrink-0">
        <span class="text-[10px] uppercase tracking-widest text-outline font-data-tabular shrink-0">참여 전문가</span>
        <div id="activeExperts" class="flex flex-wrap gap-1.5"></div>
        <button onclick="backToLobby()" class="ml-auto text-xs text-on-surface-variant border border-outline-variant/30 rounded-lg px-3 py-1.5 hover:border-secondary hover:text-secondary transition shrink-0">← 다시 시작</button>
      </div>
      <div id="chatArea" class="flex-1 overflow-y-auto p-8 space-y-6 custom-scrollbar"></div>
      <div id="typingIndicator" class="px-8 pb-1 text-xs text-secondary font-data-tabular" style="display:none">● 전문가가 답변 중...</div>
      <div id="turnControls" class="px-6 py-3 border-t border-outline-variant/20 bg-surface-container-low/40 flex items-center gap-3 flex-wrap shrink-0" style="display:none">
        <span class="text-[10px] uppercase tracking-widest text-outline font-data-tabular shrink-0">다음 발언자 ▶</span>
        <div id="tcExperts" class="flex flex-wrap gap-2"></div>
      </div>
      <div class="px-6 py-4 border-t border-outline-variant/20 bg-surface-container-low/60 flex gap-3 shrink-0">
        <input id="chatInput" onkeydown="if(event.key==='Enter'&&!event.isComposing)sendMessage()" placeholder="진행자로서 직접 발언 (전송 후 다음 발언자 선택)..." class="flex-1 bg-surface-container-lowest border border-outline-variant/40 rounded-lg px-4 py-2.5 text-sm text-on-surface focus:ring-1 focus:ring-secondary outline-none">
        <button onclick="sendMessage()" class="bg-secondary text-on-secondary-fixed font-bold px-5 rounded-lg flex items-center gap-1.5 hover:opacity-90 transition"><span class="material-symbols-outlined text-sm">send</span>내 발언</button>
      </div>
    </div>

  </div>
</main>

<script>
const EXPERTS = __EXPERTS_JSON__;
let selectedExperts = [];
let chatHistory = [];
let turnOrder = [];
let turnIdx = 0;
let busy = false;

// 전문가 카드 생성 (분야별 그룹)
const grid = document.getElementById('expertGrid');
const CAT_ORDER = ['광물','경제','정치'];
const byCat = {};
Object.entries(EXPERTS).forEach(([key, ex]) => {
  const cat = ex.category || '기타';
  (byCat[cat] = byCat[cat] || []).push([key, ex]);
});
const cats = [...CAT_ORDER.filter(c => byCat[c]), ...Object.keys(byCat).filter(c => !CAT_ORDER.includes(c))];
cats.forEach(cat => {
  const group = document.createElement('div');
  const head = document.createElement('div');
  head.className = 'text-[10px] font-bold text-secondary uppercase tracking-widest mb-3 font-data-tabular';
  head.textContent = '◆ ' + cat;
  group.appendChild(head);
  const cg = document.createElement('div');
  cg.className = 'grid grid-cols-2 gap-3';
  byCat[cat].forEach(([key, ex]) => {
    const card = document.createElement('div');
    card.className = 'expert-card bg-surface-container border border-outline-variant/30 rounded-xl p-3 flex items-center gap-3 cursor-pointer transition-all hover:border-secondary/50';
    card.dataset.key = key;
    card.innerHTML = '<span class="text-2xl">'+ex.avatar+'</span>'
      + '<div class="flex-1 min-w-0"><div class="text-sm font-bold truncate" style="color:'+ex.color+'">'+ex.name+'</div>'
      + '<div class="text-[11px] text-on-surface-variant truncate">'+ex.title+'</div></div>'
      + '<span class="ec-check material-symbols-outlined text-secondary text-lg opacity-0 transition-opacity">check_circle</span>';
    card.onclick = () => { card.classList.toggle('selected'); updateSelection(); };
    cg.appendChild(card);
  });
  group.appendChild(cg);
  grid.appendChild(group);
});

function updateSelection() {
  selectedExperts = [...document.querySelectorAll('.expert-card.selected')].map(c => c.dataset.key);
  document.getElementById('toStep2Btn').disabled = selectedExperts.length === 0;
  document.getElementById('selCount').textContent = selectedExperts.length + '명 선택됨';
}
updateSelection();

function showScreen(id) {
  ['step1Screen','step2Screen','roomScreen'].forEach(s => document.getElementById(s).style.display = 'none');
  document.getElementById(id).style.display = 'flex';
}

function goToStep2() {
  if (selectedExperts.length === 0) return;
  document.getElementById('teamSummary').innerHTML = selectedExperts.map(k => {
    const ex = EXPERTS[k];
    return '<span class="inline-flex items-center gap-1.5 text-xs font-bold px-3 py-1.5 rounded-full" style="background:'+ex.color+'22;color:'+ex.color+';border:1px solid '+ex.color+'55">'+ex.avatar+' '+ex.name+'</span>';
  }).join('');
  showScreen('step2Screen');
}

function backToStep1() { showScreen('step1Screen'); }

function startSession() {
  const q = document.getElementById('questionInput').value.trim();
  if (!q || selectedExperts.length === 0) return;
  chatHistory = [];
  turnOrder = selectedExperts.slice();
  turnIdx = 0;
  busy = false;
  showScreen('roomScreen');
  document.getElementById('activeExperts').innerHTML = selectedExperts.map(k => {
    const ex = EXPERTS[k];
    return '<span class="inline-flex items-center gap-1 text-[11px] font-bold px-2.5 py-1 rounded-full" style="background:'+ex.color+'22;color:'+ex.color+';border:1px solid '+ex.color+'44">'+ex.avatar+' '+ex.name+'</span>';
  }).join('');
  document.getElementById('chatArea').innerHTML = '';
  appendUserMsg(q);
  speakExpert(turnOrder[0]);
}

function backToLobby() {
  showScreen('step1Screen');
  document.getElementById('chatArea').innerHTML = '';
  document.getElementById('turnControls').style.display = 'none';
  chatHistory = [];
  busy = false;
}

function appendUserMsg(text) {
  const chatArea = document.getElementById('chatArea');
  const div = document.createElement('div');
  div.className = 'flex justify-end';
  div.innerHTML = '<div class="max-w-[75%]"><div class="flex items-center justify-end gap-2 mb-1"><span class="text-[11px] font-bold text-secondary">🎙️ 진행자 (나)</span></div><div class="msg-bubble bg-secondary/10 border border-secondary/30 rounded-xl rounded-tr-none px-4 py-3 text-sm text-on-surface leading-relaxed"></div></div>';
  div.querySelector('.msg-bubble').textContent = text;
  chatArea.appendChild(div);
  chatArea.scrollTop = chatArea.scrollHeight;
  chatHistory.push({role:'user', content:text});
}

function renderTurnControls() {
  const tc = document.getElementById('turnControls');
  const box = document.getElementById('tcExperts');
  const suggested = turnOrder[turnIdx % turnOrder.length];
  box.innerHTML = turnOrder.map(k => {
    const ex = EXPERTS[k];
    const sug = (k === suggested) ? ' tc-suggested' : '';
    return '<button onclick="speakExpert(\''+k+'\')" class="tc-btn inline-flex items-center gap-1.5 text-xs font-bold px-3 py-1.5 rounded-full border transition-all hover:opacity-90'+sug+'" style="color:'+ex.color+';border-color:'+ex.color+'55">'+ex.avatar+' '+ex.name+'</button>';
  }).join('');
  tc.style.display = 'flex';
}

let currentBubble = null;

function speakExpert(key) {
  if (busy || !EXPERTS[key]) return;
  busy = true;
  document.getElementById('turnControls').style.display = 'none';
  const ti = document.getElementById('typingIndicator');
  ti.style.display = 'block';
  fetch('/api/conference/chat', {
    method:'POST',
    headers:{'Content-Type':'application/json'},
    body: JSON.stringify({speaker: key, history: chatHistory})
  }).then(r => {
    const reader = r.body.getReader();
    const decoder = new TextDecoder();
    let buf = '';
    function finish() {
      ti.style.display = 'none';
      busy = false;
      turnIdx = (turnOrder.indexOf(key) + 1) % turnOrder.length;
      renderTurnControls();
    }
    function read() {
      reader.read().then(({done, value}) => {
        if (done) { finish(); return; }
        buf += decoder.decode(value, {stream:true});
        const lines = buf.split('\n');
        buf = lines.pop();
        lines.forEach(line => {
          if (!line.startsWith('data:')) return;
          const raw = line.slice(5).trim();
          if (raw === '[DONE]') { finish(); return; }
          try {
            const d = JSON.parse(raw);
            const chatArea = document.getElementById('chatArea');
            if (d.speaker_start) {
              const ex = EXPERTS[d.speaker_start] || {};
              const div = document.createElement('div');
              div.className = 'flex gap-3 max-w-[85%]';
              div.innerHTML = '<div class="shrink-0 w-9 h-9 rounded-lg border flex items-center justify-center text-lg" style="border-color:'+((ex.color||'#e9c349')+'66')+'">'+(ex.avatar||'')+'</div>'
                + '<div class="flex-1 min-w-0"><div class="flex items-baseline gap-2 mb-1"><span class="text-sm font-bold" style="color:'+(ex.color||'#e9c349')+'">'+(ex.name||d.speaker_start)+'</span></div>'
                + '<div class="msg-bubble glass-panel rounded-xl rounded-tl-none px-4 py-3 text-sm text-on-surface leading-relaxed border-l-2" style="border-left-color:'+(ex.color||'#e9c349')+'"></div></div>';
              chatArea.appendChild(div);
              currentBubble = div.querySelector('.msg-bubble');
              chatArea.scrollTop = chatArea.scrollHeight;
            } else if (d.text && currentBubble) {
              currentBubble.textContent += d.text;
              chatArea.scrollTop = chatArea.scrollHeight;
            } else if (d.speaker_end) {
              if (currentBubble) chatHistory.push({role:'assistant', name:EXPERTS[d.speaker_end]?.name||d.speaker_end, content:currentBubble.textContent});
              currentBubble = null;
            }
          } catch(e) {}
        });
        read();
      });
    }
    read();
  }).catch(e => { ti.style.display = 'none'; busy = false; renderTurnControls(); console.error(e); });
}

function sendMessage() {
  if (busy) return;
  const input = document.getElementById('chatInput');
  const msg = input.value.trim();
  if (!msg) return;
  input.value = '';
  appendUserMsg(msg);
  renderTurnControls();
}

// 실시간 시계
setInterval(() => {
  const el = document.getElementById('confClock');
  if (!el) return;
  const d = new Date();
  const p = n => String(n).padStart(2,'0');
  el.textContent = d.getFullYear()+'-'+p(d.getMonth()+1)+'-'+p(d.getDate())+' '+p(d.getHours())+':'+p(d.getMinutes())+':'+p(d.getSeconds())+' KST ● LIVE';
}, 1000);
</script>
</body>
</html>"""
    return PAGE.replace("__EXPERTS_JSON__", experts_json).replace("__NOW__", now)


if __name__ == "__main__":
    import sys
    sys.stdout.reconfigure(encoding='utf-8', errors='replace')
    print("핵심광물 대시보드 시작")
    print("브라우저 접속: http://127.0.0.1:8080")
    print("종료: Ctrl + C")
    app.run(host="0.0.0.0", port=8080, debug=False)
