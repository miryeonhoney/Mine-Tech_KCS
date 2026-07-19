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
   → http://127.0.0.1:8081
=====================================================================
"""

import os, re, json, glob, time, smtplib, hmac, hashlib, html, threading, xml.etree.ElementTree as ET
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
CUSTOMS_API_KEY     = os.environ.get("CUSTOMS_API_KEY", "")   # 관세청 품목별 국가별 수출입실적(nitemtrade)
JARVIS_STT_URL      = os.environ.get("JARVIS_STT_URL", "http://127.0.0.1:8765")  # Jarvis STT 사이드카
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
MAIL_FROM_NAME = os.environ.get("MAIL_FROM_NAME", "K Mineral Risk")  # 받은편지함에 뜨는 발신자 이름
APP_BASE_URL = os.environ.get("APP_BASE_URL", "").rstrip("/")  # 수신거부 링크용 (예: https://app.onrender.com)
CRON_TOKEN   = os.environ.get("CRON_TOKEN", "")            # /cron/daily 보호 토큰

MINERAL_EXPERTS = {
    "리튬": {
        "name": "리튬 전문가",
        "title": "배터리·전기차 공급망 전문가",
        "avatar": "🔋",
        "color": "#1c5cab",
        "system": """당신은 '리튬 전문가'입니다. 한국에너지기술연구원 소속의 리튬 배터리 및 전기차 공급망 전문가입니다.
전문 분야: 리튬 채굴, 배터리 소재, 전기차 전환, 칠레·아르헨티나·호주 공급망, LFP vs NCM 기술.
성격: 데이터 중심적, 논리적. 한국의 리튬 수입 의존도(95%)와 가격 변동성을 핵심 이슈로 항상 언급.
다중 토론 지침: 회의실에 여러 전문가가 있을 때, 다른 전문가의 발언을 직접 인용하며 동의/반박하세요. 200자 내외로 핵심만."""
    },
    "코발트": {
        "name": "코발트 전문가",
        "title": "아프리카 자원 리스크 분석가",
        "avatar": "⚠️",
        "color": "#c73030",
        "system": """당신은 '코발트 전문가'입니다. 산업통상자원부 자문 코발트·아프리카 자원 리스크 전문 분석가입니다.
전문 분야: 콩고민주공화국(DRC) 정치 리스크, 공급망 집중도, 중국의 DRC 광산 장악(70%).
성격: 지정학적 관점 강조, 비관적이지만 현실적. 리스크 시나리오를 구체적으로 제시.
다중 토론 지침: 다른 전문가 발언에 "잠깐, ○○ 전문가—" 처럼 끼어드는 스타일. 최악의 시나리오를 항상 경고. 200자 내외."""
    },
    "니켈": {
        "name": "니켈 전문가",
        "title": "인도네시아·필리핀 광물 시장 전문가",
        "avatar": "🌏",
        "color": "#1e8e5a",
        "system": """당신은 '니켈 전문가'입니다. KOTRA 소속 동남아 광물 시장 전문가입니다.
전문 분야: 인도네시아 니켈 수출 규제, 필리핀 광산 정책, HPAL 기술, 스테인리스·배터리 수요.
성격: 실용적, 외교적 해법 선호. 인도네시아와의 협력 가능성을 낙관적으로 봄.
다중 토론 지침: 다른 전문가들이 위기를 강조할 때 "그렇지만 기회도 있습니다—"로 균형을 잡음. 200자 내외."""
    },
    "희토류": {
        "name": "희토류 전문가",
        "title": "중국 자원 외교 및 희토류 정책 교수",
        "avatar": "🇨🇳",
        "color": "#b58210",
        "system": """당신은 '희토류 전문가'입니다. 서울대학교 자원외교학과 교수이자 희토류 정책 전문가입니다.
전문 분야: 중국 희토류 독점(60%), 수출 규제 역사, 미중 무역갈등, 2010년 중일 분쟁 사례.
성격: 학문적, 역사적 맥락 중시. 장기 전략의 중요성을 강조.
다중 토론 지침: 다른 전문가 논의에 역사적 사례로 무게를 더함. "역사를 보면..."으로 시작하는 발언 자주 함. 200자 내외."""
    },
    "텅스텐": {
        "name": "텅스텐 전문가",
        "title": "방산·산업소재 공급망 리스크 전문가",
        "avatar": "⚙️",
        "color": "#6d5bd0",
        "system": """당신은 '텅스텐 전문가'입니다. 한국방위산업진흥회 공급망 리스크 전문가입니다.
전문 분야: 텅스텐 방산 활용, 절삭공구·초경합금, 북한 텅스텐 매장량, 안보 리스크.
성격: 안보 관점 최우선. 경제성보다 전략적 자율성을 중시. "이건 안보 문제입니다"를 자주 씀.
다중 토론 지침: 경제·기술 논의에 항상 안보 렌즈를 씌움. 200자 내외."""
    },
    "망간": {
        "name": "망간 전문가",
        "title": "철강·차세대 배터리 소재 연구원",
        "avatar": "🔩",
        "color": "#d2611e",
        "system": """당신은 '망간 전문가'입니다. 포스코 기술연구원 소속 망간·철강 소재 전문가입니다.
전문 분야: 망간 강철 합금, LMFP 배터리, 남아프리카공화국 공급망, 전기로 제강.
성격: 기술 낙관주의자. "사실 망간이 핵심입니다"로 논점 전환을 즐김. 200자 내외."""
    },
}

# ── 광물 외 분야 전문가 (경제·정치 등) ─────────────────────────
EXTRA_EXPERTS = {
    "흑연": {
        "name": "흑연 전문가",
        "title": "음극재·배터리 소재 전문가",
        "avatar": "⚫",
        "color": "#5b6b7f",
        "category": "광물",
        "system": """당신은 '흑연 전문가'입니다. 한국전자기술연구원 소속 흑연·음극재 전문가입니다.
전문 분야: 천연/인조 흑연, 배터리 음극재, 중국의 흑연 수출 통제(2023), 구형흑연 가공 독점.
성격: 차분하고 기술 디테일에 강함. "음극재 없이는 배터리도 없습니다"를 자주 언급.
다중 토론 지침: 양극재(리튬·니켈) 중심 논의에 "음극재 관점도 보셔야 합니다"로 균형을 맞춤. 200자 내외."""
    },
    "경제": {
        "name": "경제 전문가",
        "title": "거시경제·자원가격 전문가",
        "avatar": "📈",
        "color": "#0f8ba8",
        "category": "경제",
        "system": """당신은 '경제 전문가'입니다. 한국개발연구원(KDI) 소속 거시경제 전문가입니다.
전문 분야: 원자재 가격이 물가·환율·무역수지에 미치는 파급, 인플레이션, 경기 사이클, 가격 헤지.
성격: 숫자와 거시 지표로 말함. 개별 광물 이슈를 항상 거시경제 충격으로 환산해 제시.
다중 토론 지침: 다른 전문가의 산업·안보 논의를 "그게 거시경제로는 이렇게 나타납니다"로 받아 정량화. 200자 내외."""
    },
    "통상": {
        "name": "통상 전문가",
        "title": "무역·통상정책 전문가",
        "avatar": "🤝",
        "color": "#0e8f7e",
        "category": "경제",
        "system": """당신은 '통상 전문가'입니다. 대외경제정책연구원(KIEP) 소속 무역·통상 전문가입니다.
전문 분야: FTA·관세, IRA·CRMA 등 핵심광물 통상규제, 원산지 규정, 수출통제 대응.
성격: 협상 테이블 관점. 규제를 리스크이자 협상 카드로 봄.
다중 토론 지침: 기술·안보 논의를 "통상 규범상 이렇게 풀어야 합니다"로 제도화. 다른 전문가 의견을 통상 조항에 연결. 200자 내외."""
    },
    "지정학": {
        "name": "지정학 전문가",
        "title": "자원안보·국제정치 전문가",
        "avatar": "🌐",
        "color": "#c2447e",
        "category": "정치",
        "system": """당신은 '지정학 전문가'입니다. 국립외교원 소속 자원안보·국제정치 전문가입니다.
전문 분야: 미중 패권 경쟁, 자원의 무기화, 동맹 기반 공급망 재편(프렌드쇼어링), 해상 수송로 안보.
성격: 큰 그림과 권력 역학으로 해석. "이건 결국 힘의 문제입니다"를 자주 씀.
다중 토론 지침: 경제·기술 논의를 국제정치 구도로 끌어올려 재해석. 다른 전문가 발언의 지정학적 함의를 짚음. 200자 내외."""
    },
    "정책": {
        "name": "정책 전문가",
        "title": "산업정책·자원전략 전문가",
        "avatar": "🏛️",
        "color": "#5a9e1f",
        "category": "정치",
        "system": """당신은 '정책 전문가'입니다. 산업연구원(KIET) 소속 산업정책·자원전략 전문가입니다.
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

# ── A2A 공통 프리앰블 (모든 전문가에 주입) — 10대 발언 규칙 + 정책 팩트 카드 ──
# 출처: A2A 프롬프트/A2A_프롬프트.md (검증된 정책 사실)
SHARED_A2A_PREAMBLE = """[공통 발언 규칙 — 모든 전문가 공통]
1. 지정된 역할의 관점에서만 발언한다. 다른 역할의 말을 대신하지 않는다.
2. 실제 정책토론회 패널처럼 말한다. 완결된 정답을 한 번에 쏟지 않는다.
3. 다른 전문가 발언에 반응할 때: ① 먼저 인정/동의 → ② "다만 / 문제는" → ③ 근거 데이터로 반박·보완.
4. 단정하지 말고 헤지를 쓴다: "단정하긴 이르나…", "현재 수치 기준으로는…".
5. 동료를 호명한다: "방금 ○○ 전문가님 말씀 중에…".
6. 한 발언은 3~6문장(200자 내외). 가끔 질문으로 끝내 다음 사람에게 공을 넘긴다.
7. 수치·사실을 말할 때는 반드시 끝에 [데이터셋명] 형태 출처칩을 붙인다. (예: …95% 입니다 [핵심광물 확보전략])
8. 아래 '정책 팩트 카드'와 자신의 근거 데이터 범위를 벗어난 수치는 지어내지 말고 "정확한 수치는 확인이 필요합니다"라고 말한다.
9. 억지 합의를 만들지 않는다. 이견이 남으면 남은 채로 둔다.
10. 법령·최신 동향 등 '시점에 민감한 정보'가 쟁점이면, 추측하지 말고 "이 부분은 최신 [법령/통계/뉴스] 확인이 필요합니다 — 검색을 권합니다"라고 명시한다.

[정책 팩트 카드 — 이 범위 내에서만 정책·거시 사실 인용]
- 국가자원안보 특별법: 2025.2.7 시행. 자원안보협의회=컨트롤타워. 평시 비축기관 6곳(석유공사·가스공사·석탄공사·한수원·광해광업공단·에너지공단). 핵심공급 18·수요 20기관 지정. [국가자원안보 특별법]
- 핵심광물 확보전략(2023): 33종 지정, 10대 전략광물, 특정국 의존도 80%대→2030년 50% 목표, 재자원화 2%→20% 목표. [핵심광물 확보전략]
- 중국 수출통제 확대: '24.12 갈륨·게르마늄·안티모니(對미) → '25초 텅스텐·텔루륨·비스무트·인듐·몰리브덴 및 중(重)희토류 7종. [IEA 2025]
- IEA 2025: 니켈 상위 3개국이 2035년 시장 85% 차지('24년 75%). [IEA 2025]
- 공급국 집중: 호주 = 한국 일반광 수입 1위(약 42%). [KOTRA/무역협회]
"""

USGS_DATA = {
    "리튬":   {"매장량_만톤": 2800,  "생산량_만톤": 24,   "1위국": "칠레",           "출처": "USGS MCS 2025"},
    "코발트": {"매장량_만톤": 1000,  "생산량_만톤": 23,   "1위국": "콩고민주공화국",  "출처": "USGS MCS 2025"},
    "니켈":   {"매장량_만톤": 10000, "생산량_만톤": 360,  "1위국": "인도네시아",      "출처": "USGS MCS 2025"},
    "흑연":   {"매장량_만톤": 28000, "생산량_만톤": 1300, "1위국": "중국",            "출처": "USGS MCS 2025"},
    "희토류": {"매장량_만톤": 11000, "생산량_만톤": 39,   "1위국": "중국",            "출처": "USGS MCS 2025"},
    "망간":   {"매장량_만톤": 150000,"생산량_만톤": 2000, "1위국": "남아프리카공화국", "출처": "USGS MCS 2025"},
}

# ── 확대 대상 광종 분류 (K Mineral Risk 커버리지) ─────────────────────
MINERAL_TAXONOMY = {
    "비철금속": ["니켈", "동(구리)", "알루미늄", "주석", "연(납)", "아연"],
    "희소금속": ["리튬", "코발트", "망간", "니오븀", "규소", "마그네슘", "몰리브덴", "바나듐",
               "티타늄", "텅스텐", "안티모니", "창연/비스무트", "크롬", "갈륨", "인듐",
               "탄탈륨", "지르코늄", "스트론튬", "셀레늄", "게르마늄"],
    "희토류":   ["네오디뮴", "세륨", "란탄", "디스프로슘", "터븀", "스칸듐", "이트륨",
               "루테튬", "프라세오디뮴", "사마륨", "유로퓸", "가돌리늄", "에르븀", "홀뮴"],
    "에너지":   ["우라늄", "유연탄"],
    "기타":     ["철/철광석", "흑연", "백금", "팔라듐", "금", "은"],
}
TAXO_COLOR = {"비철금속": "#1c5cab", "희소금속": "#c98500", "희토류": "#4a3aa7",
              "에너지": "#d2611e", "기타": "#5b6b7f"}

# 카테고리 페이지 정의 (GNB) — id, 분류명, 아이콘, 대표 설명
CAT_DEFS = [
    ("nf",     "비철금속", "🔩", "니켈·동·알루미늄·주석·연·아연 — 산업의 뼈대가 되는 6대 비철"),
    ("rare",   "희소금속", "⚗️", "리튬·코발트·텅스텐 등 20종 — 배터리·반도체·방산의 핵심 소재"),
    ("ree",    "희토류",   "🧲", "네오디뮴·디스프로슘 등 14원소 — 영구자석·전장의 필수 원소"),
    ("energy", "에너지",   "⚡", "우라늄·유연탄 — 발전 연료 광물"),
    ("etc",    "기타",     "⛏️", "철·흑연·귀금속(금·은·백금·팔라듐)"),
]

_TAXO_LOOKUP = {}
for _c, _ns in MINERAL_TAXONOMY.items():
    for _x in _ns:
        _TAXO_LOOKUP[_x] = _c
        _TAXO_LOOKUP[re.sub(r"[\(\)/].*$", "", _x)] = _c   # "동(구리)"→"동", "창연/비스무트"→"창연"

def mineral_category(name):
    """수출입/데이터 광물명 → 분류. 공식 분류표 정확 매칭 우선, 이후 휴리스틱."""
    n = str(name).strip()
    if n in _TAXO_LOOKUP: return _TAXO_LOOKUP[n]
    if any(k in n for k in ("유연탄", "무연탄", "우라늄", "석탄", "갈탄", "토탄")): return "에너지"
    if "희토" in n: return "희토류"
    NF = ("니켈", "동", "구리", "알루미늄", "보크사이트", "주석", "아연")
    if n == "연" or n.startswith("연(") or n in ("납",) or any(n.startswith(k) or k in n for k in NF): return "비철금속"
    RARE = ("리튬", "코발트", "망간", "니오븀", "규소", "규석", "실리콘", "마그네슘", "마그네사이트", "몰리브덴", "바나듐",
            "티타늄", "텅스텐", "안티모니", "창연", "비스무트", "크롬", "갈륨", "인듐",
            "탄탈", "지르코늄", "스트론튬", "셀레늄", "게르마늄")
    if any(k in n for k in RARE): return "희소금속"
    ETC = ("철", "흑연", "백금", "팔라듐", "금", "은")
    if any(k in n for k in ETC): return "기타"
    return None

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

NEWS_KEYWORDS = ["핵심광물", "리튬 광물", "코발트 광물", "니켈 광물", "희토류", "광물 공급망",
                 "텅스텐", "흑연 음극재", "광물 수출통제", "구리 제련", "희소금속"]

app    = Flask(__name__)
app.secret_key = SECRET_KEY

@app.after_request
def _no_html_cache(resp):
    # HTML은 항상 최신 렌더링을 받도록 — 배포 직후 '옛날 화면' 캐시 방지 (정적 파일은 캐시 유지)
    if resp.mimetype == "text/html":
        resp.headers["Cache-Control"] = "no-store, must-revalidate"
    return resp

def _conf_authed():
    """회의실 접근 허용 여부. 비밀번호 미설정(로컬)이면 항상 허용."""
    return (not CONFERENCE_PASSWORD) or (session.get("conf_ok") is True)

def _strip_surrogates(s):
    """이모지가 반쪽으로 잘린 서로게이트 문자를 제거 — UTF-8 인코딩 오류 방지."""
    return "".join(ch for ch in (s or "") if not 0xD800 <= ord(ch) <= 0xDFFF)

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
                            "한국광해광업공단_국가별 광종 수출입 현황_20251231.csv")
    if not os.path.exists(csv_path): return []
    try:
        df = None
        for _enc in ("utf-8-sig", "cp949", "euc-kr"):
            try:
                df = pd.read_csv(csv_path, encoding=_enc)
                break
            except (UnicodeDecodeError, UnicodeError):
                continue
        if df is None:
            return []
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
                    kk = str(k)
                    # 증감·비율·수지 등 파생 필드는 제외 (음수 총액 방지)
                    if any(x in kk for x in ("증감", "율", "비", "전년", "수지", "차액")):
                        continue
                    if "수입" in kk and ("금액" in kk or "액" in kk):
                        try: imp_amt = float(str(v).replace(",", "") or 0)
                        except: pass
                    if "수출" in kk and ("금액" in kk or "액" in kk):
                        try: exp_amt = float(str(v).replace(",", "") or 0)
                        except: pass
                if imp_amt < 0: imp_amt = 0   # 수입액은 음수가 될 수 없음
                if exp_amt < 0: exp_amt = 0
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

# ── 관세청 수출입 (nitemtrade) — HS코드 세트별 수집 ──
# KOMIR 통계(법정광물 28종)에 없는 전략 광물을 무역 통계로 보완
TRADE_SETS = {
    "ree":       {"2846": "희토류 화합물", "280530": "희토류 금속"},
    "strategic": {"8112": "갈륨·게르마늄·인듐 등", "8110": "안티모니", "8106": "창연(비스무트)",
                  "8104": "마그네슘", "8109": "지르코늄", "8108": "티타늄"},
    "precious":  {"7108": "금", "7106": "은", "7110": "백금족(백금·팔라듐)"},
    "uranium":   {"2844": "우라늄·방사성원소"},
}
TRADE_LABEL = {"ree": "희토류", "strategic": "전략 희소금속", "precious": "귀금속", "uranium": "우라늄"}

def _trade_snapshot(key):
    return os.path.join(os.path.dirname(__file__), f"trade_{key}_data1.json")

_trade_refreshing = set()

def fetch_trade_set(setkey):
    """HS 세트 수입 통계. 캐시 → 스냅샷 즉시 응답 + 백그라운드 갱신 → 최초엔 동기 수집."""
    c = cache_get(f"trade_{setkey}")
    if c is not None:
        return c
    snap = load_json(_trade_snapshot(setkey))
    if isinstance(snap, dict) and snap.get("by_country"):
        cache_set(f"trade_{setkey}", snap, ttl=43200)
        if CUSTOMS_API_KEY and setkey not in _trade_refreshing:
            _trade_refreshing.add(setkey)
            threading.Thread(target=_refresh_trade_set, args=(setkey,), daemon=True).start()
        return snap
    return _fetch_trade_now(setkey)

def _refresh_trade_set(setkey):
    try:
        _fetch_trade_now(setkey)
    finally:
        _trade_refreshing.discard(setkey)

def _fetch_trade_now(setkey):
    out = {"by_country": {}, "monthly": {}, "by_hs": {}, "by_item": {}, "asof": "", "total_imp": 0}
    ok = False
    if CUSTOMS_API_KEY:
        try:
            now = datetime.now()
            end = f"{now.year}{now.month:02d}"
            _m0 = now.year * 12 + (now.month - 1) - 11   # 12개월 창 (1년 이내 제한)
            strt = f"{_m0 // 12}{_m0 % 12 + 1:02d}"
            for hs, hs_name in TRADE_SETS[setkey].items():
                r = requests.get(
                    "https://apis.data.go.kr/1220000/nitemtrade/getNitemtradeList",
                    params={"serviceKey": CUSTOMS_API_KEY, "strtYymm": strt,
                            "endYymm": end, "hsSgn": hs}, timeout=20)
                if r.status_code != 200 or "<resultCode>00" not in r.text.replace(" ", ""):
                    continue
                root = ET.fromstring(r.text)
                for it in root.iter("item"):
                    g = lambda k: (it.findtext(k) or "").strip()
                    try: imp = float(g("impDlr") or 0)
                    except ValueError: imp = 0
                    cty = g("statCdCntnKor1")
                    ym = g("year").replace(".", "-")
                    if cty in ("총계", "합계", "", "-") or "총계" in ym: continue
                    if imp <= 0: continue
                    ok = True
                    out["by_country"][cty] = out["by_country"].get(cty, 0) + imp
                    out["monthly"][ym] = out["monthly"].get(ym, 0) + imp
                    out["by_hs"][hs_name] = out["by_hs"].get(hs_name, 0) + imp
                    itn = g("statKor")
                    if itn and itn != "-":
                        out["by_item"][itn] = out["by_item"].get(itn, 0) + imp
                time.sleep(0.15)
            if ok:
                out["asof"] = f"{strt[:4]}.{strt[4:]}~{end[:4]}.{end[4:]}"
                out["total_imp"] = round(sum(out["by_country"].values()))
                out["by_country"] = dict(sorted(out["by_country"].items(), key=lambda x: -x[1])[:10])
                out["monthly"] = dict(sorted(out["monthly"].items()))
                out["by_item"] = dict(sorted(out["by_item"].items(), key=lambda x: -x[1])[:8])
                try:
                    json.dump(out, open(_trade_snapshot(setkey), "w"), ensure_ascii=False)
                except Exception:
                    pass
        except Exception as e:
            print(f"[TRADE {setkey}]", e)
    if not ok:
        snap = load_json(_trade_snapshot(setkey))
        out = snap if isinstance(snap, dict) and snap.get("by_country") else out
    cache_set(f"trade_{setkey}", out, ttl=43200)
    return out

def fetch_ree_trade():
    return fetch_trade_set("ree")

# KOMIR 법정광물 교차검증용 — 광물별 HS코드(광석+금속 형태)
CORE_TRADE_HS = {
    "리튬":   ["282520", "283691"],   # 산화·수산화리튬 + 탄산리튬
    "니켈":   ["2604", "7502"],       # 니켈광 + 니켈괴
    "코발트": ["2605", "8105"],
    "텅스텐": ["2611", "8101"],
    "몰리브덴": ["2613", "8102"],
    "망간":   ["2602", "8111"],
    "동":     ["2603", "7403"],       # 동광 + 정제동
    "알루미늄": ["2606", "7601"],     # 보크사이트 + 알루미늄괴
    "아연":   ["2608", "7901"],
    "연":     ["2607", "7801"],
    "주석":   ["2609", "8001"],
    "철":     ["2601"],               # 철광석
    "흑연":   ["2504"],
    "크롬":   ["2610"],
    "규소":   ["280461", "280469"],
    "석탄":   ["2701"],
}
CORE_SNAPSHOT = os.path.join(os.path.dirname(__file__), "trade_core_data1.json")
_core_refreshing = [False]

def fetch_core_trade():
    """KOMIR 주요 광물의 관세청 월별 수입 — 스냅샷 즉시 응답 + 백그라운드 갱신."""
    c = cache_get("trade_core")
    if c is not None:
        return c
    snap = load_json(CORE_SNAPSHOT)
    if isinstance(snap, dict) and snap.get("minerals"):
        cache_set("trade_core", snap, ttl=43200)
        if CUSTOMS_API_KEY and not _core_refreshing[0]:
            _core_refreshing[0] = True
            threading.Thread(target=_refresh_core, daemon=True).start()
        return snap
    return _fetch_core_now()

def _refresh_core():
    try:
        _fetch_core_now()
    finally:
        _core_refreshing[0] = False

def _fetch_core_now():
    out = {"minerals": {}, "asof": ""}
    if not CUSTOMS_API_KEY:
        return out
    try:
        now = datetime.now()
        end = f"{now.year}{now.month:02d}"
        _m0 = now.year * 12 + (now.month - 1) - 11
        strt = f"{_m0 // 12}{_m0 % 12 + 1:02d}"
        out["asof"] = f"{strt[:4]}.{strt[4:]}~{end[:4]}.{end[4:]}"
        for mineral, hs_list in CORE_TRADE_HS.items():
            m = {"monthly": {}, "by_country": {}, "total": 0}
            for hs in hs_list:
                try:
                    r = requests.get(
                        "https://apis.data.go.kr/1220000/nitemtrade/getNitemtradeList",
                        params={"serviceKey": CUSTOMS_API_KEY, "strtYymm": strt,
                                "endYymm": end, "hsSgn": hs}, timeout=20)
                    if r.status_code != 200 or "<resultCode>00" not in r.text.replace(" ", ""):
                        continue
                    root = ET.fromstring(r.text)
                    for it in root.iter("item"):
                        g = lambda k: (it.findtext(k) or "").strip()
                        try: imp = float(g("impDlr") or 0)
                        except ValueError: imp = 0
                        cty = g("statCdCntnKor1")
                        ym = g("year").replace(".", "-")
                        if cty in ("총계", "합계", "", "-") or "총계" in ym or imp <= 0:
                            continue
                        m["monthly"][ym] = m["monthly"].get(ym, 0) + imp
                        m["by_country"][cty] = m["by_country"].get(cty, 0) + imp
                except Exception:
                    continue
                time.sleep(0.12)
            if m["monthly"]:
                m["total"] = round(sum(m["monthly"].values()))
                m["monthly"] = {k: round(v) for k, v in sorted(m["monthly"].items())}
                _bc = sorted(m["by_country"].items(), key=lambda x: -x[1])
                m["top"] = [_bc[0][0], round(_bc[0][1] / (m["total"] or 1) * 100)] if _bc else ["—", 0]
                m["by_country"] = dict(_bc[:6])
                out["minerals"][mineral] = m
        if out["minerals"]:
            try:
                json.dump(out, open(CORE_SNAPSHOT, "w"), ensure_ascii=False)
            except Exception:
                pass
    except Exception as e:
        print("[TRADE core]", e)
    cache_set("trade_core", out, ttl=43200)
    return out

def clean(t):
    s = re.sub(r"<[^>]+>", "", str(t))
    s = html.unescape(html.unescape(s))   # 이중 인코딩(&amp;lt;)까지 해제
    return s.strip()

def fetch_news():
    c = cache_get("news")
    if c is not None: return c
    if NAVER_CLIENT_ID.startswith("여기에"): return local_news()
    hdrs = {"X-Naver-Client-Id": NAVER_CLIENT_ID, "X-Naver-Client-Secret": NAVER_CLIENT_SECRET}
    all_news, seen = [], set()
    for kw in NEWS_KEYWORDS:
        try:
            r = requests.get("https://openapi.naver.com/v1/search/news.json",
                headers=hdrs, params={"query":kw,"display":6,"sort":"date"}, timeout=8)
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
                    "언론사링크": lnk, "링크": lnk,
                    "발행일시": dt, "발행일": dt,
                    "검색키워드": kw,
                })
        except: continue
        time.sleep(0.15)
    result = all_news if all_news else local_news()
    result = [n for n in result if news_relevant(n, MINERAL_NEWS_TERMS)]
    result = ai_relevance_gate(result, "핵심광물·금속 자원과 공급망, 소재 산업")
    cache_set("news", result)
    return result

def fetch_search_news(q, n=24):
    """통합검색용 — 입력어로 네이버 뉴스 검색 (관련도순)."""
    if not q or not NAVER_CLIENT_ID or NAVER_CLIENT_ID.startswith("여기에"): return []
    c = cache_get("search:" + q)
    if c is not None: return c
    hdrs = {"X-Naver-Client-Id": NAVER_CLIENT_ID, "X-Naver-Client-Secret": NAVER_CLIENT_SECRET}
    out, seen = [], set()
    try:
        r = requests.get("https://openapi.naver.com/v1/search/news.json",
            headers=hdrs, params={"query": q, "display": n, "sort": "sim"}, timeout=8)
        for it in r.json().get("items", []):
            lnk = it.get("originallink") or it.get("link", "")
            if not lnk or lnk in seen: continue
            seen.add(lnk)
            try: dt = datetime.strptime(it.get("pubDate", ""), "%a, %d %b %Y %H:%M:%S +0900").strftime("%Y-%m-%d")
            except: dt = it.get("pubDate", "")
            out.append({"제목": clean(it.get("title", "")), "요약": clean(it.get("description", "")),
                        "링크": lnk, "발행일": dt})
    except Exception as e:
        print(f"[SEARCH] {e}")
    cache_set("search:" + q, out, ttl=600)
    return out

def load_risk_data():
    try:
        with open(os.path.join(os.path.dirname(__file__), "risk_data.json"), encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return []

# ── K-RISK 종합 공급망 위험 점수 ──────────────────────────────
# 따로 발표되는 산업부 지표를 교차 계산해 광물별 공급위험을 0~100 단일 점수로 융합.
#   0.35×수급불안정(100−수급안정화지수) + 0.25×수입집중도(HHI)
# + 0.20×지정학(고위험국 수입비중·1위 생산국) + 0.20×가격변동성(파생지수 24개월 σ/μ)
# 시장위험지수·비축 항목은 공개 데이터 확보 시 가중 반영 예정 → 잔여 항목에 재배분한 산식.
K_GEO_RISK_COUNTRIES = {"중국", "러시아", "러시아연방", "콩고민주공화국", "미얀마"}
K_TOP_PRODUCER = {"리튬": "칠레", "니켈": "인도네시아", "코발트": "콩고민주공화국",
                  "동": "칠레", "텅스텐": "중국", "몰리브덴": "중국"}
K_MIDX_GROUP = {"리튬": "희소금속", "코발트": "희소금속", "텅스텐": "희소금속",
                "몰리브덴": "희소금속", "니켈": "메이저금속", "동": "메이저금속"}

def compute_k_risk():
    """K-RISK 48광종 전체 산출.
    - 수급안정화지수 제공 6광종: 4축 정식 점수 (0.35S+0.25H+0.20G+0.20V)
    - 나머지 42광종: 3축 잠정 점수 ((0.25H+0.20G+0.20V)/0.65 재정규화), "잠정" 표기
    """
    c = cache_get("k_risk")
    if c is not None: return c
    out = {}
    try:
        risk = {r["name"]: r.get("latest") for r in (load_risk_data() or [])
                if isinstance(r, dict) and r.get("latest") is not None}
        midx = load_json(os.path.join(os.path.dirname(__file__), "mineral_index_data2.json"))
        series = (midx.get("series") or {}) if isinstance(midx, dict) else {}
        vol = {}
        for g, sr in series.items():
            vals = [v for v in (sr.get("values") or [])[-24:] if isinstance(v, (int, float))]
            if len(vals) >= 6:
                m = sum(vals) / len(vals)
                sd = (sum((v - m) ** 2 for v in vals) / len(vals)) ** 0.5
                vol[g] = min(100.0, sd / m * 250) if m else 0.0

        # 수입 구조: 관세청 통계 + 그룹 스냅샷 폴백 (희토류·귀금속 등)
        imp_map, _u = by_mineral_country(fetch_customs() or [])
        try:
            groups = _v2_trade_groups()
        except Exception:
            groups = {"core": {}, "ree": {}, "strategic": {}, "precious": {}, "uranium": {}}

        usgs1 = {}
        try:
            usgs1 = _v2_usgs1() or {}
        except Exception:
            pass

        CAT_GROUP = {"비철금속": "메이저금속", "희소금속": "희소금속", "희토류": "희소금속",
                     "에너지": "에너지광물", "기타": "종합"}

        for cat, names in MINERAL_TAXONOMY.items():
            for name in names:
                base = re.sub(r"[\(\)/].*$", "", name)
                # ① 수입 구조 (직접 → 그룹 폴백)
                try:
                    byc, _src, _note = _v2_imports(name, imp_map, groups)
                except Exception:
                    byc = None
                byc = byc or {}
                tot = sum(byc.values())
                hhi = sum((v / tot) ** 2 for v in byc.values()) * 100 if tot else 50.0
                geo_share = (sum(v for c2, v in byc.items()
                                 if c2.replace(" ", "") in {g.replace(" ", "") for g in K_GEO_RISK_COUNTRIES})
                             / tot * 100 if tot else 0.0)
                # ② 세계 1위 생산국
                top = (K_TOP_PRODUCER.get(base) or K_TOP_PRODUCER.get(name)
                       or (USGS_DATA.get(base) or USGS_DATA.get(name) or {}).get("1위국") or "")
                if not top:
                    u = None
                    try:
                        u = _v2_lookup(usgs1, name)
                    except Exception:
                        pass
                    pt = (u or {}).get("prod_top") or []
                    top = pt[0] if pt else ""
                g_score = min(100.0, geo_share + (20.0 if top in K_GEO_RISK_COUNTRIES else 0.0))
                # ③ 변동성 (전용 매핑 → 분류 그룹)
                grp = K_MIDX_GROUP.get(base) or K_MIDX_GROUP.get(name) or CAT_GROUP.get(cat, "종합")
                v_score = vol.get(grp, vol.get("종합", 0.0))
                # ④ 수급불안정 (6광종만)
                ssi = risk.get(base) if risk.get(base) is not None else risk.get(name)
                if ssi is not None:
                    s_unstab = max(0.0, min(100.0, 100.0 - ssi))
                    score = max(0.0, min(100.0,
                        0.35 * s_unstab + 0.25 * hhi + 0.20 * g_score + 0.20 * v_score))
                    comp = {"수급불안정": round(s_unstab, 1), "수입집중도": round(hhi, 1),
                            "지정학": round(g_score, 1), "가격변동성": round(v_score, 1)}
                    prov = False
                else:
                    score = max(0.0, min(100.0,
                        (0.25 * hhi + 0.20 * g_score + 0.20 * v_score) / 0.65))
                    comp = {"수입집중도": round(hhi, 1), "지정학": round(g_score, 1),
                            "가격변동성": round(v_score, 1)}
                    prov = True
                grade = "위험" if score >= 70 else ("주의" if score >= 40 else "안정")
                out[name] = {"score": round(score, 1), "grade": grade, "1위국": top,
                             "요소": comp, "잠정": prov}
    except Exception as e:
        print(f"[K-RISK] 계산 오류: {e}")
    cache_set("k_risk", out, ttl=3600)
    return out

# 대상별 뉴스 — 같은 자원 이슈도 누구에게 보여줄지에 따라 다른 키워드
NEWS_AUDIENCE = {
    "투자자": ["2차전지 테마주", "핵심광물 수혜주", "희토류 관련주",
              "배터리 소재주", "리튬 관련주", "희소금속 투자"],
    "기업":   ["핵심광물 공급망", "핵심광물 수출규제", "광물 수급",
              "원자재 조달", "소재 국산화", "공급망 다변화"],
    "소비자": ["전기차 배터리 원자재", "리튬 가격", "니켈 가격",
              "배터리 가격", "전기차 가격 인상", "원자재 물가"],
    "정책":   ["핵심광물 확보전략", "자원안보", "광물 비축",
              "핵심광물 정책", "공급망 안정화", "자원 외교"],
}

# 광물 뉴스 관련성 필터 — 제목·요약에 아래 단어가 하나도 없으면 광물 뉴스로 보지 않는다
MINERAL_NEWS_TERMS = ("리튬", "니켈", "코발트", "희토류", "텅스텐", "망간", "흑연", "광물", "광산",
                      "광종", "제련", "2차전지", "양극재", "음극재", "희소금속", "몰리브덴",
                      "구리", "아연", "공급망", "원자재", "배터리")
# 검색어에 우연히 걸리는 무관 기사 차단 (운세의 '광물·비축', 병역'자원' 등)
NEWS_BLACKLIST = ("운세", "띠별", "사주", "별자리", "로또", "부고", "인사동정", "사관학교",
                  "오늘의 날씨", "TV 편성")
# 정책 뉴스는 광물 용어 외에 자원안보 계열 용어도 관련으로 인정
AUD_NEWS_TERMS = MINERAL_NEWS_TERMS + ("자원안보", "핵심광물", "확보전략", "국가 비축",
                                        "전략비축", "자원 협력", "자원외교")

def ai_relevance_gate(items, context):
    """뉴스 배치를 LLM이 한 번에 판정 — 관련 기사만 통과.
    키워드 필터(1차)를 통과한 목록에서 우연 매칭·무관 기사를 걸러내는 2차 게이트.
    LLM 실패 시 입력 그대로 반환(성능 저하 폴백), 과도한 전멸 응답도 무시."""
    if not items or not OPENAI_API_KEY:
        return items
    try:
        heads = "\n".join(f"{i}. {n.get('제목', '')} — {n.get('요약', '')}"
                           for i, n in enumerate(items))
        sysmsg = (f"너는 뉴스 큐레이터다. 아래 기사 중 '{context}'와 실질적으로 관련된 기사의 번호만 골라라. "
                  "단어만 우연히 겹치는 기사(운세·연예·스포츠·무관 정치·무관 군사 등)는 반드시 제외한다. "
                  '응답은 JSON {"keep": [번호, ...]} 형식만 출력한다.')
        r = OpenAI(api_key=OPENAI_API_KEY).chat.completions.create(
            model=DEFAULT_OPENAI_MODEL, max_completion_tokens=2000,
            response_format={"type": "json_object"},
            messages=[{"role": "system", "content": sysmsg},
                      {"role": "user", "content": heads}])
        keep = json.loads(r.choices[0].message.content or "{}").get("keep", [])
        keep = {int(i) for i in keep if str(i).isdigit()}
        sel = [n for i, n in enumerate(items) if i in keep]
        # 3건 이상에서 전멸 판정이 나오면 응답 오류로 보고 폴백
        if not sel and len(items) >= 3:
            return items
        return sel
    except Exception as e:
        print("[AI 관련성 게이트]", e)
        return items


def news_relevant(n, terms):
    t = (n.get("제목", "") or "") + " " + (n.get("요약", "") or "")
    if any(b in t for b in NEWS_BLACKLIST):
        return False
    return any(k in t for k in terms)

def mineral_relevant(n):
    return news_relevant(n, MINERAL_NEWS_TERMS)

def _fetch_audience_news(cache_key, aud_map):
    c = cache_get(cache_key)
    if c is not None: return c
    out = []
    if NAVER_CLIENT_ID and not NAVER_CLIENT_ID.startswith("여기에"):
        hdrs = {"X-Naver-Client-Id": NAVER_CLIENT_ID, "X-Naver-Client-Secret": NAVER_CLIENT_SECRET}
        seen = set()
        for aud, kws in aud_map.items():
            for kw in kws:
                try:
                    r = requests.get("https://openapi.naver.com/v1/search/news.json",
                        headers=hdrs, params={"query": kw, "display": 6, "sort": "date"}, timeout=8)
                    if r.status_code != 200: continue
                    for it in r.json().get("items", []):
                        lnk = it.get("originallink", "")
                        if lnk in seen: continue
                        seen.add(lnk)
                        try: dt = datetime.strptime(it.get("pubDate",""), "%a, %d %b %Y %H:%M:%S +0900").strftime("%Y-%m-%d %H:%M")
                        except: dt = it.get("pubDate","")
                        _item = {"제목": clean(it.get("title","")), "요약": clean(it.get("description",""))[:80],
                                 "언론사링크": lnk, "발행일시": dt, "검색키워드": kw, "aud": aud}
                        if not news_relevant(_item, AUD_NEWS_TERMS):
                            continue   # 검색어 오염(운세·무관 기사) 차단
                        out.append(_item)
                except: continue
                time.sleep(0.12)
        # 2차: 청중별 AI 관련성 판정 (키워드 필터 통과분만 대상 — 호출 4회/캐시 주기)
        _CTX = {"투자자": "광물·배터리·소재 산업의 투자 동향",
                "기업": "기업의 광물 원자재 조달·공급망 리스크",
                "소비자": "광물 가격이 전기차·전자제품 등 소비자 생활에 주는 영향",
                "정책": "핵심광물 자원안보·비축·통상 정책"}
        gated = []
        for aud in aud_map.keys():
            aud_items = [n for n in out if n.get("aud") == aud]
            gated += ai_relevance_gate(aud_items, _CTX.get(aud, "핵심광물 공급망"))
        out = gated
    cache_set(cache_key, out)
    return out

def dedup_news(items):
    seen, seen_ti, out = set(), set(), []
    for n in items:
        k = n.get("언론사링크") or n.get("제목", "")
        ti = "".join(ch for ch in (n.get("제목", "") or "") if ch.isalnum())[:10]
        if k in seen or (ti and ti in seen_ti): continue
        seen.add(k)
        if ti: seen_ti.add(ti)
        out.append(n)
    return out

# ── 지정학 상황실(3D 지구본)용 뉴스·이벤트 ──────────────────────
GEO_NEWS_KEYWORDS = ["수출통제", "경제제재", "공급망 차질", "홍해 해상운임", "중동 긴장",
                     "미중 갈등", "자원 무기화", "우크라이나 전쟁", "대만해협", "해협 봉쇄"]

GEO_EXTRA_LOCS = {
    "우크라이나": [49.0, 32.0], "이스라엘": [31.5, 34.9], "이란": [32.4, 53.7], "대만": [23.8, 121.0],
    "북한": [40.3, 127.0], "홍해": [19.0, 39.0], "호르무즈해협": [26.6, 56.5], "말라카해협": [2.5, 101.5],
    "수에즈운하": [30.4, 32.4], "남중국해": [13.0, 114.0], "파나마운하": [9.1, -79.7],
    "유럽연합": [50.5, 4.5], "사우디아라비아": [24.0, 45.0], "카타르": [25.3, 51.2], "이라크": [33.2, 43.7],
    "예멘": [15.5, 47.5], "미얀마": [21.0, 96.0], "베트남": [16.0, 107.5], "멕시코": [23.5, -102.0],
    "영국": [52.5, -1.5], "프랑스": [46.5, 2.5], "독일": [51.0, 10.3], "튀르키예": [39.0, 35.0],
    "한국": [36.5, 127.9], "홍콩": [22.3, 114.2], "싱가포르": [1.35, 103.8],
}

def _geo_locations():
    locs = {k.replace(" ", ""): v for k, v in COUNTRY_COORDS.items()}
    locs.update(GEO_EXTRA_LOCS)
    return locs

def fetch_geo_news():
    return _fetch_audience_news("geo_news", {"지정학": GEO_NEWS_KEYWORDS})

@app.route("/api/geo-events")
def geo_events():
    c = cache_get("geo_events")
    if c is not None:
        return jsonify(ok=bool(c), events=c or [])
    arts = dedup_news(fetch_geo_news())[:24]
    if not arts or not OPENAI_API_KEY:
        cache_set("geo_events", [], ttl=300)
        return jsonify(ok=False, events=[])
    locs = _geo_locations()
    heads = "\n".join(f"{i}. {n.get('제목','')} — {n.get('요약','')}" for i, n in enumerate(arts))
    events = []
    try:
        r = OpenAI(api_key=OPENAI_API_KEY).chat.completions.create(
            model=DEFAULT_OPENAI_MODEL, max_completion_tokens=1600,
            messages=[
                {"role": "system", "content":
                    "너는 자원안보 지정학 분석가다. 뉴스 목록에서 '한국 자원 공급망에 영향을 줄 수 있는' 지정학 이벤트만 골라 분류한다. "
                    "결과는 JSON 배열만 출력한다. 각 원소: {\"i\": 기사번호, \"loc\": 장소명, \"type\": \"전쟁|제재|외교|공급차질|시위\", "
                    "\"sev\": 1~3(심각도), \"res\": \"관련 자원(리튬·원유 등, 없으면 빈문자열)\", \"why\": \"한국 공급망 영향 1문장\"}. "
                    "loc은 반드시 다음 목록 중에서만 고른다: " + ", ".join(sorted(locs.keys())) + ". "
                    "무관하거나 장소를 특정할 수 없는 기사는 제외한다."},
                {"role": "user", "content": heads},
            ])
        txt = (r.choices[0].message.content or "").strip()
        m = re.search(r"\[.*\]", txt, re.S)
        for e in (json.loads(m.group(0)) if m else []):
            try:
                idx = int(e.get("i", -1))
                loc = str(e.get("loc", "")).replace(" ", "")
                if not (0 <= idx < len(arts)) or loc not in locs:
                    continue
                a = arts[idx]
                events.append({
                    "lat": locs[loc][0], "lng": locs[loc][1], "loc": loc,
                    "type": e.get("type", "외교"), "sev": max(1, min(3, int(e.get("sev", 1)))),
                    "res": str(e.get("res", ""))[:30], "why": str(e.get("why", ""))[:120],
                    "title": a.get("제목", ""), "link": a.get("언론사링크", ""),
                    "date": a.get("발행일시", ""),
                })
            except Exception:
                continue
    except Exception as ex:
        print("[GEO EVENTS]", ex)
    cache_set("geo_events", events, ttl=1800)
    return jsonify(ok=bool(events), events=events)

def fetch_audience_news():        return _fetch_audience_news("anews",   NEWS_AUDIENCE)

def by_mineral(rows):
    s = {}
    for r in rows:
        nm = r.get("광물명","기타")
        try: v = float(str(r.get("수입금액(달러)",0)).replace(",","") or 0)
        except: v = 0
        if v < 0: v = 0   # 증감치 등 오염 데이터 방어
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
        c.execute("ALTER TABLE subscribers ADD COLUMN IF NOT EXISTS minerals TEXT")
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

SUBM_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "subscribers_minerals.json")

def _subm_load():
    try:
        with open(SUBM_FILE, encoding="utf-8") as f: return json.load(f)
    except Exception: return {}

def _subm_save(m):
    with open(SUBM_FILE, "w", encoding="utf-8") as f:
        json.dump(m, f, ensure_ascii=False, indent=2)

def get_sub_minerals(email):
    """구독자의 관심 광물 목록."""
    if DATABASE_URL:
        try:
            _ensure_db()
            with _db_conn() as c:
                row = c.execute("SELECT minerals FROM subscribers WHERE email=%s", (email,)).fetchone()
            return json.loads(row[0]) if (row and row[0]) else []
        except Exception as e:
            print("[DB] get_sub_minerals:", e); return []
    return _subm_load().get(email, [])

def set_sub_minerals(email, minerals):
    if DATABASE_URL:
        try:
            _ensure_db()
            with _db_conn() as c:
                c.execute("UPDATE subscribers SET minerals=%s WHERE email=%s",
                          (json.dumps(minerals, ensure_ascii=False), email))
                c.commit()
            return
        except Exception as e:
            print("[DB] set_sub_minerals:", e); return
    m = _subm_load(); m[email] = minerals; _subm_save(m)

def load_subs_full():
    """[{email, minerals}] — 일일 발송용."""
    if DATABASE_URL:
        try:
            _ensure_db()
            with _db_conn() as c:
                rows = c.execute("SELECT email, minerals FROM subscribers ORDER BY created_at").fetchall()
            return [{"email": r[0], "minerals": (json.loads(r[1]) if r[1] else [])} for r in rows]
        except Exception as e:
            print("[DB] load_subs_full:", e); return []
    m = _subm_load()
    return [{"email": e, "minerals": m.get(e, [])} for e in load_subs()]

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
    m = _subm_load()
    if email in m:
        del m[email]; _subm_save(m)
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
        from email.header import Header
        from email.utils import formataddr
        msg["Subject"] = subj; msg["From"] = formataddr((str(Header(MAIL_FROM_NAME, "utf-8")), MAIL_FROM)); msg["To"] = to
        msg.attach(MIMEText(html, "html", "utf-8"))
        with smtplib.SMTP(SMTP_HOST, SMTP_PORT) as s:
            s.starttls(); s.login(SMTP_USER, SMTP_PASS)
            s.sendmail(MAIL_FROM, to, msg.as_string())
        return True, "OK"
    except Exception as e: return False, str(e)

def _kst_now():
    """서버 시간대와 무관하게 한국 시간 반환 (Render=UTC 대비)."""
    from datetime import timezone, timedelta
    return datetime.now(timezone.utc).astimezone(timezone(timedelta(hours=9))).replace(tzinfo=None)


def _nl_claim_today():
    """오늘자 발송 소유권 획득 — DB(워커 간 공유) 우선, 실패 시 메모리 캐시. KST 날짜 기준."""
    day = _kst_now().strftime("%Y-%m-%d")
    if DATABASE_URL:
        try:
            with _db_conn() as c:
                c.execute("CREATE TABLE IF NOT EXISTS newsletter_log (day TEXT PRIMARY KEY, sent_at TIMESTAMP DEFAULT now())")
                cur = c.execute("INSERT INTO newsletter_log (day) VALUES (%s) ON CONFLICT DO NOTHING", (day,))
                c.commit()
                return cur.rowcount == 1
        except Exception as e:
            print("[newsletter] DB 가드 실패, 캐시 폴백:", e)
    # DB 없음 → 파일 마커 (재시작해도 유지되어 재발송 방지)
    guard = os.path.join(os.path.dirname(os.path.abspath(__file__)), ".newsletter_last_sent")
    try:
        with open(guard, encoding="utf-8") as f:
            if f.read().strip() == day:
                return False
    except FileNotFoundError:
        pass
    with open(guard, "w", encoding="utf-8") as f:
        f.write(day)
    return True


def _send_daily_all():
    """구독자 전원에게 오늘의 광물 날씨 발송."""
    subs = load_subs_full()
    now = _kst_now()
    subj = f"[K Mineral Risk] {now.month}월 {now.day}일 광물 날씨"
    sent = failed = 0
    for su in subs:
        e = su["email"]
        ok, info = send_mail(e, subj, build_newsletter(e, su.get("minerals") or []))
        sent += ok; failed += (not ok)
        if not ok:
            print("[newsletter] 발송 실패:", e, info)
    print(f"[newsletter] 발송 완료 {sent}/{len(subs)} (실패 {failed})")
    return {"sent": sent, "failed": failed, "total": len(subs)}


def build_newsletter(to=None, minerals=None):
    """V2 '오늘의 광물 날씨' 이메일 — 표 기반 HTML. minerals=관심 광물(개인화 블록)."""
    if minerals is None and to:
        try: minerals = get_sub_minerals(to)
        except Exception: minerals = []
    minerals = minerals or []
    G, GD, GL = "#155BB8", "#16305C", "#EAF2FC"
    INKC, MUTC, LINE = "#222222", "#555555", "#DDE3EA"
    DGR, DGL, WRN, WRL = "#d8453c", "#fdecea", "#c97a00", "#fdf3e2"
    base = APP_BASE_URL or "http://127.0.0.1:8081"
    now = _kst_now()
    wd = "월화수목금토일"[now.weekday()]

    rows = []
    try:
        rows = _v2_rows()
    except Exception as e:
        print("[NEWSLETTER rows]", e)
    n_dg = sum(1 for r in rows if r["grade"] == "위험")
    n_wr = sum(1 for r in rows if r["grade"] == "주의")
    top = next((r for r in rows if r["grade"]), None)

    # 히어로 문장
    if top and top["grade"] == "위험":
        h_bg, h_fg, h_t = DGL, DGR, f"{top['name']}이 위험 단계예요"
    elif top and top["grade"] == "주의":
        h_bg, h_fg, h_t = WRL, WRN, f"{top['name']}을 지켜봐야 해요"
    else:
        h_bg, h_fg, h_t = GL, GD, "오늘 광물 시장은 대체로 맑아요"
    h_b = ""
    if top:
        if top.get("share") and top["share"] >= 50:
            h_b = f"수입의 {top['share']}%를 {top['top']} 한 나라에 의존하고 있어요. "
        h_b += f"{top['use']} 가격에 영향을 줄 수 있어요."
    else:
        h_b = "48개 광물의 수급·가격·수입선을 매일 살펴보고 있어요."

    # 상위 광물 8종 표
    gcol = {"위험": DGR, "주의": WRN, "안정": GD}
    lis = ""
    for r in rows[:8]:
        g = r["grade"] or "관찰"
        cs = gcol.get(r["grade"], MUTC)
        score = f" {r['score']:.0f}" if isinstance(r.get("score"), (int, float)) else ""
        from urllib.parse import quote as _q
        lis += (
            f'<tr><td style="padding:10px 4px;border-bottom:1px solid {LINE};">'
            f'<a href="{base}/m/{_q(r["name"], safe="")}" style="color:{INKC};text-decoration:none;font-weight:700;font-size:14px;">{r["name"]}</a>'
            f' <span style="color:{MUTC};font-size:12px;">{r["use"]}</span>'
            f'<div style="color:{MUTC};font-size:12px;margin-top:2px;">{r["sub"]}</div></td>'
            f'<td style="padding:10px 4px;border-bottom:1px solid {LINE};text-align:right;white-space:nowrap;">'
            f'<span style="color:{cs};font-weight:800;font-size:13.5px;">{g}{score}</span></td></tr>'
        )
    lis = lis or f'<tr><td style="color:{MUTC};padding:10px 4px;">데이터 준비 중</td></tr>'

    # 내 관심 광물 (개인화)
    my_html = ""
    if minerals:
        rmap = {r["name"]: r for r in rows}
        items = ""
        from urllib.parse import quote as _q2
        for mn in minerals[:8]:
            r = rmap.get(mn)
            if not r: continue
            g = r["grade"] or "관찰"
            cs = gcol.get(r["grade"], MUTC)
            score = f" {r['score']:.0f}" if isinstance(r.get("score"), (int, float)) else ""
            try: mnews = _v2_news(mn, 2)
            except Exception: mnews = []
            nrows = "".join(
                f'<div style="font-size:12.5px;padding:4px 0 0 10px;line-height:1.5;">· '
                f'<a href="{n.get("링크", "#")}" style="color:{INKC};text-decoration:none;">{n.get("제목", "")}</a>'
                f' <span style="color:{MUTC};font-size:11px;">{(n.get("발행일") or "")[:10]}</span></div>'
                for n in mnews)
            items += (
                f'<div style="border:1px solid {LINE};border-radius:12px;padding:12px 15px;margin:0 0 9px;">'
                f'<table style="width:100%;border-collapse:collapse;"><tr>'
                f'<td><a href="{base}/m/{_q2(mn, safe="")}" style="color:{INKC};text-decoration:none;font-weight:800;font-size:14px;">{mn}</a>'
                f' <span style="color:{MUTC};font-size:12px;">{r["use"]}</span>'
                f'<div style="color:{MUTC};font-size:12px;margin-top:2px;">{r["sub"]}</div></td>'
                f'<td style="text-align:right;white-space:nowrap;vertical-align:top;">'
                f'<span style="color:{cs};font-weight:800;font-size:13.5px;">{g}{score}</span></td></tr></table>'
                f'{nrows}</div>')
        if items:
            my_html = (f'<div style="font-size:14px;font-weight:800;color:{GD};margin:0 0 8px;">'
                       f'★ 내 관심 광물</div>{items}<div style="height:10px;"></div>')

    # AI 브리핑(캐시에 있을 때만) · 지정학 이벤트(캐시)
    brief = cache_get("news_brief_minerals") or ""
    brief_html = (
        f'<div style="background:{GL};border-radius:12px;padding:14px 18px;margin:0 0 18px;">'
        f'<div style="font-size:11.5px;font-weight:800;color:{GD};margin-bottom:5px;">AI 애널리스트 브리핑</div>'
        f'<div style="font-size:13px;color:{GD};line-height:1.65;">{brief}</div></div>'
    ) if brief else ""
    geo = (cache_get("geo_events") or [None])
    geo = geo[0] if geo else None
    geo_html = ""
    if isinstance(geo, dict) and (geo.get("why") or geo.get("title")):
        geo_html = (
            f'<div style="background:#F5F7FA;border-radius:12px;padding:13px 18px;margin:0 0 18px;">'
            f'<div style="font-size:11.5px;font-weight:800;color:{MUTC};margin-bottom:4px;">🌍 지금 세계에선</div>'
            f'<div style="font-size:13px;color:{INKC};line-height:1.6;">'
            f'{(geo.get("loc") + " — ") if geo.get("loc") else ""}{geo.get("why") or geo.get("title")}</div></div>'
        )

    # 뉴스 6건 (광물 관련성 필터)
    news = [n for n in dedup_news(fetch_news() or []) if mineral_relevant(n)][:6]
    news_html = "".join(
        f'<div style="padding:10px 0;border-bottom:1px solid {LINE};">'
        f'<a href="{n.get("링크", "#")}" style="color:{INKC};text-decoration:none;font-weight:650;font-size:13.5px;line-height:1.5;">{n.get("제목", "")}</a>'
        f'<div style="color:{MUTC};font-size:11.5px;margin-top:3px;">{n.get("발행일", "")}</div></div>'
        for n in news
    ) or f'<div style="color:{MUTC};font-size:13px;">오늘은 새 소식이 없어요.</div>'

    _link = unsub_link(to) if to else ""
    _unsub = f' · <a href="{_link}" style="color:#9aa39d;">수신거부</a>' if _link else ""

    return (
        f'<div style="max-width:620px;margin:0 auto;font-family:\'Apple SD Gothic Neo\',\'Malgun Gothic\',sans-serif;background:#F5F7FA;padding:18px 12px;">'
        # 헤더
        f'<div style="background:{GD};background-image:linear-gradient(120deg,#8C3242 0%,#2E2F72 40%,#16305C 70%,#155BB8 100%);border-radius:16px 16px 0 0;padding:22px 26px;">'
        f'<div style="color:#fff;font-size:19px;font-weight:800;">● K Mineral Risk — 오늘의 광물 날씨</div>'
        f'<div style="color:#B9CCEA;font-size:12.5px;margin-top:4px;">{now.year}년 {now.month}월 {now.day}일 {wd}요일</div></div>'
        # 본문
        f'<div style="background:#ffffff;border:1px solid {LINE};border-top:0;border-radius:0 0 16px 16px;padding:22px 26px;">'
        # 히어로
        f'<div style="background:{h_bg};border-radius:12px;padding:16px 18px;margin-bottom:14px;">'
        f'<div style="font-size:16px;font-weight:800;color:{h_fg};">{h_t}</div>'
        f'<div style="font-size:13px;color:{h_fg};margin-top:5px;line-height:1.6;">{h_b}</div></div>'
        # 카운트
        f'<div style="font-size:12.5px;color:{MUTC};margin:0 0 16px;">오늘 48개 광물 중 '
        f'<b style="color:{DGR};">위험 {n_dg}</b> · <b style="color:{WRN};">주의 {n_wr}</b></div>'
        f'{my_html}{brief_html}{geo_html}'
        # 리스트
        f'<div style="font-size:14px;font-weight:800;color:{INKC};margin:0 0 4px;">주목 광물 TOP 8</div>'
        f'<table style="width:100%;border-collapse:collapse;">{lis}</table>'
        # 뉴스
        f'<div style="font-size:14px;font-weight:800;color:{INKC};margin:20px 0 4px;">오늘의 소식</div>'
        f'{news_html}'
        # CTA
        f'<div style="text-align:center;margin:24px 0 6px;">'
        f'<a href="{base}" style="display:inline-block;background:{GD};color:#fff;text-decoration:none;'
        f'font-size:14px;font-weight:800;border-radius:999px;padding:12px 28px;">광물 날씨 전체 보기 →</a></div>'
        f'</div>'
        # 푸터
        f'<div style="padding:16px 8px;text-align:center;color:#9aa39d;font-size:11.5px;">'
        f'K Mineral Risk · 데이터: KOMIR·관세청·조달청·산업부·USGS·World Bank{_unsub}</div></div>'
    )


def build_welcome(to=None, minerals=None):
    """구독 직후 보내는 환영 메일 — 배너 + 오늘자 리포트(관심 광물 반영)."""
    banner = (
        '<div style="max-width:620px;margin:0 auto;font-family:\'Apple SD Gothic Neo\',\'Malgun Gothic\',sans-serif;padding:18px 12px 0;">'
        '<div style="background:#EAF2FC;border-radius:16px;padding:20px 24px;margin-bottom:6px;">'
        '<div style="font-size:17px;font-weight:800;color:#16305C;">구독을 환영해요 🎉</div>'
        '<div style="font-size:13.5px;color:#222222;margin-top:7px;line-height:1.65;">'
        '이제 매일 아침 <b>오늘의 광물 날씨</b>가 이 주소로 도착해요.<br>'
        '위험해진 광물, 수입선 변화, 꼭 봐야 할 소식만 골라 보내드릴게요.<br>'
        '아래는 오늘자 리포트 미리보기예요.</div></div></div>'
    )
    return banner + build_newsletter(to, minerals)


# ═══════════════════════════════════════════════════════════════
#  ② 대시보드 HTML 렌더링
# ═══════════════════════════════════════════════════════════════
# ── V2 공통 크롬 — 통계 대시보드에 본 사이트 헤더·푸터 입히기 ──
V2_CHROME_CSS = r"""
.v2ubar,.v2tbar,.v2foot{width:100%!important;min-width:100%;flex:0 0 auto!important;align-self:stretch!important;box-sizing:border-box}
.v2wrap{max-width:1200px;margin:0 auto;padding:0 24px;width:100%;box-sizing:border-box}
.v2ubar{background:#F2F4F7;border-bottom:1px solid #DDE3EA;font-size:12px;font-family:'Noto Sans KR','Pretendard',sans-serif}
.v2ubar .v2wrap{display:flex;justify-content:space-between;align-items:center;height:36px}
.v2ubar a{color:#555;text-decoration:none}.v2ubar a:hover{color:#155BB8;text-decoration:underline}
.v2ubar .l{color:#888}
.v2ubar .r a{padding:0 11px}
.v2tbar{border-bottom:2px solid;border-image:linear-gradient(100deg,#C24E59 0%,#C24E59 34%,#0047A0 66%,#0047A0 100%) 1;background:#fff;font-family:'Noto Sans KR','Pretendard',sans-serif;position:relative;z-index:80}
.v2tbar::before{content:'';display:block;height:4px;background:linear-gradient(100deg,#C24E59 0%,#C24E59 44%,#0047A0 56%,#0047A0 100%)}
.v2tbar .v2wrap{display:flex;align-items:center;height:86px;gap:18px}
.v2logo{display:flex;align-items:center;gap:11px;flex:none;text-decoration:none}
.v2logo .dot{width:30px;height:30px;border-radius:8px;background:linear-gradient(135deg,#C24E59,#0047A0);box-shadow:inset 0 0 0 3px rgba(255,255,255,.22)}
.v2logo b{font-size:21px;font-weight:900;letter-spacing:-.5px;color:#16305C;display:block;line-height:1.25}
.v2logo b em{font-style:normal;background:linear-gradient(135deg,#C24E59 15%,#7E4468 50%,#0047A0 85%);-webkit-background-clip:text;background-clip:text;-webkit-text-fill-color:transparent;color:#0047A0}
.v2logo i{display:block;font-style:normal;font-size:11px;color:#888;letter-spacing:.4px;font-weight:500}
.v2gnb{display:flex;flex:1}
.v2gnb>a,.v2drop>a{display:flex;align-items:center;height:86px;padding:0 16px;font-size:15.5px;font-weight:700;color:#222;position:relative;white-space:nowrap;text-decoration:none}
.v2gnb>a::after,.v2drop>a::after{content:'';position:absolute;left:16px;right:16px;bottom:0;height:3px;background:linear-gradient(100deg,#C24E59,#0047A0);transform:scaleX(0);transition:transform .2s}
.v2gnb>a:hover,.v2drop:hover>a{color:#155BB8}
.v2gnb>a:hover::after,.v2drop:hover>a::after{transform:scaleX(1)}
.v2drop{position:relative;display:flex}
.v2menu{display:none;position:absolute;left:50%;transform:translateX(-50%);top:86px;background:#fff;border:1px solid #DDE3EA;border-top:2px solid #155BB8;min-width:184px;padding:8px 0;box-shadow:0 10px 24px rgba(22,48,92,.12);z-index:90}
.v2drop:hover .v2menu{display:block}
.v2menu a{display:block;padding:9px 20px;font-size:14px;font-weight:500;color:#555;text-decoration:none}
.v2menu a:hover{background:#EAF2FC;color:#155BB8;font-weight:700}
.v2menu-w{padding:14px 6px 12px!important}
.v2drop:hover .v2menu-w{display:flex!important;gap:4px}
.v2menu-w .vcol{min-width:172px}
.v2menu-w .vt{font-size:11.5px;font-weight:800;color:#888;letter-spacing:.06em;padding:0 20px 7px}
.v2foot{background:#16305C;color:#B9CCEA;font-size:13px;font-family:'Noto Sans KR','Pretendard',sans-serif;margin-top:40px;position:relative;z-index:5;border-top:3px solid;border-image:linear-gradient(90deg,#C24E59,#0047A0) 1}
.v2foot .fin{max-width:1200px;margin:0 auto;padding:24px}
.v2foot a{color:#B9CCEA;text-decoration:none}.v2foot a:hover{color:#fff}
.v2foot .fb{border-top:1px solid rgba(255,255,255,.14);text-align:center;padding:12px;font-size:11.5px;color:#8FA6C6}
.cat-bar .brand-lock{display:none!important}
.cat-bar .cb-right{display:none!important}
.megapanel{display:none!important}
.v2tbar{z-index:2000!important}
.v2menu{z-index:9999!important}
/* 카테고리 줄: 구 탭 디자인 폐기 → 알약 칩 */
.cat-bar{background:#fff!important;border-bottom:1px solid #E8ECF1!important;padding:12px 24px!important;gap:9px!important;box-shadow:none!important}
.cat-bar .cat-btn{background:#fff!important;border:1.5px solid #D9DEE8!important;border-radius:999px!important;
padding:9px 22px!important;font-size:14px!important;font-weight:700!important;color:#444!important;
font-family:'Noto Sans KR','Pretendard',sans-serif!important;box-shadow:0 1px 3px rgba(22,48,92,.06)!important;
text-decoration:none!important;letter-spacing:0!important}
.cat-bar .cat-btn::after,.cat-bar .cat-btn::before{display:none!important}
.cat-bar .cat-btn:hover{border-color:#155BB8!important;color:#155BB8!important}
.cat-bar .cat-btn.active{background:#16305C!important;border-color:#16305C!important;color:#fff!important}
/* 탭 포커스 모드: 통계 탭 컨텍스트로 카테고리 진입 시 해당 기능 섹션만 표시 */
.catpage.tabfocus .stat-row{display:none!important}
.catpage.tabfocus .charts-row{display:none!important}
.catpage.tabfocus .charts-row.row-focus{display:flex!important}
.catpage.tabfocus .charts-row.row-focus > *{display:none!important}
.catpage.tabfocus .charts-row.row-focus > .sec-focus{display:block!important;flex:1!important}
@media(max-width:900px){.v2ubar{display:none}.v2gnb{display:none}.v2tbar .v2wrap{height:64px}}
"""

V2_CHROME_HEADER = """
<div class="v2tbar"><div class="v2wrap">
  <a class="v2logo" href="/"><svg width="36" height="36" viewBox="0 0 40 40" aria-hidden="true"><defs>
<linearGradient id="kmrg" x1="0" y1="0" x2="0.85" y2="1">
<stop offset="0" stop-color="#D66671"/><stop offset=".42" stop-color="#B84C59"/><stop offset=".58" stop-color="#2A3F8F"/><stop offset="1" stop-color="#0047A0"/></linearGradient></defs>
<polygon points="20,2 36,11 36,29 20,38 4,29 4,11" fill="url(#kmrg)"/>
<polygon points="20,2 36,11 20,20 4,11" fill="#ffffff" opacity=".14"/>
<polygon points="20,20 36,11 36,29 20,38" fill="#000000" opacity=".12"/>
<polyline points="8,22 14,22 17,13 22,29 25,20 32,20" fill="none" stroke="#ffffff" stroke-width="2.6" stroke-linecap="round" stroke-linejoin="round"/>
<circle cx="32" cy="20" r="2.6" fill="#FFD24A"/></svg><span><b><em>K</em> Mineral Risk</b><i>핵심광물 공급망 위험 진단 · KOREA CMR INTELLIGENCE</i></span></a>
  <nav class="v2gnb">
    <a href="/">홈</a>
    <a href="/globe">핵심광물지도</a>
    <div class="v2drop"><a href="#" onclick="return false">통계</a><div class="v2menu v2menu-w">
      <div class="vcol"><div class="vt">지표별</div>
        <a href="/dashboard#supply">수급 현황</a><a href="/dashboard#mindex">가격지수</a>
        <a href="/dashboard#forecast">가격 전망</a><a href="/dashboard#map">글로벌 매장량</a>
        <a href="/dashboard#routes">수입 루트</a><a href="/dashboard#risk">리스크 신호등</a><a href="/dashboard#mines">국내 광산</a>
      </div>
      <div class="vcol"><div class="vt">광종별</div>
        <a href="/dashboard#cat-minerals">핵심광물 종합</a><a href="/dashboard#cat-nf">비철금속 (6종)</a>
        <a href="/dashboard#cat-rare">희소금속 (20종)</a><a href="/dashboard#cat-ree">희토류 (14종)</a>
        <a href="/dashboard#cat-energy">에너지 (2종)</a><a href="/dashboard#cat-etc">기타 (6종)</a>
      </div>
    </div></div>
    <a href="/briefing">브리핑</a>
    <a href="/conference">AI 회의</a>
  </nav>
</div></div>
"""

V2_CHROME_FOOTER = """
<div class="v2foot"><div class="fin">
  <b style="color:#fff">K Mineral Risk</b> — 흩어진 광물 공공데이터를 융합해 공급망 위험을 하나의 지수로 진단합니다 ·
  <a href="/globe">핵심광물지도</a> · <a href="/briefing">브리핑</a> · <a href="/conference">AI 회의실</a> · <a href="/minerals.csv">데이터(CSV)</a>
</div><div class="fb">본 서비스는 산업통상자원부·산하기관 공공데이터를 활용합니다 · 팀 SMART-X © 2026 K Mineral Risk</div></div>
"""


def render_dashboard(home=False):
    body_cls = "is-home" if home else ""
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

    # 대상별 뉴스 — 광물 기사(기본) + 광물 대상별 기사 병합 후 관련성 필터
    anews = fetch_audience_news()
    _mmerged = dedup_news(sorted([n for n in news[:12] + anews if mineral_relevant(n)],
                                 key=lambda n: n.get("발행일시", ""), reverse=True))
    news_js = json.dumps(_mmerged, ensure_ascii=False)

    # ── 자원 리스크 신호등 (수급안정화지수, 한국광해광업공단) ──
    risk = load_risk_data()
    risk_js = json.dumps(risk, ensure_ascii=False)
    def _sig(v):
        if v is None: return ("—", "#888")
        if v >= 55: return ("안정", "#1e8e5a")
        if v >= 30: return ("주의", "#b58a12")
        return ("위험", "#d64545")
    def _risk_card(r):
        lab, col = _sig(r["latest"])
        prev = r.get("prev")
        if prev is not None:
            d = r["latest"] - prev
            sub = f'전월 {"▲" if d>0 else ("▼" if d<0 else "·")} {abs(d):.1f}'
        else:
            sub = ""
        return (f'<div class="risk-card" style="border-left:4px solid {col}">'
                f'<div class="rk-top"><span class="rk-nm">{r["name"]}</span>'
                f'<span class="rk-tag" style="background:{col}22;color:{col}">{lab}</span></div>'
                f'<div class="rk-val">{r["latest"]:.1f}<span>/100</span></div>'
                f'<div class="rk-sub">{sub} · 수급안정화지수</div></div>')
    risk_cards = "".join(_risk_card(r) for r in risk) or '<div class="empty">리스크 데이터 없음</div>'
    _risk_high = [r["name"] for r in risk if r["latest"] < 30]
    risk_summary = ("현재 <b style=\"color:#d64545\">" + " · ".join(_risk_high) + "</b> 의 수급 불안이 높습니다."
                    if _risk_high else "현재 주요 광물 수급은 비교적 안정적입니다.")

    # ── K-RISK 종합 공급망 위험 점수 카드 ──
    krisk = compute_k_risk()
    def _krisk_card(name, d):
        col = "#d64545" if d["grade"] == "위험" else ("#b58a12" if d["grade"] == "주의" else "#1e8e5a")
        ico = "🔴" if d["grade"] == "위험" else ("🟡" if d["grade"] == "주의" else "🟢")
        p = d["요소"]
        _cid = {"비철금속": "nf", "희소금속": "rare", "희토류": "ree",
                "에너지": "energy", "기타": "etc"}.get(mineral_category(name), "etc")
        _pv = ' <span style="font-size:10px;color:#888;font-weight:600">잠정</span>' if d.get("잠정") else ''
        _s1 = (f'수급불안정 {p["수급불안정"]:.0f} · ' if "수급불안정" in p else '')
        return (f'<div class="risk-card" data-cat="{_cid}" style="border-left:4px solid {col}">'
                f'<div class="rk-top"><span class="rk-nm">{name}</span>{_pv}'
                f'<span class="rk-tag" style="background:{col}22;color:{col}">{ico} {d["grade"]}</span></div>'
                f'<div class="rk-val" style="color:{col}">{d["score"]:.1f}<span>/100</span></div>'
                f'<div class="rk-sub" style="line-height:1.7">{_s1}수입집중 {p["수입집중도"]:.0f}<br>'
                f'지정학 {p["지정학"]:.0f} · 변동성 {p["가격변동성"]:.0f}</div></div>')
    krisk_cards = ("".join(_krisk_card(k, v) for k, v in
                           sorted(krisk.items(), key=lambda x: -x[1]["score"]))
                   or '<div class="empty">K-RISK 계산 데이터 없음</div>')
    _kr_red = [k for k, v in krisk.items() if v["grade"] == "위험"]
    _kr_yel = [k for k, v in krisk.items() if v["grade"] == "주의"]
    if _kr_red:
        krisk_summary = 'K-RISK 기준 <b style="color:#d64545">' + " · ".join(_kr_red) + "</b> 이(가) 위험(🔴) 단계입니다."
    elif _kr_yel:
        krisk_summary = 'K-RISK 기준 <b style="color:#b58a12">' + " · ".join(_kr_yel) + "</b> 이(가) 주의(🟡) 단계입니다."
    else:
        krisk_summary = "K-RISK 기준 주요 광물 공급위험은 안정(🟢) 범위입니다."

    # ── 광물 가격지수 (한국광해광업공단 파생지수, 2012~ 월별) ──
    midx = load_json(os.path.join(os.path.dirname(__file__), "mineral_index_data2.json"))
    if not isinstance(midx, dict):
        midx = {}
    midx_js = json.dumps(midx, ensure_ascii=False)
    def _midx_card(key, label, color):
        s = (midx.get("summary") or {}).get(key)
        if not s:
            return ""
        mom = s.get("mom") or 0
        mtxt = f'{"▲" if mom>0 else ("▼" if mom<0 else "·")} {abs(mom):.1f}% 전월'
        yoy = s.get("yoy")
        if yoy is not None:
            ycol = "#1e8e5a" if yoy >= 0 else "#d64545"
            ytxt = f'<span style="color:{ycol}">{"▲" if yoy>0 else "▼"} {abs(yoy):.1f}% 전년</span>'
        else:
            ytxt = ""
        return (f'<div class="risk-card" style="border-left:4px solid {color}">'
                f'<div class="rk-top"><span class="rk-nm">{label}</span></div>'
                f'<div class="rk-val">{s["latest"]:,.0f}</div>'
                f'<div class="rk-sub">{mtxt} · {ytxt} · {s["asof"]}</div></div>')
    midx_cards = "".join([
        _midx_card("종합", "광물종합지수", "#c98500"),
        _midx_card("에너지광물", "에너지광물 (연료탄·우라늄)", "#eb6834"),
        _midx_card("희소금속", "희소금속 (리튬·희토류)", "#155BB8"),
        _midx_card("메이저금속", "메이저금속 (철·동·니켈)", "#4a3aa7"),
    ]) or '<div class="empty">지수 데이터 없음</div>'

    # ── 데이터1 확장 데이터 로드 ──
    _pj0 = lambda f: os.path.join(os.path.dirname(__file__), f)
    forecast = load_json(_pj0("forecast_data1.json")) or {}
    reserves = load_json(_pj0("reserves_data1.json")) or []
    ppa      = load_json(_pj0("ppa_data1.json")) or {}
    steel    = load_json(_pj0("steel_data1.json")) or {}
    mines    = load_json(_pj0("mines_data1.json")) or {}
    outlook  = load_json(_pj0("outlook_data1.json")) or {}
    forecast_js = json.dumps(forecast, ensure_ascii=False)
    steel_js    = json.dumps(steel, ensure_ascii=False)
    mines_js    = json.dumps(mines, ensure_ascii=False)
    outlook_js  = json.dumps(outlook, ensure_ascii=False)

    def _tons(v):
        if v >= 1e8: return f"{v/1e8:,.1f}억t"
        if v >= 1e4: return f"{v/1e4:,.0f}만t"
        return f"{v:,.0f}t"

    # 오늘의 금속 시세 (조달청 비축물자 · LME)
    _pp = ppa.get("items") or []
    ppa_rows = "".join(
        f'<div class="pp-row"><span class="pp-nm">{i["name"]}</span>'
        f'<span class="pp-val">${(i["close"] or 0):,.0f}</span>'
        f'<span class="pp-chg" style="color:{"#d64545" if (i["chg"] or 0)>0 else ("#1e8e5a" if (i["chg"] or 0)<0 else "#888888")}">'
        f'{"▲" if (i["chg"] or 0)>0 else ("▼" if (i["chg"] or 0)<0 else "·")} {abs(i["chg"] or 0):.2f}%</span></div>'
        for i in _pp) or '<div class="empty">시세 데이터 없음</div>'
    _lme = ppa.get("lme") or {}
    ppa_lme_s = f'{(_lme.get("idx") or 0):,.1f}' if _lme.get("idx") else "—"
    ppa_date = ppa.get("date", "")

    # 홈 금속 시세 스트립
    home_metal_html = ("".join(
        f'<div class="hm-card"><span class="hm-nm">{i["name"]}</span>'
        f'<b class="hm-val">${(i["close"] or 0):,.0f}</b>'
        f'<span class="hm-chg" style="color:{"#d64545" if (i["chg"] or 0)>0 else ("#1e8e5a" if (i["chg"] or 0)<0 else "#888888")}">'
        f'{"▲" if (i["chg"] or 0)>0 else ("▼" if (i["chg"] or 0)<0 else "·")}{abs(i["chg"] or 0):.2f}%</span></div>'
        for i in _pp)) if _pp else ""

    # 광종 분류 스트립 (확대 커버리지)
    taxo_html = "".join(
        f'<div class="tx-row"><span class="tx-cat" style="background:{TAXO_COLOR[cat]}14;color:{TAXO_COLOR[cat]};border-color:{TAXO_COLOR[cat]}55">{cat}</span>'
        + '<span class="tx-list">' + "".join(f'<i>{m}</i>' for m in ms) + '</span>'
        + f'<span class="tx-cnt">{len(ms)}종</span></div>'
        for cat, ms in MINERAL_TAXONOMY.items())
    taxo_total = sum(len(v) for v in MINERAL_TAXONOMY.values())

    # ── 분류별 카테고리 페이지 데이터 ──
    _fc_cat = {}   # cat_id -> forecast 광종 리스트
    for _m in forecast.keys():
        _c = mineral_category(_m)
        for cid, cname, _, _ in CAT_DEFS:
            if cname == _c:
                _fc_cat.setdefault(cid, []).append(_m)
    catfc_js = json.dumps(_fc_cat, ensure_ascii=False)

    def _bar_rows(pairs, unit="$", fmt="{:,.0f}"):
        if not pairs: return '<div class="empty">데이터 없음</div>'
        mx = max(v for _, v in pairs) or 1
        return "".join(
            f'<div class="rk-item"><span class="rk-no">{i+1:02d}</span>'
            f'<div class="rk-mid"><div class="rk-nm">{n}</div>'
            f'<div class="rk-bar"><div class="rk-fill" style="width:{max(v/mx*100,2):.0f}%"></div></div></div>'
            f'<span class="rk-vl">{unit}{fmt.format(v)}</span></div>'
            for i, (n, v) in enumerate(pairs))

    # 분류별 관세청 무역 세트 (KOMIR 미포함 광물 보완)
    CAT_TRADE = {"rare": "strategic", "etc": "precious", "energy": "uranium"}
    def _trade_panel(setkey):
        td = fetch_trade_set(setkey)
        bc = list((td.get("by_country") or {}).items())[:8]
        if not bc:
            return (f'<div class="section" style="flex:1;padding:14px 16px;">'
                    f'<div class="chart-title">{TRADE_LABEL[setkey]} 수입 · 관세청</div>'
                    f'<div class="empty">관세청 API 연동 대기 중</div></div>')
        c_rows = "".join(
            f'<div class="rk-item"><span class="rk-no">{i+1:02d}</span>'
            f'<div class="rk-mid"><div class="rk-nm">{c}</div>'
            f'<div class="rk-bar"><div class="rk-fill" style="width:{max(v/(bc[0][1] or 1)*100,2):.0f}%"></div></div></div>'
            f'<span class="rk-vl">${v:,.0f}</span></div>'
            for i, (c, v) in enumerate(bc))
        i_rows = "".join(
            f'<div class="pp-row"><span class="pp-nm" title="{k}">{(k[:26] + "…") if len(k) > 27 else k}</span><span class="pp-val">${v:,.0f}</span></div>'
            for k, v in list((td.get("by_item") or td.get("by_hs") or {}).items())[:8])
        return (f'<div class="section" style="flex:1.2;padding:14px 16px;">'
                f'<div class="chart-title">{TRADE_LABEL[setkey]} 국가별 수입 <span style="color:var(--muted2)">· 관세청 · {td.get("asof","")} · 연 ${td.get("total_imp",0):,.0f}</span></div>'
                f'{c_rows}</div>'
                f'<div class="section" style="flex:1;padding:14px 16px;">'
                f'<div class="chart-title">{TRADE_LABEL[setkey]} 품목 구성 <span style="color:var(--muted2)">· HS 기준</span></div>'
                f'{i_rows}</div>')

    # USGS MCS 2026 — 생산·매장 카드
    usgs2 = load_json(_pj0("usgs_data1.json")) or {}
    def _ut(v, unit):
        """단위별 톤 환산 표기"""
        if not v: return "—"
        if unit == "kilograms":
            return f"{v/1000:,.0f}t" if v >= 1000 else f"{v:,.0f}kg"
        t = v * 1000 if unit == "thousand metric tons" else v
        if t >= 1e8: return f"{t/1e8:,.1f}억t"
        if t >= 1e4: return f"{t/1e4:,.0f}만t"
        return f"{t:,.0f}t"
    usgs2_cards = "".join(
        f'<div class="uc"><div class="uc-nm">{k}</div>'
        f'<div class="uc-row"><span class="uc-lb">연 생산</span><span class="uc-vl">{_ut(v["prod_total"], v["unit"])}</span></div>'
        f'<div class="uc-row"><span class="uc-lb">생산 1위</span><span class="uc-vl hi">{v["prod_top"][0]} {v["prod_top"][1]}%</span></div>'
        + (f'<div class="uc-row"><span class="uc-lb">매장량</span><span class="uc-vl">{_ut(v["rsv_total"], v["unit"])}</span></div>' if v["rsv_total"] else '')
        + f'<div class="uc-row"><span class="uc-lb">매장 1위</span><span class="uc-vl">{v["rsv_top"][0]} {v["rsv_top"][1]}%</span></div>'
        f'<div class="uc-src">USGS MCS 2026 · 2025년 기준</div></div>'
        for k, v in sorted(usgs2.items(), key=lambda x: x[0]))
    # World Bank 국제 시세
    wbp = load_json(_pj0("wb_prices_data1.json")) or {}
    wbp_js = json.dumps(wbp, ensure_ascii=False)

    # KOMIR 주요 광물 — 관세청 월별 수입 + 교차 검증
    core = fetch_core_trade()
    _cm = core.get("minerals") or {}
    core_js = json.dumps(_cm, ensure_ascii=False)
    core_asof = core.get("asof", "")
    _bm_map = dict(bm)
    def _komir_of(m):
        if m == "석탄":
            return sum(_bm_map.get(k, 0) for k in ("유연탄", "무연탄", "갈탄", "토탄"))
        if m == "흑연":
            return _bm_map.get("인상흑연", 0) + _bm_map.get("토상흑연", 0)
        return _bm_map.get(m, 0)
    cross_rows = ""
    for m, v in sorted(_cm.items(), key=lambda x: -x[1]["total"]):
        kv = _komir_of(m)
        if kv:
            ratio = v["total"] / kv
            rtxt = f"×{ratio:,.2f}"
            rcol = "#1e8e5a" if 0.5 <= ratio <= 2 else "#b58a12"
        else:
            rtxt, rcol = "미집계", "#888888"
        cross_rows += (
            f'<tr><td class="t-nm" style="padding:7px 4px;font-weight:600">{m}</td>'
            f'<td class="t-num" style="padding:7px 4px">${v["total"]/1e6:,.0f}M</td>'
            f'<td class="t-num" style="padding:7px 4px;color:var(--muted)">{("$" + format(kv/1e6, ",.0f") + "M") if kv else "—"}</td>'
            f'<td class="t-num" style="padding:7px 4px;color:{rcol};font-weight:700">{rtxt}</td>'
            f'<td class="t-nm" style="padding:7px 4px;font-size:11.5px;color:var(--muted)">{v["top"][0]} {v["top"][1]}%</td></tr>')

    ree = fetch_ree_trade()
    ree_js = json.dumps(ree.get("monthly") or {}, ensure_ascii=False)
    _ree_bc = list((ree.get("by_country") or {}).items())
    ree_country_rows = "".join(
        f'<div class="rk-item"><span class="rk-no">{i+1:02d}</span>'
        f'<div class="rk-mid"><div class="rk-nm">{c}</div>'
        f'<div class="rk-bar"><div class="rk-fill" style="width:{max(v/(_ree_bc[0][1] or 1)*100,2):.0f}%"></div></div></div>'
        f'<span class="rk-vl">${v:,.0f}</span></div>'
        for i, (c, v) in enumerate(_ree_bc)) or         '<div class="empty">관세청 API 승인 대기 중 — 키 활성화 후 자동 표시됩니다 (HS 2846·280530)</div>'
    _ree_hs_rows = "".join(
        f'<div class="pp-row"><span class="pp-nm">{k}</span><span class="pp-val">${v:,.0f}</span></div>'
        for k, v in list((ree.get("by_item") or ree.get("by_hs") or {}).items())[:8])

    cat_pages_html = ""
    for cid, cname, cicon, cdesc in CAT_DEFS:
        col = TAXO_COLOR[cname]
        _imp = [(n, v) for n, v in bm if mineral_category(n) == cname][:8]
        _imp_total = sum(v for _, v in _imp)
        if cid == "ree" and ree.get("total_imp"):
            _imp = [(k, v) for k, v in sorted((ree.get("by_hs") or {}).items(), key=lambda x: -x[1])]
            _imp_total = ree["total_imp"]
        _rsv = [x for x in reserves if x["cat"] == cname][:8]
        _rsv_rows = _bar_rows([(x["name"], x["total"]) for x in _rsv], unit="", fmt="") if False else "".join(
            f'<div class="rk-item"><span class="rk-no">{i+1:02d}</span>'
            f'<div class="rk-mid"><div class="rk-nm">{x["name"]}</div>'
            f'<div class="rk-bar"><div class="rk-fill" style="width:{max(x["total"]/(_rsv[0]["total"] or 1)*100,2):.0f}%"></div></div></div>'
            f'<span class="rk-vl">{_tons(x["total"])}</span></div>'
            for i, x in enumerate(_rsv)) or '<div class="empty">매장량 데이터 없음</div>'
        _rk_cards = "".join(_risk_card(r) for r in risk if mineral_category(r["name"]) == cname)
        _kr_cards = "".join(_krisk_card(k, v) for k, v in sorted(krisk.items(), key=lambda x: -x[1]["score"])
                            if mineral_category(k) == cname)
        _badges = "".join(f'<i>{m}</i>' for m in MINERAL_TAXONOMY.get(cname, []))
        _top_imp = _imp[0][0] if _imp else "—"
        _top_rsv = (_rsv[0]["name"].split("(")[0]) if _rsv else "—"
        _n_fc = len(_fc_cat.get(cid, []))
        _ppa_panel = (f'<div class="section" style="padding:14px 16px;">'
                      f'<div class="chart-title">오늘의 LME 시세 · {ppa_date}</div>{ppa_rows}</div>') if cid == "nf" else ""
        if cid in CAT_TRADE:
            _ppa_panel = _trade_panel(CAT_TRADE[cid])
        if cid == "ree":
            _ppa_panel = (
                f'<div class="section" style="flex:1.3;padding:14px 16px;">'
                f'<div class="chart-title">국가별 수입액 <span style="color:var(--muted2)">· 관세청 · {ree.get("asof") or "최근 12개월"} · HS 2846+280530</span></div>'
                f'{ree_country_rows}'
                + (f'<div style="margin-top:10px;border-top:1px solid var(--border);padding-top:8px">{_ree_hs_rows}</div>' if _ree_hs_rows else '')
                + f'</div>'
                f'<div class="section" style="flex:1;padding:14px 16px;">'
                f'<div class="chart-title">월별 수입 추이 <span style="color:var(--muted2)">· 달러</span></div>'
                f'<div style="height:250px;position:relative;"><canvas id="reeChart"></canvas></div></div>')
        _fc_panel = (f'<div class="section" style="padding:14px 16px;">'
                     f'<div class="chart-title" id="fcCatTitle_{cid}">가격 전망 — 실선=실측 · 점선=예측</div>'
                     f'<div style="display:flex;gap:6px;flex-wrap:wrap;margin:4px 0 10px" id="fcCatBtns_{cid}"></div>'
                     f'<div style="height:260px;position:relative;"><canvas id="fcCatChart_{cid}"></canvas></div></div>') if _n_fc else ""
        _risk_panel = (f'<div class="page-title" style="margin-top:18px">🚦 공급 리스크 <span style="color:var(--muted2);font-weight:400;font-size:12px">· K-RISK · 수급안정화지수</span></div>'
                       f'<div class="risk-grid">{_kr_cards}{_rk_cards}</div>') if (_kr_cards or _rk_cards) else ""
        cat_pages_html += f"""
<div id="cat-{cid}" class="catpage" style="display:none">
  <div class="page-title" style="font-size:15px!important">{cname} <span style="color:var(--muted2);font-weight:400;font-size:12px">· {cdesc}</span></div>
  <div class="stat-row">
    <div class="stat-card"><div class="sc-label">{cname} 수입액 합계</div><div class="sc-val">{("$" + format(_imp_total, ",.0f")) if _imp_total else "—"}</div><div class="sc-sub">{("관세청 · 최근 12개월" if cid == "ree" else "KOMIR · 최신연도") if _imp_total else "관세청 API 승인 대기"}</div></div>
    <div class="stat-card"><div class="sc-label">커버 광종</div><div class="sc-val">{len(MINERAL_TAXONOMY.get(cname, []))}<small style="font-size:14px;font-weight:600">종</small></div><div class="sc-sub">확대 대상 기준</div></div>
    <div class="stat-card"><div class="sc-label">최대 수입 광물</div><div class="sc-val" style="font-size:20px">{_top_imp}</div><div class="sc-sub">수입액 1위</div></div>
    <div class="stat-card"><div class="sc-label">최대 매장 광물</div><div class="sc-val" style="font-size:20px">{_top_rsv}</div><div class="sc-sub">세계 매장량 기준</div></div>
    <div class="stat-card"><div class="sc-label">가격 전망 제공</div><div class="sc-val">{_n_fc}<small style="font-size:14px;font-weight:600">광종</small></div><div class="sc-sub">~2028 예측</div></div>
  </div>
  <div class="charts-row" style="height:auto!important;align-items:stretch;">
    <div class="section" id="sec-{cid}-supply" style="flex:1;padding:14px 16px;">
      <div class="chart-title">광종별 수입액 <span style="color:var(--muted2)">· KOMIR</span></div>
      {_bar_rows(_imp)}
    </div>
    <div class="section" id="sec-{cid}-map" style="flex:1;padding:14px 16px;">
      <div class="chart-title">글로벌 매장량 순위 <span style="color:var(--muted2)">· KOMIR 2026</span></div>
      {_rsv_rows}
    </div>
  </div>
  <div class="charts-row" style="height:auto!important;align-items:stretch;margin-top:14px;">
    <span id="sec-{cid}-forecast" style="position:absolute"></span>
    {_fc_panel}
    <span id="sec-{cid}-mindex" style="position:absolute"></span>
    {_ppa_panel}
    <div class="section" style="flex:1;padding:14px 16px;">
      <div class="chart-title">{cname} 광종 <span style="color:var(--muted2)">· 커버리지</span></div>
      <div class="tx-list" style="line-height:2.3;padding-top:4px">{_badges}</div>
      <div style="font-size:12px;color:var(--muted2);margin-top:10px;line-height:1.7">분류 색상 <span style="display:inline-block;width:10px;height:10px;border-radius:3px;background:{col};vertical-align:-1px"></span> · 수입액·매장량·전망·리스크는 보유 데이터 기준으로 자동 필터링됩니다.</div>
    </div>
  </div>
  <div style="text-align:center;margin:18px 0 6px;"><a href="/conference" class="nav-conf">⚖️ AI 전문가 회의실에서 {cname} 토론하기 →</a></div>
</div>"""

    # 국내 광산 통계 요약
    _mst = (mines.get("stats") or {})
    _mcomp = (_mst.get("comp") or {})
    _mg = _mcomp.get("가행") or {}
    mine_active = sum(_mg.values())
    mine_closed_total = (mines.get("closed") or {}).get("total", 0)
    mine_latest_year = _mst.get("latest_year", "")
    mine_sido_rows = "".join(
        f'<div class="rk-item"><span class="rk-no">{i+1:02d}</span>'
        f'<div class="rk-mid"><div class="rk-nm">{x["s"]}</div>'
        f'<div class="rk-bar"><div class="rk-fill" style="width:{x["n"]/((mines.get("closed") or {}).get("sido") or [{"n":1}])[0]["n"]*100:.0f}%"></div></div></div>'
        f'<span class="rk-vl">{x["n"]:,}<small>개</small></span></div>'
        for i, x in enumerate((mines.get("closed") or {}).get("sido") or [])) or '<div class="empty">데이터 없음</div>'

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

    # 수급 대시보드 위젯 — 매장량 순위 / 매장 1위국 (KOMIR 2026 확대판, 폴백 USGS)
    if reserves:
        _rk = reserves[:12]
        _rmax = _rk[0]["total"] if _rk else 1
        usgs_rank_html = "".join(
            f'<div class="rk-item"><span class="rk-no">{i+1:02d}</span>'
            f'<div class="rk-mid"><div class="rk-nm">{x["name"]} <i class="rk-cat" style="color:{TAXO_COLOR.get(x["cat"], "#888888")}">{x["cat"]}</i></div>'
            f'<div class="rk-bar"><div class="rk-fill" style="width:{max(x["total"]/_rmax*100, 2):.0f}%"></div></div></div>'
            f'<span class="rk-vl">{_tons(x["total"])}</span></div>'
            for i, x in enumerate(_rk))
        prod_html = "".join(
            f'<div class="pr-item"><span class="pr-nm">{x["name"]}</span><span class="pr-co">{x["top1"]}</span></div>'
            for x in reserves[:14])
    else:
        _usgs_sorted = sorted(USGS_DATA.items(), key=lambda kv: kv[1]["매장량_만톤"], reverse=True)
        _usgs_max = _usgs_sorted[0][1]["매장량_만톤"] if _usgs_sorted else 1
        usgs_rank_html = "".join(
            f'<div class="rk-item"><span class="rk-no">{i+1:02d}</span>'
            f'<div class="rk-mid"><div class="rk-nm">{mn}</div>'
            f'<div class="rk-bar"><div class="rk-fill" style="width:{info["매장량_만톤"]/_usgs_max*100:.0f}%"></div></div></div>'
            f'<span class="rk-vl">{info["매장량_만톤"]:,}<small>만t</small></span></div>'
            for i, (mn, info) in enumerate(_usgs_sorted))
        prod_html = "".join(
            f'<div class="pr-item"><span class="pr-nm">{mn}</span><span class="pr-co">{info["1위국"]}</span></div>'
            for mn, info in USGS_DATA.items())

    komir_rows = "".join(f"""<tr>
      <td class="t-nm">{r.get('광물명','')}</td>
      <td class="t-nm">{r.get('국가명','')}</td>
      <td class="t-num">${float(str(r.get('수입금액(달러)',0)).replace(',','') or 0):,.0f}</td>
      <td class="t-num">${float(str(r.get('수출금액(달러)',0)).replace(',','') or 0):,.0f}</td>
    </tr>""" for r in customs[:30]) if customs else \
    '<tr><td colspan="4" class="empty">KOMIR 데이터 없음</td></tr>'

    # ── 실시간 리스크 티커 (K-RISK 계산값) ──
    def _tick_ico(g):
        return "🔴" if g == "위험" else ("🟡" if g == "주의" else "🟢")
    _tick = []
    for k, v in sorted(krisk.items(), key=lambda x: -x[1]["score"])[:5]:
        _tick.append(f"{_tick_ico(v['grade'])} K-RISK {k} {v['score']:.0f} {v['grade']}")
    _tick.append("산업부 공공데이터 실시간 교차 계산")
    ticker_items = " &nbsp;|&nbsp; ".join(_tick)

    # ── 메인(홈) 오늘의 리스크 스트립 ──
    def _hl_rk(icon, label, href, nm, score, grade):
        col = "#d64545" if grade == "위험" else ("#b58a12" if grade == "주의" else "#1e8e5a")
        gi = "🔴" if grade == "위험" else ("🟡" if grade == "주의" else "🟢")
        return (f'<a class="hl-rk" href="{href}"><span class="hl-rk-cat">{icon} {label}</span>'
                f'<b style="color:{col}">{nm} {score:.1f}</b>'
                f'<span class="hl-rk-gr" style="color:{col}">{gi} {grade} · 자세히 →</span></a>')
    _hl_parts = []
    for _hk, _hv in sorted(krisk.items(), key=lambda x: -x[1]["score"])[:3]:
        _hl_parts.append(_hl_rk("🔩", "핵심광물 K-RISK", "/dashboard?cat=minerals&sec=risk", _hk, _hv["score"], _hv["grade"]))
    _hl_parts.append('<a class="hl-rk cta" href="/conference">⚖️ 이 위험들, AI 전문가 회의실에서 토론 →</a>')
    home_risk_html = (
        '<div class="hl-risk"><div class="hl-risk-head">오늘의 리스크 '
        '<span>산업부 공공데이터 실시간 교차 계산 — 위험을 감지하면 회의가 소집됩니다</span></div>'
        f'<div class="hl-risk-row">{"".join(_hl_parts)}</div></div>'
    ) if krisk else ""

    DASH_OVERRIDE = r"""
/* === MINETECH 라이트 디자인 시스템 (화이트 × 네이비 × 앰버골드) === */
:root{
  --bg:#F5F7FA;--bg2:#ffffff;--bg3:#EEF1F5;
  --border:#DDE3EA;--border2:#C9D2DD;
  --text:#222222;--muted:#555555;--muted2:#888888;
  --red:#d64545;--red-dim:#fdeeee;--red-bright:#c03535;
  --accent:#16305C;--accent2:#16305C;
  --navy:#16305C;--navy2:#155BB8;
  --blue:#155BB8;--cyan:#155BB8;--green:#1e8e5a;
  --gold:#155BB8;--gold-soft:#faf3dd;
  --shadow:0 1px 2px rgba(20,35,60,.04),0 8px 24px rgba(20,35,60,.07);
  --mono:'JetBrains Mono','IBM Plex Mono',monospace;
  --sans:'Inter','Noto Sans KR',sans-serif;
}
body{background:var(--bg);padding-left:256px;color:var(--text);}
.stat-card,.chart-box,.section,.sub-box,.uc{
  background:#fff!important;border:1px solid var(--border)!important;border-radius:14px!important;box-shadow:var(--shadow)!important;
}
.ticker{background:linear-gradient(90deg,#fdf9ee,#faf3dd 50%,#fdf9ee)!important;border-bottom:1px solid #ecdfb8!important;}
.ticker-inner{color:#8a6a10!important;text-shadow:none!important;}
.nav{position:fixed!important;left:0;top:0;width:256px;height:100vh!important;flex-direction:column;align-items:stretch;
  gap:3px;padding:20px 14px!important;background:var(--bg2)!important;border-right:1px solid var(--border)!important;
  border-bottom:none!important;overflow-y:auto;z-index:200;}
.nav-brand{font-size:14px!important;color:var(--navy)!important;margin:0 0 22px 4px!important;text-shadow:none!important;letter-spacing:.08em!important;}
.nav-brand .sys-dot{background:var(--gold)!important;box-shadow:none!important;}
/* ===== 섹션 네비 ===== */
#subnav-minerals{counter-reset:nav;position:relative;}
#subnav-minerals::before{
  content:'';position:absolute;left:22px;top:16px;bottom:16px;width:1px;
  background:linear-gradient(180deg,transparent,rgba(21,91,184,.35),transparent);}
.nav a[data-tab]{position:relative;display:flex!important;align-items:center;gap:12px;width:100%;
  padding:11px 12px 11px 46px!important;border-radius:11px!important;font-size:13.5px!important;font-weight:600;
  color:var(--muted)!important;border:0!important;transition:.24s cubic-bezier(.2,.7,.3,1);overflow:hidden;}
.nav a[data-tab]::before{counter-increment:nav;content:counter(nav,decimal-leading-zero);
  position:absolute;left:11px;top:50%;transform:translateY(-50%);
  width:22px;height:22px;display:flex;align-items:center;justify-content:center;
  font-family:var(--mono);font-size:10px;font-weight:700;color:var(--muted2);
  background:var(--bg3);border:1px solid var(--border);border-radius:7px;transition:.24s;z-index:1;}
.nav a[data-tab]:hover{background:var(--gold-soft)!important;color:var(--text)!important;padding-left:50px!important;}
.nav a[data-tab]:hover::before{border-color:var(--gold);color:var(--accent);}
.nav a[data-tab].active{background:linear-gradient(90deg,rgba(21,91,184,.16),rgba(21,91,184,.02))!important;
  color:var(--navy)!important;font-weight:800;box-shadow:inset 3px 0 0 var(--gold);}
.nav a[data-tab].active::before{background:linear-gradient(135deg,#B9CCEA,#155BB8);color:#fff;
  border-color:var(--gold);}
.nav-right{margin-top:auto!important;margin-left:0!important;flex-direction:column;align-items:stretch;gap:10px;padding-top:14px;border-top:1px solid var(--border);}
.nav-time{color:var(--muted)!important;font-size:10px!important;text-align:center;}
.nav-conf{text-align:center;color:#fff!important;background:var(--navy)!important;border:none!important;font-weight:700!important;padding:10px!important;border-radius:8px!important;}
.nav-conf:hover{opacity:.92;background:var(--navy)!important;}
.sidebar{background:var(--bg2)!important;border-right:1px solid var(--border)!important;}
.stat-card.red{border-color:rgba(214,69,69,.35)!important;background:#fdf5f4!important;}
.sc-val.red{color:var(--red)!important;}
.t-num,.kp-amount{color:var(--accent)!important;}
.bf{background:var(--gold)!important;}
.mode-btn.active,.mineral-btn.active{background:var(--navy)!important;color:#fff!important;border-color:var(--navy)!important;}
.mineral-btn:hover,.mode-btn:hover{border-color:var(--navy2)!important;color:var(--text)!important;}
.nc:hover{border-color:var(--gold)!important;}
.nc-kw,.uc-nm{color:var(--accent)!important;}
.sub-btn{background:var(--navy)!important;color:#fff!important;}
.sub-btn:hover{background:var(--navy2)!important;}
.sub-input:focus{border-color:var(--navy2)!important;}
.kp-title{text-shadow:none!important;color:var(--navy)!important;}

/* ── 상단 카테고리 바 ── */
.cat-bar{flex-shrink:0;display:flex;align-items:center;gap:8px;padding:10px 18px;background:var(--bg2);border-bottom:1px solid var(--border);}
.cat-bar .cb-label{font-size:10px;color:var(--muted2);font-family:var(--mono);text-transform:uppercase;letter-spacing:.12em;margin-right:6px;}
.cat-btn{padding:8px 20px;border-radius:8px;border:1px solid var(--border2);background:var(--bg3);color:var(--muted);font-size:13px;font-weight:700;cursor:pointer;font-family:var(--mono);transition:.15s;}
.cat-btn:hover{color:var(--text);border-color:var(--gold);}
.cat-btn.active{background:var(--navy);color:#fff;border-color:var(--navy);}
#cat-minerals{flex:1;min-height:0;display:flex;flex-direction:column;}
#tab-risk{flex-direction:column;overflow-y:auto;padding:16px;}
.risk-grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(180px,1fr));gap:12px;}
.risk-card{background:#fff;border:1px solid var(--border);border-radius:12px;padding:14px 16px;}
.rk-top{display:flex;align-items:center;justify-content:space-between;margin-bottom:8px;}
.rk-nm{font-size:14px;font-weight:700;color:var(--text);}
.rk-tag{font-size:11px;font-weight:700;padding:2px 9px;border-radius:10px;}
.rk-val{font-size:26px;font-weight:800;color:var(--text);}
.rk-val span{font-size:12px;font-weight:400;color:var(--muted2);margin-left:2px;}
.rk-sub{font-size:11px;color:var(--muted);margin-top:4px;font-family:var(--mono);}

/* ── 등장 애니메이션 ── */
@keyframes cardIn{from{opacity:0;transform:translateY(12px) scale(.98);}to{opacity:1;transform:none;}}
.tab-panel.active .stat-card,
.risk-card{animation:cardIn .45s cubic-bezier(.2,.7,.3,1) both;}
.stat-card:nth-child(2),.risk-card:nth-child(2){animation-delay:.06s;}
.stat-card:nth-child(3),.risk-card:nth-child(3){animation-delay:.12s;}
.stat-card:nth-child(4),.risk-card:nth-child(4){animation-delay:.18s;}
.risk-card:nth-child(5){animation-delay:.24s;}
.risk-card:nth-child(6){animation-delay:.30s;}
.risk-card:hover,.stat-card:hover{transform:translateY(-3px);box-shadow:0 14px 34px rgba(20,35,60,.12);}
.risk-card,.stat-card{transition:transform .15s,box-shadow .15s,border-color .15s;}
@keyframes barGrow{from{width:0;}}
.ct-bar i{animation:barGrow .8s cubic-bezier(.2,.7,.3,1) both;}
.news-hero{display:block;background:linear-gradient(135deg,#fdf7e7,#ffffff 55%);border:1px solid #eddfb5;border-left:4px solid var(--gold);border-radius:14px;padding:22px 26px;margin-bottom:16px;text-decoration:none;animation:cardIn .45s ease both;transition:transform .15s,box-shadow .15s;}
.news-hero:hover{transform:translateY(-2px);box-shadow:0 12px 30px rgba(20,35,60,.12);}
.nh-badge{display:inline-block;background:var(--gold);color:#fff;font-size:11px;font-weight:800;padding:3px 11px;border-radius:8px;margin-bottom:12px;}
.nh-ti{font-size:22px;font-weight:800;color:var(--text);line-height:1.35;margin-bottom:9px;}
.nh-sm{font-size:14px;color:var(--muted);line-height:1.65;margin-bottom:12px;}
.nh-meta{font-size:11px;color:var(--muted2);font-family:'IBM Plex Mono',monospace;}
.ai-brief{background:linear-gradient(135deg,#EAF2FC,#ffffff 60%);border:1px solid #D6E6F9;border-left:4px solid var(--navy2);border-radius:12px;padding:14px 18px;margin-bottom:14px;font-size:14px;line-height:1.6;color:#0E2A55;}

/* ============================================================
   ★ 라이트 배경 레이어
   ============================================================ */
html{background:#F5F7FA;}
body{background:transparent!important;}
body[data-cat="minerals"]{--catglow:#155BB8;--catglow2:#155BB8;}

/* 밝은 페이지 배경 (은은한 네이비 그라데이션) */
#cosmos{position:fixed;inset:0;z-index:-2;background:
  radial-gradient(900px 480px at 82% -8%, rgba(21,91,184,.08), transparent 60%),
  radial-gradient(700px 420px at 8% 4%, rgba(21,91,184,.07), transparent 55%),
  linear-gradient(180deg,#F8FAFD 0%,#F2F4F7 100%);}

/* 화이트 카드 패널 */
.stat-card,.chart-box,.section,.sub-box,.uc,.sidebar,.risk-card{
  background:#fff!important;
  border:1px solid var(--border)!important;box-shadow:var(--shadow)!important;}
.section:hover,.stat-card:hover{border-color:var(--border2)!important;}
.cat-bar{background:rgba(255,255,255,.92)!important;border-bottom:1px solid var(--border)!important;backdrop-filter:blur(10px);}
.nav{background:#fff!important;border-right:1px solid var(--border)!important;}

/* 홈으로 */
.to-space{margin-left:auto;display:inline-flex;align-items:center;gap:6px;color:var(--muted)!important;
  text-decoration:none;font-size:12px;font-weight:700;border:1px solid var(--border2);
  padding:6px 14px;border-radius:20px;background:#fff;transition:.2s;}
.to-space:hover{color:var(--navy)!important;border-color:var(--gold);}

/* ===== 메가메뉴 GNB ===== */
.nav{display:none!important;}
body{padding-left:0!important;}
.cat-bar{position:relative;z-index:600;overflow:visible!important;gap:0!important;padding:12px 30px!important;align-items:center;}
.brand-lock{display:flex;align-items:center;text-decoration:none;margin-right:26px;flex-shrink:0;transition:.2s;}
.brand-lock:hover{opacity:.82;}
.brand-logo{height:34px;width:auto;display:block;}
.brand-txt{font-size:20px;font-weight:900;letter-spacing:-.02em;color:var(--navy);font-family:var(--sans);}
.brand-txt em{font-style:normal;color:var(--gold);}
.brand-txt small{display:block;font-size:9px;font-weight:700;letter-spacing:.28em;color:var(--muted2);margin-top:1px;}
.cat-btn{background:transparent!important;border:0!important;box-shadow:none!important;border-radius:0!important;
  margin:0 20px!important;padding:6px 2px!important;color:var(--muted)!important;font-family:var(--sans)!important;
  font-size:15px!important;font-weight:700!important;letter-spacing:-.01em;position:relative;}
.cat-btn:first-of-type{margin-left:24px!important;}
.cat-btn::after{content:'';position:absolute;left:0;right:0;bottom:-5px;height:2px;border-radius:2px;
  background:linear-gradient(90deg,#B9CCEA,#155BB8);transform:scaleX(0);transform-origin:center;
  transition:.24s cubic-bezier(.2,.7,.3,1);}
.cat-btn:hover{color:var(--navy)!important;}
.cat-btn.active{background:transparent!important;color:var(--navy)!important;box-shadow:none!important;}
.cat-btn.active::after,.cat-btn:hover::after{transform:scaleX(1);}
.cat-btn .cv{display:none;}
.cb-right{margin-left:auto;display:flex;align-items:center;gap:8px;}
.cb-right .to-space{margin-left:0;}
.cb-link{font-size:12px;font-weight:700;color:var(--muted)!important;text-decoration:none;padding:6px 13px;
  border-radius:20px;border:1px solid var(--border2);transition:.18s;white-space:nowrap;}
.cb-link:hover{color:var(--navy)!important;border-color:var(--gold);background:var(--gold-soft);}
.cat-menu{position:static;display:inline-flex;align-items:center;}
.megapanel{position:absolute;left:0;right:0;top:100%;z-index:590;
  background:#fff;
  border-top:1px solid var(--border);box-shadow:0 28px 60px rgba(20,35,60,.16);
  opacity:0;visibility:hidden;transform:translateY(-10px);transition:.22s cubic-bezier(.2,.7,.3,1);}
.cat-menu:hover .megapanel{opacity:1;visibility:visible;transform:none;}
.cat-bar.mega-closed .megapanel{opacity:0!important;visibility:hidden!important;transform:translateY(-10px)!important;}
.mp-grid{display:grid;gap:13px;padding:24px 40px 28px;}
.mp-c2{grid-template-columns:repeat(2,1fr);}
.mp-c3{grid-template-columns:repeat(3,1fr);}
.mp-c4{grid-template-columns:repeat(4,1fr);}
.mp-tile{position:relative;display:flex;flex-direction:column;gap:6px;padding:15px 20px;border-radius:14px;
  background:var(--bg);border:1px solid var(--border);cursor:pointer;text-decoration:none;
  transition:.18s cubic-bezier(.2,.7,.3,1);}
.mp-tile b{font-size:14px;font-weight:700;color:var(--navy);letter-spacing:-.01em;}
.mp-tile span{font-size:11px;color:var(--muted2);}
.mp-tile::after{content:'→';position:absolute;right:14px;top:50%;transform:translateY(-50%) translateX(-4px);
  color:var(--gold);font-size:14px;opacity:0;transition:.18s;}
.mp-tile:hover{background:var(--gold-soft);border-color:var(--gold);transform:translateY(-3px);
  box-shadow:0 12px 26px rgba(20,35,60,.12);}
.mp-tile:hover b{color:var(--navy);}
.mp-tile:hover::after{opacity:1;transform:translateY(-50%) translateX(0);}

/* ===== 히어로 배너 + 통합검색 ===== */
.hero{position:relative;flex-shrink:0;height:440px;margin-bottom:0;}
.hero-clip{position:absolute;inset:0;overflow:hidden;border-bottom:1px solid var(--border);}
.hero-track{display:flex;width:200%;height:100%;transition:transform .6s cubic-bezier(.4,0,.2,1);}
.hero-slide{width:50%;height:100%;display:flex;align-items:center;padding:0 7vw;text-decoration:none;cursor:pointer;}
.hs-in{max-width:640px;}
.hs-eyebrow{font-size:11px;letter-spacing:.4em;text-transform:uppercase;color:#f0d68a;font-weight:700;margin-bottom:13px;}
.hs-title{font-family:'Noto Serif KR',serif;font-size:clamp(23px,3.2vw,38px);font-weight:700;color:#fff;line-height:1.16;}
.hs-title b{color:#f4e3ad;}
.hs-sub{margin-top:11px;font-size:14px;color:rgba(255,255,255,.75);}
.hero-slide{background-size:cover!important;background-position:center!important;background-repeat:no-repeat!important;}
.hs-min{background-image:linear-gradient(90deg,rgba(14,36,68,.92) 0%,rgba(14,36,68,.72) 40%,rgba(14,36,68,.3) 72%,rgba(14,36,68,.12) 100%),url('/static/hero/minerals.png');}
.hs-ai{background-image:linear-gradient(90deg,rgba(14,36,68,.92) 0%,rgba(14,36,68,.72) 40%,rgba(14,36,68,.3) 72%,rgba(14,36,68,.12) 100%),url('/static/hero/ai.png');}
.hero-nav{position:absolute;top:44%;transform:translateY(-50%);z-index:5;width:40px;height:40px;border-radius:50%;
  background:rgba(255,255,255,.22);border:1px solid rgba(255,255,255,.4);color:#fff;font-size:22px;line-height:1;cursor:pointer;transition:.18s;}
.hero-nav:hover{background:rgba(21,91,184,.6);border-color:#155BB8;}
.hero-nav.prev{left:18px;}.hero-nav.next{right:18px;}
.hero-dots{position:absolute;bottom:98px;left:50%;transform:translateX(-50%);z-index:5;display:flex;gap:8px;}
.hero-dots button{width:8px;height:8px;border-radius:50%;background:rgba(255,255,255,.45);border:0;cursor:pointer;padding:0;transition:.25s;}
.hero-dots button.on{background:#f0c95c;width:22px;border-radius:4px;}
.hero-search{position:absolute;left:50%;bottom:26px;transform:translateX(-50%);z-index:6;width:min(640px,82%);
  display:flex;align-items:center;gap:12px;background:#fff;border-radius:15px;padding:9px 11px 9px 20px;box-shadow:0 18px 44px rgba(14,36,68,.35);}
.hsr-cat{font-size:13px;font-weight:800;color:#16305C;white-space:nowrap;border-right:1px solid #e0e0e0;padding-right:14px;}
.hero-search input{flex:1;border:0;outline:0;font-size:14px;color:#222;background:transparent;}
.hsr-btn{width:40px;height:40px;border-radius:11px;border:0;background:linear-gradient(135deg,#B9CCEA,#155BB8);color:#fff;font-size:16px;cursor:pointer;flex-shrink:0;}

/* ===== 페이지 스크롤 ===== */
body{height:auto!important;min-height:100vh;overflow-y:auto!important;overflow-x:hidden;}
.cat-bar{position:sticky;top:0;}
.ticker{position:sticky;top:62px;z-index:580;}
#cat-minerals{flex:none!important;}
.tab-panel{overflow:visible!important;}
#tab-supply{height:auto!important;overflow:visible!important;}
.dash{height:auto!important;overflow:visible!important;}
.dash-cols{min-height:66vh;}
.main{overflow:visible!important;height:auto!important;}
#cat-minerals .tab-panel:not(#tab-supply){flex-direction:column!important;align-items:stretch!important;padding:18px 24px!important;gap:0;}
#cat-minerals .tab-panel:not(#tab-supply) .risk-grid{margin-bottom:16px;}

/* ===== 메인(홈) vs 카테고리 화면 분리 ===== */
.hero{display:none;}
body.is-home .hero{display:block;}
body.is-home #cat-minerals{display:none!important;}
body.is-home .ticker{display:none!important;}
#home-landing{display:none;}
body.is-home #home-landing{display:block;padding:44px 6vw 64px;}
.hl-wrap{max-width:1120px;margin:0 auto;}
.hl-head{text-align:center;margin-bottom:36px;}
.hl-eyebrow{font-size:11px;letter-spacing:.42em;color:var(--gold);font-weight:700;text-transform:uppercase;}
.hl-title{font-family:'Noto Serif KR',serif;font-size:clamp(28px,4vw,46px);font-weight:700;color:var(--navy);margin-top:14px;letter-spacing:-.01em;}
.hl-sub{color:var(--muted);margin-top:13px;font-size:15px;}
.hl-cards{display:grid;grid-template-columns:repeat(3,1fr);margin:0 0 32px;gap:16px;}
@media(max-width:900px){.hl-cards{grid-template-columns:1fr 1fr;}}
.hl-card{position:relative;overflow:hidden;display:block;text-decoration:none;padding:32px 28px;border-radius:20px;
  background:#fff;border:1px solid var(--border);box-shadow:var(--shadow);transition:.26s cubic-bezier(.2,.7,.3,1);}
.hl-card::after{content:'';position:absolute;inset:0;background:radial-gradient(125% 130% at 82% 0%,var(--g),transparent 60%);opacity:.7;pointer-events:none;}
.hl-min{--g:rgba(21,91,184,.14);}
.hl-ic{font-size:42px;line-height:1;position:relative;}
.hl-nm{font-size:23px;font-weight:800;color:var(--navy);margin-top:16px;position:relative;}
.hl-dc{font-size:13px;color:var(--muted);margin-top:7px;position:relative;}
.hl-go{margin-top:20px;font-size:13px;font-weight:700;color:var(--accent);position:relative;}
.hl-card:hover{transform:translateY(-6px);border-color:var(--gold);box-shadow:0 22px 50px rgba(20,35,60,.14);}
.hl-risk{margin:0 0 28px;background:#fff;border:1px solid var(--border);border-radius:18px;padding:18px 22px;box-shadow:var(--shadow);}
.hl-risk-head{font-size:14px;font-weight:800;color:var(--text);margin-bottom:13px;}
.hl-risk-head span{color:var(--muted2);font-weight:500;font-size:11.5px;margin-left:7px;}
.hl-risk-row{display:grid;grid-template-columns:repeat(4,1fr);gap:12px;}
.hl-rk{display:flex;flex-direction:column;gap:5px;text-decoration:none;background:var(--bg);border:1px solid var(--border);border-radius:13px;padding:13px 15px;transition:.18s;}
.hl-rk:hover{border-color:var(--gold);transform:translateY(-2px);}
.hl-rk-cat{font-size:11px;color:var(--muted2);font-weight:700;}
.hl-rk b{font-size:17px;letter-spacing:-.01em;}
.hl-rk-gr{font-size:11px;font-weight:700;}
.hl-rk.cta{justify-content:center;align-items:center;text-align:center;font-weight:800;color:var(--navy);font-size:13px;border-color:rgba(21,91,184,.45);background:var(--gold-soft);}
@media(max-width:900px){.hl-risk-row{grid-template-columns:1fr 1fr;}}
.hl-stats{display:flex;justify-content:center;gap:46px;flex-wrap:wrap;padding:24px;border-top:1px solid var(--border);}
.hl-stat{text-align:center;}
.hl-stat span{display:block;font-size:11px;color:var(--muted2);letter-spacing:.1em;margin-bottom:7px;}
.hl-stat b{font-size:23px;font-weight:800;color:var(--navy);font-family:var(--mono);}

/* 진입 연출 */
#arrival{position:fixed;inset:0;z-index:9999;pointer-events:none;background:#F8FAFD;
  display:flex;flex-direction:column;align-items:center;justify-content:center;text-align:center;
  opacity:1;transition:opacity 1.1s ease;}
#arrival.gone{opacity:0;}
#arrival .a-eyebrow{font-family:'Cormorant Garamond',serif;font-style:italic;font-size:clamp(15px,2vw,22px);
  letter-spacing:.36em;color:#16305C;text-transform:uppercase;opacity:0;transform:translateY(14px);transition:opacity .9s .15s,transform .9s .15s;}
#arrival .a-line{width:0;height:1px;background:linear-gradient(90deg,transparent,#155BB8,transparent);margin:24px 0;transition:width 1.1s .35s;}
#arrival .a-title{font-family:'Noto Serif KR',serif;font-weight:700;font-size:clamp(40px,7.5vw,94px);line-height:1.05;color:var(--navy);
  opacity:0;transform:translateY(20px);transition:opacity 1s .3s,transform 1s .3s;}
#arrival .a-title .gold{background:linear-gradient(118deg,#5C8FD6,#1E74D8 58%,#16305C);-webkit-background-clip:text;background-clip:text;color:transparent;}
#arrival .a-name{margin-top:20px;font-size:12px;letter-spacing:.42em;color:#888888;text-transform:uppercase;opacity:0;transition:opacity .9s .5s;}
#arrival.show .a-eyebrow,#arrival.show .a-title,#arrival.show .a-name{opacity:1;transform:none;}
#arrival.show .a-line{width:200px;}
@keyframes landIn{from{opacity:0;transform:translateY(26px) scale(.992)}to{opacity:1;transform:none}}
body.landed #cat-minerals{animation:landIn .75s cubic-bezier(.2,.7,.3,1) both;}

/* 카드 입체감 (라이트 · 절제된 섀도) */
.main{perspective:1600px;}
.stat-card,.risk-card{position:relative;overflow:hidden;border-radius:16px!important;
  transform-style:preserve-3d;will-change:transform;transition:transform .18s cubic-bezier(.2,.7,.3,1),box-shadow .22s,border-color .22s;}
.stat-card:hover,.risk-card:hover{
  box-shadow:0 18px 44px rgba(20,35,60,.14)!important;
  border-color:var(--gold)!important;}

/* ====== 씬 모드 (비활성 유지) ====== */
body.scenes{padding-left:0!important;}
body.scenes .nav{display:none!important;}
body.scenes #cat-minerals{display:block!important;overflow-y:auto;scroll-snap-type:y proximity;scroll-behavior:smooth;scrollbar-width:none;}
body.scenes #cat-minerals::-webkit-scrollbar{display:none;}
body.scenes #cat-minerals .tab-panel{display:flex!important;flex-direction:column;justify-content:flex-start;align-items:stretch;
  scroll-snap-align:start;padding:5vh 5vw 7vh 96px;box-sizing:border-box;animation:sceneIn .7s cubic-bezier(.2,.7,.3,1) both;}
body.scenes #cat-minerals .main{overflow:visible!important;height:auto!important;flex:none!important;padding:0!important;}
body.scenes #cat-minerals .sidebar{height:auto!important;max-width:280px;}
@keyframes sceneIn{from{opacity:0;transform:translateY(40px)}to{opacity:1;transform:none}}
.scene-head{margin-bottom:24px;}
.scene-hero{font-size:13px;letter-spacing:.38em;text-transform:uppercase;color:var(--gold);font-weight:700;margin-bottom:10px;}
.scene-title{font-size:clamp(28px,5vw,58px);font-weight:800;letter-spacing:-.03em;line-height:1.05;color:var(--navy);}
#scene-rail{position:fixed;left:22px;top:50%;transform:translateY(-50%);z-index:350;display:none;flex-direction:column;gap:14px;}
body.scenes #scene-rail{display:flex;}
#scene-rail button{all:unset;cursor:pointer;width:11px;height:11px;border-radius:50%;background:var(--border2);transition:.25s;position:relative;}
#scene-rail button.on{background:var(--gold);transform:scale(1.4);}
#scene-rail button span{position:absolute;left:22px;top:50%;transform:translateY(-50%);white-space:nowrap;font-size:12px;color:var(--muted);opacity:0;pointer-events:none;transition:.2s;background:#fff;padding:4px 11px;border-radius:7px;border:1px solid var(--border);}
#scene-rail button:hover span{opacity:1;}

/* ============================================================
   ★ 대시보드 밀도·가독성 (KPI + 벤토)
   ============================================================ */
.stat-row{gap:14px!important;margin-bottom:18px!important}
.stat-card{padding:18px 22px!important;border-radius:16px!important;position:relative;overflow:hidden;min-height:98px;
  display:flex!important;flex-direction:column;justify-content:center;
  background:#fff!important;border:1px solid var(--border)!important}
.stat-card:first-child{flex:1.7}
.stat-card::after{content:'';position:absolute;left:0;top:0;bottom:0;width:3px;background:var(--gold);opacity:.9}
.stat-card.red::after{background:var(--red)}
.stat-card .sc-label{font-size:10px!important;letter-spacing:.18em;text-transform:uppercase;color:var(--muted2)!important;margin-bottom:9px;font-family:var(--mono)}
.stat-card .sc-val{font-size:clamp(19px,2.1vw,28px)!important;font-weight:800!important;font-family:var(--mono);letter-spacing:-.02em;line-height:1.1;white-space:nowrap;color:var(--navy);text-shadow:none!important}
.stat-card:first-child{background:linear-gradient(135deg,#16305C,#155BB8)!important;border-color:#16305C!important}
.stat-card:first-child::after{background:#B9CCEA}
.stat-card:first-child .sc-label{color:#b9cbe4!important}
.stat-card:first-child .sc-val{font-size:clamp(20px,2.7vw,36px)!important;
  color:#f4e3ad!important;-webkit-text-fill-color:#f4e3ad!important;background:none!important}
.stat-card:first-child .sc-sub{color:#9db4d4!important}
.stat-card.red .sc-val{color:var(--red)!important}
.stat-card .sc-sub{font-size:11px!important;color:var(--muted2)!important;margin-top:8px}
.stat-card:hover{border-color:var(--gold)!important;box-shadow:0 16px 40px rgba(20,35,60,.13)}
.page-title{color:var(--navy)!important;font-size:13px!important}
.chart-title,.sec-head{letter-spacing:.03em}
.charts-row{height:248px!important;gap:14px!important}
.chart-box{border-radius:14px!important;padding:16px!important}
.sb-section{margin-bottom:18px}
.sb-title{color:var(--accent)!important;opacity:.95}
.sb-stat{padding:6px 0!important}
.sb-stat:hover{background:var(--gold-soft)}

/* ====== 관제실형 대시보드 격자 (수급현황) ====== */
#tab-supply{flex-direction:column!important;overflow:hidden!important;}
.dash{display:flex;flex-direction:column;gap:14px;height:100%;padding:16px;box-sizing:border-box;overflow:hidden}
.dash-kpis{flex-shrink:0;margin-bottom:0!important}
.dash-cols{flex:1;min-height:0;display:grid;grid-template-columns:0.95fr 1.45fr 1.05fr;gap:14px}
.dash-col{min-height:0;display:flex;flex-direction:column;gap:14px}
.wpanel{background:#fff;border:1px solid var(--border);border-radius:16px;
  display:flex;flex-direction:column;min-height:0;overflow:hidden;box-shadow:var(--shadow)}
.wpanel.grow{flex:1}
.wp-head{flex-shrink:0;display:flex;align-items:center;gap:8px;font-size:11.5px;font-weight:800;letter-spacing:.04em;
  color:var(--navy);padding:13px 16px;border-bottom:1px solid var(--border)}
.wp-sub{font-size:9px;font-weight:600;letter-spacing:.12em;text-transform:uppercase;color:var(--muted2);
  background:var(--bg3);padding:2px 7px;border-radius:6px;margin-left:auto}
.wp-body{flex:1;min-height:0;overflow-y:auto;padding:8px 16px 12px}
.wp-chart{flex:1;min-height:130px;position:relative;padding:12px 14px}
.wp-table{width:100%;border-collapse:collapse;font-size:12px}
.wp-table td{padding:7px 4px;border-bottom:1px solid var(--border)}
.rk-item{display:flex;align-items:center;gap:12px;padding:9px 0;border-bottom:1px solid var(--border)}
.rk-item:last-child{border-bottom:0}
.rk-no{font-family:var(--mono);font-size:13px;font-weight:700;color:var(--muted2);width:22px;text-align:center}
.rk-mid{flex:1;min-width:0}
.rk-nm{font-size:13px;font-weight:600;color:var(--text);margin-bottom:6px}
.rk-bar{height:5px;background:var(--bg3);border-radius:3px;overflow:hidden}
.rk-fill{height:100%;background:linear-gradient(90deg,#155BB8,#B9CCEA);border-radius:3px}
.rk-vl{font-family:var(--mono);font-size:13px;font-weight:700;color:var(--accent);white-space:nowrap}
.rk-vl small{font-size:9px;color:var(--muted2);margin-left:2px;font-weight:500}
.pr-item{display:flex;justify-content:space-between;align-items:center;padding:8px 0;border-bottom:1px solid var(--border);font-size:12.5px}
/* 오늘의 금속 시세 */
.pp-row{display:flex;align-items:center;gap:8px;padding:7px 0;border-bottom:1px solid var(--border);font-size:12.5px}
.pp-row:last-child{border-bottom:0}
.pp-nm{flex:1;color:var(--text);font-weight:600}
.pp-val{font-family:var(--mono);font-weight:700;color:var(--navy)}
.pp-chg{width:74px;text-align:right;font-family:var(--mono);font-size:11.5px;font-weight:700}
/* 커버리지 분류 스트립 */
.tx-row{display:flex;align-items:flex-start;gap:12px;padding:8px 0;border-bottom:1px solid var(--border)}
.tx-row:last-child{border-bottom:0}
.tx-cat{flex-shrink:0;width:76px;text-align:center;font-size:11.5px;font-weight:800;border:1px solid;border-radius:8px;padding:4px 0;margin-top:1px}
.tx-list{flex:1;line-height:2}
.tx-list i{font-style:normal;display:inline-block;font-size:12px;color:var(--muted);background:var(--bg);border:1px solid var(--border);border-radius:12px;padding:2px 10px;margin:0 4px 4px 0}
.tx-cnt{flex-shrink:0;font-family:var(--mono);font-size:11px;font-weight:700;color:var(--muted2);margin-top:5px}
.rk-cat{font-style:normal;font-size:9.5px;font-weight:700;margin-left:5px}
/* 분류별 카테고리 페이지 */
.catpage{padding:18px 24px;}
body.is-home .catpage{display:none!important;}
.catpage .stat-row{margin-bottom:14px}
/* 홈 금속 시세 스트립 */
.hl-metal{margin:0 0 28px;background:#fff;border:1px solid var(--border);border-radius:18px;padding:18px 22px;box-shadow:var(--shadow)}
.hm-row{display:grid;grid-template-columns:repeat(6,1fr);gap:10px}
.hm-card{display:flex;flex-direction:column;gap:4px;background:var(--bg);border:1px solid var(--border);border-radius:12px;padding:11px 14px}
.hm-nm{font-size:11px;font-weight:700;color:var(--muted2)}
.hm-val{font-size:17px;font-weight:800;color:var(--navy);font-family:var(--mono);letter-spacing:-.01em}
.hm-chg{font-size:11px;font-weight:700;font-family:var(--mono)}
@media(max-width:900px){.hm-row{grid-template-columns:repeat(3,1fr)}}
.pr-item:last-child{border-bottom:0}
.pr-nm{color:var(--muted)}.pr-co{color:var(--text);font-weight:600}
@media(max-width:1100px){.dash-cols{grid-template-columns:1fr 1fr}.dash-col:last-child{grid-column:1/-1}}
/* ═══════════════════════════════════════════════
   ★ MINETECH 시그니처 패스 — 공공 대시보드 톤 탈피
   ═══════════════════════════════════════════════ */
/* 등고선 텍스처 배경 (광산 지형 모티프) */
#cosmos::after{content:'';position:absolute;inset:0;opacity:.5;
  background-image:url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='560' height='560' viewBox='0 0 560 560'%3E%3Cg fill='none' stroke='%2312325e' stroke-opacity='.045' stroke-width='1.2'%3E%3Cpath d='M60 280c40-90 150-130 230-90s110 150 60 210-190 60-250-10-60-70-40-110z'/%3E%3Cpath d='M100 280c30-70 120-100 185-70s90 120 50 168-155 48-203-8-48-56-32-90z'/%3E%3Cpath d='M140 282c22-50 88-72 135-50s66 88 36 123-113 35-148-6-35-42-23-67z'/%3E%3Cpath d='M180 284c15-33 58-47 89-33s44 58 24 81-75 23-98-4-23-27-15-44z'/%3E%3Cpath d='M420 90c50 10 80 60 60 100s-90 50-120 10 10-118 60-110z'/%3E%3Cpath d='M430 118c32 7 51 39 38 65s-58 32-77 6 7-77 39-71z'/%3E%3C/g%3E%3C/svg%3E");}

/* 숫자 디스플레이 폰트 */
html body .sc-val, html body .rk-vl, html body .hm-val, html body .pp-val,
html body .hl-stat b, html body .rk-val, html body .rk-no{
  font-family:'Archivo','Pretendard',sans-serif!important;letter-spacing:-.01em;}

/* GNB — 골드 헤어라인 + 여백 */
.cat-bar{padding:15px 34px!important;border-bottom:none!important;
  box-shadow:inset 0 -1px 0 rgba(22,48,92,.08),0 8px 28px rgba(22,48,92,.05);}
.cat-bar::before{content:'';position:absolute;top:0;left:0;right:0;height:3px;
  background:linear-gradient(90deg,#155BB8,#B9CCEA 30%,#16305C 75%);}
.brand-txt{font-size:22px;font-family:'Archivo','Pretendard',sans-serif;}
.cat-btn{font-size:15.5px!important;margin:0 22px!important;}
.cat-btn::after{height:3px;bottom:-7px;}
.ticker{background:linear-gradient(90deg,#0A1F3E,#16386b 55%,#0A1F3E)!important;border-bottom:none!important;}
.ticker-inner{color:#B9CCEA!important;}

/* KPI — 네이비 카드 밴드 (전 카드) */
.stat-card{background:linear-gradient(160deg,#123C7E,#0A1F3E 90%)!important;
  border:1px solid rgba(255,255,255,.07)!important;
  box-shadow:0 14px 34px rgba(14,38,71,.22)!important;}
.stat-card .sc-label{color:#8fb8a4!important;}
.stat-card .sc-val{color:#fff!important;}
.stat-card .sc-sub{color:#7e95b5!important;}
.stat-card::after{background:linear-gradient(180deg,#B9CCEA,#155BB8);opacity:1;width:4px;}
.stat-card:first-child{background:linear-gradient(150deg,#155BB8,#16305C 70%)!important;}
.stat-card:first-child .sc-val{color:#f0c95c!important;-webkit-text-fill-color:#f0c95c!important;}
.stat-card:hover{transform:translateY(-3px);border-color:rgba(233,198,103,.45)!important;
  box-shadow:0 22px 46px rgba(14,38,71,.32)!important;}

/* 카드 — 테두리 제거 · 레이어드 섀도 · 큰 라운드 */
.section,.chart-box,.wpanel,.uc,.sub-box,.risk-card,.hl-card,.hl-risk,.hl-metal,.dcard{
  border:none!important;border-radius:18px!important;
  box-shadow:0 1px 2px rgba(20,35,60,.04),0 12px 36px rgba(20,35,60,.08)!important;}
.wpanel .wp-head{border-bottom:1px dashed #E8ECF1;padding:15px 18px;font-size:12.5px;}
.wp-head::before{content:'';display:inline-block;width:7px;height:7px;flex-shrink:0;
  background:linear-gradient(135deg,#B9CCEA,#155BB8);border-radius:2px;transform:rotate(45deg);}
.page-title{display:flex;align-items:center;gap:9px;font-size:14px!important;}
.page-title::before{content:'';display:inline-block;width:8px;height:8px;flex-shrink:0;
  background:linear-gradient(135deg,#B9CCEA,#155BB8);border-radius:2px;transform:rotate(45deg);}
.chart-title{color:var(--navy);font-weight:800;}

/* 순위 바 — 네이비→골드 시그니처 그라데이션 */
.rk-bar{height:7px;border-radius:4px;}
.rk-fill{background:linear-gradient(90deg,#155BB8 0%,#3a72b8 55%,#155BB8 100%)!important;}
.rk-no{color:#155BB8;font-weight:800;font-size:12px;}
.bf{background:linear-gradient(90deg,#155BB8,#155BB8)!important;}

/* 버튼·필 */
.mineral-btn{border-radius:10px;font-weight:600;}
.mineral-btn.active{box-shadow:0 6px 16px rgba(22,48,92,.28)!important;}
.megapanel{border-radius:0 0 22px 22px;overflow:hidden;}

/* 테이블 line 완화 */
.wp-table td{border-bottom:1px solid #eef1f6;}
tr:hover td{background:#F8FAFD;}

/* 홈 — 통계 밴드 */
.hl-stats{background:linear-gradient(160deg,#123C7E,#0A1F3E);border-radius:20px;border-top:none;
  box-shadow:0 16px 40px rgba(14,38,71,.25);padding:28px 24px;}
.hl-stat span{color:#8fb8a4;}
.hl-stat b{color:#f0c95c;font-size:26px;}
.hm-card{border:none;box-shadow:0 1px 2px rgba(20,35,60,.05),0 8px 20px rgba(20,35,60,.07);}

"""
    CAT_JS = r"""
var CATS=['all','minerals','nf','rare','ree','energy','etc'];
// 정부 지정 핵심광물(핵심광물 확보전략 33종 중 서비스 보유 32종 + 표기 별칭)
var CORE_NMS=['리튬','니켈','코발트','망간','흑연','네오디뮴','디스프로슘','터븀','세륨','란탄','동(구리)','동','알루미늄','아연','연(납)','연','주석','티타늄','텅스텐','몰리브덴','크롬','마그네슘','안티모니','바나듐','니오븀','탄탈륨','지르코늄','갈륨','인듐','게르마늄','규소','셀레늄','창연/비스무트','비스무트','백금'];
function switchCategory(cat, el){
  if(CATS.indexOf(cat)<0) cat='minerals';
  if(cat==='all' && window._curTab!=='risk') cat='minerals';
  if(document.body.classList.contains('is-home')){ location.href='/dashboard?cat='+cat; return; }  // 메인에선 카테고리 화면으로 이동
  // 리스크 신호등: 카테고리 클릭 = K-RISK 카드 분류 필터 (페이지 전환 없음)
  if(window._curTab==='risk'){
    document.querySelectorAll('.cat-btn').forEach(function(b){b.classList.remove('active');});
    var _rb=el||document.querySelector('.cat-btn[data-cat="'+cat+'"]'); if(_rb)_rb.classList.add('active');
    var RC={'니켈':'nf','동':'nf','리튬':'rare','코발트':'rare','텅스텐':'rare','몰리브덴':'rare'};
    var grid=document.querySelector('#tab-risk .risk-grid');
    var note=document.getElementById('riskCatNote');
    if(!note && grid){ note=document.createElement('div'); note.id='riskCatNote';
      note.style.cssText='padding:12px 4px;color:#888;font-size:13.5px'; grid.parentNode.insertBefore(note,grid); }
    if(cat==='minerals'){
      document.querySelectorAll('#tab-risk .risk-card').forEach(function(c){
        var nm=((c.querySelector('.rk-nm')||{}).textContent||'').trim();
        c.style.display=(CORE_NMS.indexOf(nm)>=0)?'':'none';
      });
      if(note) note.textContent='정부 지정 핵심광물(핵심광물 확보전략 33종 기준)만 표시 중 — 전체를 보려면 [전체] 탭';
      return;
    }
    if(cat!=='all'){
      var any=false;
      document.querySelectorAll('#tab-risk .risk-card').forEach(function(c){
        var nm=((c.querySelector('.rk-nm')||{}).textContent||'').trim();
        var show=(c.dataset.cat||RC[nm])===cat; c.style.display=show?'':'none'; if(show)any=true;
      });
      if(note) note.textContent=any?'':'이 분류에는 해당 광종이 없습니다.';
      return;
    }
    document.querySelectorAll('#tab-risk .risk-card').forEach(function(c){c.style.display='';});
    if(note) note.textContent='';
    return;
  }
  // 국내 광산·뉴스: 전국/전체 단위 데이터 — 카테고리와 무관, 페이지 유지
  if((window._curTab==='mines'||window._curTab==='news') && cat!=='minerals'){
    document.querySelectorAll('.cat-btn').forEach(function(b){b.classList.remove('active');});
    var _mb=el||document.querySelector('.cat-btn[data-cat="'+cat+'"]'); if(_mb)_mb.classList.add('active');
    return;
  }
  // 가격지수 탭: 카테고리 클릭 = 그 분류의 지수 라인 필터 (페이지 전환 없음)
  if(window._curTab==='mindex' && typeof _mindexChart!=='undefined'){
    document.querySelectorAll('.cat-btn').forEach(function(b){b.classList.remove('active');});
    var _b=el||document.querySelector('.cat-btn[data-cat="'+cat+'"]'); if(_b)_b.classList.add('active');
    if(!_mindexChart && typeof drawMineralIndex==='function') drawMineralIndex();
    if(_mindexChart){
      var _g={nf:'메이저금속',rare:'희소금속',ree:'희소금속',energy:'에너지광물',etc:'종합'}[cat];
      _mindexChart.data.datasets.forEach(function(d,i){
        _mindexChart.setDatasetVisibility(i, (cat==='minerals') ? true : d.label===_g);
      });
      _mindexChart.update();
    }
    if(cat!=='minerals') return;   // 핵심광물 클릭 시엔 아래 일반 로직으로 계속(전체 복원)
  }
  var mn=document.getElementById('cat-minerals'); if(mn) mn.style.display=(cat==='minerals')?'flex':'none';
  CATS.slice(1).forEach(function(c){
    var d=document.getElementById('cat-'+c); if(d) d.style.display=(c===cat)?'block':'none';
  });
  document.querySelectorAll('.cat-btn').forEach(function(b){b.classList.remove('active');});
  if(el) el.classList.add('active');
  else { var b2=document.querySelector('.cat-btn[data-cat="'+cat+'"]'); if(b2) b2.classList.add('active'); }
  if(cat==='minerals' && typeof initMap==='function' && !window._mapInited){
    var mp=document.querySelector('#tab-map.active'); if(mp) initMap();
  }
  if(cat!=='minerals' && typeof initCatForecast==='function') initCatForecast(cat);
  document.body.dataset.cat = 'minerals';
  if(window._applyScenes) window._applyScenes('minerals');
  // 탭 포커스 모드 초기화
  document.querySelectorAll('.catpage.tabfocus').forEach(function(p){ p.classList.remove('tabfocus'); });
  document.querySelectorAll('.row-focus').forEach(function(r){ r.classList.remove('row-focus'); });
  document.querySelectorAll('.sec-focus').forEach(function(x){ x.classList.remove('sec-focus'); });
  // 통계 탭 컨텍스트 유지: 카테고리 진입 시 그 기능의 섹션만 포커스 표시
  if(cat!=='minerals' && window._curTab){
    var anchor = document.getElementById('sec-'+cat+'-'+window._curTab);
    if(anchor){
      var tgt = (anchor.classList.contains('section')) ? anchor : anchor.nextElementSibling;
      var row = tgt ? tgt.closest('.charts-row') : null;
      var page = document.getElementById('cat-'+cat);
      if(tgt && row && page){
        page.classList.add('tabfocus');
        row.classList.add('row-focus');
        tgt.classList.add('sec-focus');
        window.scrollTo({top:0});
      }
    }
  }
  if(cat==='minerals' && window._curTab){
    var tb = document.querySelector('.nav a[data-tab="'+window._curTab+'"]');
    if(tb) setTimeout(function(){ switchTab(window._curTab, tb); }, 60);
  }
}
// ── 분류 페이지 가격전망 미니차트 ──
var _catFc={};
// ── 광물별 월간 수입 동향 (수급 현황) ──
var _coreChart=null;
function buildCoreTrade(){
  var box=document.getElementById('coreBtns'); if(!box || !window.CORETRADE) return;
  var keys=Object.keys(CORETRADE); if(!keys.length) return;
  keys.forEach(function(m,i){
    var b=document.createElement('button');
    b.className='mineral-btn core-btn'+(i===0?' active':''); b.textContent=m;
    b.style.fontSize='11px'; b.style.padding='3px 11px';
    b.onclick=function(){ _setActive('.core-btn', b); drawCoreTrade(m); };
    box.appendChild(b);
  });
  drawCoreTrade(keys[0]);
}
function drawCoreTrade(m){
  var d=CORETRADE[m]; if(!d) return;
  var months=Object.keys(d.monthly);
  var note=document.getElementById('coreNote');
  if(note) note.innerHTML='최대 수입국 <b style="color:var(--navy)">'+d.top[0]+' '+d.top[1]+'%</b> · 12개월 수입 <b style="color:var(--navy)">$'+(d.total/1e6).toFixed(0)+'M</b>';
  if(_coreChart) _coreChart.destroy();
  _coreChart=new Chart(document.getElementById('coreChart'),{type:'line',
    data:{labels:months,datasets:[{label:m+' 월별 수입($)',data:months.map(function(k){return d.monthly[k];}),
      borderColor:'#155BB8',backgroundColor:'rgba(21,91,184,.08)',fill:true,borderWidth:2,tension:.25,pointRadius:2}]},
    options:{responsive:true,maintainAspectRatio:false,
      plugins:{legend:{display:false},tooltip:{callbacks:{label:function(c){return ' $'+c.raw.toLocaleString();}}}},
      scales:{x:{ticks:{color:'#555555',maxTicksLimit:12},grid:{color:'#E8ECF1'}},
        y:{ticks:{color:'#555555',callback:function(v){return '$'+(v/1e6).toFixed(0)+'M';}},grid:{color:'#E8ECF1'}}}}});
}
buildCoreTrade();

var _reeChart=null;
function drawReeChart(){
  if(_reeChart || !window.REETRADE) return;
  var keys=Object.keys(REETRADE); if(!keys.length) return;
  _reeChart=new Chart(document.getElementById('reeChart'),{type:'bar',
    data:{labels:keys,datasets:[{label:'수입액($)',data:keys.map(function(k){return REETRADE[k];}),backgroundColor:'#4a3aa7',borderRadius:3}]},
    options:{responsive:true,maintainAspectRatio:false,
      plugins:{legend:{display:false},tooltip:{callbacks:{label:function(c){return ' $'+c.raw.toLocaleString();}}}},
      scales:{x:{ticks:{color:'#555555',maxTicksLimit:12},grid:{color:'#E8ECF1'}},
        y:{ticks:{color:'#555555',callback:function(v){return '$'+(v/1e6).toFixed(0)+'M';}},grid:{color:'#E8ECF1'}}}}});
}
function initCatForecast(cat){
  if(cat==='ree') drawReeChart();
  if(_catFc[cat] || !window.CATFC) return;
  var keys=(CATFC[cat]||[]).filter(function(m){ return FORECAST[m]; });
  if(!keys.length){ _catFc[cat]=true; return; }
  var box=document.getElementById('fcCatBtns_'+cat); if(!box) return;
  keys.forEach(function(m,i){
    var b=document.createElement('button');
    b.className='mineral-btn'+(i===0?' active':''); b.textContent=m; b.style.fontSize='11.5px'; b.style.padding='4px 12px';
    b.onclick=function(){ box.querySelectorAll('.mineral-btn').forEach(function(x){x.classList.remove('active');}); b.classList.add('active'); drawCatForecast(cat,m); };
    box.appendChild(b);
  });
  _catFc[cat]={chart:null};
  drawCatForecast(cat, keys[0]);
}
function drawCatForecast(cat,m){
  var d=FORECAST[m]; if(!d) return;
  var t=document.getElementById('fcCatTitle_'+cat);
  if(t) t.textContent=m+' 가격 전망 ('+(d.unit||'')+') — 실선=실측 · 점선=예측';
  var actual=d.values.map(function(v,i){ return i<d.split? v: null; });
  var fut=d.values.map(function(v,i){ return i>=d.split-1? v: null; });
  if(_catFc[cat] && _catFc[cat].chart) _catFc[cat].chart.destroy();
  _catFc[cat]={chart:new Chart(document.getElementById('fcCatChart_'+cat),{type:'line',
    data:{labels:d.dates,datasets:[
      {label:'실측',data:actual,borderColor:'#155BB8',backgroundColor:'rgba(21,91,184,.08)',fill:true,borderWidth:2,tension:.25,pointRadius:0},
      {label:'예측',data:fut,borderColor:'#c98500',borderDash:[7,5],backgroundColor:'transparent',borderWidth:2,tension:.25,pointRadius:0}]},
    options:{responsive:true,maintainAspectRatio:false,interaction:{mode:'index',intersect:false},
      plugins:{legend:{labels:{color:'#555555',font:{size:10},boxWidth:10}}},
      scales:{x:{ticks:{color:'#555555',maxTicksLimit:10},grid:{color:'#E8ECF1'}},
        y:{ticks:{color:'#555555'},grid:{color:'#E8ECF1'}}}}})};
}
// 메가메뉴 → 해당 섹션 선택
function goSec(cat, sec){
  if(document.body.classList.contains('is-home')){ location.href='/dashboard?cat=minerals&sec='+sec; return; }  // 메인 → 카테고리 화면
  switchTab(sec, document.querySelector('.nav a[data-tab="'+sec+'"]'));
  // 클릭으로 화면이 바뀌면 펼친 메가패널 닫기 (바 영역을 벗어나면 다시 열림)
  var cb=document.querySelector('.cat-bar');
  if(cb){ cb.classList.add('mega-closed');
    cb.addEventListener('mouseleave', function(){ cb.classList.remove('mega-closed'); }, {once:true}); }
}
// ── 히어로 슬라이드 배너 + 통합검색 ──
var _heroI=0, _heroN=2, _heroT=null;
function heroSet(i){ _heroI=(i+_heroN)%_heroN;
  var t=document.getElementById('heroTrack'); if(t) t.style.transform='translateX(-'+(_heroI*50)+'%)';
  document.querySelectorAll('#heroDots button').forEach(function(d,k){ d.classList.toggle('on',k===_heroI); }); }
function heroGo(d){ heroSet(_heroI+d); heroRestart(); }
function heroRestart(){ clearInterval(_heroT); _heroT=setInterval(function(){ heroSet(_heroI+1); },5500); }
function heroSearch(){
  var q=(document.getElementById('heroQ')||{}).value||''; q=q.trim();
  if(!q) return;
  location.href='/search?q='+encodeURIComponent(q);
}
(function initHero(){ var dd=document.getElementById('heroDots'); if(!dd) return;
  for(var k=0;k<_heroN;k++){ var b=document.createElement('button'); if(k===0) b.className='on';
    (function(idx){ b.onclick=function(){ heroSet(idx); heroRestart(); }; })(k); dd.appendChild(b); }
  heroRestart();
})();
// 허브에서 다이브해 들어온 경우 → 카테고리 진입 + 시네마틱 "막 타이틀 카드"
var CAT_ACTS={
  minerals:['제 1 막 · Act I','땅속의 <span class="gold">권력</span>','핵심광물 · Core Minerals']
};
(function(){
  try{
    var ar=document.getElementById('arrival');
    // 메인(홈)에선 카테고리 진입·막 카드 없음
    if(document.body.classList.contains('is-home')){ if(ar) ar.style.display='none'; return; }
    var qp=new URLSearchParams(location.search);
    var sec=qp.get('sec'), min=qp.get('min');
    var cat0=qp.get('cat')||'minerals'; if(CATS.indexOf(cat0)<0) cat0='minerals';
    document.body.dataset.cat = 'minerals';
    var b=document.querySelector('.cat-btn[data-cat="'+cat0+'"]'); switchCategory(cat0,b);
    // 메인 검색/메가메뉴에서 넘어온 딥링크 — switchTab이 뒤 스크립트에 정의되므로 load 후 실행
    var _deep=function(){ try{
      if(min){ switchTab('map', document.querySelector('.nav a[data-tab="map"]')); if(window.selectMineral) selectMineral(min, null);
        if(qp.get('mode')==='routes' && window.setMode) setTimeout(function(){ setMode('routes', document.getElementById('modeRoutes')); }, 400); }
      else if(sec){
        switchTab(sec, document.querySelector('.nav a[data-tab="'+sec+'"]'));
      }
    }catch(e){} };
    if(document.readyState==='complete') _deep(); else window.addEventListener('load', _deep);
    if(ar){ ar.style.display='none'; }   // 막 타이틀 카드 제거 — 바로 대시보드
    document.body.classList.add('landed');
  }catch(e){ var ar=document.getElementById('arrival'); if(ar) ar.style.display='none'; }
})();
// 통계 숫자 카운트업 (막 카드가 걷힌 뒤 시작 → 대시보드 드러나며 촤르륵)
function countUp(){
  document.querySelectorAll('.stat-card .sc-val').forEach(function(el){
    if(el._cu) return;
    var txt=(el.textContent||'').trim();
    var m=txt.match(/^([^\d\-]*)([\d,]+(?:\.\d+)?)(.*)$/);
    if(!m){ el._cu=1; return; }
    var num=parseFloat(m[2].replace(/,/g,'')); if(!isFinite(num)||num<=0){ el._cu=1; return; }
    el._cu=1; var pre=m[1], suf=m[3], dur=1000, t0=performance.now();
    function step(now){ if(!(now>t0)){ el.textContent=pre+num.toLocaleString()+suf; return; }  // 타임스탬프 이상(헤드리스 등) 시 최종값 즉시 표시
      var k=Math.min((now-t0)/dur,1), e=1-Math.pow(1-k,3);
      el.textContent=pre+Math.round(num*e).toLocaleString()+suf;
      if(k<1) requestAnimationFrame(step); else el.textContent=pre+num.toLocaleString()+suf; }
    requestAnimationFrame(step);
  });
}
setTimeout(countUp, 2350);
function _setActive(sel, el){
  document.querySelectorAll(sel).forEach(function(b){b.classList.remove('active');});
  if(el) el.classList.add('active');
}
function _audLabel(a){ return {investor:'📈 투자자',business:'🏢 기업',consumer:'🛒 소비자',투자자:'📈 투자자',기업:'🏢 기업',소비자:'🛒 소비자'}[a] || a; }
function _newsCard(n, hero){
  var a=document.createElement('a'); a.href=n['언론사링크']||'#'; a.target='_blank';
  a.className = hero ? 'news-hero' : 'nc';
  var meta=(n['aud']?_audLabel(n['aud'])+' · ':'')+(n['검색키워드']||'');
  if(hero){
    var bd=document.createElement('span'); bd.className='nh-badge'; bd.textContent='🔥 주요 뉴스'; a.appendChild(bd);
    var ti=document.createElement('div'); ti.className='nh-ti'; ti.textContent=n['제목']||''; a.appendChild(ti);
    var sm=document.createElement('div'); sm.className='nh-sm'; sm.textContent=n['요약']||''; a.appendChild(sm);
    var mt=document.createElement('div'); mt.className='nh-meta'; mt.textContent=meta+' · '+(n['발행일시']||''); a.appendChild(mt);
  } else {
    var kw=document.createElement('span'); kw.className='nc-kw'; kw.textContent=meta;
    var t=document.createElement('div'); t.className='nc-ti'; t.textContent=n['제목']||'';
    var s=document.createElement('div'); s.className='nc-sm'; s.textContent=n['요약']||'';
    var d=document.createElement('div'); d.className='nc-dt'; d.textContent=n['발행일시']||'';
    a.appendChild(kw); a.appendChild(t); a.appendChild(s); a.appendChild(d);
  }
  return a;
}
function renderFeed(data, heroId, gridId, aud){
  var hero=document.getElementById(heroId), grid=document.getElementById(gridId);
  if(!hero||!grid||typeof data==='undefined') return;
  var list = aud ? data.filter(function(n){ return aud==='전체'||n.aud===aud; }) : data;
  hero.innerHTML=''; grid.innerHTML='';
  if(!list.length){ grid.innerHTML='<div class="empty">뉴스가 없습니다.</div>'; return; }
  hero.appendChild(_newsCard(list[0], true));
  list.slice(1).forEach(function(n){ grid.appendChild(_newsCard(n, false)); });
}
function filterNews(aud, el){ _setActive('.news-aud-btn', el); renderFeed(NEWS,'newsHero','newsGrid',aud); }
if(typeof NEWS!=='undefined') renderFeed(NEWS,'newsHero','newsGrid','전체');
// 카테고리별 뉴스 AI 브리핑 (한 번만 로드)
function fetchNewsBrief(cat, elId){
  var el=document.getElementById(elId); if(!el || el._loaded) return; el._loaded=true;
  fetch('/api/news-brief'+(cat?('?cat='+cat):'')).then(function(r){return r.json();}).then(function(d){
    if(d && d.ok && d.brief){ el.textContent='🤖 AI 브리핑 — '+String(d.brief).replace(/\*\*/g,''); el.style.display='block'; }
  }).catch(function(){});
}
fetchNewsBrief('', 'aiBrief');   // 핵심광물(기본)
var _riskChart=null;
function drawRiskChart(){
  if(_riskChart || typeof RISK==='undefined' || !RISK.length) return;
  var pal=['#155BB8','#c98500','#1baf7a','#e34948','#4a3aa7','#e87ba4'];
  var labels=RISK[0].months;
  var ds=RISK.map(function(r,i){return {label:r.name,data:r.vals,borderColor:pal[i%pal.length],backgroundColor:'transparent',tension:.25,pointRadius:0};});
  _riskChart=new Chart(document.getElementById('riskChart'),{type:'line',data:{labels:labels,datasets:ds},
    options:{responsive:true,maintainAspectRatio:false,interaction:{mode:'index',intersect:false},
      plugins:{legend:{labels:{color:'#555555',font:{size:10},boxWidth:10}}},
      scales:{x:{ticks:{color:'#555555',maxTicksLimit:12},grid:{color:'#E8ECF1'}},
        y:{min:0,max:100,ticks:{color:'#555555'},grid:{color:'#E8ECF1'}}}}});
}
var _mindexChart=null;
// ── 가격 전망 (실측 실선 + 예측 점선) ──
var _fcChart=null, _fcCur=null, _fcInit=false;
var _wbChart=null;
function buildWb(){
  var box=document.getElementById('wbBtns'); if(!box || !window.WBP || !WBP.series) return;
  var keys=Object.keys(WBP.series);
  keys.forEach(function(m,i){
    var b=document.createElement('button');
    b.className='mineral-btn wb-btn'+(i===0?' active':''); b.textContent=m;
    b.style.fontSize='11px'; b.style.padding='3px 11px';
    b.onclick=function(){ _setActive('.wb-btn', b); drawWb(m); };
    box.appendChild(b);
  });
  drawWb(keys[0]);
}
function drawWb(m){
  var vals=WBP.series[m]; if(!vals) return;
  var t=document.getElementById('wbTitle');
  if(t) t.innerHTML='국제 원자재 시세 — '+m+' ('+(WBP.units[m]||'')+') <span style="color:var(--muted2)">· World Bank Pink Sheet · 월별</span>';
  if(_wbChart) _wbChart.destroy();
  _wbChart=new Chart(document.getElementById('wbChart'),{type:'line',
    data:{labels:WBP.months,datasets:[{label:m,data:vals,borderColor:'#155BB8',backgroundColor:'rgba(21,91,184,.07)',fill:true,borderWidth:2,tension:.25,pointRadius:0}]},
    options:{responsive:true,maintainAspectRatio:false,interaction:{mode:'index',intersect:false},
      plugins:{legend:{display:false},tooltip:{callbacks:{label:function(c){return ' '+(c.raw||0).toLocaleString()+' '+(WBP.units[m]||'');}}}},
      scales:{x:{ticks:{color:'#555555',maxTicksLimit:14},grid:{color:'#E8ECF1'}},
        y:{ticks:{color:'#555555'},grid:{color:'#E8ECF1'}}}}});
}
function initForecast(){
  if(_fcInit){ return; } _fcInit=true;
  buildWb();
  var box=document.getElementById('fcBtns'); if(!box) return;
  var keys=Object.keys(FORECAST||{});
  keys.forEach(function(m,i){
    var b=document.createElement('button');
    b.className='mineral-btn fc-btn'+(i===0?' active':'');
    b.textContent=m;
    b.onclick=function(){ _setActive('.fc-btn', b); drawForecast(m); };
    box.appendChild(b);
  });
  if(keys.length) drawForecast(keys[0]);
  drawOutlookChart(); drawSteelCharts();
}
function drawForecast(m){
  var d=FORECAST[m]; if(!d) return; _fcCur=m;
  var t=document.getElementById('fcTitle');
  if(t) t.textContent=m+' 가격 전망 ('+(d.unit||'')+') — '+d.dates[0]+' ~ '+d.dates[d.dates.length-1];
  var actual=d.values.map(function(v,i){ return i<d.split? v: null; });
  var fut=d.values.map(function(v,i){ return i>=d.split-1? v: null; });
  if(_fcChart) _fcChart.destroy();
  _fcChart=new Chart(document.getElementById('fcChart'),{type:'line',
    data:{labels:d.dates,datasets:[
      {label:'실측',data:actual,borderColor:'#155BB8',backgroundColor:'rgba(21,91,184,.08)',fill:true,borderWidth:2,tension:.25,pointRadius:0},
      {label:'예측',data:fut,borderColor:'#c98500',borderDash:[7,5],backgroundColor:'transparent',borderWidth:2,tension:.25,pointRadius:0}]},
    options:{responsive:true,maintainAspectRatio:false,interaction:{mode:'index',intersect:false},
      plugins:{legend:{labels:{color:'#555555',font:{size:11},boxWidth:12}}},
      scales:{x:{ticks:{color:'#555555',maxTicksLimit:14},grid:{color:'#E8ECF1'}},
        y:{ticks:{color:'#555555'},grid:{color:'#E8ECF1'}}}}});
}
var _olChart=null,_stChart=null,_stChart2=null;
function drawOutlookChart(){
  if(_olChart || !OUTLOOK) return;
  var pal={'동':'#155BB8','아연':'#c98500'};
  var keys=Object.keys(OUTLOOK); if(!keys.length) return;
  var labels=OUTLOOK[keys[0]].months;
  var ds=keys.map(function(k){ return {label:k,data:OUTLOOK[k].values,borderColor:pal[k]||'#888888',backgroundColor:'transparent',borderWidth:2,tension:.25,pointRadius:0}; });
  _olChart=new Chart(document.getElementById('olChart'),{type:'line',data:{labels:labels,datasets:ds},
    options:{responsive:true,maintainAspectRatio:false,interaction:{mode:'index',intersect:false},
      plugins:{legend:{labels:{color:'#555555',font:{size:10},boxWidth:10}}},
      scales:{x:{ticks:{color:'#555555',maxTicksLimit:10},grid:{color:'#E8ECF1'}},y:{ticks:{color:'#555555'},grid:{color:'#E8ECF1'}}}}});
}
function drawSteelCharts(){
  if(_stChart || !STEEL || !STEEL.months) return;
  var conf1=[['철광석(달러_톤)','#155BB8'],['유연탄(달러_톤)','#c98500'],['철스크랩(달러_톤)','#1baf7a']];
  var conf2=[['철근(천원_톤)','#155BB8'],['열연(천원_톤)','#c98500'],['후판(천원_톤)','#1baf7a'],['냉연(천원_톤)','#e34948']];
  function mk(id,conf){
    var ds=conf.filter(function(c){return STEEL.series[c[0]];}).map(function(c){
      return {label:c[0].split('(')[0],data:STEEL.series[c[0]],borderColor:c[1],backgroundColor:'transparent',borderWidth:2,tension:.25,pointRadius:0};});
    return new Chart(document.getElementById(id),{type:'line',data:{labels:STEEL.months,datasets:ds},
      options:{responsive:true,maintainAspectRatio:false,interaction:{mode:'index',intersect:false},
        plugins:{legend:{labels:{color:'#555555',font:{size:10},boxWidth:10}}},
        scales:{x:{ticks:{color:'#555555',maxTicksLimit:12},grid:{color:'#E8ECF1'}},y:{ticks:{color:'#555555'},grid:{color:'#E8ECF1'}}}}});
  }
  _stChart=mk('stChart',conf1); _stChart2=mk('stChart2',conf2);
}
// ── 국내 광산 ──
var _mnChart=null;
function drawMines(){
  if(_mnChart || !MINES || !MINES.stats) return;
  var t=MINES.stats.trend;
  _mnChart=new Chart(document.getElementById('mnChart'),{type:'line',data:{labels:t.years,datasets:[
    {label:'가행 광산',data:t['가행'],borderColor:'#1e8e5a',backgroundColor:'rgba(30,142,90,.08)',fill:true,borderWidth:2,tension:.25,pointRadius:3},
    {label:'폐광 (누적)',data:t['폐광'],borderColor:'#d64545',backgroundColor:'transparent',borderWidth:2,tension:.25,pointRadius:3}]},
    options:{responsive:true,maintainAspectRatio:false,interaction:{mode:'index',intersect:false},
      plugins:{legend:{labels:{color:'#555555',font:{size:11},boxWidth:12}}},
      scales:{x:{ticks:{color:'#555555'},grid:{color:'#E8ECF1'}},y:{ticks:{color:'#555555'},grid:{color:'#E8ECF1'}}}}});
}
function drawMineralIndex(){
  if(_mindexChart || typeof MIDX==='undefined' || !MIDX.series) return;
  var conf=[['종합','#c98500'],['에너지광물','#eb6834'],['희소금속','#155BB8'],['메이저금속','#4a3aa7']];
  var base=MIDX.series['종합']||{}; var labels=base.months||[];
  var ds=conf.filter(function(c){return MIDX.series[c[0]];}).map(function(c){
    return {label:c[0],data:MIDX.series[c[0]].values,borderColor:c[1],backgroundColor:'transparent',borderWidth:2,tension:.25,pointRadius:0};
  });
  _mindexChart=new Chart(document.getElementById('mindexChart'),{type:'line',data:{labels:labels,datasets:ds},
    options:{responsive:true,maintainAspectRatio:false,interaction:{mode:'index',intersect:false},
      plugins:{legend:{labels:{color:'#555555',font:{size:11},boxWidth:12}}},
      scales:{x:{ticks:{color:'#555555',maxTicksLimit:12},grid:{color:'#E8ECF1'}},
        y:{ticks:{color:'#555555'},grid:{color:'#E8ECF1'}}}}});
}

"""

    # 움직이는 3D 배경 + 카드 틸트 (인트로와 같은 입체감) — 별도 raw string으로 주입
    # 카드 3D 틸트 + 풀스크린 씬 (가벼운 DOM 전용 — WebGL 배경 제거됨)
    BACKDROP = r"""
<script>
function bindTilt(){
  document.querySelectorAll('.stat-card,.basket-card,.risk-card,.fuel-card').forEach(function(card){
    if(card._tilt) return; card._tilt=1;
    card.addEventListener('animationend',function(){ card.style.animation='none'; });
    card.addEventListener('pointermove',function(e){
      var r=card.getBoundingClientRect(); var x=(e.clientX-r.left)/r.width-.5, y=(e.clientY-r.top)/r.height-.5;
      card.style.transform='perspective(760px) rotateY('+(x*9).toFixed(2)+'deg) rotateX('+(-y*9).toFixed(2)+'deg) translateZ(20px)';
    });
    card.addEventListener('pointerleave',function(){ card.style.transform=''; });
  });
}
bindTilt(); setTimeout(bindTilt,1200);

var SCENE_TITLES={
  'tab-supply':['\ud575\uc2ec\uad11\ubb3c \uc218\uc785\u00b7\uc0dd\uc0b0\uc744 \ud55c\ub208\uc5d0','\uc218\uae09 \ud604\ud669'],
  'tab-mindex':['\uad11\ud574\uad11\uc5c5\uacf5\ub2e8 \ud30c\uc0dd\uc9c0\uc218 \ucd94\uc774','\uac00\uaca9\uc9c0\uc218'],
  'tab-map':['\uc138\uacc4 \uc790\uc6d0 \ubd84\ud3ec\uc640 \uc218\uc785 \ub8e8\ud2b8','\uae00\ub85c\ubc8c \ub9e4\uc7a5\ub7c9'],
  'tab-risk':['\uc218\uae09\uc548\uc815\ud654\uc9c0\uc218 \uc9c4\ub2e8','\ub9ac\uc2a4\ud06c \uc2e0\ud638\ub4f1'],
  'tab-news':['\ub300\uc0c1\ubcc4 \uc790\uc6d0\u00b7\uc6d0\uc790\uc7ac \ub274\uc2a4','\ub274\uc2a4 \ud53c\ub4dc'],
  'tab-subscribe':['\ub9e4\uc77c \ubc1b\uc544\ubcf4\ub294 \ub3d9\ud5a5 \ub9ac\ud3ec\ud2b8','\ub9ac\ud3ec\ud2b8 \uad6c\ub3c5'],
  'tab-komir':['\uad11\uc885\ubcc4 \uad6d\uac00\ubcc4 \uc218\ucd9c\uc785','KOMIR'],
  'tab-usgs':['\uae00\ub85c\ubc8c \ub9e4\uc7a5 \ud1b5\uacc4 2025','USGS'],
};
var SCENE_DRAW={
  'tab-mindex':function(){ if(window.drawMineralIndex) window.drawMineralIndex(); },
  'tab-forecast':function(){ if(window.initForecast) window.initForecast(); },
  'tab-mines':function(){ if(window.drawMines) window.drawMines(); },
  'tab-risk':function(){ if(window.drawRiskChart) window.drawRiskChart(); },
  'tab-map':function(){ if(!window._mapInited&&window.initMap) window.initMap(); },
};
// 풀스크린 씬 비활성 — 대시보드는 밀도 높은 정돈 레이아웃(공백 방지·가독성 우선)
window._applyScenes=function(c){ document.body.classList.remove('scenes'); };
window._applyScenes('minerals');
</script>
"""

    return f"""<!DOCTYPE html>
<html lang="ko">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>K Mineral Risk — 핵심광물 인텔리전스</title>
<link rel="icon" type="image/png" href="/static/favicon.png?v=3">
<link href="https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800;900&family=Archivo:wght@700;800;900&family=Noto+Sans+KR:wght@400;500;700;900&family=IBM+Plex+Mono:wght@400;600&family=JetBrains+Mono:wght@400;500&display=swap" rel="stylesheet">
<link href="https://fonts.googleapis.com/css2?family=Cormorant+Garamond:ital,wght@0,400;0,500;0,600;1,400&family=Noto+Serif+KR:wght@300;500;700&display=swap" rel="stylesheet">
<link rel="stylesheet" href="https://cdn.jsdelivr.net/gh/orioncactus/pretendard@v1.3.9/dist/web/static/pretendard.min.css">
<style>
/* ── 글꼴 통일: Pretendard ── */
:root{{--sans:'Pretendard','Noto Sans KR',-apple-system,sans-serif;--mono:'Pretendard','Noto Sans KR',-apple-system,sans-serif;}}
html body, html body *:not([class*="material-symbols"]){{font-family:'Pretendard','Noto Sans KR',-apple-system,BlinkMacSystemFont,sans-serif !important;}}
body{{font-variant-numeric:tabular-nums;}}
</style>
<script src="https://cdnjs.cloudflare.com/ajax/libs/Chart.js/4.4.1/chart.umd.min.js"></script>
<link rel="stylesheet" href="https://unpkg.com/leaflet@1.9.4/dist/leaflet.css">
<script src="https://unpkg.com/leaflet@1.9.4/dist/leaflet.js"></script>
<style>
:root{{
  --bg:#F5F7FA;--bg2:#ffffff;--bg3:#EEF1F5;
  --border:#DDE3EA;--border2:#C9D2DD;
  --text:#222222;--muted:#555555;--muted2:#888888;
  --red:#d64545;--red-dim:#fdeeee;--red-bright:#c03535;
  --accent:#16305C;--accent2:#16305C;
  --blue:#155BB8;--cyan:#155BB8;--green:#1e8e5a;
  --sans:'Inter','Noto Sans KR',sans-serif;
  --mono:'IBM Plex Mono',monospace;
}}
*{{box-sizing:border-box;margin:0;padding:0}}
body{{background:var(--bg);color:var(--text);font-family:var(--sans);line-height:1.5;height:100vh;display:flex;flex-direction:column;overflow:hidden;}}

/* TICKER */
.ticker{{background:linear-gradient(90deg,#fdf9ee,#faf3dd 50%,#fdf9ee);border-bottom:1px solid #ecdfb8;height:30px;display:flex;align-items:center;overflow:hidden;flex-shrink:0;}}
.ticker-inner{{white-space:nowrap;animation:ticker 40s linear infinite;font-size:11px;font-weight:600;letter-spacing:.08em;color:#8a6a10;font-family:var(--mono);padding-left:100%;}}
@keyframes ticker{{from{{transform:translateX(0)}}to{{transform:translateX(-50%)}}}}

/* NAV */
.nav{{background:var(--bg2);border-bottom:1px solid var(--border);padding:0 20px;display:flex;align-items:center;gap:4px;flex-shrink:0;height:48px;}}
.nav-brand{{display:flex;align-items:center;gap:8px;font-size:13px;font-weight:800;color:var(--blue);letter-spacing:.12em;text-transform:uppercase;font-family:var(--mono);margin-right:20px;}}
.nav-brand .sys-dot{{width:8px;height:8px;border-radius:50%;background:var(--red);animation:sys-blink 1.2s steps(2,start) infinite;}}
@keyframes sys-blink{{to{{opacity:.25}}}}
.nav a{{color:var(--muted);text-decoration:none;font-size:12px;font-weight:500;padding:6px 12px;border-radius:3px;transition:.2s;cursor:pointer;border:1px solid transparent;}}
.nav a:hover{{color:var(--blue);background:var(--bg3);}}
.nav a.active{{color:var(--blue);background:var(--bg3);border-color:rgba(21,91,184,.35);}}
.nav-right{{margin-left:auto;display:flex;align-items:center;gap:12px;}}
.nav-time{{font-size:11px;color:var(--muted);font-family:var(--mono);letter-spacing:.08em;opacity:.85;}}
.nav-conf{{font-size:11px;color:var(--accent);text-decoration:none;font-weight:600;border:1px solid rgba(22,48,92,.4);padding:4px 10px;border-radius:3px;}}
.nav-conf:hover{{background:rgba(21,91,184,.1);}}

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
.choke-tip{{background:#fff;border:1px solid #d64545;color:#222222;font-size:11px;padding:4px 7px;border-radius:4px;box-shadow:0 4px 14px rgba(20,35,60,.18);}}
.leaflet-interactive{{outline:0 !important;}}
.leaflet-grab{{outline:0 !important;}}
.route-line {{ stroke-linecap: round; stroke-linejoin: round; }}
/* 루트 라인 — 점선 흐름 애니메이션 (공급국 → 부산 방향) */
.route-flow {{ stroke-dasharray: 7 11; animation: route-dash 1.1s linear infinite; }}
@keyframes route-dash {{ to {{ stroke-dashoffset: -18; }} }}
.leaflet-tooltip.map-tip {{ background:rgba(255,255,255,.97); border:1px solid var(--border2); color:var(--text); font-size:12px; padding:6px 10px; font-family:var(--mono); box-shadow:0 6px 18px rgba(20,35,60,.14); }}
.leaflet-tooltip.map-tip::before {{ border-right-color:var(--border2); }}

/* 지도 격자 오버레이 (위경도 눈금 느낌) */
.map-grid{{position:absolute;inset:0;z-index:450;pointer-events:none;
  background:
    repeating-linear-gradient(0deg,  transparent 0 79px, rgba(21,91,184,.045) 79px 80px),
    repeating-linear-gradient(90deg, transparent 0 79px, rgba(21,91,184,.045) 79px 80px);}}
.map-grid::after{{content:'';position:absolute;inset:0;
  background:linear-gradient(180deg,transparent 0%,rgba(21,91,184,.02) 50%,transparent 100%);
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
.map-korea-panel{{width:250px;background:linear-gradient(180deg,#ffffff,#f4f7fb);border-left:1px solid var(--border);padding:16px;overflow-y:auto;flex-shrink:0;position:relative;}}
.map-korea-panel::before{{content:'';position:absolute;top:0;left:0;right:0;height:2px;background:linear-gradient(90deg,transparent,var(--blue),transparent);opacity:.6;}}
.kp-title{{font-size:11px;font-weight:700;color:var(--blue);letter-spacing:.22em;text-transform:uppercase;font-family:var(--mono);margin-bottom:4px;}}
.kp-sub{{font-size:9px;color:var(--muted2);font-family:var(--mono);letter-spacing:.12em;margin-bottom:12px;}}
.kp-flag{{font-size:22px;margin-bottom:6px;}}
.kp-desc{{font-size:12px;color:var(--muted);line-height:1.6;margin-bottom:12px;}}
.kp-row{{padding:6px 0 7px;border-bottom:1px solid var(--border);opacity:0;transform:translateX(-8px);animation:kp-scan .35s ease forwards;}}
@keyframes kp-scan{{to{{opacity:1;transform:translateX(0)}}}}
.kp-line{{display:flex;justify-content:space-between;align-items:baseline;}}
.kp-country{{font-size:12px;color:var(--text);font-family:var(--mono);}}
.kp-amount{{font-size:11px;font-family:var(--mono);color:var(--accent);}}
.kp-bar{{margin-top:4px;height:4px;background:var(--bg3);border-radius:2px;overflow:hidden;}}
.kp-bar i{{display:block;height:100%;background:linear-gradient(90deg,#6da7ec,var(--blue));border-radius:2px;}}
.kp-bar.warn i{{background:linear-gradient(90deg,#B9CCEA,var(--accent));}}
.kp-bar.crit i{{background:linear-gradient(90deg,#e88a8a,var(--red));}}

/* RISK 배지 */
.risk-badge{{margin-left:auto;font-family:var(--mono);font-size:11px;font-weight:700;letter-spacing:.14em;
  padding:5px 12px;border-radius:3px;border:1px solid;display:inline-flex;align-items:center;gap:7px;}}
.risk-badge .rb-dot{{width:7px;height:7px;border-radius:50%;background:currentColor;box-shadow:0 0 7px currentColor;}}
.risk-badge.high{{color:var(--red-bright);border-color:rgba(214,69,69,.5);background:rgba(214,69,69,.07);}}
.risk-badge.high .rb-dot{{animation:cp-blink .8s steps(2,start) infinite;}}
.risk-badge.medium{{color:var(--accent);border-color:rgba(22,48,92,.5);background:rgba(21,91,184,.08);}}
.risk-badge.low{{color:var(--green);border-color:rgba(30,142,90,.5);background:rgba(30,142,90,.07);}}
.map-legend{{position:absolute;bottom:20px;left:20px;background:rgba(255,255,255,.95);border:1px solid var(--border);border-radius:8px;padding:10px 14px;z-index:1000;font-size:11px;box-shadow:0 6px 20px rgba(20,35,60,.12);}}
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
<style>
/* ══ V2 딥그린 스킨 — 신규 홈과 디자인 통일 ══ */
:root{{
  --bg:#F5F7FA;--bg2:#ffffff;--bg3:#EEF1F5;
  --border:#DDE3EA;--border2:#C9D2DD;
  --text:#222222;--muted:#555555;--muted2:#888888;
  --red:#d8453c;--red-dim:#fdecea;--red-bright:#c0392b;
  --accent:#155BB8;--accent2:#16305C;
  --navy:#16305C;--navy2:#155BB8;
  --blue:#155BB8;--cyan:#1E74D8;--green:#1e8e5a;
}}
</style>
<style>{V2_CHROME_CSS}</style>
</head>
<body class="{body_cls}">
{V2_CHROME_HEADER}
<div id="cosmos"></div>
<div id="arrival">
  <div class="a-eyebrow" id="a-eyebrow"></div>
  <div class="a-title" id="a-title"></div>
  <div class="a-line"></div>
  <div class="a-name" id="a-name"></div>
</div>

<!-- 상단 카테고리 전환 -->
<div class="cat-bar">
  <a href="/" class="brand-lock" title="허브 홈"><span class="brand-txt">K<em>MR</em><small>MINERAL INTELLIGENCE</small></span></a>
  <button class="cat-btn" data-cat="all" onclick="switchCategory('all',this)">전체</button>
  <div class="cat-menu">
    <button class="cat-btn active" data-cat="minerals" onclick="switchCategory('minerals',this)">핵심광물</button>
    <div class="megapanel"><div class="mp-grid mp-c4">
      <a class="mp-tile" onclick="goSec('minerals','supply')"><b>수급 현황</b><span>수입·생산 한눈에</span></a>
      <a class="mp-tile" onclick="goSec('minerals','mindex')"><b>가격지수</b><span>광해광업공단 파생지수</span></a>
      <a class="mp-tile" onclick="goSec('minerals','forecast')"><b>가격 전망</b><span>11광종 예측 · ~2028</span></a>
      <a class="mp-tile" onclick="goSec('minerals','map')"><b>글로벌 매장량</b><span>세계 분포·수입루트</span></a>
      <a class="mp-tile" onclick="goSec('minerals','risk')"><b>리스크 신호등</b><span>K-RISK 종합 위험</span></a>
      <a class="mp-tile" onclick="goSec('minerals','mines')"><b>국내 광산</b><span>가행·폐광 통계</span></a>
      <a class="mp-tile" onclick="goSec('minerals','news')"><b>뉴스 피드</b><span>대상별 자원 뉴스</span></a>
    </div></div>
  </div>
  <button class="cat-btn" data-cat="nf" onclick="switchCategory('nf',this)">비철금속</button>
  <button class="cat-btn" data-cat="rare" onclick="switchCategory('rare',this)">희소금속</button>
  <button class="cat-btn" data-cat="ree" onclick="switchCategory('ree',this)">희토류</button>
  <button class="cat-btn" data-cat="energy" onclick="switchCategory('energy',this)">에너지</button>
  <button class="cat-btn" data-cat="etc" onclick="switchCategory('etc',this)">기타</button>
  <div class="cb-right">
    <a href="/globe" class="cb-link">핵심광물지도</a>
    <a href="/conference" class="cb-link">AI 회의실</a>
    <a href="/" class="to-space" title="허브 홈으로">허브 홈</a>
  </div>
</div>

<!-- TICKER (핵심광물 전용) -->
<div class="ticker" id="mineralTicker">
  <div class="ticker-inner">
    ⚠ 자원 리스크 실시간 모니터링 &nbsp;|&nbsp; {ticker_items} &nbsp;|&nbsp;
    ⚠ 자원 리스크 실시간 모니터링 &nbsp;|&nbsp; {ticker_items} &nbsp;|&nbsp;
  </div>
</div>

<!-- NAV (사이드바, 카테고리별 하위 탭) -->
<nav class="nav">
  <span class="nav-brand"><span class="sys-dot"></span>K MINERAL RISK MONITOR</span>
  <div id="subnav-minerals">
    <a href="#" class="active" data-tab="supply"    onclick="switchTab('supply',this);return false;">수급 현황</a>
    <a href="#" data-tab="mindex"    onclick="switchTab('mindex',this);return false;">가격지수</a>
    <a href="#" data-tab="forecast"  onclick="switchTab('forecast',this);return false;">가격 전망</a>
    <a href="#" data-tab="map"       onclick="switchTab('map',this);return false;">글로벌 매장량</a>
    <a href="#" data-tab="risk"      onclick="switchTab('risk',this);return false;">리스크 신호등</a>
    <a href="#" data-tab="mines"     onclick="switchTab('mines',this);return false;">국내 광산</a>
    <a href="#" data-tab="news"      onclick="switchTab('news',this);return false;">뉴스 피드</a>
  </div>
  <div class="nav-right">
    <span class="nav-time" id="nav-clock">{now}</span>
    <a href="/conference" class="nav-conf">AI 전문가 회의실 →</a>
  </div>
</nav>

<!-- ===== 히어로 배너 (슬라이드쇼 + 통합검색) ===== -->
<div class="hero" id="hero">
  <div class="hero-clip">
    <div class="hero-track" id="heroTrack">
      <a class="hero-slide hs-min" onclick="goSec('minerals','supply')">
        <div class="hs-in"><div class="hs-eyebrow">Critical Minerals</div>
          <div class="hs-title">핵심광물 공급망을 <b>실시간으로</b></div>
          <div class="hs-sub">리튬·희토류·니켈 — 수입 의존 95%의 흐름을 한눈에</div></div></a>
      <a class="hero-slide hs-ai" href="/conference">
        <div class="hs-in"><div class="hs-eyebrow">AI Insight</div>
          <div class="hs-title">전문가 AI가 <b>자원을 토론한다</b></div>
          <div class="hs-sub">대상 맞춤 다중 전문가 회의 — AI 회의실 →</div></div></a>
    </div>
  </div>
  <button class="hero-nav prev" onclick="heroGo(-1)">‹</button>
  <button class="hero-nav next" onclick="heroGo(1)">›</button>
  <div class="hero-dots" id="heroDots"></div>
  <div class="hero-search">
    <span class="hsr-cat">통합 검색</span>
    <input id="heroQ" placeholder="광물·품목·키워드를 검색하세요" onkeydown="if(event.key==='Enter')heroSearch()">
    <button class="hsr-btn" onclick="heroSearch()" title="검색">🔍</button>
  </div>
</div>

<!-- ===== 메인(홈) 랜딩 — 카테고리 바로가기 + 오늘의 지표 ===== -->
<div id="home-landing">
  <div class="hl-wrap">
    <div class="hl-head">
      <div class="hl-eyebrow">Resource Intelligence</div>
      <h2 class="hl-title">무엇을 살펴볼까요?</h2>
      <p class="hl-sub">핵심광물 대시보드로 들어가세요.</p>
    </div>
    <div class="hl-cards">
      <a class="hl-card hl-min" href="/dashboard?cat=minerals"><div class="hl-ic">💎</div><div class="hl-nm">핵심광물 종합</div><div class="hl-dc">수급·가격지수·전망·매장량·리스크·국내광산</div><div class="hl-go">대시보드 →</div></a>
      <a class="hl-card hl-min" href="/dashboard?cat=nf"><div class="hl-ic">🔩</div><div class="hl-nm">비철금속</div><div class="hl-dc">니켈·동·알루미늄·주석·연·아연 + LME 시세</div><div class="hl-go">대시보드 →</div></a>
      <a class="hl-card hl-min" href="/dashboard?cat=rare"><div class="hl-ic">⚗️</div><div class="hl-nm">희소금속</div><div class="hl-dc">리튬·코발트·텅스텐 등 20종 전략 소재</div><div class="hl-go">대시보드 →</div></a>
      <a class="hl-card hl-min" href="/dashboard?cat=ree"><div class="hl-ic">🧲</div><div class="hl-nm">희토류</div><div class="hl-dc">네오디뮴·디스프로슘 등 14원소</div><div class="hl-go">대시보드 →</div></a>
      <a class="hl-card hl-min" href="/dashboard?cat=energy"><div class="hl-ic">⚡</div><div class="hl-nm">에너지</div><div class="hl-dc">우라늄·유연탄 — 발전 연료 광물</div><div class="hl-go">대시보드 →</div></a>
      <a class="hl-card hl-min" href="/dashboard?cat=etc"><div class="hl-ic">⛏️</div><div class="hl-nm">기타 광물</div><div class="hl-dc">철·흑연·금·은·백금·팔라듐</div><div class="hl-go">대시보드 →</div></a>
    </div>
    {home_risk_html}
    <div class="hl-metal"><div class="hl-risk-head">오늘의 금속 시세 <span>런던금속거래소(LME) 종가 · 조달청 비축물자 · {ppa_date}</span></div>
      <div class="hm-row">{home_metal_html}</div></div>
    <div class="hl-stats">
      <div class="hl-stat"><span>최대 수입 광물</span><b>{top_min}</b></div>
      <div class="hl-stat"><span>총 광물 수입액</span><b>${total:,.0f}</b></div>
      <div class="hl-stat"><span>수집 뉴스</span><b>{len(news)}건</b></div>
    </div>
    <div style="margin-top:34px;">
      <div class="sub-box">
        <div class="sub-title">📩 매일 자원 동향 리포트 구독</div>
        <div class="sub-desc">
          핵심광물 수급·가격의 핵심 흐름과 글로벌 이슈를 매일 이메일로 받아보세요.<br>
          현재 <strong>{len(subs)}명</strong>이 구독 중입니다.
        </div>
        <input id="sub-email" class="sub-input" type="email" placeholder="이메일 주소 입력">
        <button class="sub-btn" onclick="doSubscribe()">구독 신청</button>
        <button class="sub-btn2" onclick="doSendNow()">지금 바로 받기 (1회)</button>
        <div class="sub-msg" id="sub-msg"></div>
      </div>
    </div>
  </div>
</div>

<!-- ===== 핵심광물 카테고리 (기존 탭 6개) ===== -->
<div id="cat-minerals">

<!-- ============================
     TAB: 수급 현황
     ============================ -->
<div id="tab-supply" class="tab-panel active">
  <div class="dash">
    <!-- KPI 행 -->
    <div class="stat-row dash-kpis">
      <div class="stat-card red"><div class="sc-label">총 수입액</div><div class="sc-val red">${total:,.0f}</div><div class="sc-sub">KOMIR 기준</div></div>
      <div class="stat-card"><div class="sc-label">최대 수입 광물</div><div class="sc-val">{top_min}</div><div class="sc-sub">수입액 1위</div></div>
      <div class="stat-card"><div class="sc-label">최대 수입국</div><div class="sc-val">{top_cntry}</div><div class="sc-sub">국가별 1위</div></div>
      <div class="stat-card"><div class="sc-label">공급 리스크 경보</div><div class="sc-val">{len(_risk_high)}<small style="font-size:14px;font-weight:600">건</small></div><div class="sc-sub">수급 불안 광종</div></div>
      <div class="stat-card"><div class="sc-label">국내 가행 광산</div><div class="sc-val">{mine_active}<small style="font-size:14px;font-weight:600">개</small></div><div class="sc-sub">{mine_latest_year}년 · KOMIR</div></div>
      <div class="stat-card"><div class="sc-label">LME 지수</div><div class="sc-val">{ppa_lme_s}</div><div class="sc-sub">런던금속거래소 · {ppa_date}</div></div>
      <div class="stat-card"><div class="sc-label">뉴스</div><div class="sc-val">{len(news)}</div><div class="sc-sub">수집된 기사</div></div>
    </div>

    <!-- 3열 격자 -->
    <div class="dash-cols">
      <!-- 좌: 순위·생산국 -->
      <div class="dash-col">
        <div class="wpanel grow"><div class="wp-head">글로벌 매장량 순위 <span class="wp-sub">KOMIR 2026</span></div><div class="wp-body">{usgs_rank_html}</div></div>
        <div class="wpanel"><div class="wp-head">오늘의 금속 시세 <span class="wp-sub">LME · {ppa_date}</span></div><div class="wp-body">{ppa_rows}</div></div>
        <div class="wpanel"><div class="wp-head">매장량 1위국</div><div class="wp-body">{prod_html}</div></div>
      </div>
      <!-- 중: 차트 -->
      <div class="dash-col">
        <div class="wpanel grow"><div class="wp-head">광물별 수입액 <span class="wp-sub">상위 7</span></div><div class="wp-chart"><canvas id="chartMin"></canvas></div></div>
        <div class="wpanel grow"><div class="wp-head">국가별 수입액 <span class="wp-sub">상위 7</span></div><div class="wp-chart"><canvas id="chartCnt"></canvas></div></div>
      </div>
      <!-- 우: 데이터 표 -->
      <div class="dash-col">
        <div class="wpanel grow"><div class="wp-head">광물별 수입 현황 <span class="wp-sub">KOMIR</span></div>
          <div class="wp-body"><table class="wp-table"><tbody>{trade_rows}</tbody></table></div></div>
      </div>
    </div>

    <!-- 월간 수입 동향 + 교차 검증 -->
    <div class="charts-row" style="height:auto!important;align-items:stretch;margin-top:2px;">
      <div class="wpanel" style="flex:1.15;">
        <div class="wp-head">광물별 월간 수입 동향 <span class="wp-sub">관세청 · {core_asof}</span></div>
        <div style="padding:10px 16px 2px;display:flex;gap:6px;flex-wrap:wrap" id="coreBtns"></div>
        <div class="wp-chart" style="min-height:220px;"><canvas id="coreChart"></canvas></div>
        <div style="padding:0 16px 12px;font-size:11.5px;color:var(--muted2)" id="coreNote"></div>
      </div>
      <div class="wpanel" style="flex:1;">
        <div class="wp-head">데이터 교차 검증 <span class="wp-sub">KOMIR 연간 vs 관세청 12개월</span></div>
        <div class="wp-body">
          <table class="wp-table">
            <tr style="color:var(--muted2);font-size:10.5px"><td style="padding:4px">광물</td><td style="padding:4px;text-align:right">관세청 12M</td><td style="padding:4px;text-align:right">KOMIR 연간</td><td style="padding:4px;text-align:right">배율</td><td style="padding:4px">최대 수입국</td></tr>
            {cross_rows}
          </table>
          <div style="font-size:10.5px;color:var(--muted2);margin-top:8px;line-height:1.6">💡 KOMIR는 광종(원광 중심)·연간, 관세청은 HS코드(광석+금속 형태)·최근 12개월 기준이라 차이가 납니다. 배율 <b style="color:#1e8e5a">×0.5~2</b>는 정합, 그 외는 집계 범위 차이.</div>
        </div>
      </div>
    </div>

    <!-- 확대 광종 커버리지 -->
    <div class="wpanel" style="margin-top:2px">
      <div class="wp-head">K Mineral Risk 커버리지 — 확대 대상 광종 <span class="wp-sub">{taxo_total}종 · 5개 분류</span></div>
      <div class="wp-body" style="padding:12px 16px 14px">{taxo_html}</div>
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
      <a class="mode-btn" href="/globe" style="text-decoration:none">🌐 3D 지구본으로 보기</a>
      <span class="map-ctrl-label" style="margin-left:12px;">광물:</span>
      <button class="mineral-btn active" onclick="selectMineral('리튬',this)">리튬</button>
      <button class="mineral-btn" onclick="selectMineral('코발트',this)">코발트</button>
      <button class="mineral-btn" onclick="selectMineral('니켈',this)">니켈</button>
      <button class="mineral-btn" onclick="selectMineral('흑연',this)">흑연</button>
      <button class="mineral-btn" onclick="selectMineral('희토류',this)">희토류</button>
      <button class="mineral-btn" onclick="selectMineral('망간',this)">망간</button>
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
          <div class="legend-item"><div class="legend-color" style="background:#CBDDF6"></div> 소형 (1-5%)</div>
          <div class="legend-item"><div class="legend-color" style="background:#88bb44"></div> 미량 (&lt;1%)</div>
          <div class="legend-item"><div class="legend-color" style="background:#155BB8"></div> 🇰🇷 한국 (수입의존)</div>
        </div>
        <!-- 초크포인트 뉴스 패널 -->
        <div id="choke-panel" style="
          display:none;position:absolute;top:10px;right:10px;z-index:1500;
          width:320px;max-height:70vh;overflow-y:auto;
          background:rgba(255,255,255,0.97);border:1px solid #dbe2ec;
          border-radius:8px;padding:14px;
          box-shadow:0 10px 30px rgba(20,35,60,0.16);
          scrollbar-width:thin;scrollbar-color:#C9D2DD transparent;
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
    <div style="padding:16px 20px;background:var(--bg)">
      <div class="page-title">광종별 세계 생산·매장 <span style="color:var(--muted2);font-weight:400;font-size:12px">· USGS MCS 2026 · 25개 광종 · 2025년 기준</span></div>
      <div class="usgs-grid">{usgs2_cards}</div>
    </div>
  </div>
</div>

<!-- ============================
     TAB: 뉴스 피드
     ============================ -->
<div id="tab-mindex" class="tab-panel">
  <div class="page-title">광물 가격지수 <span style="color:var(--muted2);font-weight:400;font-size:12px">· 한국광해광업공단 파생지수 · 2016년1월=1000 기준 · 2012~2025 월별</span></div>
  <div style="background:var(--bg2);border:1px solid var(--border);border-radius:10px;padding:12px 16px;margin-bottom:14px;font-size:13px;color:var(--muted);">💡 4개 지수로 광물 시장을 한눈에 — <b>희소금속</b>엔 리튬·희토류·코발트, <b>에너지광물</b>엔 연료탄·우라늄이 들어갑니다. 지수가 오르면 해당 광물군 가격 상승.</div>
  <div class="risk-grid">{midx_cards}</div>
  <div class="section" style="padding:14px 16px;margin-top:14px;">
    <div class="chart-title">광물군별 가격지수 추이 (월별)</div>
    <div style="height:320px;position:relative;"><canvas id="mindexChart"></canvas></div>
  </div>
  <div style="text-align:center;margin-top:16px;"><a href="/conference" class="nav-conf">⚖️ AI 전문가 회의실에서 광물 시장 토론하기 →</a></div>
</div>

<div id="tab-risk" class="tab-panel">
  <div class="page-title">K-RISK 종합 공급망 위험 <span style="color:var(--muted2);font-weight:400;font-size:12px">· 산업부 공공데이터 교차 계산 · 0~100, 높을수록 위험</span></div>
  <div style="background:var(--bg2);border:1px solid var(--border);border-radius:10px;padding:12px 16px;margin-bottom:14px;font-size:13px;color:var(--muted);">💡 {krisk_summary} <span style="color:var(--muted2)">— 따로 발표되던 지표를 하나의 위험 점수로 융합. 자세한 영향은 AI 회의실에서.</span></div>
  <div class="risk-grid">{krisk_cards}</div>
  <div style="font-size:11.5px;color:var(--muted2);margin:10px 2px 0;line-height:1.75;">
    산식 <b>K-RISK = 0.35×수급불안정(100−수급안정화지수) + 0.25×수입집중도(HHI) + 0.20×지정학 + 0.20×가격변동성</b>
    · 🟢 0~39 안정 🟡 40~69 주의 🔴 70~100 위험<br>
    데이터: 광해광업공단 수급안정화지수·국가별 광종 수출입·파생지수 + USGS MCS — 갱신 시 자동 재계산.
    시장위험지수·비축 항목은 공개 데이터 확보 시 가중 반영 예정(현재 잔여 항목 재배분).
  </div>
  <div class="page-title" id="ssi" style="margin-top:22px;font-size:15px;scroll-margin-top:120px;">구성 원지표 — 수급안정화지수 <span style="color:var(--muted2);font-weight:400;font-size:12px">· 한국광해광업공단 · 지수 높을수록 수급 안정</span></div>
  <div style="background:var(--bg2);border:1px solid var(--border);border-radius:10px;padding:12px 16px;margin-bottom:14px;font-size:13px;color:var(--muted);">💡 {risk_summary} <span style="color:var(--muted2)">— K-RISK 수급불안정 요소의 원자료.</span></div>
  <div class="risk-grid">{risk_cards}</div>
  <div class="section" style="padding:14px 16px;margin-top:14px;">
    <div class="chart-title">수급안정화지수 추이 (최근 3년, 월별)</div>
    <div style="height:300px;position:relative;"><canvas id="riskChart"></canvas></div>
  </div>
  <div style="text-align:center;margin-top:16px;"><a href="/conference" class="nav-conf">⚖️ AI 전문가 회의실에서 리스크 토론하기 →</a></div>
</div>

<div id="tab-news" class="tab-panel">
  <div class="page-title">자원·원자재 뉴스 — 대상별</div>
  <div id="aiBrief" class="ai-brief" style="display:none">🤖 AI가 오늘의 뉴스를 분석 중...</div>
  <div style="display:flex;gap:8px;flex-wrap:wrap;margin-bottom:14px;">
    <button class="mineral-btn news-aud-btn active" onclick="filterNews('전체',this)">전체</button>
    <button class="mineral-btn news-aud-btn" onclick="filterNews('투자자',this)">📈 투자자용</button>
    <button class="mineral-btn news-aud-btn" onclick="filterNews('기업',this)">🏢 기업용</button>
    <button class="mineral-btn news-aud-btn" onclick="filterNews('소비자',this)">🛒 소비자용</button>
  </div>
  <div id="newsHero"></div>
  <div class="news-grid" id="newsGrid"></div>
</div>

<!-- ============================
     TAB: 가격 전망
     ============================ -->
<div id="tab-forecast" class="tab-panel">
  <div class="page-title">광물 가격 전망 <span style="color:var(--muted2);font-weight:400;font-size:12px">· 한국광해광업공단 가격예측데이터 · 분기별 · 실선=실측 · 점선=예측</span></div>
  <div style="background:var(--bg2);border:1px solid var(--border);border-radius:10px;padding:12px 16px;margin-bottom:14px;font-size:13px;color:var(--muted);">💡 광종을 선택하면 2013년부터의 실측 가격과 2028년까지의 <b>AI 예측 가격</b>이 표시됩니다. 급등 구간은 조달·비축 시점 판단에 활용하세요.</div>
  <div style="display:flex;gap:8px;flex-wrap:wrap;margin-bottom:14px;" id="fcBtns"></div>
  <div class="section" style="padding:14px 16px;">
    <div class="chart-title" id="fcTitle">가격 전망</div>
    <div style="height:340px;position:relative;"><canvas id="fcChart"></canvas></div>
  </div>
  <div class="charts-row" style="height:auto!important;margin-top:14px;">
    <div class="section" style="flex:1;padding:14px 16px;">
      <div class="chart-title">시장전망지표 — 동·아연 <span style="color:var(--muted2)">· KOMIR · 높을수록 시장 낙관</span></div>
      <div style="height:240px;position:relative;"><canvas id="olChart"></canvas></div>
    </div>
    <div class="section" style="flex:1;padding:14px 16px;">
      <div class="chart-title">철강 원자재 국제가 <span style="color:var(--muted2)">· 산업통상부 · 달러/톤</span></div>
      <div style="height:240px;position:relative;"><canvas id="stChart"></canvas></div>
    </div>
  </div>
  <div class="section" style="padding:14px 16px;margin-top:14px;">
    <div class="chart-title">국내 철강 제품가 <span style="color:var(--muted2)">· 철근·열연·후판·냉연 · 천원/톤</span></div>
    <div style="height:240px;position:relative;"><canvas id="stChart2"></canvas></div>
  </div>
  <div class="section" style="padding:14px 16px;margin-top:14px;">
    <div class="chart-title" id="wbTitle">국제 원자재 시세 <span style="color:var(--muted2)">· World Bank Pink Sheet · 월별 · 최근 15년</span></div>
    <div style="display:flex;gap:6px;flex-wrap:wrap;margin:6px 0 10px" id="wbBtns"></div>
    <div style="height:280px;position:relative;"><canvas id="wbChart"></canvas></div>
  </div>
  <div style="text-align:center;margin-top:16px;"><a href="/conference" class="nav-conf">⚖️ AI 전문가 회의실에서 가격 전망 토론하기 →</a></div>
</div>

<!-- ============================
     TAB: 국내 광산
     ============================ -->
<div id="tab-mines" class="tab-panel">
  <div class="page-title">국내 광산 현황 <span style="color:var(--muted2);font-weight:400;font-size:12px">· 한국광해광업공단 전국광산 통계 · {mine_latest_year}년</span></div>
  <div class="stat-row">
    <div class="stat-card"><div class="sc-label">가행 광산</div><div class="sc-val">{mine_active}<small style="font-size:14px;font-weight:600">개</small></div><div class="sc-sub">운영 중 · {mine_latest_year}년</div></div>
    <div class="stat-card"><div class="sc-label">폐광산 (누적)</div><div class="sc-val">{mine_closed_total:,}<small style="font-size:14px;font-weight:600">개</small></div><div class="sc-sub">전국 폐광산 위치 정보</div></div>
    <div class="stat-card"><div class="sc-label">금속 광산</div><div class="sc-val">{_mg.get("금속", 0)}<small style="font-size:14px;font-weight:600">개</small></div><div class="sc-sub">가행 기준</div></div>
    <div class="stat-card"><div class="sc-label">비금속 광산</div><div class="sc-val">{_mg.get("비금속", 0)}<small style="font-size:14px;font-weight:600">개</small></div><div class="sc-sub">가행 기준</div></div>
    <div class="stat-card"><div class="sc-label">석탄 광산</div><div class="sc-val">{_mg.get("석탄", 0)}<small style="font-size:14px;font-weight:600">개</small></div><div class="sc-sub">가행 기준</div></div>
  </div>
  <div class="charts-row" style="height:auto!important;">
    <div class="section" style="flex:1.4;padding:14px 16px;">
      <div class="chart-title">연도별 가행 · 폐광 광산 수 추이</div>
      <div style="height:280px;position:relative;"><canvas id="mnChart"></canvas></div>
    </div>
    <div class="section" style="flex:1;padding:14px 16px;">
      <div class="chart-title">폐광산 많은 지역 TOP 10 <span style="color:var(--muted2)">· 시도별</span></div>
      {mine_sido_rows}
    </div>
  </div>
  <div style="background:var(--bg2);border:1px solid var(--border);border-radius:10px;padding:12px 16px;margin-top:14px;font-size:13px;color:var(--muted);">💡 국내 금속 광산은 <b>가행 {_mg.get("금속", 0)}개</b>에 불과해 핵심광물 자급이 어렵습니다 — 수입 의존도가 높은 이유. 폐광산 {mine_closed_total:,}곳의 광해방지·재자원화가 KOMIR의 핵심 사업입니다.</div>
  <div style="text-align:center;margin-top:16px;"><a href="/conference" class="nav-conf">⚖️ AI 전문가 회의실에서 국산화·재자원화 토론하기 →</a></div>
</div>

<!-- ============================
     TAB: 리포트 구독
     ============================ -->
<!-- 리포트 구독 섹션은 메인 화면(home-landing)으로 이동함 -->

</div><!-- /#cat-minerals -->

<!-- ===== 분류별 카테고리 페이지 ===== -->
{cat_pages_html}

<script>var RISK = {risk_js}; var MIDX = {midx_js}; var NEWS = {news_js}; var FORECAST = {forecast_js}; var CATFC = {catfc_js}; var REETRADE = {ree_js}; var CORETRADE = {core_js}; var WBP = {wbp_js}; var STEEL = {steel_js}; var MINES = {mines_js}; var OUTLOOK = {outlook_js};</script>
<script>{CAT_JS}</script>

<script>
// ── 탭 전환 ──────────────────────────────────────────────────
function switchTab(name, el) {{
  window._curTab = name;
  document.querySelectorAll('.tab-panel').forEach(p => p.classList.remove('active'));
  document.querySelectorAll('.nav a[data-tab]').forEach(a => a.classList.remove('active'));
  document.getElementById('tab-' + name).classList.add('active');
  if (el) el.classList.add('active');
  if (name === 'map' && !window._mapInited) initMap();
  if (name === 'risk' && typeof drawRiskChart === 'function') drawRiskChart();
  if (name === 'risk') {{
    var ab = document.querySelector('.cat-btn.active');
    if (ab && typeof switchCategory === 'function') setTimeout(function(){{ switchCategory(ab.dataset.cat, ab); }}, 0);
  }}
  if (name === 'mindex' && typeof drawMineralIndex === 'function') drawMineralIndex();
  if (name === 'forecast' && typeof initForecast === 'function') initForecast();
  if (name === 'mines' && typeof drawMines === 'function') drawMines();
}}

// 다른 페이지(회의실 등)에서 #map / #news 등으로 들어오면 해당 탭으로 이동
function _applyHashNav(){{
  var h = (location.hash || '').replace('#','');
  if (h === 'ssi') {{
    setTimeout(function(){{
      switchTab('risk', document.querySelector('.nav a[data-tab="risk"]'));
      setTimeout(function(){{
        var el = document.getElementById('ssi');
        if (el) el.scrollIntoView({{behavior:'smooth', block:'start'}});
      }}, 200);
    }}, 0);
    return;
  }}
  if (h === 'routes') {{
    setTimeout(function(){{
      switchTab('map', document.querySelector('.nav a[data-tab="map"]'));
      setTimeout(function(){{
        var b = document.getElementById('modeRoutes');
        if (b && typeof setMode === 'function') setMode('routes', b);
      }}, 400);
    }}, 0);
    return;
  }}
  var valid = ['supply','mindex','forecast','map','news','subscribe','risk','mines'];
  if (valid.indexOf(h) >= 0) {{
    setTimeout(function(){{
      switchTab(h, document.querySelector('.nav a[data-tab="' + h + '"]'));
    }}, 0);
  }} else if (h.indexOf('cat-') === 0) {{
    var c = h.slice(4);
    setTimeout(function(){{
      if (typeof switchCategory === 'function') {{
        var b = document.querySelector('.cat-btn[data-cat="' + c + '"]');
        switchCategory(c, b);
      }}
    }}, 0);
  }}
}}
_applyHashNav();
window.addEventListener('hashchange', _applyHashNav);

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
    x: {{ ticks: {{ color: '#555555', font: {{ size: 10 }} }}, grid: {{ color: '#E8ECF1' }} }},
    y: {{ ticks: {{ color: '#555555', font: {{ size: 10 }}, callback: v => '$' + (v/1e6).toFixed(1)+'M' }}, grid: {{ color: '#E8ECF1' }} }}
  }}
}};
const isCntTon = {imports_unit_js} === '톤';
const CHART_CNT_OPTS = {{
  responsive: true, maintainAspectRatio: false,
  plugins: {{ legend: {{ display: false }}, tooltip: {{ callbacks: {{ label: ctx => ' ' + ctx.raw.toLocaleString() + (isCntTon ? ' 톤' : '') }} }} }},
  scales: {{
    x: {{ ticks: {{ color: '#555555', font: {{ size: 10 }} }}, grid: {{ color: '#E8ECF1' }} }},
    y: {{ ticks: {{ color: '#555555', font: {{ size: 10 }}, callback: v => isCntTon ? (v/1e3).toFixed(0)+'K톤' : '$'+(v/1e6).toFixed(1)+'M' }}, grid: {{ color: '#E8ECF1' }} }}
  }}
}};
new Chart(document.getElementById('chartMin'), {{
  type: 'bar',
  data: {{ labels: {cl}, datasets: [{{ data: {cd}, backgroundColor: '#155BB8', borderRadius: 3 }}] }},
  options: CHART_OPTS
}});
new Chart(document.getElementById('chartCnt'), {{
  type: 'bar',
  data: {{ labels: {cl2}, datasets: [{{ data: {cd2}, backgroundColor: '#c98500', borderRadius: 3 }}] }},
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
  L.tileLayer('https://{{s}}.basemaps.cartocdn.com/light_all/{{z}}/{{x}}/{{y}}{{r}}.png', {{
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
  if (share >=  1) return '#CBDDF6';
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
      if (iso === 'KOR') return {{ fillColor:'#155BB8', fillOpacity:.9, color:'#fff', weight:2 }};
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
  _mapMode = mode; window._mapMode = mode;
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
  {{ key:'LOMBOK',   name:'롬복 해협',     pos:WP.LOMBOK,    color:'#B9CCEA', risk:'low',
     reason:'말라카 우회 대안. 수심 깊어 대형 선박 통과 가능. 인도네시아 정세에 의존' }},
  {{ key:'TORRES',   name:'토레스 해협',   pos:WP.TORRES,    color:'#B9CCEA', risk:'low',
     reason:'호주 동부~아시아 경로. 수심 얕고 암초 많아 항법 주의. 호주산 니켈·코발트 수입에 활용' }},
  {{ key:'CAPE_HORN',name:'케이프혼',      pos:WP.CAPE_HORN, color:'#B9CCEA', risk:'low',
     reason:'칠레산 리튬·구리의 아시아행 경로. 강풍·너울로 운항 위험. 파나마 막힐 경우 대안 경로' }},
];

// 3D 지구(globe.gl) 스크립트에서 공유하도록 window에 노출 (const/let은 기본적으로 window 미등록)
window.WORLD_RESERVES = WORLD_RESERVES;
window.CHOKEPOINTS = CHOKEPOINTS;
window._currentMineral = _currentMineral;
window._mapMode = _mapMode;

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
      <span style="font-size:15px;font-weight:bold;color:#155BB8;font-family:var(--mono);">⚓ ${{cp.name}}</span>
      <button onclick="document.getElementById('choke-panel').style.display='none'"
        style="background:none;border:none;color:#888888;font-size:18px;cursor:pointer;line-height:1;">✕</button>
    </div>
    <div style="font-size:11px;color:${{cp.color}};margin-bottom:6px;font-weight:bold;">${{RISK_KO[cp.risk] || cp.risk}}</div>
    <div style="font-size:11px;color:#42536e;margin-bottom:12px;line-height:1.5;">${{cp.reason}}</div>
    <div id="choke-news-list" style="font-size:11px;color:#7D8AA3;">뉴스 로딩 중...</div>
  `;
  panel.style.display = 'block';

  fetch('/api/chokepoint-news?key=' + cp.key)
    .then(r => r.json())
    .then(data => {{
      const list = document.getElementById('choke-news-list');
      if (!data.articles || data.articles.length === 0) {{
        list.innerHTML = '<div style="color:#888888;padding:8px 0;">관련 뉴스가 없습니다</div>';
        return;
      }}
      list.innerHTML = '<div style="color:#7D8AA3;margin-bottom:6px;border-bottom:1px solid #DDE3EA;padding-bottom:4px;">📰 관련 뉴스</div>' +
        data.articles.map(a => `
          <div style="margin-bottom:8px;padding:6px;background:#F5F7FA;border-radius:4px;border-left:2px solid ${{cp.color}};">
            <div style="margin-bottom:2px;">
              <a href="${{a.link}}" target="_blank"
                style="color:#222222;text-decoration:none;font-size:11px;line-height:1.4;"
                onmouseover="this.style.color='#155BB8'" onmouseout="this.style.color='#222222'">
                ${{a.title}}
              </a>
            </div>
            ${{a.desc ? `<div style="color:#5d6b80;font-size:10px;margin-top:2px;line-height:1.3;">${{a.desc}}</div>` : ''}}
            <div style="color:#888888;font-size:10px;margin-top:3px;">${{a.date}} · ${{a.kw}}</div>
          </div>
        `).join('');
    }})
    .catch(() => {{
      const list = document.getElementById('choke-news-list');
      list.innerHTML = '<div style="color:#888888;">뉴스를 불러올 수 없습니다</div>';
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
            color: '#155BB8',
            weight: width * 2.6,
            opacity: 0.10,
            smoothFactor: 1,
          }}).addTo(_map);
          _routeLayers.push(glow);

          // Main route line — 점선 흐름 애니메이션 (CSS stroke-dashoffset)
          const line = L.polyline(pts, {{
            color: '#155BB8',
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
        fillColor: '#155BB8',
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
  _currentMineral = mineral; window._currentMineral = mineral;
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
{BACKDROP}
{V2_CHROME_FOOTER}
</body>
</html>"""


# ═══════════════════════════════════════════════════════════════
#  ③ 라우트
# ═══════════════════════════════════════════════════════════════
@app.route("/")
def index(): return Response(render_home_v2(), mimetype="text/html")   # V2 소비자 홈

@app.route("/pro")
def pro(): return Response(render_dashboard(home=False), mimetype="text/html")   # 전문가용 = 통계 대시보드

@app.route("/dashboard")
def dashboard(): return Response(render_dashboard(home=False), mimetype="text/html")  # 카테고리 화면 (검색 없음)

@app.route("/search")
def search(): return Response(render_search(request.args.get("q", "")), mimetype="text/html")

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

@app.route("/api/news-brief")
def news_brief():
    cat = request.args.get("cat", "minerals")
    ckey = f"news_brief_{cat}"
    c = cache_get(ckey)
    if c is not None:
        return jsonify(ok=bool(c), brief=c)
    if not OPENAI_API_KEY:
        return jsonify(ok=False, brief="")
    AUD_MAP = {"inv": "투자자", "biz": "기업", "con": "소비자", "pol": "정책"}
    if cat in AUD_MAP:
        aud = AUD_MAP[cat]
        items = [n for n in (fetch_audience_news() or []) if n.get("aud") == aud]
        AUD_SYS = {
            "투자자": "투자자 관점에서 오늘 뉴스가 어떤 광물·소재 섹터에 호재/악재인지 2문장으로 요약하고, "
                     "리스크 체크포인트를 한 줄 덧붙여라. 특정 종목 추천·매수매도 조언은 절대 금지, 섹터 흐름만.",
            "기업": "제조·조달 기업 관점에서 오늘 뉴스가 원자재 조달 비용·공급 안정성에 주는 영향을 2문장으로 요약하고, "
                   "조달 담당자가 점검할 액션 한 줄을 덧붙여라.",
            "소비자": "일반 소비자 관점에서 오늘 뉴스가 전기차·전자제품 등 체감 물가에 주는 영향을 아주 쉬운 말로 "
                     "2문장 요약하고, 생활 시사점 한 줄을 덧붙여라. 전문용어는 풀어서 써라.",
            "정책": "정책 연구자 관점에서 오늘 뉴스가 자원안보·비축·통상 정책에 주는 함의를 2문장으로 요약하고, "
                   "정책적 검토 포인트 한 줄을 덧붙여라.",
        }
        sysmsg = ("너는 핵심광물·공급망 애널리스트다. 아래 뉴스 헤드라인을 종합하라. "
                  + AUD_SYS[aud] + " 광물·자원과 무관한 얘기는 하지 마라. 전체 3문장 이내.")
    else:
        items = fetch_news()
        sysmsg = ("너는 핵심광물·공급망 애널리스트다. 아래 핵심광물(리튬·니켈·코발트·희토류 등) 뉴스 "
                  "헤드라인을 종합해 오늘의 광물 수급·공급망 핵심 흐름을 2문장으로 요약하고, "
                  "관련 산업(배터리·방산·반도체 소재 등) 관점 시사점을 한 줄 덧붙여라. "
                  "광물과 무관한 증시 일반·거시경제 얘기는 하지 마라. "
                  "특정 종목 추천이나 매수·매도 조언은 하지 말고 정보·교육 차원으로만. 전체 3문장 이내.")
    heads = [(n.get("제목") or n.get("title", "")) for n in items[:12] if (n.get("제목") or n.get("title"))]
    if not heads:
        cache_set(ckey, "", ttl=600); return jsonify(ok=False, brief="")
    brief = ""
    try:
        r = OpenAI(api_key=OPENAI_API_KEY).chat.completions.create(
            model=DEFAULT_OPENAI_MODEL, max_completion_tokens=500,
            messages=[
                {"role": "system", "content": sysmsg},
                {"role": "user", "content": "\n".join(heads)},
            ],
        )
        brief = (r.choices[0].message.content or "").strip()
    except Exception as e:
        print(f"[NEWS BRIEF {cat}] {e}")
    cache_set(ckey, brief, ttl=(1800 if brief else 60))
    return jsonify(ok=bool(brief), brief=brief)

@app.route("/subscribe", methods=["POST"])
def subscribe():
    body = request.get_json(silent=True) or {}
    email = (body.get("email") or "").strip().lower()
    if not valid_email(email): return jsonify(ok=False, message="올바른 이메일 형식이 아니에요.")
    valid_names = set(MIN_USES.keys())
    minerals = [m for m in (body.get("minerals") or []) if m in valid_names][:8]
    is_new = add_sub(email)
    set_sub_minerals(email, minerals)
    if not is_new:
        if minerals:
            return jsonify(ok=True, message=f"관심 광물 {len(minerals)}개로 업데이트했어요! 다음 리포트부터 반영돼요.")
        return jsonify(ok=False, message="이미 구독 중인 이메일이에요.")
    if SMTP_USER and SMTP_PASS:
        def _welcome(e=email, mi=minerals):
            try:
                ok, info = send_mail(e, "[K Mineral Risk] 구독 완료 — 매일 아침 광물 날씨를 보내드려요", build_welcome(e, mi))
                print("[welcome]", e, ok, info if not ok else "")
            except Exception as ex:
                print("[welcome] 오류:", ex)
        threading.Thread(target=_welcome, daemon=True).start()
        _mi = f" 관심 광물 {len(minerals)}개의 소식도 함께 담아드려요." if minerals else ""
        return jsonify(ok=True, message="구독 완료! 환영 메일을 보냈어요." + _mi)
    return jsonify(ok=True, message=f"구독 완료! 현재 {len(load_subs())}명이 구독 중이에요.")

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
    if request.args.get("force") != "1" and not _nl_claim_today():
        return jsonify(ok=True, message="오늘은 이미 발송했어요(중복 방지). force=1로 강제 가능")
    return jsonify(ok=True, **_send_daily_all())


def _newsletter_scheduler():
    """앱 내장 일일 발송 — 매일 NEWSLETTER_HOUR시(KST, 기본 9시) 이후 첫 기회에 1회.
    9시에 서버가 자고 있었어도 깨어난 시점에 밀린 발송을 처리(캐치업). DB 가드로 중복 방지."""
    hh = int(os.environ.get("NEWSLETTER_HOUR", "9") or 9)
    print(f"[newsletter] 자동 발송 대기 — 매일 {hh:02d}:00 KST (놓치면 깨어난 직후 발송)")
    time.sleep(30)
    while True:
        try:
            if _kst_now().hour >= hh and load_subs() and _nl_claim_today():
                _send_daily_all()
        except Exception as e:
            print("[newsletter] 스케줄러 오류:", e)
        time.sleep(240)


if os.environ.get("NEWSLETTER_AUTO", "1") == "1" and SMTP_USER and SMTP_PASS:
    threading.Thread(target=_newsletter_scheduler, daemon=True).start()

@app.route("/favicon.ico")
def favicon():
    from flask import send_file
    return send_file(os.path.join(os.path.dirname(__file__), "static", "favicon.ico"))

@app.route("/intro")
def intro():
    return Response(render_showcase(), mimetype="text/html")

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
    history = [{**h, "content": _strip_surrogates(h.get("content", ""))}
               for h in (data.get("history") or []) if isinstance(h, dict)]
    audience = data.get("audience", "consumer")
    recent_viz = data.get("recentViz", []) or []
    if not speaker or speaker not in MINERAL_EXPERTS:
        return jsonify(ok=False, message="발언할 전문가가 지정되지 않았습니다."), 400
    expert = MINERAL_EXPERTS[speaker]

    # 대상(청중)별 토론 맥락 — 같은 전문가라도 대상에 따라 토론이 달라진다
    AUDIENCE_CTX = {
        "investor": ("일반 투자자", "이 분석의 청중은 '일반 개인투자자'입니다. 해당 자원의 수급 리스크가 "
            "어떤 산업 섹터·테마(예: 2차전지, 방산, 정유·화학, 반도체, 식품주)에 호재/악재로 작용하는지 "
            "투자 관점에서 짚어주세요. 단, 특정 종목 추천이나 매수·매도 조언은 절대 하지 말고, "
            "'정보·교육 차원의 섹터 영향'으로만 설명하세요. (회의 종합 시 목표 산출물: 투자 포인트 3 + 리스크 체크리스트)"),
        "business": ("기업 조달·구매 담당", "청중은 '기업의 구매·조달 담당자'입니다. 대체 조달처 확보, 재고·비축 수준, "
            "장기계약·가격 헤지, 공급 차질 시 생산 영향 등 '실무 대응 전략' 중심으로 구체적으로 조언하세요. "
            "(회의 종합 시 목표 산출물: 리스크 히트맵 + 액션 아이템)"),
        "consumer": ("일반 소비자", "청중은 '일반 소비자'입니다. 전문용어는 풀어 쓰고, 이 이슈가 장바구니 물가·"
            "주유비·전기료 등 '생활에 미치는 영향'과 체감되는 숫자 중심으로 쉽고 친근하게 설명하세요. "
            "(회의 종합 시 목표 산출물: 3줄 요약 + 생활 Q&A)"),
        "policy": ("정책·연구자", "청중은 '정책 입안자·연구자'입니다. 국가 차원의 비축·국산화·외교·제도·전략 관점에서 "
            "근거와 사례를 들어 심도 있게 논하세요. (회의 종합 시 목표 산출물: 정책 권고안 + 우선순위 매트릭스)"),
    }
    aud_name, aud_ctx = AUDIENCE_CTX.get(audience, AUDIENCE_CTX["consumer"])

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

        # 이 전문가가 발언 근거로 띄울 수 있는 시각자료(차트) 목록 주입
        viz_block = ""
        try:
            _cat = build_viz_catalog()
            _my = [k for k in EXPERT_VIZ.get(speaker, []) if k in _cat]
            if _my:
                _lines = "\n".join(
                    f"- {k}: {_cat[k]['title']}"
                    + (f" — 최신: {_cat[k]['note']}" if _cat[k].get("note") else "")
                    for k in _my)
                viz_block = (
                    "\n\n[시각자료 — 근거 차트] 다음 자료를 화면에 띄울 수 있습니다:\n" + _lines +
                    "\n사용법: 발언에서 특정 수치·추세·자원·가격을 언급하고 위 목록에 관련 차트가 있으면, "
                    "발언 맨 마지막에 `[[viz:키]]` 한 줄을 붙여 적극적으로 보여주세요 "
                    "(예: [[viz:" + _my[0] + "]]). 데이터를 근거로 드는 발언이라면 대체로 하나 붙이는 게 좋습니다. "
                    "다만 ① 직전 발언과 똑같은 차트를 연속으로 반복하지 말고, ② 순수하게 동의·맥락·"
                    "감상만 말하는 발언에는 생략하세요. 한 발언에 차트는 최대 하나입니다."
                    + (f" 최근 화면에 이미 띄운 차트({', '.join(recent_viz)})는 다시 고르지 말고 다른 자료를 쓰거나 생략하세요." if recent_viz else "")
                )
        except Exception as _e:
            print("[VIZ prompt]", _e)

        sys_prompt = SHARED_A2A_PREAMBLE + "\n\n[당신의 역할]\n" + expert["system"] + (
            f"\n\n[대상 맞춤] {aud_ctx}"
            "\n\n[회의 형식] 이것은 여러 전문가와 진행자가 함께하는 실시간 회의입니다. "
            "아래 회의록을 읽고, 다른 전문가나 진행자의 발언을 직접 인용하며 동의하거나 반박한 뒤 "
            "자신의 핵심 의견을 200자 내외로 말하세요. 위 공통 규칙을 따르되, 특히 수치·사실 끝에는 "
            "[데이터셋명] 출처칩을 붙이세요. 이미 나온 말을 반복하지 말고 논의를 진전시키세요. "
            "발언 앞에 자신의 이름이나 '[이름]' 같은 라벨을 붙이지 말고, 바로 본문부터 말하세요."
        ) + viz_block + repeat_guard
        convo = transcript_text(history) or f"회의 주제: {topic}"
        # 진행자가 방금 끼어들어 질문했다면 — 그 질문에 대한 직접 답변이 최우선
        mod_q = ""
        if len(history) > 1 and history[-1].get("role") == "user":
            mod_q = (
                f"\n\n[진행자 질문 — 최우선] 방금 진행자가 이렇게 물었습니다: \"{history[-1].get('content', '')[:300]}\"\n"
                "이번 발언은 반드시 이 질문에 대한 직접적인 답으로 시작하세요. "
                "질문과 무관하게 하던 논지를 이어가거나 일반론을 말하지 마세요. "
                "위 [시각자료] 목록의 최신 수치가 질문과 관련 있으면 그 실제 수치·추세로 답하고 해당 차트를 띄우세요. "
                "정확한 수치가 없으면 없다고 말하고 확인 방법을 안내하세요."
            )
        user_prompt = (
            f"[회의 주제]\n{topic}\n\n"
            f"[지금까지의 회의록]\n{convo}"
            f"{mod_q}\n\n"
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


# 전문가별 OpenAI TTS 목소리 + 말투 (gpt-4o-mini-tts instructions)
EXPERT_VOICE = {
    "리튬":   ("nova",    "명료하고 자신감 있는 여성 연구원. 데이터를 짚어가며 또박또박, 속도는 빠르게."),
    "코발트": ("onyx",    "저음의 남성 분석가. 긴박하게 경고하는 톤, 말이 빠르고 힘있게."),
    "니켈":   ("echo",    "밝고 실용적인 남성. 낙관적이고 외교적인 톤, 경쾌하게 빠른 속도."),
    "희토류": ("sage",    "차분한 중년 교수 톤. 무게감 있지만 지루하지 않게, 보통보다 약간 빠르게."),
    "텅스텐": ("ash",     "단호한 남성. 안보 브리핑처럼 절도 있고 빠르게."),
    "망간":   ("verse",   "에너지 넘치는 기술 낙관주의자 남성. 활기차고 빠르게."),
    "흑연":   ("alloy",   "차분하고 꼼꼼한 톤. 기술 디테일을 명확하게, 약간 빠르게."),
    "경제":   ("shimmer", "명쾌한 여성 이코노미스트. 숫자를 강조하며 시원시원하게 빠른 속도."),
    "통상":   ("ballad",  "부드럽지만 논리적인 남성 협상가 톤, 리듬감 있게 빠르게."),
    "지정학": ("fable",   "극적인 스토리텔러 톤. 국제정치의 긴장감을 살려 빠르게."),
    "정책":   ("coral",   "실무적인 여성 정책가. 결론부터 명확하게, 빠른 속도."),
}
TTS_COMMON = ("실제 한국어 정책 토론회에서 열띤 공방 중인 패널처럼 말하세요. "
              "매우 빠른 속도로 — 시간에 쫓기는 생방송 토론자처럼 속사포로, 단 발음은 뭉개지 않게. "
              "문장 사이 쉼을 최소화하고 호흡과 강세를 살려 AI 낭독 티를 없애세요. "
              "괄호나 특수기호는 읽지 마세요. ")

@app.route("/api/conference/tts", methods=["POST"])
def conference_tts():
    """전문가 발언 → 사람 같은 음성 (OpenAI gpt-4o-mini-tts, 전문가별 목소리·말투)."""
    if not _conf_authed():
        return jsonify(ok=False, message="인증이 필요합니다."), 401
    if not OPENAI_API_KEY:
        return jsonify(ok=False, error="no_api_key"), 503
    data = request.get_json(silent=True) or {}
    text = _strip_surrogates(str(data.get("text", "")))[:800].strip()
    speaker = str(data.get("speaker", ""))
    if not text:
        return jsonify(ok=False, error="empty"), 400
    voice, style = EXPERT_VOICE.get(speaker, ("alloy", "자연스러운 한국어 구어체, 빠른 속도."))
    try:
        r = OpenAI(api_key=OPENAI_API_KEY).audio.speech.create(
            model="gpt-4o-mini-tts", voice=voice, input=text,
            instructions=TTS_COMMON + style, response_format="mp3")
        audio = getattr(r, "content", None) or r.read()
        return Response(audio, mimetype="audio/mpeg",
                        headers={"Cache-Control": "no-store"})
    except Exception as e:
        print("[TTS]", e)
        return jsonify(ok=False, error=str(e)), 502


@app.route("/api/conference/stt", methods=["POST"])
def conference_stt():
    """회의실 음성 인식 — Jarvis STT 사이드카 프록시 (WAV in, text out)."""
    if not _conf_authed():
        return jsonify(ok=False, message="인증이 필요합니다."), 401
    try:
        r = requests.post(JARVIS_STT_URL + "/stt", data=request.get_data(),
                          headers={"Content-Type": "audio/wav"}, timeout=90)
        return jsonify(r.json()), r.status_code
    except Exception:
        return jsonify(ok=False, error="stt_unavailable",
                       message="Jarvis STT 서버가 꺼져 있습니다 — ./jarvis_stt.sh 로 실행하세요."), 503


@app.route("/api/conference/summary", methods=["POST"])
def conference_summary():
    """회의 종료 시 서기 AI가 회의록을 요약 리포트로 정리한다."""
    if not _conf_authed():
        return jsonify(ok=False, message="인증이 필요합니다."), 401
    data = request.get_json(silent=True) or {}
    history = [{**h, "content": _strip_surrogates(h.get("content", ""))}
               for h in (data.get("history") or []) if isinstance(h, dict)]
    audience = data.get("audience", "consumer")
    if not any(h.get("role") == "assistant" for h in history):
        return jsonify(ok=False, message="요약할 전문가 발언이 없습니다."), 400
    AUD = {"investor": "일반 투자자", "business": "기업 조달·구매 담당",
           "consumer": "일반 소비자", "policy": "정책·연구자"}
    aud = AUD.get(audience, "일반 소비자")
    lines = []
    for h in history:
        who = "진행자" if h.get("role") == "user" else h.get("name", "전문가")
        lines.append(f"[{who}] {h.get('content', '')}")
    convo = "\n".join(lines)[-9000:]
    if not OPENAI_API_KEY:
        return jsonify(ok=False, message="API 키가 설정되지 않았습니다."), 500
    try:
        r = OpenAI(api_key=OPENAI_API_KEY).chat.completions.create(
            model=DEFAULT_OPENAI_MODEL,
            max_completion_tokens=700,
            messages=[
                {"role": "system", "content":
                    "너는 자원안보 전문가 회의의 서기다. 회의록을 읽고 한국어로 간결한 회의 요약 리포트를 작성한다. "
                    "수치를 인용할 때는 회의에서 나온 [데이터셋명] 출처 표기를 유지한다. "
                    "마크다운 굵게(**)나 헤딩(#)은 쓰지 않는다. 회의에서 나온 내용만 쓰고 새 주장을 지어내지 않는다."},
                {"role": "user", "content":
                    f"청중: {aud}\n\n[회의록]\n{convo}\n\n"
                    f"다음 형식으로 요약하라:\n"
                    f"한 줄 결론: (회의 안건에 대한 답을 40자 이내로 압축한 한 문장)\n"
                    f"■ 핵심 결론 (3줄)\n■ {aud} 관점 시사점 (번호 3~4개)\n"
                    f"■ 전문가 간 쟁점 (있으면 1~2개, 없으면 생략)\n■ 후속 확인이 필요한 데이터 (1~2개)"},
            ])
        return jsonify(ok=True, summary=(r.choices[0].message.content or "").strip())
    except Exception as e:
        return jsonify(ok=False, message=str(e)), 500


# ═══════════════════════════════════════════════════════════════
#  지정학 상황실 — 3D 지구본 (실시간 뉴스 → AI 지정학 이벤트)
# ═══════════════════════════════════════════════════════════════
GLOBE_CHOKES = [
    {"name": "말라카 해협", "lat": 2.5, "lng": 101.5, "risk": "critical"},
    {"name": "호르무즈 해협", "lat": 26.6, "lng": 56.5, "risk": "critical"},
    {"name": "밥엘만데브", "lat": 12.6, "lng": 43.4, "risk": "critical"},
    {"name": "수에즈 운하", "lat": 30.4, "lng": 32.4, "risk": "high"},
    {"name": "파나마 운하", "lat": 9.1, "lng": -79.7, "risk": "medium"},
    {"name": "희망봉", "lat": -34.4, "lng": 18.5, "risk": "medium"},
]
GLOBE_ROUTES_DEF = [
    ("칠레", "리튬·구리"), ("호주", "철광석·리튬"), ("인도네시아", "니켈"),
    ("중국", "희토류·흑연"), ("콩고민주공화국", "코발트"), ("사우디아라비아", "원유"),
    ("카타르", "LNG"), ("미국", "LNG·원유"), ("러시아", "유연탄"), ("브라질", "철광석"),
]

def render_globe():
    locs = _geo_locations()
    busan = [35.1, 129.04]
    routes = []
    for cn, res in GLOBE_ROUTES_DEF:
        k = cn.replace(" ", "")
        if k in locs:
            routes.append({"startLat": locs[k][0], "startLng": locs[k][1],
                           "endLat": busan[0], "endLng": busan[1],
                           "label": f"{cn} → 부산 · {res}"})
    kr = {}
    try:
        kr = compute_k_risk() or {}
    except Exception:
        pass
    kr_strip = " · ".join(
        f"{'🔴' if v['grade']=='위험' else ('🟡' if v['grade']=='주의' else '🟢')} {k} {v['score']:.0f}"
        for k, v in sorted(kr.items(), key=lambda x: -x[1]["score"])[:8])
    # 광물별 매장량 오버레이 (KOMIR 매장량 스냅샷 → 좌표 매핑)
    _rsv_raw = load_json(os.path.join(os.path.dirname(__file__), "reserves_data1.json")) or []
    _rsv = {}
    for _it in _rsv_raw:
        _nm = re.sub(r"[\s\(\)/·].*$", "", str(_it.get("name") or ""))
        _pts = []
        for _c in (_it.get("countries") or [])[:8]:
            _cn = (_c.get("c") or "").strip()
            _co = COUNTRY_COORDS.get(_cn) or COUNTRY_COORDS.get(_cn.replace(" ", ""))
            if _co:
                _pts.append({"c": _cn, "lat": _co[0], "lng": _co[1], "v": _c.get("v") or 0})
        if _nm and _pts:
            _rsv[_nm] = _pts
    _kmin = ["전체", "리튬", "니켈", "코발트", "텅스텐", "몰리브덴", "망간", "동", "알루미늄",
             "아연", "철", "흑연", "규소", "석탄"]
    _chips = "".join(
        f'<button class="mchip{" on" if m == "전체" else ""}" data-m="{m}">{m}</button>' for m in _kmin)
    PAGE = r"""<!DOCTYPE html>
<html lang="ko">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>핵심광물지도 — K Mineral Risk</title>
<link rel="icon" type="image/png" href="/static/favicon.png?v=3">
<link rel="stylesheet" href="https://cdn.jsdelivr.net/gh/orioncactus/pretendard@v1.3.9/dist/web/static/pretendard.min.css">
<script src="https://unpkg.com/globe.gl@2.34.4/dist/globe.gl.min.js"></script>
<style>
*{margin:0;padding:0;box-sizing:border-box}
body{background:#eef2f8;color:#16233c;font-family:Pretendard,'Apple SD Gothic Neo','Malgun Gothic',sans-serif;overflow:hidden}
#globeViz{position:fixed;inset:0}
.hd{position:fixed;top:0;left:0;right:0;z-index:10;display:flex;align-items:center;justify-content:space-between;
  padding:14px 22px;background:linear-gradient(180deg,rgba(238,242,248,.95),rgba(238,242,248,0));pointer-events:none}
.hd>*{pointer-events:auto}
.hd .brand{display:flex;align-items:center;gap:12px;text-decoration:none}
.hd .brand img{height:30px}
.hd .brand .t{font-size:15px;font-weight:900;color:#12325e;letter-spacing:.02em}
.hd .brand .s{font-size:10px;font-weight:700;color:#a97a12;letter-spacing:.24em;text-transform:uppercase}
.hd .right{display:flex;align-items:center;gap:10px}
.hd .clock{font-family:'JetBrains Mono',monospace;font-size:11px;color:#5d6b80}
.hd .clock b{color:#1e8e5a}
.hd a.btn{font-size:12px;font-weight:800;color:#16233c;text-decoration:none;border:1px solid #cdd6e4;
  border-radius:999px;padding:7px 15px;background:rgba(255,255,255,.85);backdrop-filter:blur(6px)}
.hd a.btn:hover{border-color:#c8931d;color:#8a6a10}
.panel{position:fixed;top:64px;right:16px;bottom:16px;width:min(380px,calc(100vw - 32px));z-index:9;display:flex;flex-direction:column;
  background:rgba(255,255,255,.92);border:1px solid #dbe2ec;border-radius:16px;backdrop-filter:blur(10px);overflow:hidden;box-shadow:0 14px 40px rgba(20,35,60,.14)}
.panel .ph{padding:14px 16px 10px;border-bottom:1px solid #e8edf4}
.panel .ph .t{font-size:13px;font-weight:900;color:#12325e}
.panel .ph .t span{color:#a97a12}
.panel .ph .s{font-size:10.5px;color:#7d8aa3;margin-top:3px}
.stats{display:flex;flex-wrap:wrap;gap:6px;padding:10px 16px;border-bottom:1px solid #e8edf4}
.stat{font-size:10.5px;font-weight:800;border-radius:999px;padding:4px 10px;border:1px solid #dbe2ec}
.feed{flex:1;overflow-y:auto;padding:10px 12px}
.feed::-webkit-scrollbar{width:5px}
.feed::-webkit-scrollbar-thumb{background:#cdd6e4;border-radius:3px}
.ev{border:1px solid #e3e8f0;border-left-width:3px;border-radius:11px;padding:10px 12px;margin-bottom:8px;cursor:pointer;
  background:#f9fafc;transition:.15s}
.ev:hover{background:#f1f4f9;transform:translateX(-2px)}
.ev .top{display:flex;align-items:center;gap:6px;margin-bottom:5px}
.ev .tag{font-size:9.5px;font-weight:900;border-radius:5px;padding:2px 7px;letter-spacing:.04em}
.ev .loc{font-size:10.5px;font-weight:700;color:#5d6b80}
.ev .dt{font-size:9.5px;color:#8a94a6;margin-left:auto}
.ev .ti{font-size:12.5px;font-weight:700;line-height:1.5;color:#16233c}
.ev .why{font-size:11px;line-height:1.6;color:#4a5872;margin-top:5px}
.ev .res{display:inline-block;font-size:9.5px;font-weight:800;color:#1c5cab;margin-top:5px}
.legend{position:fixed;left:16px;bottom:16px;z-index:9;background:rgba(255,255,255,.92);border:1px solid #dbe2ec;
  border-radius:13px;padding:12px 16px;backdrop-filter:blur(10px);font-size:11px;box-shadow:0 10px 26px rgba(20,35,60,.12)}
.legend .lt{font-size:10px;font-weight:900;letter-spacing:.12em;color:#7d8aa3;margin-bottom:8px}
.legend .li{display:flex;align-items:center;gap:8px;margin:4px 0;color:#42536e}
.legend .dot{width:9px;height:9px;border-radius:50%}
.legend .ln{width:16px;height:2px;background:#c8931d;border-radius:2px}
.krisk{position:fixed;left:16px;top:64px;z-index:9;max-width:calc(100vw - 440px);background:rgba(255,255,255,.92);
  border:1px solid #dbe2ec;border-radius:999px;padding:8px 16px;font-size:11px;font-weight:700;color:#42536e;backdrop-filter:blur(10px);box-shadow:0 8px 20px rgba(20,35,60,.1)}
.krisk b{color:#a97a12;margin-right:6px}
.loading{padding:40px 20px;text-align:center;color:#7d8aa3;font-size:12px;line-height:2}
@media(max-width:760px){.panel{top:auto;height:45vh}.krisk{display:none}}
</style>
</head>
<body>
<div style="position:fixed;top:0;left:0;right:0;z-index:80;display:flex;align-items:center;gap:18px;height:54px;padding:0 20px;background:rgba(10,25,48,.74);backdrop-filter:blur(10px);border-bottom:1px solid rgba(255,255,255,.08);font-family:'Pretendard Variable',Pretendard,-apple-system,sans-serif">
  <a href="/" style="color:#fff;font-weight:800;font-size:16px;text-decoration:none;display:flex;align-items:center;gap:7px"><span style="width:9px;height:9px;border-radius:50%;background:linear-gradient(135deg,#D97680,#4A7DD6);box-shadow:0 0 0 4px rgba(92,143,214,.18)"></span>K Mineral Risk</a>
  <nav style="display:flex;gap:4px">
    <a href="/" style="padding:6px 13px;border-radius:999px;color:#cfd8d2;font-size:13.5px;font-weight:650;text-decoration:none">홈</a>
    <a href="/globe" style="padding:6px 13px;border-radius:999px;background:rgba(92,143,214,.18);color:#B9CCEA;font-size:13.5px;font-weight:650;text-decoration:none">핵심광물지도</a>
    <a href="/briefing" style="padding:6px 13px;border-radius:999px;color:#cfd8d2;font-size:13.5px;font-weight:650;text-decoration:none">브리핑</a>
    <a href="/conference" style="padding:6px 13px;border-radius:999px;color:#cfd8d2;font-size:13.5px;font-weight:650;text-decoration:none">AI 회의</a>
  </nav>
</div>
<div id="globeViz"></div>

<div class="hd">
  <a class="brand" href="/"><span style="font-size:19px;font-weight:900;color:#12325e;letter-spacing:-.02em">K<em style="font-style:normal;color:#155BB8"> Mineral Risk</em></span>
    <span><span class="s">Critical Minerals Map</span><br><span class="t">핵심광물지도</span></span></a>
  <div class="right">
    <span class="clock" id="clock">— <b>● LIVE</b></span>
    <a class="btn" href="/dashboard?cat=minerals&sec=risk">🚦 리스크 신호등</a>
    <a class="btn" href="/conference">⚖️ AI 회의실</a>
    <a class="btn" href="/">🏠 허브 홈</a>
  </div>
</div>

<div class="krisk"><b>K-RISK</b> __KRSTRIP__</div>
<div id="mchips">__MCHIPS__
  <span class="mkey"><b style="color:#c8931d">●</b> 수입 루트 · <b style="color:#155BB8">●</b> 매장량</span>
</div>
<style>
#mchips{position:fixed;left:16px;top:112px;z-index:9;display:flex;gap:6px;flex-wrap:wrap;align-items:center;
  max-width:calc(100vw - 460px);font-family:Pretendard,sans-serif}
.mchip{border:1px solid #DDE3EA;background:rgba(255,255,255,.93);color:#555555;border-radius:999px;
  padding:6px 13px;font:650 12.5px Pretendard,sans-serif;cursor:pointer}
.mchip.on{background:#16305C;border-color:#16305C;color:#fff}
#mchips .mkey{font-size:11px;color:#555555;background:rgba(255,255,255,.88);border-radius:999px;padding:4px 10px}
#chokeInfo{position:fixed;left:16px;top:196px;z-index:9;display:none;background:rgba(255,255,255,.94);
  border:1px solid #DDE3EA;border-radius:14px;padding:12px 16px;width:250px;font-family:Pretendard,sans-serif;
  box-shadow:0 8px 24px rgba(22,48,92,.08)}
#chokeInfo .ct{font-size:12px;font-weight:800;color:#16305C;margin-bottom:8px}
#chokeInfo .cr{margin:7px 0}
#chokeInfo .cl{display:flex;justify-content:space-between;font-size:12px;color:#333B45;font-weight:650;margin-bottom:3px}
#chokeInfo .cb{height:6px;border-radius:999px;background:#EEF1F5;overflow:hidden}
#chokeInfo .cf{height:100%;border-radius:999px;background:#c8931d}
#chokeInfo .cnone{font-size:12px;color:#555555;line-height:1.5}
@media(max-width:760px){#mchips{display:none}#chokeInfo{display:none!important}}
</style>
<div id="chokeInfo"></div>

<div class="panel">
  <div class="ph">
    <div class="t">실시간 지정학 이벤트 <span>— AI 분류</span></div>
    <div class="s">뉴스를 AI가 읽고 한국 자원 공급망 영향 이벤트만 지구본에 표시합니다</div>
  </div>
  <div class="stats" id="stats"></div>
  <div class="feed" id="feed"><div class="loading">🤖<br>AI가 최신 지정학 뉴스를<br>분석하고 있습니다...</div></div>
</div>

<div class="legend">
  <div class="lt">LEGEND</div>
  <div class="li"><span class="dot" style="background:#ff4d4d"></span> 전쟁 · 군사 충돌</div>
  <div class="li"><span class="dot" style="background:#ff9f40"></span> 제재 · 수출통제</div>
  <div class="li"><span class="dot" style="background:#f2c94c"></span> 외교 긴장</div>
  <div class="li"><span class="dot" style="background:#ff7ab8"></span> 공급 차질</div>
  <div class="li"><span class="dot" style="background:#9b8cff"></span> 시위 · 불안</div>
  <div class="li"><span class="ln"></span> 한국 수입 루트</div>
  <div class="li"><span style="color:#ff6a6a">⚓</span> 해상 초크포인트</div>
</div>

<script>
const CHOKES = __CHOKE__;
const BUSAN = {name:'부산항 HUB', lat:35.1, lng:129.04, risk:'hub'};
const ROUTES = __ROUTES__;
const TYPE_COL = {'전쟁':'#ff4d4d','제재':'#ff9f40','외교':'#f2c94c','공급차질':'#ff7ab8','시위':'#9b8cff'};

setInterval(function(){
  var d = new Date(), p = function(n){ return String(n).padStart(2,'0'); };
  document.getElementById('clock').innerHTML = d.getFullYear()+'-'+p(d.getMonth()+1)+'-'+p(d.getDate())
    +' '+p(d.getHours())+':'+p(d.getMinutes())+':'+p(d.getSeconds())+' KST <b>● LIVE</b>';
}, 1000);

const globe = Globe()(document.getElementById('globeViz'))
  .globeImageUrl('https://unpkg.com/three-globe/example/img/earth-blue-marble.jpg')
  .bumpImageUrl('https://unpkg.com/three-globe/example/img/earth-topology.png')
  .backgroundColor('#eef2f8')
  .atmosphereColor('#7fa8d9')
  .atmosphereAltitude(0.18)
  .arcsData(ROUTES)
  .arcColor(function(){ return ['rgba(200,147,29,.9)','rgba(200,147,29,.25)']; })
  .arcStroke(0.45)
  .arcDashLength(0.45).arcDashGap(0.25).arcDashAnimateTime(2600)
  .arcLabel(function(d){ return '<div style="font-size:11px;font-weight:700">'+d.label+'</div>'; })
  .htmlElementsData(CHOKES.concat([BUSAN]))
  .htmlLat('lat').htmlLng('lng').htmlAltitude(0.012)
  .htmlElement(function(d){
    var col = d.risk==='hub' ? '#155BB8' : (d.risk==='critical' ? '#c03535' : (d.risk==='high' ? '#c2611a' : '#a97a12'));
    var el = document.createElement('div');
    el.innerHTML = (d.risk==='hub' ? '🇰🇷 ' : '⚓ ') + d.name
      + (d.badge ? '<br><span style="font-size:10px;background:rgba(255,255,255,.92);color:#8a6a10;border:1px solid #b5821088;border-radius:999px;padding:1px 8px;font-weight:800">'+d.badge+'</span>' : '');
    el.style.cssText = 'font-size:'+(d.risk==='hub'?'12px':'10.5px')+';font-weight:800;color:'+col
      +';text-shadow:0 0 5px rgba(255,255,255,.95),0 0 10px rgba(255,255,255,.85);white-space:nowrap;'
      +'transform:translate(-50%,-130%);pointer-events:none;font-family:Pretendard,sans-serif;';
    return el;
  })
  .pointOfView({lat: 25, lng: 100, altitude: 2.1});

globe.controls().autoRotate = false;   // 자동 회전 끔 — 마우스 조작 편의

// ── 핵심광물지도: 광물별 수입 루트 + 매장량 오버레이 ──
const RSV = __RSV__;
window._MODE = '전체';
function fmtV(v){ return v>=1e8 ? (v/1e8).toFixed(1)+'억' : (v>=1e4 ? Math.round(v/1e4)+'만' : Math.round(v).toLocaleString()); }
function applyEvents(evs){
  globe.arcsData(ROUTES).arcStroke(0.45);
  globe.pointsData(evs)
    .pointLat('lat').pointLng('lng')
    .pointColor(function(e){ return TYPE_COL[e.type] || '#fff'; })
    .pointAltitude(function(e){ return 0.012 * (e.sev||1); })
    .pointRadius(function(e){ return 0.32 + (e.sev||1) * 0.14; })
    .pointLabel(function(e){ return '<div style="max-width:240px;font-size:11.5px;line-height:1.5"><b style="color:'
      + (TYPE_COL[e.type]||'#fff') + '">[' + e.type + '] ' + e.loc + '</b><br>' + e.title + '</div>'; });
  globe.ringsData(evs.filter(function(e){ return (e.sev||0) >= 2; }))
    .ringLat('lat').ringLng('lng')
    .ringColor(function(e){ return function(t){ var c = TYPE_COL[e.type] || '#fff';
      return c + Math.round((1 - t) * 200).toString(16).padStart(2,'0'); }; })
    .ringMaxRadius(function(e){ return (e.sev||1) * 3.2; })
    .ringPropagationSpeed(1.6).ringRepeatPeriod(1100);
}
const CH_MAP = (function(){
  var m = {};
  var add = function(list, chokes){ list.forEach(function(c){ m[c] = chokes; }); };
  add(['사우디아라비아','아랍에미리트','아랍에미리트 연합','아랍에미리트연합','카타르','쿠웨이트','바레인','이란','이라크','오만'],
      ['호르무즈 해협','말라카 해협']);
  add(['독일','프랑스','영국','스페인','이탈리아','네덜란드','벨기에','폴란드','스웨덴','노르웨이','핀란드','오스트리아',
       '체코','스위스','튀르키예','터키','우크라이나','그리스','포르투갈','아일랜드','덴마크','이집트','모로코','알제리','튀니지'],
      ['수에즈 운하','밥엘만데브','말라카 해협']);
  add(['기니','가나','나이지리아','세네갈','코트디부아르','라이베리아','시에라리온','모리타니'],
      ['희망봉','말라카 해협']);
  add(['남아프리카공화국','남아공','마다가스카르','모잠비크','탄자니아','케냐','잠비아','짐바브웨','콩고민주공화국','나미비아','보츠와나','말라위'],
      ['말라카 해협']);
  add(['인도','스리랑카','파키스탄','방글라데시','미얀마'], ['말라카 해협']);
  add(['브라질','아르헨티나','우루과이'], ['희망봉']);
  add(['베네수엘라','콜롬비아','트리니다드토바고','쿠바'], ['파나마 운하']);
  return m;
})();
function computeChokes(rs){
  var tot = 0, acc = {};
  rs.forEach(function(r){
    tot += r.amount || 0;
    (CH_MAP[r.country] || []).forEach(function(ch){ acc[ch] = (acc[ch] || 0) + (r.amount || 0); });
  });
  var out = {};
  if(tot > 0){ Object.keys(acc).forEach(function(k){ var p = Math.round(acc[k] / tot * 100); if(p >= 1) out[k] = p; }); }
  return out;
}
function applyMineral(m, rs){
  var unit = (rs && rs._unit) || '';
  globe.ringsData([]);
  globe.arcsData(rs.map(function(r){
      return {startLat:r.lat, startLng:r.lng, endLat:35.1, endLng:129.04,
              label:r.country+' → 부산 · '+fmtV(r.amount)+(unit==='톤'?'t':'$')+' (상대점유 '+r.share+'%)', share:r.share};
    }))
    .arcStroke(function(d){ return 0.22 + (d.share||0)/100 * 1.0; });
  var pts = rs.map(function(r){ return {kind:'imp', lat:r.lat, lng:r.lng, c:r.country, v:r.amount, share:r.share, u:unit}; });
  (RSV[m]||[]).forEach(function(x){ pts.push({kind:'rsv', lat:x.lat, lng:x.lng, c:x.c, v:x.v}); });
  globe.pointsData(pts)
    .pointLat('lat').pointLng('lng')
    .pointColor(function(p){ return p.kind==='imp' ? '#c8931d' : '#155BB8'; })
    .pointAltitude(function(p){ return p.kind==='imp' ? 0.018 + (p.share||0)/100*0.05 : 0.012; })
    .pointRadius(function(p){ return p.kind==='imp' ? 0.4 + (p.share||0)/100*0.45 : 0.34; })
    .pointLabel(function(p){ return '<div style="font-size:11.5px;line-height:1.5"><b>'+p.c+'</b><br>'
      + (p.kind==='imp' ? '한국 수입 '+fmtV(p.v)+(p.u==='톤'?'t':'$') : '매장량 '+fmtV(p.v)+'t')+'</div>'; });
  var ch = computeChokes(rs);
  globe.htmlElementsData(CHOKES.map(function(c){
    var o = Object.assign({}, c);
    if(ch[c.name]) o.badge = m + ' ' + ch[c.name] + '% 통과';
    return o;
  }).concat([BUSAN]));
  var ci = document.getElementById('chokeInfo');
  var keys = Object.keys(ch).sort(function(a,b){ return ch[b]-ch[a]; });
  if(keys.length){
    ci.innerHTML = '<div class="ct">⚓ ' + m + ' 수입이 지나는 관문</div>' + keys.map(function(k){
      return '<div class="cr"><div class="cl"><span>'+k+'</span><span>'+ch[k]+'%</span></div>'
        + '<div class="cb"><div class="cf" style="width:'+ch[k]+'%"></div></div></div>';
    }).join('') + '<div class="cnone" style="margin-top:8px">해당 관문이 막히면 이 비중만큼 수입이 우회·지연될 수 있어요.</div>';
  } else {
    ci.innerHTML = '<div class="ct">⚓ ' + m + ' 수입이 지나는 관문</div>'
      + '<div class="cnone">주요 해협을 거의 지나지 않아요.<br>태평양 직항(미주·호주·중국·일본) 위주 수입입니다.</div>';
  }
  ci.style.display = 'block';
}
function setMineral(m){
  window._MODE = m;
  document.querySelectorAll('.mchip').forEach(function(b){ b.classList.toggle('on', b.dataset.m===m); });
  if(m==='전체'){
    globe.htmlElementsData(CHOKES.concat([BUSAN]));
    var ci = document.getElementById('chokeInfo');
    if(ci) ci.style.display = 'none';
    if(window._EVS) applyEvents(window._EVS);
    else { globe.arcsData(ROUTES).arcStroke(0.45); globe.pointsData([]); globe.ringsData([]); }
    return;
  }
  fetch('/api/trade-map?mineral='+encodeURIComponent(m)).then(function(r){return r.json();}).then(function(d){
    if(window._MODE!==m) return;
    var rs = (d && d.routes) || [];
    rs._unit = (d && d.unit) || '';
    applyMineral(m, rs);
  });
}
document.querySelectorAll('.mchip').forEach(function(b){ b.addEventListener('click', function(){ setMineral(b.dataset.m); }); });
window.addEventListener('resize', function(){ globe.width(innerWidth).height(innerHeight); });

fetch('/api/geo-events').then(function(r){ return r.json(); }).then(function(d){
  var evs = (d && d.events) || [];
  var feed = document.getElementById('feed');
  if(!evs.length){
    feed.innerHTML = '<div class="loading">표시할 이벤트가 없습니다.<br>잠시 후 새로고침해 주세요.</div>';
    return;
  }
  window._EVS = evs;
  if((window._MODE||'전체')==='전체') applyEvents(evs);
  // 통계 칩
  var cnt = {};
  evs.forEach(function(e){ cnt[e.type] = (cnt[e.type] || 0) + 1; });
  document.getElementById('stats').innerHTML = Object.keys(cnt).map(function(k){
    return '<span class="stat" style="color:' + (TYPE_COL[k]||'#fff') + ';border-color:' + (TYPE_COL[k]||'#fff') + '55">'
      + k + ' ' + cnt[k] + '</span>';
  }).join('') + '<span class="stat" style="color:#8fa0bd">총 ' + evs.length + '건</span>';
  // 피드
  feed.innerHTML = evs.map(function(e, i){
    var c = TYPE_COL[e.type] || '#fff';
    return '<div class="ev" style="border-left-color:' + c + '" onclick="focusEv(' + i + ')">'
      + '<div class="top"><span class="tag" style="background:' + c + '22;color:' + c + '">' + e.type + ' · S' + e.sev + '</span>'
      + '<span class="loc">📍 ' + e.loc + '</span><span class="dt">' + (e.date || '').slice(5, 16) + '</span></div>'
      + '<div class="ti">' + e.title + '</div>'
      + (e.why ? '<div class="why">' + e.why + '</div>' : '')
      + (e.res ? '<span class="res">⛏ 관련 자원 — ' + e.res + '</span>' : '')
      + '</div>';
  }).join('');
  window._EVS = evs;
});
function focusEv(i){
  var e = window._EVS[i]; if(!e) return;
  globe.pointOfView({lat: e.lat, lng: e.lng, altitude: 1.3}, 900);
}
</script>
</body>
</html>"""
    return (PAGE.replace("__CHOKE__", json.dumps(GLOBE_CHOKES, ensure_ascii=False))
                .replace("__ROUTES__", json.dumps(routes, ensure_ascii=False))
                .replace("__KRSTRIP__", kr_strip or "계산 중")
                .replace("__MCHIPS__", _chips)
                .replace("__RSV__", json.dumps(_rsv, ensure_ascii=False)))

@app.route("/globe")
def globe_page():
    return Response(render_globe(), mimetype="text/html")


# ═══════════════════════════════════════════════════════════════
def render_search(q):
    import html as _html
    qq = (q or "").strip()
    qe = _html.escape(qq)
    news = fetch_search_news(qq) if qq else []

    # ── 보유 데이터에서 검색어 관련 정보 싹 끌어오기 (통합검색) ──
    def _card(title, rows, lt, lu):
        rs = "".join(f'<div class="drow"><span>{_html.escape(str(k))}</span><b>{_html.escape(str(v))}</b></div>' for k, v in rows)
        return f'<div class="dcard"><div class="dct">{_html.escape(title)}</div>{rs}<a class="dlink" href="{lu}">{_html.escape(lt)} →</a></div>'
    data_blocks = []
    if qq:
        mineral = next((m for m in USGS_DATA if m in qq), None)
        if mineral:
            u = USGS_DATA[mineral]
            data_blocks.append(_card(f"📊 매장량 · {mineral}", [
                ("USGS 매장량", f"{u['매장량_만톤']:,} 만t"), ("연 생산량", f"{u['생산량_만톤']:,} 만t"),
                ("1위 생산국", u['1위국']), ("출처", u.get('출처', 'USGS 2025'))],
                "글로벌 매장량 보기", f"/dashboard?cat=minerals&min={mineral}"))
            rk = next((r for r in load_risk_data() if r.get('name') == mineral), None)
            if rk:
                lv = rk['latest']; sig = '🟢 안정' if lv >= 55 else ('🟡 주의' if lv >= 30 else '🔴 위험')
                data_blocks.append(_card(f"🚦 공급 리스크 · {mineral}", [
                    ("수급안정화지수", f"{lv:.1f} / 100"), ("신호", sig),
                    ("전월 대비", f"{lv - rk.get('prev', lv):+.1f}")],
                    "리스크 신호등", "/dashboard?cat=minerals&sec=risk"))
            cust = [r for r in fetch_customs() if mineral in str(r.get('광물명', ''))]
            if cust:
                cust.sort(key=lambda r: r.get('수입금액(달러)', 0) or 0, reverse=True)
                tot = sum((r.get('수입금액(달러)', 0) or 0) for r in cust)
                rows = [("총 수입액", f"${tot:,.0f}")] + [(r.get('국가명', '—'), f"${(r.get('수입금액(달러)', 0) or 0):,.0f}") for r in cust[:3]]
                data_blocks.append(_card(f"💰 수입 현황 · {mineral}", rows, "광물 수급 현황", "/dashboard?cat=minerals&sec=supply"))
    data_html = "".join(data_blocks) or '<div class="empty" style="padding:24px">보유 데이터에 직접 매칭되는 항목이 없어요 — 아래 뉴스를 확인하세요.</div>'

    sc = []
    for m in ["리튬", "코발트", "니켈", "흑연", "희토류", "망간"]:
        if m in qq: sc.append((f"🔩 {m} · 글로벌 매장량", f"/dashboard?cat=minerals&min={m}"))
    if not sc:
        sc = [("🔩 핵심광물", "/dashboard?cat=minerals")]
    sc_html = "".join(f'<a class="sc" href="{u}">{_html.escape(t)}</a>' for t, u in sc)
    news_html = "".join(
        f'<a class="rcard" href="{_html.escape(n["링크"])}" target="_blank" rel="noopener">'
        f'<div class="rt">{_html.escape(n["제목"])}</div><div class="rs">{_html.escape(n["요약"])}</div>'
        f'<div class="rd">{_html.escape(n["발행일"])}</div></a>' for n in news)
    if not news_html:
        news_html = '<div class="empty">' + ('검색 결과가 없습니다. 다른 키워드로 시도해보세요.' if qq else '검색어를 입력하세요.') + '</div>'
    PAGE = r"""<!DOCTYPE html><html lang="ko"><head>
<meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>__Q__ 검색 — K Mineral Risk</title>
<link rel="icon" type="image/png" href="/static/favicon.png?v=3">
<link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;600;700;800;900&family=Noto+Sans+KR:wght@400;500;700;900&family=Noto+Serif+KR:wght@500;700&display=swap" rel="stylesheet">
<link rel="stylesheet" href="https://cdn.jsdelivr.net/gh/orioncactus/pretendard@v1.3.9/dist/web/static/pretendard.min.css">
<style>
:root{--sans:'Pretendard','Noto Sans KR',-apple-system,sans-serif;--mono:'Pretendard','Noto Sans KR',-apple-system,sans-serif;}
html body, html body *:not([class*="material-symbols"]){font-family:'Pretendard','Noto Sans KR',-apple-system,BlinkMacSystemFont,sans-serif !important;}
body{font-variant-numeric:tabular-nums;}
</style>
<style>
*{margin:0;padding:0;box-sizing:border-box}
body{background:linear-gradient(180deg,#f7f9fc 0%,#f2f5fa 100%) fixed;color:#16233c;font-family:'Inter','Noto Sans KR',sans-serif;min-height:100vh}
.top{display:flex;align-items:center;gap:18px;padding:16px 6vw;border-bottom:1px solid #e3e8f0;position:sticky;top:0;background:rgba(255,255,255,.92);backdrop-filter:blur(10px);z-index:10}
.brand{display:flex;align-items:center;gap:10px;text-decoration:none;flex-shrink:0}
.bmark{width:32px;height:32px;border-radius:9px;background:linear-gradient(135deg,#e9c667,#c8931d);color:#fff;display:flex;align-items:center;justify-content:center;font-weight:900}
.bname{font-weight:900;letter-spacing:.05em;color:#12325e;font-size:15px}
form.sbar{flex:1;max-width:640px;display:flex;align-items:center;gap:10px;background:#fff;border:1px solid #dbe2ec;border-radius:13px;padding:8px 10px 8px 18px;box-shadow:0 2px 8px rgba(20,35,60,.06)}
form.sbar input{flex:1;border:0;outline:0;font-size:14px;color:#222}
form.sbar button{width:38px;height:38px;border:0;border-radius:10px;background:linear-gradient(135deg,#e9c667,#c8931d);color:#fff;font-size:15px;cursor:pointer}
.home{margin-left:auto;color:#4a5872;text-decoration:none;font-size:13px;font-weight:700;border:1px solid #cdd6e4;padding:7px 14px;border-radius:20px;white-space:nowrap}
.home:hover{color:#12325e;border-color:#c8931d}
.wrap{max-width:1100px;margin:0 auto;padding:34px 6vw 70px}
.qh{font-family:'Noto Serif KR',serif;font-size:clamp(22px,3vw,32px);font-weight:700;margin-bottom:6px;color:#12325e}
.qh b{color:#a97a12}
.qsub{color:#7d8aa3;font-size:13px;margin-bottom:26px}
.sclabel{font-size:11px;letter-spacing:.3em;text-transform:uppercase;color:#a97a12;font-weight:700;margin-bottom:12px}
.scs{display:flex;flex-wrap:wrap;gap:10px;margin-bottom:34px}
.sc{text-decoration:none;color:#42536e;background:#fdf8ea;border:1px solid #eddfb5;border-radius:24px;padding:9px 17px;font-size:13px;font-weight:600;transition:.18s}
.sc:hover{background:#faf0d3;color:#12325e;transform:translateY(-2px)}
.rgrid{display:grid;grid-template-columns:repeat(auto-fill,minmax(300px,1fr));gap:14px}
.rcard{display:block;text-decoration:none;background:#fff;border:1px solid #e3e8f0;border-radius:15px;padding:18px 20px;transition:.2s;box-shadow:0 1px 3px rgba(20,35,60,.05)}
.rcard:hover{border-color:#c8931d;transform:translateY(-3px);box-shadow:0 14px 32px rgba(20,35,60,.12)}
.rt{font-size:15px;font-weight:700;color:#16233c;line-height:1.45;margin-bottom:8px}
.rs{font-size:13px;color:#5d6b80;line-height:1.6;margin-bottom:10px;display:-webkit-box;-webkit-line-clamp:3;-webkit-box-orient:vertical;overflow:hidden}
.rd{font-size:11px;color:#8a94a6;font-family:monospace}
.empty{padding:60px;text-align:center;color:#8a94a6;grid-column:1/-1}
.dgrid{display:grid;grid-template-columns:repeat(auto-fill,minmax(260px,1fr));gap:14px;margin-bottom:38px}
.dcard{background:#fff;border:1px solid #e8dfc2;border-radius:16px;padding:18px 20px;display:flex;flex-direction:column;box-shadow:0 1px 3px rgba(20,35,60,.05)}
.dct{font-size:14px;font-weight:800;color:#12325e;margin-bottom:12px}
.drow{display:flex;justify-content:space-between;align-items:center;gap:10px;padding:7px 0;border-bottom:1px solid #eef1f6;font-size:13px}
.drow span{color:#7d8aa3}.drow b{color:#16233c;font-family:monospace;font-weight:700}
.dlink{margin-top:13px;font-size:12px;font-weight:700;color:#a97a12;text-decoration:none}
.dlink:hover{color:#8a6a10}
</style></head><body>
<div class="top">
  <a class="brand" href="/"><span style="font-size:18px;font-weight:900;color:#12325e;letter-spacing:-.02em">K<em style="font-style:normal;color:#155BB8"> Mineral Risk</em></span></a>
  <form class="sbar" action="/search" method="get">
    <input name="q" value="__Q__" placeholder="광물·품목·키워드를 검색하세요" autofocus>
    <button type="submit">🔍</button>
  </form>
  <a class="home" href="/">🏠 메인</a>
</div>
<div class="wrap">
  <div class="qh"><b>__Q__</b> 검색 결과</div>
  <div class="qsub">뉴스 __N__건 · 네이버 뉴스 기준</div>
  <div class="sclabel">관련 바로가기</div>
  <div class="scs">__SHORTCUTS__</div>
  <div class="sclabel">보유 데이터</div>
  <div class="dgrid">__DATA__</div>
  <div class="sclabel">뉴스</div>
  <div class="rgrid">__RESULTS__</div>
</div>
</body></html>"""
    return (PAGE.replace("__Q__", qe).replace("__N__", str(len(news)))
                .replace("__SHORTCUTS__", sc_html).replace("__DATA__", data_html).replace("__RESULTS__", news_html))


#  ④ AI 전문가 회의실 페이지
# ═══════════════════════════════════════════════════════════════
def render_showcase():
    return r"""<!DOCTYPE html>
<html lang="ko"><head>
<meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>K Mineral Risk — 핵심광물</title>
<link rel="icon" type="image/png" href="/static/favicon.png?v=3">
<link href="https://fonts.googleapis.com/css2?family=Cormorant+Garamond:ital,wght@0,300;0,400;0,500;0,600;1,400&family=Noto+Serif+KR:wght@300;500;700&family=Inter:wght@400;500;700&display=swap" rel="stylesheet">
<link rel="stylesheet" href="https://cdn.jsdelivr.net/gh/orioncactus/pretendard@v1.3.9/dist/web/static/pretendard.min.css">
<style>
:root{--sans:'Pretendard','Noto Sans KR',-apple-system,sans-serif;--mono:'Pretendard','Noto Sans KR',-apple-system,sans-serif;}
html body, html body *:not([class*="material-symbols"]){font-family:'Pretendard','Noto Sans KR',-apple-system,BlinkMacSystemFont,sans-serif !important;}
body{font-variant-numeric:tabular-nums;}
</style>
<style>
*{margin:0;padding:0;box-sizing:border-box}
html{scroll-behavior:smooth}
body{background:#f6f4ee;color:#1f2740;font-family:'Inter','Noto Sans KR',sans-serif;overflow-x:hidden}
#cine{position:fixed;inset:0;z-index:0;display:block}
.grain{position:fixed;inset:0;z-index:1;pointer-events:none;opacity:.05;
  background-image:radial-gradient(rgba(255,255,255,.6) .5px,transparent .5px);background-size:3px 3px;mix-blend-mode:overlay}
.wrap{position:relative;z-index:2}
#ghost{position:fixed;left:50%;top:50%;transform:translate(-50%,-50%);z-index:1;pointer-events:none;
  font-family:'Cormorant Garamond',serif;font-weight:600;white-space:nowrap;letter-spacing:.04em;
  font-size:26vw;line-height:1;color:transparent;-webkit-text-stroke:1px rgba(160,120,40,.16);
  opacity:0;transition:opacity .7s ease}
.brand{position:fixed;top:26px;left:34px;z-index:5;font-family:'Cormorant Garamond',serif;font-size:18px;letter-spacing:.42em;color:#a97a12;text-transform:uppercase}
.prog{position:fixed;top:0;left:0;height:2px;width:0;z-index:6;background:linear-gradient(90deg,#caa24e,#f4e3ad)}
.act{min-height:100vh;display:flex;flex-direction:column;justify-content:center;padding:0 9vw}
.eyebrow{font-size:11px;letter-spacing:.5em;text-transform:uppercase;color:#a97a12;font-weight:600;margin-bottom:28px}
.title{font-family:'Noto Serif KR',serif;font-weight:700;font-size:clamp(40px,8.4vw,108px);line-height:1.03;letter-spacing:-.01em}
.gold{background:linear-gradient(118deg,#d9a521,#a97a12 58%,#8a6708);-webkit-background-clip:text;background-clip:text;color:transparent}
.copy{margin-top:28px;max-width:540px;font-size:clamp(15px,1.5vw,19px);line-height:1.95;color:#5d6377;font-weight:400}
.actno{font-family:'Cormorant Garamond',serif;font-style:italic;font-size:clamp(17px,2.1vw,24px);letter-spacing:.36em;color:#a97a12;text-transform:uppercase;margin-bottom:20px}
.reveal{opacity:0;transform:translateY(36px);transition:opacity 1.2s cubic-bezier(.2,.7,.2,1),transform 1.2s cubic-bezier(.2,.7,.2,1)}
.in .reveal{opacity:1;transform:none}
.in .reveal.d1{transition-delay:.12s}.in .reveal.d2{transition-delay:.26s}.in .reveal.d3{transition-delay:.4s}
.hint{position:fixed;bottom:30px;left:0;right:0;text-align:center;z-index:5;color:#8b8574;font-size:10px;letter-spacing:.4em;animation:bob 2s ease-in-out infinite;transition:opacity .4s}
@keyframes bob{0%,100%{transform:translateY(0)}50%{transform:translateY(7px)}}
/* 입장(finale) */
.enter-list{margin-top:42px;max-width:680px;width:100%}
.ecat{display:flex;align-items:baseline;gap:20px;padding:24px 4px;border-top:1px solid rgba(201,162,78,.2);
  cursor:pointer;text-decoration:none;color:inherit;transition:padding .35s cubic-bezier(.2,.7,.2,1)}
.enter-list .ecat:last-child{border-bottom:1px solid rgba(201,162,78,.2)}
.ecat .no{font-family:'Cormorant Garamond',serif;font-style:italic;color:#a97a12;font-size:20px;width:42px;flex-shrink:0}
.ecat .nm{font-family:'Noto Serif KR',serif;font-size:clamp(23px,3vw,36px);font-weight:500;flex:1;transition:color .3s}
.ecat .tg{font-size:12px;color:#8b8574;letter-spacing:.06em;flex-shrink:0}
.ecat .ar{color:#a97a12;font-size:22px;transition:transform .3s;flex-shrink:0}
.ecat:hover{padding-left:20px}.ecat:hover .nm{color:#8a6708}.ecat:hover .ar{transform:translateX(9px)}
.ecat.soon{cursor:default;opacity:.4}.ecat.soon:hover{padding-left:4px}.ecat.soon:hover .nm{color:inherit}
#veil{position:fixed;inset:0;background:#f6f4ee;z-index:60;opacity:0;pointer-events:none;transition:opacity .6s ease}
#veil.on{opacity:1}
</style></head><body>
<canvas id="cine"></canvas>
<div id="ghost"></div>
<div class="grain"></div>
<div class="prog" id="prog"></div>
<div class="brand">M I N E T E C H</div>
<div class="hint" id="hint">SCROLL ↓</div>
<div id="veil"></div>

<div class="wrap">
  <section class="act"><div class="eyebrow reveal">CRITICAL MINERALS</div>
    <h1 class="title reveal d1">자원이 세상을<br><span class="gold">움직인다</span></h1>
    <p class="copy reveal d2">보이지 않는 흐름이 당신의 투자와 산업을 바꾼다.<br>땅속의 권력, 핵심광물의 시대.</p>
  </section>

  <section class="act"><div class="actno reveal">제 1 막 · Act I</div>
    <h1 class="title reveal d1">땅속의 <span class="gold">권력</span></h1>
    <p class="copy reveal d2">리튬·희토류·니켈 — 배터리와 반도체의 심장.<br>한 나라의 곳간이 세계의 공급망을 흔든다.</p>
  </section>

  <section class="act"><div class="eyebrow reveal">ENTER · 입장</div>
    <h1 class="title reveal d1">이제, <span class="gold">당신의 차례</span></h1>
    <div class="enter-list reveal d2">
      <a class="ecat" href="/dashboard?cat=minerals"><span class="no">01</span><span class="nm">핵심광물</span><span class="tg">리튬·희토류</span><span class="ar">→</span></a>
    </div>
  </section>
</div>

<script type="importmap">{"imports":{"three":"https://unpkg.com/three@0.160.0/build/three.module.js","three/addons/":"https://unpkg.com/three@0.160.0/examples/jsm/"}}</script>
<script type="module">
import * as THREE from 'three';
import { EffectComposer } from 'three/addons/postprocessing/EffectComposer.js';
import { RenderPass } from 'three/addons/postprocessing/RenderPass.js';
import { UnrealBloomPass } from 'three/addons/postprocessing/UnrealBloomPass.js';

const canvas=document.getElementById('cine');
const renderer=new THREE.WebGLRenderer({canvas,antialias:true,alpha:true});
renderer.setPixelRatio(Math.min(devicePixelRatio,1.6)); renderer.setSize(innerWidth,innerHeight);
const scene=new THREE.Scene();
const camera=new THREE.PerspectiveCamera(50,innerWidth/innerHeight,.1,100); camera.position.z=7;

const earth=new THREE.Group(); scene.add(earth);
const TX='https://threejs.org/examples/textures/planets/';
const tl=new THREE.TextureLoader(); const R=2.1;
const globe=new THREE.Mesh(new THREE.SphereGeometry(R,64,64), new THREE.MeshPhongMaterial({
  map:tl.load(TX+'earth_atmos_2048.jpg'), specularMap:tl.load(TX+'earth_specular_2048.jpg'),
  normalMap:tl.load(TX+'earth_normal_2048.jpg'), normalScale:new THREE.Vector2(.8,.8),
  specular:new THREE.Color(0x3a2f16), shininess:9,
  emissiveMap:tl.load(TX+'earth_lights_2048.png'), emissive:new THREE.Color(0xffcf66), emissiveIntensity:1.05 }));
earth.add(globe);
const clouds=new THREE.Mesh(new THREE.SphereGeometry(R*1.01,48,48),
  new THREE.MeshPhongMaterial({map:tl.load(TX+'earth_clouds_1024.png'),transparent:true,opacity:.32,depthWrite:false}));
earth.add(clouds);
const atmo=new THREE.Mesh(new THREE.SphereGeometry(R*1.16,48,48), new THREE.ShaderMaterial({
  vertexShader:'varying vec3 vN;void main(){vN=normalize(normalMatrix*normal);gl_Position=projectionMatrix*modelViewMatrix*vec4(position,1.0);}',
  fragmentShader:'varying vec3 vN;void main(){float i=pow(0.74-dot(vN,vec3(0.,0.,1.)),3.0);gl_FragColor=vec4(0.91,0.76,0.29,1.0)*clamp(i,0.,1.);}',
  blending:THREE.AdditiveBlending, side:THREE.BackSide, transparent:true, depthWrite:false }));
earth.add(atmo);

scene.add(new THREE.AmbientLight(0x141414,.5));
const key=new THREE.DirectionalLight(0xffe6b0,3.3); key.position.set(-4.5,1.6,3); scene.add(key);

// 지구 뒤 골드 광휘(halo)
function glowTex(){ const cv=document.createElement('canvas'); cv.width=cv.height=128; const x=cv.getContext('2d');
  const g=x.createRadialGradient(64,64,0,64,64,64); g.addColorStop(0,'rgba(233,200,110,.9)'); g.addColorStop(.3,'rgba(201,162,78,.35)'); g.addColorStop(1,'rgba(201,162,78,0)');
  x.fillStyle=g; x.fillRect(0,0,128,128); return new THREE.CanvasTexture(cv); }
const halo=new THREE.Sprite(new THREE.SpriteMaterial({map:glowTex(),blending:THREE.AdditiveBlending,transparent:true,depthWrite:false,depthTest:false}));
halo.scale.setScalar(8.5); halo.position.z=-2; earth.add(halo);

// 떠오르는 골드 입자(embers)
const EN=320, ep=new Float32Array(EN*3), ev=new Float32Array(EN);
for(let i=0;i<EN;i++){ ep[i*3]=(Math.random()-.5)*16; ep[i*3+1]=(Math.random()-.5)*14; ep[i*3+2]=(Math.random()-.5)*6-1; ev[i]=.002+Math.random()*.006; }
const eg=new THREE.BufferGeometry(); eg.setAttribute('position',new THREE.BufferAttribute(ep,3));
const embers=new THREE.Points(eg,new THREE.PointsMaterial({color:0xe9c86e,size:.045,transparent:true,opacity:.7,blending:THREE.AdditiveBlending,depthWrite:false}));
scene.add(embers);

const composer=new EffectComposer(renderer);
composer.addPass(new RenderPass(scene,camera));
composer.addPass(new UnrealBloomPass(new THREE.Vector2(innerWidth,innerHeight),.72,.65,.5));

let sp=0;
const GWORDS=['RESOURCE','MINERAL','ENTER']; let _gi=-1;
const ghost=document.getElementById('ghost');
function onScroll(){ const h=document.body.scrollHeight-innerHeight; sp=h>0?scrollY/h:0;
  document.getElementById('prog').style.width=(sp*100)+'%';
  document.getElementById('hint').style.opacity = sp>.03?0:1;
  // 거대 키네틱 배경 단어 — 막마다 교체 + 패럴럭스 드리프트
  const f=sp*(GWORDS.length-1), gi=Math.round(f);
  if(gi!==_gi){ _gi=gi; ghost.style.opacity=0;
    setTimeout(()=>{ ghost.textContent=GWORDS[gi]; ghost.style.opacity=1; },180); }
  ghost.style.transform='translate(-50%,-50%) translateX('+((f-gi)*-22).toFixed(1)+'vw)'; }
addEventListener('scroll',onScroll); onScroll();
addEventListener('resize',()=>{ camera.aspect=innerWidth/innerHeight; camera.updateProjectionMatrix();
  renderer.setSize(innerWidth,innerHeight); composer.setSize(innerWidth,innerHeight); });

const io=new IntersectionObserver(es=>es.forEach(e=>{ if(e.isIntersecting) e.target.classList.add('in'); }),{threshold:.32});
document.querySelectorAll('.act').forEach(s=>io.observe(s));
document.querySelectorAll('.ecat[href]').forEach(a=>a.addEventListener('click',e=>{
  e.preventDefault(); const v=document.getElementById('veil'); v.classList.add('on');
  setTimeout(()=>{ location.href=a.getAttribute('href'); },620);
}));

const clock=new THREE.Clock();
function animate(){ requestAnimationFrame(animate); if(document.hidden) return;
  const t=clock.getElapsedTime();
  earth.rotation.y=t*.03; clouds.rotation.y=t*.012;
  const e=sp*sp*(3-2*sp);
  earth.position.x = e*2.7; earth.position.y = e*1.15; earth.scale.setScalar(1 - e*0.46);
  // embers 상승 + 래핑
  const pos=embers.geometry.attributes.position;
  for(let i=0;i<EN;i++){ let y=pos.getY(i)+ev[i]; if(y>7){ y=-7; } pos.setY(i,y); }
  pos.needsUpdate=true; embers.rotation.y=t*.01;
  composer.render();
}
animate();
</script>
</body></html>"""


def render_login(err=""):
    PAGE = r"""<!DOCTYPE html>
<html class="dark" lang="ko">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>로그인 — AI 전문가 회의실</title>
<link rel="icon" type="image/png" href="/static/favicon.png?v=3">
<script src="https://cdn.tailwindcss.com"></script>
<link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;600;700;900&family=Noto+Sans+KR:wght@400;700;900&display=swap" rel="stylesheet">
<link rel="stylesheet" href="https://cdn.jsdelivr.net/gh/orioncactus/pretendard@v1.3.9/dist/web/static/pretendard.min.css">
<style>
:root{--sans:'Pretendard','Noto Sans KR',-apple-system,sans-serif;--mono:'Pretendard','Noto Sans KR',-apple-system,sans-serif;}
html body, html body *:not([class*="material-symbols"]){font-family:'Pretendard','Noto Sans KR',-apple-system,BlinkMacSystemFont,sans-serif !important;}
body{font-variant-numeric:tabular-nums;}
</style>
<style>body{background:#F5F7FA;color:#222222;font-family:'Pretendard Variable',Pretendard,'Noto Sans KR',sans-serif;}</style>
</head>
<body class="min-h-screen flex items-center justify-center p-6">
  <form method="POST" action="/conference/login" class="w-full max-w-sm bg-white border border-[#e3e8f0] rounded-2xl p-8" style="box-shadow:0 14px 40px rgba(20,35,60,.1)">
    <div class="flex items-center gap-3 mb-6">
      <div class="w-10 h-10 rounded-lg flex items-center justify-center font-black text-xl" style="background:#12325e;color:#f4e3ad">M</div>
      <div>
        <div class="font-black text-[#12325e] tracking-wider">K MINERAL RISK AI</div>
        <div class="text-[10px] uppercase tracking-widest text-[#8a94a6]">AI 전문가 회의실</div>
      </div>
    </div>
    <p class="text-sm text-[#4a5872] mb-5">이 회의실은 비밀번호로 보호되어 있습니다.</p>
    <input type="password" name="password" autofocus placeholder="비밀번호" class="w-full bg-[#f7f9fc] border border-[#cdd6e4] rounded-lg px-4 py-3 text-sm outline-none focus:border-[#c8931d] mb-3">
    <div class="text-[#c03535] text-xs mb-3" style="min-height:16px">__ERR__</div>
    <button type="submit" class="w-full font-bold py-3 rounded-lg" style="background:#12325e;color:#fff">입장하기</button>
    <a href="/" class="block text-center text-xs text-[#8a94a6] mt-4 hover:text-[#a97a12]">← 대시보드로</a>
  </form>
</body>
</html>"""
    return PAGE.replace("__ERR__", err)


# ── AI 회의실 시각자료(차트) 카탈로그 ──────────────────────────────
# 전문가가 발언하면서 근거로 띄울 수 있는 차트 스펙을 실데이터로 생성.
# 각 항목: {title, type(line|bar|doughnut|stat), labels, series|stats, note, source}
EXPERT_VIZ = {
    "리튬":   ["k_risk", "risk_리튬", "reserves", "mineral_index", "komir_trade"],
    "코발트": ["k_risk", "risk_코발트", "reserves", "komir_trade"],
    "니켈":   ["k_risk", "risk_니켈", "reserves", "resource_dev"],
    "희토류": ["reserves", "mineral_index", "komir_trade"],
    "텅스텐": ["k_risk", "risk_텅스텐", "mineral_index", "reserves"],
    "망간":   ["reserves", "resource_dev", "mineral_index"],
    "흑연":   ["reserves", "mineral_index", "komir_trade"],
    "경제":   ["mineral_index", "komir_trade", "k_risk"],
    "통상":   ["komir_trade", "reserves"],
    "지정학": ["k_risk", "reserves", "komir_trade"],
    "정책":   ["k_risk", "resource_dev", "mineral_index"],
}


def build_viz_catalog():
    """발언 근거용 차트 카탈로그. 각 데이터셋은 독립적으로 try-guard."""
    here = os.path.dirname(__file__)
    _pj = lambda f: os.path.join(here, f)
    tail = lambda a, n: (a[-n:] if isinstance(a, list) and len(a) > n else (a or []))
    cat = {}

    # 1) 광물 수급안정화지수 (광물별 라인)
    try:
        for r in (load_risk_data() or []):
            nm, ms, vs = r.get("name"), r.get("months", []), r.get("vals", [])
            if not (nm and ms and vs):
                continue
            _lt = r.get("latest"); _pv = r.get("prev")
            cat[f"risk_{nm}"] = {
                "title": f"{nm} 수급안정화지수", "type": "line",
                "labels": tail(ms, 24),
                "series": [{"name": "수급안정화지수", "data": tail(vs, 24), "color": "#1e8e5a"}],
                "headline": {"label": "최신 지수", "value": f"{_lt:.1f}" if _lt is not None else "—",
                             "sub": (f"전월 {'▲' if _lt >= _pv else '▼'} {abs(_lt - _pv):.1f}" if (_lt is not None and _pv is not None) else "")},
                "hlines": [{"v": 55, "label": "안정", "color": "#3fae7e"}, {"v": 30, "label": "위험", "color": "#e06060"}],
                "note": "≥55 안정 · 30~54 주의 · <30 위험",
                "source": "KOMIS 수급안정화지수",
            }
    except Exception as e:
        print("[VIZ risk]", e)

    # 1-b) K-RISK 종합 공급망 위험 점수 (광물별 바)
    try:
        kr = compute_k_risk() or {}
        if kr:
            items = sorted(kr.items(), key=lambda x: -x[1]["score"])
            worst = items[0]
            cat["k_risk"] = {
                "title": "K-RISK 종합 공급망 위험 점수", "type": "bar",
                "labels": [k for k, _ in items],
                "series": [{"name": "K-RISK(0~100)", "data": [v["score"] for _, v in items],
                            "color": "#d64545"}],
                "headline": {"label": "최고 위험", "value": f"{worst[1]['score']:.1f}",
                             "sub": f"{worst[0]} · {worst[1]['grade']}"},
                "bands": {"edges": [40, 70], "colors": ["#3fae7e", "#e0a92e", "#e05555"]},
                "hlines": [{"v": 70, "label": "위험", "color": "#e06060"}],
                "note": "🟢 0~39 안정 · 🟡 40~69 주의 · 🔴 70~100 위험",
                "source": "산업부 공공데이터 교차 계산(K-RISK)",
            }
    except Exception as e:
        print("[VIZ k_risk]", e)

    # 2) 광물 가격지수 (종합·카테고리)
    try:
        midx = json.load(open(_pj("mineral_index_data2.json"), encoding="utf-8"))
        S = midx.get("series", {})
        base = S.get("종합") or {}
        if base.get("months"):
            pal = {"종합": "#c98500", "에너지광물": "#eb6834", "희소금속": "#1c5cab", "메이저금속": "#4a3aa7"}
            ser = [{"name": k, "data": tail(S[k]["values"], 36), "color": pal.get(k, "#aaa")}
                   for k in pal if S.get(k)]
            sm = midx.get("summary", {}).get("종합", {})
            _mom = sm.get("mom")
            cat["mineral_index"] = {
                "title": "광물 가격지수", "type": "line",
                "labels": tail(base["months"], 36), "series": ser,
                "headline": {"label": "종합지수", "value": f"{sm.get('latest', 0):,.0f}",
                             "sub": (f"전월 {'▲' if _mom >= 0 else '▼'} {abs(_mom):.1f}%" if _mom is not None else sm.get("asof", ""))},
                "note": f"기준 {sm.get('asof','')} · 2016.1=1000",
                "source": "광해광업공단 파생지수",
            }
    except Exception as e:
        print("[VIZ midx]", e)

    # 3) 주요 광물 매장량 (USGS)
    try:
        items = sorted(USGS_DATA.items(), key=lambda kv: kv[1].get("매장량_만톤", 0), reverse=True)
        cat["reserves"] = {
            "title": "주요 광물 세계 매장량", "type": "bar",
            "headline": {"label": "매장량 1위", "value": items[0][0],
                         "sub": f"{items[0][1].get('1위국','')} 최다 보유"},
            "labels": [k for k, _ in items],
            "series": [{"name": "매장량(만톤)", "data": [v.get("매장량_만톤", 0) for _, v in items], "color": "#4a3aa7"}],
            "note": "1위국: " + ", ".join(f"{k}={v.get('1위국','')}" for k, v in items[:3]),
            "source": "USGS MCS 2025",
        }
    except Exception as e:
        print("[VIZ reserves]", e)

    # 7) 자원개발률(자주개발률)
    try:
        rdev = json.load(open(_pj("resource_dev_data2.json"), encoding="utf-8"))
        if rdev.get("series") and rdev.get("years"):
            pal = ["#c98500", "#5b6b7f", "#1c5cab", "#b58210", "#1e8e5a", "#c2447e"]
            ser = [{"name": k, "data": v, "color": pal[i % len(pal)]}
                   for i, (k, v) in enumerate(rdev["series"].items())]
            cat["resource_dev"] = {
                "title": "자원개발률(자주개발률) 추이", "type": "line",
                "labels": rdev["years"], "series": ser,
                "note": "품목별 자주개발률 %", "source": "산업통상자원부",
            }
    except Exception as e:
        print("[VIZ rdev]", e)

    # 9) 광종별 수입액 상위 (KOMIR)
    try:
        agg = {}
        for r in (local_customs() or []):
            nm = (r.get("광물명") or "").strip()
            if nm:
                agg[nm] = agg.get(nm, 0) + (r.get("수입금액(달러)", 0) or 0)
        top = [(k, v) for k, v in sorted(agg.items(), key=lambda kv: kv[1], reverse=True)[:7] if v > 0]
        if top:
            _tt = sum(v for _, v in top)
            cat["komir_trade"] = {
                "title": "광종별 수입액 상위", "type": "bar",
                "headline": {"label": "상위 광종 수입액", "value": f"${_tt/1e9:,.1f}B",
                             "sub": f"1위 {top[0][0]} ${top[0][1]/1e9:,.1f}B"},
                "labels": [k for k, _ in top],
                "series": [{"name": "수입액(억$)", "data": [round(v / 1e8, 1) for _, v in top], "color": "#c98500"}],
                "note": "최신연도 수입금액", "source": "광해광업공단 수출입",
            }
    except Exception as e:
        print("[VIZ komir]", e)

    return cat


def render_conference():
    experts_json = json.dumps(
        {k: {kk: vv for kk, vv in v.items() if kk not in ("system", "api_key")} for k, v in MINERAL_EXPERTS.items()},
        ensure_ascii=False
    )
    _cat = build_viz_catalog()
    viz_json = json.dumps(_cat, ensure_ascii=False)
    expertviz_json = json.dumps(
        {k: [vk for vk in v if vk in _cat] for k, v in EXPERT_VIZ.items()}, ensure_ascii=False
    )
    # 오늘의 리스크 안건 — K-RISK가 감지한 위험을 회의 안건 + AI 추천 전문가 조합으로 자동 제안
    def _valid_ex(lst):
        out = []
        for e in lst:
            if e in MINERAL_EXPERTS and e not in out:
                out.append(e)
        return out
    agenda = []
    kr = {}
    try:
        kr = compute_k_risk() or {}
        _tops = [(k, v) for k, v in sorted(kr.items(), key=lambda x: -x[1]["score"])
                 if v["grade"] in ("위험", "주의")][:6]
        _QTPL = {
            "위험": "{ico} [K-RISK {sc}{pv}] {k} 공급망이 '위험' 단계입니다. 원인 진단과 한국의 긴급 대응 전략은?",
            "주의": "{ico} [K-RISK {sc}{pv}] {k} 위험이 '주의' 단계로 관측됩니다. 선제적으로 무엇을 준비해야 합니까?",
        }
        for k, v in _tops:
            ico = "🔴" if v["grade"] == "위험" else "🟡"
            pv = "·잠정" if v.get("잠정") else ""
            comp = v.get("요소") or {}
            hhi = comp.get("수입집중도")
            extra = f" (수입집중 {hhi:.0f})" if isinstance(hhi, (int, float)) and hhi >= 60 else ""
            agenda.append({
                "q": _QTPL[v["grade"]].format(ico=ico, sc=v["score"], pv=pv, k=k) + extra,
                "ex": _valid_ex(([k] if k in MINERAL_EXPERTS else ["통상"]) + ["지정학", "정책", "경제"])})
    except Exception: pass
    agenda_json = json.dumps(agenda, ensure_ascii=False)
    krisk_json = json.dumps({k: {"score": v["score"], "grade": v["grade"]} for k, v in kr.items()},
                            ensure_ascii=False)
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    PAGE = r"""<!DOCTYPE html>
<html class="dark" lang="ko">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>K Mineral Risk — AI 전문가 회의실</title>
<link rel="icon" type="image/png" href="/static/favicon.png?v=3">
<script src="https://cdn.tailwindcss.com?plugins=forms,container-queries"></script>
<link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800;900&family=JetBrains+Mono:wght@400;500&family=Noto+Sans+KR:wght@400;500;700;900&display=swap" rel="stylesheet">
<link href="https://fonts.googleapis.com/css2?family=Material+Symbols+Outlined:wght,FILL@100..700,0..1&display=swap" rel="stylesheet">
<link href="https://fonts.googleapis.com/css2?family=Cormorant+Garamond:ital,wght@0,500;0,600;1,400&family=Noto+Serif+KR:wght@500;700&display=swap" rel="stylesheet">
<link rel="stylesheet" href="https://cdn.jsdelivr.net/gh/orioncactus/pretendard@v1.3.9/dist/web/static/pretendard.min.css">
<style>
:root{--sans:'Pretendard','Noto Sans KR',-apple-system,sans-serif;--mono:'Pretendard','Noto Sans KR',-apple-system,sans-serif;}
html body, html body *:not([class*="material-symbols"]){font-family:'Pretendard','Noto Sans KR',-apple-system,BlinkMacSystemFont,sans-serif !important;}
body{font-variant-numeric:tabular-nums;}
</style>
<script>
tailwind.config = {
  darkMode: "class",
  theme: { extend: {
    colors: {
      "surface-variant":"#eef2f8","outline-variant":"#dbe2ec","surface-container-low":"#f7f9fc",
      "surface-container-lowest":"#ffffff","on-surface":"#16233c","surface-container-high":"#eef2f8",
      "background":"#f4f6fa","surface-container-highest":"#e6ebf3","primary":"#12325e",
      "surface-container":"#ffffff","on-secondary":"#ffffff","outline":"#8a94a6",
      "primary-container":"#e8eef7","on-primary-container":"#12325e","on-surface-variant":"#4a5872",
      "secondary":"#b58210","on-secondary-fixed":"#ffffff","surface":"#ffffff","error":"#c03535",
      "tertiary":"#4a5872"
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
  body{background:#f4f6fa;color:#16233c;font-family:'Inter','Noto Sans KR',sans-serif;}
  .material-symbols-outlined{font-variation-settings:'FILL' 0,'wght' 400,'GRAD' 0,'opsz' 24;}
  .glass-panel{background:rgba(255,255,255,.75);backdrop-filter:blur(12px);border:1px solid #dbe2ec;}
  .src-chip{display:inline-flex;align-items:center;gap:3px;font-size:10.5px;font-weight:600;line-height:1;
    color:#8a6a10;background:rgba(200,147,29,.1);border:1px solid rgba(181,130,16,.4);
    padding:2px 7px;border-radius:10px;margin:0 2px;white-space:nowrap;vertical-align:1px;}
  .src-chip::before{content:'◆';font-size:7px;opacity:.7;}
  .custom-scrollbar::-webkit-scrollbar{width:6px;}
  .custom-scrollbar::-webkit-scrollbar-track{background:transparent;}
  .custom-scrollbar::-webkit-scrollbar-thumb{background:#cdd6e4;border-radius:10px;}
  .expert-card.selected{border-color:#c8931d !important;box-shadow:0 0 0 1px #c8931d,0 0 16px rgba(200,147,29,.2);}
  .expert-card.selected .ec-check{opacity:1 !important;}
  .tc-suggested{box-shadow:0 0 0 1px #c8931d,0 0 10px rgba(200,147,29,.35);}
  .agenda-chip.agenda-active{border-color:#c8931d;box-shadow:0 0 0 1px #c8931d,0 0 12px rgba(200,147,29,.25);color:#8a6a10;}
  .sum-hero{font-size:15px;font-weight:800;color:#7a5c0a;background:rgba(200,147,29,.08);border-left:3px solid #c8931d;border-radius:8px;padding:12px 16px;margin:2px 0 16px;line-height:1.6;}
  .sum-sec{margin-bottom:14px;}
  .sum-sec-t{font-size:12px;font-weight:800;color:#a97a12;letter-spacing:.05em;margin-bottom:7px;text-transform:uppercase;}
  .sum-ul{margin:0;padding:0;list-style:none;display:flex;flex-direction:column;gap:6px;}
  .sum-ul li{font-size:13px;line-height:1.7;background:#f7f9fc;border:1px solid #e3e8f0;border-radius:9px;padding:9px 13px;}
  .sum-meta-row{display:flex;flex-wrap:wrap;gap:6px 14px;font-size:11px;margin-bottom:14px;padding-bottom:12px;border-bottom:1px solid #e3e8f0;}
  /* ── 회의 결과 보고서 오버레이 (전면 대시보드형) ── */
  #summaryOverlay{position:fixed;inset:0;z-index:60;display:none;}
  #summaryOverlay .sum-bk{position:absolute;inset:0;background:rgba(20,35,60,.45);backdrop-filter:blur(8px);}
  .sum-doc{position:relative;width:calc(100vw - 28px);max-width:1560px;margin:14px auto;height:calc(100% - 28px);overflow-y:auto;
    border-radius:20px;background:#ffffff;border:1px solid #dbe2ec;box-shadow:0 30px 90px rgba(20,35,60,.3);}
  .sum-cover-grid{display:flex;justify-content:space-between;align-items:center;gap:30px;position:relative;z-index:1;}
  .sum-gauge{flex-shrink:0;text-align:center;}
  .sum-gauge svg{display:block;margin:0 auto;}
  .sum-gauge .gl{font-size:11px;font-weight:800;letter-spacing:.1em;color:#7d8aa3;margin-top:6px;}
  .sum-statrow{display:grid;grid-template-columns:repeat(4,1fr);gap:1px;background:#e3e8f0;border-bottom:1px solid #e3e8f0;}
  .sum-stat{background:#f7f9fc;padding:16px 24px;}
  .sum-stat .v{font-size:24px;font-weight:900;font-family:'JetBrains Mono',monospace;color:#12325e;line-height:1.2;}
  .sum-stat .v small{font-size:12px;color:#8a94a6;font-weight:600;margin-left:3px;}
  .sum-stat .l{font-size:11px;font-weight:700;color:#7d8aa3;margin-top:3px;letter-spacing:.05em;}
  .sum-cols{display:grid;grid-template-columns:1.55fr 1fr;gap:30px;align-items:start;}
  .sum-quote{display:flex;gap:11px;background:#f7f9fc;border:1px solid #e3e8f0;border-radius:12px;padding:12px 14px;margin-bottom:8px;}
  .sum-quote .av{flex-shrink:0;width:32px;height:32px;border-radius:9px;border:1px solid #dbe2ec;display:flex;align-items:center;justify-content:center;font-size:16px;}
  .sum-quote .nm{font-size:11.5px;font-weight:800;margin-bottom:3px;}
  .sum-quote .tx{font-size:12px;line-height:1.65;color:#42536e;}
  @media(max-width:1000px){.sum-cols{grid-template-columns:1fr;}.sum-statrow{grid-template-columns:repeat(2,1fr);}}
  .sum-cover{position:relative;background:linear-gradient(135deg,#12325e 0%,#1c4c8a 60%,#12325e 100%);padding:44px 48px 34px;overflow:hidden;}
  .sum-cover::after{content:'';position:absolute;inset:0;background:radial-gradient(120% 140% at 85% -10%,rgba(233,195,73,.16),transparent 55%);pointer-events:none;}
  .sum-brand{font-size:11px;font-weight:800;letter-spacing:.22em;color:#f0d68a;text-transform:uppercase;margin-bottom:14px;}
  .sum-doc-title{font-size:30px;font-weight:900;color:#fff;letter-spacing:-.01em;margin-bottom:8px;}
  .sum-agenda{font-size:14px;color:#c7d2e8;line-height:1.6;max-width:640px;}
  .sum-cover-meta{display:flex;flex-wrap:wrap;gap:8px 18px;margin-top:20px;font-size:12px;color:#b9cbe4;}
  .sum-cover-meta b{color:#e6ecf7;font-weight:700;}
  .sum-exp-chip{display:inline-flex;align-items:center;gap:5px;font-size:12px;font-weight:700;border:1px solid rgba(255,255,255,.16);
    border-radius:999px;padding:4px 12px;background:rgba(255,255,255,.04);}
  .sum-docbody{padding:34px 48px 40px;}
  .sum-hero2{font-size:22px;font-weight:900;color:#7a5c0a;line-height:1.5;padding:22px 26px;margin-bottom:28px;
    background:#fdf8ea;border:1px solid #eddfb5;border-left:5px solid #c8931d;border-radius:14px;}
  .sum-h{display:flex;align-items:center;gap:8px;font-size:13px;font-weight:900;color:#a97a12;letter-spacing:.08em;margin:26px 0 12px;text-transform:uppercase;}
  .sum-h::after{content:'';flex:1;height:1px;background:linear-gradient(90deg,rgba(181,130,16,.4),transparent);}
  .sum-grid3{display:grid;grid-template-columns:repeat(auto-fit,minmax(200px,1fr));gap:12px;}
  .sum-key{background:#f7f9fc;border:1px solid #e3e8f0;border-radius:14px;padding:16px 16px 14px;}
  .sum-key .no{font-size:24px;font-weight:900;color:#c8931d66;font-family:'JetBrains Mono',monospace;line-height:1;margin-bottom:8px;}
  .sum-key .tx{font-size:13px;line-height:1.7;color:#2b3a55;}
  .sum-imp{display:flex;gap:12px;align-items:flex-start;background:#f7f9fc;border:1px solid #e3e8f0;border-radius:12px;padding:12px 16px;margin-bottom:8px;}
  .sum-imp .bd{flex-shrink:0;width:24px;height:24px;border-radius:7px;background:rgba(200,147,29,.14);color:#a97a12;font-weight:900;font-size:12px;display:flex;align-items:center;justify-content:center;font-family:'JetBrains Mono',monospace;}
  .sum-imp .tx{font-size:13.5px;line-height:1.7;color:#2b3a55;}
  .sum-issue{background:#fdf3f2;border:1px solid #f0cfcc;border-radius:12px;padding:14px 18px;font-size:13.5px;line-height:1.7;color:#7c3a35;margin-bottom:8px;}
  .sum-next{display:inline-flex;align-items:center;gap:7px;font-size:12.5px;font-weight:600;color:#155f3e;background:rgba(30,142,90,.07);
    border:1px solid rgba(30,142,90,.3);border-radius:999px;padding:7px 15px;margin:0 8px 8px 0;line-height:1.5;}
  .sum-foot{margin-top:30px;padding-top:16px;border-top:1px solid #e3e8f0;display:flex;justify-content:space-between;flex-wrap:wrap;gap:8px;font-size:11px;color:#7d8aa3;}
  .sum-actions{position:sticky;top:14px;float:right;display:flex;gap:8px;margin:14px 16px 0 0;z-index:5;}
  .sum-actions button{font-size:12px;font-weight:800;border-radius:10px;padding:8px 14px;border:1px solid #cdd6e4;
    background:rgba(255,255,255,.92);color:#16233c;cursor:pointer;backdrop-filter:blur(4px);}
  .sum-actions button:hover{border-color:#c8931d;color:#8a6a10;}
  .sum-doc .viz-card{margin-top:10px;}
  @media(max-width:760px){.sum-grid3{grid-template-columns:1fr;}.sum-cover,.sum-docbody{padding-left:22px;padding-right:22px;}}
  .lobby-screen,#roomScreen{display:none;}
  .viz-btn{margin-left:6px;font-size:10px;font-weight:800;padding:2px 9px;border-radius:9px;cursor:pointer;
    background:rgba(200,147,29,.1);border:1px solid rgba(181,130,16,.45);color:#8a6a10;transition:.15s;}
  .viz-btn:hover{background:#c8931d;color:#fff;}
  body.cine .viz-btn{background:rgba(126,166,255,.08);border-color:rgba(126,166,255,.35);color:#9fc0ff;}
  body.cine .viz-btn:hover{background:rgba(126,166,255,.25);color:#fff;box-shadow:0 0 12px rgba(126,166,255,.35);}
  /* ═══ 시네마 모드 — 영화 속 AI 회의실 ═══ */
  body.cine{background:radial-gradient(1100px 700px at 50% -10%, #0d1830 0%, #070c18 55%, #04070f 100%) fixed !important;}
  body.cine::before{content:'';position:fixed;inset:0;pointer-events:none;z-index:1;
    background:repeating-linear-gradient(0deg, rgba(140,180,255,.022) 0 1px, transparent 1px 3px);}
  body.cine::after{content:'';position:fixed;inset:0;pointer-events:none;z-index:1;
    box-shadow:inset 0 0 220px rgba(0,0,0,.75);}
  body.cine main{background:transparent !important;}
  body.cine header{background:rgba(7,12,24,.8)!important;border-bottom:1px solid rgba(126,166,255,.14)!important;backdrop-filter:blur(14px)!important;}
  body.cine header span, body.cine header a{color:#9fb4d8!important;}
  body.cine #roomScreen > div:first-child{background:rgba(7,12,24,.6)!important;border-bottom:1px solid rgba(126,166,255,.1)!important;}
  body.cine .aud-tag{background:rgba(233,198,103,.14);color:#e9c667;border:1px solid rgba(233,198,103,.35);}
  body.cine .exp-seat{background:rgba(13,22,40,.85);border-color:rgba(126,166,255,.16);box-shadow:none;}
  body.cine .exp-seat .es-name{color:var(--exc);text-shadow:0 0 12px color-mix(in srgb,var(--exc) 55%,transparent);}
  body.cine .exp-seat.speaking{background:rgba(16,28,52,.95);
    box-shadow:0 0 0 2px color-mix(in srgb,var(--exc) 55%,transparent),0 0 26px color-mix(in srgb,var(--exc) 40%,transparent);}
  body.cine .msg-bubble{background:rgba(10,18,34,.78)!important;border:1px solid rgba(126,166,255,.12)!important;
    color:#dfe8f7!important;backdrop-filter:blur(10px);
    box-shadow:0 10px 30px rgba(0,0,0,.45),inset 0 0 22px rgba(126,166,255,.03)!important;}
  body.cine #chatArea .text-sm.font-bold{text-shadow:0 0 14px currentColor;}
  body.cine #turnControls, body.cine #consoleBar{
    background:rgba(7,12,24,.82)!important;border-color:rgba(126,166,255,.12)!important;backdrop-filter:blur(14px);}
  body.cine #chatInput{background:rgba(10,18,34,.9)!important;border-color:rgba(126,166,255,.22)!important;color:#dfe8f7!important;}
  body.cine #chatInput::placeholder{color:#5d7396;}
  body.cine #micBtn{border-color:rgba(126,166,255,.35);color:#9fb4d8;background:rgba(10,18,34,.8);}
  body.cine #micBtn.rec{background:#d64545!important;box-shadow:0 0 34px rgba(214,69,69,.6);}
  body.cine #voiceModeBtn{border-color:rgba(126,166,255,.3);color:#9fb4d8;}
  body.cine #voiceModeBtn.on{background:#e9c667!important;border-color:#e9c667!important;color:#0d1830!important;
    box-shadow:0 0 26px rgba(233,198,103,.45);}
  body.cine #micStatus{color:#e9c667!important;text-shadow:0 0 12px rgba(233,198,103,.4);}
  body.cine #typingIndicator{color:#7ea6ff!important;text-shadow:0 0 10px rgba(126,166,255,.5);}
  body.cine .src-chip{background:rgba(233,198,103,.1);border-color:rgba(233,198,103,.3);color:#e9c667;}
  body.cine .viz-card{background:rgba(10,18,34,.85)!important;border-color:rgba(126,166,255,.15)!important;}
  body.cine .viz-head{color:#dfe8f7!important;}
  body.cine .vs-val{color:#e9c667!important;}
  /* 진행자(사용자) 발언 — 커맨드 라인 느낌 */
  body.cine .user-msg-bubble{background:rgba(233,198,103,.1)!important;border:1px solid rgba(233,198,103,.3)!important;color:#f0dfae!important;}

  /* ── 스테이지 오브 (발언자 홀로그램) ── */
  #stageOrb{display:none;position:relative;height:0;z-index:5;pointer-events:none;}
  body.cine #stageOrb.on{display:block;}
  #stageOrb .so-core{position:absolute;left:50%;top:26px;transform:translateX(-50%);width:74px;height:74px;border-radius:50%;
    display:flex;align-items:center;justify-content:center;
    background:radial-gradient(circle at 38% 32%, color-mix(in srgb,var(--exc,#e9c667) 32%, #0d1830), #070c18 75%);
    border:2px solid var(--exc,#e9c667);
    box-shadow:0 0 34px color-mix(in srgb,var(--exc,#e9c667) 55%,transparent),inset 0 0 22px color-mix(in srgb,var(--exc,#e9c667) 30%,transparent);}
  #stageOrb .so-avatar{font-size:32px;filter:drop-shadow(0 0 8px var(--exc,#e9c667));}
  #stageOrb .so-ring{position:absolute;left:50%;top:63px;transform:translate(-50%,-50%);border-radius:50%;
    border:1.5px solid var(--exc,#e9c667);opacity:0;animation:soRipple 2.4s cubic-bezier(0,.4,.4,1) infinite;}
  #stageOrb .so-ring.r1{width:90px;height:90px;animation-delay:0s;}
  #stageOrb .so-ring.r2{width:90px;height:90px;animation-delay:.8s;}
  @keyframes soRipple{0%{transform:translate(-50%,-50%) scale(.9);opacity:.75;}100%{transform:translate(-50%,-50%) scale(2.6);opacity:0;}}
  /* JARVIS 코어 — 회전 세그먼트 링 */
  #stageOrb .so-spin1{position:absolute;left:50%;top:63px;width:104px;height:104px;transform:translate(-50%,-50%);
    border:1.5px dashed var(--exc,#5fd0ff);border-radius:50%;opacity:.75;animation:soSpin 7s linear infinite;}
  #stageOrb .so-spin2{position:absolute;left:50%;top:63px;width:128px;height:128px;transform:translate(-50%,-50%);border-radius:50%;
    background:conic-gradient(var(--exc,#5fd0ff) 0 60deg,transparent 60deg 100deg,var(--exc,#5fd0ff) 100deg 170deg,transparent 170deg 230deg,var(--exc,#5fd0ff) 230deg 320deg,transparent 320deg);
    -webkit-mask:radial-gradient(closest-side,transparent 88%,#000 89%);mask:radial-gradient(closest-side,transparent 88%,#000 89%);
    opacity:.85;animation:soSpinR 10s linear infinite;}
  #stageOrb .so-ticks{position:absolute;left:50%;top:63px;width:152px;height:152px;transform:translate(-50%,-50%);border-radius:50%;
    background:repeating-conic-gradient(var(--exc,#5fd0ff) 0 1.4deg,transparent 1.4deg 7.2deg);
    -webkit-mask:radial-gradient(closest-side,transparent 92%,#000 93%);mask:radial-gradient(closest-side,transparent 92%,#000 93%);
    opacity:.4;animation:soSpin 24s linear infinite;}
  @keyframes soSpin{to{transform:translate(-50%,-50%) rotate(360deg);}}
  @keyframes soSpinR{to{transform:translate(-50%,-50%) rotate(-360deg);}}
  #stageOrb .so-core{animation:soCore 2.2s ease-in-out infinite;}
  @keyframes soCore{0%,100%{filter:brightness(1);}50%{filter:brightness(1.35);}}

  /* HUD 도크 — 보조 설명 홀로그램 패널 */
  #hudDock{display:none;position:fixed;right:26px;top:50%;transform:translateY(-50%) translateX(30px);
    width:min(420px,31vw);z-index:40;opacity:0;pointer-events:auto;transition:.45s cubic-bezier(.2,.8,.3,1);}
  body.cine #hudDock.show{display:block;opacity:1;transform:translateY(-50%);}
  #hudDock .hud-title{display:flex;align-items:center;gap:8px;font-size:10px;font-weight:900;letter-spacing:.22em;
    color:var(--exc,#5fd0ff);text-shadow:0 0 12px color-mix(in srgb,var(--exc,#5fd0ff) 60%,transparent);
    padding:0 4px 8px;text-transform:uppercase;}
  #hudDock .hud-title span{margin-left:auto;letter-spacing:.04em;font-weight:600;color:#7d93b8;}
  #hudDock #hudBody{position:relative;background:rgba(8,14,28,.88);backdrop-filter:blur(12px);
    border:1px solid color-mix(in srgb,var(--exc,#5fd0ff) 35%,transparent);border-radius:6px;
    box-shadow:0 0 40px color-mix(in srgb,var(--exc,#5fd0ff) 18%,transparent),inset 0 0 30px rgba(126,166,255,.04);
    padding:6px;}
  #hudDock #hudBody::before,#hudDock #hudBody::after{content:'';position:absolute;width:16px;height:16px;
    border:2px solid var(--exc,#5fd0ff);}
  #hudDock #hudBody::before{top:-2px;left:-2px;border-right:0;border-bottom:0;}
  #hudDock #hudBody::after{bottom:-2px;right:-2px;border-left:0;border-top:0;}
  #hudDock .viz-card{background:transparent!important;border:none!important;max-width:none!important;margin:0!important;box-shadow:none!important;}
  #hudDock .viz-head{color:#dfe8f7!important;}
  #hudDock .viz-canvas{height:230px!important;}
  #hudDock .viz-kpi b{font-size:32px;text-shadow:0 0 18px currentColor;}
  #hudDock .viz-kpi i{color:#7d93b8;}
  #hudDock .viz-kpi em{color:#9fb4d8;}
  #hudDock .viz-foot{color:#7d93b8!important;}
  @media(max-width:1200px){#hudDock{display:none!important;}}
  #stageOrb .so-meta{position:absolute;left:50%;top:148px;transform:translateX(-50%);text-align:center;white-space:nowrap;}
  #stageOrb .so-meta b{display:block;font-size:13px;font-weight:900;color:var(--exc,#e9c667);
    text-shadow:0 0 14px color-mix(in srgb,var(--exc,#e9c667) 60%,transparent);letter-spacing:.04em;}
  #stageOrb .so-meta i{display:block;font-style:normal;font-size:10px;color:#7d93b8;margin-top:2px;letter-spacing:.06em;}
  /* ── 시네마 레이아웃: 우측 카톡 컬럼 + 좌측 홀로그램 무대 ── */
  body.cine #chatArea{position:fixed!important;right:0;top:64px;bottom:0;width:390px;z-index:30;
    background:rgba(6,10,22,.82);backdrop-filter:blur(16px);
    border-left:1px solid rgba(126,166,255,.14);
    padding:52px 16px 150px 16px!important;overflow-y:auto;}
  body.cine #chatArea::before{content:'LIVE TRANSCRIPT';position:fixed;right:0;top:64px;width:390px;z-index:31;
    font-size:9.5px;font-weight:900;letter-spacing:.26em;color:#7d93b8;box-sizing:border-box;
    padding:14px 16px 9px;border-bottom:1px solid rgba(126,166,255,.12);
    background:linear-gradient(rgba(6,10,22,.97),rgba(6,10,22,.88));backdrop-filter:blur(10px);}
  body.cine #chatArea .max-w-\[85\%\]{max-width:100%!important;}
  body.cine #chatArea .max-w-\[75\%\]{max-width:92%!important;}
  body.cine .msg-bubble{font-size:13px!important;padding:10px 13px!important;}
  /* 무대 중심 좌표: 우측 컬럼 제외한 중앙 */
  body.cine #stageOrb{position:fixed;left:calc((100vw - 390px)/2);top:120px;width:0;height:0;z-index:20;}
  #liveCaption{display:none!important;}   /* 무대엔 글자 없음 — 대화는 우측 트랜스크립트에만 */
  body.cine #stageOrb{display:none!important;}
  /* 홀로그램 캔버스 — 무대 중앙 */
  #holoCanvas{display:none;}
  body.cine #holoCanvas{display:block;position:fixed;z-index:15;pointer-events:none;
    left:calc((100vw - 390px)/2);top:calc(50% - 30px);transform:translate(-50%,-50%);
    width:min(78vh,820px);height:min(78vh,820px);
    filter:drop-shadow(0 0 40px rgba(80,200,255,.12));}
  /* HUD 도크 → 좌측 무대 왼편 */
  body.cine #hudDock{right:auto;left:4vw;top:52%;transform:translateY(-50%) translateX(-30px);width:min(350px,24vw);}
  body.cine #hudDock.show{transform:translateY(-50%);}
  /* 콘솔(입력바·턴컨트롤·상태)은 무대 하단 고정 */
  body.cine #consoleBar{position:fixed!important;left:0;right:390px;bottom:0;z-index:35;}
  body.cine #turnControls{position:fixed!important;left:0;right:390px;bottom:78px;z-index:35;}
  body.cine #micStatus{position:fixed!important;left:24px;right:410px;bottom:84px;z-index:36;padding:0!important;}
  body.cine #typingIndicator{position:fixed!important;left:24px;bottom:88px;z-index:36;padding:0!important;}
  @media(max-width:1100px){
    body.cine #chatArea{position:static!important;width:auto;padding:18px!important;}
    body.cine #chatArea::before{display:none;}
    body.cine #stageOrb{position:relative;left:auto;top:0;}
    body.cine #liveCaption{display:none!important;}
    body.cine #consoleBar, body.cine #turnControls{position:static!important;}
    body.cine #micStatus, body.cine #typingIndicator{position:static!important;}
  }

  /* 회의 룸 — 전문가 좌석 스트립 */
  .aud-tag{display:inline-flex;align-items:center;font-size:11px;font-weight:800;padding:7px 13px;border-radius:12px;
    background:#12325e;color:#f0d68a;letter-spacing:.02em;}
  .exp-seat{position:relative;display:inline-flex;align-items:center;gap:7px;padding:7px 13px;border-radius:12px;
    background:#fff;border:1.5px solid #e3e8f0;transition:.25s;box-shadow:0 1px 3px rgba(20,35,60,.06);}
  .es-avatar{font-size:17px;line-height:1;}
  .es-name{font-size:12px;font-weight:800;color:var(--exc,#12325e);}
  .es-eq{display:none;align-items:flex-end;gap:2px;height:13px;margin-left:2px;}
  .es-eq i{width:3px;border-radius:2px;background:var(--exc,#c8931d);animation:eqB .7s ease-in-out infinite;}
  .es-eq i:nth-child(1){height:5px;animation-delay:0s}
  .es-eq i:nth-child(2){height:11px;animation-delay:.15s}
  .es-eq i:nth-child(3){height:7px;animation-delay:.3s}
  .es-eq i:nth-child(4){height:12px;animation-delay:.45s}
  @keyframes eqB{0%,100%{transform:scaleY(.4)}50%{transform:scaleY(1)}}
  .exp-seat.speaking{border-color:var(--exc,#c8931d);background:#fffdf5;
    box-shadow:0 0 0 3px color-mix(in srgb,var(--exc,#c8931d) 22%, transparent),0 8px 22px rgba(20,35,60,.14);
    transform:translateY(-2px);}
  .exp-seat.speaking .es-eq{display:inline-flex;}
  .exp-seat.speaking::after{content:'ON AIR';position:absolute;top:-8px;right:-6px;font-size:8px;font-weight:900;
    letter-spacing:.08em;background:#d64545;color:#fff;padding:2px 6px;border-radius:6px;}
  /* 룸 배경 — 은은한 네이비 그라데이션 */
  #roomScreen #chatArea{background:
    radial-gradient(700px 300px at 85% -5%, rgba(28,92,171,.05), transparent 60%),
    radial-gradient(500px 260px at 5% 0%, rgba(200,147,29,.05), transparent 55%);}
  .msg-bubble{font-size:14.5px!important;}
  #micBtn.rec{background:#d64545!important;border-color:#d64545!important;color:#fff!important;animation:micPulse 1.1s ease-in-out infinite;}
  #voiceModeBtn.on{background:#12325e!important;border-color:#12325e!important;color:#fff!important;}
  @keyframes micPulse{0%,100%{box-shadow:0 0 0 0 rgba(214,69,69,.45)}50%{box-shadow:0 0 0 9px rgba(214,69,69,0)}}
  .aud-btn{padding:9px 16px;border-radius:12px;font-size:13px;font-weight:700;cursor:pointer;background:#fff;border:1px solid #cdd6e4;color:#4a5872;transition:.18s;}
  .aud-btn:hover{border-color:#c8931d;color:#16233c;}
  .aud-btn.active{background:#12325e;color:#fff;border-color:#12325e;box-shadow:0 4px 14px rgba(18,50,94,.25);}

  /* ===== 라이트 화이트-골드 리스킨 ===== */
  body{background:linear-gradient(180deg,#f7f9fc 0%,#f2f5fa 100%) fixed !important;}
  .glass-panel{background:rgba(255,255,255,.8)!important;backdrop-filter:blur(16px)!important;border:1px solid #dbe2ec!important;}
  /* 시네마틱 헤드라인 — 세리프 */
  .text-headline-lg{font-family:'Noto Serif KR','Cormorant Garamond',serif!important;letter-spacing:-.01em;}
  .text-headline-md{font-family:'Cormorant Garamond','Noto Serif KR',serif!important;letter-spacing:.01em;}
  /* 사이드바·헤더 */
  aside.fixed{background:rgba(255,255,255,.95)!important;border-right:1px solid #e3e8f0!important;}
  aside nav a{border-radius:12px!important;}
  header{background:rgba(255,255,255,.85)!important;border-bottom:1px solid #e3e8f0!important;}
  /* 전문가 카드 */
  .expert-card{border-radius:16px!important;transition:transform .2s cubic-bezier(.2,.7,.3,1),border-color .2s,box-shadow .2s!important;}
  .expert-card:hover{transform:translateY(-3px);border-color:rgba(200,147,29,.55)!important;box-shadow:0 14px 32px rgba(20,35,60,.12)!important;}
  .expert-card.selected{border-color:#c8931d !important;box-shadow:0 0 0 1px #c8931d,0 0 18px rgba(200,147,29,.2)!important;}
  /* 채팅 버블 */
  .msg-bubble{border-radius:16px!important;line-height:1.7!important;}
  /* 스크롤바 골드 */
  .custom-scrollbar::-webkit-scrollbar-thumb{background:rgba(181,130,16,.35)!important;}
  /* 브랜드 로고 블록 */
  aside .bg-secondary{background:linear-gradient(135deg,#e9c667,#c8931d)!important;box-shadow:0 4px 14px rgba(200,147,29,.3);}
  /* 발언 근거 시각자료 카드 */
  .viz-card{margin-top:10px;background:#ffffff;border:1px solid #e3e8f0;border-left:2px solid rgba(200,147,29,.7);
    border-radius:12px;padding:11px 13px;max-width:460px;animation:vizIn .4s cubic-bezier(.2,.7,.2,1);}
  @keyframes vizIn{from{opacity:0;transform:translateY(8px)}to{opacity:1;transform:none}}
  .viz-head{display:flex;align-items:center;gap:7px;font-size:12.5px;font-weight:700;color:#16233c;margin-bottom:7px;}
  .viz-dot{width:7px;height:7px;border-radius:50%;display:inline-block;flex-shrink:0;}
  .viz-canvas{position:relative;height:158px;}
  .viz-kpi{display:flex;align-items:baseline;gap:10px;margin:2px 0 8px;}
  .viz-kpi b{font-size:26px;font-weight:900;font-family:'Archivo','JetBrains Mono',monospace;letter-spacing:-.01em;line-height:1;}
  .viz-kpi span{display:flex;flex-direction:column;gap:1px;}
  .viz-kpi i{font-style:normal;font-size:9.5px;font-weight:700;letter-spacing:.08em;color:#8a94a6;text-transform:uppercase;}
  .viz-kpi em{font-style:normal;font-size:11px;font-weight:600;color:#5d6b80;}
  .viz-foot{margin-top:7px;font-size:9.5px;color:#8a94a6;letter-spacing:.02em;}
  .viz-stats{display:flex;gap:26px;padding:4px 2px 2px;}
  .vs-label{font-size:10.5px;color:#8a94a6;margin-bottom:2px;}
  .vs-val{font-size:25px;font-weight:800;color:#12325e;font-family:'JetBrains Mono',monospace;line-height:1;}
  .vs-val span{font-size:11px;color:#8a94a6;margin-left:3px;font-weight:500;}
  .vs-delta{font-size:11px;font-weight:700;margin-top:3px;}
</style>
<script src="https://cdnjs.cloudflare.com/ajax/libs/Chart.js/4.4.1/chart.umd.min.js"></script>
</head>
<body class="flex min-h-screen bg-background">

<!-- Main (사이드바 없음 — 회의에 집중) -->
<main class="flex-1 h-screen flex flex-col bg-background overflow-hidden">
  <header class="h-16 shrink-0 flex items-center justify-between px-8 border-b border-outline-variant/30 bg-surface/70 backdrop-blur-xl">
    <a href="/" class="flex items-center gap-3 no-underline" title="허브 홈">
      <span class="text-on-surface-variant text-sm font-bold" style="font-size:15px">AI 전문가 회의실</span>
    </a>
    <div class="flex items-center gap-4">
      <span id="confClock" class="font-data-tabular text-xs text-on-surface-variant">__NOW__ KST ● LIVE</span>
      <a href="/" class="flex items-center gap-2 px-4 py-2 rounded-full border border-secondary/40 text-secondary text-sm font-bold hover:bg-secondary hover:text-on-secondary-fixed transition">
        <span class="material-symbols-outlined text-base">logout</span> 회의실 나가기
      </a>
    </div>
  </header>

  <div class="flex-1 min-h-0 relative">

    <!-- STEP 1 -->
    <div id="step1Screen" class="lobby-screen absolute inset-0 flex-col items-center overflow-y-auto p-8 custom-scrollbar" style="display:flex">
      <div class="w-full max-w-3xl mx-auto">
        <h1 class="text-headline-lg text-on-surface mb-2">자원·원자재 AI 전문가 회의실</h1>
        <p class="text-on-surface-variant text-sm mb-6"><span class="text-secondary font-bold">STEP 1.</span> 누구를 위한 회의인지 <b class="text-secondary">대상</b>을 고르고, 회의에 데려갈 <b class="text-secondary">전문가</b>를 선택하세요. 광물·경제·정치 전문가가 함께 토론하며, 같은 전문가라도 대상(투자자·기업·소비자)에 따라 토론이 달라집니다.</p>
        <div class="mb-7">
          <div class="text-[10px] font-bold text-outline uppercase tracking-widest mb-3 font-data-tabular">① 대상 선택 — 누구를 위한 분석인가</div>
          <div class="flex flex-wrap gap-2" id="audienceRow">
            <button class="aud-btn active" data-aud="investor" onclick="setAudience('investor',this)">📈 일반 투자자</button>
            <button class="aud-btn" data-aud="business" onclick="setAudience('business',this)">🏢 기업 · 조달</button>
            <button class="aud-btn" data-aud="consumer" onclick="setAudience('consumer',this)">🛒 일반 소비자</button>
            <button class="aud-btn" data-aud="policy" onclick="setAudience('policy',this)">🏛️ 정책 · 연구</button>
          </div>
        </div>
        <div class="mb-7" id="riskAgendaWrap">
          <div class="text-[10px] font-bold text-outline uppercase tracking-widest mb-3 font-data-tabular">② 오늘의 리스크 안건 <span class="text-secondary">— K-RISK가 감지한 위험 · 클릭하면 AI 추천 전문가 조합까지 자동 선택</span></div>
          <div id="riskAgenda" class="flex flex-col gap-2"></div>
        </div>
        <div class="text-[10px] font-bold text-outline uppercase tracking-widest mb-3 font-data-tabular">③ 전문가 선택 <span class="text-outline">— 직접 고르거나, 위 안건 클릭으로 자동 구성</span></div>
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
      <div id="stageOrb"><span class="so-ring r1"></span><span class="so-ring r2"></span>
        <span class="so-ticks"></span><span class="so-spin1"></span><span class="so-spin2"></span>
        <span class="so-core"><span class="so-avatar"></span></span>
        <span class="so-meta"><b class="so-name"></b><i class="so-role"></i></span></div>
      <div id="hudDock"><div class="hud-title">◆ TACTICAL DATA <span id="hudSrc"></span></div><div id="hudBody"></div></div>
      <div id="liveCaption"></div>
      <canvas id="holoCanvas"></canvas>
      <div id="chatArea" class="flex-1 overflow-y-auto p-8 space-y-6 custom-scrollbar"></div>
      <div id="typingIndicator" class="px-8 pb-1 text-xs text-secondary font-data-tabular" style="display:none">● 전문가가 답변 중...</div>
      <div id="turnControls" class="px-6 py-3 border-t border-outline-variant/20 bg-surface-container-low/40 flex items-center gap-3 flex-wrap shrink-0" style="display:none">
        <span class="text-[10px] uppercase tracking-widest text-outline font-data-tabular shrink-0">다음 발언자 ▶</span>
        <div id="tcExperts" class="flex flex-wrap gap-2"></div>
        <button id="summaryBtn" onclick="makeSummary()" class="ml-auto shrink-0 text-xs font-bold border border-secondary/50 text-secondary rounded-lg px-3 py-1.5 hover:bg-secondary hover:text-on-secondary-fixed transition">📝 회의록 요약</button>
      </div>
      <div id="consoleBar" class="px-6 py-4 border-t border-outline-variant/20 bg-surface-container-low/60 flex gap-3 shrink-0 items-center">
        <button id="micBtn" onclick="toggleMic()" title="음성 발언 (Jarvis STT)"
          class="shrink-0 w-11 h-11 rounded-full border border-outline-variant/50 flex items-center justify-center text-on-surface-variant hover:border-secondary hover:text-secondary transition">
          <span class="material-symbols-outlined">mic</span>
        </button>
        <input id="chatInput" onkeydown="if(event.key==='Enter'&&!event.isComposing)sendMessage()" placeholder="진행자로서 직접 발언 — 🎤 버튼 또는 입력 (전송 후 다음 발언자 선택)..." class="flex-1 bg-surface-container-lowest border border-outline-variant/40 rounded-lg px-4 py-2.5 text-sm text-on-surface focus:ring-1 focus:ring-secondary outline-none">
        <button id="voiceModeBtn" onclick="toggleVoiceMode()" title="보이스 회의 — 전문가 발언 낭독 + 자동 듣기"
          class="shrink-0 text-xs font-bold border border-outline-variant/40 rounded-lg px-3 py-2.5 text-on-surface-variant hover:border-secondary hover:text-secondary transition">🔊 보이스</button>
        <button onclick="sendMessage()" class="bg-secondary text-on-secondary-fixed font-bold px-5 py-2.5 rounded-lg flex items-center gap-1.5 hover:opacity-90 transition"><span class="material-symbols-outlined text-sm">send</span>내 발언</button>
      </div>
      <div id="micStatus" class="px-8 pb-2 text-xs font-data-tabular" style="display:none;color:#c8931d"></div>
    </div>

  </div>

  <!-- 회의 결과 보고서 오버레이 -->
  <div id="summaryOverlay">
    <div class="sum-bk" onclick="closeSummary()"></div>
    <div class="sum-doc custom-scrollbar" id="sumDoc"></div>
  </div>
</main>

<script>
const EXPERTS = __EXPERTS_JSON__;
const VIZ = __VIZ_JSON__;
const EXPERT_VIZ = __EXPERTVIZ_JSON__;
const RISK_AGENDA = __AGENDA_JSON__;
const KRISK = __KRISK_JSON__;
let vizSeq = 0;
let recentViz = [];   // 최근 띄운 차트 키 (연속 중복 방지)
let selectedExperts = [];
let selectedAudience = 'investor';
function setAudience(a, el){
  selectedAudience = a;
  document.querySelectorAll('.aud-btn').forEach(function(b){ b.classList.remove('active'); });
  if(el) el.classList.add('active');
}
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
    const kr = KRISK[key];
    const krBadge = kr ? '<span class="text-[10px] font-data-tabular font-bold shrink-0" title="K-RISK 실시간 점수" style="color:'
      + (kr.grade === '위험' ? '#d64545' : (kr.grade === '주의' ? '#b58a12' : '#1e8e5a'))
      + '">' + kr.score + '</span>' : '';
    card.innerHTML = '<span class="text-2xl">'+ex.avatar+'</span>'
      + '<div class="flex-1 min-w-0"><div class="text-sm font-bold truncate" style="color:'+ex.color+'">'+ex.name+'</div>'
      + '<div class="text-[11px] text-on-surface-variant truncate">'+ex.title+'</div></div>'
      + krBadge
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

// 오늘의 리스크 안건 — 클릭 시 AI 추천 전문가 조합 자동 선택 + 안건 프리필
(function(){
  var box = document.getElementById('riskAgenda'), wrap = document.getElementById('riskAgendaWrap');
  if(!box) return;
  if(!RISK_AGENDA.length){ if(wrap) wrap.style.display='none'; return; }
  RISK_AGENDA.forEach(function(a){
    var names = (a.ex || []).map(function(k){ return EXPERTS[k] ? EXPERTS[k].avatar + ' ' + EXPERTS[k].name : k; }).join(' · ');
    var b = document.createElement('button');
    b.className = 'agenda-chip text-xs text-left border border-outline-variant/40 rounded-lg px-3 py-2.5 text-on-surface-variant hover:border-secondary hover:text-secondary transition w-full';
    b.innerHTML = '<span class="font-bold">' + a.q + '</span>'
      + '<span class="block mt-1 text-[10px] text-outline">🤖 AI 추천 조합 — ' + names + '</span>';
    b.onclick = function(){ applyAgenda(a, b); };
    box.appendChild(b);
  });
})();
// 회의록 요약 — 화면 전체를 덮는 보고서형 카드뉴스 오버레이
let _summing = false;
function makeSummary(){
  if(_summing || busy) return;
  if(!chatHistory.some(function(h){ return h.role === 'assistant'; })) return;
  _summing = true;
  var ov = document.getElementById('summaryOverlay'), doc = document.getElementById('sumDoc');
  ov.style.display = 'block';
  doc.innerHTML = '<div style="padding:100px 40px;text-align:center;color:#8a8a95;font-size:14px;line-height:2">📝<br>서기 AI가 회의록을 읽고<br>결과 보고서를 작성하는 중입니다...</div>';
  fetch('/api/conference/summary', {method:'POST', headers:{'Content-Type':'application/json'},
    body: JSON.stringify({history: chatHistory, audience: selectedAudience})})
    .then(function(r){ return r.json(); })
    .then(function(d){
      if(d && d.ok){ renderSummaryDoc(d.summary, doc); }
      else { doc.innerHTML = '<div style="padding:80px 40px;text-align:center;color:#d64545">요약 실패: ' + ((d && d.message) || '알 수 없는 오류') + '</div>'; }
      _summing = false;
    })
    .catch(function(e){
      doc.innerHTML = '<div style="padding:80px 40px;text-align:center;color:#d64545">요약 실패: ' + e + '</div>';
      _summing = false;
    });
}
function closeSummary(){ document.getElementById('summaryOverlay').style.display = 'none'; }
function _escH(s){ return String(s || '').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;'); }
function parseSummary(raw){
  var esc = _escH(raw).replace(/\*\*([^*]+)\*\*/g, '$1');
  var lines = esc.split('\n'), oneline = '', sections = [], cur = null;
  lines.forEach(function(l){
    var t = l.trim();
    if(!t) return;
    if(t.indexOf('한 줄 결론') === 0){ oneline = t.replace(/^한 줄 결론\s*[:：]\s*/, ''); return; }
    if(t.charAt(0) === '■'){ cur = {title: t.slice(1).trim(), items: []}; sections.push(cur); return; }
    if(cur) cur.items.push(t.replace(/^[-·•]\s*|^\d+[)\.]\s*/, ''));
    else if(!oneline) oneline = t;
  });
  return {oneline: oneline, sections: sections};
}

function renderSummaryDoc(raw, doc){
  window._lastSummary = raw;
  var p = parseSummary(raw);
  function chip(t){ return t.replace(/\[([^\[\]]{1,40})\]/g, '<span class="src-chip">$1</span>'); }
  var topicMsg = '';
  for(var i = 0; i < chatHistory.length; i++){ if(chatHistory[i].role === 'user'){ topicMsg = chatHistory[i].content; break; } }
  var AUD_LB = {investor:'📈 일반 투자자', business:'🏢 기업·조달', consumer:'🛒 일반 소비자', policy:'🏛️ 정책·연구'};
  var chips = selectedExperts.map(function(k){ var ex = EXPERTS[k];
    return ex ? '<span class="sum-exp-chip" style="color:'+ex.color+'">'+ex.avatar+' '+ex.name+'</span>' : ''; }).join('');
  var nTalk = chatHistory.filter(function(h){ return h.role === 'assistant'; }).length;
  var srcCount = (String(raw).match(/\[[^\[\]]{1,40}\]/g) || []).length;
  // 안건 관련 K-RISK 게이지 (안건에 광물명이 있으면)
  var krName = null, krD = null;
  Object.keys(KRISK || {}).forEach(function(k){ if(!krName && topicMsg.indexOf(k) !== -1){ krName = k; krD = KRISK[k]; } });
  var gauge = '';
  if(krD){
    var col = krD.grade === '위험' ? '#d64545' : (krD.grade === '주의' ? '#b58a12' : '#1e8e5a');
    var C = 2 * Math.PI * 46, fill = (krD.score / 100 * C).toFixed(1);
    gauge = '<div class="sum-gauge"><svg width="128" height="128" viewBox="0 0 128 128">'
      + '<circle cx="64" cy="64" r="46" fill="none" stroke="rgba(255,255,255,.08)" stroke-width="11"/>'
      + '<circle cx="64" cy="64" r="46" fill="none" stroke="' + col + '" stroke-width="11" stroke-linecap="round" '
      + 'stroke-dasharray="' + fill + ' ' + C.toFixed(1) + '" transform="rotate(-90 64 64)"/>'
      + '<text x="64" y="60" text-anchor="middle" fill="' + col + '" font-size="26" font-weight="900" font-family="JetBrains Mono,monospace">' + krD.score + '</text>'
      + '<text x="64" y="80" text-anchor="middle" fill="#8fa0bd" font-size="11" font-weight="700">' + krD.grade + '</text></svg>'
      + '<div class="gl">' + krName + ' K-RISK</div></div>';
  }
  // 차트: 토론에서 인용된 최근 차트 최대 2개 (없으면 K-RISK 차트)
  var vizKeys = [];
  for(var vi = recentViz.length - 1; vi >= 0 && vizKeys.length < 2; vi--){
    if(vizKeys.indexOf(recentViz[vi]) === -1 && VIZ[recentViz[vi]]) vizKeys.push(recentViz[vi]);
  }
  if(!vizKeys.length && VIZ['k_risk']) vizKeys.push('k_risk');
  // 패널 한마디: 전문가별 마지막 발언 발췌
  var quotes = '';
  selectedExperts.forEach(function(k){
    var ex = EXPERTS[k]; if(!ex) return;
    var last = null;
    chatHistory.forEach(function(h){ if(h.role === 'assistant' && h.name === ex.name) last = h.content; });
    if(!last) return;
    var t = last.replace(/\[\[\s*viz\s*:[^\]]*\]\]/gi, '').trim();
    if(t.length > 110) t = t.slice(0, 110) + '…';
    quotes += '<div class="sum-quote"><div class="av" style="border-color:' + ex.color + '66">' + ex.avatar + '</div>'
      + '<div><div class="nm" style="color:' + ex.color + '">' + ex.name + '</div>'
      + '<div class="tx">“ ' + chip(_escH(t)) + ' ”</div></div></div>';
  });
  var html = '<div class="sum-actions">'
    + '<button onclick="downloadMinutes(window._lastSummary||\'\')">⬇ 회의록 저장</button>'
    + '<button onclick="closeSummary()">✕ 닫기</button></div>';
  html += '<div class="sum-cover"><div class="sum-cover-grid"><div>'
    + '<div class="sum-brand">◆ K-RESOURCE · AI 전문가 회의실</div>'
    + '<div class="sum-doc-title">회의 결과 보고서</div>'
    + '<div class="sum-agenda">안건 — ' + _escH(topicMsg) + '</div>'
    + '<div class="sum-cover-meta"><span>일시 <b>' + new Date().toLocaleString('ko-KR') + '</b></span>'
    + '<span>청중 <b>' + (AUD_LB[selectedAudience] || selectedAudience) + '</b></span></div>'
    + '<div style="margin-top:16px;display:flex;flex-wrap:wrap;gap:8px">' + chips + '</div>'
    + '</div>' + gauge + '</div></div>';
  html += '<div class="sum-statrow">'
    + (krD ? '<div class="sum-stat"><div class="v" style="color:' + (krD.grade==='위험'?'#ff7a7a':(krD.grade==='주의'?'#f2c94c':'#5ad1b0')) + '">' + krD.score + '<small>/100</small></div><div class="l">' + krName + ' K-RISK</div></div>'
           : '<div class="sum-stat"><div class="v">' + selectedExperts.length + '<small>명</small></div><div class="l">참여 전문가</div></div>')
    + '<div class="sum-stat"><div class="v">' + nTalk + '<small>회</small></div><div class="l">전문가 발언</div></div>'
    + '<div class="sum-stat"><div class="v">' + srcCount + '<small>건</small></div><div class="l">인용된 데이터 출처</div></div>'
    + '<div class="sum-stat"><div class="v">' + vizKeys.length + '<small>개</small></div><div class="l">근거 차트</div></div></div>';
  html += '<div class="sum-docbody">';
  if(p.oneline) html += '<div class="sum-hero2">“ ' + chip(p.oneline) + ' ”</div>';
  html += '<div class="sum-cols"><div>';
  p.sections.forEach(function(s){
    if(s.title.indexOf('핵심 결론') !== -1){
      html += '<div class="sum-h">🎯 ' + s.title + '</div><div class="sum-grid3">'
        + s.items.map(function(it, i){ return '<div class="sum-key"><div class="no">0' + (i + 1) + '</div><div class="tx">' + chip(it) + '</div></div>'; }).join('') + '</div>';
    } else if(s.title.indexOf('시사점') !== -1){
      html += '<div class="sum-h">💡 ' + s.title + '</div>'
        + s.items.map(function(it, i){ return '<div class="sum-imp"><div class="bd">' + (i + 1) + '</div><div class="tx">' + chip(it) + '</div></div>'; }).join('');
    } else if(s.title.indexOf('쟁점') !== -1){
      html += '<div class="sum-h">⚔️ ' + s.title + '</div>'
        + s.items.map(function(it){ return '<div class="sum-issue">' + chip(it) + '</div>'; }).join('');
    } else if(s.title.indexOf('후속') !== -1){
      html += '<div class="sum-h">🔍 ' + s.title + '</div><div>'
        + s.items.map(function(it){ return '<span class="sum-next">📎 ' + chip(it) + '</span>'; }).join('') + '</div>';
    } else {
      html += '<div class="sum-h">📌 ' + s.title + '</div>'
        + s.items.map(function(it){ return '<div class="sum-imp"><div class="bd">·</div><div class="tx">' + chip(it) + '</div></div>'; }).join('');
    }
  });
  html += '</div><div>';
  html += '<div class="sum-h">📊 근거 데이터</div><div id="sumVizBox"></div>';
  if(quotes) html += '<div class="sum-h">🗣 패널 한마디</div>' + quotes;
  html += '</div></div>';
  html += '<div class="sum-foot"><span>산업통상부·산하기관 공공데이터 기반 · 모든 수치는 [데이터셋] 출처칩을 따릅니다</span><span>K-RESOURCE 자동 생성 보고서</span></div></div>';
  doc.innerHTML = html;
  var box = document.getElementById('sumVizBox');
  if(box){ vizKeys.forEach(function(k){ renderVizCard(VIZ[k], box, '#e9c349'); }); }
  doc.scrollTop = 0;
}

function downloadMinutes(summary){
  var lines = ['K-RESOURCE AI 전문가 회의록', '일시: ' + new Date().toLocaleString('ko-KR'), '', '[요약]', summary, '', '[전체 회의록]'];
  chatHistory.forEach(function(h){
    lines.push((h.role === 'user' ? '[진행자] ' : '[' + (h.name || '전문가') + '] ') + h.content);
  });
  var blob = new Blob([lines.join('\n')], {type: 'text/plain;charset=utf-8'});
  var a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = 'K-RESOURCE_회의록_' + new Date().toISOString().slice(0,10) + '.txt';
  a.click(); URL.revokeObjectURL(a.href);
}

function applyAgenda(a, btn){
  document.querySelectorAll('.agenda-chip').forEach(function(x){ x.classList.remove('agenda-active'); });
  if(btn) btn.classList.add('agenda-active');
  document.querySelectorAll('.expert-card').forEach(function(c){
    c.classList.toggle('selected', (a.ex || []).indexOf(c.dataset.key) >= 0);
  });
  updateSelection();
  var t = document.getElementById('questionInput');
  if(t) t.value = a.q.replace(/^(🔴|🟡|🟢)\s*/, '');
  var g = document.getElementById('expertGrid');
  if(g) g.scrollIntoView({behavior:'smooth', block:'start'});
}

function showScreen(id) {
  ['step1Screen','step2Screen','roomScreen'].forEach(s => document.getElementById(s).style.display = 'none');
  document.getElementById(id).style.display = 'flex';
  document.body.classList.toggle('cine', id === 'roomScreen');   // 영화 모드
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
  const audLabel = {investor:'📈 일반 투자자', business:'🏢 기업·조달', consumer:'🛒 일반 소비자', policy:'🏛️ 정책·연구'}[selectedAudience] || selectedAudience;
  document.getElementById('activeExperts').innerHTML =
    '<span class="aud-tag">'+audLabel+'</span>' +
    selectedExperts.map(k => {
    const ex = EXPERTS[k];
    return '<div class="exp-seat" id="seat-'+k+'" style="--exc:'+ex.color+'">'
      + '<span class="es-avatar">'+ex.avatar+'</span>'
      + '<span class="es-name">'+ex.name.replace(' 전문가','')+'</span>'
      + '<span class="es-eq"><i></i><i></i><i></i><i></i></span>'
      + '</div>';
  }).join('');
  document.getElementById('chatArea').innerHTML = '';
  recentViz = [];
  appendUserMsg(q);
  busy = false;
  renderTurnControls(true);   // 첫 발언자 직접 선택
}

function backToLobby() {
  showScreen('step1Screen');
  document.getElementById('chatArea').innerHTML = '';
  document.getElementById('turnControls').style.display = 'none';
  chatHistory = [];
  recentViz = [];
  busy = false;
}

function appendUserMsg(text) {
  const chatArea = document.getElementById('chatArea');
  const div = document.createElement('div');
  div.className = 'flex justify-end';
  div.innerHTML = '<div class="max-w-[75%]"><div class="flex items-center justify-end gap-2 mb-1"><span class="text-[11px] font-bold text-secondary">🎙️ 진행자 (나)</span></div><div class="msg-bubble user-msg-bubble bg-secondary/10 border border-secondary/30 rounded-xl rounded-tr-none px-4 py-3 text-sm text-on-surface leading-relaxed"></div></div>';
  div.querySelector('.msg-bubble').textContent = text;
  chatArea.appendChild(div);
  chatArea.scrollTop = chatArea.scrollHeight;
  chatHistory.push({role:'user', content:text});
}

function renderTurnControls(first) {
  const tc = document.getElementById('turnControls');
  const box = document.getElementById('tcExperts');
  const lbl = tc.querySelector('span');
  if (lbl) lbl.textContent = first ? '첫 발언자 ▶' : '다음 발언자 ▶';
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
    body: JSON.stringify({speaker: key, history: chatHistory, audience: selectedAudience, recentViz: recentViz})
  }).then(r => {
    if (!r.ok) {
      ti.style.display = 'none'; busy = false; renderTurnControls();
      const chatArea = document.getElementById('chatArea');
      const warn = document.createElement('div');
      warn.className = 'text-xs font-bold px-4 py-3 rounded-lg border';
      warn.style.cssText = 'color:#ff7a7a;border-color:#ff7a7a55;background:#ff7a7a11';
      warn.textContent = (r.status === 401)
        ? '⚠️ 세션이 만료되었습니다. 회의실을 나갔다가 다시 로그인해 주세요.'
        : ('⚠️ 서버 오류 (HTTP ' + r.status + ') — 잠시 후 다시 시도해 주세요.');
      chatArea.appendChild(warn); chatArea.scrollTop = chatArea.scrollHeight;
      return;
    }
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
              hideHud();
              setSpeaking(d.speaker_start, true);
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
              if (document.body.classList.contains('cine')) {
                const cap = document.getElementById('liveCaption');
                if (cap) { cap.textContent = currentBubble.textContent; cap.scrollTop = cap.scrollHeight; }
              }
              chatArea.scrollTop = chatArea.scrollHeight;
            } else if (d.error) {
              if (currentBubble) {
                currentBubble.textContent = '⚠️ 발언 생성 실패: ' + d.error;
                currentBubble.style.color = '#ff7a7a';
              }
            } else if (d.speaker_end) {
              if (currentBubble) {
                var _btxt = currentBubble.textContent;
                // 시각자료 태그 추출: 모델이 [[viz:키]]를 명시한 경우만, 최근 띄운 차트면 생략
                var _vizKey = null;
                var _vm = _btxt.match(/\[\[\s*viz\s*:\s*([^\]]+?)\s*\]\]/i);
                if (_vm) { _vizKey = _vm[1].trim(); _btxt = _btxt.replace(_vm[0], '').trim(); }
                if (_vizKey && recentViz.indexOf(_vizKey) !== -1) { _vizKey = null; }
                chatHistory.push({role:'assistant', name:EXPERTS[d.speaker_end]?.name||d.speaker_end, content:_btxt});
                // 텍스트 렌더: HTML 이스케이프 → **굵게** → [출처칩]
                var _esc = _btxt.replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
                _esc = _esc.replace(/\*\*([^*]+)\*\*/g, '<strong>$1</strong>');
                _esc = _esc.replace(/\[([^\[\]]{1,40})\]/g, '<span class="src-chip">$1</span>');
                currentBubble.innerHTML = _esc;
                if (_vizKey && VIZ[_vizKey]) {
                  var _exc = (EXPERTS[d.speaker_end]||{}).color;
                  if (document.body.classList.contains('cine')) {
                    showHudViz(VIZ[_vizKey], _exc);
                  } else {
                    renderVizCard(VIZ[_vizKey], currentBubble.parentNode, _exc);
                  }
                  // 채팅 헤더에 '시각자료' 버튼 — 언제든 다시 보기
                  var _hd = currentBubble.parentNode.querySelector('.flex.items-baseline');
                  if (_hd) {
                    var _vb = document.createElement('button');
                    _vb.className = 'viz-btn';
                    _vb.textContent = '📊 시각자료';
                    (function(spec, col){ _vb.onclick = function(){
                      if (document.body.classList.contains('cine')) showHudViz(spec, col);
                      else { renderVizCard(spec, currentBubble ? currentBubble.parentNode : _hd.parentNode, col); }
                    }; })(VIZ[_vizKey], _exc);
                    _hd.appendChild(_vb);
                  }
                  recentViz.push(_vizKey); if (recentViz.length > 3) recentViz.shift();
                  document.getElementById('chatArea').scrollTop = 1e9;
                }
                voiceSpeak(_btxt, d.speaker_end);
              }
              currentBubble = null;
              if (!voiceMode) setSpeaking(null, false);
            }
          } catch(e) {}
        });
        read();
      });
    }
    read();
  }).catch(e => { ti.style.display = 'none'; busy = false; renderTurnControls(); console.error(e); });
}

// 발언 근거 시각자료 카드 렌더 (말풍선 아래)
let _hudTimer = null;
function hideHud(){
  var dock = document.getElementById('hudDock');
  if (dock) dock.classList.remove('show');
  if (_hudTimer){ clearTimeout(_hudTimer); _hudTimer = null; }
}
function showHudViz(spec, color){
  var dock = document.getElementById('hudDock'); if(!dock) return;
  var body = document.getElementById('hudBody');
  body.innerHTML = '';
  dock.style.setProperty('--exc', color || '#5fd0ff');
  document.getElementById('hudSrc').textContent = spec.source || '';
  renderVizCard(spec, body, color);
  dock.classList.remove('show'); void dock.offsetWidth;   // 애니메이션 리셋
  dock.classList.add('show');
  if (_hudTimer) clearTimeout(_hudTimer);
  if (!voiceMode) _hudTimer = setTimeout(hideHud, 14000);   // 음성 없을 땐 14초 후 자동 정리
}

// 기준선(hlines) 플러그인 — 위험/안정 가이드
var _hlinePlugin = {
  id: 'hlines',
  afterDraw: function(chart, args, opts){
    var lines = (opts && opts.lines) || [];
    if (!lines.length) return;
    var c = chart.ctx, area = chart.chartArea, yS = chart.scales.y;
    lines.forEach(function(h){
      if (h.v < yS.min || h.v > yS.max) return;
      var y = yS.getPixelForValue(h.v);
      c.save();
      c.strokeStyle = h.color || '#e06060'; c.globalAlpha = .55;
      c.setLineDash([5, 5]); c.lineWidth = 1;
      c.beginPath(); c.moveTo(area.left, y); c.lineTo(area.right, y); c.stroke();
      c.setLineDash([]); c.globalAlpha = .9;
      c.font = '9px Pretendard, sans-serif'; c.fillStyle = h.color || '#e06060';
      c.fillText(h.label || '', area.right - c.measureText(h.label || '').width - 3, y - 4);
      c.restore();
    });
  }
};

function _hexA(hex, a){       // #rrggbb → rgba
  var n = parseInt(hex.slice(1), 16);
  return 'rgba('+(n>>16&255)+','+(n>>8&255)+','+(n&255)+','+a+')';
}

function renderVizCard(spec, container, color){
  if(!spec || !container) return;
  color = color || '#e9c349';
  var cine = document.body.classList.contains('cine');
  var grid = cine ? 'rgba(126,166,255,.09)' : '#e8ecf3';
  var tickc = cine ? '#8fa6c6' : '#6f7b90';
  var head = '<div class="viz-head"><span class="viz-dot" style="background:'+color+'"></span>'+(spec.title||'')+'</div>';
  // 헤드라인 KPI (큰 숫자 + 서브)
  var kpi = spec.headline ? ('<div class="viz-kpi"><b style="color:'+color+'">'+spec.headline.value+'</b>'
    + '<span><i>'+(spec.headline.label||'')+'</i>'+(spec.headline.sub ? '<em>'+spec.headline.sub+'</em>' : '')+'</span></div>') : '';
  var foot = '<div class="viz-foot">'+[spec.note, spec.source].filter(Boolean).join(' · ')+'</div>';
  var card = document.createElement('div');
  card.className = 'viz-card';
  if(spec.type === 'stat'){
    var body = '<div class="viz-stats">'+(spec.stats||[]).map(function(s){
      var col = ((s.delta||'').indexOf('-')===0) ? '#ff8a8a' : '#5ad1b0';
      return '<div class="viz-stat"><div class="vs-label">'+s.label+'</div>'
        +'<div class="vs-val">'+s.value+'<span>'+(s.unit||'')+'</span></div>'
        +(s.delta ? ('<div class="vs-delta" style="color:'+col+'">'+s.delta+'</div>') : '')+'</div>';
    }).join('')+'</div>';
    card.innerHTML = head + kpi + body + foot;
    container.appendChild(card);
    return;
  }
  var cid = 'vz'+(vizSeq++);
  card.innerHTML = head + kpi + '<div class="viz-canvas"><canvas id="'+cid+'"></canvas></div>' + foot;
  container.appendChild(card);
  if(typeof Chart === 'undefined') return;
  var ctx = document.getElementById(cid), type = spec.type || 'line';
  if(type === 'doughnut'){
    var pal = ['#e9c349','#5fd0ff','#5ad1b0','#f472b6','#f59e0b','#9b8cff'];
    new Chart(ctx, {type:'doughnut',
      data:{labels:spec.labels, datasets:[{data:(spec.series[0]||{}).data||[], backgroundColor:pal,
        borderColor:(cine?'#0a1222':'#ffffff'), borderWidth:2, hoverOffset:6}]},
      options:{responsive:true, maintainAspectRatio:false, cutout:'62%',
        plugins:{legend:{position:'right', labels:{color:tickc, font:{size:10}, boxWidth:10}}}}});
    return;
  }
  var bands = spec.bands;
  var ds = (spec.series||[]).map(function(s, si){
    var col = s.color || color;
    var bg;
    if (type === 'bar') {
      bg = bands ? (s.data||[]).map(function(v){
        var bi = 0; (bands.edges||[]).forEach(function(e){ if (v >= e) bi++; });
        return bands.colors[bi] || col;
      }) : (s.data||[]).map(function(){ return _hexA(col, .85); });
    } else {
      bg = function(c2){                       // 라인 아래 글로우 그라데이션
        var area = c2.chart.chartArea; if(!area) return _hexA(col, .06);
        var g2 = c2.chart.ctx.createLinearGradient(0, area.top, 0, area.bottom);
        g2.addColorStop(0, _hexA(col, cine ? .32 : .2)); g2.addColorStop(1, _hexA(col, 0));
        return g2;
      };
    }
    var n = (s.data||[]).length;
    return {label:s.name, data:s.data, borderColor:col, backgroundColor:bg,
      fill:(type!=='bar' && si===0), borderWidth:2, tension:.3,
      pointRadius:(type==='bar'?0:(s.data||[]).map(function(_,i){ return i===n-1?4:0; })),
      pointBackgroundColor:col, pointBorderColor:(cine?'#0a1222':'#fff'), pointBorderWidth:2,
      borderRadius:6, maxBarThickness:26};
  });
  new Chart(ctx, {type:type, data:{labels:spec.labels, datasets:ds},
    plugins:[_hlinePlugin],
    options:{responsive:true, maintainAspectRatio:false, interaction:{mode:'index', intersect:false},
      plugins:{legend:{display:(ds.length>1), labels:{color:tickc, font:{size:9}, boxWidth:9}},
               hlines:{lines:spec.hlines||[]}},
      scales:{x:{ticks:{color:tickc, font:{size:9}, maxTicksLimit:8}, grid:{color:grid, drawTicks:false}},
              y:{ticks:{color:tickc, font:{size:9}}, grid:{color:grid, drawTicks:false}, border:{display:false}}}}});
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

/* ═══ Jarvis 음성 회의 — STT(마이크) + TTS(낭독) ═══ */
let _mic = null;          // {ctx, proc, src, stream, chunks[], rate, hadVoice, silentMs}
let voiceMode = false;
const SIL_TH = 0.013, SIL_MS = 1800, MAX_MS = 30000;

function _micStatus(t){ const el=document.getElementById('micStatus');
  if(!el) return; el.style.display = t ? 'block' : 'none'; el.textContent = t || ''; }

function toggleMic(){ _mic ? stopMic(true) : startMic(); }

async function startMic(){
  if (_mic || busy) return;
  try {
    const stream = await navigator.mediaDevices.getUserMedia({audio:{echoCancellation:true, noiseSuppression:true}});
    const ctx = new (window.AudioContext||window.webkitAudioContext)();
    const src = ctx.createMediaStreamSource(stream);
    const proc = ctx.createScriptProcessor(4096, 1, 1);
    _mic = {ctx, proc, src, stream, chunks:[], rate:ctx.sampleRate, hadVoice:false, silentMs:0, startedAt:Date.now()};
    proc.onaudioprocess = (e) => {
      if (!_mic) return;
      const buf = e.inputBuffer.getChannelData(0);
      _mic.chunks.push(new Float32Array(buf));
      let sum = 0; for (let i=0;i<buf.length;i++) sum += buf[i]*buf[i];
      const rms = Math.sqrt(sum/buf.length);
      const frameMs = buf.length/_mic.rate*1000;
      if (rms > SIL_TH) { _mic.hadVoice = true; _mic.silentMs = 0; }
      else if (_mic.hadVoice) { _mic.silentMs += frameMs; }
      if ((_mic.hadVoice && _mic.silentMs > SIL_MS) || (Date.now()-_mic.startedAt > MAX_MS)) stopMic(true);
    };
    src.connect(proc); proc.connect(ctx.destination);
    document.getElementById('micBtn').classList.add('rec');
    _micStatus('🎤 듣는 중… 말씀이 끝나면 자동으로 전송됩니다 (Jarvis STT)');
  } catch(e) {
    _micStatus(''); alert('마이크 권한이 필요합니다: ' + e.message);
  }
}

function _teardownMic(){
  if (!_mic) return null;
  const m = _mic; _mic = null;
  try { m.proc.disconnect(); m.src.disconnect(); m.stream.getTracks().forEach(t=>t.stop()); m.ctx.close(); } catch(e){}
  document.getElementById('micBtn').classList.remove('rec');
  return m;
}

function stopMic(send){
  const m = _teardownMic();
  if (!m || !send) { _micStatus(''); return; }
  if (!m.hadVoice || !m.chunks.length) { _micStatus(''); return; }
  _micStatus('⏳ Jarvis가 알아듣는 중…');
  const wav = _toWav16k(m.chunks, m.rate);
  fetch('/api/conference/stt', {method:'POST', headers:{'Content-Type':'audio/wav'}, body:wav})
    .then(r=>r.json()).then(d=>{
      _micStatus('');
      if (d && d.ok && d.text) {
        const input = document.getElementById('chatInput');
        input.value = d.text;
        sendMessage();
      } else if (d && d.error === 'stt_unavailable') {
        _micStatus('⚠️ Jarvis STT 서버가 꺼져 있어요 — 터미널에서 ./jarvis_stt.sh 실행 후 다시 시도');
      } else { _micStatus('… 음성을 인식하지 못했어요. 다시 눌러 말씀해주세요'); }
    }).catch(()=>{ _micStatus('⚠️ STT 연결 실패 — ./jarvis_stt.sh 확인'); });
}

function _toWav16k(chunks, rate){
  let len = 0; chunks.forEach(c=>len+=c.length);
  let audio = new Float32Array(len); let o=0;
  chunks.forEach(c=>{ audio.set(c,o); o+=c.length; });
  if (rate !== 16000) {                       // 선형 리샘플 → 16k
    const n = Math.round(len*16000/rate), out = new Float32Array(n);
    for (let i=0;i<n;i++){ const x=i*(len-1)/(n-1), i0=Math.floor(x), f=x-i0;
      out[i] = audio[i0]*(1-f) + (audio[i0+1]||audio[i0])*f; }
    audio = out;
  }
  const pcm = new Int16Array(audio.length);
  for (let i=0;i<audio.length;i++){ const v=Math.max(-1,Math.min(1,audio[i])); pcm[i]=v<0?v*32768:v*32767; }
  const buf = new ArrayBuffer(44+pcm.length*2), dv = new DataView(buf);
  const wstr=(off,s)=>{for(let i=0;i<s.length;i++)dv.setUint8(off+i,s.charCodeAt(i));};
  wstr(0,'RIFF'); dv.setUint32(4,36+pcm.length*2,true); wstr(8,'WAVE'); wstr(12,'fmt ');
  dv.setUint32(16,16,true); dv.setUint16(20,1,true); dv.setUint16(22,1,true);
  dv.setUint32(24,16000,true); dv.setUint32(28,32000,true); dv.setUint16(32,2,true); dv.setUint16(34,16,true);
  wstr(36,'data'); dv.setUint32(40,pcm.length*2,true);
  new Int16Array(buf,44).set(pcm);
  return buf;
}

/* ── TTS: 전문가 발언 낭독 — OpenAI 목소리 11종 (전문가별) + 브라우저 폴백 ── */
let _koVoice = null, _curAudio = null;
function _pickVoice(){
  const vs = speechSynthesis.getVoices();
  _koVoice = vs.find(v=>v.lang==='ko-KR' && /Yuna|유나/i.test(v.name)) || vs.find(v=>v.lang&&v.lang.indexOf('ko')===0) || null;
}
if (window.speechSynthesis){ _pickVoice(); speechSynthesis.onvoiceschanged = _pickVoice; }

function setSpeaking(key, on){
  document.querySelectorAll('.exp-seat.speaking').forEach(el=>{ if(!on || el.id!=='seat-'+key) el.classList.remove('speaking'); });
  if (on && key){ const el=document.getElementById('seat-'+key); if(el) el.classList.add('speaking'); }
  const cap = document.getElementById('liveCaption');
  if (cap){
    if (on && key){ cap.style.setProperty('--exc', (EXPERTS[key]||{}).color || '#7ea6ff'); cap.textContent=''; cap.classList.add('on'); }
    else { cap.classList.remove('on'); }
  }
  const st = document.getElementById('stageOrb');
  if (st){
    if (on && key){
      const ex = EXPERTS[key] || {};
      st.style.setProperty('--exc', ex.color || '#e9c667');
      st.querySelector('.so-avatar').textContent = ex.avatar || '◆';
      st.querySelector('.so-name').textContent = ex.name || key;
      st.querySelector('.so-role').textContent = ex.title || '';
      st.classList.add('on');
    } else {
      st.classList.remove('on');
    }
  }
  if (window._holoSet){
    if (on && key) _holoSet((EXPERTS[key]||{}).color || '#4fd8ff', true);
    else _holoSet('#4fd8ff', false);
  }
}
function _afterSpeak(){
  if (voiceMode && !busy && !_mic) setTimeout(()=>{ if(voiceMode && !busy && !_mic) startMic(); }, 350);
}
function _cleanForSpeech(text){
  return text.replace(/\[\[\s*viz[^\]]*\]\]/gi,'').replace(/\[[^\[\]]{1,40}\]/g,'').replace(/\*\*/g,'').trim();
}
function voiceSpeak(text, exKey){
  if (!voiceMode) return;
  const clean = _cleanForSpeech(text);
  if (!clean) return;
  _micStatus('🔊 ' + ((EXPERTS[exKey]||{}).name || '전문가') + ' 발언 중…');
  fetch('/api/conference/tts', {method:'POST', headers:{'Content-Type':'application/json'},
    body: JSON.stringify({text: clean, speaker: exKey})})
    .then(r => { if(!r.ok) throw new Error('tts'); return r.blob(); })
    .then(b => {
      if (!voiceMode) return;
      if (_curAudio){ try{_curAudio.pause();}catch(e){} }
      const a = new Audio(URL.createObjectURL(b));
      _curAudio = a;
      a.playbackRate = 1.35;                     // 토론 템포 — 빠르게
      a.onended = () => { setSpeaking(null,false); hideHud(); _micStatus(''); _afterSpeak(); };
      a.onerror = () => { setSpeaking(null,false); _micStatus(''); _afterSpeak(); };
      a.play().catch(()=>{ _micStatus(''); _fallbackSpeak(clean, exKey); });
    })
    .catch(() => { _micStatus(''); _fallbackSpeak(clean, exKey); });
}
function _fallbackSpeak(clean, exKey){          // OpenAI TTS 실패 시 브라우저 TTS
  if (!voiceMode || !window.speechSynthesis) return;
  const u = new SpeechSynthesisUtterance(clean);
  if (_koVoice) u.voice = _koVoice;
  u.lang = 'ko-KR'; u.rate = 1.15;
  let h = 0; for (let i=0;i<(exKey||'').length;i++) h = (h*31 + exKey.charCodeAt(i)) & 1023;
  u.pitch = 0.85 + (h % 8) * 0.05;
  u.onend = _afterSpeak;
  speechSynthesis.speak(u);
}

function toggleVoiceMode(){
  voiceMode = !voiceMode;
  document.getElementById('voiceModeBtn').classList.toggle('on', voiceMode);
  if (voiceMode){ _micStatus('🔊 보이스 회의 모드 — 전문가 발언을 낭독하고, 끝나면 자동으로 마이크가 켜집니다'); }
  else { speechSynthesis.cancel(); if(_curAudio){try{_curAudio.pause();}catch(e){}} stopMic(false); _micStatus(''); }
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
<script type="importmap">{"imports":{"three":"https://unpkg.com/three@0.160.0/build/three.module.js"}}</script>
<script type="module">
import * as THREE from 'three';
const cv = document.getElementById('holoCanvas');
if (cv) {
  const SZ = 640;
  const renderer = new THREE.WebGLRenderer({canvas:cv, alpha:true, antialias:true});
  renderer.setPixelRatio(Math.min(devicePixelRatio, 2));
  renderer.setSize(SZ, SZ, false);
  const scene = new THREE.Scene();
  const cam = new THREE.PerspectiveCamera(42, 1, .1, 50); cam.position.z = 4.4;
  const g = new THREE.Group(); scene.add(g);
  const C = new THREE.Color('#4fd8ff');           // 현재 색
  const target = new THREE.Color('#4fd8ff');      // 목표 색 (전문가)
  const mats = [];
  function M(op){ const m = new THREE.LineBasicMaterial({color:C.clone(), transparent:true, opacity:op, blending:THREE.AdditiveBlending, depthWrite:false}); mats.push(m); return m; }
  // 와이어프레임 구체 (겉·속) — 점 구름의 뼈대 역할, 은은하게
  const outer = new THREE.LineSegments(new THREE.WireframeGeometry(new THREE.IcosahedronGeometry(1, 2)), M(.13)); g.add(outer);
  const inner = new THREE.LineSegments(new THREE.WireframeGeometry(new THREE.IcosahedronGeometry(.62, 1)), M(.28)); g.add(inner);
  // 자비스식 점 구름 — 무수한 입자로 이루어진 구
  function pointSphere(n, r, jitter){
    const a = new Float32Array(n*3);
    for (let i=0;i<n;i++){
      const phi = Math.acos(1 - 2*(i+.5)/n), th = Math.PI*(1+Math.sqrt(5))*i;   // 피보나치 분포
      const rr = r*(1 + (Math.random()-.5)*jitter);
      a[i*3]   = rr*Math.sin(phi)*Math.cos(th);
      a[i*3+1] = rr*Math.cos(phi);
      a[i*3+2] = rr*Math.sin(phi)*Math.sin(th);
    }
    const geo = new THREE.BufferGeometry();
    geo.setAttribute('position', new THREE.BufferAttribute(a, 3));
    return geo;
  }
  const cloudMat = new THREE.PointsMaterial({color:C.clone(), size:.013, transparent:true, opacity:.75, blending:THREE.AdditiveBlending, depthWrite:false});
  mats.push(cloudMat);
  g.add(new THREE.Points(pointSphere(2800, 1.0, .02), cloudMat));      // 촘촘한 표면 입자
  const cloud2Mat = new THREE.PointsMaterial({color:C.clone(), size:.02, transparent:true, opacity:.9, blending:THREE.AdditiveBlending, depthWrite:false});
  mats.push(cloud2Mat);
  g.add(new THREE.Points(pointSphere(260, 1.005, .01), cloud2Mat));    // 밝은 포인트
  const hazeMat = new THREE.PointsMaterial({color:C.clone(), size:.01, transparent:true, opacity:.3, blending:THREE.AdditiveBlending, depthWrite:false});
  mats.push(hazeMat);
  g.add(new THREE.Points(pointSphere(900, .82, .5), hazeMat));         // 내부 안개 입자
  // 궤도 링 2개 (스우시)
  const ring1 = new THREE.Mesh(new THREE.TorusGeometry(1.3, .005, 8, 160),
    new THREE.MeshBasicMaterial({color:C.clone(), transparent:true, opacity:.6, blending:THREE.AdditiveBlending, depthWrite:false}));
  mats.push(ring1.material); ring1.rotation.x = 1.25; g.add(ring1);
  const ring2 = ring1.clone(); ring2.material = ring1.material.clone(); mats.push(ring2.material);
  ring2.rotation.x = 1.9; ring2.rotation.y = .6; ring2.scale.setScalar(1.12); g.add(ring2);
  // 중심 글로우 스프라이트
  function glowTex(){ const c = document.createElement('canvas'); c.width = c.height = 128;
    const x = c.getContext('2d'); const gr = x.createRadialGradient(64,64,0,64,64,64);
    gr.addColorStop(0,'rgba(255,255,255,.9)'); gr.addColorStop(.25,'rgba(120,220,255,.35)'); gr.addColorStop(1,'rgba(120,220,255,0)');
    x.fillStyle = gr; x.fillRect(0,0,128,128); return new THREE.CanvasTexture(c); }
  const glowMat = new THREE.SpriteMaterial({map:glowTex(), color:C.clone(), transparent:true, opacity:.5, blending:THREE.AdditiveBlending, depthWrite:false});
  mats.push(glowMat);
  const glow = new THREE.Sprite(glowMat); glow.scale.setScalar(2.2); g.add(glow);
  // 바닥 투사빔
  const beamMat = new THREE.MeshBasicMaterial({color:C.clone(), transparent:true, opacity:.05, blending:THREE.AdditiveBlending, depthWrite:false, side:THREE.DoubleSide});
  mats.push(beamMat);
  const beam = new THREE.Mesh(new THREE.ConeGeometry(.9, 1.6, 40, 1, true), beamMat);
  beam.position.y = -1.55; scene.add(beam);
  const baseGlow = new THREE.Sprite(glowMat.clone()); mats.push(baseGlow.material);
  baseGlow.scale.set(1.6,.5,1); baseGlow.position.y = -2.2; scene.add(baseGlow);

  // 유기적 꿀렁임 — 각 지오메트리의 원본 좌표 저장 후 노이즈 파동으로 방사형 변형
  const deformables = [];
  g.traverse(o => {
    if ((o.isPoints || o.isLineSegments) && o.geometry && o.geometry.attributes.position) {
      o.geometry.userData.base = o.geometry.attributes.position.array.slice();
      deformables.push(o.geometry);
    }
  });
  function organic(t, amp){
    deformables.forEach(geo => {
      const pos = geo.attributes.position.array, base = geo.userData.base;
      for (let i = 0; i < pos.length; i += 3) {
        const x = base[i], y = base[i+1], z = base[i+2];
        const n = Math.sin(2.2*x + t*.7) * .5
                + Math.sin(2.6*y - t*.55 + 1.7) * .3
                + Math.sin(2.9*z + t*.62 + 3.1) * .2
                + Math.sin(1.3*(x+y+z) + t*.9) * .25;
        const sc = 1 + amp * n;
        pos[i] = x*sc; pos[i+1] = y*sc; pos[i+2] = z*sc;
      }
      geo.attributes.position.needsUpdate = true;
    });
  }
  let speaking = false, amp = .02;
  window._holoSet = (hex, on) => { try{ target.set(hex); }catch(e){} speaking = !!on; };
  const clock = new THREE.Clock();
  let _mf = 0;
  window._holoFrame = () => {                      // 단일 프레임 (외부 강제 렌더용)
    _mf += .13;
    C.lerp(target, .2);
    mats.forEach(m => m.color.copy(C));
    g.rotation.y += .008; inner.rotation.y -= .012; inner.rotation.x = .35;
    ring1.rotation.z += .004; ring2.rotation.z -= .005;
    organic(_mf, speaking ? .075 : .03);
    renderer.render(scene, cam);
  };
  (function tick(){
    requestAnimationFrame(tick);
    if (!document.body.classList.contains('cine')) return;
    const t = clock.getElapsedTime();
    C.lerp(target, .06);
    mats.forEach(m => m.color.copy(C));
    const spd = speaking ? 1 : .5;
    g.rotation.y += (speaking ? .0028 : .0013);        // 부드러운 유영
    inner.rotation.y -= .0022 * spd; inner.rotation.x = .35;
    ring1.rotation.z += .0009 * spd; ring2.rotation.z -= .0013 * spd;
    g.rotation.z = speaking ? .03*Math.sin(t*.9) : .015*Math.sin(t*.5);
    g.position.y = .04*Math.sin(t*.5);
    // 유기적 꿀렁임 — 발언 중 진폭 상승 (부드럽게 전환)
    amp += ((speaking ? .075 : .022) - amp) * .04;
    organic(t, amp);
    glowMat.opacity = speaking ? .46+.08*Math.sin(t*1.8) : .34;
    outer.material.opacity = speaking ? .18 : .12;
    renderer.render(scene, cam);
  })();
}
</script>
</body>
</html>"""
    return (PAGE.replace("__EXPERTS_JSON__", experts_json)
                .replace("__VIZ_JSON__", viz_json)
                .replace("__EXPERTVIZ_JSON__", expertviz_json)
                .replace("__AGENDA_JSON__", agenda_json)
                .replace("__KRISK_JSON__", krisk_json)
                .replace("__NOW__", now))



# ═══════════════════════════════════════════════════════════════════════
#  V2 소비자 UI — "광물 날씨" 리디자인 (2026-07)
#  홈(/) · 광종 상세(/m/이름) · 브리핑(/briefing) — 일반 사용자용
#  이전 대시보드는 /pro, 카테고리 화면은 /dashboard 에 유지(전문가용)
# ═══════════════════════════════════════════════════════════════════════
from urllib.parse import quote as _v2q

# 생활 카테고리 — 일반 사용자 언어로 광종을 묶는다 (한 광종이 여러 곳에 속할 수 있음)
LIFE_CATS = [
    ("bat",  "배터리·전기차", ["리튬", "코발트", "니켈", "망간", "흑연", "알루미늄", "동(구리)",
                          "네오디뮴", "프라세오디뮴", "디스프로슘", "터븀"]),
    ("chip", "반도체·전자",  ["규소", "갈륨", "게르마늄", "인듐", "탄탈륨", "주석", "은", "팔라듐",
                          "안티모니", "지르코늄", "셀레늄", "이트륨", "세륨", "란탄", "유로퓸",
                          "가돌리늄", "에르븀", "홀뮴", "루테튬", "스칸듐"]),
    ("ind",  "자동차·중공업", ["철/철광석", "크롬", "몰리브덴", "텅스텐", "바나듐", "니오븀", "티타늄",
                          "마그네슘", "아연", "연(납)", "사마륨", "스트론튬", "창연/비스무트"]),
    ("ene",  "에너지",      ["우라늄", "유연탄"]),
    ("gold", "귀금속·자산",  ["금", "은", "백금", "팔라듐"]),
]

# 광종별 쉬운 용도 한 줄 (소비자 언어)
MIN_USES = {
    "니켈": "배터리·스테인리스", "동(구리)": "전선·전자제품", "알루미늄": "차체·캔·창호",
    "주석": "납땜·통조림", "연(납)": "자동차 배터리", "아연": "철 도금(부식 방지)",
    "리튬": "전기차 배터리", "코발트": "배터리 양극재", "망간": "철강·배터리",
    "니오븀": "고강도 철강", "규소": "반도체 웨이퍼", "마그네슘": "경량 차체·노트북",
    "몰리브덴": "특수강", "바나듐": "고강도 강철·ESS", "티타늄": "항공기·임플란트",
    "텅스텐": "절삭공구·방산", "안티모니": "난연제", "창연/비스무트": "의약품·합금",
    "크롬": "스테인리스", "갈륨": "반도체·LED", "인듐": "터치스크린",
    "탄탈륨": "스마트폰 콘덴서", "지르코늄": "원자로·세라믹", "스트론튬": "불꽃·자석",
    "셀레늄": "유리·태양전지", "게르마늄": "광섬유·적외선",
    "네오디뮴": "전기차 모터 자석", "세륨": "유리 연마·촉매", "란탄": "카메라 렌즈·촉매",
    "디스프로슘": "고온 자석 첨가", "터븀": "자석·형광체", "스칸듐": "경량 합금",
    "이트륨": "LED·세라믹", "루테튬": "의료 PET", "프라세오디뮴": "자석·합금",
    "사마륨": "자석·원자로", "유로퓸": "형광체(빨강)", "가돌리늄": "MRI 조영제",
    "에르븀": "광통신 증폭", "홀뮴": "레이저·자석",
    "우라늄": "원자력 연료", "유연탄": "발전·제철",
    "철/철광석": "철강 원료", "흑연": "배터리 음극재", "백금": "촉매·수소차",
    "팔라듐": "자동차 촉매", "금": "반도체·자산", "은": "전자·태양광",
}


# ── 전 광종 전문가 자동 증강 — 수기 전문가(리튬 등 7인)는 유지, 나머지 광종은 템플릿 생성 ──
_GEN_COLORS = ["#155BB8", "#0E7A4F", "#B8720A", "#7A5195", "#C0392B", "#2E8B8B",
               "#8A5A2B", "#4A3AA7", "#1E74D8", "#D2611E"]
_GEN_AVATAR = {"비철금속": "🔩", "희소금속": "💠", "희토류": "⚛️", "에너지": "⚡", "기타": "⛏️"}

def _gen_expert(nm, cat, use, idx):
    stance = ("리스크를 먼저 경고하되 근거 수치를 반드시 제시" if idx % 3 == 0 else
              ("기회 요인과 대체 공급선을 균형 있게 제시" if idx % 3 == 1 else
               "가격·수급 데이터의 추세 해석에 집중"))
    return {
        "name": f"{nm} 전문가",
        "title": f"{use} 공급망 분석가",
        "avatar": _GEN_AVATAR.get(cat, "⛏️"),
        "color": _GEN_COLORS[idx % len(_GEN_COLORS)],
        "category": cat,
        "model": DEFAULT_OPENAI_MODEL,
        "api_key": "",
        "system": (f"당신은 '{nm} 전문가'입니다. 한국의 {nm}({use}) 공급망을 전담 분석합니다.\n"
                   f"전문 분야: {nm}의 용도({use})와 수요 산업, 한국 수입 구조·주요 공급국, 가격 동향, 분류({cat}) 시장 맥락.\n"
                   f"성격: {stance}. 제공된 팩트카드·차트 범위의 수치만 인용한다.\n"
                   "다중 토론 지침: 다른 전문가 발언을 직접 인용해 동의/반박하고, 200자 내외로 핵심만 말한다."),
    }

_idx = 0
for _cat, _names in MINERAL_TAXONOMY.items():
    for _nm in _names:
        _base = re.sub(r"[\(\)/].*$", "", _nm)
        if _base in MINERAL_EXPERTS or _nm in MINERAL_EXPERTS:
            _idx += 1
            continue
        MINERAL_EXPERTS[_base] = _gen_expert(_base, _cat, MIN_USES.get(_nm, _cat), _idx)
        _idx += 1


def _v2_norm(s):
    return re.sub(r"[\s\(\)/·]", "", str(s or ""))


def _v2_lookup(d, name):
    """이름 이형(동/동(구리), 철/철광석 등)을 흡수하는 사전 조회."""
    if not isinstance(d, dict) or not d:
        return None
    if name in d:
        return d[name]
    n = _v2_norm(name)
    for k, v in d.items():
        k2 = _v2_norm(k)
        if not k2:
            continue
        if k2 == n:
            return v
        if len(k2) >= 2 and len(n) >= 2 and (k2 in n or n in k2):
            return v
    return None



_V2_REE_SET   = set(MINERAL_TAXONOMY["희토류"])
_V2_STRAT_SET = {"갈륨", "게르마늄", "인듐", "안티모니", "창연/비스무트", "마그네슘", "지르코늄", "티타늄"}
_V2_PREC_SET  = {"금", "은", "백금", "팔라듐"}

_V2_USGS1_CACHE = None
def _v2_usgs1():
    global _V2_USGS1_CACHE
    if _V2_USGS1_CACHE is None:
        _V2_USGS1_CACHE = load_json(os.path.join(os.path.dirname(__file__), "usgs_data1.json")) or {}
    return _V2_USGS1_CACHE


def _v2_trade_groups():
    """관세청 그룹 스냅샷 일괄 로드(스냅샷 우선이라 즉시 응답)."""
    return {
        "core":      (fetch_core_trade() or {}).get("minerals") or {},
        "ree":       fetch_trade_set("ree") or {},
        "strategic": fetch_trade_set("strategic") or {},
        "precious":  fetch_trade_set("precious") or {},
        "uranium":   fetch_trade_set("uranium") or {},
    }


def _v2_imports(name, imp_map, groups):
    """광종별 수입국 분포 — (dict, 출처라벨, 그룹주석). 없으면 (None, '', '')."""
    byc = _v2_lookup(imp_map, name)
    if byc:
        return byc, "관세청 수출입 통계", ""
    core = _v2_lookup(groups["core"], name)
    if isinstance(core, dict) and core.get("by_country"):
        return core["by_country"], "관세청 12개월 수입액", ""
    if name == "유연탄":
        core = (groups["core"] or {}).get("석탄") or {}
        if core.get("by_country"):
            return core["by_country"], "관세청 12개월 수입액", "석탄 기준"
    if name in _V2_REE_SET and (groups["ree"] or {}).get("by_country"):
        return groups["ree"]["by_country"], "관세청 12개월 수입액", "희토류 전체 기준"
    if name in _V2_STRAT_SET and (groups["strategic"] or {}).get("by_country"):
        return groups["strategic"]["by_country"], "관세청 12개월 수입액", "전략 희소금속 그룹 기준"
    if name in _V2_PREC_SET and (groups["precious"] or {}).get("by_country"):
        return groups["precious"]["by_country"], "관세청 12개월 수입액", "귀금속 그룹 기준"
    if name == "우라늄" and (groups["uranium"] or {}).get("by_country"):
        return groups["uranium"]["by_country"], "관세청 12개월 수입액", ""
    return None, "", ""


_V2_NEWS_Q = {"금": "금값 시세", "은": "은 시세 귀금속", "동(구리)": "구리 가격", "연(납)": "납 금속 가격",
              "철/철광석": "철광석", "주석": "주석 금속", "규소": "규소 웨이퍼", "유연탄": "유연탄 석탄",
              "백금": "백금 시세", "팔라듐": "팔라듐 가격", "아연": "아연 가격", "크롬": "크롬 금속"}


def _v2_news(name, limit=4):
    """광종별 실시간 뉴스 — 네이버 검색(30분 캐시), 실패 시 전체 광물 뉴스에서 매칭."""
    ck = "v2news_" + _v2_norm(name)
    c = cache_get(ck)
    if c is None:
        c = []
        base = re.sub(r"[\(\)/].*$", "", name)
        q = _V2_NEWS_Q.get(name) or (base + (" 광물" if len(base) <= 2 else ""))
        if NAVER_CLIENT_ID and not NAVER_CLIENT_ID.startswith("여기에"):
            try:
                r = requests.get("https://openapi.naver.com/v1/search/news.json",
                                 headers={"X-Naver-Client-Id": NAVER_CLIENT_ID,
                                          "X-Naver-Client-Secret": NAVER_CLIENT_SECRET},
                                 params={"query": q, "display": 12, "sort": "date"}, timeout=8)
                seen = set()
                from email.utils import parsedate_to_datetime as _pdt
                for it in (r.json().get("items") or []):
                    t = clean(it.get("title", ""))
                    d = clean(it.get("description", ""))[:90]
                    if any(b in (t + d) for b in NEWS_BLACKLIST):
                        continue
                    if name not in _V2_NEWS_Q and base not in (t + d):
                        continue
                    key = re.sub(r"[^0-9A-Za-z가-힣]", "", t)[:18]
                    if key in seen:
                        continue
                    seen.add(key)
                    try: dt = _pdt(it.get("pubDate", "")).strftime("%Y-%m-%d %H:%M")
                    except Exception: dt = ""
                    c.append({"제목": t, "요약": d,
                              "링크": it.get("originallink") or it.get("link", ""), "발행일": dt})
            except Exception as e:
                print("[v2news]", name, e)
        c = ai_relevance_gate(c, f"'{name}' 광물(자원·소재)의 시장·수급·산업 동향")
        cache_set(ck, c, ttl=1800)
    if not c:
        base = re.sub(r"[\(\)/].*$", "", name)
        c = [n for n in (fetch_news() or [])
             if base and base in ((n.get("제목") or "") + (n.get("요약") or ""))]
    return c[:limit]


def _v2_grade_pill(grade, score=None, prov=False):
    cls = {"위험": "p-dg", "주의": "p-wr", "안정": "p-ok"}.get(grade, "p-mu")
    txt = f"{grade} {score:.0f}" if isinstance(score, (int, float)) else (grade or "관찰")
    if prov and grade:
        txt += "*"
    return f'<span class="pill {cls}">{txt}</span>'


def _v2_rows():
    """홈 리스트용 광종 행 데이터(48종)."""
    krisk = compute_k_risk() or {}
    imp_map, _unit = by_mineral_country(fetch_customs() or [])
    groups = _v2_trade_groups()
    rows = []
    for cat, names in MINERAL_TAXONOMY.items():
        for nm in names:
            k = _v2_lookup(krisk, nm)
            byc, _src, _note = _v2_imports(nm, imp_map, groups)
            byc = byc or {}
            tot = sum(byc.values())
            top = max(byc.items(), key=lambda x: x[1]) if byc else None
            share = round(top[1] / tot * 100) if (top and tot) else None
            if _note:
                share = None  # 그룹 합산이라 광종별 %로 말하지 않음
            if top and _note:
                sub = f"{_note.replace(' 기준', '')} 수입 1위 {top[0]}"
            elif top and share is not None and share >= 50:
                sub = f"수입 {share}%가 {top[0]}에서 와요"
            elif top:
                sub = f"주요 수입국 {top[0]}"
            else:
                ug = _v2_lookup(USGS_DATA, nm) or {}
                pt = (_v2_lookup(_v2_usgs1(), nm) or {}).get("prod_top") or []
                t1 = ug.get("1위국") or (pt[0] if pt else "")
                sub = f"세계 1위 생산 {t1}" if t1 else cat
            life = [key for key, _label, mins in LIFE_CATS if nm in mins]
            rows.append({
                "name": nm, "cat": cat, "use": MIN_USES.get(nm, cat),
                "life": life, "sub": sub,
                "grade": (k or {}).get("grade"), "score": (k or {}).get("score"),
                "prov": bool((k or {}).get("잠정")),
                "top": top[0] if top else "", "share": share, "imp": tot,
            })
    grade_rank = {"위험": 0, "주의": 1, "안정": 2}
    rows.sort(key=lambda r: (grade_rank.get(r["grade"], 3),
                             -(r["score"] or 0), -(r["imp"] or 0), r["name"]))
    return rows


V2_SHELL = r"""<!DOCTYPE html>
<html lang="ko">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>__TITLE__</title>
<link rel="icon" type="image/png" href="/static/favicon.png?v=3">
<link rel="preconnect" href="https://fonts.googleapis.com">
<link href="https://fonts.googleapis.com/css2?family=Noto+Sans+KR:wght@400;500;700;900&display=swap" rel="stylesheet">
<style>
:root{
  --blue:#155BB8;--blue2:#0E4A99;--navy:#16305C;--sky:#EAF2FC;--sky2:#D6E6F9;
  --kred:#C24E59;--kredl:#FBE9EC;--kblue:#0047A0;--taeg:linear-gradient(100deg,#C24E59 0%,#C24E59 44%,#0047A0 56%,#0047A0 100%);
  --ink:#222;--ink2:#555;--ink3:#888;--line:#DDE3EA;--line2:#C9D2DD;
  --bg:#F5F7FA;--card:#fff;
  --dgr:#D0342C;--dgl:#FCEBEA;--wrn:#B8720A;--wrl:#FCF3E2;--okc:#1E7B45;--okl:#E7F4EC;
  /* 기존 페이지 CSS 호환 별칭 */
  --g:#155BB8;--gd:#16305C;--gl:#EAF2FC;--mut:#555;
}
*{margin:0;padding:0;box-sizing:border-box}
html{scroll-behavior:smooth}
body{font-family:'Noto Sans KR',-apple-system,'Malgun Gothic',sans-serif;background:#fff;color:var(--ink);
font-size:15px;line-height:1.6;-webkit-font-smoothing:antialiased}
a{color:inherit;text-decoration:none}
button{font-family:inherit;cursor:pointer}
.wrap{max-width:1200px;margin:0 auto;padding:0 24px}
/* ── 유틸바 ── */
.ubar{background:#F2F4F7;border-bottom:1px solid var(--line);font-size:12px}
.ubar .wrap{display:flex;justify-content:space-between;align-items:center;height:36px}
.ubar .u-l{color:var(--ink3);display:flex;gap:4px}
.ubar a{color:var(--ink2)} .ubar a:hover{color:var(--blue);text-decoration:underline}
.ubar .u-r a{padding:0 11px;position:relative}
.ubar .u-r a+a::before{content:'';position:absolute;left:0;top:50%;transform:translateY(-50%);width:1px;height:10px;background:var(--line2)}
.ubar .u-r span{display:none}
/* ── 헤더 + GNB (한 줄, 86px) ── */
.tbar{border-bottom:2px solid;border-image:linear-gradient(100deg,#C24E59 0%,#C24E59 34%,#0047A0 66%,#0047A0 100%) 1;background:#fff;position:sticky;top:0;z-index:60}
.tbar::before{content:'';display:block;height:4px;background:var(--taeg)}
.tbar .wrap{display:flex;align-items:center;height:86px;gap:18px}
.logo{display:flex;align-items:center;gap:11px;flex:none}
.logo .dot{width:30px;height:30px;border-radius:8px;background:linear-gradient(135deg,var(--kred),var(--kblue));flex:none;
box-shadow:inset 0 0 0 3px rgba(255,255,255,.22)}
.logo b{font-size:21px;font-weight:900;letter-spacing:-.5px;color:var(--navy);display:block;line-height:1.25}
.logo b em{font-style:normal;background:linear-gradient(135deg,#C24E59 15%,#7E4468 50%,#0047A0 85%);-webkit-background-clip:text;background-clip:text;-webkit-text-fill-color:transparent;color:var(--kblue)}
.logo i{display:block;font-style:normal;font-size:11px;color:var(--ink3);letter-spacing:.4px;font-weight:500}
.gnbbar{display:flex;flex:1;height:86px}
.gnbbar .wrap{display:contents}
.gnbbar a,.gdrop>a{display:flex;align-items:center;height:86px;padding:0 16px;font-size:15.5px;font-weight:700;color:var(--ink);position:relative;white-space:nowrap}
.gnbbar a::after,.gdrop>a::after{content:'';position:absolute;left:16px;right:16px;bottom:0;height:3px;background:linear-gradient(100deg,#C24E59,#0047A0);transform:scaleX(0);transition:transform .2s}
.gnbbar a:hover,.gdrop:hover>a{color:var(--blue)}
.gnbbar a:hover::after,.gdrop:hover>a::after,.gnbbar a.on::after{transform:scaleX(1)}
.gnbbar a.on{color:var(--blue)}
.gdrop{position:relative;display:flex}
.gmenu{display:none;position:absolute;left:50%;transform:translateX(-50%);top:86px;background:#fff;border:1px solid var(--line);
border-top:2px solid var(--blue);min-width:184px;padding:8px 0;box-shadow:0 10px 24px rgba(22,48,92,.12);z-index:70}
.gdrop:hover .gmenu{display:block}
.gmenu a{display:block;height:auto;padding:9px 20px;font-size:14px;font-weight:500;color:var(--ink2)}
.gmenu a::after{display:none}
.gmenu a:hover{background:var(--sky);color:var(--blue);font-weight:700}
.gmenu-w{display:none;min-width:0;padding:14px 6px 12px}
.gdrop:hover .gmenu-w{display:flex;gap:4px}
.gmenu-w .gcol{min-width:172px}
.gmenu-w .gt{font-size:11.5px;font-weight:800;color:var(--ink3);letter-spacing:.06em;padding:0 20px 7px}
.sbox{display:flex;align-items:center;gap:8px;background:var(--bg);border:1px solid var(--line2);border-radius:100px;padding:9px 16px;width:185px;flex:none}
.sbox svg{width:15px;height:15px;stroke:var(--ink3);flex:none}
.sbox input{border:0;outline:0;background:none;font:inherit;font-size:13px;width:100%;color:var(--ink)}
.mtab{display:none}
@media(max-width:900px){
 .gnbbar{display:none}
 .sbox{margin-left:auto;width:150px}
 .tbar .wrap{height:64px}
 .mtab{display:flex;position:fixed;left:0;right:0;bottom:0;z-index:80;background:#fff;border-top:1px solid var(--line);
  justify-content:space-around;padding:8px 0 calc(8px + env(safe-area-inset-bottom))}
 .mtab a{display:flex;flex-direction:column;align-items:center;gap:3px;font-size:11px;color:var(--ink3);font-weight:700}
 .mtab a.on{color:var(--blue)}
 .mtab svg{width:22px;height:22px;stroke:currentColor}
 body{padding-bottom:74px}
}
/* ── 공통 컴포넌트 ── */
.card{background:var(--card);border:1px solid var(--line);border-radius:10px;padding:20px 22px}
.pill{display:inline-flex;align-items:center;border-radius:100px;padding:3px 11px;font-size:12.5px;font-weight:700;white-space:nowrap}
.p-dg{background:var(--dgl);color:var(--dgr)}.p-wr{background:var(--wrl);color:var(--wrn)}
.p-ok{background:var(--okl);color:var(--okc)}.p-mu{background:#EEF1F5;color:var(--ink3)}
h1{font-size:24px;font-weight:900;letter-spacing:-.5px}
.sec-t{font-size:16px;font-weight:900;margin:0 0 10px}
.modt{display:flex;justify-content:space-between;align-items:flex-end;border-bottom:2px solid var(--navy);padding-bottom:9px;margin:30px 0 0}
.modt b{font-size:19px;font-weight:900;letter-spacing:-.3px}
.modt a{font-size:12.5px;color:var(--ink3);font-weight:700}
.modt a:hover{color:var(--blue)}
/* ── 관련 사이트 ── */
.relsites{border-top:1px solid var(--line);background:#fff;padding:15px 0;margin-top:56px}
.relsites .wrap{display:flex;gap:12px;align-items:center;flex-wrap:wrap}
.relsites b{font-size:13.5px}
.relsites select{border:1px solid var(--line2);border-radius:4px;padding:9px 12px;font:inherit;font-size:13px;min-width:230px;background:#fff;color:var(--ink)}
/* ── 푸터 (네이비) ── */
footer{margin:0;background:var(--navy);color:#B9CCEA;font-size:13px;border-top:3px solid;border-image:linear-gradient(90deg,#C24E59,#0047A0) 1}
footer .f-wrap{max-width:1200px;margin:0 auto;padding:34px 24px 26px;display:grid;grid-template-columns:1.3fr 1fr 1fr;gap:28px}
footer h4{color:#fff;font-size:13.5px;font-weight:700;margin-bottom:10px}
footer .f-brand{font-size:16px;font-weight:900;color:#fff;display:flex;align-items:center;gap:8px;margin-bottom:8px}
footer .f-brand .dot{width:9px;height:9px;border-radius:50%;background:linear-gradient(135deg,#D97680,#4A7DD6)}
footer p{line-height:1.65;font-size:12.5px}
footer ul{list-style:none} footer li{margin:5px 0;font-size:12.5px}
footer a{color:#B9CCEA} footer a:hover{color:#fff;text-decoration:underline}
footer .f-status{display:inline-flex;align-items:center;gap:6px;background:rgba(255,255,255,.08);border-radius:100px;padding:5px 13px;font-size:12px;margin-top:10px}
footer .f-status .ok{width:7px;height:7px;border-radius:50%;background:#4CD68C}
footer .f-bot{border-top:1px solid rgba(255,255,255,.14);text-align:center;padding:14px;font-size:11.5px;color:#8FA6C6}
@media(max-width:900px){.ubar{display:none}footer .f-wrap{grid-template-columns:1fr;gap:18px}}
__EXTRA_CSS__
</style>
</head>
<body>
<div class="tbar"><div class="wrap">
  <a class="logo" href="/"><svg width="36" height="36" viewBox="0 0 40 40" aria-hidden="true"><defs>
<linearGradient id="kmrg" x1="0" y1="0" x2="0.85" y2="1">
<stop offset="0" stop-color="#D66671"/><stop offset=".42" stop-color="#B84C59"/><stop offset=".58" stop-color="#2A3F8F"/><stop offset="1" stop-color="#0047A0"/></linearGradient></defs>
<polygon points="20,2 36,11 36,29 20,38 4,29 4,11" fill="url(#kmrg)"/>
<polygon points="20,2 36,11 20,20 4,11" fill="#ffffff" opacity=".14"/>
<polygon points="20,20 36,11 36,29 20,38" fill="#000000" opacity=".12"/>
<polyline points="8,22 14,22 17,13 22,29 25,20 32,20" fill="none" stroke="#ffffff" stroke-width="2.6" stroke-linecap="round" stroke-linejoin="round"/>
<circle cx="32" cy="20" r="2.6" fill="#FFD24A"/></svg><span><b><em>K</em> Mineral Risk</b><i>핵심광물 공급망 위험 진단 · KOREA CMR INTELLIGENCE</i></span></a>
  <nav class="gnbbar">
  <a href="/" class="__A_HOME__">홈</a>
  <a href="/globe" class="__A_MAP__">핵심광물지도</a>
  <div class="gdrop"><a href="#" onclick="return false">통계</a><div class="gmenu gmenu-w">
    <div class="gcol"><div class="gt">지표별</div>
      <a href="/dashboard#supply">수급 현황</a>
      <a href="/dashboard#mindex">가격지수</a>
      <a href="/dashboard#forecast">가격 전망</a>
      <a href="/dashboard#map">글로벌 매장량</a>
      <a href="/dashboard#routes">수입 루트</a>
      <a href="/dashboard#risk">리스크 신호등</a>
      <a href="/dashboard#mines">국내 광산</a>
    </div>
    <div class="gcol"><div class="gt">광종별</div>
      <a href="/dashboard#cat-minerals">핵심광물 종합</a>
      <a href="/dashboard#cat-nf">비철금속 (6종)</a>
      <a href="/dashboard#cat-rare">희소금속 (20종)</a>
      <a href="/dashboard#cat-ree">희토류 (14종)</a>
      <a href="/dashboard#cat-energy">에너지 (2종)</a>
      <a href="/dashboard#cat-etc">기타 (6종)</a>
    </div>
  </div></div>
  <a href="/briefing" class="__A_BRF__">브리핑</a>
  <a href="/conference" class="__A_AI__">AI 회의</a>
  </nav>
  <div class="sbox"><svg viewBox="0 0 24 24" fill="none" stroke-width="2"><circle cx="11" cy="11" r="7"/><path d="m20 20-3.5-3.5"/></svg><input id="gq" placeholder="광물·용도 검색" value="__Q__"></div>
</div></div>
__CONTENT__
<div class="relsites"><div class="wrap"><b>관련 사이트</b>
  <select onchange="if(this.value){window.open(this.value);this.value=''}">
    <option value="">데이터 출처 바로가기</option>
    <option value="https://www.komis.or.kr">KOMIS 한국자원정보서비스</option>
    <option value="https://www.data.go.kr">공공데이터포털</option>
    <option value="https://tradedata.go.kr">관세청 수출입무역통계</option>
    <option value="https://www.usgs.gov">USGS</option>
    <option value="https://www.worldbank.org/en/research/commodity-markets">World Bank 원자재 시장</option>
  </select>
  <select onchange="if(this.value){location=this.value;this.value=''}">
    <option value="">K Mineral Risk 서비스 바로가기</option>
    <option value="/globe">핵심광물지도</option>
    <option value="/briefing">브리핑 · 리포트 구독</option>
    <option value="/conference">AI 전문가 회의실</option>
    <option value="/minerals.csv">데이터 내려받기(CSV)</option>
  </select>
</div></div>
<footer>
<div class="f-wrap">
  <div>
    <div class="f-brand"><span class="dot"></span>K Mineral Risk</div>
    <p>흩어진 광물 공공데이터를 융합해 공급망 위험을 하나의 지수로 진단하는
    핵심광물 인텔리전스 서비스입니다. 모든 수치에 출처가 표기됩니다.</p>
    <div class="f-status"><span class="ok"></span>정상 운영 · 데이터 기준 __ASOF__</div>
  </div>
  <div>
    <h4>데이터 출처 · 갱신</h4>
    <ul>
      <li>KOMIR 수급안정화지수·파생지수 — 월간</li>
      <li>관세청 수출입 통계(OpenAPI) — 12개월 실시간</li>
      <li>조달청 LME 시세 · 산업부 철강 — 일·월간</li>
      <li>USGS MCS · World Bank CMO — 연·월간</li>
    </ul>
  </div>
  <div>
    <h4>바로가기</h4>
    <ul>
      <li><a href="/globe">핵심광물지도</a></li>
      <li><a href="/briefing">브리핑 · 리포트 구독</a></li>
      <li><a href="/conference">AI 전문가 회의실</a></li>
      <li><a href="/minerals.csv">데이터 내려받기(CSV)</a></li>
    </ul>
  </div>
</div>
<div class="f-bot">본 서비스는 산업통상자원부·산하기관 공공데이터를 활용합니다 · 팀 SMART-X, 세종대학교 에너지자원공학과 · © 2026 K Mineral Risk</div>
</footer>
<nav class="mtab">
 <a href="/" class="__A_HOME__"><svg viewBox="0 0 24 24" fill="none" stroke-width="2"><path d="M3 11 12 3l9 8v9a1 1 0 0 1-1 1h-5v-6h-6v6H4a1 1 0 0 1-1-1z"/></svg>홈</a>
 <a href="/globe" class="__A_MAP__"><svg viewBox="0 0 24 24" fill="none" stroke-width="2"><circle cx="12" cy="12" r="9"/><path d="M3 12h18M12 3c3 3.5 3 14 0 18M12 3c-3 3.5-3 14 0 18"/></svg>광물지도</a>
 <a href="/briefing" class="__A_BRF__"><svg viewBox="0 0 24 24" fill="none" stroke-width="2"><path d="M5 4h14a1 1 0 0 1 1 1v14a1 1 0 0 1-1 1H5a1 1 0 0 1-1-1V5a1 1 0 0 1 1-1z"/><path d="M8 9h8M8 13h8M8 17h5"/></svg>브리핑</a>
 <a href="/conference" class="__A_AI__"><svg viewBox="0 0 24 24" fill="none" stroke-width="2"><path d="M21 12a8 8 0 0 1-8 8H4l2-3a8 8 0 1 1 15-5z"/></svg>AI 회의</a>
</nav>
<script>
(function(){
  var q=document.getElementById('gq');
  if(q){q.addEventListener('keydown',function(e){
    if(e.key==='Enter'&&!document.getElementById('mlist')){location='/?q='+encodeURIComponent(q.value);}
  });}
})();
__JS__
</script>
</body></html>"""


def _v2_shell(active, title, content, extra_css="", js="", q=""):
    html = (V2_SHELL
            .replace("__ASOF__", datetime.now().strftime("%Y-%m-%d %H:%M"))
            .replace("__TITLE__", title)
            .replace("__EXTRA_CSS__", extra_css)
            .replace("__CONTENT__", content)
            .replace("__JS__", js)
            .replace("__Q__", q.replace('"', "&quot;")))
    for key, mark in (("home", "__A_HOME__"), ("map", "__A_MAP__"),
                      ("brf", "__A_BRF__"), ("ai", "__A_AI__")):
        html = html.replace(mark, "on" if active == key else "")
    return html


V2_HOME_CSS = r"""
.sbox{display:none}
.hero{position:relative;background:linear-gradient(150deg,#96323F 0%,#6E3160 24%,#2E2F72 48%,#123C7E 66%,#1E74D8 100%);color:#fff;overflow:hidden}
.hero::before{content:'';position:absolute;right:-120px;top:-140px;width:520px;height:520px;border-radius:50%;border:70px solid rgba(255,255,255,.05)}
.hero::after{content:'';position:absolute;right:120px;bottom:-200px;width:420px;height:420px;border-radius:50%;border:54px solid rgba(255,255,255,.05)}
.hero .wrap{position:relative;z-index:2;padding-top:60px;padding-bottom:92px}
.h-badge{display:inline-block;font-size:12.5px;font-weight:700;background:rgba(255,255,255,.16);border:1px solid rgba(255,255,255,.3);padding:5px 14px;border-radius:100px;margin-bottom:16px}
.h-tt{font-size:38px;font-weight:900;letter-spacing:-1px;line-height:1.3}
.h-sub{font-size:15.5px;color:#CBDDF6;margin-top:12px}
.h-cta{display:inline-block;margin-top:22px;background:#fff;color:var(--blue2);font-size:14px;font-weight:700;padding:11px 26px;border-radius:4px}
.h-cta:hover{background:var(--sky)}
.gsearch{position:relative;z-index:5;margin-top:-34px}
.h-search{display:flex;align-items:center;background:#fff;border:2px solid var(--blue);border-radius:100px;padding:5px 5px 5px 26px;max-width:760px;margin:0 auto;box-shadow:0 12px 30px rgba(22,48,92,.16)}
.h-search input{flex:1;border:0;outline:0;font:inherit;font-size:15.5px;color:var(--ink);background:none}
.h-search button{border:0;background:var(--blue);color:#fff;border-radius:100px;width:48px;height:48px;font-size:16px;cursor:pointer;flex:none}
.h-search button:hover{background:var(--blue2)}
.h-pop{margin-top:14px;font-size:12.5px;color:var(--ink3);font-weight:700;display:flex;gap:8px;align-items:center;flex-wrap:wrap;justify-content:center}
.h-pop button{border:1px solid var(--line);background:var(--bg);color:var(--ink2);border-radius:100px;padding:4px 13px;font:500 12.5px/1.5 inherit;font-family:inherit;cursor:pointer}
.h-pop button:hover{border-color:var(--blue);color:var(--blue)}
.hgrid{display:grid;grid-template-columns:2fr 1fr 1fr;gap:12px;margin-bottom:6px}
.alert{border-radius:8px;padding:18px 20px}
.alert.a-dg{background:var(--dgl)} .alert.a-wr{background:var(--wrl)} .alert.a-ok{background:var(--okl)}
.alert .a-t{font-size:17px;font-weight:800;display:flex;align-items:center;gap:8px}
.alert.a-dg .a-t,.alert.a-dg .a-b,.alert.a-dg a{color:var(--dgr)}
.alert.a-wr .a-t,.alert.a-wr .a-b,.alert.a-wr a{color:var(--wrn)}
.alert.a-ok .a-t,.alert.a-ok .a-b,.alert.a-ok a{color:var(--gd)}
.alert .a-b{font-size:13.5px;margin-top:5px;line-height:1.55}
.alert a{font-size:13px;font-weight:750;display:inline-block;margin-top:9px}
.stat{background:var(--card);border:1px solid var(--line);border-radius:8px;padding:16px 18px}
.stat .s-l{font-size:12.5px;color:var(--mut)}
.stat .s-v{font-size:26px;font-weight:800;margin-top:2px}
.chips{display:flex;gap:7px;flex-wrap:wrap;margin:2px 0 16px}
.chip{padding:7px 15px;border-radius:999px;border:1px solid var(--line);background:var(--card);
font-size:13.5px;font-weight:650;color:var(--mut);cursor:pointer;transition:.15s}
.chip.on{background:var(--gd);border-color:var(--gd);color:#fff}
.hgrid2{display:grid;grid-template-columns:1.65fr 1fr;gap:16px;align-items:start}
.mlist{background:var(--card);border:1px solid var(--line);border-radius:8px;overflow:hidden}
.mrow{display:flex;align-items:center;gap:13px;padding:13px 18px;border-bottom:1px solid var(--line);transition:.12s}
.mrow:last-child{border-bottom:0}
.mrow:hover{background:var(--bg)}
.sig{width:10px;height:10px;border-radius:50%;flex:none}
.sig.dg{background:var(--dgr)}.sig.wr{background:#E8A13C}.sig.ok{background:#1E7B45}.sig.mu{background:#C9D2DD}
.m-main{flex:1;min-width:0}
.m-main b{font-size:15px;font-weight:750}
.m-main i{font-style:normal;font-size:12.5px;color:var(--mut);margin-left:7px}
.m-main em{display:block;font-style:normal;font-size:12.5px;color:var(--mut);margin-top:1px}
.more-btn{display:block;width:100%;padding:13px;background:var(--card);border:0;border-top:1px solid var(--line);
font:inherit;font-size:13.5px;font-weight:700;color:var(--gd);cursor:pointer}
.rail{display:flex;flex-direction:column;gap:12px;position:sticky;top:78px}
.rail .r-l{font-size:12.5px;color:var(--mut);font-weight:700;margin-bottom:7px;display:flex;align-items:center;gap:6px}
.rail .r-b{font-size:13.5px;line-height:1.6}
.rail .ai-card{background:var(--gd);border:0;color:#fff}
.rail .ai-card .r-t{font-size:14.5px;font-weight:750}
.rail .ai-card .r-q{font-size:12.5px;opacity:.85;margin-top:4px}
.rail .ai-card a{display:inline-block;margin-top:10px;background:rgba(255,255,255,.16);border-radius:999px;
padding:6px 14px;font-size:12.5px;font-weight:700;color:#fff}
.qmenu{display:grid;grid-template-columns:repeat(7,1fr);gap:6px;padding:34px 0 4px;margin:0 0 8px}
.qmenu a{display:flex;flex-direction:column;align-items:center;gap:9px;background:none;
border:0;padding:10px 4px;font-size:13px;font-weight:700;color:var(--ink2);transition:.12s}
.qmenu a:hover{color:var(--blue)}
.qmenu .qi{width:56px;height:56px;border-radius:50%;background:var(--sky);display:flex;align-items:center;justify-content:center;font-size:23px;transition:.12s}
.qmenu a:hover .qi{background:var(--sky2)}
.popq{font-size:12.5px;color:var(--mut);font-weight:700;margin:0 0 10px;display:flex;align-items:center;gap:6px;flex-wrap:wrap}
.popq button{border:1px solid var(--line);background:var(--card);color:var(--gd);border-radius:999px;
padding:3px 11px;font:650 12px/1.5 inherit;font-family:inherit;cursor:pointer}
.popq button:hover{border-color:var(--g)}
.notice{list-style:none}
.notice li{border-bottom:1px solid var(--line);padding:8px 0}
.notice li:last-child{border-bottom:0}
.notice a{display:block;font-size:12.5px;line-height:1.5;color:var(--ink)}
.notice span{display:block;font-size:11px;color:var(--mut);margin-top:1px}
@media(max-width:860px){.qmenu{display:flex;flex-wrap:wrap;justify-content:center;gap:18px 26px}}
.mod3{display:grid;grid-template-columns:1.25fr 1fr 1fr;gap:14px;margin:20px 0 0}
.mh{display:flex;justify-content:space-between;align-items:center;margin-bottom:10px}
.mh b{font-size:15px;font-weight:800}
.mh a{font-size:12px;color:var(--mut);font-weight:650}
.srcline{font-size:11px;color:var(--ink3);margin-top:12px;padding-top:10px;border-top:1px dashed var(--line)}
.ptable{width:100%;border-collapse:collapse;font-size:13px}
.ptable th{font-size:11.5px;color:var(--mut);font-weight:700;border-bottom:1px solid var(--line);padding:4px 2px;text-align:left}
.ptable td{padding:7px 2px;border-bottom:1px solid var(--line)}
.ptable tr:last-child td{border-bottom:0}
.ptable .num{text-align:right;font-variant-numeric:tabular-nums}
.krow{display:flex;align-items:center;gap:8px;padding:7.5px 0;border-bottom:1px solid var(--line);font-size:13.5px}
.krow:last-of-type{border-bottom:0}
.krow b{font-weight:750}
.krow i{font-style:normal;font-size:11.5px;color:var(--mut);flex:1}
.nrow{display:flex;gap:8px;align-items:baseline;padding:7px 0;border-bottom:1px solid var(--line);font-size:13px;line-height:1.45}
.nrow:last-of-type{border-bottom:0}
.nrow span{flex:1;overflow:hidden;text-overflow:ellipsis;display:-webkit-box;-webkit-line-clamp:2;-webkit-box-orient:vertical}
.nrow em{font-style:normal;font-size:11px;color:var(--mut);flex:none}
.chiprow{display:flex;gap:10px;align-items:flex-start;margin-top:12px}
.chiprow>span{font-size:12px;font-weight:800;color:var(--mut);padding:8px 0 0;flex:none;width:44px}
.chiprow .chips{margin:0}
.bnr4{display:grid;grid-template-columns:repeat(4,1fr);gap:12px;margin-top:26px}
.bnr4 a{background:var(--card);border:1px solid var(--line);border-radius:8px;padding:16px 18px;transition:.12s}
.bnr4 a:hover{border-color:var(--g)}
.bnr4 b{display:block;font-size:14px;font-weight:800}
.bnr4 span{display:block;font-size:12px;color:var(--mut);margin-top:3px}
@media(max-width:860px){.hgrid{grid-template-columns:1fr 1fr}.hgrid .alert{grid-column:1/3}
.hgrid2{grid-template-columns:1fr}.rail{position:static}.mod3{grid-template-columns:1fr}
.bnr4{grid-template-columns:1fr 1fr}.chiprow>span{display:none}}
"""


def render_home_v2():
    rows = _v2_rows()
    q0 = (request.args.get("q") or "").strip()
    now = datetime.now()
    wd = "월화수목금토일"[now.weekday()]
    n_dg = sum(1 for r in rows if r["grade"] == "위험")
    n_wr = sum(1 for r in rows if r["grade"] == "주의")

    top = next((r for r in rows if r["grade"]), None)
    if top and top["grade"] == "위험":
        acls, icon = "a-dg", "⚠️"
        at = f"{top['name']}이 위험 단계예요"
    elif top and top["grade"] == "주의":
        acls, icon = "a-wr", "👀"
        at = f"{top['name']}을 지켜봐야 해요"
    else:
        acls, icon = "a-ok", "☀️"
        at = "오늘 광물 시장은 대체로 맑아요"
    ab = ""
    if top:
        if top["share"] and top["share"] >= 50:
            ab = f"수입의 {top['share']}%를 {top['top']} 한 나라에 의존하고 있어요. "
        ab += f"{top['use']} 가격에 영향을 줄 수 있어요."
        alink = f'<a href="/m/{_v2q(top["name"], safe="")}">왜 그런가요 →</a>'
    else:
        ab = "48개 광물의 수급·가격·수입선을 매일 살펴보고 있어요."
        alink = ""

    chips = ['<button class="chip on" data-c="all">전체</button>'] + [
        f'<button class="chip" data-c="{key}">{label}</button>'
        for key, label, _m in LIFE_CATS
    ]
    pchips = "".join(
        f'<button class="chip pchip" data-c="p:{c}">{c}</button>'
        for c in ["비철금속", "희소금속", "희토류", "에너지", "기타"])

    _rk = load_risk_data() or []
    _colors = ["#155BB8", "#c8931d", "#1c5cab", "#c0392b", "#7a5195", "#2e8b8b"]
    risk_js = json.dumps({
        "labels": (_rk[0].get("months") if _rk else []) or [],
        "series": [{"name": r["name"], "vals": r.get("vals") or [], "color": _colors[i % 6]}
                   for i, r in enumerate(_rk)],
    }, ensure_ascii=False)

    _ppa = load_json(os.path.join(os.path.dirname(__file__), "ppa_data1.json")) or {}
    _lme_rows = ""
    for it in (_ppa.get("items") or [])[:6]:
        chg = it.get("chg") or 0
        cc = "#c0392b" if chg > 0 else ("#1c5cab" if chg < 0 else "#888888")
        arrow = "▲" if chg > 0 else ("▼" if chg < 0 else "—")
        _lme_rows += (f'<tr><td>{it.get("name", "")}</td>'
                      f'<td class="num">{(it.get("close") or 0):,.0f}</td>'
                      f'<td class="num" style="color:{cc}">{arrow} {abs(chg):.2f}%</td></tr>')
    lme_html = (f'<table class="ptable"><thead><tr><th>품목</th><th class="num">종가($/t)</th>'
                f'<th class="num">등락</th></tr></thead><tbody>{_lme_rows}</tbody></table>'
                f'<div class="srcline">조달청 비축물자·LME · 기준 {_ppa.get("date", "—")}'
                + (f' · LME지수 {_ppa["lme"]["idx"]:,.0f}' if _ppa.get("lme") else "") + "</div>")                if _lme_rows else '<div class="srcline">시세 데이터 준비 중</div>'

    _top6 = [r for r in rows if r["grade"]][:6]
    ktop_html = "".join(
        f'<a class="krow" href="/m/{_v2q(r["name"], safe="")}"><b>{r["name"]}</b>'
        f'<i>{r["cat"]}</i>{_v2_grade_pill(r["grade"], r["score"], r.get("prov"))}</a>'
        for r in _top6) or '<div class="srcline">계산 중</div>'

    _nws = [n for n in dedup_news(fetch_news() or []) if mineral_relevant(n)][:5]
    news_mod = "".join(
        f'<a class="nrow" href="{n.get("링크", "#")}" target="_blank" rel="noopener">'
        f'<span>{n.get("제목", "")}</span><em>{(n.get("발행일") or "")[5:10]}</em></a>'
        for n in _nws) or '<div class="srcline">뉴스 준비 중</div>' 

    row_html = []
    for r in rows:
        sig = {"위험": "dg", "주의": "wr", "안정": "ok"}.get(r["grade"], "mu")
        side = _v2_grade_pill(r["grade"], r["score"], r.get("prov")) if r["grade"] else '<span class="pill p-mu">관찰</span>'
        dataq = f'{r["name"]} {r["use"]} {r["cat"]} {r["sub"]}'
        row_html.append(
            f'<a class="mrow" href="/m/{_v2q(r["name"], safe="")}" data-cats="{" ".join(r["life"])} p:{r["cat"]}" '
            f'data-q="{dataq}">'
            f'<span class="sig {sig}"></span>'
            f'<span class="m-main"><b>{r["name"]}</b><i>{r["use"]}</i><em>{r["sub"]}</em></span>'
            f'{side}</a>'
        )

    content = f"""
<div class="hero">
  <div class="wrap">
    <span class="h-badge">핵심광물 48종 · 5개 분류 실시간 관측</span>
    <div class="h-tt">흩어진 광물 데이터를 융합해<br>공급망 위험을 하나의 지수로 진단합니다</div>
    <div class="h-sub">{now.month}월 {now.day}일 {wd}요일 · KOMIR·관세청·조달청·USGS 공개 지표 교차 계산 — 오늘의 광물 날씨</div>
    <a class="h-cta" href="#mlist">전체 광종 현황 바로가기</a>
  </div>
</div>
<div class="wrap gsearch">
  <div class="h-search"><input id="heroq" placeholder="광종명·용도를 입력하세요 (예: 리튬, 배터리, 자석)"><button aria-label="검색">🔍</button></div>
  <div class="h-pop">인기 검색어
    <button data-q="리튬">#리튬</button><button data-q="텅스텐">#텅스텐</button>
    <button data-q="흑연">#흑연</button><button data-q="네오디뮴">#네오디뮴</button>
    <button data-q="갈륨">#갈륨</button><button data-q="배터리">#배터리</button></div>
</div>
<div class="wrap">
  <div class="qmenu">
    <a href="#mlist"><span class="qi">◉</span>전체 광종</a>
    <a href="/globe"><span class="qi">🌍</span>광물 지도</a>
    <a href="/briefing"><span class="qi">📰</span>브리핑</a>
    <a href="/conference"><span class="qi">🎙</span>AI 회의</a>
    <a href="/briefing#sub"><span class="qi">✉️</span>리포트 구독</a>
    <a href="/minerals.csv"><span class="qi">⬇</span>데이터 받기</a>
    <a href="/dashboard"><span class="qi">📊</span>통계 대시보드</a>
  </div>
  <div class="hgrid">
    <div class="alert {acls}">
      <div class="a-t">{icon} {at}</div>
      <div class="a-b">{ab}</div>
      {alink}
    </div>
    <div class="stat"><div class="s-l">위험 광물</div>
      <div class="s-v" style="color:var(--dgr)">{n_dg}<span style="font-size:13px;color:var(--mut);font-weight:600"> / 48</span></div></div>
    <div class="stat"><div class="s-l">주의 광물</div>
      <div class="s-v" style="color:var(--wrn)">{n_wr}<span style="font-size:13px;color:var(--mut);font-weight:600"> / 48</span></div></div>
  </div>

  <div class="mod3">
    <div class="card"><div class="mh"><b>수급안정화지수</b><a href="/dashboard#ssi">더보기 +</a></div>
      <div style="height:215px"><canvas id="cRisk"></canvas></div>
      <div class="srcline">출처: KOMIR 수급안정화지수(핵심 6광종·월간)</div></div>
    <div class="card"><div class="mh"><b>오늘의 금속 시세</b><a href="/dashboard#supply">더보기 +</a></div>
      {lme_html}</div>
    <div class="card"><div class="mh"><b>K-RISK 위험 상위</b><a href="/dashboard#risk">더보기 +</a></div>
      {ktop_html}
      <div class="srcline">공급망 위험지수 높은 순 · 1시간 자동 갱신</div></div>
  </div>

  <div class="modt"><b>광물 종합 현황</b><a href="/minerals.csv">CSV 내려받기 ↓</a></div>
  <div class="chiprow"><span>용도별</span><div class="chips">{''.join(chips)}</div></div>
  <div class="chiprow"><span>분류별</span><div class="chips">{pchips}</div></div>
  <div class="srcline" style="border:0;margin-top:6px;padding-top:0">* 표시는 수급안정화지수 미제공 광종 — 수입집중도·지정학·변동성 3축 잠정 점수</div>

  <div class="hgrid2">
    <div>
      <div class="mlist" id="mlist">{''.join(row_html)}</div>
      <button class="more-btn" id="moreBtn" style="border-radius:0 0 8px 8px;margin-top:-1px">광물 더 보기</button>
    </div>
    <div class="rail">
      <div class="card"><div class="r-l">📰 오늘의 브리핑</div><div class="r-b" id="railBrief">불러오는 중…</div></div>
      <div class="card"><div class="r-l">🌍 지금 세계에선</div><div class="r-b" id="railGeo">불러오는 중…</div></div>
      <div class="card ai-card"><div class="r-t">AI 전문가에게 물어보기</div>
        <div class="r-q">"흑연이 위험하면 나한테 뭐가 문제야?"</div>
        <a href="/conference">회의실 입장 →</a></div>
      <div class="card"><div class="mh" style="margin-bottom:4px"><b>자원 뉴스</b><a href="/briefing">더보기 +</a></div>
        {news_mod}</div>
      <div class="card"><div class="r-l">📌 새 소식</div>
        <ul class="notice">
          <li><a href="/briefing">관심 광물 맞춤 리포트 오픈 — 매일 09:00 발송<span>07-17</span></a></li>
          <li><a href="/globe">핵심광물지도 개편 — 수입 루트·해협 통과율<span>07-17</span></a></li>
          <li><a href="/">'오늘의 광물 날씨' 새 홈 오픈<span>07-16</span></a></li>
        </ul></div>
    </div>
  </div>

  <div class="bnr4">
    <a href="/globe"><b>🌍 핵심광물지도</b><span>수입 루트·해협 통과율</span></a>
    <a href="/conference"><b>🎙 AI 전문가 회의실</b><span>발언마다 출처 첨부</span></a>
    <a href="/briefing#sub"><b>✉️ 데일리 리포트 구독</b><span>매일 09:00 이메일 발송</span></a>
    <a href="/minerals.csv"><b>⬇ 데이터 개방(CSV)</b><span>48광종 현황 내려받기</span></a>
  </div>
</div>
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.1/dist/chart.umd.min.js"></script>"""

    js = r"""
(function(){
  var LIMIT=12, rows=[].slice.call(document.querySelectorAll('.mrow')),
      chips=[].slice.call(document.querySelectorAll('.chip')),
      more=document.getElementById('moreBtn'), cat='all', expanded=false,
      q=document.getElementById('heroq')||document.getElementById('gq');
  function apply(){
    var kw=((q&&q.value)||'').trim().toLowerCase(), shown=0;
    var limitOn=(cat==='all'&&!kw&&!expanded);
    rows.forEach(function(r){
      var okC=(cat==='all')||((r.dataset.cats||'').split(' ').indexOf(cat)>=0);
      var okQ=!kw||(r.dataset.q||'').toLowerCase().indexOf(kw)>=0;
      var ok=okC&&okQ;
      if(ok) shown++;
      r.style.display=(ok&&(!limitOn||shown<=LIMIT))?'':'none';
    });
    more.style.display=limitOn?'':'none';
  }
  chips.forEach(function(c){c.addEventListener('click',function(){
    chips.forEach(function(x){x.classList.remove('on')}); c.classList.add('on');
    cat=c.dataset.c; apply();
  });});
  more.addEventListener('click',function(){expanded=true;apply();});
  [].slice.call(document.querySelectorAll('.h-pop button')).forEach(function(b){
    b.addEventListener('click',function(){ if(q){ q.value=b.dataset.q; apply(); q.focus(); } });
  });
  if(q){q.addEventListener('input',apply); if(q.value) apply(); else apply();}
  fetch('/api/news-brief?cat=minerals').then(function(r){return r.json()}).then(function(d){
    document.getElementById('railBrief').textContent=(d&&d.brief)?d.brief:'오늘은 새 브리핑이 없어요.';
  }).catch(function(){document.getElementById('railBrief').textContent='브리핑을 불러오지 못했어요.';});
  function applyPcatHash(){
    var ph=(location.hash||'').replace('#','');
    if(ph.indexOf('pcat-')!==0) return;
    var pc=decodeURIComponent(ph.slice(5));
    var pb=document.querySelector('.chip[data-c="p:'+pc+'"]');
    if(pb){ pb.click(); var ml=document.getElementById('mlist'); if(ml) ml.scrollIntoView({behavior:'smooth'}); }
  }
  applyPcatHash();
  window.addEventListener('hashchange', applyPcatHash);
  function drawRisk(){
    if(!window.Chart) return setTimeout(drawRisk, 150);
    var R=__RISKJS__;
    if(!R.labels.length) return;
    new Chart(document.getElementById('cRisk'),{type:'line',
      data:{labels:R.labels,datasets:R.series.map(function(sr){
        return {label:sr.name,data:sr.vals,borderColor:sr.color,borderWidth:1.8,pointRadius:0,tension:.3};})},
      options:{maintainAspectRatio:false,interaction:{mode:'index',intersect:false},
        plugins:{legend:{position:'bottom',labels:{boxWidth:8,boxHeight:8,font:{size:10.5},padding:8}}},
        scales:{x:{grid:{display:false},ticks:{maxTicksLimit:6,font:{size:10}}},
                y:{grid:{color:'#EEF1F5'},ticks:{font:{size:10}}}}}});
  }
  drawRisk();
  fetch('/api/geo-events').then(function(r){return r.json()}).then(function(d){
    var ev=(d&&d.events&&d.events[0])||null;
    document.getElementById('railGeo').textContent=ev?((ev.loc?ev.loc+' — ':'')+(ev.why||ev['제목']||'')):'특별한 이슈가 없어요.';
  }).catch(function(){document.getElementById('railGeo').textContent='이슈를 불러오지 못했어요.';});
})();
"""
    return _v2_shell("home", "K Mineral Risk — 오늘의 광물 날씨",
                     content, V2_HOME_CSS, js.replace("__RISKJS__", risk_js), q=q0)


V2_DETAIL_CSS = r"""
.bk{display:inline-flex;align-items:center;gap:6px;margin:22px 0 14px;font-size:13.5px;font-weight:700;color:var(--mut)}
.dhead{display:flex;align-items:flex-start;gap:14px;flex-wrap:wrap}
.dhead h1{font-size:28px}
.dhead .u{font-size:14px;color:var(--mut);margin-top:3px}
.deasy{font-size:15px;line-height:1.65;margin:14px 0 20px;background:var(--card);border:1px solid var(--line);
border-radius:18px;padding:16px 20px}
.dgrid{display:grid;grid-template-columns:1fr 1fr;gap:14px;margin-bottom:14px}
.dgrid .card h3{font-size:14.5px;font-weight:800;margin-bottom:12px}
.facts{display:flex;gap:10px;flex-wrap:wrap;margin:0 0 18px}
.fact{background:var(--card);border:1px solid var(--line);border-radius:14px;padding:10px 16px;font-size:13px}
.fact b{display:block;font-size:14.5px;margin-top:1px}
.newsl a{display:block;padding:10px 0;border-bottom:1px solid var(--line);font-size:13.5px;line-height:1.5}
.newsl a:last-child{border-bottom:0}
.newsl .nd{font-size:11.5px;color:var(--mut);margin-top:2px}
details.card{margin-top:14px}
details.card summary{cursor:pointer;font-weight:800;font-size:14.5px;color:var(--mut)}
.kbar{margin:12px 0 4px}
.kbar .kl{display:flex;justify-content:space-between;font-size:12.5px;color:var(--mut);margin-bottom:3px}
.kbar .kt{height:8px;border-radius:999px;background:var(--bg);overflow:hidden}
.kbar .kf{height:100%;border-radius:999px;background:var(--gd)}
@media(max-width:860px){.dgrid{grid-template-columns:1fr}}
"""


def render_mineral_v2(name):
    rows = {r["name"]: r for r in _v2_rows()}
    r = rows.get(name) or _v2_lookup(rows, name)
    if not r:
        return _v2_shell("home", "K Mineral Risk", f'<div class="wrap"><div class="bk">← <a href="/">홈으로</a></div>'
                                          f'<h1>{name}</h1><p style="margin-top:10px">아직 준비되지 않은 광물이에요.</p></div>')
    name = r["name"]
    krisk = _v2_lookup(compute_k_risk() or {}, name) or {}
    imp_map, _unit0 = by_mineral_country(fetch_customs() or [])
    byc, imp_src, imp_note = _v2_imports(name, imp_map, _v2_trade_groups())
    byc = byc or {}
    ug = _v2_lookup(USGS_DATA, name) or {}

    # 쉬운 말 요약
    easy = f"<b>{name}</b>은(는) {r['use']}에 쓰이는 광물이에요. "
    if r["share"] and r["share"] >= 50:
        easy += f"우리나라는 수입의 <b>{r['share']}%</b>를 <b>{r['top']}</b>에서 들여오고 있어서, " \
                f"그 나라 상황에 따라 값이 출렁일 수 있어요. "
    elif r["top"]:
        easy += f"주로 <b>{r['top']}</b>에서 수입해요. "
    if krisk:
        g = krisk.get("grade")
        if g == "위험":
            easy += "지금은 공급이 불안해질 수 있는 <b>위험 단계</b>예요."
        elif g == "주의":
            easy += "지금은 흐름을 지켜봐야 하는 <b>주의 단계</b>예요."
        else:
            easy += "지금은 수급이 안정적인 편이에요."

    facts = []
    if r["top"]:
        top_txt = r["top"] + (f' ({r["share"]}%)' if r["share"] else "")
        facts.append(f'<div class="fact">주요 수입국<b>{top_txt}</b></div>')
    if ug.get("1위국"):
        facts.append(f'<div class="fact">세계 1위 생산<b>{ug["1위국"]}</b></div>')
    facts.append(f'<div class="fact">전문 분류<b>{r["cat"]}</b></div>')

    # 수입국 차트 데이터
    pairs = sorted(byc.items(), key=lambda x: x[1], reverse=True)[:6]
    imp_labels = json.dumps([p[0] for p in pairs], ensure_ascii=False)
    imp_vals = json.dumps([round(p[1]) for p in pairs])
    if pairs:
        _lab = imp_src + (f" · {imp_note}" if imp_note else "")
        imp_card = (f'<div class="card"><h3>어디서 수입하나요? '
                    f'<span style="font-weight:600;color:var(--mut);font-size:12px">({_lab})</span></h3>'
                    f'<div style="height:230px"><canvas id="cImp"></canvas></div></div>')
    else:
        pt = (_v2_lookup(_v2_usgs1(), name) or {}).get("prod_top") or []
        _info = (f"세계 생산 1위는 <b>{pt[0]}</b>이에요 (점유 {pt[1]}%)."
                 if len(pt) == 2 and pt[0] else "수입·생산 통계가 아직 연결되지 않은 광물이에요.")
        imp_card = ('<div class="card"><h3>공급은 어디서 오나요?</h3>'
                    f'<div style="padding:64px 12px;text-align:center;color:var(--mut);font-size:14px">{_info}'
                    '<div style="font-size:11.5px;margin-top:6px">출처: USGS MCS 2026</div></div></div>')

    # 가격지수(그룹) 라인
    grp = K_MIDX_GROUP.get(name.replace("(구리)", "").replace("/철광석", ""), None) or \
          ("메이저금속" if r["cat"] == "비철금속" else ("에너지광물" if r["cat"] == "에너지" else "희소금속" if r["cat"] == "희소금속" else "종합"))
    midx = load_json(os.path.join(os.path.dirname(__file__), "mineral_index_data2.json")) or {}
    s = ((midx.get("series") or {}).get(grp)) or ((midx.get("series") or {}).get("종합")) or {}
    months = (s.get("months") or [])[-36:]
    vals = (s.get("values") or [])[-36:]
    px_labels = json.dumps(months, ensure_ascii=False)
    px_vals = json.dumps(vals)

    # 관련 뉴스 — 광종별 실시간 검색
    news = _v2_news(name)
    news_html = "".join(
        f'<a href="{n.get("링크", "#")}" target="_blank" rel="noopener">{n.get("제목", "")}'
        f'<div class="nd">{n.get("발행일", "")}</div></a>' for n in news
    ) or '<div style="color:var(--mut);font-size:13.5px">최근 관련 뉴스가 없어요.</div>'

    # 전문가용 K-RISK 4요소
    kbars = ""
    if krisk.get("요소"):
        for kk, vv in krisk["요소"].items():
            kbars += (f'<div class="kbar"><div class="kl"><span>{kk}</span><span>{vv:.0f}</span></div>'
                      f'<div class="kt"><div class="kf" style="width:{min(100, vv):.0f}%"></div></div></div>')
    expert = f"""
<details class="card"><summary>전문가용 데이터 보기</summary>
  <div style="margin-top:12px;font-size:13px;color:var(--mut)">K-RISK 구성요소 (0~100, 높을수록 위험)</div>
  {kbars or '<div style="margin-top:8px;font-size:13.5px;color:var(--mut)">K-RISK 계산 데이터가 없어요.</div>'}
  {'<div style="margin-top:8px;font-size:12px;color:var(--mut)">※ 수급안정화지수 미제공 광종 — 수입집중도·지정학·변동성 3축 잠정 점수입니다.</div>' if krisk.get("잠정") else ''}
  <div style="margin-top:14px;font-size:13.5px">
    <a href="/dashboard" style="color:var(--gd);font-weight:700">카테고리 대시보드 →</a></div>
</details>"""

    side = _v2_grade_pill(r["grade"], r["score"], r.get("prov")) if r["grade"] else '<span class="pill p-mu">관찰 대상</span>'
    content = f"""
<div class="wrap">
  <div class="bk"><a href="/">← 홈으로</a></div>
  <div class="dhead"><div><h1>{name}</h1><div class="u">{r['use']} · {r['cat']}</div></div>{side}</div>
  <div class="deasy">{easy}</div>
  <div class="facts">{''.join(facts)}</div>
  <div class="dgrid">
    {imp_card}
    <div class="card"><h3>가격 흐름 <span style="font-weight:600;color:var(--mut);font-size:12px">({grp} 지수 · 최근 3년)</span></h3>
      <div style="height:230px"><canvas id="cPx"></canvas></div></div>
  </div>
  <div class="card newsl"><h3 style="font-size:14.5px;font-weight:800;margin-bottom:6px">관련 소식</h3>{news_html}</div>
  {expert}
</div>
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.1/dist/chart.umd.min.js"></script>"""

    js = f"""
(function(){{
  if(!window.Chart) return;
  Chart.defaults.font.family="'Pretendard Variable',Pretendard,sans-serif";
  Chart.defaults.color='#555555';
  var IL={imp_labels}, IV={imp_vals}, elImp=document.getElementById('cImp');
  if(elImp&&IL.length){{
    new Chart(elImp,{{type:'bar',
      data:{{labels:IL,datasets:[{{data:IV,backgroundColor:'#155BB8',borderRadius:8,barThickness:18}}]}},
      options:{{indexAxis:'y',maintainAspectRatio:false,plugins:{{legend:{{display:false}}}},
        scales:{{x:{{grid:{{color:'#EEF1F5'}},ticks:{{font:{{size:11}}}}}},y:{{grid:{{display:false}},ticks:{{font:{{size:12}}}}}}}}}}}});
  }}
  var PL={px_labels}, PV={px_vals};
  if(PL.length){{
    new Chart(document.getElementById('cPx'),{{type:'line',
      data:{{labels:PL,datasets:[{{data:PV,borderColor:'#155BB8',borderWidth:2.5,pointRadius:0,tension:.3,fill:true,
        backgroundColor:'rgba(21,91,184,.07)'}}]}},
      options:{{maintainAspectRatio:false,plugins:{{legend:{{display:false}}}},
        scales:{{x:{{grid:{{display:false}},ticks:{{maxTicksLimit:6,font:{{size:11}}}}}},
                y:{{grid:{{color:'#EEF1F5'}},ticks:{{font:{{size:11}}}}}}}}}}}});
  }}
}})();"""
    return _v2_shell("home", f"{name} — K Mineral Risk", content, V2_DETAIL_CSS, js)


V2_BRF_CSS = r"""
.bh{margin-top:26px}
.brf-ai{background:var(--gl);border-radius:18px;padding:18px 20px;margin:14px 0 18px;font-size:14.5px;line-height:1.65;color:var(--gd)}
.brf-ai .bl{font-size:12.5px;font-weight:800;margin-bottom:6px}
.bgrid{display:grid;grid-template-columns:1.65fr 1fr;gap:16px;align-items:start}
.audtabs{display:flex;gap:7px;margin-bottom:12px}
.audtab{border:1.5px solid var(--line2);background:var(--card);color:var(--ink2);border-radius:100px;
padding:7px 20px;font:700 13.5px/1.5 inherit;font-family:inherit;cursor:pointer}
.audtab:hover{border-color:var(--blue);color:var(--blue)}
.audtab.on{background:var(--navy);border-color:var(--navy);color:#fff}
.ncard a{display:block;padding:13px 18px;border-bottom:1px solid var(--line)}
.ncard a:last-child{border-bottom:0}
.ncard .nt{font-size:14.5px;font-weight:700;line-height:1.45}
.ncard .ns{font-size:13px;color:var(--mut);margin-top:3px;line-height:1.5}
.ncard .nd{font-size:11.5px;color:var(--mut);margin-top:4px}
.sub-card{scroll-margin-top:120px}
.sub-card input{width:100%;border:1px solid var(--line);border-radius:12px;padding:11px 14px;font:inherit;font-size:14px;margin:10px 0 8px;outline-color:var(--g)}
.subchips{display:flex;flex-wrap:wrap;gap:5px;margin-bottom:10px}
.sub-card .subchips .subchip{width:auto;border:1px solid var(--line);background:var(--card);color:var(--mut);
border-radius:999px;padding:4px 11px;font-size:12px;font-weight:650;font-family:inherit;cursor:pointer}
.sub-card .subchips .subchip.on{background:var(--gd);border-color:var(--gd);color:#fff}
.sub-card .subchips .subchip.ext{display:none}
.sub-card .subchips.open .subchip.ext{display:inline-block}
.sub-card .subchips .subchip.more{border-style:dashed;color:var(--gd);background:var(--card)}
.sub-card .subchips.open .subchip.more{display:none}
.sub-card button{width:100%;border:0;border-radius:12px;padding:12px;background:var(--gd);color:#fff;font:inherit;font-size:14px;font-weight:750;cursor:pointer}
.sub-card .msg{font-size:12.5px;margin-top:8px;color:var(--gd)}
@media(max-width:860px){.bgrid{grid-template-columns:1fr}}
"""


def render_briefing_v2():
    _pop = ["리튬", "니켈", "코발트", "흑연", "텅스텐", "망간", "동(구리)", "알루미늄",
            "네오디뮴", "갈륨", "우라늄", "금"]
    _rest = sorted(m for m in MIN_USES.keys() if m not in _pop)
    _chips = "".join(
        f'<button type="button" class="subchip{" ext" if m in _rest else ""}" data-m="{m}">{m}</button>'
        for m in _pop + _rest)
    news = [n for n in dedup_news(fetch_news() or []) if mineral_relevant(n)][:14]
    aud_news = fetch_audience_news() or []

    def _aud_items(aud):
        its = [n for n in aud_news if n.get("aud") == aud][:10]
        return "".join(
            f'<a href="{n.get("언론사링크", "#")}" target="_blank" rel="noopener">'
            f'<div class="nt">{n.get("제목", "")}</div>'
            f'<div class="ns">{n.get("요약", "")}</div>'
            f'<div class="nd">#{n.get("검색키워드", "")} · {n.get("발행일시", "")}</div></a>'
            for n in its) or '<div style="padding:18px;color:var(--mut)">뉴스가 없어요.</div>'

    items = "".join(
        f'<a href="{n.get("링크", "#")}" target="_blank" rel="noopener">'
        f'<div class="nt">{n.get("제목", "")}</div>'
        f'<div class="ns">{n.get("요약", "")}</div>'
        f'<div class="nd">{n.get("발행일", "")}</div></a>' for n in news
    ) or '<div style="padding:18px;color:var(--mut)">불러올 뉴스가 없어요.</div>'

    content = f"""
<div class="wrap">
  <div class="bh"><h1>오늘의 브리핑</h1></div>
  <div class="audtabs" style="margin:14px 0 12px">
    <button class="audtab on" data-a="all">전체</button>
    <button class="audtab" data-a="inv">투자자</button>
    <button class="audtab" data-a="biz">기업</button>
    <button class="audtab" data-a="con">소비자</button>
    <button class="audtab" data-a="pol">정책</button>
  </div>
  <div class="brf-ai"><div class="bl" id="aiBriefLabel">AI 애널리스트 3문장 요약 — 전체</div><div id="aiBrief">불러오는 중…</div></div>
  <div class="bgrid">
    <div>
      <div class="card ncard aud-list" id="aud-all" style="padding:4px 0">{items}</div>
      <div class="card ncard aud-list" id="aud-inv" style="padding:4px 0;display:none">{_aud_items("투자자")}</div>
      <div class="card ncard aud-list" id="aud-biz" style="padding:4px 0;display:none">{_aud_items("기업")}</div>
      <div class="card ncard aud-list" id="aud-con" style="padding:4px 0;display:none">{_aud_items("소비자")}</div>
      <div class="card ncard aud-list" id="aud-pol" style="padding:4px 0;display:none">{_aud_items("정책")}</div>
    </div>
    <div class="rail" style="position:sticky;top:78px">
      <div class="card sub-card" id="sub">
        <div class="sec-t" style="margin-bottom:2px">매일 아침 메일로 받기</div>
        <div style="font-size:13px;color:var(--mut)">광물 날씨와 주요 소식을 보내드려요.</div>
        <input id="subEmail" type="email" placeholder="이메일 주소">
        <div style="font-size:12.5px;color:var(--mut);font-weight:700;margin:2px 0 7px">관심 광물 선택 <span style="font-weight:500">(선택 · 최대 8개 — 위험도와 뉴스를 따로 담아드려요)</span></div>
        <div class="subchips" id="subChips">__SUBCHIPS__<button type="button" class="subchip more" id="chipMore">+ 전체 보기</button></div>
        <button id="subBtn">구독하기</button>
        <div class="msg" id="subMsg"></div>
      </div>
      <div class="card"><div class="r-l" style="font-size:12.5px;color:var(--mut);font-weight:700;margin-bottom:7px">🌍 지금 세계에선</div>
        <div style="font-size:13.5px;line-height:1.6" id="railGeo">불러오는 중…</div></div>
    </div>
  </div>
</div>"""

    js = r"""
(function(){
  fetch('/api/news-brief?cat=minerals').then(function(r){return r.json()}).then(function(d){
    document.getElementById('aiBrief').textContent=(d&&d.brief)?d.brief:'오늘은 새 브리핑이 없어요.';
  }).catch(function(){document.getElementById('aiBrief').textContent='브리핑을 불러오지 못했어요.';});
  fetch('/api/geo-events').then(function(r){return r.json()}).then(function(d){
    var ev=(d&&d.events&&d.events[0])||null;
    document.getElementById('railGeo').textContent=ev?((ev.loc?ev.loc+' — ':'')+(ev.why||ev['제목']||'')):'특별한 이슈가 없어요.';
  }).catch(function(){});
  var AUD_LB={all:'전체',inv:'투자자',biz:'기업',con:'소비자',pol:'정책'};
  function loadBrief(a){
    var cat=(a==='all')?'minerals':a;
    var lb=document.getElementById('aiBriefLabel');
    if(lb) lb.textContent='AI 애널리스트 3문장 요약 — '+AUD_LB[a];
    document.getElementById('aiBrief').textContent='분석 생성 중…';
    fetch('/api/news-brief?cat='+cat).then(function(r){return r.json()}).then(function(d){
      document.getElementById('aiBrief').textContent=(d&&d.brief)?d.brief:'아직 분석할 뉴스가 부족해요.';
    }).catch(function(){document.getElementById('aiBrief').textContent='분석을 불러오지 못했어요.';});
  }
  [].slice.call(document.querySelectorAll('.audtab')).forEach(function(t){
    t.addEventListener('click', function(){
      document.querySelectorAll('.audtab').forEach(function(x){x.classList.remove('on')});
      t.classList.add('on');
      document.querySelectorAll('.aud-list').forEach(function(l){l.style.display='none'});
      var el=document.getElementById('aud-'+t.dataset.a);
      if(el) el.style.display='';
      loadBrief(t.dataset.a);
    });
  });
  var sel=[];
  document.querySelectorAll('.subchip[data-m]').forEach(function(c){
    c.addEventListener('click',function(){
      var m=c.dataset.m, i=sel.indexOf(m);
      if(i>=0){ sel.splice(i,1); c.classList.remove('on'); }
      else if(sel.length<8){ sel.push(m); c.classList.add('on'); }
      else { document.getElementById('subMsg').textContent='관심 광물은 최대 8개까지 고를 수 있어요.'; }
    });
  });
  var mo=document.getElementById('chipMore');
  if(mo){ mo.addEventListener('click',function(){ document.getElementById('subChips').classList.add('open'); }); }
  var btn=document.getElementById('subBtn');
  btn.addEventListener('click',function(){
    var em=document.getElementById('subEmail').value.trim();
    btn.disabled=true; btn.textContent='처리 중…';
    fetch('/subscribe',{method:'POST',headers:{'Content-Type':'application/json'},
      body:JSON.stringify({email:em, minerals:sel})})
      .then(function(r){return r.json()}).then(function(d){
        document.getElementById('subMsg').textContent=d.message||'';
        btn.disabled=false; btn.textContent='구독하기';
      }).catch(function(){
        document.getElementById('subMsg').textContent='잠시 후 다시 시도해 주세요.';
        btn.disabled=false; btn.textContent='구독하기';
      });
  });
})();
"""
    return _v2_shell("brf", "브리핑 — K Mineral Risk",
                     content.replace("__SUBCHIPS__", _chips), V2_BRF_CSS, js)


@app.route("/m/<path:name>")
def mineral_v2(name):
    return Response(render_mineral_v2(name), mimetype="text/html")


@app.route("/briefing")
def briefing_v2():
    return Response(render_briefing_v2(), mimetype="text/html")


@app.route("/minerals.csv")
def minerals_csv():
    """광종 현황 CSV 내려받기 — '실제 서비스' 데이터 개방."""
    rows = _v2_rows()
    lines = ["광물,분류,용도,위험등급,K-RISK점수,주요수입국,수입점유율(%)"]
    for r in rows:
        lines.append(",".join([
            r["name"], r["cat"], r["use"].replace(",", "·"),
            r["grade"] or "관찰", str(r["score"] if r["score"] is not None else ""),
            r["top"] or "", str(r["share"] if r["share"] is not None else ""),
        ]))
    csv = "\ufeff" + "\n".join(lines)
    return Response(csv, mimetype="text/csv",
                    headers={"Content-Disposition": "attachment; filename=minetech_minerals.csv"})



if __name__ == "__main__":
    import sys
    sys.stdout.reconfigure(encoding='utf-8', errors='replace')
    print("핵심광물 대시보드 시작")
    _port = int(os.environ.get("PORT", "8081"))
    print(f"브라우저 접속: http://127.0.0.1:{_port}")
    print("종료: Ctrl + C")
    app.run(host="0.0.0.0", port=_port, debug=False)
