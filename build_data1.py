# -*- coding: utf-8 -*-
"""데이터1/ 원본 → 대시보드용 컴팩트 JSON 스냅샷 생성.
원본(데이터1/)은 .gitignore 로 배포 제외, 산출 *_data1.json 만 커밋/배포.
실행: python3 build_data1.py
"""
import csv, json, os, unicodedata, datetime

BASE = "데이터1"
OUT = "."

def _files():
    return os.listdir(BASE)

def _find(key):
    """NFD 파일명 대응 — NFC 정규화 후 부분일치"""
    for f in _files():
        if key in unicodedata.normalize("NFC", f):
            return os.path.join(BASE, f)
    return None

def _open(path):
    raw = open(path, "rb").read()
    for enc in ("utf-8-sig", "cp949", "euc-kr", "utf-8"):
        try:
            return raw.decode(enc).splitlines()
        except Exception:
            continue
    return raw.decode("utf-8", errors="replace").splitlines()

def _f(s):
    try:
        return float(str(s).replace(",", "").strip())
    except Exception:
        return None


# ── 1) 광종별 가격 전망 (KOMIR 가격예측데이터 · 분기별 2013~2028) ──
FORECAST_MINERALS = ["니켈", "동", "리튬", "망간", "몰리브덴", "아연",
                     "우라늄", "유연탄", "철", "코발트", "텅스텐"]

def build_forecast(today):
    out = {}
    for m in FORECAST_MINERALS:
        p = _find(f"{m}가격예측데이터")
        if not p:
            print("  (없음)", m); continue
        rd = csv.reader(_open(p)); next(rd)
        pts = []
        for r in rd:
            if len(r) < 3: continue
            d, v = r[1].strip(), _f(r[2])
            if d and v is not None:
                pts.append((d, v))
        pts.sort()
        unit = ""
        rd2 = csv.reader(_open(p)); next(rd2)
        for r in rd2:
            if len(r) >= 4 and r[3].strip():
                unit = r[3].strip(); break
        split = sum(1 for d, _ in pts if d <= today)   # 실측/예측 경계
        out[m] = {"dates": [d[:7] for d, _ in pts],
                  "values": [round(v, 2) for _, v in pts],
                  "unit": unit, "split": split}
    return out


# ── 2) 글로벌 매장량 (KOMIR 광종별 매장량 2026.6 · 56광종 국가별) ──
EN_KO = {  # 사용자 확대 대상 광종 분류 기준
    "nickel": ("니켈", "비철금속"), "copper": ("동", "비철금속"), "bauxite": ("알루미늄(보크사이트)", "비철금속"),
    "tin": ("주석", "비철금속"), "lead": ("연(납)", "비철금속"), "zinc": ("아연", "비철금속"),
    "lithium": ("리튬", "희소금속"), "cobalt": ("코발트", "희소금속"), "manganese": ("망간", "희소금속"),
    "niobium": ("니오븀", "희소금속"), "magnesium": ("마그네슘", "희소금속"), "molybdenum": ("몰리브덴", "희소금속"),
    "vanadium": ("바나듐", "희소금속"), "tungsten": ("텅스텐", "희소금속"), "antimony": ("안티모니", "희소금속"),
    "chromium": ("크롬", "희소금속"), "tantalum": ("탄탈륨", "희소금속"), "strontium": ("스트론튬", "희소금속"),
    "zirconium": ("지르코늄", "희소금속"), "zirconium and hafnium": ("지르코늄", "희소금속"),
    "titanium(ilmentite)": ("티타늄", "희소금속"), "titanium mineral concentrates(limenite)": ("티타늄", "희소금속"),
    "rare earths": ("희토류", "희토류"),
    "iron": ("철", "기타"), "graphite": ("흑연", "기타"), "platinum": ("백금", "기타"),
    "palladium": ("팔라듐", "기타"), "gold": ("금", "기타"), "silver": ("은", "기타"),
}
C_KO = {"australia": "호주", "bolivia": "볼리비아", "canada": "캐나다", "china": "중국", "chile": "칠레",
        "peru": "페루", "russia": "러시아", "indonesia": "인도네시아", "brazil": "브라질", "india": "인도",
        "united states": "미국", "usa": "미국", "kazakhstan": "카자흐스탄", "south africa": "남아프리카공화국",
        "guinea": "기니", "vietnam": "베트남", "viet nam": "베트남", "argentina": "아르헨티나", "mexico": "멕시코",
        "poland": "폴란드", "zambia": "잠비아", "congo": "콩고민주공화국", "dr congo": "콩고민주공화국",
        "mauritania": "모리타니", "turkiye": "튀르키예", "bolivia ": "볼리비아", "other countries ": "기타국가",
        "congo(kinshasa)": "콩고민주공화국", "congo (kinshasa)": "콩고민주공화국", "turkey": "튀르키예",
        "philippines": "필리핀", "new caledonia": "뉴칼레도니아", "cuba": "쿠바", "madagascar": "마다가스카르",
        "mozambique": "모잠비크", "tanzania": "탄자니아", "zimbabwe": "짐바브웨", "ukraine": "우크라이나",
        "sweden": "스웨덴", "norway": "노르웨이", "finland": "핀란드", "spain": "스페인", "portugal": "포르투갈",
        "greece": "그리스", "iran": "이란", "morocco": "모로코", "gabon": "가봉", "ghana": "가나",
        "myanmar": "미얀마", "malaysia": "말레이시아", "thailand": "태국", "south korea": "한국",
        "korea": "한국", "japan": "일본", "mongolia": "몽골", "saudi arabia": "사우디아라비아",
        "world total": "세계합계", "other countries": "기타국가", "others": "기타국가"}

def build_reserves():
    p = _find("광종별 매장량")
    if not p: return None
    rd = csv.reader(_open(p)); next(rd)
    acc = {}   # ko -> {"cat":, "unit":, "countries": {c: v}}
    for r in rd:
        if len(r) < 4: continue
        en = r[0].strip().casefold()
        if en not in EN_KO: continue
        ko, cat = EN_KO[en]
        unit = r[1].strip().casefold()
        c_en = r[2].strip().casefold()
        v = _f(r[3])
        if v is None or v <= 0: continue
        mult = {"ton": 1, "k ton": 1000, "kton": 1000, "kg": 0.001}.get(unit, 1)  # → ton 환산
        c = C_KO.get(c_en, r[2].strip())
        if c in ("세계합계",): continue
        a = acc.setdefault(ko, {"cat": cat, "countries": {}})
        a["countries"][c] = max(a["countries"].get(c, 0), v * mult)
    out = []
    for ko, a in acc.items():
        cs = sorted(a["countries"].items(), key=lambda x: -x[1])
        total = sum(v for _, v in cs)
        top = [{"c": c, "v": round(v)} for c, v in cs if c != "기타국가"][:8]
        out.append({"name": ko, "cat": a["cat"], "total": round(total),
                    "top1": top[0]["c"] if top else "", "countries": top})
    out.sort(key=lambda x: -x["total"])
    return out


# ── 3) 조달청 비축물자 비철금속 일일가격 (LME) ──
def build_ppa():
    p = _find("비철금속_일일가격")
    if not p: return None
    date = ""
    nfc = unicodedata.normalize("NFC", os.path.basename(p))
    for tok in nfc.replace(".csv", "").split("_"):
        if tok.isdigit() and len(tok) == 8:
            date = f"{tok[:4]}-{tok[4:6]}-{tok[6:]}"
    rd = csv.reader(_open(p)); next(rd)
    items, lme = [], None
    for r in rd:
        if len(r) < 6: continue
        items.append({"name": r[0].strip(), "close": _f(r[2]), "chg": _f(r[3])})
        if lme is None:
            lme = {"idx": _f(r[4]), "chg": _f(r[5])}
    return {"date": date, "lme": lme, "items": items}


# ── 4) 철강 원자재 가격동향 (산업통상부 · 월별) ──
def build_steel():
    p = _find("철강원자재 가격동향")
    if not p: return None
    rows = list(csv.reader(_open(p)))
    header = [h.strip() for h in rows[0]]
    data = [r for r in rows[1:] if r and r[0].strip()]
    data.sort(key=lambda r: r[0])
    months = [r[0].strip() for r in data]
    series = {}
    for i, h in enumerate(header[1:], start=1):
        vals = [_f(r[i]) if i < len(r) else None for r in data]
        if any(v is not None for v in vals):
            series[h] = vals
    return {"months": months, "series": series}


# ── 5) 국내 광산 통계 + 폐광산 실태 ──
def build_mines():
    out = {}
    p = _find("전국광산 통계 현황(광종별")
    if p:
        rd = csv.reader(_open(p)); next(rd)
        rows = [r for r in rd if len(r) >= 5]
        years = sorted({r[1][:4] for r in rows})
        trend = {"years": years, "가행": [], "폐광": []}
        def _state(nm):
            if "가행" in nm: return "가행"
            if "휴지" in nm: return "휴지"
            if "폐광" in nm: return "폐광"
            return None
        for y in years:
            for st in ("가행", "폐광"):
                tot = sum(sum(_f(x) or 0 for x in r[2:5])
                          for r in rows if r[1].startswith(y) and _state(r[0]) == st)
                trend[st].append(round(tot))
        latest = max(years)
        comp = {}
        for r in rows:
            if r[1].startswith(latest):
                key = _state(r[0])
                if not key: continue
                comp[key] = {"금속": round(_f(r[2]) or 0), "비금속": round(_f(r[3]) or 0),
                             "석탄": round(_f(r[4]) or 0)}
        out["stats"] = {"trend": trend, "latest_year": latest, "comp": comp}
    p = _find("전국 폐광산 위치")
    if p:
        rd = csv.reader(_open(p)); next(rd)
        sido = {}
        n = 0
        for r in rd:
            if len(r) < 4: continue
            n += 1
            sido[r[3].strip()] = sido.get(r[3].strip(), 0) + 1
        top = sorted(sido.items(), key=lambda x: -x[1])[:10]
        out["closed"] = {"total": n, "sido": [{"s": s, "n": c} for s, c in top]}
    return out or None


# ── 6) 시장전망지표 (동·아연 · 월별) ──
def build_outlook():
    out = {}
    for m in ("동", "아연"):
        p = _find(f"시장전망지표_{m}")
        if not p: continue
        rd = csv.reader(_open(p)); next(rd)
        pts = sorted((r[1].strip(), _f(r[2])) for r in rd if len(r) >= 3 and _f(r[2]) is not None)
        out[m] = {"months": [d[:7] for d, _ in pts][-60:],
                  "values": [round(v, 1) for _, v in pts][-60:]}
    return out or None


if __name__ == "__main__":
    today = datetime.date.today().isoformat()

    print("[1] 가격 전망 (11광종)…")
    fc = build_forecast(today)
    json.dump(fc, open(f"{OUT}/forecast_data1.json", "w"), ensure_ascii=False)
    print("    광종:", list(fc.keys()))

    print("[2] 글로벌 매장량 (확대)…")
    rv = build_reserves()
    if rv:
        json.dump(rv, open(f"{OUT}/reserves_data1.json", "w"), ensure_ascii=False)
        print("    광종 수:", len(rv), "| 상위:", [x["name"] for x in rv[:5]])

    print("[3] 조달청 비축물자 일일가격…")
    ppa = build_ppa()
    if ppa:
        json.dump(ppa, open(f"{OUT}/ppa_data1.json", "w"), ensure_ascii=False)
        print("    기준일:", ppa["date"], "| 품목:", [i["name"] for i in ppa["items"]])

    print("[4] 철강 원자재 가격…")
    st = build_steel()
    if st:
        json.dump(st, open(f"{OUT}/steel_data1.json", "w"), ensure_ascii=False)
        print("    기간:", st["months"][0], "~", st["months"][-1], "| 지표:", list(st["series"].keys()))

    print("[5] 국내 광산 통계…")
    mn = build_mines()
    if mn:
        json.dump(mn, open(f"{OUT}/mines_data1.json", "w"), ensure_ascii=False)
        print("    ", {k: (v if k != "closed" else v["total"]) for k, v in mn.items()} if mn else "")

    print("[6] 시장전망지표…")
    ol = build_outlook()
    if ol:
        json.dump(ol, open(f"{OUT}/outlook_data1.json", "w"), ensure_ascii=False)
        print("    ", list(ol.keys()))
    print("완료.")
