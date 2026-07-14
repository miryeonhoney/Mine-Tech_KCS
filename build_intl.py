# -*- coding: utf-8 -*-
"""해외 공개 데이터 스냅샷 생성 (키 불필요).
- USGS MCS 2026: 광종별 세계 생산량·매장량 → usgs_data1.json
- World Bank CMO(Pink Sheet): 월별 국제 원자재 가격 → wb_prices_data1.json
실행: python3 build_intl.py   (갱신 시 재실행 후 커밋)
"""
import csv, io, json, re, requests

USGS_URL = ("https://www.sciencebase.gov/catalog/file/get/"
            "696a75d5d4be0228872d3bf8?name=MCS2026_Commodities_Data.csv")
WB_PAGE = "https://www.worldbank.org/en/research/commodity-markets"
WB_FALLBACK = ("https://thedocs.worldbank.org/en/doc/"
               "74e8be41ceb20fa0da750cda2f6b9e4e-0050012026/related/CMO-Historical-Data-Monthly.xlsx")

UA = {"User-Agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7)"}

# USGS Commodity → (한글명, 분류)
USGS_KO = {
    "lithium": "리튬", "cobalt": "코발트", "nickel": "니켈", "graphite (natural)": "흑연",
    "rare earths": "희토류", "manganese": "망간", "copper": "동", "bauxite and alumina": "알루미늄(보크사이트)",
    "zinc": "아연", "tin": "주석", "tungsten": "텅스텐", "molybdenum": "몰리브덴",
    "chromium": "크롬", "titanium mineral concentrates": "티타늄", "antimony": "안티모니",
    "vanadium": "바나듐", "niobium (columbium)": "니오븀", "tantalum": "탄탈륨",
    "platinum-group metals": "백금족", "silver": "은", "gold": "금", "iron ore": "철",
    "magnesium compounds": "마그네슘", "zirconium and hafnium": "지르코늄",
    "gallium": "갈륨", "germanium": "게르마늄", "indium": "인듐", "bismuth": "창연(비스무트)",
}
C_KO = {"united states": "미국", "china": "중국", "australia": "호주", "chile": "칠레",
        "argentina": "아르헨티나", "brazil": "브라질", "canada": "캐나다", "russia": "러시아",
        "indonesia": "인도네시아", "india": "인도", "south africa": "남아프리카공화국",
        "congo (kinshasa)": "콩고민주공화국", "peru": "페루", "mexico": "멕시코",
        "kazakhstan": "카자흐스탄", "guinea": "기니", "vietnam": "베트남", "myanmar (burma)": "미얀마",
        "myanmar": "미얀마", "philippines": "필리핀", "new caledonia": "뉴칼레도니아",
        "zimbabwe": "짐바브웨", "mozambique": "모잠비크", "madagascar": "마다가스카르",
        "tanzania": "탄자니아", "mali": "말리", "gabon": "가봉", "turkey": "튀르키예",
        "turkiye": "튀르키예", "bolivia": "볼리비아", "japan": "일본", "south korea": "한국",
        "korea, republic of": "한국", "ukraine": "우크라이나", "norway": "노르웨이",
        "finland": "핀란드", "poland": "폴란드", "spain": "스페인", "portugal": "포르투갈",
        "morocco": "모로코", "iran": "이란", "thailand": "태국", "malaysia": "말레이시아"}

def _num(v):
    v = str(v).replace(",", "").strip()
    try:
        return float(v)
    except ValueError:
        return None

def build_usgs():
    r = requests.get(USGS_URL, headers=UA, timeout=120)
    r.raise_for_status()
    rows = list(csv.DictReader(io.StringIO(r.content.decode("utf-8", errors="replace"))))
    out = {}
    for r0 in rows:
        ckey = r0["Commodity"].strip().casefold()
        if ckey not in USGS_KO:
            continue
        ko = USGS_KO[ckey]
        stat = r0["Statistics"].strip()
        cty = r0["Country"].strip()
        year = r0["Year"].strip()
        val = _num(r0["Value"])
        unit = r0["Unit"].strip()
        if year != "2025" or val is None or stat not in ("Production", "Reserves"):
            continue
        # 광산 생산만 (정련·제련 제외)
        det = r0["Statistics_detail"].casefold()
        if stat == "Production" and not ("mine" in det or "production" == det.split(":")[0].strip()):
            pass
        m = out.setdefault(ko, {"prod": {}, "rsv": {}, "unit": unit})
        key = "prod" if stat == "Production" else "rsv"
        ck = cty.casefold()
        if ck == "world total":
            m[key]["_total"] = max(m[key].get("_total", 0), val)
        elif ck not in ("other countries",):
            nm = C_KO.get(ck, cty)
            m[key][nm] = max(m[key].get(nm, 0), val)
    result = {}
    for ko, m in out.items():
        def top(d):
            items = [(c, v) for c, v in d.items() if c != "_total"]
            if not items:
                return ["—", 0]
            c, v = max(items, key=lambda x: x[1])
            tot = d.get("_total") or sum(v2 for _, v2 in items) or 1
            return [c, round(v / tot * 100)]
        result[ko] = {
            "prod_total": round(m["prod"].get("_total", 0)),
            "prod_top": top(m["prod"]),
            "rsv_total": round(m["rsv"].get("_total", 0)),
            "rsv_top": top(m["rsv"]),
            "unit": m["unit"],
        }
    return result

# World Bank 컬럼 → 한글
WB_COLS = {
    "Aluminum": ("알루미늄", "$/mt"), "Copper": ("동(구리)", "$/mt"), "Lead": ("연(납)", "$/mt"),
    "Tin": ("주석", "$/mt"), "Nickel": ("니켈", "$/mt"), "Zinc": ("아연", "$/mt"),
    "Iron ore, cfr spot": ("철광석", "$/dmtu"), "Coal, Australian": ("석탄(호주)", "$/mt"),
    "Gold": ("금", "$/toz"), "Silver": ("은", "$/toz"), "Platinum": ("백금", "$/toz"),
}

def build_wb():
    import openpyxl
    url = WB_FALLBACK
    try:
        pg = requests.get(WB_PAGE, headers=UA, timeout=30).text
        m = re.search(r'https://[^"]*CMO-Historical-Data-Monthly[^"]*\.xlsx', pg)
        if m:
            url = m.group(0)
    except Exception:
        pass
    r = requests.get(url, headers=UA, timeout=120)
    r.raise_for_status()
    wb = openpyxl.load_workbook(io.BytesIO(r.content), read_only=True)
    ws = wb["Monthly Prices"]
    rows = list(ws.iter_rows(values_only=True))
    header = rows[4]
    col_idx = {}
    for i, h in enumerate(header):
        if h and str(h).strip() in WB_COLS:
            col_idx[i] = WB_COLS[str(h).strip()]
    months, series = [], {ko: [] for ko, _ in col_idx.values()}
    for r0 in rows[6:]:
        m0 = str(r0[0] or "")
        if not re.match(r"^\d{4}M\d{2}$", m0):
            continue
        months.append(m0.replace("M", "-"))
        for i, (ko, _) in col_idx.items():
            v = r0[i]
            series[ko].append(round(float(v), 2) if isinstance(v, (int, float)) else None)
    # 최근 15년만
    keep = 180
    months = months[-keep:]
    series = {k: v[-keep:] for k, v in series.items()}
    units = {ko: u for ko, u in col_idx.values()}
    return {"months": months, "series": series, "units": units,
            "source": "World Bank CMO (Pink Sheet)"}

if __name__ == "__main__":
    print("[1] USGS MCS 2026…")
    u = build_usgs()
    json.dump(u, open("usgs_data1.json", "w"), ensure_ascii=False)
    print("    광종:", len(u), "|", {k: v["rsv_top"] for k, v in list(u.items())[:4]})

    print("[2] World Bank Pink Sheet…")
    w = build_wb()
    json.dump(w, open("wb_prices_data1.json", "w"), ensure_ascii=False)
    print("    시계열:", list(w["series"].keys()), "|", w["months"][0], "~", w["months"][-1])
    print("완료.")
