# -*- coding: utf-8 -*-
"""K-RISK 백테스트 — 우크라이나 침공(2022-02-24) 직전 시점 소급 적용.

방법:
  T0 = 2022-01 (침공 직전 마지막 관측월)
  사전 점수(T0 기준):
    S = 100 − 수급안정화지수(2022-01)          [KOMIR 월별 CSV, 2018~]
    V = 250 × 직전 24개월 가격 변동계수          [KOMIR 분기 실측가]
    H, G = 수입집중도·지정학                     [최신 수입 구조로 고정 — 한계로 명시]
  사후 결과:
    가격 변화율 = 침공 후 2개 분기(2022Q2·Q3) 평균가 ÷ 직전 분기(2021Q4) − 1
  검증:
    사전 점수 순위 vs 사후 상승률 순위 (스피어만 순위상관)
"""
import json, os, glob, unicodedata
import pandas as pd

BASE = os.path.dirname(os.path.abspath(__file__))
T0 = "2022-01"

def nfc(s): return unicodedata.normalize("NFC", s)

# ── 1) 수급안정화지수 (2022-01 값) ── (macOS NFD 파일명 → NFC 정규화 매칭)
ssi = {}
_files = [os.path.join(BASE, f) for f in os.listdir(BASE)
          if "수급안정화지수" in nfc(f) and nfc(f).endswith(".csv")]
for f in _files:
    for enc in ("utf-8-sig", "cp949", "euc-kr"):
        try:
            df = pd.read_csv(f, encoding=enc); break
        except Exception: continue
    nm = str(df["광종"].iloc[0]).strip()
    df["기간"] = df["기간"].astype(str).str[:7]
    row = df[df["기간"] == T0]
    if len(row):
        ssi[nm] = float(row["수급안정화지수"].iloc[0])
print("① 2022-01 수급안정화지수:", {k: round(v, 1) for k, v in ssi.items()})

# ── 2) 분기 실측가 (KOMIR 가격예측 스냅샷의 실측 구간) ──
fc = json.load(open(os.path.join(BASE, "forecast_data1.json")))

def q_series(nm):
    d = fc.get(nm) or {}
    dates, vals, split = d.get("dates") or [], d.get("values") or [], d.get("split") or 0
    return list(zip(dates[:split], vals[:split]))          # 실측만

def cv24_before(nm, t0):
    """t0 이전 8개 분기(≈24개월) 변동계수 ×250 (상한 100)."""
    s = [(d, v) for d, v in q_series(nm) if d < t0 and isinstance(v, (int, float))][-8:]
    vals = [v for _, v in s]
    if len(vals) < 4: return None
    m = sum(vals) / len(vals)
    sd = (sum((v - m) ** 2 for v in vals) / len(vals)) ** 0.5
    return min(100.0, sd / m * 250) if m else 0.0

def price_change_after(nm):
    """직전 분기(2021Q4≈2021-10) 대비 침공 후 2개 분기(2022-04·07) 평균 변화율(%)."""
    s = dict(q_series(nm))
    base = s.get("2021-10")
    after = [s.get("2022-04"), s.get("2022-07")]
    after = [v for v in after if isinstance(v, (int, float))]
    if not (isinstance(base, (int, float)) and base and after): return None
    return (sum(after) / len(after) / base - 1) * 100

# ── 3) H·G — 2021년 실측 수입 구조 (KOMIR 국가별 광종 수출입) ──
_tf = [x for x in os.listdir(BASE) if "국가별 광종" in nfc(x) and nfc(x).endswith(".csv")][0]
for enc in ("utf-8-sig", "cp949", "euc-kr"):
    try:
        tdf = pd.read_csv(os.path.join(BASE, _tf), encoding=enc); break
    except Exception: continue
tdf["기간"] = tdf["기간"].astype(str)
t21 = tdf[tdf["기간"].str.startswith("2021")]
GEO_RISK = {"중국", "러시아", "러시아 연방", "콩고민주공화국", "미얀마"}
TOP1_2021 = {"리튬": "호주", "니켈": "인도네시아", "코발트": "콩고민주공화국",
             "동": "칠레", "텅스텐": "중국", "몰리브덴": "중국"}

def hg_2021(nm):
    sub = t21[t21["품목명"] == nm]
    by = sub.groupby("국가명")["수입금액(천불)"].sum()
    by = by[by > 0]
    tot = by.sum()
    if not tot: return 50.0, (20.0 if TOP1_2021.get(nm) in GEO_RISK else 0.0)
    H = float(((by / tot) ** 2).sum() * 100)
    geo = float(by[by.index.map(lambda c: c.replace(" ", "") in {g.replace(" ", "") for g in GEO_RISK})].sum() / tot * 100)
    G = min(100.0, geo + (20.0 if TOP1_2021.get(nm) in GEO_RISK else 0.0))
    return H, G

rows = []
for nm in ["리튬", "니켈", "코발트", "동", "텅스텐", "몰리브덴"]:
    if nm not in ssi: continue
    S = max(0.0, min(100.0, 100.0 - ssi[nm]))
    V = cv24_before(nm, T0)
    H, G = hg_2021(nm)
    if V is None: continue
    score = max(0.0, min(100.0, 0.35 * S + 0.25 * H + 0.20 * G + 0.20 * V))
    chg = price_change_after(nm)
    rows.append({"광물": nm, "S(수급)": round(S, 1), "H(집중)": round(H, 1),
                 "G(지정학)": round(G, 1), "V(변동)": round(V, 1),
                 "사전점수": round(score, 1), "침공후 가격변화%": round(chg, 1) if chg is not None else None})

df = pd.DataFrame(rows).sort_values("사전점수", ascending=False)
print("\n② 2022-01 시점 소급 K-RISK vs 침공 후 6개월 가격")
print(df.to_string(index=False))

# ── 4) 스피어만 순위상관 ──
v = df.dropna(subset=["침공후 가격변화%"])
rho = v["사전점수"].rank().corr(v["침공후 가격변화%"].rank())
print(f"\n③ 스피어만 순위상관: {rho:.2f}  (표본 {len(v)}광종)")
json.dump({"t0": T0, "rows": rows, "spearman": round(float(rho), 3)},
          open(os.path.join(BASE, "backtest_2022_result.json"), "w", encoding="utf-8"),
          ensure_ascii=False, indent=2)
print("결과 저장: backtest_2022_result.json")
