# -*- coding: utf-8 -*-
"""K-RISK 가중치 민감도 분석 — 롤링 백테스트(96관측) 조건에서 가중 조합별 ρ 비교.

결론:
  채택안(0.35/0.25/0.20/0.20) ρ=0.230 ≈ 균등(0.25×4) ρ=0.229 ≫ 수급지수 단독 ρ=0.063
  → 성능의 원천은 특정 가중치가 아니라 4축 융합 구조 자체.
  ±0.10 섭동 23개 조합 범위 0.105~0.334 — 최적화(0.334)를 취하지 않은 것은 과적합 방지 원칙.
"""
import json, os, unicodedata
import pandas as pd

BASE = os.path.dirname(os.path.abspath(__file__))
MINERALS = ["리튬", "니켈", "코발트", "동", "텅스텐", "몰리브덴"]
GEO_RISK = {"중국", "러시아", "러시아 연방", "콩고민주공화국", "미얀마"}
TOP1 = {"리튬": "호주", "니켈": "인도네시아", "코발트": "콩고민주공화국",
        "동": "칠레", "텅스텐": "중국", "몰리브덴": "중국"}
nfc = lambda s: unicodedata.normalize("NFC", s)

ssi = {}
for f in [x for x in os.listdir(BASE) if "수급안정화지수" in nfc(x) and nfc(x).endswith(".csv")]:
    for enc in ("utf-8-sig", "cp949", "euc-kr"):
        try: df = pd.read_csv(os.path.join(BASE, f), encoding=enc); break
        except Exception: continue
    ssi[str(df["광종"].iloc[0]).strip()] = dict(zip(df["기간"].astype(str).str[:7],
                                                   df["수급안정화지수"].astype(float)))

fc = json.load(open(os.path.join(BASE, "forecast_data1.json")))
def qs(nm):
    d = fc.get(nm) or {}
    sp = d.get("split") or 0
    return dict(zip(d["dates"][:sp], d["values"][:sp]))

tf = [x for x in os.listdir(BASE) if "국가별 광종" in nfc(x) and nfc(x).endswith(".csv")][0]
for enc in ("utf-8-sig", "cp949", "euc-kr"):
    try: tdf = pd.read_csv(os.path.join(BASE, tf), encoding=enc); break
    except Exception: continue
tdf["기간"] = tdf["기간"].astype(str)

def hg(nm, year):
    y = str(max(2021, min(2025, year)))
    sub = tdf[tdf["기간"].str.startswith(y) & (tdf["품목명"] == nm)]
    by = sub.groupby("국가명")["수입금액(천불)"].sum()
    by = by[by > 0]; tot = by.sum()
    if not tot: return 50.0, (20.0 if TOP1.get(nm) in GEO_RISK else 0.0)
    H = float(((by / tot) ** 2).sum() * 100)
    geo = float(by[by.index.map(lambda c: c.replace(" ", "") in {g.replace(" ", "") for g in GEO_RISK})].sum() / tot * 100)
    return H, min(100.0, geo + (20.0 if TOP1.get(nm) in GEO_RISK else 0.0))

QT = [f"{y}-{m:02d}" for y in range(2021, 2025) for m in (1, 4, 7, 10)]
obs = []
for nm in MINERALS:
    s_q = qs(nm); dates = sorted(s_q)
    for t in QT:
        if t not in s_q or ssi.get(nm, {}).get(t) is None: continue
        hist = [s_q[d] for d in dates if d < t][-8:]
        if len(hist) < 4: continue
        m = sum(hist) / len(hist)
        sd = (sum((v - m) ** 2 for v in hist) / len(hist)) ** 0.5
        V = min(100.0, sd / m * 250) if m else 0.0
        S = max(0.0, min(100.0, 100.0 - ssi[nm][t]))
        H, G = hg(nm, int(t[:4]))
        idx = dates.index(t)
        fut = [s_q[d] for d in dates[idx + 1: idx + 3]]
        if len(fut) < 2 or not s_q[t]: continue
        obs.append({"S": S, "H": H, "G": G, "V": V,
                    "abs": abs((sum(fut) / 2 / s_q[t] - 1) * 100)})

df = pd.DataFrame(obs)
def rho(w):
    sc = w[0] * df.S + w[1] * df.H + w[2] * df.G + w[3] * df.V
    return sc.rank().corr(df["abs"].rank())

print(f"관측 {len(df)}건")
print(f"채택 가중 (0.35/0.25/0.20/0.20): ρ = {rho((.35, .25, .20, .20)):.3f}")
print(f"균등 가중 (0.25×4):              ρ = {rho((.25, .25, .25, .25)):.3f}")
print(f"S 단독:                          ρ = {rho((1, 0, 0, 0)):.3f}")

rs = []
for dS in (-.1, 0, .1):
    for dH in (-.1, 0, .1):
        for dG in (-.1, 0, .1):
            w = [.35 + dS, .25 + dH, .20 + dG, 0]
            w[3] = 1 - sum(w[:3])
            if min(w) < 0.05: continue
            rs.append(rho(tuple(w)))
print(f"가중 ±0.10 섭동 {len(rs)}개 조합: ρ 범위 {min(rs):.3f} ~ {max(rs):.3f}")
