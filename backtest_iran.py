# -*- coding: utf-8 -*-
"""K-RISK 사건 검증 2 — 미국-이란 충돌(2026-02-28) 직전 시점 소급 적용.

방법 (우크라이나 검증과 동일 설계):
  T0 = 2026-01 (충돌 직전 마지막 분기 관측)
  사전 점수(T0 기준):
    S = 100 − 수급안정화지수(공표 최신 2025-03 — 시차 한계로 명시)
    V = 250 × 직전 8분기 가격 변동계수 (T0 이전 실측만)
    H, G = 2025년 실측 수입 구조 (관세청 국가별 광종 수출입)
  사후 결과:
    가격 변화율 = 충돌 후 2개 분기(2026-04·07) 평균가 ÷ 직전 분기(2026-01) − 1
  검증:
    사전 점수 순위 vs 사후 |변화| 및 변화 순위 (스피어만)
"""
import json, os, unicodedata
import pandas as pd

BASE = os.path.dirname(os.path.abspath(__file__))
T0 = "2026-01"
MINERALS = ["리튬", "니켈", "코발트", "동", "텅스텐", "몰리브덴"]
GEO_RISK = {"중국", "러시아", "러시아 연방", "콩고민주공화국", "미얀마", "이란"}

def nfc(s): return unicodedata.normalize("NFC", s)

# ── 1) 수급안정화지수 (공표 최신값) ──
ssi = {}
for f in [x for x in os.listdir(BASE) if "수급안정화지수" in nfc(x) and nfc(x).endswith(".csv")]:
    for enc in ("utf-8-sig", "cp949", "euc-kr"):
        try: df = pd.read_csv(os.path.join(BASE, f), encoding=enc); break
        except Exception: continue
    nm = str(df["광종"].iloc[0]).strip()
    df = df.sort_values("기간")
    ssi[nm] = (str(df["기간"].iloc[-1])[:7], float(df["수급안정화지수"].iloc[-1]))
print("① 수급안정화지수 최신 공표값:", {k: v for k, v in ssi.items()})

# ── 2) 분기 실측가 ──
fc = json.load(open(os.path.join(BASE, "forecast_data1.json")))
def q_series(nm):
    d = fc.get(nm) or {}
    sp = d.get("split") or 0
    return dict(zip((d.get("dates") or [])[:sp], (d.get("values") or [])[:sp]))

def cv8_before(nm, t0):
    s = sorted((d, v) for d, v in q_series(nm).items() if d < t0 and isinstance(v, (int, float)))[-8:]
    vals = [v for _, v in s]
    if len(vals) < 4: return None
    m = sum(vals) / len(vals)
    sd = (sum((v - m) ** 2 for v in vals) / len(vals)) ** 0.5
    return min(100.0, sd / m * 250) if m else 0.0

def price_change_after(nm):
    s = q_series(nm)
    base = s.get(T0)
    after = [v for v in (s.get("2026-04"), s.get("2026-07")) if isinstance(v, (int, float))]
    if not (isinstance(base, (int, float)) and base and after): return None
    return (sum(after) / len(after) / base - 1) * 100

# ── 3) H·G — 2025년 수입 구조 ──
_tf = [x for x in os.listdir(BASE) if "국가별 광종" in nfc(x) and nfc(x).endswith(".csv")][0]
for enc in ("utf-8-sig", "cp949", "euc-kr"):
    try: tdf = pd.read_csv(os.path.join(BASE, _tf), encoding=enc); break
    except Exception: continue
tdf["기간"] = tdf["기간"].astype(str)
t25 = tdf[tdf["기간"].str.startswith("2025")]

def hg_2025(nm):
    sub = t25[t25["품목명"] == nm]
    by = sub.groupby("국가명")["수입금액(천불)"].sum()
    by = by[by > 0]; tot = by.sum()
    if not tot: return 50.0, 0.0, "?"
    top1 = by.idxmax()
    H = float(((by / tot) ** 2).sum() * 100)
    geo = float(by[by.index.map(lambda c: c.replace(" ", "") in {g.replace(" ", "") for g in GEO_RISK})].sum() / tot * 100)
    G = min(100.0, geo + (20.0 if top1.replace(" ", "") in {g.replace(" ", "") for g in GEO_RISK} else 0.0))
    return H, G, top1

rows = []
for nm in MINERALS:
    if nm not in ssi: continue
    ssi_month, ssi_v = ssi[nm]
    S = max(0.0, min(100.0, 100.0 - ssi_v))
    V = cv8_before(nm, T0)
    H, G, top1 = hg_2025(nm)
    if V is None: continue
    score = max(0.0, min(100.0, 0.35 * S + 0.25 * H + 0.20 * G + 0.20 * V))
    chg = price_change_after(nm)
    rows.append({"광물": nm, "S(수급)": round(S, 1), "H(집중)": round(H, 1),
                 "G(지정학)": round(G, 1), "V(변동)": round(V, 1), "1위공급국": top1,
                 "사전점수": round(score, 1),
                 "충돌후 가격변화%": round(chg, 1) if chg is not None else None})

df = pd.DataFrame(rows).sort_values("사전점수", ascending=False)
print("\n② 2026-01 시점 소급 K-RISK vs 충돌(2026-02-28) 후 가격")
print(df.to_string(index=False))

v = df.dropna(subset=["충돌후 가격변화%"])
rho_sgn = v["사전점수"].rank().corr(v["충돌후 가격변화%"].rank())
rho_abs = v["사전점수"].rank().corr(v["충돌후 가격변화%"].abs().rank())
print(f"\n③ 스피어만 — 방향 포함 ρ={rho_sgn:.2f} · 충격 크기(|변화|) ρ={rho_abs:.2f} (표본 {len(v)}광종)")

json.dump({"t0": T0, "event": "미국-이란 충돌 격화 (2026-02-28)", "ssi_asof": {k: v0[0] for k, v0 in ssi.items()},
           "rows": rows, "spearman_signed": round(float(rho_sgn), 3), "spearman_abs": round(float(rho_abs), 3)},
          open(os.path.join(BASE, "backtest_iran_result.json"), "w", encoding="utf-8"),
          ensure_ascii=False, indent=2)
print("저장: backtest_iran_result.json")
