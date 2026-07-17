# -*- coding: utf-8 -*-
"""K-RISK 롤링 백테스트 — 2021Q1~2024Q4 매 분기 소급 산출 → 이후 6개월 가격과 비교.

핵심 질문: "그 시점에 점수가 높았던 광물일수록, 이후 6개월 가격 충격(|변화율|)이 컸는가?"
 - 사전 점수: S=당시 수급안정화지수(월별), V=직전 8분기 변동계수, H·G=해당 연도 실측 수입구조
 - 사후 결과: |다음 2개 분기 평균가 ÷ 당시 분기가 − 1|  (조기경보이므로 방향보다 충격 크기)
 - 표본: 6광종 × 16분기 = 최대 96 관측
"""
import json, os, unicodedata
import pandas as pd

BASE = os.path.dirname(os.path.abspath(__file__))
MINERALS = ["리튬", "니켈", "코발트", "동", "텅스텐", "몰리브덴"]
GEO_RISK = {"중국", "러시아", "러시아 연방", "콩고민주공화국", "미얀마"}
TOP1 = {"리튬": "호주", "니켈": "인도네시아", "코발트": "콩고민주공화국",
        "동": "칠레", "텅스텐": "중국", "몰리브덴": "중국"}

def nfc(s): return unicodedata.normalize("NFC", s)

# ── 수급안정화지수 월별 (2018-01~2025-03) ──
ssi = {}   # {광종: {YYYY-MM: 값}}
for f in [x for x in os.listdir(BASE) if "수급안정화지수" in nfc(x) and nfc(x).endswith(".csv")]:
    for enc in ("utf-8-sig", "cp949", "euc-kr"):
        try: df = pd.read_csv(os.path.join(BASE, f), encoding=enc); break
        except Exception: continue
    nm = str(df["광종"].iloc[0]).strip()
    ssi[nm] = dict(zip(df["기간"].astype(str).str[:7], df["수급안정화지수"].astype(float)))

# ── 분기 실측가 ──
fc = json.load(open(os.path.join(BASE, "forecast_data1.json")))
def qs(nm):
    d = fc.get(nm) or {}
    sp = d.get("split") or 0
    return dict(zip((d.get("dates") or [])[:sp], (d.get("values") or [])[:sp]))

# ── 연도별 수입 구조 (2021~2025) ──
_tf = [x for x in os.listdir(BASE) if "국가별 광종" in nfc(x) and nfc(x).endswith(".csv")][0]
for enc in ("utf-8-sig", "cp949", "euc-kr"):
    try: tdf = pd.read_csv(os.path.join(BASE, _tf), encoding=enc); break
    except Exception: continue
tdf["기간"] = tdf["기간"].astype(str)

def hg(nm, year):
    y = str(max(2021, min(2025, year)))          # 2021 이전은 2021 구조로 근사
    sub = tdf[tdf["기간"].str.startswith(y) & (tdf["품목명"] == nm)]
    by = sub.groupby("국가명")["수입금액(천불)"].sum()
    by = by[by > 0]; tot = by.sum()
    if not tot: return 50.0, (20.0 if TOP1.get(nm) in GEO_RISK else 0.0)
    H = float(((by / tot) ** 2).sum() * 100)
    geo = float(by[by.index.map(lambda c: c.replace(" ", "") in {g.replace(" ", "") for g in GEO_RISK})].sum() / tot * 100)
    return H, min(100.0, geo + (20.0 if TOP1.get(nm) in GEO_RISK else 0.0))

# ── 롤링 산출 ──
QT = [f"{y}-{m:02d}" for y in range(2021, 2025) for m in (1, 4, 7, 10)]
obs = []
for nm in MINERALS:
    s_q = qs(nm)
    dates = sorted(s_q)
    for t in QT:
        if t not in s_q: continue
        ssi_v = ssi.get(nm, {}).get(t)
        if ssi_v is None: continue
        hist = [s_q[d] for d in dates if d < t][-8:]
        if len(hist) < 4: continue
        m = sum(hist) / len(hist)
        sd = (sum((v - m) ** 2 for v in hist) / len(hist)) ** 0.5
        V = min(100.0, sd / m * 250) if m else 0.0
        S = max(0.0, min(100.0, 100.0 - ssi_v))
        H, G = hg(nm, int(t[:4]))
        score = max(0.0, min(100.0, 0.35 * S + 0.25 * H + 0.20 * G + 0.20 * V))
        idx = dates.index(t)
        fut = [s_q[d] for d in dates[idx + 1: idx + 3]]
        if len(fut) < 2 or not s_q[t]: continue
        chg = (sum(fut) / len(fut) / s_q[t] - 1) * 100
        obs.append({"광물": nm, "분기": t, "score": round(score, 1),
                    "chg": round(chg, 2), "abs_chg": round(abs(chg), 2)})

df = pd.DataFrame(obs)
print(f"관측치: {len(df)}개 (광물 {df['광물'].nunique()}종 × 분기)")

# ── ① 풀링 순위상관 ──
rho_abs = df["score"].rank().corr(df["abs_chg"].rank())
rho_sgn = df["score"].rank().corr(df["chg"].rank())
n = len(df)
import math
t_stat = rho_abs * math.sqrt((n - 2) / (1 - rho_abs ** 2))
print(f"\n① 사전 점수 vs 사후 6개월 |가격변화|  스피어만 ρ = {rho_abs:.3f} (t={t_stat:.2f}, n={n})")
print(f"   (참고) 방향 포함 변화율과는 ρ = {rho_sgn:.3f}")

# ── ② 등급 버킷별 사후 충격 ──
df["등급"] = pd.cut(df["score"], [0, 40, 70, 100], labels=["안정(<40)", "주의(40-70)", "위험(70+)"])
bucket = df.groupby("등급", observed=True).agg(관측수=("abs_chg", "size"),
                                              평균충격=("abs_chg", "mean"),
                                              큰충격10pct비율=("abs_chg", lambda x: (x >= 10).mean() * 100))
print("\n② 사전 등급별 사후 6개월 가격 충격")
print(bucket.round(1).to_string())

# ── ③ 단독 지표 대비 (융합의 가치) ──
df["S만"] = None
rho_parts = {}
for part, w in [("S", None)]: pass
# S 단독: 100-SSI 만으로 같은 상관 계산
solo = []
for nm in MINERALS:
    s_q = qs(nm); dates = sorted(s_q)
    for t in QT:
        if t not in s_q or ssi.get(nm, {}).get(t) is None: continue
        idx = dates.index(t)
        fut = [s_q[d] for d in dates[idx + 1: idx + 3]]
        if len(fut) < 2 or not s_q[t]: continue
        solo.append({"s": 100 - ssi[nm][t], "abs_chg": abs((sum(fut) / 2 / s_q[t] - 1) * 100)})
sdf = pd.DataFrame(solo)
rho_solo = sdf["s"].rank().corr(sdf["abs_chg"].rank())
print(f"\n③ 벤치마크 — 수급안정화지수 단독 사용 시 ρ = {rho_solo:.3f}  (K-RISK 융합 {rho_abs:.3f})")

json.dump({"n": n, "spearman_abs": round(float(rho_abs), 3), "spearman_signed": round(float(rho_sgn), 3),
           "solo_ssi": round(float(rho_solo), 3), "t_stat": round(float(t_stat), 2),
           "buckets": json.loads(bucket.round(2).to_json(force_ascii=False)),
           "obs": obs},
          open(os.path.join(BASE, "backtest_rolling_result.json"), "w", encoding="utf-8"),
          ensure_ascii=False, indent=2)
print("\n저장: backtest_rolling_result.json")
