# -*- coding: utf-8 -*-
"""데이터2/ 원본 → 대시보드용 컴팩트 JSON 스냅샷 생성.
원본(데이터2/)은 .gitignore 로 배포 제외, 산출 *_data2.json 만 커밋/배포."""
import csv, json, os, datetime

BASE = "데이터2"
OUT = "."

def _open(path):
    for enc in ("utf-8-sig", "cp949", "euc-kr", "utf-8"):
        try:
            f = open(path, encoding=enc, newline="")
            f.readline(); f.seek(0)
            return f
        except Exception:
            continue
    return open(path, encoding="utf-8", errors="replace", newline="")

def _date(s):
    """YYYY-MM-DD 또는 엑셀 시리얼(45744) → date"""
    s = (s or "").strip()
    if not s:
        return None
    if s.isdigit():
        try:
            return datetime.date(1899, 12, 30) + datetime.timedelta(days=int(s))
        except Exception:
            return None
    for fmt in ("%Y-%m-%d", "%Y/%m/%d", "%Y.%m.%d"):
        try:
            return datetime.datetime.strptime(s, fmt).date()
        except Exception:
            continue
    return None

def _f(s):
    try:
        return float(str(s).replace(",", "").strip())
    except Exception:
        return None


# ── 1) 광물 파생지수 (월말 리샘플) ───────────────────────────────
def build_mineral_index():
    folder = f"{BASE}/한국광해광업공단_파생지수_20250328/파생지수"
    files = {
        "종합": "한국광해광업공단_파생지수_광물종합지수.csv",
        "에너지광물": "한국광해광업공단_파생지수_에너지광물지수.csv",
        "희소금속": "한국광해광업공단_파생지수_희소금속지수.csv",
        "메이저금속": "한국광해광업공단_파생지수_메이저금속지수.csv",
    }
    out = {"series": {}, "components": {}}
    for key, fn in files.items():
        p = os.path.join(folder, fn)
        if not os.path.exists(p):
            print("  (없음)", fn); continue
        f = _open(p); rd = csv.reader(f)
        header = next(rd)
        # 지수 컬럼 = 헤더에 '지수' 포함된 첫 컬럼
        idx_col = next((i for i, h in enumerate(header) if "지수" in h), 1)
        comp_cols = [(i, h.replace(" 가격변동률", "").replace("가격변동률", ""))
                     for i, h in enumerate(header) if "변동률" in h]
        monthly = {}      # 'YYYY-MM' -> (date, value)
        last_comps = None
        prev = None
        for row in rd:
            if len(row) <= idx_col:
                continue
            d = _date(row[0]); v = _f(row[idx_col])
            if d is None or v is None or v <= 1:
                continue
            # 깨진 행 제거: 전일 대비 50% 초과 급변은 데이터 오류로 간주
            if prev is not None and abs(v - prev) / prev > 0.5:
                continue
            prev = v
            monthly[d.strftime("%Y-%m")] = (d, v)
            if comp_cols:
                last_comps = {name: _f(row[i]) for i, name in comp_cols if i < len(row)}
        f.close()
        keys = sorted(monthly)
        months = keys
        vals = [round(monthly[k][1], 1) for k in keys]
        out["series"][key] = {"months": months, "values": vals}
        if last_comps:
            out["components"][key] = last_comps
    # 최신/등락 요약
    out["summary"] = {}
    for key, s in out["series"].items():
        v = s["values"]
        if len(v) >= 2:
            cur, prev = v[-1], v[-2]
            out["summary"][key] = {
                "latest": cur, "asof": s["months"][-1],
                "mom": round((cur - prev) / prev * 100, 2) if prev else 0,
                "yoy": round((cur - v[-13]) / v[-13] * 100, 2) if len(v) >= 13 and v[-13] else None,
            }
    return out


# ── 2) 석유제품 국가별 수입 ───────────────────────────────────────
def build_energy_import():
    p = f"{BASE}/015_한국석유공사_석유제품수입_국가별.json"
    if not os.path.exists(p):
        return None
    f = _open(p); rd = csv.reader(f); header = next(rd)
    # '국가_물량' 컬럼만 추출
    vol_cols = [(i, h.split("_")[0]) for i, h in enumerate(header)
                if h.endswith("_물량")]
    rows = []
    for r in rd:
        if not r or not r[0].strip().isdigit():
            continue
        yr = int(r[0])
        vols = {name: _f(r[i]) for i, name in vol_cols if i < len(r) and _f(r[i])}
        rows.append((yr, vols))
    f.close()
    rows.sort()
    years = [str(y) for y, _ in rows]
    # 최신연도 국가별 상위
    latest_year, latest = rows[-1]
    top = sorted(latest.items(), key=lambda kv: kv[1], reverse=True)[:8]
    return {
        "years": years,
        "latest_year": str(latest_year),
        "top_countries": [{"name": n, "vol": round(v)} for n, v in top],
        "total_latest": round(sum(latest.values())),
    }


# ── 3) 자원개발률 (자주개발률) ────────────────────────────────────
def build_resource_dev():
    p = f"{BASE}/[석유]산업통상부_자원개발률.csv"
    if not os.path.exists(p):
        return None
    f = _open(p); rd = csv.reader(f); header = next(rd)
    cols = header[1:]
    years, data = [], {c: [] for c in cols}
    for r in rd:
        if not r or not r[0].strip().isdigit():
            continue
        years.append(r[0].strip())
        for i, c in enumerate(cols):
            data[c].append(_f(r[i + 1]) if i + 1 < len(r) else None)
    f.close()
    return {"years": years, "series": data}


if __name__ == "__main__":
    print("[1] 광물 파생지수…")
    mi = build_mineral_index()
    json.dump(mi, open(f"{OUT}/mineral_index_data2.json", "w"), ensure_ascii=False)
    print("    series:", {k: len(v["months"]) for k, v in mi["series"].items()},
          "| summary:", mi.get("summary"))

    print("[2] 석유제품 국가별 수입…")
    ei = build_energy_import()
    if ei:
        json.dump(ei, open(f"{OUT}/energy_import_data2.json", "w"), ensure_ascii=False)
        print("    years:", ei["years"][0], "~", ei["years"][-1],
              "| top:", [c["name"] for c in ei["top_countries"]])

    print("[3] 자원개발률…")
    rd_ = build_resource_dev()
    if rd_:
        json.dump(rd_, open(f"{OUT}/resource_dev_data2.json", "w"), ensure_ascii=False)
        print("    years:", rd_["years"][0], "~", rd_["years"][-1],
              "| cols:", list(rd_["series"].keys()))
    print("완료.")
