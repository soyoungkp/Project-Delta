"""
Major Item 조회 기능.

지정한 Major Item 리스트에 대해 Previous/Revised BQ 에서 해당 아이템을 찾아
(Description 키워드 매칭) Quantity 와 Unit Rate / Cost 를 집계한다.

기존 비교 로직(compare_bq / compare_rate)은 수정하지 않고, 추출 함수만 재사용한다:
  - bq_to_long(path, drop_zero=False) : FWBS×PABS Quantity (long)
  - compare_rate.extract_rates(path)  : FWBS×Compare_Item 단가/원가 (long, blank=0)

매칭은 대소문자/공백/괄호/복수형/오타 변형을 흡수하도록
normalize + partial keyword(substring) 방식으로 한다.
"""

import io
import re

import numpy as np
import pandas as pd

from bq_to_long import bq_to_long
from compare_rate import extract_rates, ITEM_ORDER

# (Major Item, Unit, [검색 키워드/별칭])
MAJOR_ITEMS = [
    ("Excavation (above water)", "M3", ["excavation above water", "excavation"]),
    ("Excavation (Below water)", "M3", ["excavation below water", "underwater excavation"]),
    ("Disposal", "M3", ["disposal", "dispose"]),
    ("Excavated backfilling", "M3", ["excavated backfilling", "backfilling with excavated soil"]),
    ("Borrow Soil backfilling", "M3", ["borrow soil backfilling", "imported soil backfilling"]),
    ("Borrow Sand backfilling", "M3", ["borrow sand backfilling", "sand backfilling"]),
    ("Concrete", "M3", ["concrete", "conrete"]),
    ("Lean concrete", "M3", ["lean concrete"]),
    ("Rebar", "Ton", ["rebar", "reinforcement bar", "reinforcing steel"]),
    ("Form (General)", "M2", ["form", "formwork", "general form"]),
    ("Steel", "KG", ["steel", "structural steel"]),
    ("Subgrade", "M2", ["subgrade"]),
    ("Sub Base Course (200mm)", "M2", ["sub base course", "subbase", "200mm"]),
    ("Asphalt Base Course (230mm)", "M2", ["asphalt base course", "230mm"]),
    ("Asphalt Base Course (150mm)", "M2", ["asphalt base course", "150mm"]),
    ("Surface Base Course (90mm)", "M2", ["surface base course", "90mm"]),
    ("Wearing Base Course (50mm)", "M2", ["wearing base course", "50mm"]),
]
ITEM_UNIT = {name: unit for name, unit, _ in MAJOR_ITEMS}
ITEM_NAMES = [name for name, _, _ in MAJOR_ITEMS]

# Compare_Item -> 깔끔한 컬럼 베이스명
ITEM_COL = {
    "Manhour U/P": "Manhour_UP",
    "Labor U/P": "Labor_UP",
    "Material U/P": "Material_UP",
    "Equipment U/P": "Equipment_UP",
    "Tool & Consume U/P": "Tool_Consume_UP",
    "Indirect U/P": "Indirect_UP",
    "Total U/P": "Total_UP",
    "Total Cost": "Total_Cost",
}
BROAD_MATCH_THRESHOLD = 60   # 이 이상 매칭되면 '광범위 매칭' 경고

# Change Rate 집계용 단가 베이스(7종, Total Cost 는 별도) — Indirect 포함
UP_BASES = ["Manhour_UP", "Labor_UP", "Material_UP", "Equipment_UP",
            "Tool_Consume_UP", "Indirect_UP", "Total_UP"]
RATE_TOL = 1e-4   # 단가 동일 판정 허용오차


def _norm(s):
    """소문자화 + 영숫자/점 외 문자를 공백으로 → 공백 단일화."""
    s = re.sub(r"[^a-z0-9.]+", " ", str(s).lower())
    return re.sub(r"\s+", " ", s).strip()


def _kw_pattern(k):
    """키워드를 단어 경계로 매칭(부분문자열 오매칭 방지).

    예: '50mm' 이 '350mm' 안에서 매칭되지 않도록 앞뒤를 영숫자 비경계로 제한.
    """
    return re.compile(r"(?<![a-z0-9])" + re.escape(k) + r"(?![a-z0-9])")


# (name, [(키워드, 컴파일된 패턴), ...])
_NORM_ITEMS = [
    (name, [(nk, _kw_pattern(nk)) for nk in (_norm(k) for k in kws) if nk])
    for name, _, kws in MAJOR_ITEMS
]


def match_major_items(description):
    """description 과 매칭되는 Major Item 이름 리스트(단어경계 부분일치)."""
    d = _norm(description)
    if not d:
        return []
    return [name for name, kws in _NORM_ITEMS if any(p.search(d) for _, p in kws)]


def _best_match(description):
    """가장 구체적인 Major Item 1개로 배정.

    점수 = (매칭된 키워드 수, 매칭 키워드 길이 합) 최대값.
    동점이면 ambiguous(애매)로 표시. 반환: (assigned|None, candidates, is_tie)
    """
    d = _norm(description)
    if not d:
        return None, [], False
    scored = []
    for name, kws in _NORM_ITEMS:
        matched = [nk for nk, p in kws if p.search(d)]
        if matched:
            scored.append((name, (len(matched), sum(len(k) for k in matched))))
    if not scored:
        return None, [], False
    best = max(s for _, s in scored)
    top = [name for name, s in scored if s == best]
    candidates = [name for name, _ in scored]
    return top[0], candidates, len(top) > 1


def _src(x):
    """bytes 면 매번 새 BytesIO, 아니면(경로) 그대로."""
    return io.BytesIO(x) if isinstance(x, (bytes, bytearray)) else x


def _fwbs_frame(file):
    """파일 1개 → FWBS 단위 [Description, Unit, Special_Note, Qty, 8개 단가/원가] + PABS."""
    ql = bq_to_long(_src(file), drop_zero=False)
    qsum = ql.groupby("FWBS", as_index=False)["Quantity"].sum()
    nz = ql[ql["Quantity"] != 0]
    pabs = nz.groupby("FWBS")["PABS"].apply(lambda s: sorted(set(s)))
    # Special Note: FWBS 별 첫 비어있지 않은 값
    if "Special_Note" in ql.columns:
        note = ql.groupby("FWBS")["Special_Note"].apply(
            lambda s: next((str(x).strip() for x in s if str(x).strip()), ""))
    else:
        note = pd.Series(dtype=object)

    rates = extract_rates(_src(file))
    meta = rates.drop_duplicates("FWBS")[["FWBS", "Description", "Unit"]]
    wide = rates.pivot_table(index="FWBS", columns="Compare_Item",
                             values="Value", aggfunc="first", observed=True)
    wide.columns = [str(c) for c in wide.columns]
    wide = wide.reset_index()

    df = meta.merge(qsum, on="FWBS", how="outer").merge(wide, on="FWBS", how="outer")
    df["Quantity"] = df["Quantity"].fillna(0)
    for item in ITEM_ORDER:
        if item in df.columns:
            df[item] = df[item].fillna(0)
        else:
            df[item] = 0.0
    df["PABS_set"] = df["FWBS"].map(pabs).apply(lambda v: v if isinstance(v, list) else [])
    df["Special_Note"] = df["FWBS"].map(note).fillna("")
    return df


def _coalesce(a, b):
    """a(우선) 비었으면 b. 문자열 기준."""
    a = "" if a is None or (isinstance(a, float) and pd.isna(a)) else str(a)
    return a if a.strip() else ("" if b is None else str(b))


def _pct(diff, base):
    """diff/base. base 가 0/NaN 이면 NaN (blank)."""
    base = pd.to_numeric(base, errors="coerce")
    return np.where((~base.isna()) & (base != 0), diff / base.replace(0, np.nan), np.nan)


# all-zero dummy 판정 컬럼 (수량 + Total U/P + 6개 구성단가의 Previous/Revised)
_DUMMY_ZERO_COLS = [
    "Previous_Qty", "Revised_Qty", "Total_UP_Previous", "Total_UP_Revised",
    "Manhour_UP_Previous", "Manhour_UP_Revised", "Labor_UP_Previous", "Labor_UP_Revised",
    "Material_UP_Previous", "Material_UP_Revised",
    "Equipment_UP_Previous", "Equipment_UP_Revised",
    "Tool_Consume_UP_Previous", "Tool_Consume_UP_Revised",
    "Indirect_UP_Previous", "Indirect_UP_Revised",
]


def _all_zero_dummy_mask(df, tol=RATE_TOL):
    """수량·단가·변경이 모두 0인 all-zero dummy 행 마스크.

    하나라도 0이 아닌 값(신규/삭제 물량, 단가 입력 등)이 있으면 False(유지).
    """
    mask = pd.Series(True, index=df.index)
    for c in _DUMMY_ZERO_COLS:
        mask &= pd.to_numeric(df[c], errors="coerce").fillna(0).abs() <= tol
    return mask


def build_major_item_view(previous_file, revised_file, include_dummy=False):
    """Major Item 조회 결과.

    include_dummy=False(기본): 수량·단가·변경이 모두 0인 all-zero dummy 행을 제외.
    반환 dict: summary, detail, rate_mismatch, unmatched, ambiguous, broad
    """
    prev = _fwbs_frame(previous_file)
    rev = _fwbs_frame(revised_file)

    merged = prev.merge(rev, on="FWBS", how="outer", suffixes=("__P", "__R"))

    rows = []
    for _, r in merged.iterrows():
        desc = _coalesce(r.get("Description__R"), r.get("Description__P"))
        unit = _coalesce(r.get("Unit__R"), r.get("Unit__P"))
        note = _coalesce(r.get("Special_Note__R"), r.get("Special_Note__P"))
        assigned, candidates, is_tie = _best_match(desc)
        if assigned is None:
            continue  # Major Item 아님
        pabs_set = sorted(set(
            (r.get("PABS_set__P") or []) + (r.get("PABS_set__R") or [])))
        # Rate Consistency Key = Major Item + Unit + Activity(Description) + Special Note
        rc_key = " | ".join([assigned, _norm(unit), _norm(desc), _norm(note)])
        row = {
            "Major_Item": assigned,
            "FWBS": r["FWBS"],
            "Description": desc,
            "Unit": unit,
            "Special_Note": note,
            "Rate_Consistency_Key": rc_key,
            "PABS": ", ".join(pabs_set),
            "Previous_Qty": float(r.get("Quantity__P") or 0),
            "Revised_Qty": float(r.get("Quantity__R") or 0),
            "_candidates": candidates,
            "_is_tie": is_tie,
        }
        for item, base in ITEM_COL.items():
            row[f"{base}_Previous"] = float(r.get(f"{item}__P") or 0)
            row[f"{base}_Revised"] = float(r.get(f"{item}__R") or 0)
        rows.append(row)

    detail_full = pd.DataFrame(rows)
    if not detail_full.empty:
        detail_full["Qty_Difference"] = (
            detail_full["Revised_Qty"] - detail_full["Previous_Qty"]).round(6)

    # 미매칭(등록했지만 매칭 0)은 원본 매칭 기준 — dummy 제외 영향 없음
    matched_names = set() if detail_full.empty else set(detail_full["Major_Item"])
    unmatched = [n for n in ITEM_NAMES if n not in matched_names]

    # all-zero dummy 행 제외 (기본). 표시/요약/진단(광범위·중복후보)은 이 기준으로 산출
    kept = detail_full
    if not detail_full.empty and not include_dummy:
        kept = detail_full[~_all_zero_dummy_mask(detail_full)]

    # --- 진단: 애매(중복 후보) / 광범위 ---
    if kept.empty:
        ambiguous = pd.DataFrame(
            columns=["FWBS", "Description", "Assigned", "Candidates", "Tie"])
        broad = []
    else:
        amb = kept[kept["_candidates"].apply(len) > 1]
        ambiguous = pd.DataFrame({
            "FWBS": amb["FWBS"],
            "Description": amb["Description"],
            "Assigned": amb["Major_Item"],
            "Candidates": amb["_candidates"].apply(lambda c: ", ".join(c)),
            "Tie": amb["_is_tie"],
        }).reset_index(drop=True)
        counts = kept["Major_Item"].value_counts()
        broad = [(n, int(c)) for n, c in counts.items() if c >= BROAD_MATCH_THRESHOLD]

    # 요약 + Rate Consistency 는 (dummy 제외된) 전체 항목 기준 계산
    summary, rate_mismatch = summarize_major_items(kept)

    # 공개 Detail 테이블 (기존 컬럼 구성 유지)
    detail = (kept.drop(columns=["_candidates", "_is_tie"])
              if not kept.empty else _empty_detail())
    if not detail.empty:
        detail["Major_Item"] = pd.Categorical(detail["Major_Item"],
                                              categories=ITEM_NAMES, ordered=True)
        detail = detail[_DETAIL_COLS].sort_values(["Major_Item", "FWBS"]).reset_index(drop=True)

    return {"summary": summary, "detail": detail, "rate_mismatch": rate_mismatch,
            "unmatched": unmatched, "ambiguous": ambiguous, "broad": broad}


def _cap_codes(codes, n=20):
    if not codes:
        return ""
    if len(codes) <= n:
        return ", ".join(codes)
    return ", ".join(codes[:n]) + f" ...(+{len(codes) - n})"


def _distinct_within_tol(vals, tol):
    """tolerance 내 서로 다른 대표값 리스트."""
    reps = []
    for v in sorted(vals):
        if all(abs(v - rep) > tol for rep in reps):
            reps.append(v)
    return reps


# Revised 구성단가(6종) — dummy/parent(전부 0) 판정용
_COMPONENT_REV = ["Manhour_UP_Revised", "Labor_UP_Revised", "Material_UP_Revised",
                  "Equipment_UP_Revised", "Tool_Consume_UP_Revised", "Indirect_UP_Revised"]


def _valid_rows(group, tol=RATE_TOL):
    """0단가 parent/dummy 행(Revised Total U/P 와 구성단가 모두 0) 제외한 행."""
    dummy = group["Total_UP_Revised"].abs() <= tol
    for c in _COMPONENT_REV:
        dummy &= group[c].abs() <= tol
    return group[~dummy]


def _key_consistency(group, tol=RATE_TOL):
    """Rate_Consistency_Key 그룹 1개의 일치 여부.

    "단일 항목" / "일치" / "상이" / "검토 필요"
    """
    valid = _valid_rows(group, tol)
    if valid.empty:                       # 전부 0단가 → 의미 불명
        return "검토 필요"
    if len(valid) == 1:
        return "단일 항목"
    reps = _distinct_within_tol(
        pd.to_numeric(valid["Total_UP_Revised"], errors="coerce").dropna().tolist(), tol)
    if not reps:
        return "검토 필요"
    return "일치" if len(reps) == 1 else "상이"


def _wavg(vals, weights):
    """물량 가중평균 단가 (가중치 합 0 이면 단순평균)."""
    v = pd.to_numeric(vals, errors="coerce").fillna(0)
    w = pd.to_numeric(weights, errors="coerce").fillna(0)
    denom = float(w.sum())
    if denom > 0:
        return float((v * w).sum() / denom)
    return float(v.mean()) if len(v) else 0.0


def summarize_major_items(full):
    """전체 FWBS 단위 프레임 → Major Item 요약 + Rate Mismatch 상세.

    Rate Consistency 는 Rate_Consistency_Key(Major+Unit+Activity+Special Note) 단위로 판단.
    반환: (summary_df, rate_mismatch_df)
    """
    if full is None or full.empty:
        return _empty_summary(), _empty_mismatch()

    sum_rows, mism_frames = [], []
    for name, mg in full.groupby("Major_Item", observed=True):
        # Key 단위 일치 판단
        key_labels = {}
        for kval, g in mg.groupby("Rate_Consistency_Key", observed=True):
            lab = _key_consistency(g)
            key_labels[kval] = lab
            if lab == "상이":
                sub = _valid_rows(g)[[
                    "Major_Item", "Rate_Consistency_Key", "Unit", "FWBS", "Description",
                    "Special_Note", "Manhour_UP_Revised", "Labor_UP_Revised",
                    "Material_UP_Revised", "Equipment_UP_Revised", "Tool_Consume_UP_Revised",
                    "Indirect_UP_Revised", "Total_UP_Previous", "Total_UP_Revised"]].copy()
                sub["Rate_Consistency"] = "상이"
                mism_frames.append(sub)

        labels = set(key_labels.values())
        if len(mg) == 1:
            major_label = "단일 항목"
        elif "상이" in labels:
            major_label = "상이"
        elif "일치" in labels:
            major_label = "일치"
        else:
            major_label = "검토 필요"

        # Total Cost = Σ(Qty × Total U/P) — 단가 동일·물량만 변해도 차이가 반영됨
        prev_cost = float((pd.to_numeric(mg["Previous_Qty"], errors="coerce").fillna(0)
                           * pd.to_numeric(mg["Total_UP_Previous"], errors="coerce").fillna(0)).sum())
        rev_cost = float((pd.to_numeric(mg["Revised_Qty"], errors="coerce").fillna(0)
                          * pd.to_numeric(mg["Total_UP_Revised"], errors="coerce").fillna(0)).sum())
        sum_rows.append({
            "Major_Item": name,
            "Unit": ITEM_UNIT.get(name),
            "Matched_FWBS_Count": len(mg),
            "Rate_Consistency": major_label,
            "Previous_Qty": float(mg["Previous_Qty"].sum()),
            "Revised_Qty": float(mg["Revised_Qty"].sum()),
            "Previous_Total_UP": _wavg(mg["Total_UP_Previous"], mg["Previous_Qty"]),
            "Revised_Total_UP": _wavg(mg["Total_UP_Revised"], mg["Revised_Qty"]),
            "Total_Cost_Previous": prev_cost,
            "Total_Cost_Revised": rev_cost,
        })

    s = pd.DataFrame(sum_rows)
    s["Qty_Difference"] = (s["Revised_Qty"] - s["Previous_Qty"]).round(6)
    s["Total_Cost_Difference"] = (s["Total_Cost_Revised"] - s["Total_Cost_Previous"]).round(6)
    s["Major_Item"] = pd.Categorical(s["Major_Item"], categories=ITEM_NAMES, ordered=True)
    s = s[_SUMMARY_COLS].sort_values("Major_Item").reset_index(drop=True)

    rate_mismatch = (pd.concat(mism_frames, ignore_index=True)
                     if mism_frames else _empty_mismatch())
    if not rate_mismatch.empty:
        rate_mismatch["Major_Item"] = pd.Categorical(
            rate_mismatch["Major_Item"], categories=ITEM_NAMES, ordered=True)
        rate_mismatch = (rate_mismatch[_MISMATCH_COLS].sort_values(
            ["Major_Item", "Rate_Consistency_Key", "Total_UP_Revised", "FWBS"])
            .reset_index(drop=True))

    return s, rate_mismatch


_DETAIL_COLS = [
    "Major_Item", "FWBS", "Description", "Unit", "PABS",
    "Previous_Qty", "Revised_Qty", "Qty_Difference",
    "Manhour_UP_Previous", "Manhour_UP_Revised",
    "Labor_UP_Previous", "Labor_UP_Revised",
    "Material_UP_Previous", "Material_UP_Revised",
    "Equipment_UP_Previous", "Equipment_UP_Revised",
    "Tool_Consume_UP_Previous", "Tool_Consume_UP_Revised",
    "Total_UP_Previous", "Total_UP_Revised",
    "Total_Cost_Previous", "Total_Cost_Revised",
]
_SUMMARY_COLS = [
    "Major_Item", "Unit", "Matched_FWBS_Count", "Rate_Consistency",
    "Previous_Qty", "Revised_Qty", "Qty_Difference",
    "Previous_Total_UP", "Revised_Total_UP",
    "Total_Cost_Previous", "Total_Cost_Revised", "Total_Cost_Difference",
]
_MISMATCH_COLS = [
    "Major_Item", "Rate_Consistency_Key", "Unit", "FWBS", "Description", "Special_Note",
    "Manhour_UP_Revised", "Labor_UP_Revised", "Material_UP_Revised",
    "Equipment_UP_Revised", "Tool_Consume_UP_Revised", "Indirect_UP_Revised",
    "Total_UP_Previous", "Total_UP_Revised", "Rate_Consistency",
]


def _empty_detail():
    return pd.DataFrame(columns=_DETAIL_COLS)


def _empty_summary():
    return pd.DataFrame(columns=_SUMMARY_COLS)


def _empty_mismatch():
    return pd.DataFrame(columns=_MISMATCH_COLS)
