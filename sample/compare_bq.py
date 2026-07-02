"""
Previous_BQ.xlsx 와 Revised_BQ.xlsx 를 FWBS + PABS 기준으로 비교한다.

핵심:
- 두 파일 모두 bq_to_long(drop_zero=False) 로 변환한다.
  (0 -> 숫자, 숫자 -> 0 변경도 잡아야 하므로 0 행을 버리지 않는다.)
- 비교 Key = FWBS + PABS

주의 (원본 시트 특성):
  BOQ 시트 하단에는 PABS 수량표가 아니라 'CCS/Area 원가 요약표'가 같이 들어있어,
  bq_to_long 이 끝까지 읽으면서 요약표 행까지 긁어온다. 이로 인해
   (1) C112, C113 ... 같은 상위 코드가 메인 표와 하단 요약표에 중복 등장하고
   (2) '30~513', 'FWBS C11 ... Summary' 같은 잡음 FWBS 가 섞인다.
  -> 정규식(^C+숫자)으로 정상 FWBS 만 남기고,
     (FWBS, PABS) 중복은 먼저 나오는 메인 표 행을 채택(keep='first')해 1:1 키를 만든다.
"""

import re
from pathlib import Path

import numpy as np
import pandas as pd

from bq_to_long import bq_to_long

HERE = Path(__file__).parent
VALID_FWBS = re.compile(r"^C\d")   # 정상 FWBS: C + 숫자 (C11, C1122, C11231Aa ...)
QTY_TOL = 6                        # 수량 비교/차이 반올림 자릿수 (부동소수 잡음 방지)


def load_clean(path):
    """bq_to_long(drop_zero=False) 후 잡음 제거 + (FWBS,PABS) 키 1:1 정리."""
    df = bq_to_long(path, drop_zero=False)
    df = df[df["FWBS"].str.match(VALID_FWBS)].copy()        # 하단 요약표 잡음 제거
    df = df.drop_duplicates(["FWBS", "PABS"], keep="first")  # 메인 표 행 채택
    return df


def compare_bq(prev_path, rev_path):
    prev = load_clean(prev_path)
    rev = load_clean(rev_path)

    keys = ["FWBS", "PABS"]
    # 메타(Description/Unit/U_Code)는 비교 대상이 아니라 참고용 → 한 묶음으로 합친다.
    meta_cols = ["Description", "Unit", "U_Code"]

    prev_q = prev[keys + ["Quantity"]].rename(columns={"Quantity": "Previous_Qty"})
    rev_q = rev[keys + ["Quantity"]].rename(columns={"Quantity": "Revised_Qty"})

    merged = prev_q.merge(rev_q, on=keys, how="outer")

    # 메타데이터: Revised 우선, 없으면 Previous 로 보강 (Deleted 행은 Previous 만 존재)
    meta = (
        rev[keys + meta_cols]
        .set_index(keys)
        .combine_first(prev[keys + meta_cols].set_index(keys))
        .reset_index()
    )
    merged = merged.merge(meta, on=keys, how="left")

    p = merged["Previous_Qty"]
    r = merged["Revised_Qty"]
    has_p = p.notna()
    has_r = r.notna()

    # Change_Type
    same_qty = np.isclose(p.fillna(0), r.fillna(0), rtol=0, atol=10 ** -QTY_TOL)
    conditions = [
        has_p & ~has_r,            # Previous 에만 존재
        ~has_p & has_r,            # Revised 에만 존재
        has_p & has_r & ~same_qty,  # 둘 다 있고 값이 다름
    ]
    choices = ["Deleted", "Added", "Changed"]
    merged["Change_Type"] = np.select(conditions, choices, default="No Change")

    # Difference = Revised - Previous (한쪽만 있으면 없는 쪽을 0으로 간주)
    merged["Difference"] = (r.fillna(0) - p.fillna(0)).round(QTY_TOL)

    # Difference_Percent = Difference / Previous_Qty
    #   Previous_Qty 가 0 또는 빈 값이면 빈 값 처리
    denom = p.where(has_p & (p != 0))            # 0/NaN -> NaN
    merged["Difference_Percent"] = merged["Difference"] / denom

    cols = [
        "FWBS", "Description", "Unit", "PABS", "U_Code",
        "Previous_Qty", "Revised_Qty", "Difference",
        "Difference_Percent", "Change_Type",
    ]
    return merged[cols].sort_values(["FWBS", "PABS"]).reset_index(drop=True)


def main():
    result = compare_bq(HERE / "Previous_BQ.xlsx", HERE / "Revised_BQ.xlsx")

    full_csv = HERE / "BQ_comparison_result.csv"
    result.to_csv(full_csv, index=False, encoding="utf-8-sig")

    counts = result["Change_Type"].value_counts()
    print("=== BQ Comparison (Key = FWBS + PABS) ===")
    for t in ["Added", "Deleted", "Changed", "No Change"]:
        print(f"  {t:<10}: {int(counts.get(t, 0)):,}")
    print(f"  {'TOTAL':<10}: {len(result):,}")

    changes = result[result["Change_Type"] != "No Change"]
    changes_csv = HERE / "BQ_changes_only.csv"
    changes.to_csv(changes_csv, index=False, encoding="utf-8-sig")

    print(f"\nSaved full result   -> {full_csv}")
    print(f"Saved changes only  -> {changes_csv}  ({len(changes):,} rows)")


if __name__ == "__main__":
    main()
