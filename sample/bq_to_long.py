"""
BQ(Bill of Quantities) 와이드 시트를 Long Table로 변환.

원본 구조 (Previous_BQ.xlsx 기준):
    - FWBS CODE  : B열 (계층형 코드, 예: C11 / C112 / C1121 ...)
    - Description: C~H열 (들여쓰기 계층 — 행마다 한 칸만 채워짐)
    - Special Note: J열
    - UNIT       : K열
    - PABS CODE  : 'PABS CODE' 라벨이 있는 행(기본 14행)의 오른쪽, A01~A20
    - U-Code     : PABS CODE 바로 아랫행의 area 코드 (U11100 ...)
    - Quantity   : FWBS(행) × PABS(열) 교차 영역

결과: FWBS / Description / Unit / PABS / U_Code / Quantity 형태의 Long Table
"""

from pathlib import Path

import pandas as pd


def bq_to_long(
    path,
    sheet=0,
    pabs_label="PABS CODE",
    desc_cols=("C", "D", "E", "F", "G", "H"),
    fwbs_col="B",
    unit_col="K",
    note_col="J",
    drop_zero=True,
):
    """BQ 와이드 시트를 Long Table(DataFrame)로 변환한다.

    Parameters
    ----------
    path : str | Path
        엑셀 파일 경로.
    sheet : int | str
        시트 인덱스 또는 이름.
    pabs_label : str
        PABS CODE 행을 찾기 위한 라벨 (해당 셀이 있는 행을 헤더로 인식).
    desc_cols : tuple[str]
        Description 계층이 들어있는 열들. 행마다 채워진 칸을 골라 합친다.
    fwbs_col, unit_col, note_col : str
        각 항목의 엑셀 열 문자.
    drop_zero : bool
        True면 Quantity가 0/NaN인 행을 제거한다.

    Returns
    -------
    pandas.DataFrame
        columns = [FWBS, Description, Unit, Special_Note, PABS, U_Code, Quantity]
    """
    # 헤더 없이 통째로 읽어 행/열 인덱스를 직접 다룬다.
    raw = pd.read_excel(path, sheet_name=sheet, header=None, dtype=object)

    col = _col_letter_to_idx  # 'B' -> 1

    # 1) PABS CODE 라벨이 있는 행 찾기
    pabs_row = _find_label_row(raw, pabs_label)
    ucode_row = pabs_row + 1          # PABS 바로 아랫행 = U-Code(area)
    data_start = pabs_row + 2         # 데이터 시작행

    # 2) PABS 코드가 들어있는 열들 수집 (라벨 셀 오른쪽에서 값이 있는 칸)
    pabs_label_col = raw.iloc[pabs_row].eq(pabs_label).idxmax()
    pabs_cols = []
    for c in range(pabs_label_col + 1, raw.shape[1]):
        code = raw.iat[pabs_row, c]
        if pd.isna(code) or str(code).strip() == "":
            continue
        pabs_cols.append(c)
    pabs_map = {c: str(raw.iat[pabs_row, c]).strip() for c in pabs_cols}
    ucode_map = {c: _clean(raw.iat[ucode_row, c]) for c in pabs_cols}

    fwbs_i = col(fwbs_col)
    unit_i = col(unit_col)
    note_i = col(note_col)
    desc_i = [col(c) for c in desc_cols]

    records = []
    for r in range(data_start, raw.shape[0]):
        fwbs = _clean(raw.iat[r, fwbs_i])
        if not fwbs:                  # FWBS 코드 없는 행은 건너뜀
            continue

        # Description: 계층 열 중 채워진 칸을 사용 (가장 구체적인 레벨)
        desc = ""
        for di in desc_i:
            v = _clean(raw.iat[r, di])
            if v:
                desc = v
        unit = _clean(raw.iat[r, unit_i])
        note = _clean(raw.iat[r, note_i])

        for c in pabs_cols:
            qty = pd.to_numeric(raw.iat[r, c], errors="coerce")  # 'INPUT' 등 텍스트 -> NaN
            if pd.isna(qty):
                continue
            if drop_zero and qty == 0:
                continue
            records.append(
                {
                    "FWBS": fwbs,
                    "Description": desc,
                    "Unit": unit,
                    "Special_Note": note,
                    "PABS": pabs_map[c],
                    "U_Code": ucode_map[c],
                    "Quantity": qty,
                }
            )

    return pd.DataFrame.from_records(
        records,
        columns=[
            "FWBS",
            "Description",
            "Unit",
            "Special_Note",
            "PABS",
            "U_Code",
            "Quantity",
        ],
    )


def _find_label_row(df, label):
    """label 문자열이 들어있는 첫 행 인덱스를 반환."""
    mask = df.apply(lambda s: s.astype(str).str.strip().eq(label).any(), axis=1)
    if not mask.any():
        raise ValueError(f"'{label}' 라벨을 찾을 수 없습니다.")
    return int(mask.idxmax())


def _col_letter_to_idx(letter):
    """'A' -> 0, 'B' -> 1, 'AA' -> 26 ..."""
    idx = 0
    for ch in letter.upper():
        idx = idx * 26 + (ord(ch) - ord("A") + 1)
    return idx - 1


def _clean(v):
    """NaN/공백 정리 후 문자열 반환 (없으면 빈 문자열)."""
    if pd.isna(v):
        return ""
    return str(v).strip()


if __name__ == "__main__":
    src = Path(__file__).with_name("Previous_BQ.xlsx")
    long_df = bq_to_long(src)
    print(f"Long Table rows: {len(long_df):,}")
    print(long_df.head(20).to_string(index=False))
    out = Path(__file__).with_name("Previous_BQ_long.csv")
    long_df.to_csv(out, index=False, encoding="utf-8-sig")
    print(f"\nSaved -> {out}")
