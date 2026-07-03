import html
import io
import sys
from datetime import datetime
from pathlib import Path

import altair as alt
import pandas as pd
import streamlit as st

# 비교 모듈은 sample/ 폴더에 있으므로 import 경로에 추가
sys.path.insert(0, str(Path(__file__).parent / "sample"))
from compare_bq import compare_bq                       # noqa: E402
from compare_rate import compare_rate, ITEM_ORDER       # noqa: E402
# 집계/요약 헬퍼만 사용 (Excel 리포트 생성 기능은 화면 Report/Copy 로 대체)
from build_report import (                               # noqa: E402
    n_changes, revision_type, revision_comment, changed_items,
)
from major_items import build_major_item_view           # noqa: E402  (Major Item 조회)

st.set_page_config(page_title="Project Delta", layout="wide")

# ---- 대시보드용 CSS (Report/Copy 탭 카드·결론박스 등) -------------------------
st.markdown(
    """
    <style>
      .pd-card {
        background: #ffffff;
        border: 1px solid #ececf1;
        border-radius: 16px;
        padding: 18px 20px;
        box-shadow: 0 1px 4px rgba(16, 24, 40, 0.06);
        height: 100%;
      }
      .pd-card .pd-label {
        font-size: 0.78rem;
        font-weight: 500;
        color: #8a90a2;
        letter-spacing: .02em;
        margin-bottom: 8px;
        text-transform: uppercase;
      }
      .pd-card .pd-value {
        font-size: 1.7rem;
        font-weight: 700;
        color: #1a1c23;
        line-height: 1.25;
        word-break: keep-all;
      }
      .pd-card .pd-sub { font-size: .8rem; color: #9aa0b0; margin-top: 4px; }
      .pd-conclusion {
        background: #f4f6ff;
        border: 1px solid #e3e8ff;
        border-left: 5px solid #5b6cff;
        border-radius: 12px;
        padding: 16px 20px;
        color: #2a2f45;
        font-size: 1.02rem;
        line-height: 1.5;
      }

      /* ---- 랜딩(첫 화면) 전용 ------------------------------------------ */
      .pd-topbar {
        display: flex; align-items: center; justify-content: space-between;
        padding: 6px 4px 0 4px;
      }
      .pd-topbar .pd-brand {
        font-size: 1.05rem; font-weight: 800; color: #16336b; letter-spacing: .01em;
      }
      .pd-topbar .pd-author {
        font-size: .9rem; font-weight: 600; color: #64748b;
      }
      .pd-hero {
        text-align: center; padding: 58px 0 6px 0; position: relative; z-index: 1;
      }
      .pd-hero-label {
        font-size: .8rem; letter-spacing: .34em; color: #2563eb; font-weight: 700;
      }
      .pd-hero-title {
        font-size: 4rem; font-weight: 800; line-height: 1.08; margin-top: 16px;
        background: linear-gradient(90deg, #16336b 0%, #2563eb 60%, #3b82f6 100%);
        -webkit-background-clip: text; background-clip: text; color: transparent;
      }
      .pd-hero-subtitle {
        font-size: 1.15rem; color: #1e293b; font-weight: 600;
        margin: 22px auto 0 auto; line-height: 1.6; max-width: 720px;
      }
      .pd-hero-desc {
        font-size: .97rem; color: #64748b;
        margin: 14px auto 0 auto; line-height: 1.7; max-width: 680px;
      }
      .pd-chips {
        display: flex; flex-wrap: wrap; gap: 12px; justify-content: center;
        margin-top: 28px;
      }
      .pd-chip {
        display: inline-flex; align-items: center; gap: 8px;
        background: rgba(255, 255, 255, .78); border: 1px solid #e2e8f2;
        border-radius: 999px; padding: 9px 18px;
        font-size: .9rem; font-weight: 600; color: #334155;
        box-shadow: 0 1px 3px rgba(16, 24, 40, .05);
      }
      .pd-chip .pd-cico { color: #2563eb; font-size: 1rem; }
      .pd-note {
        text-align: center; font-size: .8rem; color: #94a3b8;
        margin-top: 8px; line-height: 1.6;
      }
    </style>
    """,
    unsafe_allow_html=True,
)

# ---- 시작(랜딩) 화면 ----------------------------------------------------------
if "pd_started" not in st.session_state:
    st.session_state.pd_started = False

if not st.session_state.pd_started:
    # ---- 랜딩 전용 배경 / 버튼 스타일 (첫 화면에서만 주입 → 메인 앱은 기본 흰 배경) --
    st.markdown(
        """
        <style>
          .stApp {
            background:
              radial-gradient(1100px 620px at 10% 6%, rgba(191,219,254,.45), transparent 55%),
              radial-gradient(1000px 680px at 92% 10%, rgba(199,210,254,.40), transparent 55%),
              radial-gradient(900px 600px at 82% 96%, rgba(186,230,253,.32), transparent 55%),
              linear-gradient(180deg, #fbfdff 0%, #eef3fb 100%);
          }
          header[data-testid="stHeader"] { background: transparent; }
          .stApp::before, .stApp::after {
            content: ""; position: absolute; z-index: 0; pointer-events: none;
          }
          .stApp::before {
            top: -130px; right: -70px; width: 420px; height: 420px;
            background: linear-gradient(135deg, rgba(147,197,253,.34), rgba(196,181,253,.22));
            clip-path: polygon(50% 0, 100% 38%, 82% 100%, 18% 100%, 0 38%);
          }
          .stApp::after {
            bottom: -150px; left: -110px; width: 360px; height: 360px;
            background: linear-gradient(135deg, rgba(191,219,254,.30), rgba(165,243,252,.20));
            clip-path: polygon(25% 0, 100% 25%, 75% 100%, 0 75%);
          }
          .block-container { position: relative; z-index: 1; }
          .stButton > button[kind="primary"] {
            background: linear-gradient(135deg, #2563eb 0%, #3b82f6 100%);
            border: none; color: #fff; font-weight: 700; border-radius: 12px;
            box-shadow: 0 8px 18px rgba(37, 99, 235, .28);
          }
          .stButton > button[kind="primary"]:hover { filter: brightness(1.06); }
          .stButton > button[kind="secondary"] {
            background: #ffffff; border: 1px solid #cbd5e1; color: #1e293b;
            font-weight: 600; border-radius: 12px;
          }
          .stButton > button[kind="secondary"]:hover { border-color: #2563eb; color: #2563eb; }
        </style>
        """,
        unsafe_allow_html=True,
    )

    # ---- Top header ----------------------------------------------------------
    st.markdown(
        '<div class="pd-topbar"><span class="pd-brand">Project Delta</span>'
        '<span class="pd-author">By. Soyoung Katie Park</span></div>',
        unsafe_allow_html=True,
    )

    # ---- Hero ----------------------------------------------------------------
    st.markdown(
        """
        <div class="pd-hero">
          <div class="pd-hero-label">BQ REVISION COMPARISON</div>
          <div class="pd-hero-title">Project Delta</div>
          <div class="pd-hero-subtitle">
            Previous / Revised Civil BQ를 자동 비교해 물량·단가·금액 영향과<br>
            Rate Consistency 이슈를 한 화면에서 검토하는 Cost Review Dashboard
          </div>
          <div class="pd-hero-desc">
            복잡한 Civil BQ Revision 검토를 자동화해 반복 비교 시간을 줄이고,
            주요 변경사항과 단가 불일치 검토 포인트를 빠르게 확인합니다.
          </div>
          <div class="pd-chips">
            <span class="pd-chip"><span class="pd-cico">⚖</span> Quantity · Rate · Cost 자동 비교</span>
            <span class="pd-chip"><span class="pd-cico">🔍</span> Rate Consistency 이슈 식별</span>
            <span class="pd-chip"><span class="pd-cico">📄</span> Report-ready Summary 즉시 활용</span>
          </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    # ---- CTA 버튼 ------------------------------------------------------------
    st.write("")
    bc = st.columns([1.4, 1, 1, 1.4])
    with bc[1]:
        if st.button("샘플 데이터로 시연하기", type="primary", use_container_width=True):
            st.session_state.pd_started = True
            st.session_state.use_demo = True
            st.rerun()
    with bc[2]:
        if st.button("BQ 파일 업로드로 시작하기", use_container_width=True):
            st.session_state.pd_started = True
            st.session_state.use_demo = False
            st.rerun()

    # ---- 하단 note -----------------------------------------------------------
    st.write("")
    st.write("")
    st.markdown(
        '<div class="pd-note">Local-first prototype&nbsp;·&nbsp;'
        'Company BQ files are excluded from Git tracking&nbsp;·&nbsp;'
        'Built with Python, pandas, openpyxl, Streamlit</div>',
        unsafe_allow_html=True,
    )
    st.stop()

st.title("Project Delta")
st.caption("BQ Revision Comparison Tool · V1.1")

with st.sidebar:
    st.markdown("### 📊 Project Delta")
    st.caption("BQ Revision Comparison Tool")
    st.divider()

    st.markdown("#### 데이터 입력")
    use_demo = st.checkbox(
        "샘플 데이터로 체험하기", key="use_demo",
        help="회사 데이터를 익명화한 샘플 BQ로 모든 탭 기능을 바로 확인할 수 있습니다.",
    )
    if use_demo:
        st.caption("업로드 없이 익명화된 데모 BQ로 실행합니다.")

    previous_file = st.file_uploader("Previous BQ", type=["xlsx"], disabled=use_demo)
    revised_file = st.file_uploader("Revised BQ", type=["xlsx"], disabled=use_demo)

    st.divider()
    st.caption("v1.1 · Built with Streamlit")


@st.cache_data(show_spinner="Quantity 비교 중...")
def run_quantity(prev_bytes, rev_bytes):
    return compare_bq(io.BytesIO(prev_bytes), io.BytesIO(rev_bytes))


@st.cache_data(show_spinner="Unit Rate / Cost 비교 중...")
def run_rate(prev_bytes, rev_bytes):
    return compare_rate(io.BytesIO(prev_bytes), io.BytesIO(rev_bytes))


@st.cache_data(show_spinner="Major Item 집계 중...")
def run_major(prev_bytes, rev_bytes, include_dummy=False):
    return build_major_item_view(prev_bytes, rev_bytes, include_dummy=include_dummy)


def _contains(series, query):
    """대소문자 무시 부분일치 마스크 (빈 검색어 → 전체 True)."""
    if not query or not query.strip():
        return pd.Series(True, index=series.index)
    return series.astype(str).str.contains(query.strip(), case=False, na=False)


def _download_csv(label, df, fname, key):
    st.download_button(
        label, data=df.to_csv(index=False).encode("utf-8-sig"),
        file_name=fname, mime="text/csv", use_container_width=True, key=key,
    )


def _change_rate_display(prev, rev, new_label="New Qty"):
    """증감률 표시문자열. blank/NaN 은 0 으로 취급.

    1) prev==0, rev>0 → new_label ("New Qty" / "New Value")
    2) prev>0,  rev==0 → "-100%"
    3) prev>0,  rev>0  → "±%.1f%%"
    4) prev==0, rev==0 → "0%"
    """
    p = pd.to_numeric(prev, errors="coerce").fillna(0)
    r = pd.to_numeric(rev, errors="coerce").fillna(0)
    out = pd.Series("0%", index=p.index, dtype=object)
    m3 = (p > 0) & (r > 0)
    out[m3] = ((r[m3] - p[m3]) / p[m3] * 100).map(lambda v: f"{v:.1f}%")
    out[(p == 0) & (r > 0)] = new_label
    out[(p > 0) & (r == 0)] = "-100%"
    return out


# Quantity 표시용 컬럼 순서 (내부 컬럼은 그대로 두고 표시 DataFrame 에서만 사용)
_QTY_DISPLAY_COLS = [
    "FWBS", "Description", "Unit", "PABS_Key", "PABS_Name",
    "Previous_Qty", "Revised_Qty", "Difference", "Change_Rate_Display", "Change_Type",
]
# 화면 컬럼 헤더 라벨
_QTY_COLCFG = {"Change_Rate_Display": st.column_config.TextColumn("Change Rate (증감률)")}


def qty_display(df):
    """compare_bq 결과 → Quantity 화면/CSV 표시용 DataFrame.

    내부 비교 로직/컬럼은 유지하고, 표시용으로만 rename + Change_Rate_Display 추가.
      PABS → PABS_Key, U_Code → PABS_Name, Difference_Percent → Change_Rate_Display
    """
    d = df.copy()
    d["Change_Rate_Display"] = _change_rate_display(d["Previous_Qty"], d["Revised_Qty"])
    d = d.rename(columns={"PABS": "PABS_Key", "U_Code": "PABS_Name"})
    return d[_QTY_DISPLAY_COLS]


# Rate/Cost 표시용 컬럼 순서 (Difference_Percent → Change_Rate_Display)
_RATE_DISPLAY_COLS = [
    "FWBS", "Description", "Unit", "Compare_Item",
    "Previous_Value", "Revised_Value", "Difference", "Change_Rate_Display", "Change_Type",
]
_RATE_COLCFG = {"Change_Rate_Display": st.column_config.TextColumn("Change Rate (증감률)")}


def rate_display(df):
    """compare_rate 결과 → Rate/Cost 화면/CSV 표시용 DataFrame.

    내부 비교 로직/컬럼은 유지하고, 표시용으로만 Change_Rate_Display 추가
    (Previous_Value / Revised_Value 기준, prev==0&rev>0 은 "New Value").
    Difference_Percent 는 표시에서 제외.
    """
    d = df.copy()
    d["Change_Rate_Display"] = _change_rate_display(
        d["Previous_Value"], d["Revised_Value"], new_label="New Value")
    return d[_RATE_DISPLAY_COLS]


# ---- Major Items Summary 표시용 -------------------------------------------
_MAJOR_SUMMARY_COLS = [
    "Major Item", "Unit", "Matched FWBS Count", "Rate Consistency",
    "Previous Qty", "Revised Qty", "Qty Difference",
    "Previous Total U/P", "Revised Total U/P", "Total U/P Change Rate",
    "Total Cost Difference",
]
# 4개 컬럼 그룹 (라벨, 컬럼들, 은은한 배경색)
_SUMMARY_GROUPS = [
    ("기본 정보", ["Major Item", "Unit", "Matched FWBS Count", "Rate Consistency"], "#eef1f4"),
    ("물량", ["Previous Qty", "Revised Qty", "Qty Difference"], "#e9f1fb"),
    ("단가", ["Previous Total U/P", "Revised Total U/P", "Total U/P Change Rate"], "#eaf6ee"),
    ("금액 영향", ["Total Cost Difference"], "#fdf2e1"),
]
_SUMMARY_RIGHT_ALIGN = [
    "Previous Qty", "Revised Qty", "Qty Difference",
    "Previous Total U/P", "Revised Total U/P", "Total U/P Change Rate",
    "Total Cost Difference",
]

# 단가 상이 상세 테이블 (FWBS별 Revised 구성단가)
_MISMATCH_NUM_COLS = [
    "Revised Manhour U/P", "Revised Labor U/P", "Revised Material U/P",
    "Revised Equipment U/P", "Revised Tool & Consume U/P", "Revised Indirect U/P",
    "Revised Total U/P", "Previous Total U/P",
]
_MISMATCH_DISPLAY_COLS = (
    ["Major Item", "Rate Consistency Key", "FWBS", "Description", "Special Note", "Unit"]
    + _MISMATCH_NUM_COLS + ["Total U/P Change Rate"]
)


def _fmt_num(series, dp=2):
    return pd.to_numeric(series, errors="coerce").map(
        lambda v: "" if pd.isna(v) else f"{v:,.{dp}f}")


def _fmt_signed(series, dp=2):
    """천단위 + 부호 표시 (양수 +, 음수 -)."""
    return pd.to_numeric(series, errors="coerce").map(
        lambda v: "" if pd.isna(v) else f"{v:+,.{dp}f}")


def major_summary_display(summary):
    """Major Item Summary 화면/CSV 표시용 (가중평균 단가 + 금액영향 중심)."""
    if summary is None or summary.empty:
        return pd.DataFrame(columns=_MAJOR_SUMMARY_COLS)
    s = summary.reset_index(drop=True)
    d = pd.DataFrame(index=range(len(s)))
    d["Major Item"] = s["Major_Item"].astype(str).values
    d["Unit"] = s["Unit"].values
    d["Matched FWBS Count"] = s["Matched_FWBS_Count"].astype(int).astype(str).values
    d["Rate Consistency"] = s["Rate_Consistency"].values
    d["Previous Qty"] = _fmt_num(s["Previous_Qty"]).values
    d["Revised Qty"] = _fmt_num(s["Revised_Qty"]).values
    d["Qty Difference"] = _fmt_num(s["Qty_Difference"]).values
    d["Previous Total U/P"] = _fmt_num(s["Previous_Total_UP"]).values
    d["Revised Total U/P"] = _fmt_num(s["Revised_Total_UP"]).values
    d["Total U/P Change Rate"] = _change_rate_display(
        s["Previous_Total_UP"], s["Revised_Total_UP"], "New Value").values
    d["Total Cost Difference"] = _fmt_signed(s["Total_Cost_Difference"]).values
    return d[_MAJOR_SUMMARY_COLS]


def summary_group_legend():
    """4개 컬럼 그룹 색상 범례(chip) HTML — chip 사이 간격 확보."""
    chips = "".join(
        f'<span style="padding:3px 12px;border-radius:8px;background:{color};'
        f'border:1px solid #e3e6ea;font-size:0.8rem;color:#333;">{label}</span>'
        for label, _cols, color in _SUMMARY_GROUPS)
    return (
        '<div style="display:flex;flex-wrap:wrap;gap:10px;align-items:center;'
        f'margin:2px 0 10px 0;">{chips}</div>'
    )


def style_major_summary(disp):
    """그룹별 배경색 + Total Cost Difference 부호 색상 Styler."""
    sty = disp.style
    for _label, cols, color in _SUMMARY_GROUPS:
        sty = sty.set_properties(subset=cols, **{"background-color": color})
    sty = sty.set_properties(subset=_SUMMARY_RIGHT_ALIGN, **{"text-align": "right"})

    def _sign_color(v):
        if isinstance(v, str) and v.startswith("-"):
            return "color:#c0392b; font-weight:600"
        if isinstance(v, str) and v.startswith("+") and not v.startswith("+0.00"):
            return "color:#1e7e34; font-weight:600"
        return ""

    # pandas 2.1+ 는 Styler.map, 구버전은 applymap (deprecated) fallback
    if hasattr(sty, "map"):
        return sty.map(_sign_color, subset=["Total Cost Difference"])
    return sty.applymap(_sign_color, subset=["Total Cost Difference"])


def mismatch_display(df):
    """단가 상이 상세 (Activity/Special Note 단위, FWBS별 Revised 구성단가)."""
    if df is None or df.empty:
        return pd.DataFrame(columns=_MISMATCH_DISPLAY_COLS)
    d = df.reset_index(drop=True)
    out = pd.DataFrame(index=range(len(d)))
    out["Major Item"] = d["Major_Item"].astype(str).values
    out["Rate Consistency Key"] = d["Rate_Consistency_Key"].values
    out["FWBS"] = d["FWBS"].values
    out["Description"] = d["Description"].values
    out["Special Note"] = d["Special_Note"].values
    out["Unit"] = d["Unit"].values
    out["Revised Manhour U/P"] = pd.to_numeric(d["Manhour_UP_Revised"], errors="coerce").values
    out["Revised Labor U/P"] = pd.to_numeric(d["Labor_UP_Revised"], errors="coerce").values
    out["Revised Material U/P"] = pd.to_numeric(d["Material_UP_Revised"], errors="coerce").values
    out["Revised Equipment U/P"] = pd.to_numeric(d["Equipment_UP_Revised"], errors="coerce").values
    out["Revised Tool & Consume U/P"] = pd.to_numeric(d["Tool_Consume_UP_Revised"], errors="coerce").values
    out["Revised Indirect U/P"] = pd.to_numeric(d["Indirect_UP_Revised"], errors="coerce").values
    out["Revised Total U/P"] = pd.to_numeric(d["Total_UP_Revised"], errors="coerce").values
    out["Previous Total U/P"] = pd.to_numeric(d["Total_UP_Previous"], errors="coerce").values
    out["Total U/P Change Rate"] = _change_rate_display(
        d["Total_UP_Previous"], d["Total_UP_Revised"], "New Value").values
    out = out.sort_values(
        ["Major Item", "Rate Consistency Key", "Revised Total U/P", "FWBS"]).reset_index(drop=True)
    return out[_MISMATCH_DISPLAY_COLS]


# ---- 입력 소스 결정 (샘플 데모 or 업로드) ------------------------------------
_DEMO_DIR = Path(__file__).parent / "sample"
_DEMO_PREV = _DEMO_DIR / "demo_Previous_BQ.xlsx"
_DEMO_REV = _DEMO_DIR / "demo_Revised_BQ.xlsx"

if use_demo:
    if not (_DEMO_PREV.exists() and _DEMO_REV.exists()):
        st.error("샘플 데이터 파일을 찾을 수 없습니다. (sample/demo_*.xlsx)")
        st.stop()
    prev_bytes, rev_bytes = _DEMO_PREV.read_bytes(), _DEMO_REV.read_bytes()
    prev_name, rev_name = _DEMO_PREV.name, _DEMO_REV.name
    st.sidebar.success("샘플 데이터로 실행 중 (익명화된 데모 BQ)")
elif previous_file and revised_file:
    st.sidebar.success("두 파일 업로드 완료")
    prev_bytes, rev_bytes = previous_file.getvalue(), revised_file.getvalue()
    prev_name, rev_name = previous_file.name, revised_file.name
else:
    st.info("👈 왼쪽 사이드바에서 Previous / Revised BQ 파일을 업로드하거나 "
            "'샘플 데이터로 체험하기'를 켜세요.")
    st.stop()

try:
    qty_result = run_quantity(prev_bytes, rev_bytes)
    rate_result = run_rate(prev_bytes, rev_bytes)
except Exception as exc:  # noqa: BLE001
    st.error(
        "비교 중 오류가 발생했습니다. 업로드한 파일이 올바른 BQ 양식(xlsx)인지 "
        "확인해주세요. ('PABS CODE' 행, FWBS/Quantity/단가 영역이 있어야 합니다.)"
    )
    st.exception(exc)
    st.stop()

q_chg, r_chg = n_changes(qty_result), n_changes(rate_result)
rtype = revision_type(q_chg, r_chg)

# Major Item 집계는 1회만 계산해 Report/Copy · Major Items 탭에서 공유
try:
    major_view = run_major(prev_bytes, rev_bytes)
    major_error = None
except Exception as exc:  # noqa: BLE001
    major_view, major_error = None, exc

generated = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
ts = datetime.now().strftime("%Y%m%d_%H%M")

tab_summary, tab_major, tab_report, tab_qty, tab_rate = st.tabs(
    ["Summary", "Major Items", "Report / Copy", "Quantity Changes", "Rate / Cost Changes"]
)

# ============================== Summary =======================================
with tab_summary:
    c1, c2 = st.columns(2)
    c1.markdown(f"**Previous File Name**\n\n{prev_name}")
    c2.markdown(f"**Revised File Name**\n\n{rev_name}")
    st.caption(f"Generated Time: {generated}")

    m1, m2, m3 = st.columns(3)
    m1.metric("Quantity Changed", f"{q_chg:,}")
    m2.metric("Rate/Cost Changed", f"{r_chg:,}")
    m3.metric("Revision Type", rtype)

    st.divider()
    st.markdown("##### 변경 건수 요약")
    _chg_df = pd.DataFrame(
        {"구분": ["Quantity", "Rate / Cost"], "변경 건수": [q_chg, r_chg]}
    )
    _chg_chart = (
        alt.Chart(_chg_df)
        .mark_bar(cornerRadiusTopLeft=6, cornerRadiusTopRight=6, size=70)
        .encode(
            x=alt.X("구분:N", axis=alt.Axis(labelAngle=0, title=None)),
            y=alt.Y("변경 건수:Q", title=None),
            color=alt.Color(
                "구분:N",
                scale=alt.Scale(domain=["Quantity", "Rate / Cost"],
                                range=["#5b6cff", "#22c1a4"]),
                legend=None),
            tooltip=["구분", "변경 건수"],
        )
        .properties(height=260)
    )
    st.altair_chart(_chg_chart, use_container_width=True)

    comment = revision_comment(qty_result, rate_result)
    (st.info if rtype == "No Revision Detected" else st.success)(
        f"**Revision Type: {rtype}**\n\n{comment}"
    )

# =========================== Quantity Changes =================================
with tab_qty:
    st.caption("**Key:** FWBS + PABS_Key  ·  **Target:** Quantity only")
    f1, f2, f3, f4 = st.columns([2, 2, 2, 1.4])
    q_fwbs = f1.text_input("FWBS 검색", key="q_fwbs")
    q_pabs = f2.text_input("PABS_Key 검색", key="q_pabs")
    q_types = f3.multiselect(
        "Change Type", ["Added", "Deleted", "Changed"],
        default=["Added", "Deleted", "Changed"], key="q_types",
    )
    q_nc = f4.checkbox("No Change 포함", value=False, key="q_nc")

    allowed = set(q_types) | ({"No Change"} if q_nc else set())
    view = qty_result[qty_result["Change_Type"].isin(allowed)]
    view = view[_contains(view["FWBS"], q_fwbs) & _contains(view["PABS"], q_pabs)]

    if view.empty:
        st.info("Quantity 변경사항이 없습니다.")
    else:
        st.caption(f"표시 행: {len(view):,} / 전체 {len(qty_result):,}")
        st.dataframe(qty_display(view), use_container_width=True,
                     hide_index=True, column_config=_QTY_COLCFG)

    st.divider()
    qc = qty_result[qty_result["Change_Type"] != "No Change"]
    cc1, cc2 = st.columns(2)
    with cc1:
        _download_csv("Quantity 전체 CSV", qty_display(qty_result),
                      f"BQ_comparison_result_{ts}.csv", "dl_q_all")
    with cc2:
        _download_csv("Quantity 변경분 CSV", qty_display(qc),
                      f"BQ_changes_only_{ts}.csv", "dl_q_chg")

# ========================= Rate / Cost Changes ================================
with tab_rate:
    st.caption("**Key:** FWBS + Compare_Item  ·  **Target:** Unit Rate / Cost")

    # Compare_Item 별 Changed Count 요약 (행/열 전환: 컬럼=Compare Item, 행=Changed Count)
    chg = rate_result[rate_result["Change_Type"] != "No Change"]
    _counts = (chg.groupby("Compare_Item", observed=False).size()
               .reindex(ITEM_ORDER, fill_value=0))
    summary_t = _counts.to_frame("Changed Count").T   # 1행(Changed Count) × 항목 컬럼
    st.dataframe(summary_t, use_container_width=True)

    # --- Description별 단가(Total U/P) 변동 차트 -------------------------------
    # FWBS가 달라도 Description이 같으면 같은 아이템으로 묶어 Previous/Revised 단가 비교.
    # 같은 Description인데 단가쌍이 다르면 (1),(2)… 로 분리하고 소수 그룹은 확인 경고.
    _tup = rate_result[(rate_result["Compare_Item"] == "Total U/P")
                       & (rate_result["Change_Type"] == "Changed")].copy()
    _tup["_desc"] = _tup["Description"].astype(str).str.strip()
    _n_blank = int((_tup["_desc"] == "").sum())
    _tup = _tup[_tup["_desc"] != ""]

    if not _tup.empty:
        st.markdown("##### Description별 단가(Total U/P) 변동")
        st.caption("FWBS가 달라도 Description이 같으면 같은 아이템으로 묶어 표시합니다. "
                   "같은 Description인데 단가가 다르면 (1), (2)… 로 구분합니다.")

        def _desc_label(d):
            dl = d.lower()
            # Above/Below Water Level 은 Excavation 하위항목 → 표시명에 prefix
            if dl.startswith("above water level") or dl.startswith("below water level"):
                return "Excavation " + d
            return d

        _rows, _warns = [], []
        for _d, _g in _tup.groupby("_desc"):
            _pairs = pd.Series(
                list(zip(_g["Previous_Value"].round(4), _g["Revised_Value"].round(4))),
                index=_g.index)
            _vc = _pairs.value_counts()          # 다수 단가 그룹부터
            _multi = len(_vc) > 1
            for _i, (_pair, _cnt) in enumerate(_vc.items(), 1):
                _label = _desc_label(_d) + (f" ({_i})" if _multi else "")
                _rows.append({"Item": _label, "구분": "Previous",
                              "Total U/P": _pair[0], "FWBS 수": int(_cnt)})
                _rows.append({"Item": _label, "구분": "Revised",
                              "Total U/P": _pair[1], "FWBS 수": int(_cnt)})
                if _multi and _i > 1:            # 소수 단가 그룹 → 단가 확인 필요
                    _codes = _g.loc[_pairs[_pairs == _pair].index, "FWBS"].tolist()
                    _warns.append((_label, _codes))

        _cdf = pd.DataFrame(_rows)
        _n_items = _cdf["Item"].nunique()
        _dchart = (
            alt.Chart(_cdf)
            .mark_bar()
            .encode(
                y=alt.Y("Item:N", title=None,
                        sort=alt.EncodingSortField(field="Total U/P",
                                                   op="max", order="descending")),
                yOffset=alt.YOffset("구분:N", sort=["Previous", "Revised"]),
                x=alt.X("Total U/P:Q", title=None),
                color=alt.Color(
                    "구분:N",
                    scale=alt.Scale(domain=["Previous", "Revised"],
                                    range=["#b9c0d4", "#5b6cff"]),
                    legend=alt.Legend(orient="top", title=None)),
                tooltip=["Item", "구분",
                         alt.Tooltip("Total U/P:Q", format=",.2f"), "FWBS 수"],
            )
            .properties(height=max(240, 32 * _n_items))
        )
        st.altair_chart(_dchart, use_container_width=True)

        for _label, _codes in _warns:
            st.markdown(
                f'<div style="color:#c0392b; font-size:0.9rem; margin:2px 0;">'
                f'⚠️ <b>{html.escape(_label)}</b> · FWBS: '
                f'{html.escape(", ".join(_codes))} — 단가 확인 필요.</div>',
                unsafe_allow_html=True)
        if _n_blank:
            st.caption(f"Description이 비어 있는 {_n_blank}개 변경 행은 차트에서 제외했습니다.")

    f1, f2, f3, f4 = st.columns([2, 2.4, 2, 1.6])
    r_fwbs = f1.text_input("FWBS 검색", key="r_fwbs")
    r_items = f2.multiselect("Compare_Item", ITEM_ORDER, default=ITEM_ORDER, key="r_items")
    r_types = f3.multiselect(
        "Change Type", ["Added", "Deleted", "Changed", "No Change"],
        default=["Added", "Deleted", "Changed"], key="r_types",
    )
    r_sort = f4.checkbox("|Difference| 큰 순", value=True, key="r_sort")

    view = rate_result[
        rate_result["Change_Type"].isin(r_types)
        & rate_result["Compare_Item"].isin(r_items)
    ]
    view = view[_contains(view["FWBS"], r_fwbs)]
    if r_sort and not view.empty:
        view = view.assign(_abs=view["Difference"].abs()) \
            .sort_values("_abs", ascending=False).drop(columns="_abs")

    if view.empty:
        st.info("Rate/Cost 변경사항이 없습니다.")
    else:
        st.caption(f"표시 행: {len(view):,} / 전체 {len(rate_result):,}")
        st.dataframe(rate_display(view), use_container_width=True,
                     hide_index=True, column_config=_RATE_COLCFG)

    st.divider()
    rc = rate_result[rate_result["Change_Type"] != "No Change"]
    cc1, cc2 = st.columns(2)
    with cc1:
        _download_csv("Rate/Cost 전체 CSV", rate_display(rate_result),
                      f"BQ_rate_comparison_result_{ts}.csv", "dl_r_all")
    with cc2:
        _download_csv("Rate/Cost 변경분 CSV", rate_display(rc),
                      f"BQ_rate_changes_only_{ts}.csv", "dl_r_chg")

# ============================== Report / Copy =================================
def _card(label, value, sub=""):
    sub_html = f'<div class="pd-sub">{html.escape(sub)}</div>' if sub else ""
    return (
        f'<div class="pd-card"><div class="pd-label">{html.escape(label)}</div>'
        f'<div class="pd-value">{html.escape(str(value))}</div>{sub_html}</div>'
    )


with tab_report:
    st.subheader("Executive Summary")
    st.caption("화면에서 바로 확인하고 복사해 메일 / Notion / 보고자료에 붙여넣으세요.")

    # 변경 항목별 집계 (Changed_Count + Difference 합계) — 비교 결과만 가공
    rchg = rate_result[rate_result["Change_Type"] != "No Change"]
    items_tbl = (
        rchg.groupby("Compare_Item", observed=True)
        .agg(Changed_Count=("Change_Type", "size"),
             Difference_Sum=("Difference", "sum"))
        .reset_index()
        .sort_values("Changed_Count", ascending=False)
    )
    if not items_tbl.empty:
        top_item = str(items_tbl.iloc[0]["Compare_Item"])
        top_sub = f"{int(items_tbl.iloc[0]['Changed_Count']):,}건 변경"
    else:
        top_item, top_sub = "—", "변경 없음"

    # --- Executive Summary 카드 4개 ---
    cols = st.columns(4)
    cols[0].markdown(_card("Revision Type", rtype), unsafe_allow_html=True)
    cols[1].markdown(_card("Quantity Change", f"{q_chg:,}", "건"), unsafe_allow_html=True)
    cols[2].markdown(_card("Rate/Cost Change", f"{r_chg:,}", "건"), unsafe_allow_html=True)
    cols[3].markdown(_card("Top Changed Item", top_item, top_sub), unsafe_allow_html=True)

    # --- 한 줄 결론 (커스텀 박스) ---
    st.write("")
    comment = revision_comment(qty_result, rate_result)
    st.markdown(f'<div class="pd-conclusion">{html.escape(comment)}</div>',
                unsafe_allow_html=True)

    st.divider()

    # --- Rate/Cost Change Summary (행=Changed Count, 열=Compare Item) ---
    st.subheader("Rate/Cost Change Summary")
    _rc_counts = (rchg.groupby("Compare_Item", observed=False).size()
                  .reindex(ITEM_ORDER, fill_value=0))
    rc_summary_t = _rc_counts.to_frame("Changed Count").T   # 1행 × 항목 컬럼
    if items_tbl.empty:
        st.info("변경된 Rate/Cost 항목이 없습니다.")
    else:
        st.dataframe(rc_summary_t, use_container_width=True)

    st.divider()

    # --- Major Item 요약 (req 7) ---
    st.subheader("Major Item 요약")
    st.caption(
        "Major Item 기준으로 Concrete, Rebar, Excavation 등 주요 항목의 "
        "Quantity 및 Unit Rate를 별도 집계했습니다. (상세는 Major Items 탭)"
    )
    major_report_lines = [
        "",
        "[Major Item 요약]",
        "Major Item 기준으로 Concrete, Rebar, Excavation 등 주요 항목의 "
        "Quantity 및 Unit Rate를 별도 집계했습니다.",
    ]
    msum = major_view["summary"] if major_view is not None else None
    if msum is not None and not msum.empty:
        msum = msum.assign(
            Total_Cost_Difference=(msum["Total_Cost_Revised"]
                                   - msum["Total_Cost_Previous"]).round(6))
        mtop = (msum.assign(_a=msum["Total_Cost_Difference"].abs())
                .query("_a > 0").sort_values("_a", ascending=False).head(5))
        if mtop.empty:
            st.caption("Major Item 기준 Cost 변경은 없습니다.")
            major_report_lines.append("- Major Item 기준 Cost 변경 없음")
        else:
            st.dataframe(
                mtop[["Major_Item", "Unit", "Total_Cost_Revised", "Total_Cost_Difference"]],
                use_container_width=True, hide_index=True,
                column_config={
                    "Total_Cost_Revised": st.column_config.NumberColumn("Revised Total Cost", format="%.0f"),
                    "Total_Cost_Difference": st.column_config.NumberColumn("Cost Difference", format="%.0f"),
                },
            )
            major_report_lines.append("[Major Item Top Changes (Cost)]")
            for row in mtop.itertuples(index=False):
                major_report_lines.append(
                    f"- {row.Major_Item}: Cost Δ {row.Total_Cost_Difference:,.0f}")
    else:
        st.info("매칭된 Major Item이 없습니다.")
        major_report_lines.append("- 매칭된 Major Item 없음")

    st.divider()

    # --- 복사용 텍스트/표 (expander 안에) ---
    top5 = changed_items(rate_result, top=5)
    lines = [
        "[Project Delta 분석 결과]",
        f"Previous File : {prev_name}",
        f"Revised File  : {rev_name}",
        f"Generated     : {generated}",
        "",
        f"Quantity 변경 건수   : {q_chg:,}",
        f"Rate/Cost 변경 건수  : {r_chg:,}",
        f"Revision Type        : {rtype}",
        "",
        "[주요 변경 항목 Top 5]",
    ]
    if top5 is not None and not top5.empty:
        for i, row in enumerate(top5.itertuples(index=False), 1):
            lines.append(f"{i}. {row.Compare_Item} : {int(row.Changed_Count):,}건")
    else:
        lines.append("(변경 항목 없음)")
    lines += ["", "[해석]", comment] + major_report_lines
    report_text = "\n".join(lines)

    if items_tbl.empty:
        tsv = "Changed Count\n(변경 항목 없음)"
    else:
        tsv = rc_summary_t.to_csv(sep="\t")   # 행=Changed Count, 열=Compare Item

    with st.expander("Copy-ready report text"):
        st.code(report_text, language="text")
    with st.expander("Copy table as TSV"):
        st.code(tsv, language="text")

# ============================== Major Items ===================================
with tab_major:
    st.subheader("Major Items 조회")
    st.caption("지정 Major Item 리스트를 BQ Description 으로 매칭해 "
               "Quantity · Unit Rate / Cost 를 별도 집계합니다.")

    include_dummy = st.checkbox(
        "0값/dummy 행 포함", value=False, key="mi_dummy",
        help="수량·단가·변경이 모두 0인 all-zero dummy 행까지 표시합니다.")
    try:
        mview = run_major(prev_bytes, rev_bytes, include_dummy)
    except Exception as exc:  # noqa: BLE001
        st.error("Major Item 집계 중 오류가 발생했습니다.")
        st.exception(exc)
        mview = None

    if mview is None:
        st.info("집계 결과를 표시할 수 없습니다.")
    else:
        m_summary = mview["summary"]
        m_detail = mview["detail"]

        if m_summary.empty:
            st.warning("매칭된 Major Item이 없습니다. 키워드 또는 BQ Description 을 확인해주세요.")
        else:
            # --- 매칭 진단 (개발/검토용) — 기본 접힘 expander ---
            with st.expander("매칭 진단 보기", expanded=False):
                st.caption(
                    "Description 기반 자동 매칭 결과를 검토하기 위한 진단 정보입니다. "
                    "Tie Flag는 여러 Major Item 후보가 동일/유사 점수로 매칭되어 "
                    "수동 확인이 필요한 항목을 의미합니다."
                )
                amb = mview["ambiguous"]
                has_diag = False
                if mview["unmatched"]:
                    st.warning("등록했지만 매칭되지 않은 Major Item: "
                               + ", ".join(mview["unmatched"]))
                    has_diag = True
                if mview["broad"]:
                    st.warning("매칭 검토 권장: "
                               + ", ".join(f"{n} ({c})" for n, c in mview["broad"]))
                    has_diag = True
                if not amb.empty:
                    st.markdown(f"**중복 후보 매칭 항목 ({len(amb):,}건)**")
                    st.dataframe(amb.rename(columns={"Tie": "Tie Flag"}),
                                 use_container_width=True, hide_index=True)
                    has_diag = True
                if not has_diag:
                    st.caption("진단 사항이 없습니다.")

            m_mismatch = mview.get("rate_mismatch")

            # --- 필터 ---
            unit_map = dict(zip(m_summary["Major_Item"].astype(str), m_summary["Unit"]))
            opt_items = list(m_summary["Major_Item"].astype(str))
            opt_units = sorted(set(unit_map.values()))
            fc = st.columns([2.6, 1.3, 1.7, 1.7])
            sel_items = fc[0].multiselect("Major Item", opt_items, default=opt_items, key="mi_items")
            sel_units = fc[1].multiselect("Unit", opt_units, default=opt_units, key="mi_units")
            q_only = fc[2].checkbox("Quantity changed only", value=False, key="mi_qonly")
            r_only = fc[3].checkbox("Rate/Cost changed only", value=False, key="mi_ronly")
            sc = st.columns(2)
            f_fwbs = sc[0].text_input("FWBS 검색", key="mi_fwbs")
            f_desc = sc[1].text_input("Description 검색", key="mi_desc")

            # Unit 필터는 Major Item(=고정 Unit) 기준으로 적용 → 상세 Unit 표기 차이 영향 없음
            allowed = [it for it in sel_items if unit_map.get(it) in sel_units]

            # --- Summary Table (Summary 행만, 변화율 중심) ---
            s_view = m_summary[m_summary["Major_Item"].astype(str).isin(allowed)]
            if q_only:
                s_view = s_view[s_view["Qty_Difference"].round(6) != 0]
            if r_only:
                s_view = s_view[(s_view["Total_Cost_Revised"]
                                 - s_view["Total_Cost_Previous"]).round(6) != 0]
            shown_items = list(s_view["Major_Item"].astype(str))

            st.markdown("**Summary Table**")
            if s_view.empty:
                st.info("조건에 해당하는 Major Item이 없습니다.")
            else:
                summary_disp = major_summary_display(s_view)
                st.markdown(summary_group_legend(), unsafe_allow_html=True)
                st.dataframe(style_major_summary(summary_disp),
                             use_container_width=True, hide_index=True)

                _mi = s_view.assign(
                    _diff=(s_view["Total_Cost_Revised"]
                           - s_view["Total_Cost_Previous"]).round(6))
                _mi = _mi[_mi["_diff"] != 0]
                if not _mi.empty:
                    _mi_df = pd.DataFrame({
                        "Major Item": _mi["Major_Item"].astype(str).values,
                        "Cost Difference": _mi["_diff"].values,
                    })
                    _mi_chart = (
                        alt.Chart(_mi_df)
                        .mark_bar()
                        .encode(
                            x=alt.X("Cost Difference:Q", title=None),
                            y=alt.Y("Major Item:N", sort="-x", title=None),
                            color=alt.condition(
                                "datum['Cost Difference'] > 0",
                                alt.value("#1e7e34"), alt.value("#c0392b")),
                            tooltip=["Major Item", "Cost Difference"],
                        )
                        .properties(height=280)
                    )
                    st.caption("Major Item별 금액 영향 (Cost Difference)")
                    st.altair_chart(_mi_chart, use_container_width=True)

                with st.expander("Copy Major Item Summary as TSV"):
                    st.code(summary_disp.to_csv(sep="\t", index=False), language="text")

            # --- Rate Mismatch Details (단가 상이 상세) ---
            st.markdown("**Rate Mismatch Details (단가 상이 상세)**")
            mm_view = None
            if m_mismatch is not None and not m_mismatch.empty:
                mm_view = m_mismatch[m_mismatch["Major_Item"].astype(str).isin(shown_items)]
                mm_view = mm_view[_contains(mm_view["FWBS"], f_fwbs)
                                  & _contains(mm_view["Description"], f_desc)]
            if mm_view is None or mm_view.empty:
                st.caption("단가 상이(Rate Consistency=상이) 항목이 없습니다.")
            else:
                mm_disp = mismatch_display(mm_view)
                mm_colcfg = {c: st.column_config.NumberColumn(format="%.2f")
                             for c in _MISMATCH_NUM_COLS}
                for item in shown_items:                       # Major Item 순서 유지
                    sub = mm_disp[mm_disp["Major Item"] == item]
                    if sub.empty:
                        continue
                    with st.expander(f"단가 상이 상세 - {item} ({len(sub):,})"):
                        st.dataframe(sub, use_container_width=True, hide_index=True,
                                     column_config=mm_colcfg)
                        st.code(sub.to_csv(sep="\t", index=False), language="text")

            # --- Detail Table (기존 구성 유지) ---
            d_view = m_detail[m_detail["Major_Item"].astype(str).isin(allowed)]
            if q_only:
                d_view = d_view[d_view["Qty_Difference"].round(6) != 0]
            if r_only:
                d_view = d_view[(d_view["Total_Cost_Revised"]
                                 - d_view["Total_Cost_Previous"]).round(6) != 0]
            d_view = d_view[_contains(d_view["FWBS"], f_fwbs)
                            & _contains(d_view["Description"], f_desc)]

            st.markdown("**Detail Table**")
            if d_view.empty:
                st.info("조건에 해당하는 상세 항목이 없습니다.")
            else:
                st.caption(f"표시 행: {len(d_view):,} / 전체 {len(m_detail):,}")
                # PABS(=A코드 매칭 Key) 표시명 통일
                st.dataframe(d_view.rename(columns={"PABS": "PABS_Key"}),
                             use_container_width=True, hide_index=True)
