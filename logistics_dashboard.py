"""
물류비 운영 대시보드 (Streamlit)
--------------------------------
실행 방법:
    pip install -r requirements.txt
    streamlit run logistics_dashboard.py

기준 원본 파일: data(2).xlsx (또는 data.xlsx) — 이 스크립트와 같은 폴더에 두세요.
    - 08_물류비 항목 : 월별 물류비 원장 (CJ, 카카오, DHL, EMS, 퀵/화물, 도급사, 고고밴)
    - 04_출고        : 출고일, 상품코드, 상품명, 출고수량, 판매채널
    - 03_입고        : 입고일, 상품코드, 상품명, 입고수량, 중국공장
    - 05_불량관리    : 발생일, 상품코드, 상품명, 불량수량, 불량유형, 중국공장

채널 매핑 기준
    - 해외   = DHL + EMS 비용            / 출고채널 '해외' 물동량
    - 온라인 = CJ + 카카오 비용          / 출고채널 '온라인' 물동량
    - 도급사 = 도급사 비용 1개 계정      / 출고(오프라인+일본+협찬+샘플) + 입고 + 불량 물동량 합산
    - 제외   = 퀵/화물, 고고밴(밀크런)
"""

import os
import glob
import pandas as pd
import numpy as np
import streamlit as st
import plotly.graph_objects as go

# ----------------------------------------------------------------------
# 0. 페이지 설정 & 스타일
# ----------------------------------------------------------------------
st.set_page_config(page_title="물류비 운영 대시보드", layout="wide", page_icon="📦")

INK = "#22262F"
INK_SOFT = "#6B7280"
PAPER = "#F5F4F0"
LINE = "#E4E1D8"
OVERSEAS = "#2E6F63"
ONLINE = "#35507E"
CONTRACTOR = "#AD5D2E"
WARN = "#B33A3A"
WARN_BG = "#FBEDEA"

st.markdown(
    f"""
    <style>
    .stApp {{ background-color: {PAPER}; }}
    .block-container {{ padding-top: 1.6rem; }}
    div[data-testid="stMetric"] {{
        background: #fff; border: 1px solid {LINE}; border-radius: 10px;
        padding: 14px 18px 10px 18px;
    }}
    div[data-testid="stMetricLabel"] {{ color: {INK_SOFT}; font-weight: 600; }}
    div[data-testid="stMetricValue"] {{ color: {INK}; }}
    .warn-banner {{
        background: {WARN_BG}; border: 1px solid {WARN}55; border-radius: 10px;
        padding: 12px 16px; color: {WARN}; font-size: 13.5px; line-height: 1.6; margin-bottom: 14px;
    }}
    .footnote {{ color: {INK_SOFT}; font-size: 11.5px; text-align: right; margin-top: 6px; }}
    </style>
    """,
    unsafe_allow_html=True,
)


# ----------------------------------------------------------------------
# 1. 원본 파일 로드
# ----------------------------------------------------------------------
def find_source_file():
    here = os.path.dirname(os.path.abspath(__file__))
    candidates = ["data(2).xlsx", "data (2).xlsx", "data2.xlsx", "data.xlsx"]
    for c in candidates:
        p = os.path.join(here, c)
        if os.path.exists(p):
            return p
    matches = glob.glob(os.path.join(here, "data*.xlsx"))
    if matches:
        return matches[0]
    return None


@st.cache_data
def load_data(path):
    xls = pd.ExcelFile(path)

    cost_raw = pd.read_excel(xls, "08_물류비 항목", header=1)
    cost_raw = cost_raw.dropna(subset=["연도", "월"]).copy()
    cost_raw["연도"] = cost_raw["연도"].astype(int)
    cost_raw["월"] = cost_raw["월"].astype(int)
    cost_raw["ym"] = cost_raw["연도"].astype(str) + "-" + cost_raw["월"].astype(str).str.zfill(2)
    for c in ["CJ", "카카오", "DHL", "EMS", "퀵/화물", "도급사", "고고밴(밀크런)"]:
        if c not in cost_raw.columns:
            cost_raw[c] = 0
        cost_raw[c] = cost_raw[c].fillna(0)

    cost = pd.DataFrame({
        "ym": cost_raw["ym"],
        "overseas": cost_raw["DHL"] + cost_raw["EMS"],
        "online": cost_raw["CJ"] + cost_raw["카카오"],
        "contractor": cost_raw["도급사"],
    }).reset_index(drop=True)
    # 데이터 이상치 의심 플래그 (해외 급등 + 도급사 급락 패턴)
    cost["flagged"] = False
    for i in range(1, len(cost)):
        prev, cur = cost.iloc[i - 1], cost.iloc[i]
        if prev["contractor"] > 0 and cur["contractor"] < prev["contractor"] * 0.3 and cur["overseas"] > prev["overseas"] * 2:
            cost.loc[i, "flagged"] = True

    outbound = pd.read_excel(xls, "04_출고")
    inbound = pd.read_excel(xls, "03_입고")
    defect = pd.read_excel(xls, "05_불량관리")

    outbound["ym"] = pd.to_datetime(outbound["출고일"]).dt.strftime("%Y-%m")
    inbound["ym"] = pd.to_datetime(inbound["입고일"]).dt.strftime("%Y-%m")
    defect["ym"] = pd.to_datetime(defect["발생일"]).dt.strftime("%Y-%m")

    vol_months = sorted(outbound["ym"].unique())
    rows = []
    for ym in vol_months:
        ob = outbound[outbound["ym"] == ym]
        by_ch = ob.groupby("판매채널")["출고수량"].sum()
        offline = by_ch.get("오프라인", 0)
        japan = by_ch.get("일본", 0)
        sponsor = by_ch.get("협찬", 0)
        sample = by_ch.get("샘플", 0)
        overseas_v = by_ch.get("해외", 0)
        online_v = by_ch.get("온라인", 0)
        in_v = inbound[inbound["ym"] == ym]["입고수량"].sum()
        def_v = defect[defect["ym"] == ym]["불량수량"].sum()
        contractor_v = offline + japan + sponsor + sample + in_v + def_v
        rows.append(dict(ym=ym, overseas=overseas_v, online=online_v, contractor=contractor_v))
    volume = pd.DataFrame(rows)

    unit = cost.merge(volume, on="ym", suffixes=("_cost", "_vol"))
    for ch in ["overseas", "online", "contractor"]:
        unit[f"{ch}_unit"] = np.where(unit[f"{ch}_vol"] > 0, unit[f"{ch}_cost"] / unit[f"{ch}_vol"], np.nan)

    return cost, volume, unit


def won(n):
    try:
        return "₩" + f"{round(n):,}"
    except Exception:
        return "-"


def won_short(n):
    n = round(n)
    if abs(n) >= 1e8:
        return f"{n/1e8:.1f}억"
    if abs(n) >= 1e4:
        return f"{n/1e4:.0f}만"
    return f"{n:,}"


# ----------------------------------------------------------------------
# 2. 데이터 로드
# ----------------------------------------------------------------------
src_path = find_source_file()
if src_path is None:
    st.error(
        "원본 데이터 파일을 찾을 수 없습니다. `data(2).xlsx` (또는 data.xlsx)를 "
        "이 스크립트와 같은 폴더에 넣어주세요."
    )
    st.stop()

try:
    cost, volume, unit = load_data(src_path)
except Exception as e:
    st.error(f"데이터를 읽는 중 오류가 발생했습니다: {e}")
    st.stop()

CHANNELS = [
    ("overseas", "해외", "DHL + EMS", OVERSEAS),
    ("online", "온라인", "CJ + 카카오", ONLINE),
    ("contractor", "도급사", "오프라인+일본+협찬+샘플+입고+불량", CONTRACTOR),
]

# ----------------------------------------------------------------------
# 3. 헤더
# ----------------------------------------------------------------------
st.markdown(
    f"""
    <div style="font-size:24px; font-weight:700; color:{INK};">물류비 운영 대시보드</div>
    <div style="font-size:13px; color:{INK_SOFT}; margin-top:4px; margin-bottom:18px;">
    비용 {cost['ym'].min()} ~ {cost['ym'].max()} · 물동량·단가 {volume['ym'].min()} ~ {volume['ym'].max()}
    (그 외 기간은 물동량 데이터 없음) · <b style="color:{INK};">퀵/화물, 고고밴(밀크런)은 범위 제외</b>
    </div>
    """,
    unsafe_allow_html=True,
)

# ----------------------------------------------------------------------
# 4. KPI 카드 (최신월 기준)
# ----------------------------------------------------------------------
last_row = cost.iloc[-1]
prev_row = cost.iloc[-2] if len(cost) > 1 else None

cols = st.columns(3)
for (key, label, sub, color), col in zip(CHANNELS, cols):
    cur = last_row[key]
    mom = None
    if prev_row is not None and prev_row[key] not in (0, None):
        mom = (cur - prev_row[key]) / prev_row[key] * 100
    with col:
        st.metric(
            label=f"{label} · {sub}",
            value=won(cur),
            delta=f"{mom:+.1f}% (전월대비)" if mom is not None else None,
            delta_color="inverse",
        )
if last_row["flagged"]:
    st.caption(f"⚠ {last_row['ym']} 데이터는 이상치 의심 구간입니다 (아래 안내 참고).")

# ----------------------------------------------------------------------
# 5. 이상치 경고 배너
# ----------------------------------------------------------------------
flagged_months = cost.loc[cost["flagged"], "ym"].tolist()
if flagged_months:
    st.markdown(
        f"""
        <div class="warn-banner">
        <b>⚠ 데이터 확인 필요</b> — {", ".join(flagged_months)} 기간은 해외(DHL/EMS)와 도급사 값이
        서로 뒤바뀐 것으로 보입니다. 도급사 단가가 갑자기 급락하고 해외 비용이 급등하는 패턴이 근거입니다.
        원본 08_물류비 항목 시트를 재확인해 주세요. (이 대시보드는 원본 값을 임의로 수정하지 않고 그대로 표시합니다.)
        </div>
        """,
        unsafe_allow_html=True,
    )

# ----------------------------------------------------------------------
# 6. 월별 비용 추이
# ----------------------------------------------------------------------
st.subheader("월별 물류비 추이")
range_opt = st.radio(
    "기간", ["전체", "2025~2026", "2026년"], horizontal=True, label_visibility="collapsed"
)
if range_opt == "2026년":
    plot_cost = cost[cost["ym"].str.startswith("2026")]
elif range_opt == "2025~2026":
    plot_cost = cost[cost["ym"] >= "2025-01"]
else:
    plot_cost = cost

fig = go.Figure()
for key, label, sub, color in CHANNELS:
    fig.add_trace(go.Scatter(
        x=plot_cost["ym"], y=plot_cost[key], name=label,
        mode="lines", line=dict(color=color, width=2.6),
        hovertemplate="%{x}<br>" + label + ": ₩%{y:,.0f}<extra></extra>",
    ))
for ym in flagged_months:
    if ym in plot_cost["ym"].values:
        fig.add_vrect(x0=ym, x1=ym, fillcolor=WARN, opacity=0.08, line_width=0)
fig.update_layout(
    height=340, plot_bgcolor="white", paper_bgcolor="white",
    margin=dict(t=10, l=10, r=10, b=10),
    legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="left", x=0),
    yaxis=dict(tickformat=",", gridcolor=LINE),
    xaxis=dict(gridcolor=LINE),
)
st.plotly_chart(fig, use_container_width=True)

# ----------------------------------------------------------------------
# 7. 누적 비중 + 2026 단가 추이
# ----------------------------------------------------------------------
c1, c2 = st.columns([1, 1.4])

with c1:
    st.subheader("누적 비용 비중")
    st.caption(f"{cost['ym'].min()} ~ {cost['ym'].max()}")
    totals = {key: cost[key].sum() for key, *_ in CHANNELS}
    grand_total = sum(totals.values())
    for key, label, sub, color in CHANNELS:
        pct = totals[key] / grand_total * 100 if grand_total else 0
        st.markdown(
            f"""
            <div style="margin-bottom:12px;">
              <div style="display:flex; justify-content:space-between; font-size:13px; margin-bottom:4px;">
                <span style="color:{INK}; font-weight:600;">{label}</span>
                <span style="color:{INK_SOFT};">{won(totals[key])} · {pct:.1f}%</span>
              </div>
              <div style="height:8px; border-radius:6px; background:#EEECE5; overflow:hidden;">
                <div style="width:{pct}%; height:100%; background:{color}; border-radius:6px;"></div>
              </div>
            </div>
            """,
            unsafe_allow_html=True,
        )
    st.markdown(f"**총 누적 물류비: {won(grand_total)}**")

with c2:
    st.subheader("2026년 채널별 단가 추이")
    st.caption("비용 ÷ 물동량 (원/개)")
    fig2 = go.Figure()
    for key, label, sub, color in CHANNELS:
        fig2.add_trace(go.Bar(
            x=unit["ym"], y=unit[f"{key}_unit"], name=label, marker_color=color,
            hovertemplate="%{x}<br>" + label + ": ₩%{y:,.0f}/개<extra></extra>",
        ))
    fig2.update_layout(
        height=300, barmode="group", plot_bgcolor="white", paper_bgcolor="white",
        margin=dict(t=10, l=10, r=10, b=10),
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="left", x=0),
        yaxis=dict(tickformat=",", gridcolor=LINE),
        xaxis=dict(gridcolor=LINE),
    )
    st.plotly_chart(fig2, use_container_width=True)
    if any(unit["flagged"]):
        st.markdown(
            f'<span style="color:{WARN}; font-size:12px;">⚠ 이상치 의심 구간은 단가 신뢰도가 낮습니다.</span>',
            unsafe_allow_html=True,
        )

st.markdown(
    f"""
    <div class="footnote">
    기준: 해외=DHL+EMS · 온라인=CJ+카카오 · 도급사=오프라인+일본+협찬+샘플 출고, 입고, 불량 합산 · 퀵/화물·고고밴 제외
    </div>
    """,
    unsafe_allow_html=True,
)
