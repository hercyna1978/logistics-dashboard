# -*- coding: utf-8 -*-
"""
물류비 운영 대시보드
====================
기준 Excel:
01_상품Master
02_재고스냅샷
03_입고
04_출고
05_불량관리
06_재고관리기준
07_대시보드 KPI용 월별 집계
08_물류비 항목

1차 운영 범위
- 온라인  : CJ + 카카오
- 해외    : DHL + EMS
- 도급사  : 오프라인 + 입고 + 불량 + 샘플 + 협찬 + 일본 관련 운영비
- 제외    : 퀵/화물, 고고밴(밀크런)

중요:
도급사 비용은 현재 월별 총액만 있으므로 오프라인/입고/불량/샘플/협찬/일본별
비용을 임의 배분하지 않습니다. 도급사 총액과 도급사 대상 처리량을 별도로 비교합니다.
"""

import os
from pathlib import Path

import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st


# -----------------------------
# 기본 설정
# -----------------------------
st.set_page_config(
    page_title="물류비 운영 대시보드",
    page_icon="🚚",
    layout="wide",
    initial_sidebar_state="expanded",
)

COST_COLS = ["CJ", "카카오", "DHL", "EMS", "도급사"]
EXCLUDED_COST_COLS = ["퀵/화물", "고고밴(밀크런)"]

CHANNELS = ["온라인", "해외", "오프라인", "입고", "불량", "샘플", "협찬", "일본"]

CHANNEL_COST_MAP = {
    "온라인": ["CJ", "카카오"],
    "해외": ["DHL", "EMS"],
    "오프라인": ["도급사"],
    "입고": ["도급사"],
    "불량": ["도급사"],
    "샘플": ["도급사"],
    "협찬": ["도급사"],
    "일본": ["도급사"],
}

DOCK_CHANNELS = ["오프라인", "일본", "샘플", "협찬"]
DOCK_INBOUND = ["입고"]
DOCK_DEFECT = ["불량"]


# -----------------------------
# 유틸
# -----------------------------
def money(v):
    if pd.isna(v):
        return "-"
    return f"{v:,.0f}원"


def qty(v):
    if pd.isna(v):
        return "-"
    return f"{v:,.0f}"


def pct(v):
    if pd.isna(v):
        return "-"
    return f"{v:.1f}%"


def get_theme():
    try:
        return st.context.theme.type
    except Exception:
        try:
            return st.get_option("theme.base")
        except Exception:
            return "light"


def plot_template():
    return "plotly_dark" if get_theme() == "dark" else "plotly_white"


def clean_num(df, cols):
    for c in cols:
        if c in df.columns:
            df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0)
    return df


def normalize_columns(df):
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]
    return df


def read_excel_source(uploaded_file=None):
    """
    업로드 파일이 없으면 현재 실행 폴더의 data(2).xlsx 또는 data.xlsx를 찾습니다.
    """
    if uploaded_file is not None:
        return uploaded_file

    candidates = [
        Path("data(2).xlsx"),
        Path("data.xlsx"),
        Path("raw_data.xlsx"),
        Path("물류비_raw_data.xlsx"),
    ]
    for p in candidates:
        if p.exists():
            return p

    return None


@st.cache_data(show_spinner=False)
def load_workbook(file_source):
    xls = pd.ExcelFile(file_source)

    # 물류비 시트는 첫 행이 제목, 실제 헤더가 두 번째 행
    logistics = pd.read_excel(file_source, sheet_name="08_물류비 항목", header=1)
    logistics = normalize_columns(logistics)

    inbound = pd.read_excel(file_source, sheet_name="03_입고")
    outbound = pd.read_excel(file_source, sheet_name="04_출고")
    defect = pd.read_excel(file_source, sheet_name="05_불량관리")
    master = pd.read_excel(file_source, sheet_name="01_상품Master")
    stock = pd.read_excel(file_source, sheet_name="02_재고스냅샷")

    inbound = normalize_columns(inbound)
    outbound = normalize_columns(outbound)
    defect = normalize_columns(defect)
    master = normalize_columns(master)
    stock = normalize_columns(stock)

    # 날짜
    for df, col in [
        (logistics, None),
        (inbound, "입고일"),
        (outbound, "출고일"),
        (defect, "발생일"),
        (stock, "기준일"),
    ]:
        if col and col in df.columns:
            df[col] = pd.to_datetime(df[col], errors="coerce")

    # 물류비
    logistics["연도"] = pd.to_numeric(logistics["연도"], errors="coerce")
    logistics["월"] = pd.to_numeric(logistics["월"], errors="coerce")
    logistics = logistics.dropna(subset=["연도", "월"]).copy()
    logistics["연도"] = logistics["연도"].astype(int)
    logistics["월"] = logistics["월"].astype(int)
    logistics["년월"] = pd.to_datetime(
        logistics["연도"].astype(str) + "-" + logistics["월"].astype(str) + "-01",
        errors="coerce",
    )
    clean_num(logistics, COST_COLS + EXCLUDED_COST_COLS)
    logistics["총물류비"] = logistics[COST_COLS].sum(axis=1)
    logistics["전체입력비용"] = logistics[COST_COLS + EXCLUDED_COST_COLS].sum(axis=1)

    # 거래 수량
    clean_num(inbound, ["입고수량"])
    clean_num(outbound, ["출고수량"])
    clean_num(defect, ["불량수량"])

    inbound["년월"] = inbound["입고일"].dt.to_period("M").astype(str)
    outbound["년월"] = outbound["출고일"].dt.to_period("M").astype(str)
    defect["년월"] = defect["발생일"].dt.to_period("M").astype(str)

    # 판매채널 정규화
    if "판매채널" in outbound.columns:
        outbound["판매채널"] = outbound["판매채널"].astype(str).str.strip()
    else:
        outbound["판매채널"] = "미상"

    # 2026 월별 물동량
    inbound_m = inbound.groupby("년월", as_index=False)["입고수량"].sum()
    outbound_m = outbound.groupby("년월", as_index=False)["출고수량"].sum()
    defect_m = defect.groupby("년월", as_index=False)["불량수량"].sum()

    monthly = logistics.copy()
    monthly["년월키"] = monthly["년월"].dt.to_period("M").astype(str)

    monthly = monthly.merge(inbound_m, left_on="년월키", right_on="년월", how="left")
    monthly = monthly.drop(columns=["년월_y"], errors="ignore").rename(columns={"년월_x": "년월"})
    monthly = monthly.merge(outbound_m, left_on="년월키", right_on="년월", how="left", suffixes=("", "_out"))
    monthly = monthly.drop(columns=["년월_out"], errors="ignore")
    monthly = monthly.merge(defect_m, left_on="년월키", right_on="년월", how="left", suffixes=("", "_def"))
    monthly = monthly.drop(columns=["년월_def"], errors="ignore")

    for c in ["입고수량", "출고수량", "불량수량"]:
        if c not in monthly.columns:
            monthly[c] = 0
        monthly[c] = pd.to_numeric(monthly[c], errors="coerce").fillna(0)

    monthly["출고검수량"] = monthly["출고수량"] + monthly["불량수량"]
    monthly["전체처리량"] = monthly["입고수량"] + monthly["출고수량"] + monthly["불량수량"]
    monthly["도급사대상처리량"] = (
        monthly["입고수량"] + monthly["불량수량"] + monthly["출고수량"]
    )

    monthly["출고당물류비"] = np.where(
        monthly["출고수량"] > 0,
        monthly["총물류비"] / monthly["출고수량"],
        np.nan,
    )
    monthly["전체처리당물류비"] = np.where(
        monthly["전체처리량"] > 0,
        monthly["총물류비"] / monthly["전체처리량"],
        np.nan,
    )
    monthly["도급사처리당비용"] = np.where(
        monthly["도급사대상처리량"] > 0,
        monthly["도급사"] / monthly["도급사대상처리량"],
        np.nan,
    )

    # 비용 전월 증감
    monthly = monthly.sort_values("년월").reset_index(drop=True)
    monthly["물류비전월증감률"] = monthly["총물류비"].pct_change() * 100
    monthly["출고당비용전월증감률"] = monthly["출고당물류비"].pct_change() * 100

    return {
        "logistics": logistics,
        "monthly": monthly,
        "inbound": inbound,
        "outbound": outbound,
        "defect": defect,
        "master": master,
        "stock": stock,
        "sheets": xls.sheet_names,
    }


def make_month_label(df):
    x = df.copy()
    x["표시월"] = x["년월"].dt.strftime("%Y-%m")
    return x


def safe_sum(df, cols):
    return df[cols].sum().sum() if cols else 0


def style_fig(fig, height=380):
    fig.update_layout(
        template=plot_template(),
        height=height,
        margin=dict(l=20, r=20, t=55, b=30),
        hovermode="x unified",
        legend=dict(orientation="h", yanchor="bottom", y=1.02, x=0),
    )
    return fig


# -----------------------------
# 사이드바
# -----------------------------
st.sidebar.title("🚚 물류비 운영")

source = read_excel_source()

uploaded = st.sidebar.file_uploader(
    "Excel 파일 선택",
    type=["xlsx", "xls"],
    help="현재 사용 중인 로우데이터 Excel을 선택하세요.",
)

if uploaded is not None:
    source = uploaded

if source is None:
    st.error(
        "Excel 파일을 찾지 못했습니다. `data(2).xlsx`를 앱 파일과 같은 폴더에 두거나 "
        "왼쪽에서 Excel 파일을 업로드하세요."
    )
    st.stop()

try:
    data = load_workbook(source)
except Exception as e:
    st.error(f"Excel을 읽는 중 오류가 발생했습니다: {e}")
    st.stop()

monthly = data["monthly"].copy()
logistics = data["logistics"].copy()
outbound = data["outbound"].copy()
inbound = data["inbound"].copy()
defect = data["defect"].copy()

# 기간
available_years = sorted(monthly["연도"].dropna().unique().astype(int).tolist())
selected_years = st.sidebar.multiselect(
    "분석 연도",
    options=available_years,
    default=available_years[-1:] if available_years else [],
)

if not selected_years:
    selected_years = available_years

year_monthly = monthly[monthly["연도"].isin(selected_years)].copy()

if year_monthly.empty:
    st.warning("선택한 연도에 물류비 데이터가 없습니다.")
    st.stop()

# 비용항목
selected_costs = st.sidebar.multiselect(
    "물류비 항목",
    options=COST_COLS,
    default=COST_COLS,
)

if not selected_costs:
    selected_costs = COST_COLS

# 2026년 물동량 채널
selected_channels = st.sidebar.multiselect(
    "출고 채널",
    options=sorted(outbound["판매채널"].dropna().unique().tolist()),
    default=sorted(outbound["판매채널"].dropna().unique().tolist()),
)

if not selected_channels:
    selected_channels = sorted(outbound["판매채널"].dropna().unique().tolist())

# 기간 필터
min_month = year_monthly["년월"].min()
max_month = year_monthly["년월"].max()

st.sidebar.caption(
    f"물류비 데이터: {min_month:%Y-%m} ~ {max_month:%Y-%m}"
)
st.sidebar.caption("퀵/화물 · 고고밴은 1차 대시보드에서 제외")


# -----------------------------
# 데이터 필터
# -----------------------------
m = year_monthly.copy()
m["선택물류비"] = m[selected_costs].sum(axis=1)

# 선택 채널 기준 출고량
out_sel = outbound[outbound["판매채널"].isin(selected_channels)].copy()
out_channel_m = (
    out_sel.groupby("년월", as_index=False)["출고수량"]
    .sum()
    .rename(columns={"출고수량": "선택채널출고수량"})
)

m = m.merge(out_channel_m, left_on="년월키", right_on="년월", how="left", suffixes=("", "_channel"))
m["선택채널출고수량"] = pd.to_numeric(
    m["선택채널출고수량"], errors="coerce"
).fillna(0)

m["선택출고당물류비"] = np.where(
    m["선택채널출고수량"] > 0,
    m["선택물류비"] / m["선택채널출고수량"],
    np.nan,
)

# -----------------------------
# 헤더
# -----------------------------
st.title("🚚 물류비 운영 대시보드")
st.caption(
    "08_물류비 항목을 기준으로 물류비 구조와 물동량을 연결해 보는 1차 운영 버전"
)

# -----------------------------
# KPI
# -----------------------------
total_cost = m["선택물류비"].sum()
total_out = m["선택채널출고수량"].sum()
total_processing = m["전체처리량"].sum()
avg_monthly_cost = m["선택물류비"].mean()

cost_per_out = total_cost / total_out if total_out else np.nan
cost_per_process = total_cost / total_processing if total_processing else np.nan

prev = m["선택물류비"].iloc[-2] if len(m) >= 2 else np.nan
last = m["선택물류비"].iloc[-1] if len(m) else np.nan
mom = ((last / prev) - 1) * 100 if pd.notna(prev) and prev != 0 else np.nan

k1, k2, k3, k4, k5, k6 = st.columns(6)
k1.metric("총 물류비", money(total_cost))
k2.metric("월평균 물류비", money(avg_monthly_cost))
k3.metric("출고수량", qty(total_out))
k4.metric("출고당 물류비", money(cost_per_out))
k5.metric("전체 처리량당 비용", money(cost_per_process))
k6.metric("최근월 전월 대비", pct(mom))

st.divider()

# -----------------------------
# 1. 월별 물류비
# -----------------------------
st.subheader("① 월별 물류비 추이")

c1, c2 = st.columns([1.5, 1])

with c1:
    trend = m.copy()
    trend["표시월"] = trend["년월"].dt.strftime("%Y-%m")

    fig = go.Figure()
    for col in selected_costs:
        fig.add_trace(
            go.Bar(
                x=trend["표시월"],
                y=trend[col],
                name=col,
            )
        )
    fig.add_trace(
        go.Scatter(
            x=trend["표시월"],
            y=trend["선택물류비"],
            name="선택항목 합계",
            mode="lines+markers",
        )
    )
    fig.update_layout(barmode="stack", yaxis_title="비용(원)", xaxis_title="")
    st.plotly_chart(style_fig(fig, 430), use_container_width=True)

with c2:
    latest = m.sort_values("년월").iloc[-1]
    latest_cost = latest["선택물류비"]

    pie_df = pd.DataFrame({
        "항목": selected_costs,
        "비용": [latest[c] for c in selected_costs],
    })
    pie_df = pie_df[pie_df["비용"] > 0]

    if not pie_df.empty:
        fig = px.pie(
            pie_df,
            names="항목",
            values="비용",
            hole=0.55,
            title=f"{latest['년월']:%Y-%m} 비용 구성",
        )
        st.plotly_chart(style_fig(fig, 430), use_container_width=True)

# -----------------------------
# 2. 물류비 vs 물동량
# -----------------------------
st.subheader("② 물류비와 물동량")

fig = go.Figure()
fig.add_trace(
    go.Bar(
        x=m["년월"].dt.strftime("%Y-%m"),
        y=m["출고수량"],
        name="전체 출고수량",
    )
)
fig.add_trace(
    go.Scatter(
        x=m["년월"].dt.strftime("%Y-%m"),
        y=m["총물류비"],
        name="총 물류비",
        mode="lines+markers",
        yaxis="y2",
    )
)
fig.update_layout(
    yaxis=dict(title="출고수량"),
    yaxis2=dict(title="물류비(원)", overlaying="y", side="right"),
)
st.plotly_chart(style_fig(fig, 420), use_container_width=True)

st.info(
    "해석 포인트: 물류비가 증가했더라도 출고량이 함께 증가하면 정상적인 물동량 증가일 수 있습니다. "
    "반대로 출고량은 비슷한데 물류비 또는 출고당 물류비가 급등하면 운영 점검 대상입니다."
)

# -----------------------------
# 3. 단위 물류비
# -----------------------------
st.subheader("③ 단위 물류비 추이")

fig = go.Figure()
fig.add_trace(
    go.Scatter(
        x=m["년월"].dt.strftime("%Y-%m"),
        y=m["출고당물류비"],
        mode="lines+markers",
        name="전체 출고당 물류비",
    )
)
fig.add_trace(
    go.Scatter(
        x=m["년월"].dt.strftime("%Y-%m"),
        y=m["전체처리당물류비"],
        mode="lines+markers",
        name="전체 처리량당 물류비",
    )
)
fig.update_layout(yaxis_title="원 / 수량")
st.plotly_chart(style_fig(fig, 400), use_container_width=True)

# -----------------------------
# 4. 온라인
# -----------------------------
st.subheader("④ 온라인 물류비 — CJ + 카카오")

online_cost = logistics[logistics["연도"].isin(selected_years)].copy()
online_cost["온라인물류비"] = online_cost[["CJ", "카카오"]].sum(axis=1)
online_out = (
    outbound[
        (outbound["판매채널"] == "온라인")
        & (outbound["출고일"].dt.year.isin(selected_years))
    ]
    .groupby("년월", as_index=False)["출고수량"]
    .sum()
    .rename(columns={"출고수량": "온라인출고수량"})
)

online = online_cost.merge(
    online_out,
    left_on="년월키",
    right_on="년월",
    how="left",
)
online["온라인출고수량"] = online["온라인출고수량"].fillna(0)
online["온라인출고당비용"] = np.where(
    online["온라인출고수량"] > 0,
    online["온라인물류비"] / online["온라인출고수량"],
    np.nan,
)

a, b = st.columns(2)
with a:
    fig = px.bar(
        online,
        x=online["년월"].dt.strftime("%Y-%m"),
        y=["CJ", "카카오"],
        title="CJ / 카카오 비용",
        barmode="stack",
    )
    st.plotly_chart(style_fig(fig, 380), use_container_width=True)

with b:
    fig = go.Figure()
    fig.add_trace(
        go.Bar(
            x=online["년월"].dt.strftime("%Y-%m"),
            y=online["온라인출고수량"],
            name="온라인 출고량",
        )
    )
    fig.add_trace(
        go.Scatter(
            x=online["년월"].dt.strftime("%Y-%m"),
            y=online["온라인출고당비용"],
            name="온라인 출고당 비용",
            mode="lines+markers",
            yaxis="y2",
        )
    )
    fig.update_layout(
        title="온라인 물동량 vs 출고당 비용",
        yaxis=dict(title="출고수량"),
        yaxis2=dict(title="원/출고수량", overlaying="y", side="right"),
    )
    st.plotly_chart(style_fig(fig, 380), use_container_width=True)

st.caption(
    "주의: 현재 데이터에는 CJ/카카오별 실제 배송 건수 또는 송장 연결값이 없으므로 "
    "CJ·카카오 각각의 실제 건당 배송단가는 계산하지 않습니다."
)

# -----------------------------
# 5. 해외
# -----------------------------
st.subheader("⑤ 해외 물류비 — DHL + EMS")

overseas_cost = logistics[logistics["연도"].isin(selected_years)].copy()
overseas_cost["해외물류비"] = overseas_cost[["DHL", "EMS"]].sum(axis=1)

overseas_out = (
    outbound[
        (outbound["판매채널"] == "해외")
        & (outbound["출고일"].dt.year.isin(selected_years))
    ]
    .groupby("년월", as_index=False)["출고수량"]
    .sum()
    .rename(columns={"출고수량": "해외출고수량"})
)

overseas = overseas_cost.merge(
    overseas_out,
    left_on="년월키",
    right_on="년월",
    how="left",
)
overseas["해외출고수량"] = overseas["해외출고수량"].fillna(0)
overseas["해외출고당비용"] = np.where(
    overseas["해외출고수량"] > 0,
    overseas["해외물류비"] / overseas["해외출고수량"],
    np.nan,
)

a, b = st.columns(2)
with a:
    fig = px.bar(
        overseas,
        x=overseas["년월"].dt.strftime("%Y-%m"),
        y=["DHL", "EMS"],
        title="DHL / EMS 비용",
        barmode="stack",
    )
    st.plotly_chart(style_fig(fig, 380), use_container_width=True)

with b:
    fig = go.Figure()
    fig.add_trace(
        go.Bar(
            x=overseas["년월"].dt.strftime("%Y-%m"),
            y=overseas["해외출고수량"],
            name="해외 출고량",
        )
    )
    fig.add_trace(
        go.Scatter(
            x=overseas["년월"].dt.strftime("%Y-%m"),
            y=overseas["해외출고당비용"],
            name="해외 출고당 비용",
            mode="lines+markers",
            yaxis="y2",
        )
    )
    fig.update_layout(
        title="해외 물동량 vs 출고당 비용",
        yaxis=dict(title="출고수량"),
        yaxis2=dict(title="원/출고수량", overlaying="y", side="right"),
    )
    st.plotly_chart(style_fig(fig, 380), use_container_width=True)

# -----------------------------
# 6. 도급사
# -----------------------------
st.subheader("⑥ 도급사 운영비")

st.markdown(
    "**도급사 비용은 오프라인 + 입고 + 불량 + 샘플 + 협찬 + 일본 업무를 포괄하는 월 총액으로 관리합니다.** "
    "현재 데이터만으로 세부 업무별 비용을 임의 배분하지 않습니다."
)

contract = m.copy()
contract["도급사비용"] = contract["도급사"]

a, b = st.columns(2)

with a:
    fig = go.Figure()
    fig.add_trace(
        go.Bar(
            x=contract["년월"].dt.strftime("%Y-%m"),
            y=contract["도급사비용"],
            name="도급사 비용",
        )
    )
    fig.add_trace(
        go.Scatter(
            x=contract["년월"].dt.strftime("%Y-%m"),
            y=contract["도급사대상처리량"],
            name="도급사 대상 처리량",
            mode="lines+markers",
            yaxis="y2",
        )
    )
    fig.update_layout(
        title="도급사 비용 vs 대상 처리량",
        yaxis=dict(title="도급사 비용(원)"),
        yaxis2=dict(title="처리량", overlaying="y", side="right"),
    )
    st.plotly_chart(style_fig(fig, 400), use_container_width=True)

with b:
    fig = px.line(
        contract,
        x=contract["년월"].dt.strftime("%Y-%m"),
        y="도급사처리당비용",
        markers=True,
        title="도급사 처리량당 비용",
    )
    fig.update_yaxes(title="원 / 처리수량")
    st.plotly_chart(style_fig(fig, 400), use_container_width=True)

# 도급사 대상 처리량 구성
dock_parts = pd.DataFrame({
    "업무": ["입고", "불량", "오프라인", "일본", "샘플", "협찬"],
    "처리량": [
        inbound[inbound["년월"].str[:4].astype(int).isin(selected_years)]["입고수량"].sum(),
        defect[defect["년월"].str[:4].astype(int).isin(selected_years)]["불량수량"].sum(),
        outbound[
            (outbound["판매채널"] == "오프라인")
            & (outbound["출고일"].dt.year.isin(selected_years))
        ]["출고수량"].sum(),
        outbound[
            (outbound["판매채널"] == "일본")
            & (outbound["출고일"].dt.year.isin(selected_years))
        ]["출고수량"].sum(),
        outbound[
            (outbound["판매채널"] == "샘플")
            & (outbound["출고일"].dt.year.isin(selected_years))
        ]["출고수량"].sum(),
        outbound[
            (outbound["판매채널"] == "협찬")
            & (outbound["출고일"].dt.year.isin(selected_years))
        ]["출고수량"].sum(),
    ],
})

fig = px.bar(
    dock_parts,
    x="업무",
    y="처리량",
    title="도급사 관련 업무별 처리량",
    text_auto=".3s",
)
st.plotly_chart(style_fig(fig, 360), use_container_width=True)

# -----------------------------
# 7. 채널별 물동량
# -----------------------------
st.subheader("⑦ 출고 채널별 물동량")

channel_out = (
    outbound[
        outbound["출고일"].dt.year.isin(selected_years)
        & outbound["판매채널"].isin(selected_channels)
    ]
    .groupby(["년월", "판매채널"], as_index=False)["출고수량"]
    .sum()
)

if not channel_out.empty:
    channel_pivot = channel_out.pivot(
        index="년월", columns="판매채널", values="출고수량"
    ).fillna(0).reset_index()

    fig = px.bar(
        channel_out,
        x="년월",
        y="출고수량",
        color="판매채널",
        title="월별 채널별 출고량",
    )
    fig.update_xaxes(type="category")
    st.plotly_chart(style_fig(fig, 420), use_container_width=True)

# -----------------------------
# 8. 이상징후
# -----------------------------
st.subheader("⑧ 물류비 이상징후 / 운영 점검")

alerts = []

for _, r in m.iterrows():
    ym = r["년월"].strftime("%Y-%m")
    if pd.notna(r["물류비전월증감률"]) and r["물류비전월증감률"] >= 20:
        alerts.append(
            {
                "월": ym,
                "구분": "물류비 급증",
                "내용": f"전월 대비 물류비 {r['물류비전월증감률']:.1f}% 증가",
                "수치": r["물류비전월증감률"],
            }
        )
    if pd.notna(r["출고당비용전월증감률"]) and r["출고당비용전월증감률"] >= 20:
        alerts.append(
            {
                "월": ym,
                "구분": "단위 물류비 급증",
                "내용": f"출고당 물류비 {r['출고당비용전월증감률']:.1f}% 증가",
                "수치": r["출고당비용전월증감률"],
            }
        )

# 2026년 9월처럼 물류비가 아직 없는 달은 경고하지 않고,
# 물류비 최신 월 이후의 거래월이 있는 경우 참고 표시
latest_logistics_month = logistics["년월"].max()
latest_outbound_month = outbound["출고일"].max().to_period("M").to_timestamp()

if latest_outbound_month > latest_logistics_month:
    alerts.append(
        {
            "월": latest_outbound_month.strftime("%Y-%m"),
            "구분": "물류비 미입력",
            "내용": f"출고 데이터는 {latest_outbound_month:%Y-%m}까지 있으나 물류비 데이터는 {latest_logistics_month:%Y-%m}까지 입력됨",
            "수치": np.nan,
        }
    )

if alerts:
    alert_df = pd.DataFrame(alerts)
    st.dataframe(
        alert_df[["월", "구분", "내용"]],
        use_container_width=True,
        hide_index=True,
    )
else:
    st.success("현재 설정된 기준에서 확인할 이상징후가 없습니다.")

# -----------------------------
# 9. 월별 상세표
# -----------------------------
st.subheader("⑨ 월별 운영 상세")

detail_cols = [
    "년월",
    *selected_costs,
    "선택물류비",
    "입고수량",
    "출고수량",
    "불량수량",
    "출고검수량",
    "전체처리량",
    "출고당물류비",
    "전체처리당물류비",
    "도급사처리당비용",
]

detail = m[detail_cols].copy()
detail["년월"] = detail["년월"].dt.strftime("%Y-%m")

rename_map = {
    "선택물류비": "선택 물류비",
    "입고수량": "입고",
    "출고수량": "출고",
    "불량수량": "불량",
    "출고검수량": "출고 검수량",
    "전체처리량": "전체 처리량",
    "출고당물류비": "출고당 물류비",
    "전체처리당물류비": "전체 처리당 물류비",
    "도급사처리당비용": "도급사 처리당 비용",
}
detail = detail.rename(columns=rename_map)

st.dataframe(
    detail,
    use_container_width=True,
    hide_index=True,
    column_config={
        c: st.column_config.NumberColumn(format="₩%d")
        for c in selected_costs + ["선택 물류비"]
        if c in detail.columns
    },
)

# -----------------------------
# 10. 데이터 구조 / 향후 개선
# -----------------------------
with st.expander("🔎 데이터 연결 기준 및 향후 개선사항", expanded=False):
    st.markdown(
        """
### 현재 적용한 비용 연결 기준
- 온라인 → CJ + 카카오
- 해외 → DHL + EMS
- 오프라인 → 도급사
- 입고 → 도급사
- 불량 → 도급사
- 샘플 → 도급사
- 협찬 → 도급사
- 일본 → 도급사
- 퀵/화물, 고고밴 → 제외

### 현재 데이터로 정확히 계산하지 않는 것
- CJ 실제 건당 배송단가
- 카카오 실제 건당 배송단가
- DHL 실제 건당 배송단가
- EMS 실제 건당 배송단가
- 도급사 비용의 업무별 세부 배분액

위 항목은 배송사/송장/주문번호/업무별 비용 등의 추가 데이터가 있어야 정확하게 계산할 수 있습니다.

### 다음 개선 단계
1. 출고 시트에 배송사(CJ/카카오/DHL/EMS 등) 추가
2. 주문번호 또는 송장번호 추가
3. 도급사 작업유형별 처리량 추가
4. 물류비 상세내역과 연결
5. 업체별 실제 건당 배송비 및 비용 효율 분석
"""
    )

st.caption("※ 현재 버전은 08_물류비 항목의 퀵/화물·고고밴을 계산에서 제외합니다.")
