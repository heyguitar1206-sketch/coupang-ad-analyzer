import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from io import BytesIO

st.set_page_config(
    page_title="쿠팡 광고 분석기",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)


def auto_map_columns(df):
    mapping_rules = {
        "날짜": ["날짜", "date", "일자", "기간"],
        "캠페인명": ["캠페인명", "캠페인", "campaign"],
        "광고그룹명": ["광고그룹명", "광고그룹", "adgroup"],
        "상품명": ["상품명", "광고상품명", "product", "제품명"],
        "상품ID": ["상품id", "productid", "상품번호"],
        "키워드": ["키워드", "keyword", "검색어"],
        "노출수": ["노출수", "impression", "impressions"],
        "클릭수": ["클릭수", "click", "clicks"],
        "광고비": ["광고비", "집행광고비", "cost", "spend"],
        "직접전환판매수": ["직접전환판매수", "직접전환판매수(14일)"],
        "간접전환판매수": ["간접전환판매수", "간접전환판매수(14일)"],
        "직접전환매출액": ["직접전환매출액", "직접전환매출액(14일)"],
        "간접전환매출액": ["간접전환매출액", "간접전환매출액(14일)"],
        "총전환판매수": ["총전환판매수", "총전환판매수(14일)", "전환판매수", "광고전환판매수"],
        "총전환매출액": ["총전환매출액", "총전환매출액(14일)", "전환매출액", "광고전환매출액"],
    }
    result = {}
    used = set()
    for standard, candidates in mapping_rules.items():
        for col in df.columns:
            if col in used:
                continue
            col_clean = col.lower().replace(" ", "").replace("_", "").replace("(", "").replace(")", "")
            for cand in candidates:
                cand_clean = cand.lower().replace(" ", "").replace("_", "").replace("(", "").replace(")", "")
                if cand_clean in col_clean or col_clean in cand_clean:
                    result[col] = standard
                    used.add(col)
                    break
            if list(result.values()).count(standard) >= 1:
                break
    return result


def clean_numeric(df, cols):
    df = df.copy()
    for col in cols:
        if col in df.columns:
            df[col] = pd.to_numeric(
                df[col].astype(str).str.replace(",", "").str.replace("원", "").str.replace("%", "").str.strip(),
                errors="coerce",
            ).fillna(0)
    return df


def calc_metrics(df):
    df = df.copy()
    eps = 1e-10
    df["CTR(%)"] = (df["클릭수"] / (df["노출수"] + eps) * 100).round(2)
    df["CPC(원)"] = (df["광고비"] / (df["클릭수"] + eps)).round(0)
    df["CVR(%)"] = (df["총전환판매수"] / (df["클릭수"] + eps) * 100).round(2)
    df["ROAS(%)"] = (df["총전환매출액"] / (df["광고비"] + eps) * 100).round(0)
    df["ACoS(%)"] = (df["광고비"] / (df["총전환매출액"] + eps) * 100).round(2)
    df["건당전환비용"] = (df["광고비"] / (df["총전환판매수"] + eps)).round(0)
    for c in ["CTR(%)", "CPC(원)", "CVR(%)", "ROAS(%)", "ACoS(%)", "건당전환비용"]:
        df.loc[df[c] > 1e8, c] = 0
    return df


def safe_div(a, b):
    return (a / b) if b > 0 else 0


def format_number(n):
    if abs(n) >= 1e8:
        return f"{n/1e8:,.1f}억"
    elif abs(n) >= 1e4:
        return f"{n/1e4:,.0f}만"
    else:
        return f"{n:,.0f}"


def to_excel_bytes(sheet_dict):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        for name, dataframe in sheet_dict.items():
            dataframe.to_excel(writer, sheet_name=name[:31], index=False)
    return output.getvalue()
with st.sidebar:
    st.title("📊 쿠팡 광고 분석기")
    st.caption("v1.0")
    st.markdown("---")
    st.markdown("### 사용법")
    st.markdown("1. 쿠팡 WING → 광고보고서\n2. 엑셀 다운로드\n3. 여기에 업로드!")
    st.markdown("---")
    target_roas = st.number_input("🎯 목표 ROAS (%)", min_value=0, value=300, step=50)
    uploaded = st.file_uploader("📁 엑셀 파일 업로드", type=["xlsx", "xls", "csv"])

if uploaded is None:
    st.markdown(
        "<div style='text-align:center; padding:80px 20px;'>"
        "<h1>📊</h1><h2>쿠팡 광고 보고서를 업로드해주세요</h2>"
        "<p style='color:gray;'>왼쪽 사이드바에서 엑셀 파일을 업로드하면<br>"
        "모든 분석이 자동으로 실행됩니다.</p></div>",
        unsafe_allow_html=True,
    )
    st.stop()


@st.cache_data(show_spinner="데이터 로딩 중...")
def load(file):
    if file.name.endswith(".csv"):
        return pd.read_csv(file, encoding="utf-8-sig")
    try:
        return pd.read_excel(file, engine="openpyxl")
    except Exception:
        return pd.read_excel(file, engine="openpyxl", header=1)


raw = load(uploaded)
col_map = auto_map_columns(raw)
df = raw.rename(columns=col_map)

num_cols = ["노출수", "클릭수", "광고비", "직접전환판매수", "간접전환판매수",
            "직접전환매출액", "간접전환매출액", "총전환판매수", "총전환매출액"]
df = clean_numeric(df, num_cols)

if "총전환판매수" not in df.columns:
    df["총전환판매수"] = df.get("직접전환판매수", 0) + df.get("간접전환판매수", 0)
if "총전환매출액" not in df.columns:
    df["총전환매출액"] = df.get("직접전환매출액", 0) + df.get("간접전환매출액", 0)

df = calc_metrics(df)

if "키워드" in df.columns:
    df["영역구분"] = df["키워드"].apply(
        lambda x: "비검색영역" if pd.isna(x) or str(x).strip() == "" else "검색영역"
    )
else:
    df["영역구분"] = "정보없음"

if "날짜" in df.columns:
    df["날짜"] = pd.to_datetime(df["날짜"], errors="coerce")
st.header("① 전체 성과 요약")

t_cost = df["광고비"].sum()
t_rev = df["총전환매출액"].sum()
t_imp = df["노출수"].sum()
t_click = df["클릭수"].sum()
t_conv = df["총전환판매수"].sum()
t_profit = t_rev - t_cost

o_ctr = safe_div(t_click, t_imp) * 100
o_cpc = safe_div(t_cost, t_click)
o_cvr = safe_div(t_conv, t_click) * 100
o_roas = safe_div(t_rev, t_cost) * 100

c1, c2, c3, c4, c5, c6 = st.columns(6)
c1.metric("💰 총 광고비", f"{format_number(t_cost)}원")
c2.metric("📈 전환매출", f"{format_number(t_rev)}원")
c3.metric("💵 순수익", f"{format_number(t_profit)}원")
c4.metric("🎯 ROAS", f"{o_roas:,.0f}%")
c5.metric("👆 CTR", f"{o_ctr:.2f}%")
c6.metric("💳 CPC", f"{o_cpc:,.0f}원")

if o_roas >= 500:
    st.success(f"ROAS {o_roas:.0f}% — 매우 우수합니다!")
elif o_roas >= target_roas:
    st.info(f"ROAS {o_roas:.0f}% — 목표({target_roas}%) 이상입니다.")
elif o_roas >= 100:
    st.warning(f"ROAS {o_roas:.0f}% — 목표({target_roas}%) 미달. 최적화 필요.")
else:
    st.error(f"ROAS {o_roas:.0f}% — 광고비가 매출보다 큽니다! 점검하세요.")

if "날짜" in df.columns and df["날짜"].notna().sum() > 0:
    st.markdown("---")
    st.header("② 일별 추이 분석")

    daily = (
        df.groupby("날짜")
        .agg({"노출수": "sum", "클릭수": "sum", "광고비": "sum",
              "총전환판매수": "sum", "총전환매출액": "sum"})
        .reset_index().sort_values("날짜")
    )
    daily["ROAS(%)"] = (daily["총전환매출액"] / (daily["광고비"] + 1e-10) * 100).round(0)
    daily.loc[daily["ROAS(%)"] > 1e8, "ROAS(%)"] = 0

    tab_d1, tab_d2 = st.tabs(["광고비 vs 매출", "ROAS 추이"])

    with tab_d1:
        fig1 = go.Figure()
        fig1.add_trace(go.Bar(x=daily["날짜"], y=daily["광고비"], name="광고비", marker_color="#FF6B6B"))
        fig1.add_trace(go.Bar(x=daily["날짜"], y=daily["총전환매출액"], name="전환매출", marker_color="#4ECDC4"))
        fig1.update_layout(barmode="group", height=420)
        st.plotly_chart(fig1, use_container_width=True)

    with tab_d2:
        fig2 = px.line(daily, x="날짜", y="ROAS(%)", markers=True)
        fig2.add_hline(y=target_roas, line_dash="dash", line_color="green",
                       annotation_text=f"목표 {target_roas}%")
        fig2.update_layout(height=380)
        st.plotly_chart(fig2, use_container_width=True)
if "키워드" in df.columns:
    st.markdown("---")
    st.header("③ 키워드 성과 분석")

    search = df[df["영역구분"] == "검색영역"]
    kw = (
        search.groupby("키워드")
        .agg({"노출수": "sum", "클릭수": "sum", "광고비": "sum",
              "총전환판매수": "sum", "총전환매출액": "sum"})
        .reset_index()
    )
    kw = calc_metrics(kw)

    show_cols = ["키워드", "노출수", "클릭수", "광고비", "총전환판매수",
                 "총전환매출액", "CTR(%)", "CPC(원)", "CVR(%)", "ROAS(%)"]

    tab_k1, tab_k2, tab_k3 = st.tabs(["🏆 효율 키워드 TOP", "💸 돈 먹는 키워드", "📊 전체"])

    with tab_k1:
        top = kw[kw["총전환판매수"] > 0].sort_values("ROAS(%)", ascending=False).head(20)
        st.dataframe(top[show_cols], use_container_width=True, hide_index=True)

        if len(top) > 0:
            fig_kw = px.bar(top.head(15), x="ROAS(%)", y="키워드", orientation="h",
                            color="ROAS(%)", color_continuous_scale="Greens")
            fig_kw.add_vline(x=target_roas, line_dash="dash", line_color="red")
            fig_kw.update_layout(height=500, yaxis=dict(autorange="reversed"))
            st.plotly_chart(fig_kw, use_container_width=True)

    with tab_k2:
        drain = kw[(kw["총전환판매수"] == 0) & (kw["광고비"] > 0)].sort_values("광고비", ascending=False)
        waste = drain["광고비"].sum()

        mc1, mc2 = st.columns(2)
        mc1.metric("🚨 낭비 키워드 수", f"{len(drain)}개")
        mc2.metric("💸 낭비 광고비", f"{format_number(waste)}원")

        st.dataframe(
            drain[["키워드", "노출수", "클릭수", "광고비", "CTR(%)", "CPC(원)"]].head(50),
            use_container_width=True, hide_index=True,
        )

        if len(drain) > 0:
            st.download_button(
                "📥 제외 추천 키워드 다운로드",
                data="\n".join(drain["키워드"].tolist()),
                file_name="제외_추천_키워드.txt",
                mime="text/plain",
            )

    with tab_k3:
        st.dataframe(kw[show_cols].sort_values("광고비", ascending=False),
                     use_container_width=True, hide_index=True)

if "상품명" in df.columns:
    st.markdown("---")
    st.header("④ 상품별 성과 분석")

    prod = (
        df.groupby("상품명")
        .agg({"노출수": "sum", "클릭수": "sum", "광고비": "sum",
              "총전환판매수": "sum", "총전환매출액": "sum"})
        .reset_index()
    )
    prod = calc_metrics(prod)

    pa, pb = st.columns(2)
    with pa:
        fig_pie = px.pie(prod, values="광고비", names="상품명", title="상품별 광고비 비중", hole=0.4)
        fig_pie.update_layout(height=420, showlegend=False)
        st.plotly_chart(fig_pie, use_container_width=True)

    with pb:
        fig_bar = px.bar(prod.sort_values("ROAS(%)", ascending=True),
                         x="ROAS(%)", y="상품명", orientation="h",
                         color="ROAS(%)", color_continuous_scale="RdYlGn")
        fig_bar.add_vline(x=target_roas, line_dash="dash", line_color="red")
        fig_bar.update_layout(height=420)
        st.plotly_chart(fig_bar, use_container_width=True)

    p_cols = ["상품명", "노출수", "클릭수", "광고비", "총전환판매수",
              "총전환매출액", "CTR(%)", "CPC(원)", "CVR(%)", "ROAS(%)"]
    st.dataframe(prod[p_cols].sort_values("광고비", ascending=False),
                 use_container_width=True, hide_index=True)
if df["영역구분"].nunique() > 1:
    st.markdown("---")
    st.header("⑤ 검색영역 vs 비검색영역")

    area = (
        df.groupby("영역구분")
        .agg({"노출수": "sum", "클릭수": "sum", "광고비": "sum",
              "총전환판매수": "sum", "총전환매출액": "sum"})
        .reset_index()
    )
    area = calc_metrics(area)
    st.dataframe(area, use_container_width=True, hide_index=True)

if "캠페인명" in df.columns:
    st.markdown("---")
    st.header("⑥ 캠페인별 성과 비교")

    camp = (
        df.groupby("캠페인명")
        .agg({"노출수": "sum", "클릭수": "sum", "광고비": "sum",
              "총전환판매수": "sum", "총전환매출액": "sum"})
        .reset_index()
    )
    camp = calc_metrics(camp)
    st.dataframe(camp.sort_values("광고비", ascending=False),
                 use_container_width=True, hide_index=True)

st.markdown("---")
with st.expander("📄 원본 데이터 전체 보기"):
    st.dataframe(df, use_container_width=True, hide_index=True)

st.markdown("---")
st.header("📥 분석 결과 다운로드")

sheets = {"전체데이터": df}
if "키워드" in df.columns:
    sheets["키워드분석"] = kw.sort_values("광고비", ascending=False)
    sheets["제외추천키워드"] = kw[(kw["총전환판매수"] == 0) & (kw["광고비"] > 0)].sort_values("광고비", ascending=False)
if "상품명" in df.columns:
    sheets["상품별분석"] = prod.sort_values("광고비", ascending=False)

st.download_button(
    label="📊 전체 분석 결과 엑셀 다운로드",
    data=to_excel_bytes(sheets),
    file_name="쿠팡_광고_분석_결과.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)

st.markdown("---")
st.caption("쿠팡 광고 보고서 자동 분석기 v1.0")
