import streamlit as st
import pandas as pd
import numpy as np

# 페이지 기본 설정
st.set_page_config(page_title="쿠팡 광고 분석기", layout="wide")

st.title("📊 쿠팡 광고보고서 자동 분석기 (3단계 심층분석)")
st.markdown("업로드하신 광고보고서를 바탕으로 **전체 요약, 핵심 키워드, 전체 상세 데이터**를 단계별로 분석해 드립니다.")

# 파일 업로드 영역
uploaded_file = st.file_uploader("광고보고서 엑셀 파일을 업로드하세요", type=['xlsx', 'xls'])

if uploaded_file is not None:
    try:
        # --- 데이터 전처리 ---
        df_raw = pd.read_excel(uploaded_file, sheet_name="Sheet1")
        
        if '키워드' in df_raw.columns:
            df_raw['키워드'] = df_raw['키워드'].fillna('nan')
        
        pivot_df = pd.pivot_table(
            df_raw, 
            index='키워드', 
            values=['노출수', '클릭수', '광고비', '총 주문수(14일)', '총 판매수량(14일)', '총 전환매출액(14일)'], 
            aggfunc='sum'
        ).reset_index()

        pivot_df['CPC'] = np.where(pivot_df['클릭수'] > 0, round(pivot_df['광고비'] / pivot_df['클릭수'], 0), 0)
        pivot_df['ROAS'] = np.where(pivot_df['광고비'] > 0, round((pivot_df['총 전환매출액(14일)'] / pivot_df['광고비']) * 100, 2), 0)

        kw_str = pivot_df['키워드'].astype(str).str.strip().str.lower()
        non_search_condition = kw_str.isin(['-', 'nan', 'none', ''])
        
        df_total = pivot_df.sum(numeric_only=True)
        df_non_search = pivot_df[non_search_condition].sum(numeric_only=True)
        df_search = df_total - df_non_search

        def safe_div(a, b):
            return a / b if b and b > 0 else 0

        # ════════════════════════════════════════════════════════
        # [1단계] 전체 성과 및 검색/비검색 영역 요약 (디자인 대폭 개선)
        # ════════════════════════════════════════════════════════
        st.header("1️⃣ 전체 성과 및 영역별 요약")
        
        total_ad_spend = df_total.get('광고비', 0)
        total_sales = df_total.get('총 전환매출액(14일)', 0)
        total_roas = safe_div(total_sales, total_ad_spend) * 100
        total_orders = df_total.get('총 주문수(14일)', 0)

        # 1. 최상단 전체 요약 대시보드
        st.subheader("🏆 전체 광고 성과")
        col_t1, col_t2, col_t3, col_t4 = st.columns(4)
        col_t1.metric("총 전환매출액", f"{total_sales:,.0f}원")
        col_t2.metric("총 지출 광고비", f"{total_ad_spend:,.0f}원")
        col_t3.metric("전체 평균 ROAS", f"{total_roas:,.2f}%")
        col_t4.metric("총 주문수", f"{total_orders:,.0f}건")

        st.write("") 
        
        # 2. 커스텀 CSS를 적용한 가독성 높은 영역별 상세 표
        st.subheader("🔍 검색 vs 🌐 비검색 상세 비교 (비율 포함)")
        
        # 데이터 계산
        summary_data = [
            {
                '구분': '총합계', 
                '노출수': df_total.get('노출수',0), '클릭수': df_total.get('클릭수',0),
                'CPC': safe_div(total_ad_spend, df_total.get('클릭수',0)), '광고비': total_ad_spend,
                '광고비비중': 100.0, '총주문수': total_orders, '총판매수량': df_total.get('총 판매수량(14일)',0),
                '총전환매출액': total_sales, '매출비중': 100.0, 'ROAS': total_roas
            },
            {
                '구분': '비검색영역', 
                '노출수': df_non_search.get('노출수',0), '클릭수': df_non_search.get('클릭수',0),
                'CPC': safe_div(df_non_search.get('광고비',0), df_non_search.get('클릭수',0)), '광고비': df_non_search.get('광고비',0),
                '광고비비중': safe_div(df_non_search.get('광고비',0), total_ad_spend) * 100, '총주문수': df_non_search.get('총 주문수(14일)',0), '총판매수량': df_non_search.get('총 판매수량(14일)',0),
                '총전환매출액': df_non_search.get('총 전환매출액(14일)',0), '매출비중': safe_div(df_non_search.get('총 전환매출액(14일)',0), total_sales) * 100, 'ROAS': safe_div(df_non_search.get('총 전환매출액(14일)',0), df_non_search.get('광고비',0)) * 100
            },
            {
                '구분': '검색영역', 
                '노출수': df_search.get('노출수',0), '클릭수': df_search.get('클릭수',0),
                'CPC': safe_div(df_search.get('광고비',0), df_search.get('클릭수',0)), '광고비': df_search.get('광고비',0),
                '광고비비중': safe_div(df_search.get('광고비',0), total_ad_spend) * 100, '총주문수': df_search.get('총 주문수(14일)',0), '총판매수량': df_search.get('총 판매수량(14일)',0),
                '총전환매출액': df_search.get('총 전환매출액(14일)',0), '매출비중': safe_div(df_search.get('총 전환매출액(14일)',0), total_sales) * 100, 'ROAS': safe_div(df_search.get('총 전환매출액(14일)',0), df_search.get('광고비',0)) * 100
            }
        ]

        # HTML 구조 생성
        html_table = f"""
        <style>
        .custom-table {{
            width: 100%;
            border-collapse: collapse;
            font-family: 'Malgun Gothic', sans-serif;
            margin-top: 10px;
            color: #1F2937; /* 진하고 선명한 텍스트 색상 */
        }}
        .custom-table th {{
            background-color: #F3F4F6;
            color: #111827;
            font-size: 16px; /* 제목 폰트 크기 확대 */
            font-weight: 800; /* 제목을 두껍게 통일 */
            padding: 14px 10px;
            text-align: right;
            border-bottom: 2px solid #D1D5DB;
        }}
        .custom-table th:first-child {{ text-align: center; }}
        
        .custom-table td {{
            font-size: 15px; /* 일반 텍스트 크기 일치 */
            font-weight: 600; /* 일반 텍스트도 진하게 표시 */
            padding: 12px 10px;
            text-align: right;
            border-bottom: 1px solid #E5E7EB;
        }}
        .custom-table td:first-child {{ text-align: center; font-weight: 800; }}
        
        /* 총합계 행 하이라이트 (주황색) */
        .custom-table tr:nth-child(2) td {{
            background-color: #FFEDD5;
            color: #9A3412;
            font-size: 16px;
        }}
        
        .custom-table tr:hover td {{ background-color: #F9FAFB; }}
        </style>

        <table class="custom-table">
            <tr>
                <th>구분</th><th>노출수</th><th>클릭수</th><th>CPC</th><th>광고비</th>
                <th>광고비 비중</th><th>총 주문수</th><th>총 판매수량</th><th>총 전환매출액</th>
                <th>매출 비중</th><th>ROAS</th>
            </tr>
        """
        
        for row in summary_data:
            html_table += f"""
            <tr>
                <td>{row['구분']}</td>
                <td>{row['노출수']:,.0f}</td>
                <td>{row['클릭수']:,.0f}</td>
                <td>{row['CPC']:,.0f}원</td>
                <td>{row['광고비']:,.0f}원</td>
                <td>{row['광고비비중']:.1f}%</td>
                <td>{row['총주문수']:,.0f}건</td>
                <td>{row['총판매수량']:,.0f}개</td>
                <td>{row['총전환매출액']:,.0f}원</td>
                <td>{row['매출비중']:.1f}%</td>
                <td>{row['ROAS']:,.2f}%</td>
            </tr>
            """
        html_table += "</table>"
        
        # HTML 렌더링
        st.markdown(html_table, unsafe_allow_html=True)

        st.divider()

        # ════════════════════════════════════════════════════════
        # [2단계] 키워드 분석 (블랙홀 / 고비용)
        # ════════════════════════════════════════════════════════
        st.header("2️⃣ 핵심 키워드 점검 (검색 영역 기준)")
        st.markdown("의미 없이 **광고비를 많이 갉아먹는 키워드**와 **클릭당 비용(CPC)이 지나치게 높은 키워드**를 빠르게 찾아냅니다.")
        
        df_keywords = pivot_df[~non_search_condition].copy()

        col_kw1, col_kw2 = st.columns(2)
        
        with col_kw1:
            st.subheader("💸 광고비 지출 TOP 10")
            top_spend = df_keywords.sort_values(by='광고비', ascending=False).head(10)
            
            st.dataframe(top_spend[['키워드', '광고비', 'ROAS', '총 전환매출액(14일)']].style.format({
                '광고비': '{:,.0f}', 'ROAS': '{:,.2f}', '총 전환매출액(14일)': '{:,.0f}'
            }), use_container_width=True, hide_index=True)

        with col_kw2:
            st.subheader("📈 평균 CPC TOP 10")
            top_cpc = df_keywords[df_keywords['클릭수'] >= 3].sort_values(by='CPC', ascending=False).head(10)
            
            st.dataframe(top_cpc[['키워드', 'CPC', '클릭수', '광고비']].style.format({
                'CPC': '{:,.0f}', '클릭수': '{:,.0f}', '광고비': '{:,.0f}'
            }), use_container_width=True, hide_index=True)

        st.divider()

        # ════════════════════════════════════════════════════════
        # [3단계] 키워드별 상세 분석 (전체)
        # ════════════════════════════════════════════════════════
        st.header("3️⃣ 키워드별 상세 분석 표")
        
        def highlight_roas(row):
            color = 'background-color: #FFFF99; color: black' if row['ROAS'] > 0 else ''
            return [color] * len(row)

        cols_order = ['키워드', '노출수', '클릭수', 'CPC', '광고비', '총 주문수(14일)', '총 판매수량(14일)', '총 전환매출액(14일)', 'ROAS']
        final_df = pivot_df.sort_values(by='총 전환매출액(14일)', ascending=False)[cols_order]

        styled_df = final_df.style.apply(highlight_roas, axis=1).format({
            '노출수': '{:,.0f}', '클릭수': '{:,.0f}', 'CPC': '{:,.0f}',
            '광고비': '{:,.0f}', '총 주문수(14일)': '{:,.0f}', 
            '총 판매수량(14일)': '{:,.0f}', '총 전환매출액(14일)': '{:,.0f}', 
            'ROAS': '{:,.2f}'
        })

        st.dataframe(styled_df, use_container_width=True, height=600, hide_index=True)

    except Exception as e:
        st.error(f"데이터를 처리하는 중 오류가 발생했습니다. (상세 에러: {e})")
