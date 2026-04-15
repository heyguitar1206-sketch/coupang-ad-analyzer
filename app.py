import streamlit as st
import pandas as pd
import numpy as np

# 페이지 기본 설정
st.set_page_config(page_title="쿠팡 광고 분석기", layout="wide")

st.title("📊 쿠팡 광고보고서 자동 분석기 (3단계 심층분석)")
st.markdown("업로드하신 광고보고서를 바탕으로 **영역별 효율, 핵심 키워드, 전체 상세 데이터**를 단계별로 분석해 드립니다.")

# 파일 업로드 영역
uploaded_file = st.file_uploader("광고보고서 엑셀 파일을 업로드하세요", type=['xlsx', 'xls'])

if uploaded_file is not None:
    try:
        # --- 데이터 전처리 ---
        df_raw = pd.read_excel(uploaded_file, sheet_name="Sheet1")
        
        # [핵심 수정] 피벗 테이블 생성 시 빈칸 데이터가 날아가지 않도록 'nan' 문자로 채워줌
        if '키워드' in df_raw.columns:
            df_raw['키워드'] = df_raw['키워드'].fillna('nan')
        
        # 피벗 테이블 생성
        pivot_df = pd.pivot_table(
            df_raw, 
            index='키워드', 
            values=['노출수', '클릭수', '광고비', '총 주문수(14일)', '총 판매수량(14일)', '총 전환매출액(14일)'], 
            aggfunc='sum'
        ).reset_index()

        # 개별 키워드의 CPC, ROAS 계산
        pivot_df['CPC'] = np.where(pivot_df['클릭수'] > 0, round(pivot_df['광고비'] / pivot_df['클릭수'], 0), 0)
        pivot_df['ROAS'] = np.where(pivot_df['광고비'] > 0, round((pivot_df['총 전환매출액(14일)'] / pivot_df['광고비']) * 100, 2), 0)

        # ════════════════════════════════════════════════════════
        # [수정] 비검색영역 조건 파악 (매크로 로직 완벽 적용)
        # ════════════════════════════════════════════════════════
        # 키워드를 문자열로 바꾸고, 소문자로 변경 후, 양옆 공백을 제거하여 비교
        kw_str = pivot_df['키워드'].astype(str).str.strip().str.lower()
        non_search_condition = kw_str.isin(['-', 'nan', 'none', ''])
        
        # 전체, 비검색, 검색 영역 데이터 분리 (합계용)
        df_total = pivot_df.sum(numeric_only=True)
        df_non_search = pivot_df[non_search_condition].sum(numeric_only=True)
        
        # 검색영역 = 총합계 - 비검색영역 (기존 매크로와 동일한 수식 적용)
        df_search = df_total - df_non_search

        # 0으로 나누기 방지를 위한 안전한 나눗셈 함수
        def safe_div(a, b):
            return a / b if b and b > 0 else 0

        # ════════════════════════════════════════════════════════
        # [1단계] 검색영역 vs 비검색영역 분석
        # ════════════════════════════════════════════════════════
        st.header("1️⃣ 검색 vs 비검색 영역 효율 분석")
        st.markdown("전체 지출 광고비와 발생 매출이 어느 영역에서 주로 일어났는지 비중과 효율을 확인합니다.")
        
        total_ad_spend = df_total.get('광고비', 0)
        total_sales = df_total.get('총 전환매출액(14일)', 0)

        # 비중 계산 (%)
        search_spend_ratio = safe_div(df_search.get('광고비', 0), total_ad_spend) * 100
        non_search_spend_ratio = safe_div(df_non_search.get('광고비', 0), total_ad_spend) * 100
        search_sales_ratio = safe_div(df_search.get('총 전환매출액(14일)', 0), total_sales) * 100
        non_search_sales_ratio = safe_div(df_non_search.get('총 전환매출액(14일)', 0), total_sales) * 100

        # 효율(ROAS, CPC) 계산
        search_roas = safe_div(df_search.get('총 전환매출액(14일)', 0), df_search.get('광고비', 0)) * 100
        non_search_roas = safe_div(df_non_search.get('총 전환매출액(14일)', 0), df_non_search.get('광고비', 0)) * 100
        search_cpc = safe_div(df_search.get('광고비', 0), df_search.get('클릭수', 0))
        non_search_cpc = safe_div(df_non_search.get('광고비', 0), df_non_search.get('클릭수', 0))

        # 화면 좌우 분할 출력
        col1, col2 = st.columns(2)
        with col1:
            st.info("#### 🔍 검색 영역 (고객 직접 검색)")
            st.metric("총 지출 광고비", f"{df_search.get('광고비', 0):,.0f}원", f"전체의 {search_spend_ratio:.1f}%", delta_color="off")
            st.metric("총 발생 매출액", f"{df_search.get('총 전환매출액(14일)', 0):,.0f}원", f"전체의 {search_sales_ratio:.1f}%", delta_color="off")
            st.metric("평균 ROAS (효율)", f"{search_roas:,.2f}%")
            st.metric("평균 CPC (클릭당비용)", f"{search_cpc:,.0f}원")

        with col2:
            st.warning("#### 🌐 비검색 영역 (쿠팡 자동 노출)")
            st.metric("총 지출 광고비", f"{df_non_search.get('광고비', 0):,.0f}원", f"전체의 {non_search_spend_ratio:.1f}%", delta_color="off")
            st.metric("총 발생 매출액", f"{df_non_search.get('총 전환매출액(14일)', 0):,.0f}원", f"전체의 {non_search_sales_ratio:.1f}%", delta_color="off")
            st.metric("평균 ROAS (효율)", f"{non_search_roas:,.2f}%")
            st.metric("평균 CPC (클릭당비용)", f"{non_search_cpc:,.0f}원")

        st.divider()

        # ════════════════════════════════════════════════════════
        # [2단계] 키워드 분석 (블랙홀 / 고비용)
        # ════════════════════════════════════════════════════════
        st.header("2️⃣ 핵심 키워드 점검 (검색 영역 기준)")
        st.markdown("의미 없이 **광고비를 많이 갉아먹는 키워드**와 **클릭당 비용(CPC)이 지나치게 높은 키워드**를 빠르게 찾아냅니다.")
        
        # 검색영역 키워드만 필터링 (비검색영역 제외)
        df_keywords = pivot_df[~non_search_condition].copy()

        col_kw1, col_kw2 = st.columns(2)
        
        with col_kw1:
            st.subheader("💸 광고비 지출 TOP 10")
            # 광고비 기준 내림차순 정렬
            top_spend = df_keywords.sort_values(by='광고비', ascending=False).head(10)
            
            st.dataframe(top_spend[['키워드', '광고비', 'ROAS', '총 전환매출액(14일)']].style.format({
                '광고비': '{:,.0f}', 'ROAS': '{:,.2f}', '총 전환매출액(14일)': '{:,.0f}'
            }), use_container_width=True)

        with col_kw2:
            st.subheader("📈 평균 CPC TOP 10")
            # 클릭수가 3회 이상인 것들 중에서 CPC가 가장 높은 순으로 정렬 (노이즈 제거)
            top_cpc = df_keywords[df_keywords['클릭수'] >= 3].sort_values(by='CPC', ascending=False).head(10)
            
            st.dataframe(top_cpc[['키워드', 'CPC', '클릭수', '광고비']].style.format({
                'CPC': '{:,.0f}', '클릭수': '{:,.0f}', '광고비': '{:,.0f}'
            }), use_container_width=True)

        st.divider()

        # ════════════════════════════════════════════════════════
        # [3단계] 키워드별 상세 분석 (전체)
        # ════════════════════════════════════════════════════════
        st.header("3️⃣ 키워드별 상세 분석 표")
        
        # 엑셀처럼 노란색 하이라이트 적용 함수
        def highlight_roas(row):
            color = 'background-color: #FFFF99; color: black' if row['ROAS'] > 0 else ''
            return [color] * len(row)

        # 전체 데이터를 매출액 기준 내림차순 정렬
        cols_order = ['키워드', '노출수', '클릭수', 'CPC', '광고비', '총 주문수(14일)', '총 판매수량(14일)', '총 전환매출액(14일)', 'ROAS']
        final_df = pivot_df.sort_values(by='총 전환매출액(14일)', ascending=False)[cols_order]

        # 숫자 포맷팅 및 색상 적용
        styled_df = final_df.style.apply(highlight_roas, axis=1).format({
            '노출수': '{:,.0f}', '클릭수': '{:,.0f}', 'CPC': '{:,.0f}',
            '광고비': '{:,.0f}', '총 주문수(14일)': '{:,.0f}', 
            '총 판매수량(14일)': '{:,.0f}', '총 전환매출액(14일)': '{:,.0f}', 
            'ROAS': '{:,.2f}'
        })

        st.dataframe(styled_df, use_container_width=True, height=600)

    except Exception as e:
        st.error(f"데이터를 처리하는 중 오류가 발생했습니다. (상세 에러: {e})")
