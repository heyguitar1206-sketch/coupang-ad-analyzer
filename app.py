import streamlit as st
import pandas as pd
import numpy as np

# 페이지 기본 설정
st.set_page_config(page_title="쿠팡 광고 분석기", layout="wide")

st.title("📊 쿠팡 광고보고서 자동 분석기")
st.markdown("쿠팡에서 다운받은 광고보고서 엑셀 파일을 업로드하면 핵심 지표를 자동으로 분석해 드립니다.")

# 파일 업로드 영역
uploaded_file = st.file_uploader("광고보고서 엑셀 파일을 업로드하세요", type=['xlsx', 'xls'])

if uploaded_file is not None:
    try:
        # 1. 데이터 불러오기 (Sheet1 기준)
        df_raw = pd.read_excel(uploaded_file, sheet_name="Sheet1")
        
        # 2. 피벗 테이블 생성
        pivot_df = pd.pivot_table(
            df_raw, 
            index='키워드', 
            values=['노출수', '클릭수', '광고비', '총 주문수(14일)', '총 판매수량(14일)', '총 전환매출액(14일)'], 
            aggfunc='sum'
        ).reset_index()

        # 3. 요약 데이터 계산 (비검색영역 분리)
        non_search_condition = pivot_df['키워드'].isin(['-', 'nan']) | pivot_df['키워드'].isna()
        df_non_search = pivot_df[non_search_condition].sum(numeric_only=True)
        df_total = pivot_df.sum(numeric_only=True)
        df_search = df_total - df_non_search

        # 4. CPC 및 ROAS 계산
        pivot_df['CPC'] = np.where(pivot_df['클릭수'] > 0, round(pivot_df['광고비'] / pivot_df['클릭수'], 1), 0)
        pivot_df['ROAS'] = np.where(pivot_df['광고비'] > 0, round((pivot_df['총 전환매출액(14일)'] / pivot_df['광고비']) * 100, 2), 0)

        # 5. 정렬 (매출 기준 내림차순)
        final_df = pivot_df.sort_values(by='총 전환매출액(14일)', ascending=False)

        # --- 화면 출력 ---
        
        # KPI 요약 대시보드
        st.subheader("💡 전체 광고 성과 요약")
        col1, col2, col3, col4 = st.columns(4)
        
        total_ad_spend = df_total.get('광고비', 0)
        total_sales = df_total.get('총 전환매출액(14일)', 0)
        total_roas = round((total_sales / total_ad_spend) * 100, 2) if total_ad_spend > 0 else 0
        
        col1.metric("총 광고비", f"{total_ad_spend:,.0f}원")
        col2.metric("총 전환매출액", f"{total_sales:,.0f}원")
        col3.metric("전체 ROAS", f"{total_roas}%")
        col4.metric("총 주문수", f"{df_total.get('총 주문수(14일)', 0):,.0f}건")

        st.divider()

        # 상세 데이터 테이블
        st.subheader("🔍 키워드별 상세 분석 (ROAS 발생 키워드 하이라이트)")
        
        # 엑셀처럼 노란색 하이라이트 적용 함수
        def highlight_roas(row):
            color = 'background-color: #FFFF99; color: black' if row['ROAS'] > 0 else ''
            return [color] * len(row)

        # 컬럼 순서 보기 좋게 정리
        cols_order = ['키워드', '노출수', '클릭수', 'CPC', '광고비', '총 주문수(14일)', '총 판매수량(14일)', '총 전환매출액(14일)', 'ROAS']
        final_df = final_df[cols_order]

        # 숫자 포맷팅 (천단위 콤마 등)
        styled_df = final_df.style.apply(highlight_roas, axis=1).format({
            '노출수': '{:,.0f}', '클릭수': '{:,.0f}', 'CPC': '{:,.0f}',
            '광고비': '{:,.0f}', '총 주문수(14일)': '{:,.0f}', 
            '총 판매수량(14일)': '{:,.0f}', '총 전환매출액(14일)': '{:,.0f}', 
            'ROAS': '{:,.2f}'
        })

        # 화면에 테이블 출력
        st.dataframe(styled_df, use_container_width=True, height=600)

    except Exception as e:
        st.error(f"데이터를 처리하는 중 오류가 발생했습니다. 쿠팡 원본 파일이 맞는지 확인해주세요. (상세 에러: {e})")
