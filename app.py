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
        
        # 빈칸 데이터가 날아가지 않도록 'nan' 문자로 채워줌
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

        # 비검색영역 조건 파악
        kw_str = pivot_df['키워드'].astype(str).str.strip().str.lower()
        non_search_condition = kw_str.isin(['-', 'nan', 'none', ''])
        
        # 전체, 비검색, 검색 영역 데이터 분리
        df_total = pivot_df.sum(numeric_only=True)
        df_non_search = pivot_df[non_search_condition].sum(numeric_only=True)
        df_search = df_total - df_non_search

        # 0으로 나누기 방지를 위한 안전한 나눗셈 함수
        def safe_div(a, b):
            return a / b if b and b > 0 else 0

        # ════════════════════════════════════════════════════════
        # [1단계] 전체 성과 및 검색/비검색 영역 요약 (엑셀 매크로 완벽 구현)
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

        st.write("") # 간격 띄우기
        
        # 2. 매크로 형태의 영역별 상세 표
        st.subheader("🔍 검색 vs 🌐 비검색 상세 비교 (비율 포함)")
        
        # 데이터프레임 조립 (엑셀 상단 3줄과 동일)
        summary_data = {
            '구분': ['총합계', '비검색영역', '검색영역'],
            '노출수': [df_total.get('노출수',0), df_non_search.get('노출수',0), df_search.get('노출수',0)],
            '클릭수': [df_total.get('클릭수',0), df_non_search.get('클릭수',0), df_search.get('클릭수',0)],
            '광고비': [total_ad_spend, df_non_search.get('광고비',0), df_search.get('광고비',0)],
            '총 주문수': [total_orders, df_non_search.get('총 주문수(14일)',0), df_search.get('총 주문수(14일)',0)],
            '총 판매수량': [df_total.get('총 판매수량(14일)',0), df_non_search.get('총 판매수량(14일)',0), df_search.get('총 판매수량(14일)',0)],
            '총 전환매출액': [total_sales, df_non_search.get('총 전환매출액(14일)',0), df_search.get('총 전환매출액(14일)',0)]
        }
        summary_df = pd.DataFrame(summary_data)
        
        # 테이블 전용 CPC, ROAS 및 비중(%) 계산
        summary_df['CPC'] = summary_df.apply(lambda x: round(safe_div(x['광고비'], x['클릭수']), 0), axis=1)
        summary_df['ROAS'] = summary_df.apply(lambda x: round(safe_div(x['총 전환매출액'], x['광고비']) * 100, 2), axis=1)
        summary_df['매출 비중(%)'] = summary_df['총 전환매출액'].apply(lambda x: round(safe_div(x, total_sales) * 100, 1))
        summary_df['광고비 비중(%)'] = summary_df['광고비'].apply(lambda x: round(safe_div(x, total_ad_spend) * 100, 1))
        
        # 컬럼 순서 보기 좋게 정리
        summary_df = summary_df[['구분', '노출수', '클릭수', 'CPC', '광고비', '광고비 비중(%)', '총 주문수', '총 판매수량', '총 전환매출액', '매출 비중(%)', 'ROAS']]

        # 엑셀처럼 주황색 배경 적용 함수
        def highlight_summary(row):
            return ['background-color: #FFC78C; color: black; font-weight: bold'] * len(row)

        # 테이블 스타일 및 숫자 포맷팅
        styled_summary = summary_df.style.apply(highlight_summary, axis=1).format({
            '노출수': '{:,.0f}', '클릭수': '{:,.0f}', 'CPC': '{:,.0f}',
            '광고비': '{:,.0f}', '광고비 비중(%)': '{:,.1f}%', '총 주문수': '{:,.0f}', 
            '총 판매수량': '{:,.0f}', '총 전환매출액': '{:,.0f}', '매출 비중(%)': '{:,.1f}%',
            'ROAS': '{:,.2f}'
        })
        
        # 화면에 출력 (hide_index=True 로 지저분한 인덱스 번호 숨김)
        st.dataframe(styled_summary, use_container_width=True, hide_index=True)

        st.divider()

        # ════════════════════════════════════════════════════════
        # [2단계] 키워드 분석 (블랙홀 / 고비용)
        # ════════════════════════════════════════════════════════
        st.header("2️⃣ 핵심 키워드 점검 (검색 영역 기준)")
        st.markdown("의미 없이 **광고비를 많이 갉아먹는 키워드**와 **클릭당 비용(CPC)이 지나치게 높은 키워드**를 빠르게 찾아냅니다.")
        
        # 검색영역 키워드만 필터링
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
