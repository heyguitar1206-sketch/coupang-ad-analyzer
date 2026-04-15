import streamlit as st
import pandas as pd
import numpy as np

# [디자인] 페이지 기본 설정 및 모던 폰트 적용을 위한 CSS
st.set_page_config(page_title="쿠팡 광고 분석기", layout="wide")

st.markdown("""
    <style>
    /* 전체 폰트 및 스타일 통일 */
    @import url('https://fonts.googleapis.com/css2?family=Noto+Sans+KR:wght@400;700;900&display=swap');
    
    html, body, [class*="css"] {
        font-family: 'Noto Sans KR', sans-serif !important;
        color: #111111 !important;
    }
    
    /* 헤더 스타일 커스텀 (쿠팡 블루) */
    .stHeader h1, .stHeader h2, .stHeader h3 {
        color: #007AFF !important;
        font-weight: 800 !important;
    }
    
    /* 메트릭 카드 디자인 개선 */
    [data-testid="stMetricValue"] {
        font-size: 28px !important;
        font-weight: 900 !important;
        color: #007AFF !important;
    }
    [data-testid="stMetricLabel"] {
        font-size: 16px !important;
        font-weight: 700 !important;
        color: #555555 !important;
    }
    
    /* 표 헤더 색상 및 텍스트 선명도 */
    thead tr th {
        background-color: #F0F7FF !important;
        color: #0056b3 !important;
        font-weight: 700 !important;
    }
    </style>
    """, unsafe_allow_html=True)

st.title("📊 쿠팡 광고보고서 자동 분석기")
st.markdown("쿠팡 윙(Wing) 스타일의 직관적인 인터페이스로 광고 성과를 심층 분석합니다.")

# 파일 업로드 영역
uploaded_file = st.file_uploader("분석할 광고보고서 엑셀 파일을 업로드하세요", type=['xlsx', 'xls'])

if uploaded_file is not None:
    try:
        # --- [로직 유지] 데이터 전처리 ---
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
        # [1단계] 전체 성과 및 영역별 요약 (A컨셉 스타일)
        # ════════════════════════════════════════════════════════
        st.subheader("1️⃣ 전체 성과 및 영역별 요약")
        
        total_ad_spend = df_total.get('광고비', 0)
        total_sales = df_total.get('총 전환매출액(14일)', 0)
        total_roas = safe_div(total_sales, total_ad_spend) * 100
        total_orders = df_total.get('총 주문수(14일)', 0)

        # 요약 메트릭
        col_t1, col_t2, col_t3, col_t4 = st.columns(4)
        col_t1.metric("총 전환매출액", f"{total_sales:,.0f}원")
        col_t2.metric("총 지출 광고비", f"{total_ad_spend:,.0f}원")
        col_t3.metric("전체 평균 ROAS", f"{total_roas:,.2f}%")
        col_t4.metric("총 주문수", f"{total_orders:,.0f}건")

        st.write("") 
        
        # 영역별 상세 비교 테이블
        search_sales_pct = safe_div(df_search.get('총 전환매출액(14일)', 0), total_sales) * 100
        non_search_sales_pct = safe_div(df_non_search.get('총 전환매출액(14일)', 0), total_sales) * 100
        search_roas_val = safe_div(df_search.get('총 전환매출액(14일)', 0), df_search.get('광고비', 0)) * 100
        non_search_roas_val = safe_div(df_non_search.get('총 전환매출액(14일)', 0), df_non_search.get('광고비', 0)) * 100

        summary_data = [
            {
                '구분': '총합계', '노출수': df_total.get('노출수',0), '클릭수': df_total.get('클릭수',0),
                'CPC': safe_div(total_ad_spend, df_total.get('클릭수',0)), '광고비': total_ad_spend,
                '광고비비중': 100.0, '주문수': total_orders, '판매수량': df_total.get('총 판매수량(14일)',0),
                '매출액': total_sales, '매출비중': 100.0, 'ROAS': total_roas
            },
            {
                '구분': '비검색영역', '노출수': df_non_search.get('노출수',0), '클릭수': df_non_search.get('클릭수',0),
                'CPC': safe_div(df_non_search.get('광고비',0), df_non_search.get('클릭수',0)), '광고비': df_non_search.get('광고비',0),
                '광고비비중': safe_div(df_non_search.get('광고비',0), total_ad_spend) * 100, '주문수': df_non_search.get('총 주문수(14일)',0), '판매수량': df_non_search.get('총 판매수량(14일)',0),
                '매출액': df_non_search.get('총 전환매출액(14일)',0), '매출비중': non_search_sales_pct, 'ROAS': non_search_roas_val
            },
            {
                '구분': '검색영역', '노출수': df_search.get('노출수',0), '클릭수': df_search.get('클릭수',0),
                'CPC': safe_div(df_search.get('광고비',0), df_search.get('클릭수',0)), '광고비': df_search.get('광고비',0),
                '광고비비중': safe_div(df_search.get('광고비',0), total_ad_spend) * 100, '주문수': df_search.get('총 주문수(14일)',0), '판매수량': df_search.get('총 판매수량(14일)',0),
                '매출액': df_search.get('총 전환매출액(14일)',0), '매출비중': search_sales_pct, 'ROAS': search_roas_val
            }
        ]
        
        summary_df = pd.DataFrame(summary_data)
        
        # [디자인] 총합계 행 하이라이트 (쿠팡 오렌지톤 가미)
        def highlight_summary(row):
            if row['구분'] == '총합계':
                return ['background-color: #FFF4E5; color: #E65100; font-weight: 800; font-size: 16px; border-bottom: 2px solid #E65100'] * len(row)
            return ['background-color: white; color: #333333; font-weight: 600; font-size: 15px'] * len(row)

        styled_summary = summary_df.style.apply(highlight_summary, axis=1).format({
            '노출수': '{:,.0f}', '클릭수': '{:,.0f}', 'CPC': '{:,.0f}원',
            '광고비': '{:,.0f}원', '광고비비중': '{:,.1f}%', '주문수': '{:,.0f}건', 
            '판매수량': '{:,.0f}개', '매출액': '{:,.0f}원', '매출비중': '{:,.1f}%', 'ROAS': '{:,.2f}%'
        })
        
        st.table(styled_summary)

        # 가이드 코멘트 영역 (쿠팡 블루 테두리)
        with st.container():
            st.markdown("---")
            if total_sales > 0:
                if search_sales_pct >= non_search_sales_pct:
                    st.info(f"**💡 쿠팡 광고 가이드:** 전체 매출의 **{search_sales_pct:.1f}%**가 검색영역에서 발생 중입니다. 효율이 좋은 핵심 키워드를 선별하여 수동 입찰가를 높이는 전략이 유효합니다.")
                else:
                    st.warning(f"**💡 쿠팡 광고 가이드:** 매출의 **{non_search_sales_pct:.1f}%**가 비검색영역에 쏠려 있습니다. 자동 캠페인의 기본 입찰가를 방어적으로 조절하여 광고비 누수를 막으세요.")

        # ════════════════════════════════════════════════════════
        # [2단계] 제외 키워드 추출 (A컨셉 스타일)
        # ════════════════════════════════════════════════════════
        st.subheader("2️⃣ 자동 제외 키워드 추출 (Top 30)")
        
        df_keywords = pivot_df[~non_search_condition].copy()
        top_spend = df_keywords.sort_values(by='광고비', ascending=False).head(30)
        top_cpc = df_keywords[df_keywords['클릭수'] >= 3].sort_values(by='CPC', ascending=False).head(30)

        bad_spend_kw = top_spend[top_spend['총 전환매출액(14일)'] == 0]['키워드'].tolist()
        bad_cpc_kw = top_cpc[top_cpc['총 전환매출액(14일)'] == 0]['키워드'].tolist()
        negative_keywords = list(set(bad_spend_kw + bad_cpc_kw))
        
        # 제외 키워드 박스 (선명한 레드)
        if len(negative_keywords) > 0:
            st.error("❗ 아래 키워드들을 쿠팡 광고센터의 [제외 키워드] 란에 즉시 추가하세요.")
            st.text_area(label="전체 복사 (매출 0원 & 고비용 키워드)", value=", ".join(negative_keywords), height=100)
        
        st.write("") 

        # [디자인] 매출 여부에 따른 선명한 색상 대비
        def highlight_sales_status(row):
            if row['총 전환매출액(14일)'] > 0:
                return ['background-color: #EBF7EE; color: #1E4620; font-weight: 700'] * len(row)
            return ['background-color: #FFF0F0; color: #C62828; font-weight: 700'] * len(row)

        col_kw1, col_kw2 = st.columns(2)
        with col_kw1:
            st.markdown("**💸 광고비 지출 Top 30**")
            st.dataframe(top_spend[['키워드', '광고비', 'ROAS', '총 전환매출액(14일)']].style.apply(highlight_sales_status, axis=1).format({
                '광고비': '{:,.0f}', 'ROAS': '{:,.2f}', '총 전환매출액(14일)': '{:,.0f}'
            }), use_container_width=True, hide_index=True, height=500)

        with col_kw2:
            st.markdown("**📈 평균 CPC Top 30**")
            st.dataframe(top_cpc[['키워드', 'CPC', '클릭수', '광고비', '총 전환매출액(14일)']].style.apply(highlight_sales_status, axis=1).format({
                'CPC': '{:,.0f}', '클릭수': '{:,.0f}', '광고비': '{:,.0f}', '총 전환매출액(14일)': '{:,.0f}'
            }), use_container_width=True, hide_index=True, height=500)

        # ════════════════════════════════════════════════════════
        # [3단계] 키워드별 상세 분석 (A컨셉 스타일)
        # ════════════════════════════════════════════════════════
        st.subheader("3️⃣ 키워드별 상세 분석 전체 시트")
        
        final_df = pivot_df.copy()
        final_df.loc[non_search_condition, '키워드'] = '비검색영역'
        final_df = final_df.rename(columns={'총 주문수(14일)': '주문', '총 판매수량(14일)': '수량', '총 전환매출액(14일)': '매출액'})
        
        def highlight_roas_bold(row):
            color = 'background-color: #FFFDE7; color: #000000; font-weight: 700' if row['ROAS'] > 0 else 'color: #333333'
            return [color] * len(row)

        cols_order = ['키워드', '노출수', '클릭수', 'CPC', '광고비', '주문', '수량', '매출액', 'ROAS']
        final_df = final_df.sort_values(by='매출액', ascending=False)[cols_order]

        st.dataframe(
            final_df.style.apply(highlight_roas_bold, axis=1).format({
                '노출수': '{:,.0f}', '클릭수': '{:,.0f}', 'CPC': '{:,.0f}',
                '광고비': '{:,.0f}', '주문': '{:,.0f}', '수량': '{:,.0f}', '매출액': '{:,.0f}', 'ROAS': '{:,.2f}'
            }), 
            use_container_width=True, 
            height=700, 
            hide_index=True,
            column_config={
                "키워드": st.column_config.TextColumn(width="medium"),
                "주문": st.column_config.NumberColumn(width="small"),
                "수량": st.column_config.NumberColumn(width="small")
            }
        )

    except Exception as e:
        st.error(f"데이터 처리 중 오류가 발생했습니다. (에러: {e})")
