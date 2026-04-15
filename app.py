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
        # [1단계] 전체 성과 및 검색/비검색 영역 요약
        # ════════════════════════════════════════════════════════
        st.header("1️⃣ 전체 성과 및 영역별 요약")
        
        total_ad_spend = df_total.get('광고비', 0)
        total_sales = df_total.get('총 전환매출액(14일)', 0)
        total_roas = safe_div(total_sales, total_ad_spend) * 100
        total_orders = df_total.get('총 주문수(14일)', 0)

        # 최상단 대시보드
        st.subheader("🏆 전체 광고 성과")
        col_t1, col_t2, col_t3, col_t4 = st.columns(4)
        col_t1.metric("총 전환매출액", f"{total_sales:,.0f}원")
        col_t2.metric("총 지출 광고비", f"{total_ad_spend:,.0f}원")
        col_t3.metric("전체 평균 ROAS", f"{total_roas:,.2f}%")
        col_t4.metric("총 주문수", f"{total_orders:,.0f}건")

        st.write("") 
        
        # 정적 테이블 요약
        st.subheader("🔍 검색 vs 🌐 비검색 상세 비교")
        
        search_sales_pct = safe_div(df_search.get('총 전환매출액(14일)', 0), total_sales) * 100
        non_search_sales_pct = safe_div(df_non_search.get('총 전환매출액(14일)', 0), total_sales) * 100
        search_roas_val = safe_div(df_search.get('총 전환매출액(14일)', 0), df_search.get('광고비', 0)) * 100
        non_search_roas_val = safe_div(df_non_search.get('총 전환매출액(14일)', 0), df_non_search.get('광고비', 0)) * 100

        summary_data = [
            {
                '구분': '총합계', 
                '노출수': df_total.get('노출수',0), '클릭수': df_total.get('클릭수',0),
                'CPC': safe_div(total_ad_spend, df_total.get('클릭수',0)), '광고비': total_ad_spend,
                '광고비 비중': 100.0, '총 주문수': total_orders, '총 판매수량': df_total.get('총 판매수량(14일)',0),
                '총 전환매출액': total_sales, '매출 비중': 100.0, 'ROAS': total_roas
            },
            {
                '구분': '비검색영역', 
                '노출수': df_non_search.get('노출수',0), '클릭수': df_non_search.get('클릭수',0),
                'CPC': safe_div(df_non_search.get('광고비',0), df_non_search.get('클릭수',0)), '광고비': df_non_search.get('광고비',0),
                '광고비 비중': safe_div(df_non_search.get('광고비',0), total_ad_spend) * 100, '총 주문수': df_non_search.get('총 주문수(14일)',0), '총 판매수량': df_non_search.get('총 판매수량(14일)',0),
                '총 전환매출액': df_non_search.get('총 전환매출액(14일)',0), '매출 비중': non_search_sales_pct, 'ROAS': non_search_roas_val
            },
            {
                '구분': '검색영역', 
                '노출수': df_search.get('노출수',0), '클릭수': df_search.get('클릭수',0),
                'CPC': safe_div(df_search.get('광고비',0), df_search.get('클릭수',0)), '광고비': df_search.get('광고비',0),
                '광고비 비중': safe_div(df_search.get('광고비',0), total_ad_spend) * 100, '총 주문수': df_search.get('총 주문수(14일)',0), '총 판매수량': df_search.get('총 판매수량(14일)',0),
                '총 전환매출액': df_search.get('총 전환매출액(14일)',0), '매출 비중': search_sales_pct, 'ROAS': search_roas_val
            }
        ]
        
        summary_df = pd.DataFrame(summary_data)
        
        def highlight_summary(row):
            if row['구분'] == '총합계':
                return ['background-color: #FFEDD5; color: #9A3412; font-weight: bold; font-size: 15px'] * len(row)
            else:
                return ['font-weight: normal; font-size: 15px'] * len(row)

        styled_summary = summary_df.style.apply(highlight_summary, axis=1).format({
            '노출수': '{:,.0f}', '클릭수': '{:,.0f}', 'CPC': '{:,.0f}원',
            '광고비': '{:,.0f}원', '광고비 비중': '{:,.1f}%', '총 주문수': '{:,.0f}건', 
            '총 판매수량': '{:,.0f}개', '총 전환매출액': '{:,.0f}원', '매출 비중': '{:,.1f}%',
            'ROAS': '{:,.2f}%'
        })
        
        st.table(styled_summary)

        # 💡 [신규 추가] 데이터 기반 자동 코멘트 생성
        st.markdown("### 🤖 데이터 기반 캠페인 개선 가이드")
        
        if total_sales == 0 and total_ad_spend == 0:
            st.info("광고 데이터가 충분하지 않아 가이드를 생성할 수 없습니다.")
        else:
            if search_sales_pct >= non_search_sales_pct and search_roas_val >= non_search_roas_val:
                st.success(f"**[진단] 전체 매출의 {search_sales_pct:.1f}%가 검색영역에서 발생하며, 효율(ROAS) 역시 비검색영역보다 우수합니다.**")
                st.markdown("""
                **🎯 액션 플랜: 검색영역 집중 및 수동 캠페인 극대화**
                * 비검색 영역으로 소진되는 광고비(스마트 캠페인, 매출최적화 등)의 비중을 줄이거나 자동 캠페인 입찰가를 낮추세요.
                * 아래 2단계, 3단계 분석에서 발견된 **효율이 좋은 핵심 키워드를 수동 캠페인으로 분리하여 입찰가를 상향** 조절하세요.
                * 비검색 영역에서 쓸데없이 돈이 나가는 것을 막기 위해 제외 키워드를 적극 설정하세요.
                """)
                
            elif search_sales_pct >= non_search_sales_pct and search_roas_val < non_search_roas_val:
                st.warning(f"**[진단] 매출 비중은 검색영역({search_sales_pct:.1f}%)이 더 크지만, 광고 효율(ROAS)은 비검색영역이 더 높습니다.**")
                st.markdown("""
                **🎯 액션 플랜: 검색영역 단가 최적화 및 비검색 유지**
                * 검색영역에서 뼈대 역할을 하는 매출이 나오고 있으나, 광고비 누수가 발생하고 있습니다.
                * 아래 2단계 분석을 통해 **광고비만 먹고 전환이 안 되는(블랙홀) 키워드를 찾아 입찰가를 대폭 낮추거나 제외**하세요.
                * 비검색 영역(추천 상품 등)에서 효율적으로 판매가 일어나고 있으므로, 현재의 자동 캠페인 세팅은 당분간 유지해 주세요.
                """)
                
            elif search_sales_pct < non_search_sales_pct and search_roas_val >= non_search_roas_val:
                st.info(f"**[진단] 전체 매출의 {non_search_sales_pct:.1f}%가 비검색영역에서 발생하지만, 실질적인 효율(ROAS)은 검색영역이 뛰어납니다.**")
                st.markdown("""
                **🎯 액션 플랜: 검색 키워드 발굴 및 자동 캠페인 단가 조절**
                * 쿠팡의 추천 로직에 의해 비검색 영역에서 많이 팔리고 있으나 단가가 비싼 편입니다. 
                * 자동 캠페인의 기본 입찰가를 50원~100원 단위로 조금씩 낮춰 비검색 영역의 전체 ROAS를 끌어올려 보세요.
                * 효율이 좋은 검색영역의 파이를 키우기 위해, 세부/중소형 키워드들을 새롭게 소싱하여 수동으로 추가해 보세요.
                """)
                
            else:
                st.info(f"**[진단] 전체 매출의 {non_search_sales_pct:.1f}%가 비검색영역에서 발생하며, 광고 효율(ROAS) 역시 가장 뛰어납니다.**")
                st.markdown("""
                **🎯 액션 플랜: 쿠팡 추천 알고리즘 적극 활용**
                * 이 상품은 키워드 검색보다는 타 상품 페이지 하단의 '연관 상품'으로 노출되었을 때 가장 잘 팔리는 상품입니다.
                * 억지로 검색 키워드 입찰가를 높여서 싸우지 마시고, 현재 잘 돌아가는 자동 캠페인을 유지하세요.
                * 검색 영역의 경우, 지나치게 비싼 키워드가 있다면 과감하게 입찰가를 낮추거나 끄는 것이 이득입니다.
                """)

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
