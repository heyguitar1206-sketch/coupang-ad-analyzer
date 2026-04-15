import streamlit as st
import pandas as pd
import numpy as np

# 페이지 기본 설정
st.set_page_config(page_title="쿠팡 광고 분석기", layout="wide")

st.title("📊 쿠팡 광고보고서 자동 분석기 (3단계 심층분석)")
st.markdown("업로드하신 광고보고서를 바탕으로 **전체 요약, 제외 키워드 추출, 전체 상세 데이터**를 단계별로 분석해 드립니다.")

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

        st.subheader("🏆 전체 광고 성과")
        col_t1, col_t2, col_t3, col_t4 = st.columns(4)
        col_t1.metric("총 전환매출액", f"{total_sales:,.0f}원")
        col_t2.metric("총 지출 광고비", f"{total_ad_spend:,.0f}원")
        col_t3.metric("전체 평균 ROAS", f"{total_roas:,.2f}%")
        col_t4.metric("총 주문수", f"{total_orders:,.0f}건")

        st.write("") 
        
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

        st.markdown("### 🤖 데이터 기반 캠페인 개선 가이드")
        if total_sales == 0 and total_ad_spend == 0:
            st.info("광고 데이터가 충분하지 않아 가이드를 생성할 수 없습니다.")
        else:
            if search_sales_pct >= non_search_sales_pct and search_roas_val >= non_search_roas_val:
                st.success(f"**[진단] 전체 매출의 {search_sales_pct:.1f}%가 검색영역에서 발생하며, 효율(ROAS) 역시 우수합니다.**")
                st.markdown("**🎯 액션 플랜:** 비검색 영역(스마트/매출최적화)의 예산 비중을 줄이고, 효율이 좋은 핵심 검색 키워드를 수동 캠페인으로 분리하여 입찰가를 상향하세요.")
            elif search_sales_pct >= non_search_sales_pct and search_roas_val < non_search_roas_val:
                st.warning(f"**[진단] 매출 비중은 검색영역({search_sales_pct:.1f}%)이 크지만, 효율(ROAS)은 비검색영역이 더 높습니다.**")
                st.markdown("**🎯 액션 플랜:** 검색영역에서 뼈대 매출이 나오나 광고비 누수가 심합니다. 2단계의 '제외 키워드'를 자동 캠페인에서 반드시 제외 처리하세요.")
            elif search_sales_pct < non_search_sales_pct and search_roas_val >= non_search_roas_val:
                st.info(f"**[진단] 매출은 비검색영역({non_search_sales_pct:.1f}%)에서 많이 나오나, 실질적인 효율(ROAS)은 검색영역이 뛰어납니다.**")
                st.markdown("**🎯 액션 플랜:** 쿠팡 추천 노출 비중이 높습니다. 자동 캠페인 기본 입찰가를 50~100원 단위로 낮춰 비검색 영역의 전체 ROAS를 방어하세요.")
            else:
                st.info(f"**[진단] 전체 매출의 {non_search_sales_pct:.1f}%가 비검색영역에서 발생하며, 광고 효율(ROAS) 역시 가장 뛰어납니다.**")
                st.markdown("**🎯 액션 플랜:** 쿠팡 추천(연관 상품) 로직을 가장 잘 타는 상품입니다. 억지로 비싼 키워드를 잡으려 하지 마시고 현재의 자동 캠페인 세팅을 유지하세요.")

        st.divider()

        # ════════════════════════════════════════════════════════
        # [2단계] 핵심 키워드 점검 및 제외 키워드 자동 추출 (TOP 30)
        # ════════════════════════════════════════════════════════
        st.header("2️⃣ 자동 제외 키워드 추출 (블랙홀 / 고비용)")
        st.markdown("""
        쿠팡 매출최적화(자동) 광고에서 쓸데없이 돈만 나가는 키워드를 골라냅니다.  
        아래 추출된 **'제외 키워드'를 복사하여 쿠팡 광고센터에 그대로 붙여넣기** 하시면 광고 효율이 즉각적으로 개선됩니다.
        """)
        
        df_keywords = pivot_df[~non_search_condition].copy()

        # TOP 30 데이터 추출
        top_spend = df_keywords.sort_values(by='광고비', ascending=False).head(30)
        top_cpc = df_keywords[df_keywords['클릭수'] >= 3].sort_values(by='CPC', ascending=False).head(30)

        # 💡 [핵심 로직] 제외 키워드 선정 (매출 0원이면서 광고비 TOP 30 이거나 CPC TOP 30 인 경우)
        bad_spend_kw = top_spend[top_spend['총 전환매출액(14일)'] == 0]['키워드'].tolist()
        bad_cpc_kw = top_cpc[top_cpc['총 전환매출액(14일)'] == 0]['키워드'].tolist()
        
        # 중복 제거 및 리스트화
        negative_keywords = list(set(bad_spend_kw + bad_cpc_kw))
        
        # ✂️ 제외 키워드 복사 UI
        st.error("### ✂️ 즉시 적용 가능한 제외 키워드 (매출 0원 & 고비용)")
        if len(negative_keywords) > 0:
            st.markdown("우측 상단의 **복사 버튼(📋)** 을 눌러 쿠팡 광고 캠페인 - [제외 키워드] 란에 붙여넣으세요.")
            # st.code 를 사용하면 기본적으로 복사 버튼이 활성화되어 수강생들이 매우 편하게 쓸 수 있습니다.
            st.code(", ".join(negative_keywords), language='text')
        else:
            st.success("현재 조건(광고비/CPC 상위 30위 내 매출 0원)에 해당하는 악성 제외 키워드가 없습니다! 광고가 아주 효율적으로 돌아가고 있습니다.")

        st.write("") # 간격 띄우기

        col_kw1, col_kw2 = st.columns(2)
        
        with col_kw1:
            st.subheader("💸 광고비 지출 TOP 30")
            st.dataframe(top_spend[['키워드', '광고비', 'ROAS', '총 전환매출액(14일)']].style.format({
                '광고비': '{:,.0f}', 'ROAS': '{:,.2f}', '총 전환매출액(14일)': '{:,.0f}'
            }), use_container_width=True, hide_index=True, height=500)

        with col_kw2:
            st.subheader("📈 평균 CPC TOP 30")
            st.dataframe(top_cpc[['키워드', 'CPC', '클릭수', '광고비', '총 전환매출액(14일)']].style.format({
                'CPC': '{:,.0f}', '클릭수': '{:,.0f}', '광고비': '{:,.0f}', '총 전환매출액(14일)': '{:,.0f}'
            }), use_container_width=True, hide_index=True, height=500)

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
