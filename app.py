import streamlit as st
import pandas as pd
import numpy as np
import io  

# [디자인] 페이지 기본 설정
st.set_page_config(page_title="쿠팡 광고 분석기", layout="wide")

st.markdown("""
    <style>
    /* 콘텐츠 가로 길이를 전체 화면의 70%로 강제 제한하고 중앙 정렬 */
    .block-container {
        max-width: 70% !important;
        padding-top: 3rem !important;
        padding-bottom: 5rem !important;
    }
    
    /* 세련된 '프리텐다드(Pretendard)' 폰트 임포트 */
    @import url('https://cdn.jsdelivr.net/gh/orioncactus/pretendard@v1.3.9/dist/web/static/pretendard.min.css');
    
    /* 💡 [버그 해결 1] 아이콘을 깨뜨리는 전체 선택자(*) 사용을 중지하고 기본 태그만 타겟팅 */
    html, body, div, p, span, text {
        font-family: 'Pretendard', -apple-system, BlinkMacSystemFont, system-ui, Roboto, sans-serif;
    }
    
    /* 💡 [버그 해결 2] Streamlit 고유의 아이콘 폰트(upload 등)가 글자로 깨지는 현상을 완벽 방어 */
    i, .stIcon, .material-symbols-rounded, .material-icons {
        font-family: 'Material Symbols Rounded', 'Material Icons' !important;
    }
    
    /* 본문(마크다운)에만 줄간격 안전하게 적용 */
    .stMarkdown p, .stMarkdown li {
        color: #1f2937 !important;
        line-height: 1.8 !important;
        font-size: 15px !important;
    }
    
    /* 파일 업로드 박스 여백 꼬임 방지 */
    div[data-testid="stFileUploadDropzone"] * {
        line-height: initial !important;
        letter-spacing: normal !important;
        margin: 0 !important;
        padding: 0 !important;
    }
    
    /* 제목 디자인 */
    h1, h2, h3, .stHeader h1, .stHeader h2, .stHeader h3 {
        color: #2563EB !important; 
        font-weight: 700 !important; 
        letter-spacing: -0.5px !important; 
        line-height: 1.4 !important;
        font-family: 'Pretendard', sans-serif !important;
    }
    
    h1 { padding-bottom: 1rem !important; }
    h2 { margin-top: 3.5rem !important; margin-bottom: 1.5rem !important; }
    h3 { margin-top: 2rem !important; margin-bottom: 1.2rem !important; }
    
    /* 메트릭 카드 디자인 */
    [data-testid="stMetricValue"] {
        font-size: 28px !important;
        font-weight: 800 !important;
        color: #2563EB !important;
        letter-spacing: -0.5px !important;
        font-family: 'Pretendard', sans-serif !important;
    }
    [data-testid="stMetricLabel"] * {
        font-size: 16px !important;
        font-weight: 700 !important;
        color: #4B5563 !important;
    }
    
    /* 요약 표(st.table) 셀 정렬 */
    table.stTable td:first-child { text-align: center !important; }
    table.stTable td:not(:first-child) { text-align: right !important; }
    
    /* 모든 표의 헤더(첫 줄)를 강제로 가운데 정렬 및 15px 통일 */
    thead tr th, table.stTable th {
        background-color: #F3F6FF !important; 
        color: #1D4ED8 !important;
        font-weight: 700 !important;
        font-size: 15px !important;
        padding: 12px 10px !important; 
        text-align: center !important;
    }
    
    tbody tr td { padding: 10px 10px !important; }
    </style>
    """, unsafe_allow_html=True)

st.title("📊 쿠팡 광고보고서 자동 분석기")
st.markdown("쿠팡 윙(Wing) 스타일의 직관적인 인터페이스로 광고 성과를 심층 분석합니다.")
st.write("<br>", unsafe_allow_html=True) 

uploaded_file = st.file_uploader("분석할 광고보고서 엑셀 파일을 업로드하세요", type=['xlsx', 'xls'])

if uploaded_file is not None:
    try:
        # --- 데이터 전처리 ---
        df_raw = pd.read_excel(uploaded_file, sheet_name="Sheet1")
        if '키워드' in df_raw.columns:
            df_raw['키워드'] = df_raw['키워드'].fillna('nan')
            
        ad_type_detected = "매출최적화" 
        for col in df_raw.columns:
            if '유형' in col.replace(" ", "") or '방식' in col.replace(" ", "") or '캠페인' in col.replace(" ", ""):
                types = df_raw[col].astype(str).unique()
                if any('수동' in t for t in types) and not any('매출' in t for t in types):
                    ad_type_detected = "수동성과형"
                elif any('매출' in t for t in types):
                    ad_type_detected = "매출최적화"
                break
        
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
        # [1단계] 전체 성과 및 영역별 요약
        # ════════════════════════════════════════════════════════
        st.header("1️⃣ 전체 성과 및 영역별 요약")
        
        total_ad_spend = df_total.get('광고비', 0)
        total_sales = df_total.get('총 전환매출액(14일)', 0)
        total_roas = safe_div(total_sales, total_ad_spend) * 100
        total_orders = df_total.get('총 주문수(14일)', 0)

        col_t1, col_t2, col_t3, col_t4 = st.columns(4)
        col_t1.metric("총 전환매출액", f"{total_sales:,.0f}원")
        col_t2.metric("총 지출 광고비", f"{total_ad_spend:,.0f}원")
        col_t3.metric("전체 평균 ROAS", f"{total_roas:,.2f}%")
        col_t4.metric("총 주문수", f"{total_orders:,.0f}건")

        st.write("<br>", unsafe_allow_html=True) 
        
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
        
        def highlight_summary(row):
            if row['구분'] == '총합계':
                return ['background-color: #FFF4E5; color: #EA580C; font-weight: 700; font-size: 16px; border-bottom: 2px solid #EA580C'] * len(row)
            return ['background-color: white; color: #374151; font-weight: 500; font-size: 15px'] * len(row)

        styled_summary = summary_df.style.apply(highlight_summary, axis=1)\
            .set_properties(subset=['구분'], **{'text-align': 'center'})\
            .set_properties(subset=['노출수', '클릭수', 'CPC', '광고비', '광고비비중', '주문수', '판매수량', '매출액', '매출비중', 'ROAS'], **{'text-align': 'right'})\
            .set_table_styles([dict(selector='th', props=[('text-align', 'center')])])\
            .format({
                '노출수': '{:,.0f}', '클릭수': '{:,.0f}', 'CPC': '{:,.0f}원',
                '광고비': '{:,.0f}원', '광고비비중': '{:,.1f}%', '주문수': '{:,.0f}건', 
                '판매수량': '{:,.0f}개', '매출액': '{:,.0f}원', '매출비중': '{:,.1f}%', 'ROAS': '{:,.2f}%'
            })
        
        st.table(styled_summary)

        with st.container():
            st.write("<br>", unsafe_allow_html=True)
            if total_sales > 0:
                st.markdown(f"#### 💡 끝장캐리 실전 가이드 (현재 주력 광고: **{ad_type_detected}**)")
                
                if ad_type_detected == "매출최적화":
                    if non_search_sales_pct >= search_sales_pct and non_search_roas_val >= search_roas_val:
                        st.success(f"**[진단] 전형적인 매출최적화 성공 패턴! 비검색영역 매출({non_search_sales_pct:.1f}%)과 효율이 모두 우수합니다.**")
                        st.markdown("""
                        * **액션 플랜:** 현재 매최 광고가 제품과 찰떡궁합으로 잘 돌고 있습니다. 볼륨을 키우기 위해 **목표수익률(ROAS) 세팅값을 평소보다 50% ~ 100% 정도 상향**시켜 마진율 극대화를 시도해 보세요.
                        * **단가 세팅:** 아래 2단계에서 추출된 '제외 키워드'를 쿠팡에 꾸준히 입력하여 검색영역에서 새는 돈만 막아주면 됩니다.
                        """)
                    elif search_sales_pct > non_search_sales_pct and search_roas_val > non_search_roas_val:
                        st.info(f"**[진단] 매최 광고임에도 검색영역 성과({search_sales_pct:.1f}%)가 두드러지게 좋습니다.**")
                        st.markdown("""
                        * **액션 플랜 (투트랙 전략):** 검색을 통한 유입과 전환이 아주 훌륭합니다. 이때 효율만 보고 매최를 끄면 기존 매출이 박살 납니다! **기존 매출최적화 광고는 볼륨 방어용으로 그대로 켜두고, 성과 좋은 핵심 키워드만 따로 빼서 '수동성과형 광고'를 새롭게 추가 개설(투트랙 테스트)** 하세요.
                        * **단가 세팅:** 3단계 표에서 효율이 검증된 키워드만 수동으로 세팅하고, 추후 두 캠페인의 데이터를 비교 분석하여 비중 조절하세요.
                        """)
                    elif non_search_roas_val < (total_roas * 0.5) or non_search_sales_pct < 20:
                        st.warning(f"**[진단] 비검색영역의 효율이 심각하게 부진하며 돈만 까먹고 있습니다.**")
                        st.markdown("""
                        * **액션 플랜 (수동 갈아타기):** 매최 광고의 알고리즘이 비검색 영역에서 타겟을 전혀 찾지 못하고 있습니다. 이럴 때는 과감하게 **매출최적화 광고를 완전히 끄고 '수동성과형 광고'로 갈아타서** 검색 상단을 직접 점령하는 것이 훨씬 유리합니다.
                        """)
                    else:
                        st.warning(f"**[진단] 비검색영역 볼륨은 크지만 실질적인 효율은 검색이 더 낫습니다.**")
                        st.markdown("""
                        * **액션 플랜 (방어적 투트랙):** 매출 볼륨을 당장 포기할 수 없으니 매최는 유지하세요. 대신 **매최 광고의 목표 ROAS를 살짝 높여 방어적으로 돌리고, 수동성과형 광고를 병행하여 검색 타겟팅을 강화**하는 투트랙 테스트를 권장합니다.
                        """)
                
                else:
                    if search_sales_pct >= non_search_sales_pct and search_roas_val >= non_search_roas_val:
                        st.success(f"**[진단] 수동광고의 정석! 검색영역 매출({search_sales_pct:.1f}%)과 효율이 모두 훌륭합니다.**")
                        st.markdown("""
                        * **액션 플랜:** 직접 세팅하신 키워드들이 시장에서 정확히 먹히고 있습니다. 3단계 표를 확인하여 **효율이 좋은 핵심 키워드의 CPC 입찰가를 조금 더 상향**하여 상단 점유율을 꽉 잡으세요.
                        * **단가 세팅:** 클릭만 많고 돈만 나가는 2단계 블랙홀 키워드들은 가차없이 OFF 처리하여 일예산을 방어하세요.
                        """)
                    elif non_search_sales_pct > search_sales_pct:
                        st.warning(f"**[진단] 수동광고임에도 비검색영역(스마트타겟팅 등)의 매출({non_search_sales_pct:.1f}%)이 더 큽니다.**")
                        st.markdown("""
                        * **액션 플랜 (광고방식 변경 고려):** 수동으로 설정한 키워드가 빗나갔거나, 오히려 쿠팡 알고리즘이 제품 타겟을 더 잘 찾고 있습니다. 수동을 끄고 **'매출최적화 광고'로 전환하여 쿠팡 AI에게 전적으로 맡겨보는 것**을 추천합니다.
                        """)
                    else:
                        st.error(f"**[진단] 검색/비검색 모두 전반적인 ROAS 효율이 너무 낮습니다.**")
                        st.markdown("""
                        * **액션 플랜 (키워드 다이어트 및 리셋):** 수동 키워드에서 클릭만 일어날 뿐 구매가 나오지 않습니다. 2단계 제외 키워드를 대폭 솎아내시고, 며칠 더 지켜봐도 개선되지 않는다면 광고를 끄고 썸네일/상세페이지를 먼저 점검하세요.
                        """)

        st.markdown("<br><br>", unsafe_allow_html=True)
        st.divider()
        st.markdown("<br>", unsafe_allow_html=True)

        # ════════════════════════════════════════════════════════
        # [2단계] 제외 키워드 추출
        # ════════════════════════════════════════════════════════
        st.header("2️⃣ 자동 제외 키워드 추출 (Top 30)")
        
        df_keywords = pivot_df[~non_search_condition].copy()
        
        top_spend = df_keywords.sort_values(by='광고비', ascending=False).head(30)
        top_cpc = df_keywords.sort_values(by='CPC', ascending=False).head(30)

        bad_spend_kw = top_spend[top_spend['총 전환매출액(14일)'] == 0]['키워드'].tolist()
        bad_cpc_kw = top_cpc[top_cpc['총 전환매출액(14일)'] == 0]['키워드'].tolist()
        negative_keywords = list(set(bad_spend_kw + bad_cpc_kw))
        
        if len(negative_keywords) > 0:
            st.error("❗ 아래 키워드들을 쿠팡 광고센터의 [제외 키워드] 란에 즉시 추가하세요.")
            st.text_area(label="전체 복사 (매출 0원 & 고비용 키워드)", value=", ".join(negative_keywords), height=300)
        
        st.write("<br>", unsafe_allow_html=True) 

        def highlight_sales_status(row):
            if row['총 전환매출액(14일)'] > 0:
                return ['background-color: #F0FDF4; color: #166534; font-weight: 500; font-size: 16px'] * len(row)
            return ['background-color: #FEF2F2; color: #B91C1C; font-weight: 400; font-size: 16px'] * len(row)

        col_kw1, col_kw2 = st.columns(2)
        with col_kw1:
            st.subheader("💸 광고비 지출 Top 30")
            st.dataframe(top_spend[['키워드', '광고비', 'ROAS', '총 전환매출액(14일)']]\
                .style.apply(highlight_sales_status, axis=1)\
                .set_properties(subset=['키워드'], **{'text-align': 'center'})\
                .set_properties(subset=['광고비', 'ROAS', '총 전환매출액(14일)'], **{'text-align': 'right'})\
                .set_table_styles([dict(selector='th', props=[('text-align', 'center')])])\
                .format({
                    '광고비': '{:,.0f}', 'ROAS': '{:,.2f}', '총 전환매출액(14일)': '{:,.0f}'
                }), use_container_width=True, hide_index=True)

        with col_kw2:
            st.subheader("📈 평균 CPC Top 30")
            st.dataframe(top_cpc[['키워드', 'CPC', '클릭수', '광고비', '총 전환매출액(14일)']]\
                .style.apply(highlight_sales_status, axis=1)\
                .set_properties(subset=['키워드'], **{'text-align': 'center'})\
                .set_properties(subset=['CPC', '클릭수', '광고비', '총 전환매출액(14일)'], **{'text-align': 'right'})\
                .set_table_styles([dict(selector='th', props=[('text-align', 'center')])])\
                .format({
                    'CPC': '{:,.0f}', '클릭수': '{:,.0f}', '광고비': '{:,.0f}', '총 전환매출액(14일)': '{:,.0f}'
                }), use_container_width=True, hide_index=True)

        st.markdown("<br><br>", unsafe_allow_html=True)
        st.divider()
        st.markdown("<br>", unsafe_allow_html=True)

        # ════════════════════════════════════════════════════════
        # [3단계] 키워드별 상세 분석
        # ════════════════════════════════════════════════════════
        st.header("3️⃣ 키워드별 상세 분석 전체 시트")
        
        final_df = pivot_df.copy()
        final_df.loc[non_search_condition, '키워드'] = '비검색영역'
        final_df = final_df.rename(columns={'총 주문수(14일)': '주문', '총 판매수량(14일)': '수량', '총 전환매출액(14일)': '매출액'})
        
        def highlight_roas_soft(row):
            if row['ROAS'] > 0:
                color = 'background-color: #F0FDF4; color: #1f2937; font-weight: 500; font-size: 16px'
            else:
                color = 'color: #1f2937; font-weight: 400; font-size: 16px'
            return [color] * len(row)

        cols_order = ['키워드', '노출수', '클릭수', 'CPC', '광고비', '주문', '수량', '매출액', 'ROAS']
        final_df = final_df.sort_values(by='매출액', ascending=False)[cols_order]

        col_title, col_btn = st.columns([8, 2])
        with col_title:
            st.markdown("전체 데이터를 확인하고 우측 버튼을 눌러 **엑셀 파일(.xlsx)** 로 다운로드하세요.")
        with col_btn:
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                final_df.to_excel(writer, index=False, sheet_name='분석결과')
            
            st.download_button(
                label="📥 엑셀 다운로드",
                data=buffer.getvalue(),
                file_name="쿠팡_광고분석_상세데이터.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )

        st.dataframe(
            final_df.style.apply(highlight_roas_soft, axis=1)\
                .set_properties(subset=['키워드'], **{'text-align': 'center'})\
                .set_properties(subset=['노출수', '클릭수', 'CPC', '광고비', '주문', '수량', '매출액', 'ROAS'], **{'text-align': 'right'})\
                .set_table_styles([dict(selector='th', props=[('text-align', 'center')])])\
                .format({
                    '노출수': '{:,.0f}', '클릭수': '{:,.0f}', 'CPC': '{:,.0f}',
                    '광고비': '{:,.0f}', '주문': '{:,.0f}', '수량': '{:,.0f}', '매출액': '{:,.0f}', 'ROAS': '{:,.2f}'
                }), 
            use_container_width=True, 
            hide_index=True
        )

    except Exception as e:
        st.error(f"데이터 처리 중 오류가 발생했습니다. (에러: {e})")
