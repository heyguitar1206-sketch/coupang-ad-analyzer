import streamlit as st
import pandas as pd
import numpy as np
import io  

# [디자인] 페이지 기본 설정
st.set_page_config(page_title="쿠팡 광고 분석기", layout="wide")

st.markdown("""
    <style>
    /* 콘텐츠 가로 길이를 제한하고 중앙 정렬 */
    .block-container { max-width: 70% !important; padding-top: 3rem !important; padding-bottom: 5rem !important; }
    
    /* 세련된 '프리텐다드(Pretendard)' 폰트 임포트 */
    @import url('https://cdn.jsdelivr.net/gh/orioncactus/pretendard@v1.3.9/dist/web/static/pretendard.min.css');
    
    /* 전체 폰트 모양 적용 및 기본 UI 보호 */
    html, body, div, p, span, text { font-family: 'Pretendard', -apple-system, BlinkMacSystemFont, system-ui, Roboto, sans-serif; }
    i, .stIcon, .material-symbols-rounded, .material-icons { font-family: 'Material Symbols Rounded', 'Material Icons' !important; }
    .stMarkdown p, .stMarkdown li { color: #1f2937 !important; line-height: 1.8 !important; font-size: 15px !important; }
    
    div[data-testid="stFileUploadDropzone"] * { line-height: initial !important; letter-spacing: normal !important; margin: 0 !important; padding: 0 !important; }
    
    h1, h2, h3, .stHeader h1, .stHeader h2, .stHeader h3 { color: #2563EB !important; font-weight: 700 !important; letter-spacing: -0.8px !important; line-height: 1.4 !important; font-family: 'Pretendard', sans-serif !important; }
    
    /* 로그인 화면 컴팩트 디자인 */
    .login-wrapper { display: flex; flex-direction: column; align-items: center; justify-content: center; padding-top: 40px; }
    .login-card { background-color: white; padding: 50px 35px; border-radius: 20px; box-shadow: 0 10px 30px rgba(0, 0, 0, 0.04); border: 1px solid #F1F5F9; width: 100%; max-width: 420px; text-align: center; }
    div[data-testid="stForm"] button { background-color: #2563EB !important; color: white !important; font-weight: 700 !important; border-radius: 8px !important; padding: 0.6rem 2rem !important; border: none !important; margin-top: 5px !important; font-size: 16px !important; }
    div[data-testid="stTextInput"] input { border-radius: 8px !important; border: 1px solid #E2E8F0 !important; padding: 10px !important; background-color: #F8FAFC !important; }
    div[data-testid="stForm"] { border: none !important; padding: 0 !important; }

    [data-testid="stMetricValue"] { font-size: 28px !important; font-weight: 800 !important; color: #2563EB !important; letter-spacing: -0.5px !important; }
    [data-testid="stMetricLabel"] * { font-size: 16px !important; font-weight: 700 !important; color: #4B5563 !important; }
    table.stTable td:first-child { text-align: center !important; }
    table.stTable td:not(:first-child) { text-align: right !important; }
    thead tr th, table.stTable th { background-color: #F3F6FF !important; color: #1D4ED8 !important; font-weight: 700 !important; font-size: 15px !important; padding: 12px 10px !important; text-align: center !important; }
    tbody tr td { padding: 10px 10px !important; font-size: 16px !important; }
    
    /* 다운로드 버튼 스타일링 */
    .stDownloadButton button { border-radius: 8px !important; font-weight: bold !important; }
    </style>
    """, unsafe_allow_html=True)


# ════════════════════════════════════════════════════════
# 🔐 [로그인 시스템]
# ════════════════════════════════════════════════════════

if 'logged_in' not in st.session_state:
    st.session_state['logged_in'] = False

if not st.session_state['logged_in']:
    st.markdown('<div class="login-wrapper">', unsafe_allow_html=True)
    st.markdown("<h1 style='text-align: center; font-size: 36px; margin-bottom: 8px;'>🎯 끝장캐리 수동끝판왕</h1>", unsafe_allow_html=True)
    st.markdown("<h2 style='text-align: center; color: #1e293b; font-size: 22px; margin-top: 0; font-weight: 500;'>수강생 전용 시스템</h2>", unsafe_allow_html=True)
    st.markdown("<p style='text-align: center; color: #64748b; font-size: 14px; margin-bottom: 25px;'>발급받은 아이디와 비밀번호를 입력해 주세요.</p>", unsafe_allow_html=True)
    
    st.markdown('<div class="login-card">', unsafe_allow_html=True)
    with st.form("login_form"):
        input_id = st.text_input("아이디 입력", placeholder="ID")
        input_pw = st.text_input("비밀번호 입력", type="password", placeholder="Password")
        submit_btn = st.form_submit_button("접속하기", use_container_width=True)
        
        if submit_btn:
            try:
                users_df = pd.read_csv("user.csv")
                users_df['id'] = users_df['id'].astype(str).str.strip()
                users_df['password'] = users_df['password'].astype(str).str.strip()
                match = users_df[(users_df['id'] == input_id.strip()) & (users_df['password'] == input_pw.strip())]
                
                if not match.empty:
                    st.session_state['logged_in'] = True
                    st.rerun()
                else:
                    st.error("정보가 일치하지 않습니다.")
            except FileNotFoundError:
                st.error("시스템 에러: 'user.csv' 파일을 찾을 수 없습니다.")
    st.markdown('</div>', unsafe_allow_html=True)
    st.markdown('</div>', unsafe_allow_html=True)


# 🟢 로그인 성공 시 메인 분석 화면
else:
    col_empty, col_logout = st.columns([9, 1])
    with col_logout:
        if st.button("로그아웃"):
            st.session_state['logged_in'] = False
            st.rerun()

    st.title("📊 쿠팡 광고보고서 자동 분석기")
    st.markdown("쿠팡 윙(Wing) 스타일의 직관적인 인터페이스로 광고 성과를 심층 분석합니다.")
    st.write("<br>", unsafe_allow_html=True) 

    uploaded_file = st.file_uploader("분석할 광고보고서 엑셀 파일을 업로드하세요", type=['xlsx', 'xls'])

    if uploaded_file is not None:
        try:
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
            
            pivot_df = pd.pivot_table(df_raw, index='키워드', values=['노출수', '클릭수', '광고비', '총 주문수(14일)', '총 판매수량(14일)', '총 전환매출액(14일)'], aggfunc='sum').reset_index()
            pivot_df['CPC'] = np.where(pivot_df['클릭수'] > 0, round(pivot_df['광고비'] / pivot_df['클릭수'], 0), 0)
            pivot_df['ROAS'] = np.where(pivot_df['광고비'] > 0, round((pivot_df['총 전환매출액(14일)'] / pivot_df['광고비']) * 100, 2), 0)
            kw_str = pivot_df['키워드'].astype(str).str.strip().str.lower()
            non_search_condition = kw_str.isin(['-', 'nan', 'none', ''])
            df_total = pivot_df.sum(numeric_only=True)
            df_non_search = pivot_df[non_search_condition].sum(numeric_only=True)
            df_search = df_total - df_non_search
            def safe_div(a, b): return a / b if b and b > 0 else 0

            # 1단계
            st.header("1️⃣ 전체 성과 및 영역별 요약")
            total_ad_spend, total_sales = df_total.get('광고비', 0), df_total.get('총 전환매출액(14일)', 0)
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

            summary_df = pd.DataFrame([
                {'구분': '총합계', '노출수': df_total.get('노출수',0), '클릭수': df_total.get('클릭수',0), 'CPC': safe_div(total_ad_spend, df_total.get('클릭수',0)), '광고비': total_ad_spend, '광고비비중': 100.0, '주문수': total_orders, '판매수량': df_total.get('총 판매수량(14일)',0), '매출액': total_sales, '매출비중': 100.0, 'ROAS': total_roas},
                {'구분': '비검색영역', '노출수': df_non_search.get('노출수',0), '클릭수': df_non_search.get('클릭수',0), 'CPC': safe_div(df_non_search.get('광고비',0), df_non_search.get('클릭수',0)), '광고비': df_non_search.get('광고비',0), '광고비비중': safe_div(df_non_search.get('광고비',0), total_ad_spend) * 100, '주문수': df_non_search.get('총 주문수(14일)',0), '판매수량': df_non_search.get('총 판매수량(14일)',0), '매출액': df_non_search.get('총 전환매출액(14일)',0), '매출비중': non_search_sales_pct, 'ROAS': non_search_roas_val},
                {'구분': '검색영역', '노출수': df_search.get('노출수',0), '클릭수': df_search.get('클릭수',0), 'CPC': safe_div(df_search.get('광고비',0), df_search.get('클릭수',0)), '광고비': df_search.get('광고비',0), '광고비비중': safe_div(df_search.get('광고비',0), total_ad_spend) * 100, '주문수': df_search.get('총 주문수(14일)',0), '판매수량': df_search.get('총 판매수량(14일)',0), '매출액': df_search.get('총 전환매출액(14일)',0), '매출비중': search_sales_pct, 'ROAS': search_roas_val}
            ])
            
            def highlight_summary(row):
                if row['구분'] == '총합계': return ['background-color: #FFF4E5; color: #EA580C; font-weight: 700; font-size: 16px; border-bottom: 2px solid #EA580C'] * len(row)
                return ['background-color: white; color: #374151; font-weight: 500; font-size: 15px'] * len(row)

            st.table(summary_df.style.apply(highlight_summary, axis=1).set_properties(subset=['구분'], **{'text-align': 'center'}).set_properties(subset=['노출수', '클릭수', 'CPC', '광고비', '광고비비중', '주문수', '판매수량', '매출액', '매출비중', 'ROAS'], **{'text-align': 'right'}).set_table_styles([dict(selector='th', props=[('text-align', 'center')])]).format({'노출수': '{:,.0f}', '클릭수': '{:,.0f}', 'CPC': '{:,.0f}원', '광고비': '{:,.0f}원', '광고비비중': '{:,.1f}%', '주문수': '{:,.0f}건', '판매수량': '{:,.0f}개', '매출액': '{:,.0f}원', '매출비중': '{:,.1f}%', 'ROAS': '{:,.2f}%'}))

            # 💡 [리포트 연동을 위한 진단 로직 변수화]
            diagnosis_state = ""
            diagnosis_title = ""
            diagnosis_desc = ""

            if ad_type_detected == "매출최적화":
                if non_search_sales_pct >= search_sales_pct and non_search_roas_val >= search_roas_val:
                    diagnosis_state = "success"
                    diagnosis_title = f"[진단] 전형적인 매출최적화 성공 패턴! 비검색영역 매출({non_search_sales_pct:.1f}%)과 효율이 모두 우수합니다."
                    diagnosis_desc = "현재 매최 광고가 제품과 찰떡궁합으로 잘 돌고 있습니다. 볼륨을 키우기 위해 목표수익률(ROAS) 세팅값을 평소보다 50% ~ 100% 정도 상향시켜 마진율 극대화를 시도해 보세요. 또한 추출된 '제외 키워드'를 꾸준히 입력하여 검색영역에서 새는 돈만 막아주면 됩니다."
                elif search_sales_pct > non_search_sales_pct and search_roas_val > non_search_roas_val:
                    diagnosis_state = "info"
                    diagnosis_title = f"[진단] 매최 광고임에도 검색영역 성과({search_sales_pct:.1f}%)가 두드러지게 좋습니다."
                    diagnosis_desc = "검색을 통한 유입과 전환이 아주 훌륭합니다. 이때 효율만 보고 매최를 끄면 기존 매출이 박살 납니다! 기존 매출최적화 광고는 볼륨 방어용으로 그대로 켜두고, 성과 좋은 핵심 키워드만 따로 빼서 '수동성과형 광고'를 새롭게 추가 개설(투트랙 테스트) 하세요."
                elif non_search_roas_val < (total_roas * 0.5) or non_search_sales_pct < 20:
                    diagnosis_state = "warning"
                    diagnosis_title = "[진단] 비검색영역의 효율이 심각하게 부진하며 돈만 까먹고 있습니다."
                    diagnosis_desc = "매최 광고의 알고리즘이 비검색 영역에서 타겟을 전혀 찾지 못하고 있습니다. 이럴 때는 과감하게 매출최적화 광고를 완전히 끄고 '수동성과형 광고'로 갈아타서 검색 상단을 직접 점령하는 것이 훨씬 유리합니다."
                else:
                    diagnosis_state = "warning"
                    diagnosis_title = "[진단] 비검색영역 볼륨은 크지만 실질적인 효율은 검색이 더 낫습니다."
                    diagnosis_desc = "매출 볼륨을 당장 포기할 수 없으니 매최는 유지하세요. 대신 매최 광고의 목표 ROAS를 살짝 높여 방어적으로 돌리고, 수동성과형 광고를 병행하여 검색 타겟팅을 강화하는 투트랙 테스트를 권장합니다."
            else:
                if search_sales_pct >= non_search_sales_pct and search_roas_val >= non_search_roas_val:
                    diagnosis_state = "success"
                    diagnosis_title = f"[진단] 수동광고의 정석! 검색영역 매출({search_sales_pct:.1f}%)과 효율이 모두 훌륭합니다."
                    diagnosis_desc = "직접 세팅하신 키워드들이 시장에서 정확히 먹히고 있습니다. 효율이 좋은 핵심 키워드의 CPC 입찰가를 조금 더 상향하여 상단 점유율을 꽉 잡으세요. 클릭만 많고 돈만 나가는 블랙홀 키워드들은 가차없이 OFF 처리하여 일예산을 방어하세요."
                elif non_search_sales_pct > search_sales_pct:
                    diagnosis_state = "warning"
                    diagnosis_title = f"[진단] 수동광고임에도 비검색영역(스마트타겟팅 등)의 매출({non_search_sales_pct:.1f}%)이 더 큽니다."
                    diagnosis_desc = "수동으로 설정한 키워드가 빗나갔거나, 오히려 쿠팡 알고리즘이 제품 타겟을 더 잘 찾고 있습니다. 수동을 끄고 '매출최적화 광고'로 전환하여 쿠팡 AI에게 전적으로 맡겨보는 것을 추천합니다."
                else:
                    diagnosis_state = "error"
                    diagnosis_title = "[진단] 검색/비검색 모두 전반적인 ROAS 효율이 너무 낮습니다."
                    diagnosis_desc = "수동 키워드에서 클릭만 일어날 뿐 구매가 나오지 않습니다. 제외 키워드를 대폭 솎아내시고, 며칠 더 지켜봐도 개선되지 않는다면 광고를 끄고 썸네일/상세페이지를 먼저 점검하세요."

            with st.container():
                st.write("<br>", unsafe_allow_html=True)
                if total_sales > 0:
                    st.markdown(f"#### 💡 끝장캐리 실전 가이드 (현재 주력 광고: **{ad_type_detected}**)")
                    if diagnosis_state == "success":
                        st.success(f"**{diagnosis_title}**\n\n* **액션 플랜:** {diagnosis_desc}")
                    elif diagnosis_state == "info":
                        st.info(f"**{diagnosis_title}**\n\n* **액션 플랜:** {diagnosis_desc}")
                    elif diagnosis_state == "warning":
                        st.warning(f"**{diagnosis_title}**\n\n* **액션 플랜:** {diagnosis_desc}")
                    else:
                        st.error(f"**{diagnosis_title}**\n\n* **액션 플랜:** {diagnosis_desc}")

            st.markdown("<br><br>", unsafe_allow_html=True); st.divider(); st.markdown("<br>", unsafe_allow_html=True)

            # 2단계
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
            
            def highlight_sales_status(row):
                if row['총 전환매출액(14일)'] > 0: return ['background-color: #F0FDF4; color: #166534; font-weight: 500; font-size: 16px'] * len(row)
                return ['background-color: #FEF2F2; color: #B91C1C; font-weight: 400; font-size: 16px'] * len(row)

            col_kw1, col_kw2 = st.columns(2)
            with col_kw1:
                st.subheader("💸 광고비 지출 Top 30")
                st.dataframe(top_spend[['키워드', '광고비', 'ROAS', '총 전환매출액(14일)']].style.apply(highlight_sales_status, axis=1).set_properties(subset=['키워드'], **{'text-align': 'center'}).set_properties(subset=['광고비', 'ROAS', '총 전환매출액(14일)'], **{'text-align': 'right'}).set_table_styles([dict(selector='th', props=[('text-align', 'center')])]).format({'광고비': '{:,.0f}', 'ROAS': '{:,.2f}', '총 전환매출액(14일)': '{:,.0f}'}), use_container_width=True, hide_index=True)
            with col_kw2:
                st.subheader("📈 평균 CPC Top 30")
                st.dataframe(top_cpc[['키워드', 'CPC', '클릭수', '광고비', '총 전환매출액(14일)']].style.apply(highlight_sales_status, axis=1).set_properties(subset=['키워드'], **{'text-align': 'center'}).set_properties(subset=['CPC', '클릭수', '광고비', '총 전환매출액(14일)'], **{'text-align': 'right'}).set_table_styles([dict(selector='th', props=[('text-align', 'center')])]).format({'CPC': '{:,.0f}', '클릭수': '{:,.0f}', '광고비': '{:,.0f}', '총 전환매출액(14일)': '{:,.0f}'}), use_container_width=True, hide_index=True)

            st.markdown("<br><br>", unsafe_allow_html=True); st.divider(); st.markdown("<br>", unsafe_allow_html=True)

            # 3단계
            st.header("3️⃣ 키워드별 상세 분석 전체 시트")
            final_df = pivot_df.copy()
            final_df.loc[non_search_condition, '키워드'] = '비검색영역'
            final_df = final_df.rename(columns={'총 주문수(14일)': '주문', '총 판매수량(14일)': '수량', '총 전환매출액(14일)': '매출액'})
            
            def highlight_roas_soft(row):
                if row['ROAS'] > 0: color = 'background-color: #F0FDF4; color: #1f2937; font-weight: 500; font-size: 16px'
                else: color = 'color: #1f2937; font-weight: 400; font-size: 16px'
                return [color] * len(row)
            
            final_df = final_df.sort_values(by='매출액', ascending=False)[['키워드', '노출수', '클릭수', 'CPC', '광고비', '주문', '수량', '매출액', 'ROAS']]
            
            # 💡 [핵심 추가] HTML 리포트 생성 및 데이터 포맷팅
            html_summary_df = summary_df.copy()
            html_summary_df['노출수'] = html_summary_df['노출수'].apply(lambda x: f"{x:,.0f}")
            html_summary_df['클릭수'] = html_summary_df['클릭수'].apply(lambda x: f"{x:,.0f}")
            html_summary_df['CPC'] = html_summary_df['CPC'].apply(lambda x: f"{x:,.0f}원")
            html_summary_df['광고비'] = html_summary_df['광고비'].apply(lambda x: f"{x:,.0f}원")
            html_summary_df['광고비비중'] = html_summary_df['광고비비중'].apply(lambda x: f"{x:,.1f}%")
            html_summary_df['주문수'] = html_summary_df['주문수'].apply(lambda x: f"{x:,.0f}건")
            html_summary_df['판매수량'] = html_summary_df['판매수량'].apply(lambda x: f"{x:,.0f}개")
            html_summary_df['매출액'] = html_summary_df['매출액'].apply(lambda x: f"{x:,.0f}원")
            html_summary_df['매출비중'] = html_summary_df['매출비중'].apply(lambda x: f"{x:,.1f}%")
            html_summary_df['ROAS'] = html_summary_df['ROAS'].apply(lambda x: f"{x:,.2f}%")

            kw_list_str = ", ".join(negative_keywords) if len(negative_keywords) > 0 else "발견된 악성 제외 키워드가 없습니다."

            html_report = f"""
            <!DOCTYPE html>
            <html lang="ko">
            <head>
                <meta charset="UTF-8">
                <title>끝장캐리 광고 분석 리포트</title>
                <style>
                    @import url('https://cdn.jsdelivr.net/gh/orioncactus/pretendard@v1.3.9/dist/web/static/pretendard.min.css');
                    body {{ font-family: 'Pretendard', sans-serif; color: #1f2937; line-height: 1.6; padding: 40px; max-width: 900px; margin: auto; }}
                    h1 {{ color: #2563EB; border-bottom: 2px solid #2563EB; padding-bottom: 10px; font-weight: 800; }}
                    h2 {{ color: #1D4ED8; margin-top: 30px; font-weight: 700; }}
                    .metric-box {{ display: flex; justify-content: space-between; background: #F8FAFC; padding: 20px; border-radius: 12px; margin-bottom: 20px; border: 1px solid #E2E8F0; }}
                    .metric {{ text-align: center; }}
                    .metric .value {{ font-size: 24px; font-weight: bold; color: #2563EB; }}
                    .metric .label {{ font-size: 14px; color: #64748b; font-weight: bold; }}
                    .diagnosis-box {{ background: #F0FDF4; border-left: 5px solid #166534; padding: 20px; border-radius: 8px; margin-bottom: 30px; }}
                    .diagnosis-box h3 {{ margin-top: 0; color: #166534; font-weight: bold; }}
                    table {{ width: 100%; border-collapse: collapse; margin-top: 15px; font-size: 14px; }}
                    th, td {{ border: 1px solid #E2E8F0; padding: 10px; text-align: right; }}
                    th {{ background-color: #F3F6FF; color: #1D4ED8; text-align: center; font-weight: bold; }}
                    td:first-child {{ text-align: center; font-weight: bold; }}
                    .kw-box {{ background: #FEF2F2; padding: 15px; border-radius: 8px; color: #B91C1C; font-weight: bold; word-break: break-all; line-height: 1.8; }}
                    @media print {{ body {{ padding: 0; }} .no-print {{ display: none !important; }} }}
                </style>
            </head>
            <body>
                <h1>🎯 끝장캐리 수동끝판왕 - 주간 광고 분석 리포트</h1>
                
                <div class="metric-box">
                    <div class="metric"><div class="label">총 전환매출액</div><div class="value">{total_sales:,.0f}원</div></div>
                    <div class="metric"><div class="label">총 지출 광고비</div><div class="value">{total_ad_spend:,.0f}원</div></div>
                    <div class="metric"><div class="label">전체 평균 ROAS</div><div class="value">{total_roas:,.2f}%</div></div>
                    <div class="metric"><div class="label">총 주문수</div><div class="value">{total_orders:,.0f}건</div></div>
                </div>

                <div class="diagnosis-box">
                    <h3>💡 AI 코칭 진단 (현재 주력: {ad_type_detected})</h3>
                    <p><strong>{diagnosis_title}</strong></p>
                    <p>👉 {diagnosis_desc}</p>
                </div>

                <h2>1. 전체 성과 및 영역별 요약</h2>
                {html_summary_df.to_html(index=False, classes='table')}

                <h2>2. 즉시 추가해야 할 제외 키워드 (매출 0원 & 광고비 누수)</h2>
                <div class="kw-box">{kw_list_str}</div>
                
                <div class="no-print" style="margin-top: 50px; text-align: center; background: #F9FAFB; padding: 30px; border-radius: 12px; border: 1px dashed #D1D5DB;">
                    <h3 style="margin-top:0; color:#4B5563;">🖨️ PDF 리포트 저장 안내</h3>
                    <p style="color:#6B7280; font-size:14px;">아래 버튼을 누르면 인쇄 창이 열립니다. 대상(프린터)을 <strong>'PDF로 저장'</strong>으로 선택하시고 저장해 주세요.</p>
                    <button onclick="window.print()" style="background:#2563EB; color:white; border:none; padding:14px 28px; border-radius:8px; font-size:16px; cursor:pointer; font-weight:bold; box-shadow: 0 4px 6px rgba(37,99,235,0.2);">📄 PDF로 저장하기 (인쇄)</button>
                </div>
            </body>
            </html>
            """

            # 💡 다운로드 버튼 배치
            st.markdown("### 📥 리포트 및 데이터 다운로드")
            col_dl1, col_dl2 = st.columns(2)
            
            with col_dl1:
                st.download_button(
                    label="📄 종합 리포트 다운로드 (PDF 저장용)",
                    data=html_report,
                    file_name="끝장캐리_종합_분석리포트.html",
                    mime="text/html",
                    use_container_width=True
                )
                
            with col_dl2:
                buffer = io.BytesIO()
                with pd.ExcelWriter(buffer, engine='openpyxl') as writer: 
                    final_df.to_excel(writer, index=False, sheet_name='분석결과')
                st.download_button(
                    label="📥 상세 키워드 엑셀 다운로드 (원시데이터)", 
                    data=buffer.getvalue(), 
                    file_name="쿠팡_광고분석_상세데이터.xlsx", 
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", 
                    use_container_width=True
                )

            st.write("<br>", unsafe_allow_html=True) 
            
            # 마지막 데이터 프레임 출력
            st.dataframe(final_df.style.apply(highlight_roas_soft, axis=1).set_properties(subset=['키워드'], **{'text-align': 'center'}).set_properties(subset=['노출수', '클릭수', 'CPC', '광고비', '주문', '수량', '매출액', 'ROAS'], **{'text-align': 'right'}).set_table_styles([dict(selector='th', props=[('text-align', 'center')])]).format({'노출수': '{:,.0f}', '클릭수': '{:,.0f}', 'CPC': '{:,.0f}', '광고비': '{:,.0f}', '주문': '{:,.0f}', '수량': '{:,.0f}', '매출액': '{:,.0f}', 'ROAS': '{:,.2f}'}), use_container_width=True, hide_index=True)

        except Exception as e:
            st.error(f"데이터 처리 중 오류가 발생했습니다. (에러: {e})")
