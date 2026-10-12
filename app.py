import pandas as pd
import streamlit as st
import re

# ==========================================
# 1. 페이지 기본 설정 및 모바일 CSS 스타일링
# ==========================================
st.set_page_config(
    page_title="현장 점유율 및 채용 현황",
    layout="wide",
    initial_sidebar_state="collapsed",
)

st.markdown(
    """
    <style>
    /* 전체 글꼴 및 기본 폰트 크기 */
    html, body, [class*="css"] {
        font-size: 17px !important;
    }
    
    /* 사이드바 메뉴 글씨 크기 확대 */
    [data-testid="stSidebar"] .stRadio label p, 
    [data-testid="stSidebar"] .stRadio div[role="radiogroup"] label span {
        font-size: 1.35rem !important;
        font-weight: 800 !important;
    }
    [data-testid="stSidebar"] .stRadio div[role="radiogroup"] {
        gap: 12px;
    }
    
    /* 제목 및 헤더 크기 */
    h1 { font-size: 1.9rem !important; font-weight: bold !important; color: #1E3A8A; }
    h2 { font-size: 1.5rem !important; font-weight: bold !important; color: #1E40AF; }
    h3 { font-size: 1.3rem !important; font-weight: bold !important; }
    
    /* 여백 설정 */
    .block-container {
        padding-top: 1.2rem !important;
        padding-bottom: 2.5rem !important;
        padding-left: 1rem !important;
        padding-right: 1rem !important;
        max-width: 100% !important;
    }
    
    /* 테이블 글씨 크기 */
    .dataframe {
        font-size: 15px !important;
    }
    
    /* 전화번호 연결 버튼 스타일 */
    .phone-btn {
        display: inline-block;
        background-color: #25D366;
        color: white !important;
        padding: 9px 16px;
        font-size: 16px;
        font-weight: bold;
        text-decoration: none;
        border-radius: 8px;
        margin-top: 5px;
        margin-bottom: 8px;
        text-align: center;
        box-shadow: 0 2px 4px rgba(0,0,0,0.2);
    }

    /* 메인 화면 카드 스타일 */
    .main-total-card {
        background: linear-gradient(135deg, #1E3A8A, #3B82F6);
        color: white;
        padding: 14px 18px;
        border-radius: 10px;
        text-align: center;
        margin-bottom: 15px;
        box-shadow: 0 3px 5px rgba(0,0,0,0.1);
    }
    .main-stat-card {
        background-color: #FFFFFF;
        border-left: 6px solid #3B82F6;
        border-radius: 8px;
        padding: 10px 14px;
        margin-bottom: 10px;
        box-shadow: 0 2px 4px rgba(0,0,0,0.08);
    }
    .border-hanno { border-left-color: #EF4444 !important; }
    .border-minno { border-left-color: #3B82F6 !important; }
    .border-seomoo { border-left-color: #10B981 !important; }
    .border-geunsan { border-left-color: #F59E0B !important; }
    .border-jikwon { border-left-color: #8B5CF6 !important; }
    .border-mijeong { border-left-color: #6B7280 !important; }

    /* 타워사 모바일 전용 카드 스타일 */
    .tower-card {
        background-color: #FFFFFF;
        border-radius: 10px;
        padding: 14px 16px;
        margin-bottom: 12px;
        box-shadow: 0 2px 5px rgba(0,0,0,0.08);
        border: 1px solid #E5E7EB;
    }
    .tower-card-uho {
        border-left: 6px solid #2563EB !important;
        background-color: #F0F6FF;
    }
    .tower-card-bihyop {
        border-left: 6px solid #DC2626 !important;
        background-color: #FEF2F2;
    }
    .tower-card-normal {
        border-left: 6px solid #9CA3AF !important;
    }
    .badge-uho {
        background-color: #2563EB;
        color: white;
        padding: 3px 8px;
        border-radius: 12px;
        font-size: 0.85rem;
        font-weight: bold;
    }
    .badge-bihyop {
        background-color: #DC2626;
        color: white;
        padding: 3px 8px;
        border-radius: 12px;
        font-size: 0.85rem;
        font-weight: bold;
    }
    .badge-normal {
        background-color: #6B7280;
        color: white;
        padding: 3px 8px;
        border-radius: 12px;
        font-size: 0.85rem;
        font-weight: bold;
    }
    </style>
""",
    unsafe_allow_html=True,
)

EXCEL_FILE = "채용비율-3.xlsx"

@st.cache_data(ttl=5)
def load_all_sheets():
    try:
        xls = pd.ExcelFile(EXCEL_FILE)
        sheets = {sheet: pd.read_excel(xls, sheet, header=None) for sheet in xls.sheet_names}
        return sheets
    except Exception as e:
        st.error(f"엑셀 파일을 읽는 중 오류가 발생했습니다. '{EXCEL_FILE}' 파일이 같은 폴더에 있는지 확인해주세요.\n오류 내용: {e}")
        return None

sheets = load_all_sheets()
if sheets is None:
    st.stop()

sheet_names = list(sheets.keys())
s1_name = sheet_names[0] if len(sheet_names) > 0 else "1.26년 전체 소속별 점유율"
s2_name = sheet_names[1] if len(sheet_names) > 1 else "2.각8개지부 현장점유율"
s3_name = sheet_names[2] if len(sheet_names) > 2 else "3.타워사별 점유현황"
s4_name = sheet_names[3] if len(sheet_names) > 3 else "4.반도체현장"
s5_name = sheet_names[4] if len(sheet_names) > 4 else "5.1~9월채용추이"

def safe_int(val):
    try:
        if pd.isna(val) or val == "" or val == "-":
            return 0
        return int(float(val))
    except:
        return 0

# ==========================================
# 사이드바 메뉴 설정
# ==========================================
st.sidebar.title("📌 채용 비율 메뉴 선택")

menu_options = [
    "📊 1. 26년 소속별 점유율(반도체제외)",
    "🏢 2. 각 8개지부 현장 점유율",
    "🏗️ 3. 타워사별 점유 현황",
    "🏭 4. 반도체 현장 현황",
    "📅 5. 월별 신규 채용 추이",
]

menu = st.sidebar.radio("원하시는 화면을 선택하세요", menu_options)

# ==========================================
# 메뉴별 화면 구현
# ==========================================

# ------------------------------------------
# [메뉴 1] 26년 소속별 점유율(반도체제외) 및 반도체 포함 통합
# ------------------------------------------
if menu == "📊 1. 26년 소속별 점유율(반도체제외)":
    st.title("📊 26년 소속별 점유율 (반도체 제외)")
    df1 = sheets[s1_name].copy()

    try:
        hanno_cnt1 = safe_int(df1.iloc[2, 1])
        minno_cnt1 = safe_int(df1.iloc[2, 2])
        seomoo_cnt1 = safe_int(df1.iloc[2, 3])
        geunsan_cnt1 = safe_int(df1.iloc[2, 4])
        jikwon_cnt1 = safe_int(df1.iloc[2, 5])
        mijeong_cnt1 = safe_int(df1.iloc[2, 6])
        tot_cnt1 = safe_int(df1.iloc[2, 7])

        def calc_p(cnt, total):
            return f"{(cnt / total) * 100:.1f}%" if total > 0 else "0.0%"

        stat_items1 = [
            {"title": "한국노총 (한노)", "cnt": hanno_cnt1, "pct": calc_p(hanno_cnt1, tot_cnt1), "class": "border-hanno", "color": "#EF4444"},
            {"title": "민주노총 (민노)", "cnt": minno_cnt1, "pct": calc_p(minno_cnt1, tot_cnt1), "class": "border-minno", "color": "#3B82F6"},
            {"title": "섬유노조", "cnt": seomoo_cnt1, "pct": calc_p(seomoo_cnt1, tot_cnt1), "class": "border-seomoo", "color": "#10B981"},
            {"title": "건설산업 (건산)", "cnt": geunsan_cnt1, "pct": calc_p(geunsan_cnt1, tot_cnt1), "class": "border-geunsan", "color": "#F59E0B"},
            {"title": "직원", "cnt": jikwon_cnt1, "pct": calc_p(jikwon_cnt1, tot_cnt1), "class": "border-jikwon", "color": "#8B5CF6"},
            {"title": "미정", "cnt": mijeong_cnt1, "pct": calc_p(mijeong_cnt1, tot_cnt1), "class": "border-mijeong", "color": "#6B7280"},
        ]

        st.markdown(
            f"""
        <div class="main-total-card">
            <span style="font-size: 1.05rem; font-weight:bold;">🏆 전체 현장 총 대수 (반도체 제외): </span>
            <span style="font-size: 1.5rem; font-weight:bold; margin-left:8px;">{tot_cnt1}대 (100.0%)</span>
        </div>
        """,
            unsafe_allow_html=True,
        )

        st.markdown("<div style='font-size: 1.1rem; font-weight: bold; color: #1E40AF; margin-bottom: 8px;'>👇 소속별 점유 현황 (반도체 제외)</div>", unsafe_allow_html=True)

        for item in stat_items1:
            st.markdown(
                f"""
            <div class="main-stat-card {item['class']}">
                <div style="display:flex; justify-content:space-between; align-items:center;">
                    <span style="font-size: 1.0rem; font-weight:bold; color:{item['color']};">{item['title']}</span>
                    <span style="font-size: 1.15rem; font-weight:bold; color:#111827;">{item['pct']}</span>
                </div>
                <div style="margin-top: 4px; font-size: 0.95rem; color: #4B5563;">
                    보유 수량: <b>{item['cnt']}대</b>
                </div>
            </div>
            """,
                unsafe_allow_html=True,
            )

        # 반도체 현장 합산
        df4_raw = sheets[s4_name].copy()
        parsed_sites_semi = []
        for idx in range(10, len(df4_raw)):
            row = df4_raw.iloc[idx]
            site_name = str(row[1]).strip() if pd.notna(row[1]) else ""
            if not site_name or site_name == "nan" or site_name in ["합 계", "합계"]:
                continue
            parsed_sites_semi.append({
                "total_cnt": safe_int(row[2]),
                "hanno": safe_int(row[3]),
                "minno": safe_int(row[4]),
                "gunsan": safe_int(row[5]),
                "non_union": safe_int(row[6]),
                "mijung": safe_int(row[7])
            })

        semi_tot = sum(s["total_cnt"] for s in parsed_sites_semi)
        semi_hanno = sum(s["hanno"] for s in parsed_sites_semi)
        semi_minno = sum(s["minno"] for s in parsed_sites_semi)
        semi_gunsan = sum(s["gunsan"] for s in parsed_sites_semi) 
        semi_non = sum(s["non_union"] for s in parsed_sites_semi)
        semi_mijung = sum(s["mijung"] for s in parsed_sites_semi)

        total_combined = tot_cnt1 + semi_tot
        comb_hanno = hanno_cnt1 + semi_hanno
        comb_minno = minno_cnt1 + semi_minno
        comb_seomoo = seomoo_cnt1 + semi_gunsan 
        comb_geunsan = geunsan_cnt1    
        comb_jikwon = jikwon_cnt1 + semi_non   # 반도체 비노조 -> 직원 합산
        comb_mijung = mijeong_cnt1 + semi_mijung

        st.markdown("---")
        st.markdown("<br>", unsafe_allow_html=True)
        st.title("📊 26년 소속별 점유율 (반도체 포함 통합)")

        st.markdown(
            f"""
        <div class="main-total-card">
            <span style="font-size: 1.05rem; font-weight:bold;">🏆 전체 현장 총 대수 (반도체 포함): </span>
            <span style="font-size: 1.5rem; font-weight:bold; margin-left:8px;">{total_combined}대 (100.0%)</span>
        </div>
        """,
            unsafe_allow_html=True,
        )

        comb_items = [
            {"title": "한국노총 (한노)", "cnt": comb_hanno, "pct": calc_p(comb_hanno, total_combined), "class": "border-hanno", "color": "#EF4444"},
            {"title": "민주노총 (민노)", "cnt": comb_minno, "pct": calc_p(comb_minno, total_combined), "class": "border-minno", "color": "#3B82F6"},
            {"title": "섬유노조", "cnt": comb_seomoo, "pct": calc_p(comb_seomoo, total_combined), "class": "border-seomoo", "color": "#10B981"},
            {"title": "건설산업 (건산)", "cnt": comb_geunsan, "pct": calc_p(comb_geunsan, total_combined), "class": "border-geunsan", "color": "#F59E0B"},
            {"title": "직원 (반도체 비노조 포함)", "cnt": comb_jikwon, "pct": calc_p(comb_jikwon, total_combined), "class": "border-jikwon", "color": "#8B5CF6"},
            {"title": "미정", "cnt": comb_mijung, "pct": calc_p(comb_mijung, total_combined), "class": "border-mijeong", "color": "#6B7280"},
        ]

        st.markdown("<div style='font-size: 1.1rem; font-weight: bold; color: #1E40AF; margin-bottom: 8px;'>👇 소속별 점유 현황 (반도체 포함 통합)</div>", unsafe_allow_html=True)

        for item in comb_items:
            st.markdown(
                f"""
            <div class="main-stat-card {item['class']}">
                <div style="display:flex; justify-content:space-between; align-items:center;">
                    <span style="font-size: 1.0rem; font-weight:bold; color:{item['color']};">{item['title']}</span>
                    <span style="font-size: 1.15rem; font-weight:bold; color:#111827;">{item['pct']}</span>
                </div>
                <div style="margin-top: 4px; font-size: 0.95rem; color: #4B5563;">
                    보유 수량: <b>{item['cnt']}대</b>
                </div>
            </div>
            """,
                unsafe_allow_html=True,
            )

    except Exception as e:
        st.error(f"데이터를 처리하는 중 오류가 발생했습니다: {e}")
        st.dataframe(df1.fillna(""), use_container_width=True, hide_index=True)


# ------------------------------------------
# [메뉴 2] 각 8개지부 현장 점유율
# ------------------------------------------
elif menu == "🏢 2. 각 8개지부 현장 점유율":
    st.title("🏢 지부별 현장 점유율")
    df2 = sheets[s2_name].copy()
    rows = df2.values.tolist()

    branches = ["북부지부", "남부지부", "남서지부", "동부지부", "서부지부", "북서지부", "용인지부", "중부지부"]

    branch_data = {}
    current_branch = "북부지부"

    for r in rows[1:]:
        branch_col = str(r[0]).strip() if pd.notna(r[0]) else ""
        for b in ["북부", "남부", "남서", "동부", "서부", "북서", "용인", "중부"]:
            if branch_col == b or branch_col == b + "지부":
                current_branch = b + "지부"
                break

        if current_branch not in branch_data:
            branch_data[current_branch] = {"sites": [], "subtotal": None, "pct": None}

        if "소계" in branch_col and "퍼센테이지" not in branch_col:
            branch_data[current_branch]["subtotal"] = r
        elif "소계퍼센테이지" in branch_col or "퍼센테이지" in branch_col:
            branch_data[current_branch]["pct"] = r
        elif pd.notna(r[1]) and str(r[1]).strip() != "현장명":
            branch_data[current_branch]["sites"].append(r)

    st.subheader("🔍 타워사/현장명 통합 검색")
    search_query = st.text_input("타워사 또는 현장명을 입력하면 어느 지부인지 바로 확인 가능합니다", "", placeholder="예: 백산, 복정 등")
    
    if search_query:
        all_searched_sites = []
        for b_name, b_info in branch_data.items():
            for s in b_info["sites"]:
                site_name = str(s[1]) if len(s) > 1 and pd.notna(s[1]) else ""
                tower_name = str(s[2]) if len(s) > 2 and pd.notna(s[2]) else ""
                memo = str(s[3]) if len(s) > 3 and pd.notna(s[3]) else ""
                note = str(s[11]) if len(s) > 11 and pd.notna(s[11]) else ""
                
                if search_query.lower() in site_name.lower() or search_query.lower() in tower_name.lower():
                    is_excluded = "통계제외" in note or "통계제외" in memo or "통계제외" in str(s)
                    all_searched_sites.append({
                        "지부명": b_name,
                        "현장명": f"🔒 [통계제외] {site_name}" if is_excluded else site_name,
                        "타워회사": tower_name,
                        "한노": safe_int(s[4]),
                        "민노": safe_int(s[5]),
                        "섬유": safe_int(s[6]),
                        "건산": safe_int(s[7]),
                        "직원": safe_int(s[8]),
                        "미정": safe_int(s[9]),
                        "총대수": safe_int(s[10]),
                        "비고": note if note else memo
                    })
        if all_searched_sites:
            st.markdown(f"### 🔎 '{search_query}' 검색 결과")
            st.dataframe(pd.DataFrame(all_searched_sites), use_container_width=True, hide_index=True)
        else:
            st.warning("검색 결과가 없습니다.")
        st.markdown("---")

    st.subheader("📌 각 8개 지부 현장 점유율")

    for b_name in branches:
        if b_name in branch_data:
            sites_list = branch_data[b_name]["sites"]

            # 통계제외 항목을 제외하고 합산하도록 필터링
            valid_sites = []
            for s in sites_list:
                memo = str(s[3]) if len(s) > 3 and pd.notna(s[3]) else ""
                note = str(s[11]) if len(s) > 11 and pd.notna(s[11]) else ""
                is_excluded = "통계제외" in note or "통계제외" in memo or "통계제외" in str(s)
                if not is_excluded:
                    valid_sites.append(s)

            hanno_sum = sum(safe_int(s[4]) for s in valid_sites)
            minno_sum = sum(safe_int(s[5]) for s in valid_sites)
            seomoo_sum = sum(safe_int(s[6]) for s in valid_sites)
            geunsan_sum = sum(safe_int(s[7]) for s in valid_sites)
            jikwon_sum = sum(safe_int(s[8]) for s in valid_sites)
            mijeong_sum = sum(safe_int(s[9]) for s in valid_sites)
            total_sum = sum(safe_int(s[10]) for s in valid_sites)

            if total_sum > 0:
                hanno_pct = f"{(hanno_sum / total_sum) * 100:.1f}%"
                minno_pct = f"{(minno_sum / total_sum) * 100:.1f}%"
                seomoo_pct = f"{(seomoo_sum / total_sum) * 100:.1f}%"
                geunsan_pct = f"{(geunsan_sum / total_sum) * 100:.1f}%"
                jikwon_pct = f"{(jikwon_sum / total_sum) * 100:.1f}%"
                mijeong_pct = f"{(mijeong_sum / total_sum) * 100:.1f}%"
            else:
                hanno_pct = minno_pct = seomoo_pct = geunsan_pct = jikwon_pct = mijeong_pct = "0.0%"

            expander_label = (
                f"📁 **{b_name}** | "
                f":red[**🔥 한노 {hanno_sum}({hanno_pct})**] | "
                f":blue[민노 {minno_sum}({minno_pct})] | "
                f":green[섬유 {seomoo_sum}({seomoo_pct})] | "
                f":orange[건산 {geunsan_sum}({geunsan_pct})] | "
                f":violet[직원 {jikwon_sum}({jikwon_pct})] | "
                f":gray[미정 {mijeong_sum}({mijeong_pct})] | "
                f"**총대수 {total_sum}(100%)**"
            )

            with st.expander(expander_label):
                if sites_list:
                    clean_sites = []
                    for s in sites_list:
                        site_name = str(s[1]) if len(s) > 1 and pd.notna(s[1]) else ""
                        memo = str(s[3]) if len(s) > 3 and pd.notna(s[3]) else ""
                        note = str(s[11]) if len(s) > 11 and pd.notna(s[11]) else ""
                        
                        is_excluded = "통계제외" in note or "통계제외" in memo or "통계제외" in str(s)
                        display_name = f"🔒 [통계제외] {site_name}" if is_excluded else site_name

                        clean_sites.append({
                            "현장명": display_name,
                            "타워회사": s[2] if len(s) > 2 and pd.notna(s[2]) else "",
                            "한노": safe_int(s[4]) if len(s) > 4 else 0,
                            "민노": safe_int(s[5]) if len(s) > 5 else 0,
                            "섬유": safe_int(s[6]) if len(s) > 6 else 0,
                            "건산": safe_int(s[7]) if len(s) > 7 else 0,
                            "직원": safe_int(s[8]) if len(s) > 8 else 0,
                            "미정": safe_int(s[9]) if len(s) > 9 else 0,
                            "총대수": safe_int(s[10]) if len(s) > 10 else 0,
                            "비고": note if note else memo
                        })
                    
                    display_df = pd.DataFrame(clean_sites)
                    st.markdown(f"📍 **{b_name} 상세 현장 목록** (🔒 표기된 현장은 통계에서 제외된 항목입니다)")
                    st.dataframe(display_df, use_container_width=True, hide_index=True)
                else:
                    st.info("해당 지부의 상세 현장 데이터가 없습니다.")


# ------------------------------------------
# [메뉴 3] 타워사별 점유 현황
# ------------------------------------------
elif menu == "🏗️ 3. 타워사별 점유 현황":
    st.title("🏗️ 임대사(타워사)별 점유 현황")
    df3 = sheets[s3_name].copy()

    header_idx = None
    for idx, row in df3.iterrows():
        if "타워사" in row.values:
            header_idx = idx
            break

    if header_idx is not None:
        df3.columns = df3.iloc[header_idx]
        df3 = df3[header_idx + 1:].reset_index(drop=True)

    if "타워사" in df3.columns:
        df3 = df3.dropna(subset=["타워사"]).fillna(0)

        st.subheader("🔍 임대사(타워사) 검색")
        tower_list = ["전체 보기"] + list(df3["타워사"].astype(str).unique())
        selected_tower = st.selectbox("타워사를 선택하세요", tower_list)

        search_kw = st.text_input("또는 타워사 이름 일부 직접 입력", "", placeholder="예: 비엠, 대원 등")

        filtered_df = df3.copy()
        if selected_tower and selected_tower != "전체 보기":
            filtered_df = filtered_df[filtered_df["타워사"].astype(str) == selected_tower]

        if search_kw:
            filtered_df = filtered_df[filtered_df["타워사"].astype(str).str.contains(search_kw, case=False)]

        st.markdown(
            """
            <div style="background-color: #F3F4F6; padding: 10px 14px; border-radius: 8px; font-size: 0.92rem; margin-bottom: 15px;">
                💡 <b>점유율 구분:</b> 
                <span style="color:#2563EB; font-weight:bold;">💙 40% 이상 (우호)</span> | 
                <span style="color:#DC2626; font-weight:bold;">🔴 30% 미만 (비협조)</span>
            </div>
            """,
            unsafe_allow_html=True
        )

        st.markdown(f"### 📋 타워사 목록 (총 {len(filtered_df)}개사)")

        for _, row in filtered_df.iterrows():
            tower_name = str(row["타워사"])
            site_cnt = safe_int(row.get("현장수", 0))
            hanno = safe_int(row.get("한노", 0))
            minno = safe_int(row.get("민노", 0))
            seomoo = safe_int(row.get("섬유", 0))
            geunsan = safe_int(row.get("건산", 0))
            jikwon = safe_int(row.get("직원", 0))
            gita = safe_int(row.get("기타", 0))
            total = safe_int(row.get("합계", 0))

            raw_pct = row.get("한노점유율", 0)
            try:
                f_pct = float(raw_pct)
                pct_val = f_pct * 100 if f_pct <= 1.0 else f_pct
            except:
                pct_val = 0.0

            if pct_val >= 40.0:
                card_class = "tower-card-uho"
                badge_html = f'<span class="badge-uho">💙 우호 {pct_val:.1f}%</span>'
            elif pct_val < 30.0:
                card_class = "tower-card-bihyop"
                badge_html = f'<span class="badge-bihyop">🔴 비협조 {pct_val:.1f}%</span>'
            else:
                card_class = "tower-card-normal"
                badge_html = f'<span class="badge-normal">⚪ 일반 {pct_val:.1f}%</span>'

            card_html = f"""
            <div class="tower-card {card_class}">
                <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 8px;">
                    <span style="font-size: 1.15rem; font-weight: bold; color: #111827;">{tower_name}</span>
                    {badge_html}
                </div>
                <div style="font-size: 0.92rem; color: #4B5563; margin-bottom: 6px;">
                    📍 현장수: <b>{site_cnt}개소</b> | 총 합계대수: <b>{total}대</b>
                </div>
                <div style="font-size: 0.88rem; background-color: rgba(255,255,255,0.7); padding: 8px; border-radius: 6px; line-height: 1.5;">
                    <span style="color:#EF4444; font-weight:bold;">한노 {hanno}대</span> | 
                    <span style="color:#3B82F6; font-weight:bold;">민노 {minno}대</span> | 
                    <span style="color:#10B981; font-weight:bold;">섬유 {seomoo}대</span> | 
                    <span style="color:#F59E0B; font-weight:bold;">건산 {geunsan}대</span><br>
                    <span style="color:#8B5CF6; font-weight:bold;">직원 {jikwon}대</span> | 
                    <span style="color:#6B7280; font-weight:bold;">기타 {gita}대</span>
                </div>
            </div>
            """
            st.markdown(card_html, unsafe_allow_html=True)

    else:
        st.error("3번 시트에서 '타워사' 헤더 항목을 찾을 수 없습니다. 엑셀 파일 형식을 확인해주세요.")


# ------------------------------------------
# [메뉴 4] 반도체 현장 현황
# ------------------------------------------
elif menu == "🏭 4. 반도체 현장 현황":
    st.title("🏭 반도체 현장 타워크레인 현황")
    
    df4_raw = sheets[s4_name].copy()

    parsed_sites = []
    for idx in range(10, len(df4_raw)):
        row = df4_raw.iloc[idx]
        site_name = str(row[1]).strip() if pd.notna(row[1]) else ""
        if not site_name or site_name == "nan" or site_name in ["합 계", "합계"]:
            continue

        total_cnt = safe_int(row[2])
        hanno = safe_int(row[3])
        minno = safe_int(row[4])
        gunsan = safe_int(row[5])
        non_union = safe_int(row[6])
        mijung = safe_int(row[7])
        contact_info = str(row[8]) if pd.notna(row[8]) else ""

        parsed_sites.append({
            "site_name": site_name,
            "total_cnt": total_cnt,
            "hanno": hanno,
            "minno": minno,
            "gunsan": gunsan,
            "non_union": non_union,
            "mijung": mijung,
            "contact_info": contact_info
        })

    tot_d = sum(s["total_cnt"] for s in parsed_sites)
    hanno_cnt = sum(s["hanno"] for s in parsed_sites)
    minno_cnt = sum(s["minno"] for s in parsed_sites)
    gunsan_cnt = sum(s["gunsan"] for s in parsed_sites)
    non_cnt = sum(s["non_union"] for s in parsed_sites)
    mi_cnt = sum(s["mijung"] for s in parsed_sites)

    if tot_d > 0:
        hanno_p = f"{(hanno_cnt / tot_d) * 100:.1f}%"
        minno_p = f"{(minno_cnt / tot_d) * 100:.1f}%"
        gunsan_p = f"{(gunsan_cnt / tot_d) * 100:.1f}%"
        non_p = f"{(non_cnt / tot_d) * 100:.1f}%"
        mi_p = f"{(mi_cnt / tot_d) * 100:.1f}%"
    else:
        hanno_p = minno_p = gunsan_p = non_p = mi_p = "0.0%"

    st.subheader("📊 반도체 현장 전체 대수 비교 (실시간 합산)")
    try:
        st.markdown(
            f"""
        <div class="main-total-card">
            <span style="font-size: 1.1rem; font-weight:bold;">🏆 반도체 현장 전체 대수: </span>
            <span style="font-size: 1.6rem; font-weight:bold; margin-left:8px;">{tot_d}대 (100.0%)</span>
        </div>
        """,
            unsafe_allow_html=True,
        )

        col_s1, col_s2 = st.columns(2)
        with col_s1:
            st.markdown(
                f"""
            <div class="main-stat-card border-hanno">
                <div style="font-size: 0.95rem; font-weight:bold; color:#EF4444;">🔥 한국노총</div>
                <div style="font-size: 1.15rem; font-weight:bold; color:#111827;">{hanno_cnt}대 ({hanno_p})</div>
            </div>
            <div class="main-stat-card border-minno">
                <div style="font-size: 0.95rem; font-weight:bold; color:#3B82F6;">🟦 민주노총</div>
                <div style="font-size: 1.15rem; font-weight:bold; color:#111827;">{minno_cnt}대 ({minno_p})</div>
            </div>
            <div class="main-stat-card border-seomoo">
                <div style="font-size: 0.95rem; font-weight:bold; color:#10B981;">🟩 건설노조(건산,섬유)</div>
                <div style="font-size: 1.15rem; font-weight:bold; color:#111827;">{gunsan_cnt}대 ({gunsan_p})</div>
            </div>
            """,
                unsafe_allow_html=True,
            )
        with col_s2:
            st.markdown(
                f"""
            <div class="main-stat-card border-geunsan">
                <div style="font-size: 0.95rem; font-weight:bold; color:#F59E0B;">🟧 비노조</div>
                <div style="font-size: 1.15rem; font-weight:bold; color:#111827;">{non_cnt}대 ({non_p})</div>
            </div>
            <div class="main-stat-card border-mijeong">
                <div style="font-size: 0.95rem; font-weight:bold; color:#6B7280;">⬛ 미정</div>
                <div style="font-size: 1.15rem; font-weight:bold; color:#111827;">{mi_cnt}대 ({mi_p})</div>
            </div>
            """,
                unsafe_allow_html=True,
            )

    except Exception as e:
        st.info(f"상단 요약 데이터 표시 중 오류 발생: {e}")

    st.markdown("---")
    st.subheader("📍 각 현장별 상세 현황 및 담당자")

    for s in parsed_sites:
        site_name = s["site_name"]
        total_cnt = s["total_cnt"]
        hanno = s["hanno"]
        minno = s["minno"]
        gunsan = s["gunsan"]
        non_union = s["non_union"]
        mijung = s["mijung"]
        contact_info = s["contact_info"]

        with st.expander(f"🏢 {site_name} (총 {total_cnt}대)", expanded=True):
            st.markdown(
                f"• <span style='color:#EF4444; font-weight:bold;'>한국노총:</span> <b>{hanno}대</b> | "
                f"<span style='color:#3B82F6; font-weight:bold;'>민주노총:</span> <b>{minno}대</b> | "
                f"<span style='color:#10B981; font-weight:bold;'>건설노조(건산,섬유):</span> <b>{gunsan}대</b>",
                unsafe_allow_html=True
            )
            st.markdown(
                f"• <span style='color:#F59E0B; font-weight:bold;'>비노조:</span> <b>{non_union}대</b> | "
                f"<span style='color:#6B7280; font-weight:bold;'>미정:</span> <b>{mijung}대</b>",
                unsafe_allow_html=True
            )

            if contact_info and contact_info != "nan":
                phone_match = re.search(r"01[016789][-\s]?\d{3,4}[-\s]?\d{4}", contact_info)
                if phone_match:
                    phone_num = phone_match.group().replace("-", "").replace(" ", "")
                    st.markdown(
                        f"👤 **담당자:** {contact_info}<br>"
                        f'<a href="tel:{phone_num}" class="phone-btn">📞 담당자 바로 전화걸기</a>',
                        unsafe_allow_html=True,
                    )
                else:
                    st.write(f"👤 **담당자:** {contact_info}")


# ------------------------------------------
# [메뉴 5] 월별 신규 채용 추이
# ------------------------------------------
elif menu == "📅 5. 월별 신규 채용 추이":
    st.title("📅 월별 신규 채용 추이 현황")
    
    df5_raw = sheets[s5_name].copy()
    rows = df5_raw.values.tolist()
    
    month_blocks = {}
    current_month = "1월 채용 현황"
    current_data = []

    for r in rows:
        title_cell = str(r[0]).strip() if pd.notna(r[0]) else ""
        if "채용 현황" in title_cell or "채용현황" in title_cell:
            if current_month and current_data:
                month_blocks[current_month] = current_data
            current_month = title_cell
            current_data = []
        elif "임대사" in title_cell or "1~10월" in title_cell:
            continue
        elif current_month:
            if any(pd.notna(cell) for cell in r):
                current_data.append(r)

    if current_month and current_data:
        month_blocks[current_month] = current_data

    if not month_blocks:
        st.warning("월별 채용 데이터를 불러오지 못했습니다. 엑셀 5번 시트의 양식을 확인해주세요.")
    else:
        st.subheader("📌 월별 채용현황 합계 및 비율 요약")

        processed_months = {}
        for m_name, m_rows in month_blocks.items():
            detail_list = []
            for r in m_rows:
                cell0 = str(r[0]).strip() if pd.notna(r[0]) else ""
                if cell0 and "임대사" not in cell0 and "합계" not in cell0 and "비율" not in cell0:
                    detail_list.append({
                        "임대사": r[0],
                        "원청사": r[1] if len(r) > 1 and pd.notna(r[1]) else "",
                        "현장명": r[2] if len(r) > 2 and pd.notna(r[2]) else "",
                        "총대수": safe_int(r[3]) if len(r) > 3 else 0,
                        "한노": safe_int(r[4]) if len(r) > 4 else 0,
                        "민노": safe_int(r[5]) if len(r) > 5 else 0,
                        "기타(건산섬유)": safe_int(r[6]) if len(r) > 6 else 0,
                        "직원": safe_int(r[7]) if len(r) > 7 else 0,
                        "미정": safe_int(r[8]) if len(r) > 8 else 0,
                        "비고": r[9] if len(r) > 9 and pd.notna(r[9]) else "",
                    })
            
            t_tot = sum(d["총대수"] for d in detail_list)
            t_han = sum(d["한노"] for d in detail_list)
            t_min = sum(d["민노"] for d in detail_list)
            t_git = sum(d["기타(건산섬유)"] for d in detail_list)
            t_jik = sum(d["직원"] for d in detail_list)
            t_mi  = sum(d["미정"] for d in detail_list)

            if t_tot > 0:
                p_han = f"{(t_han / t_tot) * 100:.1f}%"
                p_min = f"{(t_min / t_tot) * 100:.1f}%"
                p_git = f"{(t_git / t_tot) * 100:.1f}%"
                p_jik = f"{(t_jik / t_tot) * 100:.1f}%"
                p_mi  = f"{(t_mi / t_tot) * 100:.1f}%"
            else:
                p_han = p_min = p_git = p_jik = p_mi = "0.0%"

            processed_months[m_name] = {
                "details": detail_list,
                "sum": {"총대수": t_tot, "한노": t_han, "민노": t_min, "기타": t_git, "직원": t_jik, "미정": t_mi},
                "pct": {"한노": p_han, "민노": p_min, "기타": p_git, "직원": p_jik, "미정": p_mi}
            }

        for m_name, m_info in processed_months.items():
            s = m_info["sum"]
            p = m_info["pct"]

            expander_title = (
                f"🗓️ **{m_name}** | "
                f"🔥 한노 {s['한노']}({p['한노']}) | "
                f"민노 {s['민노']}({p['민노']}) | "
                f"기타 {s['기타']}({p['기타']}) | "
                f"직원 {s['직원']}({p['직원']}) | "
                f"미정 {s['미정']}({p['미정']}) | "
                f"**총대수 {s['총대수']}대**"
            )

            with st.expander(expander_title):
                if m_info["details"]:
                    st.dataframe(pd.DataFrame(m_info["details"]), use_container_width=True, hide_index=True)
                else:
                    st.info("해당 월의 상세 채용 데이터가 없습니다.")
