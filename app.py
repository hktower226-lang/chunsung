import streamlit as st
import pandas as pd

st.set_page_config(
    page_title="2026 채용 현황 대시보드",
    page_icon="📱",
    layout="centered", # 모바일 디스플레이 최적화
    initial_sidebar_state="collapsed"
)

# 모바일 전용 스타일 커스텀
st.markdown("""
<style>
    .block-container {
        padding-top: 1rem;
        padding-bottom: 2rem;
        padding-left: 0.8rem;
        padding-right: 0.8rem;
    }
    
    .main-card {
        background: linear-gradient(135deg, #1e3c72 0%, #2a5298 100%);
        color: white;
        padding: 16px;
        border-radius: 12px;
        box-shadow: 0 4px 10px rgba(0,0,0,0.15);
        margin-bottom: 15px;
    }
    
    .branch-card {
        background-color: #f8f9fa;
        border-left: 5px solid #2a5298;
        padding: 12px;
        border-radius: 8px;
        margin-bottom: 8px;
    }
</style>
""", unsafe_content_policy=True)

# 깃허브 원본 엑셀 RAW URL (매일 수정 시 자동 반영)
GITHUB_EXCEL_URL = "https://raw.githubusercontent.com/chunsung/chunsung/main/%EC%B2%84%EC%9A%A9%EB%B9%84%EC%9C%A8.xlsx"

@st.cache_data(ttl=60) # 1분 주기로 엑셀 변경사항 캐시 갱신
def load_excel_data():
    try:
        xl = pd.ExcelFile(GITHUB_EXCEL_URL)
        df_total = pd.read_excel(xl, sheet_name='채용비율')
        return df_total
    except Exception as e:
        st.error(f"GitHub 엑셀 연동 오류: {e}")
        return None

df_raw = load_excel_data()

if df_raw is not None:
    st.title("📱 2026년 채용 현황")
    st.caption("🔄 GitHub 연동 (매일 자동 갱신)")

    # 1. 2026년 전체 소속별 점유율 (메인 고정)
    st.markdown("### 📊 2026년 전체 소속별 점유율")
    
    try:
        cnt_vals = df_raw.iloc[1, :9].tolist()
        pct_vals = df_raw.iloc[2, :9].tolist()
        total_cnt = cnt_vals[7]
        
        st.markdown(f"""
        <div class="main-card">
            <div style="font-size:1rem; opacity:0.9;">총 채용 대수</div>
            <div style="font-size:1.8rem; font-weight:bold;">{total_cnt} 대 <span style="font-size:1rem; font-weight:normal;">(100%)</span></div>
        </div>
        """, unsafe_content_policy=True)

        cols1 = st.columns(2)
        cols2 = st.columns(2)
        cols3 = st.columns(2)
        
        unions = [
            ("한노", cnt_vals[1], pct_vals[1]),
            ("민노", cnt_vals[2], pct_vals[2]),
            ("직원", cnt_vals[5], pct_vals[5]),
            ("건산", cnt_vals[4], pct_vals[4]),
            ("미정", cnt_vals[6], pct_vals[6]),
            ("섬유", cnt_vals[3], pct_vals[3]),
        ]
        
        grid_cols = cols1 + cols2 + cols3
        for i, (u_name, u_cnt, u_pct) in enumerate(unions):
            try:
                pct_str = f"{float(u_pct)*100:.1f}%"
            except:
                pct_str = str(u_pct)
            with grid_cols[i]:
                st.metric(
                    label=f"🔴 {u_name}" if u_name in ["한노", "민노"] else f"⚪ {u_name}", 
                    value=f"{u_cnt}대", 
                    delta=pct_str
                )
    except Exception as e:
        st.warning("전체 점유율 데이터 로딩 중 오류가 발생했습니다.")

    st.divider()

    # 2. 경기지역본부 현장 점유율 현황
    st.markdown("### 🏢 경기지역본부 지부별 점유율 현황")

    branch_data = {}
    current_branch_sites = []

    for idx in range(6, len(df_raw)):
        row = df_raw.iloc[idx].tolist()
        b_col = str(row[0]).strip() if pd.notna(row[0]) else ""

        if "소계" in b_col:
            b_name = b_col.replace("소계", "").strip()
            try:
                hanno = int(row[4]) if pd.notna(row[4]) else 0
                minno = int(row[5]) if pd.notna(row[5]) else 0
                seom = int(row[6]) if pd.notna(row[6]) else 0
                geonsan = int(row[7]) if pd.notna(row[7]) else 0
                staff = int(row[8]) if pd.notna(row[8]) else 0
                mijeong = int(row[9]) if pd.notna(row[9]) else 0
                total = int(row[10]) if pd.notna(row[10]) else (hanno + minno + seom + geonsan + staff + mijeong)
            except:
                hanno = minno = seom = geonsan = staff = mijeong = total = 0

            branch_data[b_name] = {
                "summary": {
                    "총대수": total, "한노": hanno, "민노": minno, 
                    "섬유": seom, "건산": geonsan, "직원": staff, "미정": mijeong
                },
                "sites": current_branch
