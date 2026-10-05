import pandas as pd
import streamlit as st

# 페이지 설정 (모바일 최적화 레이아웃)
st.set_page_config(
    page_title="26년 채용비율 및 현황", page_icon="📊", layout="centered"
)

# 모바일 가독성 및 색상 구분을 위한 CSS 스타일링
st.markdown(
    """
    <style>
    html, body, [class*="css"] {
        font-size: 16px;
    }
    
    /* 상단 고정 영역 스타일 */
    .sticky-header {
        position: sticky;
        top: 0;
        background-color: #ffffff;
        z-index: 999;
        padding: 10px;
        border-bottom: 3px solid #ff4b4b;
        box-shadow: 0 4px 6px rgba(0,0,0,0.05);
        margin-bottom: 15px;
    }
    
    /* 소속별 뱃지 스타일 */
    .badge-hanno { background-color: #ffebee; color: #c62828; padding: 3px 6px; border-radius: 4px; font-weight: bold; }
    .badge-minno { background-color: #e3f2fd; color: #1565c0; padding: 3px 6px; border-radius: 4px; font-weight: bold; }
    .badge-seom { background-color: #e8f5e9; color: #2e7d32; padding: 3px 6px; border-radius: 4px; font-weight: bold; }
    .badge-geonsan { background-color: #fff3e0; color: #ef6c00; padding: 3px 6px; border-radius: 4px; font-weight: bold; }
    .badge-staff { background-color: #f3e5f5; color: #7b1fa2; padding: 3px 6px; border-radius: 4px; font-weight: bold; }
    .badge-mijung { background-color: #eceff1; color: #37474f; padding: 3px 6px; border-radius: 4px; font-weight: bold; }
    
    .metric-box {
        background-color: #f8f9fa;
        border-radius: 8px;
        padding: 6px;
        text-align: center;
        border: 1px solid #e9ecef;
        margin-bottom: 6px;
    }
    </style>
""",
    unsafe_allow_html=True,
)


@st.cache_data
def load_data():
  excel_path = "채용비율.xlsx"
  df = pd.read_excel(excel_path, sheet_name="채용비율")
  return df


try:
  df = load_data()
except Exception as e:
  st.error(
      f"엑셀 파일을 찾을 수 없습니다. '채용비율.xlsx' 파일이 같은 폴더에 있는지 확인해주세요. 에러: {e}"
  )
  st.stop()

# ==========================================
# 1. 상단 고정 영역 (순서: 한노, 민노, 섬유, 건산, 직원, 미정, 합계)
# ==========================================
st.markdown('<div class="sticky-header">', unsafe_allow_html=True)
st.markdown("### 📌 26년 전체 소속별 점유율")

try:
  hanno_c, minno_c, seom_c, geonsan_c, staff_c, mijung_c, total_c, _ = (
      df.iloc[1, 1:9].values
  )
  hanno_p, minno_p, seom_p, geonsan_p, staff_p, mijung_p = df.iloc[2, 1:7].values

  # 첫 번째 줄: 한노, 민노, 섬유
  r1_c1, r1_c2, r1_c3 = st.columns(3)
  with r1_c1:
    st.markdown(
        f"""<div class="metric-box">
        <span class="badge-hanno">한노</span><br>
        <b style="font-size: 1.1rem;">{hanno_c}대</b><br>
        <span style="color: #c62828; font-size: 0.85rem; font-weight: bold;">({float(hanno_p)*100:.1f}%)</span>
        </div>""",
        unsafe_allow_html=True,
    )
  with r1_c2:
    st.markdown(
        f"""<div class="metric-box">
        <span class="badge-minno">민노</span><br>
        <b style="font-size: 1.1rem;">{minno_c}대</b><br>
        <span style="color: #1565c0; font-size: 0.85rem; font-weight: bold;">({float(minno_p)*100:.1f}%)</span>
        </div>""",
        unsafe_allow_html=True,
    )
  with r1_c3:
    st.markdown(
        f"""<div class="metric-box">
        <span class="badge-seom">섬유</span><br>
        <b style="font-size: 1.1rem;">{seom_c}대</b><br>
        <span style="color: #2e7d32; font-size: 0.85rem; font-weight: bold;">({float(seom_p)*100:.1f}%)</span>
        </div>""",
        unsafe_allow_html=True,
    )

  # 두 번째 줄: 건산, 직원, 미정
  r2_c1, r2_c2, r2_c3 = st.columns(3)
  with r2_c1:
    st.markdown(
        f"""<div class="metric-box">
        <span class="badge-geonsan">건산</span><br>
        <b style="font-size: 1.1rem;">{geonsan_c}대</b><br>
        <span style="color: #ef6c00; font-size: 0.85rem; font-weight: bold;">({float(geonsan_p)*100:.1f}%)</span>
        </div>""",
        unsafe_allow_html=True,
    )
  with r2_c2:
    st.markdown(
        f"""<div class="metric-box">
        <span class="badge-staff">직원</span><br>
        <b style="font-size: 1.1rem;">{staff_c}대</b><br>
        <span style="color: #7b1fa2; font-size: 0.85rem; font-weight: bold;">({float(staff_p)*10
