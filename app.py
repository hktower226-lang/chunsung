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
        padding: 12px 10px;
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
        padding: 8px;
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
# 1. 상단 고정 영역 (26년 전체 소속별 점유율 - 순서: 한,민,섬,건,직,미,합계)
# ==========================================
st.markdown('<div class="sticky-header">', unsafe_allow_html=True)
st.markdown("### 📌 26년 전체 소속별 점유율")

try:
  # 엑셀 원본 순서대로 추출: 한노, 민노, 섬유, 건산, 직원, 미정, 합계
  hanno_c, minno_c, seom_c, geonsan_c, staff_c, mijung_c, total_c, _ = (
      df.iloc[1, 1:9].values
  )
  hanno_p, minno_p, seom_p, geonsan_p, staff_p, mijung_p = df.iloc[2, 1:7].values

  col1, col2, col3 = st.columns(3)
  with col1:
    st.markdown(
        f"""<div class="metric-box">
        <span class="badge-hanno">한노</span><br>
        <b style="font-size: 1.1rem;">{hanno_c}대</b><br>
        <span style="color: #c62828; font-size: 0.85rem; font-weight: bold;">({float(hanno_p)*100:.1f}%)</span>
        </div>""",
        unsafe_allow_html=True,
    )
    st.markdown(
        f"""<div class="metric-box">
        <span class="badge-geonsan">건산</span><br>
        <b style="font-size: 1.1rem;">{geonsan_c}대</b><br>
        <span style="color: #ef6c00; font-size: 0.85rem; font-weight: bold;">({float(geonsan_p)*100:.1f}%)</span>
        </div>""",
        unsafe_allow_html=True,
    )

  with col2:
    st.markdown(
        f"""<div class="metric-box">
        <span class="badge-minno">민노</span><br>
        <b style="font-size: 1.1rem;">{minno_c}대</b><br>
        <span style="color: #1565c0; font-size: 0.85rem; font-weight: bold;">({float(minno_p)*100:.1f}%)</span>
        </div>""",
        unsafe_allow_html=True,
    )
    st.markdown(
        f"""<div class="metric-box">
        <span class="badge-staff">직원</span><br>
        <b style="font-size: 1.1rem;">{staff_c}대</b><br>
        <span style="color: #7b1fa2; font-size: 0.85rem; font-weight: bold;">({float(staff_p)*100:.1f}%)</span>
        </div>""",
        unsafe_allow_html=True,
    )

  with col3:
    st.markdown(
        f"""<div class="metric-box">
        <span class="badge-seom">섬유</span><br>
        <b style="font-size: 1.1rem;">{seom_c}대</b><br>
        <span style="color: #2e7d32; font-size: 0.85rem; font-weight: bold;">({float(seom_p)*100:.1f}%)</span>
        </div>""",
        unsafe_allow_html=True,
    )
    st.markdown(
        f"""<div class="metric-box">
        <span class="badge-mijung">미정</span><br>
        <b style="font-size: 1.1rem;">{mijung_c}대</b><br>
        <span style="color: #37474f; font-size: 0.85rem; font-weight: bold;">({float(mijung_p)*100:.1f}%)</span>
        </div>""",
        unsafe_allow_html=True,
    )

  st.markdown(
      f"<div style='text-align: center; margin-top: 4px; font-weight: bold; font-size: 1rem;'>총대수 합계: {total_c}대 (100%)</div>",
      unsafe_allow_html=True,
  )

except Exception as e:
  st.warning(f"상단 요약 로딩 중: {e}")

st.markdown("</div>", unsafe_allow_html=True)

# ==========================================
# 데이터 전처리: 지부별 현장 정보 파싱
# ==========================================
raw_data = df.iloc[6:76].copy()
raw_data.columns = [
    "지부",
    "현장명",
    "타워회사",
    "특이사항",
    "한노",
    "민노",
    "섬유",
    "건산",
    "직원",
    "미정",
    "총대수",
    "비고",
]

jibues = ["북부", "남부", "남서", "동부", "서부", "북서", "용인", "중부"]

# ==========================================
# 2. 8개 지역 지부별 소계 및 상세 수치 출력
# ==========================================
st.markdown("### 🏢 8개 지역 지부별 현황 및 소계")
st.markdown(
    "<p style='font-size: 0.9rem; color: gray;'>지부 소계 탭을 누르면 각 현장의 세부 내역이 펼쳐집니다.</p>",
    unsafe_allow_html=True,
)

for jibu in jibues:
  jibu_df = raw_data[raw_data["지부"] == jibu]
  if jibu_df.empty:
    continue

  subtotal_row = jibu_df[jibu_df["현장명"].isna()]
  if not subtotal_row.empty:
    st_idx = subtotal_row.index[0]
    hanno_cnt = raw_data.loc[st_idx, "한노"]
    minno_cnt = raw_data.loc[st_idx, "민노"]
    seom_cnt = raw_data.loc[st_idx, "섬유"]
    geonsan_cnt = raw_data.loc[st_idx, "건산"]
    staff_cnt = raw_data.loc[st_idx, "직원"]
    mijung_cnt = raw_data.loc[st_idx, "미정"]
    total_cnt = raw_data.loc[st_idx, "총대수"]

    # 퍼센테이지 행 추출
    pct_row_idx = st_idx + 1
    if pct_row_idx in raw_data.index:
      hp = (
          float(raw_data.loc[pct_row_idx, "한노"]) * 100
          if pd.notna(raw_data.loc[pct_row_idx, "한노"])
          else 0
      )
      mp = (
          float(raw_data.loc[pct_row_idx, "민노"]) * 100
          if pd.notna(raw_data.loc[pct_row_idx, "민노"])
          else 0
      )
      sp = (
          float(raw_data.loc[pct_row_idx, "섬유"]) * 100
          if pd.notna(raw_data.loc[pct_row_idx, "섬유"])
          else 0
      )
        
