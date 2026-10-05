import pandas as pd
import streamlit as st

# 페이지 설정 (모바일 최적화 레이아웃)
st.set_page_config(
    page_title="26년 채용비율 및 현황", page_icon="📊", layout="centered"
)

# 커스텀 CSS (모바일 가독성을 위한 글씨 키우기 및 상단 고정 스타일)
st.markdown(
    """
    <style>
    /* 전체 폰트 크기 확대 및 모바일 최적화 */
    html, body, [class*="css"] {
        font-size: 16px;
    }
    h1 { font-size: 1.8rem !important; }
    h2 { font-size: 1.5rem !important; }
    h3 { font-size: 1.2rem !important; }
    
    /* 상단 고정 영역 스타일 */
    .sticky-header {
        position: sticky;
        top: 0;
        background-color: white;
        z-index: 999;
        padding: 10px 0;
        border-bottom: 2px solid #ff4b4b;
        margin-bottom: 15px;
    }
    
    /* 한노 퍼센테이지 강조 스타일 */
    .hanno-highlight {
        font-size: 1.2rem;
        font-weight: bold;
        color: #d9534f;
        background-color: #f9f2f4;
        padding: 2px 6px;
        border-radius: 4px;
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
      f"엑셀 파일을 불러오는 중 오류가 발생했습니다. '채용비율.xlsx' 파일이 같은 폴더에 있는지 확인해주세요. 에러: {e}"
  )
  st.stop()

# ==========================================
# 1. 상단 고정 영역 (26년 전체 소속별 점유율)
# ==========================================
st.markdown('<div class="sticky-header">', unsafe_allow_html=True)
st.markdown("### 📌 26년 전체 소속별 점유율 요약")

# 전체 점유율 데이터 추출 (엑셀 행 기준 인덱스 0~2 활용)
try:
  headers = df.iloc[0, 1:9].values  # 소속, 한노, 민노... 합계, 비율
  values_cnt = df.iloc[1, 1:9].values  # 전체 수치
  values_pct = df.iloc[2, 1:7].values  # 비율

  # 요약을 보기 쉽게 카드 형태로 표현
  col1, col2, col3, col4 = st.columns(4)
  with col1:
    st.metric(label="한노", value=f"{values_cnt[0]}대", delta="27.8%")
  with col2:
    st.metric(label="민노", value=f"{values_cnt[1]}대", delta="41.2%")
  with col3:
    st.metric(label="합계", value=f"{values_cnt[6]}대")
  with col4:
    st.metric(label="통계제외", value="현장제외")
except Exception:
  st.info("상단 요약 데이터를 불러오는 중입니다.")

st.markdown("</div>", unsafe_allow_html=True)

st.divider()

# ==========================================
# 데이터 가공: 지부별 현장 및 소계 분리
# ==========================================
# 경기지역본부 현장 데이터 파싱 (행 6부터 75까지)
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

# 지부별로 데이터 그룹화
jibues = [
    "북부",
    "남부",
    "남서",
    "동부",
    "서부",
    "북서",
    "용인",
    "중부",
]

# ==========================================
# 2. 지부별 검색 및 소계 / 세부내역 뷰
# ==========================================
st.markdown("### 🏢 지역 지부별 현황 및 소계")

search_jibู = st.selectbox(
    "🔍 조회할 지부를 선택하세요 (전체보기 가능)", ["전체보기"] + jibues
)

for jibu in jibues:
  if search_jibู != "전체보기" and search_jibู != jibu:
    continue

  # 해당 지부 데이터 필터링
  jibu_df = raw_data[raw_data["지부"] == jibu]
  if jibu_df.empty:
    continue

  # 소계 행 찾기 (현장명이 NaN이거나 '소계'인 행)
  subtotal_row = jibu_df[jibu_df["현장명"].isna()]
  if not subtotal_row.empty:
    st_idx = subtotal_row.index[0]
    # 소계 수치
    hanno_cnt = raw_data.loc[st_idx, "한노"]
    minno_cnt = raw_data.loc[st_idx, "민노"]
    total_cnt = raw_data.loc[st_idx, "총대수"]

    # 퍼센테이지 행 (소계 바로 다음 행)
    pct_row_idx = st_idx + 1
    if pct_row_idx in raw_data.index:
      hanno_pct = (
          float(raw_data.loc[pct_row_idx, "한노"]) * 100
          if pd.notna(raw_data.loc[pct_row_idx, "한노"])
          else 0
      )
    else:
      hanno_pct = (
          (float(hanno_cnt) / float(total_cnt) * 100) if total_cnt > 0 else 0
      )

    # 지부 카드 UI 출력
    with st.container():
      st.markdown(
          f"#### 📍 **{jibu}지부 소계** &nbsp;&nbsp;|&nbsp;&nbsp; 총대수: **{total_cnt}대**"
      )

      col_a, col_b, col_c = st.columns(3)
      with col_a:
        st.markdown(
            f"🔴 한노: <span class='hanno-highlight'>{hanno_cnt}대 ({hanno_pct:.1f}%)</span>",
            unsafe_allow_html=True,
        )
      with col_b:
        st.markdown(f"🔵 민노: **{minno_cnt}대**")
      with col_c:
        st.markdown(f"📦 총대수: **{total_cnt}대**")

      # 세부내역 토글 (평소에는 숨겨져 있음)
      with st.expander(f"📂 [{jibu}지부] 세부 현장 내역 보기"):
        detailed_rows = jibu_df[~jibu_df["현장명"].isna()]
        # 깔끔하게 보기 위해 필요한 컬럼만 추출
        display_cols = [
            "현장명",
            "타워회사",
            "한노",
            "민노",
            "섬유",
            "건산",
            "직원",
            "미정",
            "총대수",
        ]
        st.dataframe(
            detailed_rows[display_cols].reset_index(drop=True),
            use_container_width=True,
        )

      st.markdown("---")

# ==========================================
# 3. 임대사별 한노 점유율 검색 기능
# ==========================================
st.markdown("### 🏗️ 타워(렌탈)사별 한노 점유율 검색")

# 임대사 데이터 파싱 (행 81부터 끝까지)
rental_data = df.iloc[81:].copy()
rental_data.columns = [
    "타워사",
    "현장수",
    "한노",
    "민노",
    "섬유",
    "건산",
    "직원",
    "기타",
    "합계",
    "한노점유율",
    "NaN1",
    "NaN2",
]
rental_data = rental_data.dropna(subset=["타워사"])
rental_data["한노점유율_pct"] = pd.to_numeric(
    rental_data["한노점유율"], errors="coerce"
).fillna(0) * 100

search_rental = st.text_input(
    "🔎 임대사 이름을 검색하세요 (예: 비엠케이, 우정타워 등)"
)

if search_rental:
  filtered_rental = rental_data[
      rental_data["타워사"].str.contains(search_rental, na=False)
  ]
else:
  filtered_rental = rental_data.head(5)  # 검색어 없을 때는 상위 5개만 기본 표시

if not filtered_rental.empty:
  for idx, row in filtered_rental.iterrows():
    st.markdown(
        f"🏢 **{row['타워사']}** (현장수: {row['현장수']}개, 총합계: {row['합계']}대)"
    )
    st.markdown(
        f"👉 한노 점유율: <span class='hanno-highlight'>{row['한노점유율_pct']:.1f}%</span> (한노: {row['한노']}대 / 민노: {row['민노']}대)",
        unsafe_allow_html=True,
    )
    st.markdown("---")
else:
  st.warning("검색 결과가 없습니다.")
