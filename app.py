import streamlit as st
import pandas as pd
import os

# 페이지 제목 설정
st.set_page_config(page_title="현장 점유율 및 통계 관리", layout="wide")

st.title("🏗️ 경기지역본부 현장 점유율 및 통계 조회")
st.markdown("임직원 및 타워사 점유율 현황 조회 전용 프로그램입니다. (조회만 가능)")

# 엑셀 파일 불러오기
current_dir = os.path.dirname(os.path.abspath(__file__))
FILE_NAME = os.path.join(current_dir, "경기지역본부_현장점유율-7.xlsx")

try:
    df_search = pd.read_excel(FILE_NAME, sheet_name="지부별 현장명단", header=1)
    df_total = pd.read_excel(FILE_NAME, sheet_name="전체현황", header=1)
    df_tower = pd.read_excel(FILE_NAME, sheet_name="타워사별", header=1)
    
    df_search = df_search.dropna(subset=['지부', '현장명']).fillna("")
    df_total = df_total.fillna("")
    df_tower = df_tower.fillna("")
except Exception as e:
    st.error(f"엑셀 파일을 읽지 못했습니다. 파일을 확인해주세요: {e}")
    st.stop()

# 탭 메뉴 만들기
tab1, tab2, tab3 = st.tabs(["🔍 1. 지부/현장 검색", "📊 2. 전체 통계", "🏢 3. 타워사별 점유율"])

# [탭 1] 검색 기능
with tab1:
    st.subheader("지부별 현장 명단 검색")
    keyword = st.text_input("지부명을 입력하세요 (예: 북부)", "")
    
    if keyword:
        filtered_df = df_search[df_search['지부'].astype(str).str.contains(keyword)]
    else:
        filtered_df = df_search
        
    st.dataframe(filtered_df, use_container_width=True)

# [탭 2] 전체 통계
with tab2:
    st.subheader("전체 소속별 점유율 현황")
    st.dataframe(df_total.head(2), use_container_width=True)

# [탭 3] 타워사별 점유율
with tab3:
    st.subheader("타워(렌탈)사별 점유 현황")
    st.dataframe(df_tower, use_container_width=True)