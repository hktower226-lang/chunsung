import streamlit as st
import pandas as pd
import os

st.set_page_config(page_title="채용 및 현장 현황 대시보드", layout="wide")

st.title("📊 채용 및 현장 현황 대시보드")

# 불러올 3개 시트 명칭 고정
target_sheets = ["26년 채용비율", "반도체현장", "예전채용비율"]

# 파일 업로드 버튼
uploaded_file = st.file_uploader("엑셀 파일(.xlsx)을 업로드하세요", type=["xlsx", "xls"])

file_to_load = None

# 업로드된 파일이 있으면 우선 사용하고, 없으면 서버 내부의 기본 파일 확인
if uploaded_file is not None:
    file_to_load = uploaded_file
elif os.path.exists("채용비율.xlsx"):
    file_to_load = "채용비율.xlsx"

if file_to_load:
    try:
        excel_file = pd.ExcelFile(file_to_load)
        available_sheets = excel_file.sheet_names

        # Streamlit 상단 탭 생성 (클릭하여 시트 전환)
        tab1, tab2, tab3 = st.tabs([f"📌 {name}" for name in target_sheets])
        tabs_map = {
            "26년 채용비율": tab1,
            "반도체현장": tab2,
            "예전채용비율": tab3
        }

        for sheet_name, tab_obj in tabs_map.items():
            with tab_obj:
                if sheet_name in available_sheets:
                    # 해당 시트 데이터 불러오기
                    df = pd.read_excel(file_to_load, sheet_name=sheet_name)
                    st.dataframe(df, use_container_width=True)
                else:
                    st.warning(f"엑셀 파일 내에 '{sheet_name}' 시트가 없습니다.")

    except Exception as e:
        st.error(f"엑셀 파일을 읽는 중 오류가 발생했습니다: {e}")
else:
    st.info("상단에서 엑셀 파일을 업로드하거나, 깃허브 리포지토리에 '채용비율.xlsx' 파일을 함께 올려주세요.")
