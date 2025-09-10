import streamlit as st
import pandas as pd
import io

# Set page title and icon
st.set_page_config(
    page_title="경비예산 Merge Tool",
    page_icon="💸"
)

# App title and description
st.title("💸 경비예산 Merge Tool")
st.markdown("왼쪽 사이드바에서 엑셀 파일을 업로드하면, 필요한 데이터를 추출하여 새로운 형식으로 변환해 드립니다.")

# Sidebar for file upload and year input
st.sidebar.header("엑셀 파일 업로드")
uploaded_file = st.sidebar.file_uploader(
    "여기에 파일을 끌어다 놓거나 클릭하여 업로드하세요.", 
    type=["xls", "xlsx"]
)

# User input for the year
year_input = st.sidebar.text_input(
    "데이터의 년도를 입력하세요 (예: 2024)", 
    pd.Timestamp.now().year
)

if uploaded_file is not None:
    try:
        # Read the uploaded Excel file, skipping to the 6th row for the header
        # The header is at index 5 (0-indexed)
        df_original = pd.read_excel(uploaded_file, header=5)

        st.subheader("업로드된 원본 데이터 미리보기")
        st.dataframe(df_original)

        # Sanity check: Ensure required columns exist
        required_cols = ["비용센터코드", "계정코드"]
        if not all(col in df_original.columns for col in required_cols):
            st.error("업로드된 파일에 '비용센터코드' 또는 '계정코드' 컬럼이 없습니다. 올바른 형식의 파일을 업로드해 주세요.")
        else:
            # Find all columns that contain '월' to identify month columns
            month_cols = [col for col in df_original.columns if "월" in str(col)]

            # Melt the DataFrame to long format to handle month columns
            # This turns the month columns into rows, with a new column for the month name and its value
            df_melted = pd.melt(
                df_original,
                id_vars=["비용센터코드", "계정코드"],
                value_vars=month_cols,
                var_name="월",
                value_name="고정비금액"
            )

            # Drop rows where '고정비금액' is NaN
            df_melted.dropna(subset=['고정비금액'], inplace=True)

            # Convert month names to two-digit month strings (e.g., '1월' -> '01')
            df_melted["월_str"] = df_melted["월"].str.extract(r"(\d+)").astype(int).astype(str).str.zfill(2)

            # Create the '계획년월' column in YYYYMM format
            if year_input:
                df_melted["계획년월"] = year_input + df_melted["월_str"]
            else:
                st.warning("년도가 입력되지 않아 '계획년월' 컬럼을 생성할 수 없습니다.")
                df_melted["계획년월"] = ""

            # Add required columns
            df_melted["차대구분코드"] = "1"
            df_melted["원가요소코드"] = df_melted["계정코드"]

            # Select and reorder the final columns
            df_final = df_melted[[
                "계획년월",
                "비용센터코드",
                "차대구분코드",
                "원가요소코드",
                "고정비금액"
            ]]

            # Convert numeric columns to string type for consistency with requirement
            df_final['계획년월'] = df_final['계획년월'].astype(str)
            df_final['비용센터코드'] = df_final['비용센터코드'].astype(str)
            df_final['차대구분코드'] = df_final['차대구분코드'].astype(str)
            df_final['원가요소코드'] = df_final['원가요소코드'].astype(str)

            st.subheader("변환된 데이터 미리보기")
            st.dataframe(df_final)

            # Create an in-memory Excel file for download
            buffer = io.BytesIO()
            df_final.to_excel(buffer, index=False, engine="openpyxl")
            buffer.seek(0)

            # Download button for the new Excel file
            st.download_button(
                label="결과 다운로드",
                data=buffer,
                file_name="경비예산_병합_결과.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                help="변환된 데이터를 새로운 엑셀 파일로 다운로드합니다."
            )

    except Exception as e:
        st.error(f"파일 처리 중 오류가 발생했습니다: {e}")
        st.info("파일 형식이나 내용이 올바른지 확인해 주세요.")

else:
    st.info("시작하려면 왼쪽 사이드바에 엑셀 파일을 업로드하세요.")
    st.markdown("---")
    st.markdown("""
    **사용 방법:**
    1. 왼쪽 사이드바에서 '찾아보기' 버튼을 클릭하여 엑셀 파일을 선택하거나 파일을 끌어다 놓으세요.
    2. 데이터에 맞는 년도를 입력하세요.
    3. 업로드가 완료되면 변환된 데이터와 다운로드 버튼이 나타납니다.
    """)
