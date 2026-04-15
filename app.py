import streamlit as st
import pandas as pd
from io import BytesIO

# 页面设置
st.set_page_config(
    page_title="Excel合并工具",
    page_icon="📊",
    layout="wide"
)

# 标题
st.title("📊 Excel 批量合并工具")
st.markdown("### 上传多个 Excel 文件，一键合并并下载")
st.divider()

# 上传文件
uploaded_files = st.file_uploader(
    "请选择多个 .xlsx 文件",
    type=["xlsx"],
    accept_multiple_files=True
)

if uploaded_files:
    with st.spinner("正在合并文件..."):
        df_list = []
        for file in uploaded_files:
            df = pd.read_excel(file)
            df_list.append(df)

        # 合并
        merged_df = pd.concat(df_list, ignore_index=True)

        # 展示结果
        st.success(f"✅ 合并完成！共 {len(merged_df)} 行数据")
        st.dataframe(merged_df, use_container_width=True)

        # 生成下载文件
        output = BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            merged_df.to_excel(writer, index=False, sheet_name="合并结果")
        output.seek(0)

        # 下载按钮
        st.download_button(
            label="💾 下载合并后的 Excel",
            data=output,
            file_name="合并结果.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )