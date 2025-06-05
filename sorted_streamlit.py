import streamlit as st
import pandas as pd
import numpy as np
import io

# TOPSIS 算法函数
def calculate_topsis(df, selected_columns, direction_list):
    df_vars = df[selected_columns].astype(float)

    # 标准化
    df_std = df_vars.copy()
    for i, col in enumerate(selected_columns):
        if direction_list[i] == 1:
            df_std[col] = (df_vars[col] - df_vars[col].min()) / (df_vars[col].max() - df_vars[col].min())
        else:
            df_std[col] = (df_vars[col].max() - df_vars[col]) / (df_vars[col].max() - df_vars[col].min())

    # 熵权法计算权重
    X = df_std.values
    n, m = X.shape
    P = X / X.sum(axis=0)
    epsilon = 1e-12
    k = 1 / np.log(n)
    E = -k * np.sum(P * np.log(P + epsilon), axis=0)
    d = 1 - E
    w = d / d.sum()

    # 加权规范化
    weighted_matrix = df_std * w
    Z_plus = weighted_matrix.max(axis=0)
    Z_neg = weighted_matrix.min(axis=0)
    D_plus = np.sqrt(((weighted_matrix - Z_plus)**2).sum(axis=1))
    D_minus = np.sqrt(((weighted_matrix - Z_neg)**2).sum(axis=1))
    CR = D_minus / (D_plus + D_minus)

    result = pd.DataFrame({
        "FCIL_CDE": df["FCIL_CDE"].values,
        "经度": df["经度"],
        "纬度": df["纬度"],
        "TOPSIS得分": CR
    })
    min_lat = result['经度'].min()
    min_lon = result['纬度'].min()
    result['Distance'] = np.sqrt((result['经度'] - min_lat)**2 + (result['纬度'] - min_lon)**2)
    result_sorted = result.sort_values(by='Distance').reset_index(drop=True)
    result_sorted = result_sorted.drop(columns='Distance')
    return result_sorted

# Streamlit 网页
st.title("📊 TOPSIS 决策分析工具")

uploaded_file = st.file_uploader("📁 上传 Excel 文件", type=["xlsx"])
if uploaded_file:
    df = pd.read_excel(uploaded_file)
    st.success("文件读取成功！")

    numeric_cols = df.select_dtypes(include=[np.number]).columns.tolist()
    col1, col2 = st.columns([3, 1])
    with col1:
        selected_columns = st.multiselect("📌 请选择参与计算的变量", options=numeric_cols)

    if selected_columns:
        direction_list = []
        for col in selected_columns:
            direction = st.radio(f"{col} 的方向", ["极大化", "极小化"], key=col, horizontal=True)
            direction_list.append(1 if direction == "极大化" else -1)

        if st.button("✅ 开始计算"):
            result_df = calculate_topsis(df, selected_columns, direction_list)
            st.dataframe(result_df)
    
            # 导出下载  

            # 将结果写入内存中的 Excel 文件
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                result_df.to_excel(writer, index=False)
            output.seek(0)  # 回到文件开头

            # 下载按钮
            st.download_button(
                label="📥 下载结果 Excel 文件",
                data=output,
                file_name="topsis_result.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)
