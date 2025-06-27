
import pandas as pd
import streamlit as st
import numpy as np
import io
from io import BytesIO
import xlsxwriter

st.set_page_config(page_title="动态多维拆解分析工具", layout="wide")

st.title("📊 动态多维拆解分析工具")
st.markdown("上传你的 CSV 文件，自动识别维度列，进行结构效应和退费率效应的中心化拆解分析。")

uploaded_file = st.file_uploader("📁 上传 CSV 文件", type=["csv"])

if uploaded_file:
    df = pd.read_csv(uploaded_file)
    st.success("✅ 文件上传成功！")

    numeric_cols = df.select_dtypes(include=[np.number]).columns.tolist()
    dimension_cols = [col for col in df.columns if col not in numeric_cols]

    if len(numeric_cols) < 4:
        st.error("❌ 数值列少于4列，无法拆解。请确保包含：基期在班、当期在班、基期退费、当期退费。")
    else:
        with st.form("config_form"):
            selected_dims = st.multiselect("🔘 请选择拆解维度列（可多选）", options=dimension_cols, default=dimension_cols)
            submit_btn = st.form_submit_button("开始拆解")

        if submit_btn:
            in0, in1, ref0, ref1 = numeric_cols[:4]
            grouped = df.groupby(selected_dims, dropna=False).agg({
                in0: 'sum',
                in1: 'sum',
                ref0: 'sum',
                ref1: 'sum'
            }).reset_index()

            in0_ratio_col = in0 + "占比"
            in1_ratio_col = in1 + "占比"
            rate0_col = ref0.replace("人数", "") + "率"
            rate1_col = ref1.replace("人数", "") + "率"

            sum_in0 = grouped[in0].sum()
            sum_in1 = grouped[in1].sum()
            sum_ref0 = grouped[ref0].sum()
            sum_ref1 = grouped[ref1].sum()

            grouped[in0_ratio_col] = grouped[in0] / sum_in0
            grouped[in1_ratio_col] = grouped[in1] / sum_in1
            grouped[rate0_col] = grouped[ref0] / grouped[in0].replace(0, np.nan)
            grouped[rate1_col] = grouped[ref1] / grouped[in1].replace(0, np.nan)

            R0 = sum_ref0 / sum_in0
            grouped["结构效应(pp)"] = (grouped[in1_ratio_col] - grouped[in0_ratio_col]) * (grouped[rate0_col] - R0) * 100
            grouped["退费率效应(pp)"] = grouped[in1_ratio_col] * (grouped[rate1_col] - grouped[rate0_col]) * 100
            grouped["合计影响(pp)"] = grouped["结构效应(pp)"] + grouped["退费率效应(pp)"]

            total_row = pd.DataFrame({
                **{dim: ["总计"] for dim in selected_dims},
                in0: [sum_in0],
                in1: [sum_in1],
                ref0: [sum_ref0],
                ref1: [sum_ref1],
                in0_ratio_col: [1.0],
                in1_ratio_col: [1.0],
                rate0_col: [sum_ref0 / sum_in0],
                rate1_col: [sum_ref1 / sum_in1],
                "结构效应(pp)": [grouped["结构效应(pp)"].sum()],
                "退费率效应(pp)": [grouped["退费率效应(pp)"].sum()],
                "合计影响(pp)": [grouped["合计影响(pp)"].sum()]
            })

            result = pd.concat([grouped, total_row], ignore_index=True)

            # 四舍五入
            for col in [in0_ratio_col, in1_ratio_col]:
                result[col] = (result[col] * 100).map(lambda x: f"{x:.2f}%")
            for col in [rate0_col, rate1_col]:
                result[col] = (result[col] * 100).map(lambda x: f"{x:.2f}%")
            for col in ["结构效应(pp)", "退费率效应(pp)", "合计影响(pp)"]:
                result[col] = result[col].round(4)

            st.subheader("📄 拆解结果")
            st.dataframe(result, use_container_width=True)

            # 下载为格式化 Excel
            output = BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                result.to_excel(writer, sheet_name='Result', index=False)
                workbook = writer.book
                worksheet = writer.sheets['Result']
                format1 = workbook.add_format({'align': 'center', 'valign': 'vcenter', 'font_name': 'Arial', 'font_size': 11})
                percent_fmt = workbook.add_format({'num_format': '0.00%', 'align': 'center'})
                bold_fmt = workbook.add_format({'bold': True, 'bg_color': '#F0F0F0'})
                for col_num, value in enumerate(result.columns.values):
                    worksheet.set_column(col_num, col_num, 14, format1)
                worksheet.set_row(len(result), None, bold_fmt)

            st.download_button(
                label="📥 下载格式化 Excel",
                data=output.getvalue(),
                file_name="decomposition_result.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
