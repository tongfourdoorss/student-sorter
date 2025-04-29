import streamlit as st
import pandas as pd
import os
from io import BytesIO

st.title("📊 เรียงลำดับชื่อนักเรียนตามเลขประจำตัว")

uploaded_file = st.file_uploader("📂 อัปโหลดไฟล์ Excel (.xlsx หรือ .xls)", type=["xlsx", "xls"])

if uploaded_file:
    file_ext = os.path.splitext(uploaded_file.name)[1]

    try:
        if file_ext == ".xls":
            df = pd.read_excel(uploaded_file, engine="xlrd")
        else:
            df = pd.read_excel(uploaded_file, engine="openpyxl")
    except Exception as e:
        st.error(f"❌ ไม่สามารถอ่านไฟล์ได้: {e}")
    else:
        required_columns = {"ชื่อ", "เพศ", "เลขประจำตัว", "ชั้น", "ห้อง"}
        if not required_columns.issubset(df.columns):
            st.error(f"❌ ไฟล์ต้องมีคอลัมน์: {required_columns}")
        else:
            output = BytesIO()
            with pd.ExcelWriter(output, engine="openpyxl") as writer:
                levels = ["อนุบาล 1", "อนุบาล 2", "อนุบาล 3",
                          "ป.1", "ป.2", "ป.3", "ป.4", "ป.5", "ป.6"]
                for level in levels:
                    for room in [1, 2, 3]:
                        class_df = df[(df["ชั้น"] == level) & (df["ห้อง"] == room)]
                        if class_df.empty:
                            continue
                        class_df["เลขประจำตัว"] = class_df["เลขประจำตัว"].astype(int)
                        class_df = class_df.sort_values(by="เลขประจำตัว")

                        males = class_df[class_df["เพศ"] == "ชาย"]
                        females = class_df[class_df["เพศ"] == "หญิง"]

                        males.to_excel(writer, sheet_name=f"{level}-ห้อง{room}-ชาย", index=False)
                        females.to_excel(writer, sheet_name=f"{level}-ห้อง{room}-หญิง", index=False)

            st.success("✅ เรียบร้อย! กดปุ่มด้านล่างเพื่อดาวน์โหลดไฟล์")

            st.download_button(
                label="📥 ดาวน์โหลดไฟล์ที่จัดเรียงแล้ว",
                data=output.getvalue(),
                file_name="students_sorted.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
