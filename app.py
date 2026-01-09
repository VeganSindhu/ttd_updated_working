import streamlit as st
import tempfile
import os
import sys

# 🔴 REQUIRED FOR STREAMLIT CLOUD
sys.path.append(os.path.dirname(__file__))

from ttd_filler_logic import generate_output

st.set_page_config(page_title="TTD Excel Processor", layout="centered")

st.title("📦 TTD Excel Processor")

st.markdown("""
Upload the following files:
1. **TTD Orders Excel**
2. **TTD Postal Excel**

The system will generate the final India Post upload file.
""")

orders_file = st.file_uploader("Upload TTD Orders Excel", type=["xlsx"])
postal_file = st.file_uploader("Upload TTD Postal Excel", type=["xlsx"])

if orders_file and postal_file:
    with st.spinner("Processing files..."):
        with tempfile.TemporaryDirectory() as tmpdir:
            orders_path = os.path.join(tmpdir, "orders.xlsx")
            postal_path = os.path.join(tmpdir, "postal.xlsx")
            output_path = os.path.join(tmpdir, "Matching_Output.xlsx")

            # 🔒 FIXED FILES (NOT USER INPUT)
            template_path = "TTD Template.xlsx"
            volumetric_path = "Volumetric Measurement.xlsx"

            # Save uploads
            with open(orders_path, "wb") as f:
                f.write(orders_file.getbuffer())

            with open(postal_path, "wb") as f:
                f.write(postal_file.getbuffer())

            # 🔴 CORRECT FUNCTION CALL (5 ARGUMENTS)
            count = generate_output(
                orders_path,
                postal_path,
                template_path,
                volumetric_path,
                output_path
            )

            st.success(f"✅ File generated successfully!")
            st.info(f"📄 Number of articles created: **{count}**")

            with open(output_path, "rb") as f:
                st.download_button(
                    label="⬇️ Download Output Excel",
                    data=f,
                    file_name="TTD_Final_Output.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

