import streamlit as st
from terrarosa_order_processor import main
import tempfile, os
from datetime import datetime

st.title("☕ 테라로사 주문취합 자동화")

st.write("📁 파일 두 개를 올려주세요")

order_file = st.file_uploader("① 주문취합 엑셀", type=["xlsx"])
code_file  = st.file_uploader("② 자사몰상품코드 엑셀", type=["xlsx"])

if order_file and code_file:
    if st.button("🚀 취합 시작!"):
        with tempfile.TemporaryDirectory() as tmp:
            o_path = os.path.join(tmp, "order.xlsx")
            c_path = os.path.join(tmp, "code.xlsx")
            out_path = os.path.join(tmp, "result.xlsx")

            with open(o_path, "wb") as f:
                f.write(order_file.read())
            with open(c_path, "wb") as f:
                f.write(code_file.read())

            main(o_path, c_path, out_path)

            today = datetime.today().strftime("%Y%m%d")

            with open(out_path, "rb") as f:
                st.download_button(
                    "📥 완성 파일 다운로드",
                    data=f,
                    file_name=f"자사몰주문취합_{today}.xlsx"
                )
