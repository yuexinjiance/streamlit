import streamlit as st
from datetime import date


st.title('越鑫检测证书生成2023-06-30')
st.header("1.证书生成")
with st.expander("填写公司相关信息："):
    with st.form("form", True):
        company_name = st.text_input('公司名称：')
        all_number = st.number_input('探头数量：', step=1, format="%d")
        selected_date = st.date_input("选择日期", date.today())
        has_guessed = st.form_submit_button("Submit Guess!")

