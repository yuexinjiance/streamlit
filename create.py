import streamlit as st
from datetime import date


st.title('越鑫检测证书生成')
st.header("1.证书生成")

with st.expander("填写公司相关信息："):
    with st.form("form", True):
        company_name = st.text_input('公司名称：')
        all_number = st.number_input('探头数量：', step=1, format="%d")
        selected_date = st.date_input("选择日期", date.today())
        temperature = st.number_input('温度：', step=1, format="%d")
        humidity = st.number_input('湿度：', step=1, format="%d")
        setions = st.text_input('探头分布区域（空格分隔）：')
        setions_number = st.text_input('各区域探头数量（空格分隔）：')
        start_num = st.number_input('证书起始编号：', value=1, step=1, format="%d")
        has_guessed = st.form_submit_button("开始生成")

if has_guessed:
    with st.sidebar:
        st.header('弹窗内容')
        st.write('这是一个弹窗的内容。')

