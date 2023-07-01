import streamlit as st
from datetime import date
import sys
import os
from tools import create_new_company_data, get_new_data, change_brand

# 获取 tools 文件夹的绝对路径
tools_path = os.path.abspath("tools")

# 将 tools 文件夹所在的路径添加到 sys.path
sys.path.append(tools_path)

st.title('越鑫检测证书生成')
st.header("1.证书生成")

full_data_list, short_name_list = get_new_data.get_new_data()

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
        press_create = st.form_submit_button("开始生成")

if press_create:
    user_input = {
        "company_name": company_name,
        "all_nums": all_number,
        "date": selected_date.strftime("%Y%m%d"),
        "temperature": temperature,
        "humidity": humidity,
        "sections": setions.strip().split(' '),
        "sections_num": [int(num) for num in setions_number.strip().split(' ')],
        "start_num": start_num
    }
    with st.spinner("正在生成证书，请等待..."):
        create_new_company_data.write_save_all_company(user_input)
    st.success("证书生成完毕！")

st.header("2.品牌及型号修改")
with st.expander("选择品牌及型号"):
    company_name_change_brand = st.text_input('公司名称：')
    file_num = st.text_input("编号（以空格分割，示例：1-5 6 11）")
    col1, col2 = st.columns([1, 1])
    with col1:
        factory_choice = st.selectbox('制造商', options=short_name_list, key=1)
    with col2:
        type_option = [item['list'] for item in full_data_list if item['name'] == factory_choice][0]
        type_choice = st.selectbox('品牌', options=type_option, key=2)
    if st.button("更改品牌"):
        product_company_full_name = [item['fullname'] for item in full_data_list if item['name'] == factory_choice][0]
        dongzuozhi = [item['alarm'] for item in full_data_list if item['name'] == factory_choice][0]
        with st.spinner("正在更改品牌，请等待..."):
            change_brand.change_all_brand(company_name_change_brand, product_company_full_name, type_choice, dongzuozhi, file_num)
        st.success("品牌更改完毕！")
