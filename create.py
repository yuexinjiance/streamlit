import streamlit as st
from datetime import date
from tools import create_new_company_data, get_new_data, change_brand, change_broken, combine_docx_to_one_pdf


st.title('越鑫检测证书生成 :sunglasses:')
st.subheader(":one: 证书生成 :clipboard:")

full_data_list, short_name_list = get_new_data.get_new_data()

with st.expander("填写公司相关信息："):
    with st.form("form", True):
        col1, col2 = st.columns([1, 1])
        with col1:
            company_name = st.text_input('公司名称：')
        with col2:
            selected_date = st.date_input("检测日期：", date.today())

        col3, col4, col5 = st.columns([1,1, 1])
        with col3:
            all_number = st.number_input('探头数量：', step=1, format="%d")
        with col4:
            temperature = st.number_input('温度：', step=1, format="%d")
        with col5:
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

st.text("")
st.subheader(":two: 品牌及型号修改 :pencil:")
with st.expander("选择品牌及型号"):

    col1, col2 = st.columns([1, 1])
    with col1:
        company_name_change_brand = st.text_input('公司名称：')
    with col2:
        file_num = st.text_input("编号（以空格分割，示例：1-5 6 11）")

    col3, col4 = st.columns([1, 1])
    with col3:
        factory_choice = st.selectbox('制造商', options=short_name_list, key=1)
    with col4:
        type_option = [item['list'] for item in full_data_list if item['name'] == factory_choice][0]
        type_choice = st.selectbox('品牌', options=type_option, key=2)

    if st.button("更改品牌"):
        product_company_full_name = [item['fullname'] for item in full_data_list if item['name'] == factory_choice][0]
        dongzuozhi = [item['alarm'] for item in full_data_list if item['name'] == factory_choice][0]
        with st.spinner("正在更改品牌，请等待..."):
            change_brand.change_all_brand(company_name_change_brand, product_company_full_name, type_choice, dongzuozhi, file_num)
        st.success("品牌更改完毕！")

st.text("")
st.subheader(":three: 故障探头修改 :mute:")
with st.expander("填写故障探头信息"):
    col1, col2 = st.columns([1, 1])
    with col1:
        company_name_change_broken = st.text_input('公司名称：', key='22')
    with col2:
        file_num_broken = st.text_input("编号（以空格分割，示例：1-5 6 12）", key='01')
    if st.button("提交更改"):
        with st.spinner("正在更改故障探头，请等待..."):
            change_broken.change_all_problem_file(company_name_change_broken, file_num_broken)
        st.success("故障探头更改完毕！")

st.text("")
st.subheader(":four: 证书合并 :link:")
company_name_combine = st.text_input('公司名称：', key='12')
if st.button("开始合并"):
    with st.spinner("正在合并证书，请等待..."):
        combine_docx_to_one_pdf.combine_process(company_name_combine)
    st.success("证书合并完成！")