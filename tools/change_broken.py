# 安装修改word文档的库
from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.shared import Pt
from .get_all_problem_files_name import get_all_problem_files_name


# 更改单个探头为故障
def change_problem_file(problem_filename):
    document = Document(problem_filename)

    document.paragraphs[60].runs[2].text = "异常"  # 更改报警功能的值
    document.paragraphs[62].runs[3].text = "/"  # 更改报警动作值
    document.paragraphs[62].runs[4].text = ""
    document.paragraphs[65].runs[3].text = "/"  # 更改重复性
    document.paragraphs[65].runs[4].text = ""
    document.paragraphs[66].runs[3].text = "  /   "  # 更改响应时间
    document.paragraphs[66].runs[4].text = ""

    tables = document.tables
    table = tables[1]
    table.cell(1, 1).text = "/"
    table.cell(1, 1).paragraphs[0].paragraph_format.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.cell(1, 1).paragraphs[0].runs[0].font.size = Pt(12)
    table.cell(1, 2).text = "/"
    table.cell(1, 2).paragraphs[0].paragraph_format.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.cell(1, 2).paragraphs[0].runs[0].font.size = Pt(12)
    table.cell(1, 3).text = "/"
    table.cell(1, 3).paragraphs[0].paragraph_format.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.cell(1, 3).paragraphs[0].runs[0].font.size = Pt(12)
    table.cell(2, 1).text = "/"
    table.cell(2, 1).paragraphs[0].paragraph_format.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.cell(2, 1).paragraphs[0].runs[0].font.size = Pt(12)
    table.cell(2, 2).text = "/"
    table.cell(2, 2).paragraphs[0].paragraph_format.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.cell(2, 2).paragraphs[0].runs[0].font.size = Pt(12)
    table.cell(2, 3).text = "/"
    table.cell(2, 3).paragraphs[0].paragraph_format.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.cell(2, 3).paragraphs[0].runs[0].font.size = Pt(12)
    table.cell(3, 1).text = "/"
    table.cell(3, 1).paragraphs[0].paragraph_format.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.cell(3, 1).paragraphs[0].runs[0].font.size = Pt(12)
    table.cell(3, 2).text = "/"
    table.cell(3, 2).paragraphs[0].paragraph_format.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.cell(3, 2).paragraphs[0].runs[0].font.size = Pt(12)
    table.cell(3, 3).text = "/"
    table.cell(3, 3).paragraphs[0].paragraph_format.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.cell(3, 3).paragraphs[0].runs[0].font.size = Pt(12)

    document.save(problem_filename)


# 更改所有的故障探头
def change_all_problem_file(company_name, product_num_list):
    problem_files_name = get_all_problem_files_name(company_name, product_num_list)
    for problem_file_name in problem_files_name:
        change_problem_file(problem_file_name)


# 查找指定文本内容在模板word文档中的对应位置
def find_text_in_word(text):
    filename = "test.docx"
    document = Document(filename)
    target_text = text
    # 遍历文档中的每个段落
    for paragraph_index, paragraph in enumerate(document.paragraphs):
        # 遍历段落中的每个运行
        for run_index, run in enumerate(paragraph.runs):
            # 判断运行的文本是否与目标文本匹配
            if target_text in run.text:
                # 返回匹配到的段落序号和运行序号
                result = f"段落序号：{paragraph_index}，运行序号：{run_index}"
                return paragraph_index, run_index

        else:
            continue
    else:
        # 如果没有找到目标文本，将 result 设置为 None
        result = None
