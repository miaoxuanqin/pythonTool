import os
import re
from docx import Document
from docx.shared import Pt
from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph  # 导入段落包装类


def solve_task():
    txt_path = r'D:\DESKTOP\txt.txt'
    docx_path = r'D:\DESKTOP\目录 模板 苗 内容空白.docx'
    output_path = r'D:\DESKTOP\更新后的文档_最终版.docx'

    if not os.path.exists(txt_path) or not os.path.exists(docx_path):
        print("错误：未找到文件，请检查路径。")
        return

    # 1. 解析 TXT 文件
    txt_data = {}
    current_title = None
    current_content = []

    try:
        with open(txt_path, 'r', encoding='utf-8-sig') as f:
            lines = f.readlines()
    except UnicodeDecodeError:
        with open(txt_path, 'r', encoding='gbk') as f:
            lines = f.readlines()

    # 提取逻辑：匹配标题及其下方的正文
    for line in lines:
        line_strip = line.strip()
        header_match = re.match(r'^#+\s+([\d\.]+)\s*(.*)', line_strip)
        if header_match:
            if current_title:
                txt_data[current_title] = current_content
            current_title = header_match.group(2).strip()
            current_content = []
        else:
            if current_title is not None and line_strip:
                current_content.append(line_strip)

    if current_title:
        txt_data[current_title] = current_content

    # 2. 操作 Word 文档
    doc = Document(docx_path)

    # 遍历当前文档中的段落（这些是标题所在段落）
    for para in doc.paragraphs:
        para_text = para.text.strip()

        if para_text in txt_data:
            print(f"匹配成功，正在插入正文: {para_text}")
            contents = txt_data[para_text]

            # 为了保证正文顺序，我们需要从最后一行开始“紧跟标题”插入
            # 或者记录最后一次插入的位置。这里采用“紧跟标题倒序插入”逻辑：
            for text_line in reversed(contents):
                # 创建一个空段落元素
                new_p_element = doc.add_paragraph()._element
                # 将该元素移到当前标题段落的下方
                para._element.addnext(new_p_element)
                # **核心修复**：直接将 XML 元素包装成 Paragraph 对象，不再搜索列表
                p_obj = Paragraph(new_p_element, doc)

                # --- 设置段落格式 ---
                # 小四号字 = 12磅，2字符缩进 = 24磅
                p_obj.paragraph_format.first_line_indent = Pt(24)
                # 也可以设置行间距（可选，例如1.5倍行距）：
                # p_obj.paragraph_format.line_spacing = 1.5

                # --- 设置字体格式 ---
                run = p_obj.add_run(text_line)
                run.font.name = 'SimSun'  # 英文字体名
                run._element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')  # 中文字体名
                run.font.size = Pt(12)  # 小四

    # 3. 保存
    try:
        doc.save(output_path)
        print(f"\n处理完成！")
        print(f"结果已保存至: {output_path}")
    except PermissionError:
        print(f"保存失败：请关闭已打开的 Word 文件 {output_path} 后重试。")


if __name__ == "__main__":
    solve_task()