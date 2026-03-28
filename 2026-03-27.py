import os
import re
from docx import Document
from docx.shared import Pt
from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph


def set_text_format(paragraph, text):
    """设置正文格式：宋体, 小四, 1.5倍行距, 首行缩进"""
    # 清除段落现有内容并添加新内容
    paragraph.clear()
    run = paragraph.add_run(text.strip())

    # 字体设置
    run.font.name = '宋体'
    run._element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')
    run.font.size = Pt(12)  # 小四
    run.bold = False

    # 段落格式设置
    paragraph.paragraph_format.line_spacing = 1.5  # 1.5倍行距
    paragraph.paragraph_format.first_line_indent = Pt(24)  # 约2字符缩进
    paragraph.paragraph_format.space_before = Pt(0)
    paragraph.paragraph_format.space_after = Pt(0)


def process_txt(txt_path):
    """解析txt，返回 {标题关键信息: {"num": 编号, "body": 正文}}"""
    if not os.path.exists(txt_path):
        return {}

    with open(txt_path, 'r', encoding='utf-8') as f:
        content = f.read()

    # 按照 ##### 分割
    sections = re.split(r'\n(?=#+ )', content)
    data_map = {}

    for sec in sections:
        lines = sec.strip().split('\n')
        if not lines: continue

        # 匹配标题行：提取编号(group 1) 和 标题文字(group 2)
        header_match = re.search(r'#+\s+([\d\.]+)\s+(.*)', lines[0])
        if header_match:
            num = header_match.group(1).strip('.')
            title_text = header_match.group(2).strip()
            body = "\n".join(lines[1:]).strip()
            if body:
                # 使用标题文字作为Key
                data_map[title_text] = {"num": num, "body": body}
    return data_map


def insert_to_word(word_path, data_map):
    doc = Document(word_path)

    # 记录已经匹配到的标题，防止重复插入
    matched_titles = set()

    # 遍历Word文档中的段落
    # 使用列表快照，因为我们会动态增加段落
    paragraphs = list(doc.paragraphs)

    for para in paragraphs:
        para_text = para.text.strip()
        if not para_text: continue

        target_info = None
        # 尝试匹配：1. 标题文字匹配  2. 编号匹配
        for title_key, info in data_map.items():
            # 只要Word里的标题包含txt里的标题文字，或者包含那个长编号
            if (title_key in para_text or info['num'] in para_text) and title_key not in matched_titles:
                target_info = info
                matched_titles.add(title_key)
                break

        if target_info:
            # 提取正文并拆分成多段
            body_paragraphs = target_info['body'].split('\n')

            # 在当前标题段落之后依次插入
            current_cursor = para
            for text in body_paragraphs:
                if not text.strip(): continue

                # 在底层XML层面插入新段落
                new_p_element = doc.add_paragraph()._element
                current_cursor._element.addnext(new_p_element)

                # 将XML元素包装回Paragraph对象以便操作格式
                new_para = Paragraph(new_p_element, doc)
                set_text_format(new_para, text)

                # 移动指针，下一段插在这一段后面
                current_cursor = new_para

            print(
                f"成功填充章节: {target_info['num']} {list(data_map.keys())[list(data_map.values()).index(target_info)]}")

    output_path = word_path.replace(".docx", "_填充完成.docx")
    doc.save(output_path)
    print(f"\n全部处理完成！\n匹配并填充了 {len(matched_titles)} 个章节。\n保存路径: {output_path}")


if __name__ == "__main__":
    # 请确保路径正确
    # TXT_FILE = r'D:\DESKTOP\txt.txt'
    # WORD_FILE = r'D:\DESKTOP\分工.docx'

    TXT_FILE = r'G:\UserDirectory\txt.txt'
    WORD_FILE = r'G:\UserDirectory\分工.docx'

    insert_to_word(WORD_FILE, process_txt(TXT_FILE))