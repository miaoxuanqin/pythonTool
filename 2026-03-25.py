import os
import re
from docx import Document


def main():
    txt_path = r"D:\DESKTOP\txt.txt"
    docx_path = r"D:\DESKTOP\分工.docx"
    output_path = r"D:\DESKTOP\分工_自动编号最终版.docx"

    if not os.path.exists(docx_path):
        print("未找到Word文件")
        return

    doc = Document(docx_path)

    # 1. 解析 TXT：提取文字并计算真实的自动编号级别
    items = []
    with open(txt_path, 'r', encoding='utf-8') as f:
        for line in f:
            line = line.strip()
            if not line or line.startswith("["): continue

            # 提取编号部分：比如 "1.5.1.5.3.16.4.4"
            match_num = re.match(r'^[#\s]*([\d\.]+)', line)
            if match_num:
                num_str = match_num.group(1).strip('.')
                # 计算级别：1.5.1.5 是 4 级, 1.5.1.5.3.16.4.4 是 8 级
                level = num_str.count('.') + 1

                # 提取纯文字：删掉编号和 # 号，只留标题内容
                # 例如从 "## 1.5.1.5.1 业务逻辑描述" 提取出 "业务逻辑描述"
                clean_text = re.sub(r'^[#\s\d\.]+', '', line).strip()

                items.append({"text": clean_text, "level": level})

    # 2. 定位锚点
    target_idx = -1
    keyword = "技术路线（苗）"
    for i, p in enumerate(doc.paragraphs):
        if keyword in p.text.replace(" ", ""):
            target_idx = i
            break

    if target_idx == -1:
        print("未找到 1.5.1.6 章节")
        return

    # 3. 插入标题并挂载自动编号样式
    current_pos = doc.paragraphs[target_idx]

    for item in items:
        # 跳过 1.5.1.5 本身（因为 Word 里已经有了）
        if item['text'] == keyword: continue

        # 创建新段落（只写文字）
        new_p = doc.add_paragraph(item['text'])

        # 核心：应用对应的标题样式名，触发自动编号
        # Word 内置样式名通常为 '标题 1', '标题 2' ... '标题 9'
        level = item['level']
        if level > 9: level = 9  # Word 样式最高支持到 9 级

        style_name = f'标题 {level}'
        try:
            new_p.style = doc.styles[style_name]
        except:
            try:
                new_p.style = doc.styles[f'Heading {level}']
            except:
                # 如果模板里没定义这么深的样式，则手动设置缩进并加粗
                new_p.style = doc.styles['Normal']
                run = new_p.runs[0]
                run.bold = True if level < 6 else False

        # 移动节点
        current_pos._element.addnext(new_p._element)
        current_pos = new_p

    doc.save(output_path)
    print(f"处理完成！请查看：{output_path}")


if __name__ == "__main__":
    main()