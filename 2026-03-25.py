import os
import re
from docx import Document


def main():
    # 路径配置
    txt_path = r"D:\DESKTOP\txt.txt"
    docx_path = r"D:\DESKTOP\目录.docx"
    output_path = r"D:\DESKTOP\目录_填充.docx"

    if not os.path.exists(docx_path):
        print("错误：未找到分工.docx")
        return

    doc = Document(docx_path)

    # 1. 解析 TXT 标题信息
    parsed_items = []
    with open(txt_path, 'r', encoding='utf-8') as f:
        for line in f:
            line = line.strip()
            if not line or line.startswith("["):
                continue

            # 正则匹配开头的数字编号（如 1.5.1.5.3.16.4.4）
            # 逻辑：匹配所有数字和点的组合
            num_match = re.search(r'(\d+(\.\d+)+)', line)
            if num_match:
                num_str = num_match.group(1)
                # --- 核心逻辑：数数字的数量 ---
                # 将 "1.5.1.5.1" 拆分为 ['1', '5', '1', '5', '1']，长度为 5
                level = len(num_str.split('.'))

                # 提取标题文字：删掉数字编号、点号和开头的 # 符号
                # 例如：从 "## 1.5.1.5.1 业务逻辑描述" 提取出 "业务逻辑描述"
                clean_text = re.sub(r'^[#\s\d\.]+', '', line).strip()

                if clean_text:
                    parsed_items.append({"text": clean_text, "level": level})

    # 2. 定位锚点：交通运输行业数据资源可视化管理系统
    target_idx = -1
    keyword = "业务逻辑描述"
    for i, p in enumerate(doc.paragraphs):
        # 移除空格比对文字内容
        if keyword in p.text.replace(" ", ""):
            target_idx = i
            break

    if target_idx == -1:
        print(f"错误：未能在Word中定位到包含 '{keyword}' 的段落")
        return

    # 3. 顺序插入并挂载自动编号样式
    current_pos = doc.paragraphs[target_idx]

    for item in parsed_items:
        # 跳过 1.5.1.5 本身（已经在文档里了）
        if item['text'] == keyword:
            continue

        # 插入新段落（只写入文字）
        new_p = doc.add_paragraph(item['text'])

        # --- 根据数字个数匹配 Word 样式 ---
        level = item['level']
        if level > 9: level = 9  # Word 默认标题样式上限为 9

        # 尝试匹配样式名（支持中英文 Word 环境）
        style_names = [f'标题 {level}', f'Heading {level}']
        applied = False
        for s_name in style_names:
            if s_name in doc.styles:
                new_p.style = doc.styles[s_name]
                applied = True
                break

        # 如果模板没定义该级别，则保持正文并加粗（保底方案）
        if not applied:
            new_p.style = doc.styles['Normal']
            if level <= 6:
                new_p.runs[0].bold = True

        # 将新段落移动到锚点之后
        current_pos._element.addnext(new_p._element)
        current_pos = new_p

    # 4. 保存文件
    try:
        doc.save(output_path)
        print(f"处理完成！已根据数字个数确定标题级别。")
        print(f"生成的 8 级标题已自动关联 Word 样式。请查看：{output_path}")
    except PermissionError:
        print("错误：无法保存文件，请先关闭已打开的 Word。")


if __name__ == "__main__":
    main()