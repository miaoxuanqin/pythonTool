from docx import Document


def check_empty_heading_8(file_path):
    try:
        doc = Document(file_path)
        paragraphs = doc.paragraphs
        results = []

        for i in range(len(paragraphs)):
            # 检查当前段落是否为 8 级标题
            # 注意：在 Word 中，样式名通常是 "Heading 8" 或者是中文 "标题 8"
            if paragraphs[i].style.name in ['Heading 8', '标题 8']:
                current_heading_text = paragraphs[i].text.strip()

                # 检查下一个段落是否存在
                is_empty = True
                if i + 1 < len(paragraphs):
                    next_para = paragraphs[i + 1]
                    # 如果下一个段落不是标题，且内容不为空，则认为该标题下有正文
                    if not next_para.style.name.startswith('Heading') and \
                            not next_para.style.name.startswith('标题') and \
                            next_para.text.strip():
                        is_empty = False

                if is_empty:
                    results.append(f"第 {i + 1} 行附近 - 标题内容: '{current_heading_text}'")

        # 输出结果
        if results:
            print("--- 发现以下 8 级标题的正文为空 ---")
            for r in results:
                print(r)
        else:
            print("所有 8 级标题下似乎都有内容，或未发现 8 级标题。")

    except Exception as e:
        print(f"读取文件时出错: {e}")


# 执行检测
file_path = r'D:\DESKTOP\111.docx'
check_empty_heading_8(file_path)