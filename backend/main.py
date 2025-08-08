import os

from docx import Document


def identify_docx_labels(docx_folder, zip_from=''):
    """
    遍历指定目录下的所有 .docx 文件，判断每个文件是商务标、技术标还是价格标。
    
    :param directory: str, 待检查的目录路径
    :return: list of dict, 每个字典表示一个文件及其标签
    """

    def check_file_content(doc_path, keywords):
        """检查文件内容中是否包含指定的关键字"""
        try:
            doc = Document(doc_path)
            for paragraph in doc.paragraphs:
                if paragraph.text == keywords:
                    return True
                for run in paragraph.runs:
                    if run.text == keywords:
                        return True
            return False
        except Exception as e:
            print(f"读取文件 {doc_path} 时出错: {e}")
            return False

    result = {}
    # 定义关键词
    labels = {
        "商务": {"filename_keyword": "商务", "content_keyword": "商务应答文件"},
        "技术": {"filename_keyword": "技术", "content_keyword": "技术应答文件"},
        "价格": {"filename_keyword": "价", "content_keyword": "价格应答文件"}
    }

    # 遍历目录下的所有文件
    for rootpath, _, filenames in os.walk(docx_folder):
        for filename in filenames:
            file_path = os.path.join(rootpath, filename)
            matched_labels = []

            # if filename.endswith('.doc'):
            #     cres = doc_to_docx(file_path)
            #     if cres:
            #         os.remove(file_path)
            #         file_path = cres
            #         filename = os.path.basename(file_path)

            # 根据文件名匹配
            for label, info in labels.items():
                if info["filename_keyword"] in filename:
                    matched_labels.append(label)

            if filename.endswith('.docx'):
                for label, info in labels.items():
                    if check_file_content(file_path, info["content_keyword"]):
                        matched_labels.append(label)

            # elif filename.endswith('.doc'):
            #     cres = doc_to_docx(file_path)
            #     if cres:
            #         for label, info in labels.items():
            #             if check_file_content(cres, info["content_keyword"]):
            #                 matched_labels.append(label)

            elif (filename.endswith('.xlsx') or filename.endswith('.xls')) and ('商务' in zip_from or '价' in filename):
                matched_labels.append('价格')

            elif filename.endswith('.dwg'):
                matched_labels.append('技术')

            elif '技术' in zip_from:
                matched_labels.append('技术')

            # 去重
            matched_labels = list(set(matched_labels))
            result[filename] = matched_labels

    return result
