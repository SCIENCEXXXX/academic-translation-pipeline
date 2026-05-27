
import fitz
import os
import re
import json

def audit_translation(md_path, start_page, end_page):
    if not os.path.exists(md_path):
        return [{"error": f"File not found: {md_path}"}]
        
    with open(md_path, 'r', encoding='utf-8') as f:
        md_content = f.read()
    
    pdf_path = os.path.expanduser('~/Desktop/Book_Project/Original_Book.pdf')
    doc = fitz.open(pdf_path)
    report = []

    for p in range(start_page - 1, end_page):
        page_num = p + 1
        page = doc.load_page(p)
        real_img_count = len(page.get_images())
        
        # 查找 MD 中是否存在该页的图片标签
        # 匹配格式: page_XX_img_X
        tags_found = re.findall(f'page_{{page_num}}_img_\d+', md_content)
        
        if real_img_count > 0:
            if len(tags_found) == 0:
                report.append({"page": page_num, "issue": "MISSING_ALL", "expected": real_img_count, "found": 0})
            elif len(tags_found) < real_img_count:
                report.append({"page": page_num, "issue": "MISSING_PARTS", "expected": real_img_count, "found": len(tags_found)})
            
    return report

if __name__ == "__main__":
    import sys
    # 示例用法: python3 auditor.py <md_file> <start_page> <end_page>
    if len(sys.argv) > 3:
        res = audit_translation(sys.argv[1], int(sys.argv[2]), int(sys.argv[3]))
        print(json.dumps(res, indent=2))
    else:
        print("Usage: python3 auditor.py <md_file> <start_page> <end_page>")
