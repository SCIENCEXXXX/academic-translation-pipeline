import fitz
import os
import re
import argparse
import json

def audit_file(md_path, pdf_path, page_num):
    """
    Audits a single markdown file against its corresponding PDF page.
    """
    if not os.path.exists(md_path):
        return {"page": page_num, "status": "ERROR", "message": "MD file not found"}

    with open(md_path, 'r', encoding='utf-8') as f:
        content = f.read()

    issues = []

    # 1. Check Images
    try:
        doc = fitz.open(pdf_path)
        page = doc.load_page(page_num - 1)
        expected_images = len(page.get_images())
        found_images = len(re.findall(r'!\[.*?\]\(.*?\)', content))
        
        if found_images < expected_images:
            issues.append(f"Missing images: Expected {expected_images}, found {found_images}")
        elif found_images > expected_images:
            # Sometimes fitz counts small icons as images, but MD might have more if manually added
            pass
        doc.close()
    except Exception as e:
        issues.append(f"PDF Error: {e}")

    # 2. Check KaTeX Formulas
    # Count inline $ (should be even)
    inline_math = re.findall(r'(?<!\\)\$', content)
    if len(inline_math) % 2 != 0:
        issues.append("Unbalanced inline math markers ($)")

    # Count display $$ (should be even)
    display_math = re.findall(r'\$\$', content)
    if len(display_math) % 2 != 0:
        issues.append("Unbalanced display math markers ($$)")

    # 3. Check for typical broken text patterns (like missing starting chars)
    # This is harder to automate perfectly, but we can look for sentences starting with lowercase or punctuation
    # (though in Chinese this is different)
    
    status = "OK" if not issues else "ISSUE"
    return {
        "page": page_num,
        "status": status,
        "issues": issues
    }

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Audit translated MD files")
    parser.add_argument("--dir", type=str, required=True, help="Directory containing translated MD files")
    parser.add_argument("--pdf", type=str, default="/Users/wangchen/Desktop/Essenti..ng.pdf", help="Path to PDF")
    parser.add_argument("--start", type=int, required=True, help="Start page")
    parser.add_argument("--end", type=int, required=True, help="End page")
    
    args = parser.parse_args()
    
    results = []
    for p in range(args.start, args.end + 1):
        md_file = os.path.join(args.dir, f"translated_page_{p}.md")
        # If not found with that name, try others or just skip
        if not os.path.exists(md_file):
            # Try to find any file that contains the page num? 
            # For now, expect naming convention
            pass
        res = audit_file(md_file, args.pdf, p)
        results.append(res)
    
    print(json.dumps(results, indent=2, ensure_ascii=False))
