
import os
import re

project_root = os.path.expanduser("~/Desktop/Book_Translation_Project")
translated_dir = os.path.join(project_root, "translated")
images_dir = os.path.join(project_root, "images")
output_master = os.path.join(translated_dir, "BOOK_Master_Final.md")

# 获取翻译文件列表并按页码排序
# 之前我们是 Cat 出来的，现在我们按原始章节小块读取
files = [
    "Frontmatter_CN.md",
    "Chapter_01_Complete_CN.md",
    "Chapter_02_Part1_CN.md",
    "Chapter_02_Part2_CN.md",
    "Chapter_02_Part3_CN.md",
    "Chapter_02_Part4_CN.md",
    "Chapter_02_Part5_A.md",
    "Chapter_02_Part5_B.md",
    "Chapter_03_Part1_CN.md",
    "Chapter_03_Part2_CN.md",
    "Chapter_04_Part1_CN.md",
    "Chapter_04_Part2_CN.md",
    "Chapter_05_Part1_CN.md",
    "Chapter_05_Part2_CN.md"
]

full_content = ""

for f_name in files:
    f_path = os.path.join(translated_dir, f_name)
    if not os.path.exists(f_path): continue
    
    with open(f_path, 'r', encoding='utf-8') as f:
        content = f.read()
        
    # 💡 强制注入缺失的图片标签逻辑：
    # 我们根据译文中的 "Page XXX" 标记，找到对应页码，然后检查 images 文件夹
    def inject_images(match):
        page_num = match.group(1)
        # 搜寻该页所有的图片碎片
        img_tags = ""
        for i in range(10): # 假设一页最多10个图
            img_file = f"page_{{page_num}}_img_{i}.jpeg"
            if os.path.exists(os.path.join(images_dir, img_file)):
                img_tags += f"\n\n![图 {page_num}-{i}](images/{img_file})\n\n"
        return f"\n\n--- [Page {page_num} Start] ---\n\n{img_tags}"

    # 先清理掉可能存在的旧标签（防止重复）
    content = re.sub(r'!\\[.*?\\]\(images/.*?\\]', '', content)
    # 注入新标签
    content = re.sub(r'Page (\\d+)', inject_images, content)
    
    full_content += content + "\n\n---\n\n"

with open(output_master, 'w', encoding='utf-8') as f:
    f.write(full_content)

print(f"✅ 全书终极合稿已重塑：{output_master}")
