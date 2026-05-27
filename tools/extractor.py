import fitz
import os
import argparse

def extract_pages(pdf_path, output_dir, image_dir, start_page, end_page):
    os.makedirs(output_dir, exist_ok=True)
    os.makedirs(image_dir, exist_ok=True)

    try:
        doc = fitz.open(pdf_path)
    except Exception as e:
        print(f"Error opening PDF: {e}")
        return

    for p in range(start_page - 1, min(end_page, doc.page_count)):
        page_num = p + 1
        page = doc.load_page(p)
        
        blocks = page.get_text("blocks")
        img_info = page.get_image_info(xrefs=True)
        drawings = page.get_drawings()
        
        content_elements = []
        # 1. 文本块
        for b in blocks:
            if b[6] == 0: 
                content_elements.append({
                    "type": "text", "y": b[1], "content": b[4]
                })
        
        # 2. 位图提取
        has_real_images = False
        for i, img in enumerate(img_info):
            xref = img.get('xref')
            if not xref: continue
            bbox = img['bbox']
            try:
                base_image = doc.extract_image(xref)
                image_name = f"page_{page_num}_img_{i}.{base_image['ext']}"
                with open(os.path.join(image_dir, image_name), "wb") as f:
                    f.write(base_image["image"])
                
                content_elements.append({
                    "type": "image", "y": bbox[1],
                    "content": f"\n\n![Image page_{page_num}_img_{i}]({image_name})\n\n"
                })
                has_real_images = True
            except: continue

        # 3. 矢量图补偿 (关键修复)
        # 如果没有位图但有复杂的绘图路径，我们将整个页面渲染为一个图片，供翻译对齐
        if not has_real_images and len(drawings) > 10:
            image_name = f"page_{page_num}_draw.jpeg"
            # 提高精度渲染
            pix = page.get_pixmap(matrix=fitz.Matrix(2, 2)) 
            pix.save(os.path.join(image_dir, image_name))
            content_elements.append({
                "type": "image", "y": 50, # 放在页首附近供 AI 识别
                "content": f"\n\n![Image page_{page_num}_draw]({image_name})\n\n"
            })

        # 4. 排序并生成
        content_elements.sort(key=lambda x: x["y"])
        page_text = "".join([el["content"] for el in content_elements])
        
        with open(os.path.join(output_dir, f"raw_page_{page_num}.txt"), "w", encoding="utf-8") as f:
            f.write(page_text)
            
        print(f"Processed Page {page_num} (V3 Hybrid)")

    doc.close()

if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--pdf", default="/Users/wangchen/Desktop/Essenti..ng.pdf")
    parser.add_argument("--output", default="temp/raw")
    parser.add_argument("--images", default="images")
    parser.add_argument("--start", type=int, required=True)
    parser.add_argument("--end", type=int, required=True)
    args = parser.parse_args()
    extract_pages(args.pdf, args.output, args.images, args.start, args.end)
