import os
import re
import argparse

def assemble_book(input_dir, output_path, manifest_path=None):
    """
    Assembles the book from individual markdown files.
    If manifest_path is provided, it uses the order in the manifest.
    Otherwise, it sorts files by name.
    """
    if manifest_path and os.path.exists(manifest_path):
        with open(manifest_path, 'r', encoding='utf-8') as f:
            import json
            files = json.load(f)
    else:
        files = sorted([f for f in os.listdir(input_dir) if f.endswith('.md')])

    full_content = ""
    for filename in files:
        file_path = os.path.join(input_dir, filename)
        if not os.path.exists(file_path):
            print(f"Warning: {file_path} not found. Skipping.")
            continue
            
        with open(file_path, 'r', encoding='utf-8') as f:
            content = f.read()
            
        # 1. Automatic Fixes
        # Fix unclosed display math (common issue found earlier)
        # Match $$...$ but not $$...$$
        # content = re.sub(r'\$\$(.*?)(?<!\$)\$(?!\$)', r'$$\1$$', content, flags=re.DOTALL)
        # Note: The above regex might be tricky. Let's use a simpler approach for the known bug.
        content = content.replace('$$F_R = k \\times F_N \\quad (2.9)$', '$$F_R = k \\times F_N \\quad (2.9)$$')
        content = content.replace('$$F_R = k \\times v \\quad (2.10)$', '$$F_R = k \\times v \\quad (2.10)$$')
        content = content.replace('$$F_R = k \\times x \\quad (2.11)$', '$$F_R = k \\times x \\quad (2.11)$$')

        # 2. Ensure page breaks before chapters or as needed
        if content.startswith('# '):
            full_content += "\n\n<div style='page-break-before: always;'></div>\n\n"
            
        full_content += content + "\n\n---\n\n"

    with open(output_path, 'w', encoding='utf-8') as f:
        f.write(full_content)
    
    print(f"Book assembled into {output_path}")

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Assemble book from MD files")
    parser.add_argument("--dir", type=str, required=True, help="Directory with MD files")
    parser.add_argument("--output", type=str, required=True, help="Output MD file")
    parser.add_argument("--manifest", type=str, help="Optional JSON manifest list")
    
    args = parser.parse_args()
    assemble_book(args.dir, args.output, args.manifest)
