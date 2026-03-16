import zipfile
import shutil
import os

file_path = "D:\\rocky-work\\ReportDesign\\恒洁卫浴\\3-4\\docx\\HG-QC-ELC-0005-08-V1.0-检测报告.docx"
output_dir = "D:\\rocky-work\\ReportDesign\\恒洁卫浴\\3-4\\docx\\xml_output"

# 解压 docx（本质是 ZIP）
with zipfile.ZipFile(file_path, 'r') as z:
    z.extractall(output_dir)

print(f"已解压到：{output_dir}")
print("\n包含以下文件：")
for root, dirs, files in os.walk(output_dir):
    for f in files:
        full = os.path.join(root, f)
        print(full)