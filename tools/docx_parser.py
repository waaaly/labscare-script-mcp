import warnings
warnings.filterwarnings('ignore')

from docx2python import docx2python
from docling.document_converter import DocumentConverter
from openai import OpenAI
from rapidfuzz import fuzz
import pandas as pd
import json
import os
import sys
from dotenv import load_dotenv
import zipfile
from lxml import etree
from collections import defaultdict
import tempfile
import shutil
import subprocess
# 加载 .env 文件内容
load_dotenv()

client = OpenAI(
    base_url="https://dashscope.aliyuncs.com/compatible-mode/v1",
    api_key=os.getenv("qwenKey")
)

def _get_subprocess_kwargs():
    """Windows 下完全隐藏所有子进程窗口"""
    kwargs = {}
    if sys.platform == 'win32':
        si = subprocess.STARTUPINFO()
        si.dwFlags |= subprocess.STARTF_USESHOWWINDOW
        si.wShowWindow = subprocess.SW_HIDE  # 隐藏窗口，而非 CREATE_NO_WINDOW
        kwargs['startupinfo'] = si
        kwargs['creationflags'] = (
            subprocess.CREATE_NO_WINDOW |
            subprocess.DETACHED_PROCESS  # 与父进程分离，防止继承控制台
        )
    return kwargs

def convert_doc_to_docx(doc_path):
    """
    将 .doc 文件转换为 .docx 格式
    优先使用 LibreOffice，失败则使用 Microsoft Word
    返回转换后的 .docx 文件路径（临时文件），如果转换失败返回 None
    """
    if not doc_path.lower().endswith('.doc'):
        return doc_path
    
    if not os.path.exists(doc_path):
        print(f"❌ 文件不存在: {doc_path}")
        sys.stdout.flush()
        return None
    
    # 创建临时目录
    temp_dir = tempfile.mkdtemp()
    docx_path = os.path.join(temp_dir, os.path.splitext(os.path.basename(doc_path))[0] + '.docx')
    
    # LibreOffice 路径列表
    soffice_paths = [
        'soffice',
        '/usr/bin/soffice',
        '/usr/lib/libreoffice/program/soffice',
        'C:\\Program Files\\LibreOffice\\program\\soffice.exe',
        'C:\\Program Files (x86)\\LibreOffice\\program\\soffice.exe',
    ]
    
    # 尝试使用 LibreOffice 转换
    soffice_cmd = None
    for path in soffice_paths:
        try:
            subprocess.run(
                [path, '--version'],
                capture_output=True,
                stdin=subprocess.DEVNULL,
                timeout=5,
                **_get_subprocess_kwargs()
            )
            soffice_cmd = path
            break
        except (FileNotFoundError, subprocess.TimeoutExpired):
            continue
    
    if soffice_cmd:
        print(f"📄 检测到 .doc 格式文件，使用 LibreOffice 转换为 .docx...")
        sys.stdout.flush()
        try:
            result = subprocess.run(
                [
                    soffice_cmd,
                    '--headless',
                    '--norestore',
                    '--nofirststartwizard',
                    '--nolockcheck',
                    '--convert-to', 'docx',
                    '--outdir', temp_dir,
                    doc_path
                ],
                capture_output=True,
                text=True,
                timeout=120,
                stdin=subprocess.DEVNULL,
                **_get_subprocess_kwargs()
            )
            if result.returncode == 0 and os.path.exists(docx_path) and zipfile.is_zipfile(docx_path):
                print(f"✅ LibreOffice 转换成功！临时文件: {docx_path}")
                sys.stdout.flush()
                return docx_path
            else:
                print(f"⚠️ LibreOffice 转换失败: {result.stderr}")
                sys.stdout.flush()
        except subprocess.TimeoutExpired:
            print("⚠️ LibreOffice 转换超时")
            sys.stdout.flush()
        except Exception as e:
            print(f"⚠️ LibreOffice 转换出错: {e}")
            sys.stdout.flush()
    
    # 尝试使用 Microsoft Word 转换
    try:
        import win32com.client
        print(f"📄 使用 Microsoft Word 转换为 .docx...")
        sys.stdout.flush()

        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False

        try:
            doc = word.Documents.Open(os.path.abspath(doc_path))
            wdFormatXMLDocument = 12
            doc.SaveAs2(os.path.abspath(docx_path), FileFormat=wdFormatXMLDocument)
            doc.Close(False)
            if os.path.exists(docx_path) and zipfile.is_zipfile(docx_path):
                print(f"✅ Microsoft Word 转换成功！临时文件: {docx_path}")
                sys.stdout.flush()
                return docx_path
            else:
                print(f"❌ Microsoft Word 输出不是有效的 docx 文件")
                sys.stdout.flush()
        finally:
            word.Quit()
    except ImportError:
        print("⚠️ 未安装 pywin32，无法使用 Microsoft Word 转换")
        sys.stdout.flush()
    except Exception as e:
        print(f"⚠️ Microsoft Word 转换出错: {e}")
        sys.stdout.flush()
    
    # 清理临时目录
    try:
        shutil.rmtree(temp_dir)
    except:
        pass
    
    print("❌ doc 转 docx 失败，将尝试直接处理原文件（不支持批注提取）")
    sys.stdout.flush()
    return None

def extract_comments_manual(docx_path):
    """
    手动从 word/comments.xml 和 document.xml 中提取批注 + 被批注原文
    返回格式与 docx2python.comments 类似
    """
    comments_dict = {}          # id -> comment_text, author, date
    comment_targets = {}        # id -> target_text (被批注的原文)

    with zipfile.ZipFile(docx_path) as z:
        # 1. 读取所有批注内容
        if 'word/comments.xml' in z.namelist():
            comments_xml = z.read('word/comments.xml')
            root = etree.fromstring(comments_xml)
            ns = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
            
            for c in root.xpath('//w:comment', namespaces=ns):
                c_id = c.get(f"{{{ns['w']}}}id")
                author = c.get(f"{{{ns['w']}}}author", "")
                date = c.get(f"{{{ns['w']}}}date", "")
                text = " ".join(t.text for t in c.xpath('.//w:t', namespaces=ns) if t.text)
                
                comments_dict[c_id] = {
                    "author": author,
                    "date": date,
                    "comment": text.strip()
                }

        # 2. 读取 document.xml，找到批注引用的原文（commentRangeStart + commentReference）
        if 'word/document.xml' in z.namelist():
            doc_xml = z.read('word/document.xml')
            root = etree.fromstring(doc_xml)
            ns = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
            
            # 遍历所有段落，收集批注范围内的文本
            for paragraph in root.xpath('//w:p', namespaces=ns):
                current_text = []
                current_comment_id = None
                
                for elem in paragraph.iter():
                    if elem.tag == f"{{{ns['w']}}}commentRangeStart":
                        current_comment_id = elem.get(f"{{{ns['w']}}}id")
                        current_text = []
                    elif elem.tag == f"{{{ns['w']}}}t" and elem.text:
                        if current_comment_id is not None:
                            current_text.append(elem.text)
                    elif elem.tag == f"{{{ns['w']}}}commentRangeEnd":
                        if current_comment_id and current_comment_id in comments_dict:
                            target = " ".join(current_text).strip()
                            if target:
                                comment_targets[current_comment_id] = target
                        current_comment_id = None
                    elif elem.tag == f"{{{ns['w']}}}commentReference":
                        # 有些文档只有 commentReference，没有 range
                        c_id = elem.get(f"{{{ns['w']}}}id")
                        if c_id and c_id not in comment_targets:
                            comment_targets[c_id] = ""  # 至少记录有批注

    # 组合输出（兼容你之前的代码）
    comments_list = []
    for cid, info in comments_dict.items():
        ref_text = comment_targets.get(cid, "")
        comments_list.append((ref_text, info["author"], info["date"], info["comment"]))

    print(f"手动提取到 {len(comments_list)} 条批注")
    return comments_list
# -----------------------------
# 1. 混合提取：批注 + 表格结构 + 列头
# -----------------------------
def hybrid_extract(docx_path):
    # Step 1: docx2python 获取批注
    comments_list = extract_comments_manual(docx_path)

    # Step 2: Docling AI 提取表格结构（自动处理合并单元格 + 列头）
    converter = DocumentConverter()
    result = converter.convert(docx_path)
    docling_tables = result.document.tables  # Docling 自动识别的表格列表

    tables = []
    cell_comment_map = {}  # 用于后续 anchor

    for t_idx, table in enumerate(docling_tables):
        # Docling 自动导出 DataFrame（列头已智能识别）
        df: pd.DataFrame = table.export_to_dataframe(doc=result.document)
        
        # 转为 grid（第一行 = 自动识别的列头）
        headers = list(df.columns)
        data = df.values.tolist()
        grid = [headers] + data

        # 构建 cells（兼容原代码的 mapping 逻辑）
        cells = []
        for r_idx, row in enumerate(grid):
            for c_idx, text in enumerate(row):
                cell_text = str(text).strip()
                cells.append({
                    "text": cell_text,
                    "row": r_idx,
                    "col": c_idx,
                    "table_id": t_idx
                })

        tables.append({
            "table_id": t_idx,
            "cells": cells,
            "grid": grid,          # 已包含正确列头
            "header_rows": 1       # Docling 通常把第一行当作表头
        })

        # Step 3: 批注映射到具体单元格（fuzzy）
        for ref_text, author, date, cmt_text in comments_list:
            if not ref_text or not cmt_text:
                continue
            best_score = 0
            best_cell = None
            for cell in cells:
                score = fuzz.partial_ratio(ref_text.lower(), cell["text"].lower())
                if score > best_score:
                    best_score = score
                    best_cell = cell
            if best_cell and best_score >= 65:  # 可调阈值
                key = f"{best_cell['table_id']}_{best_cell['row']}_{best_cell['col']}"
                cell_comment_map.setdefault(key, []).append({
                    "comment": cmt_text,
                    "author": author,
                    "date": date,
                    "target_text": ref_text,
                    "score": best_score
                })

    return tables, cell_comment_map


# -----------------------------
# 4. 构建 grid（保留原逻辑，但现在 grid 已正确）
# -----------------------------
def build_grid(table):
    return table["grid"]  # 直接使用 Docling 生成的（已含正确列头）


# -----------------------------
# 5. Anchor 绑定（保留原逻辑）
# -----------------------------
def attach_mapping_to_grid(cell_comment_map, tables):
    anchor_map = {}
    for table in tables:
        for cell in table["cells"]:
            key = f"{cell['table_id']}_{cell['row']}_{cell['col']}"
            if key in cell_comment_map:
                anchor_map[key] = cell_comment_map[key]  # 支持多个批注
    return anchor_map


# -----------------------------
# 6. 邻近关系（保留原逻辑）
# -----------------------------
def build_neighbors(grid):
    neighbors = {}
    rows = len(grid)
    cols = len(grid[0]) if grid else 0

    for r in range(rows):
        for c in range(cols):
            key = f"0_{r}_{c}"  # table_id 固定为 0（单表简化，可扩展）
            neighbors[key] = {
                "self": grid[r][c],
                "right": grid[r][c+1] if c+1 < cols else None,
                "left": grid[r][c-1] if c-1 >= 0 else None,
                "down": grid[r+1][c] if r+1 < rows else None,
                "up": grid[r-1][c] if r-1 >= 0 else None,
            }
    return neighbors


# -----------------------------
# 7. 构建 LLM 输入（保留 + 增强）
# -----------------------------
def build_llm_input(tables, cell_comment_map):
    structured_tables = []
    anchor_map = attach_mapping_to_grid(cell_comment_map, tables)
    for table in tables:
        grid = build_grid(table)
        neighbors = build_neighbors(grid)

        structured_tables.append({
            # "table_id": table["table_id"],
            # "grid": grid,
            "anchors": anchor_map,
            "neighbors": neighbors,
        })

    return structured_tables


# -----------------------------
# 8. 调用 Qwen（结构推理，prompt 增强）
# @table_desc: 用户输入的子表格描述，包含什么列名
# -----------------------------
def call_qwen(structured_data, table_desc):
    sys_prompt = f"""
你是一个专业的文档表格结构解析助手。
输入数据：
{json.dumps(structured_data, ensure_ascii=False, indent=2)}
输入说明：
- anchors：单元格与批注的映射
- neighbors：单元格邻接关系

严格按照以下输出格式：
{{
    "tables": [
    //tables数组长度要严格要求与用户说明子表格个数相等
        {{
            "columns": [
                {{
                    "label": "取自单元格文本"
                    "field": "此字段内容严格中单元格对应的批注内的提取字段名称（英文），禁止自己创造生成",
                }},
            ],
        }}
    ],
    "fileds":[
    // fileds数组要从anchors中只保留不属于子表格的字段
        {{
            "label": "取自单元格文本"
            "field": "此字段内容严格中单元格对应的批注内的提取字段名称（英文），禁止自己创造生成",
        }},
    ]
}}
你的任务：
1. 处理合并单元格 + 空单元格
2. 将批注内容映射到正确字段
3. 输出结构化 JSON
4. 理解用户输入prompt中子表格描述，其中用户会提到有几个子表格，列头分别是什么你要在返回结构中体现出来

"""
    print("\n📤 正在调用 LLM 进行表格结构解析...")
    print("⏳ 请稍候，这可能需要几秒钟时间...\n")
    sys.stdout.flush()
    
    response = client.chat.completions.create(
        model="qwen2.5-14b-instruct",
        messages=[
            {"role": "system", "content": sys_prompt},
            {"role": "user", "content": table_desc},
        ],
        temperature=0.0,          # 改成 0.0（更激进的 greedy）
        top_p=1.0,                # 必须加上
        seed=42,                  # 固定随机种子（DashScope 支持）
        max_tokens=4096,          # 根据你的输出长度调整，避免截断
        response_format={"type": "json_object"}   # 强制 JSON 输出（Qwen 支持）
    )
    
    print("✅ LLM 调用完成！\n")
    sys.stdout.flush()
    
    return response.choices[0].message.content

#暴露给外部调用的函数
def handle_parse_docx(docx_path, table_desc):
    # 检查并转换 .doc 到 .docx
    actual_docx_path = convert_doc_to_docx(docx_path)
    is_temp_file = actual_docx_path is not None and actual_docx_path != docx_path
    use_docling_only = actual_docx_path is None  # 如果转换失败，使用 Docling 直接处理
    
    if use_docling_only:
        actual_docx_path = docx_path  # 使用原文件
    
    try:
        tables, cell_comment_map = hybrid_extract(actual_docx_path)
        structured_tables = build_llm_input(tables, cell_comment_map)
        qwen_output = call_qwen(structured_tables[0], table_desc)
        return qwen_output
    finally:
        # 如果创建了临时文件，清理它
        if is_temp_file and os.path.exists(actual_docx_path):
            temp_dir = os.path.dirname(actual_docx_path)
            try:
                shutil.rmtree(temp_dir)
                print(f"🧹 已清理临时文件")
                sys.stdout.flush()
            except Exception as e:
                print(f"⚠️  清理临时文件失败: {e}")
                sys.stdout.flush()
# -----------------------------
# 主流程
# -----------------------------
def run(docx_path, table_desc):
    print("\n" + "="*60)
    print("🚀 开始执行文档表格解析流程")
    print("="*60 + "\n")
    sys.stdout.flush()
    
    # 检查并转换 .doc 到 .docx
    actual_docx_path = convert_doc_to_docx(docx_path)
    is_temp_file = actual_docx_path is not None and actual_docx_path != docx_path
    use_docling_only = actual_docx_path is None  # 如果转换失败，使用 Docling 直接处理
    
    if use_docling_only:
        actual_docx_path = docx_path  # 使用原文件
        print("📄 将使用 Docling 直接处理文件（批注提取功能不可用）")
        sys.stdout.flush()
    
    try:
        print("📋 步骤 1/3: 混合提取表格和批注...")
        sys.stdout.flush()
        tables, cell_comment_map = hybrid_extract(actual_docx_path)
        print(f"✅ 提取完成！")
        print(f"   - 表格数量: {len(tables)}")
        print(f"   - 批注映射数: {len(cell_comment_map)}")

        sys.stdout.flush()
        
        print("\n📋 步骤 2/3: 构建 LLM 输入结构...")
        sys.stdout.flush()
        structured_tables = build_llm_input(tables, cell_comment_map)
        print(f"✅ 结构构建完成！")
        print(f"   - 结构化表格数: {len(structured_tables)}")
        sys.stdout.flush()

        print(f"\n📋 步骤 3/3: 调用 LLM 进行语义分析...")
        sys.stdout.flush()
        qwen_output = call_qwen(structured_tables[0], table_desc)
        print(f"✅ LLM 分析完成！")
        sys.stdout.flush()

        print("\n" + "="*60)
        print("🎉 所有步骤执行完成！")
        print("="*60 + "\n")
        sys.stdout.flush()

        return {
            "raw_structure": structured_tables,
            "semantic": qwen_output
        }
    finally:
        # 如果创建了临时文件，清理它
        if is_temp_file and os.path.exists(actual_docx_path):
            temp_dir = os.path.dirname(actual_docx_path)
            try:
                shutil.rmtree(temp_dir)
                print(f"🧹 已清理临时文件")
                sys.stdout.flush()
            except Exception as e:
                print(f"⚠️  清理临时文件失败: {e}")
                sys.stdout.flush()


# -----------------------------
# 执行
# -----------------------------
if __name__ == "__main__":
    try:
        result = run(r".\labscareXML\example.doc","包含有一个子表格，该表格的列头分别是：测试项目，测试要求，样品编号，测试结果，结果判定。")

        print("✅ SUCCESS! 混合提取完成")
        print(f"Tables: {len(result['semantic']['tables'])}")
        print(f"Fileds: {len(result['semantic']['fields'])}")
        print("\n====== LLM 语义化结果 ======\n")
        print(result["semantic"])

    except Exception as e:
        import traceback
        print("❌ ERROR!")
        print(e)
        traceback.print_exc()