
import warnings
warnings.filterwarnings('ignore')

from docx import Document
from openai import OpenAI
from rapidfuzz import fuzz
from zipfile import ZipFile
from lxml import etree
import json

client = OpenAI(
    base_url="https://dashscope.aliyuncs.com/compatible-mode/v1",
    api_key="sk-fcbe3d702dc949eb9e1061206c13ff6c"
)

NAMESPACE = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}


# -----------------------------
# 1. 提取批注
# -----------------------------
def extract_comments(docx_path):
    comments = {}

    with ZipFile(docx_path) as z:
        xml = z.read("word/comments.xml")
        tree = etree.fromstring(xml)

        for c in tree.findall(".//w:comment", NAMESPACE):
            cid = c.get("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}id")
            text = "".join(c.itertext()).strip()
            comments[cid] = text

    return comments


# -----------------------------
# 2. 提取批注对应文本
# -----------------------------
def extract_comment_targets(docx_path):
    with ZipFile(docx_path) as z:
        xml = z.read("word/document.xml")
        tree = etree.fromstring(xml)

    results = []

    for start in tree.findall(".//w:commentRangeStart", NAMESPACE):
        cid = start.get("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}id")

        texts = []
        node = start.getnext()

        while node is not None:
            if node.tag.endswith("commentRangeEnd"):
                break
            texts.extend(node.itertext())
            node = node.getnext()

        results.append({
            "comment_id": cid,
            "target_text": "".join(texts).strip()
        })

    return results


# -----------------------------
# 3. 提取表格结构
# -----------------------------
def parse_with_python_docx(file_path):
    doc = Document(file_path)
    tables = []

    for t_idx, table in enumerate(doc.tables):
        table_data = []

        for row_idx, row in enumerate(table.rows):
            for col_idx, cell in enumerate(row.cells):
                table_data.append({
                    "text": cell.text.strip(),
                    "row": row_idx,
                    "col": col_idx,
                    "bbox": None
                })

        tables.append({
            "table_id": t_idx,
            "cells": table_data
        })

    return tables


# -----------------------------
# 4. fuzz 映射
# -----------------------------
def map_comments_to_cells_advanced(comments, targets, tables):
    mapped = []

    for t in targets:
        target_text = t["target_text"]

        best_match = None
        best_score = 0

        for table in tables:
            for cell in table["cells"]:
                score = fuzz.partial_ratio(target_text, cell["text"])

                if score > best_score:
                    best_score = score
                    best_match = cell

        mapped.append({
            "comment": comments.get(t["comment_id"], ""),
            "target_text": target_text,
            "cell": best_match,
            "score": best_score
        })

    return mapped


# -----------------------------
# 5. 构建 grid（关键）
# -----------------------------
def build_grid(table):
    max_row = max(cell["row"] for cell in table["cells"])
    max_col = max(cell["col"] for cell in table["cells"])

    grid = [["" for _ in range(max_col + 1)] for _ in range(max_row + 1)]

    for cell in table["cells"]:
        grid[cell["row"]][cell["col"]] = cell["text"]

    return grid


# -----------------------------
# 6. Anchor 绑定
# -----------------------------
def attach_mapping_to_grid(mappings):
    anchor_map = {}

    for m in mappings:
        if not m["cell"]:
            continue

        if m["score"] < 60:
            continue  # 过滤低质量匹配

        r = m["cell"]["row"]
        c = m["cell"]["col"]

        key = f"{r}_{c}"

        anchor_map[key] = {
            "comment": m["comment"],
            "target_text": m["target_text"],
            "score": m["score"]
        }

    return anchor_map


# -----------------------------
# 7. 邻近关系（关键增强）
# -----------------------------
def build_neighbors(grid):
    neighbors = {}

    rows = len(grid)
    cols = len(grid[0])

    for r in range(rows):
        for c in range(cols):
            key = f"{r}_{c}"

            neighbors[key] = {
                "self": grid[r][c],
                "right": grid[r][c+1] if c+1 < cols else None,
                "left": grid[r][c-1] if c-1 >= 0 else None,
                "down": grid[r+1][c] if r+1 < rows else None,
                "up": grid[r-1][c] if r-1 >= 0 else None,
            }

    return neighbors


# -----------------------------
# 8. 构建 LLM 输入
# -----------------------------
def build_llm_input(tables, mappings):
    structured_tables = []

    anchor_map = attach_mapping_to_grid(mappings)

    for table in tables:
        grid = build_grid(table)
        neighbors = build_neighbors(grid)

        structured_tables.append({
            "table_id": table["table_id"],
            "grid": grid,
            "anchors": anchor_map,
            "neighbors": neighbors
        })

    return structured_tables


# -----------------------------
# 9. 调用 Qwen（结构推理）
# -----------------------------
def call_qwen(structured_data):

    prompt = f"""
你是一个专业的文档表格结构解析助手。

输入说明：
- grid：二维表格
- anchors：单元格与字段的映射
- neighbors：单元格邻接关系

你的任务：

1. 识别表格结构（kv / 子表）
2. 识别列头
3. 将字段(comment)映射到正确列
4. 处理空单元格 + 邻近推理

输出JSON：

{{
  "tables": [
    {{
      "type": "kv",
      "data": {{}}
    }},
    {{
      "type": "table",
      "columns": [],
      "rows": []
    }}
  ]
}}

输入：
{json.dumps(structured_data, ensure_ascii=False, indent=2)}
"""

    response = client.chat.completions.create(
        model="qwen2.5-14b-instruct",  # ✅ 比 VL 更适合
        messages=[{"role": "user", "content": prompt}],
        temperature=0
    )

    return response.choices[0].message.content


# -----------------------------
# 主流程
# -----------------------------
def run(docx_path):
    comments = extract_comments(docx_path)
    targets = extract_comment_targets(docx_path)

    tables = parse_with_python_docx(docx_path)

    mappings = map_comments_to_cells_advanced(comments, targets, tables)

    structured_tables = build_llm_input(tables, mappings)

    structured_input = {
        "tables": structured_tables
    }

    qwen_output = call_qwen(structured_input)

    return {
        "raw_structure": structured_input,
        "semantic": qwen_output
    }


# -----------------------------
# 执行
# -----------------------------
if __name__ == "__main__":
    try:
        result = run(r"labscareXML\example.docx")

        print("SUCCESS!")
        print("Tables:", len(result["raw_structure"]["tables"]))
        print("\n====== LLM RESULT ======\n")
        print(result["semantic"])

    except Exception as e:
        import traceback
        print("ERROR!")
        print(e)
        traceback.print_exc()