#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
LabsCare .xmreport 结构样本提取工具 v5

真实结构（header / body / footer 三区域完全一致）：
  <header/body/footer>
    <id> <parent-id> <z-index> <visible> <rotate>
    <data-bind> <border> <background> <height> <shrink>
    <children>
      <grid>
        <id> <parent-id> ... <column-widths> <row-colors> <col-colors>
        <rows>
          <row> × N    ← 数量巨大

精简策略：
  body   → 保留前 body_rows  个 row（默认3，用于学习数据格式）
  header → 保留前 header_rows 个 row（默认1，仅供了解结构）
  footer → 保留前 footer_rows 个 row（默认1，仅供了解结构）
  其余所有骨架节点原样保留
  preview-data 清空
"""

import xml.etree.ElementTree as ET
import copy, os, sys, argparse


def clone(el):
    return copy.deepcopy(el)


def to_xml_str(el):
    return "<?xml version='1.0' encoding='utf-8'?>\n" + ET.tostring(el, encoding="unicode")


def trim_section(section_el, keep_rows, section_label):
    """
    裁剪 header / body / footer 内部的 grid > rows，只保留前 keep_rows 个 row。
    对 children 下所有 grid 都生效。
    返回 (裁剪后的节点副本, 原始row总数)
    """
    new_sec = clone(section_el)
    children = new_sec.find("children")
    if children is None:
        return new_sec, 0

    total = 0
    for grid in list(children):
        rows_el = grid.find("rows")
        if rows_el is None:
            continue
        all_rows = list(rows_el.findall("row"))
        total += len(all_rows)
        kept = min(keep_rows, len(all_rows))
        # 删除多余 row
        for row in all_rows[kept:]:
            rows_el.remove(row)
        # 加说明占位节点
        if len(all_rows) > kept:
            note = ET.SubElement(rows_el, "sample-note")
            note.set("section", section_label)
            note.set("total-rows", str(len(all_rows)))
            note.set("shown", str(kept))
            note.text = (
                f"[样本说明] {section_label} 实际共 {len(all_rows)} 个row，"
                f"此处仅展示前 {kept} 个供格式参考，其余已省略"
            )
    return new_sec, total


def extract_sample(input_file, output_file=None,
                   body_rows=3, header_rows=1, footer_rows=1):

    in_kb = os.path.getsize(input_file) / 1024
    print(f"[*] 读取: {input_file}  ({in_kb:.0f} KB)")

    try:
        tree = ET.parse(input_file)
    except ET.ParseError as e:
        print(f"[错误] XML 解析失败: {e}")
        sys.exit(1)

    root = tree.getroot()
    pages_node = root.find("pages")
    if pages_node is None:
        print("[错误] 未找到 <pages>")
        sys.exit(1)

    new_root = ET.Element(root.tag, root.attrib)

    for child in root:
        tag = child.tag

        if tag == "preview-data":
            ET.SubElement(new_root, tag).text = ""
            print(f"  <{tag}>  → 已清空")
            continue

        if tag != "pages":
            new_root.append(clone(child))
            print(f"  <{tag}>  → 完整保留")
            continue

        # ── <pages> ────────────────────────────────────────────────────
        new_pages = ET.SubElement(new_root, "pages")
        all_pages = list(pages_node.findall("page"))
        print(f"  <pages>  共 {len(all_pages)} 个 <page>")

        for p_idx, page_el in enumerate(all_pages):
            page_id = page_el.findtext("id") or f"page{p_idx+1}"
            new_page = ET.SubElement(new_pages, "page", page_el.attrib)

            for pchild in page_el:
                ptag = pchild.tag

                if ptag == "header":
                    new_sec, total = trim_section(pchild, header_rows, "header")
                    new_page.append(new_sec)
                    kept = min(header_rows, total)
                    print(f"    [{page_id}] <header>  {total} rows → 保留 {kept} 个")

                elif ptag == "footer":
                    new_sec, total = trim_section(pchild, footer_rows, "footer")
                    new_page.append(new_sec)
                    kept = min(footer_rows, total)
                    print(f"    [{page_id}] <footer>  {total} rows → 保留 {kept} 个")

                elif ptag == "body":
                    new_sec, total = trim_section(pchild, body_rows, "body")
                    new_page.append(new_sec)
                    kept = min(body_rows, total)
                    print(f"    [{page_id}] <body>    {total} rows → 保留 {kept} 个")

                else:
                    # id / labscare-datasource / paper / design-region /
                    # paper-orientation / page-padding / economic-print /
                    # background / data-bind  → 原样保留
                    new_page.append(clone(pchild))

    # ── 输出 ────────────────────────────────────────────────────────────
    if output_file is None:
        base, ext = os.path.splitext(input_file)
        output_file = base + "_sample" + ext

    comment = (
        "<!-- ================================================================\n"
        "  LabsCare .xmreport 结构样本（供 Claude 学习格式 / docx转换用）\n"
        f"  body:   保留前 {body_rows} 个 row\n"
        f"  header: 保留前 {header_rows} 个 row\n"
        f"  footer: 保留前 {footer_rows} 个 row\n"
        "  其余骨架节点完整保留，preview-data 已清空\n"
        "================================================================ -->\n"
    )
    final = comment + to_xml_str(new_root)

    with open(output_file, "w", encoding="utf-8") as f:
        f.write(final)

    out_kb = os.path.getsize(output_file) / 1024
    ratio  = (1 - out_kb / in_kb) * 100

    print()
    print(f"[完成]")
    print(f"  原始文件 : {in_kb:>8.1f} KB")
    print(f"  样本文件 : {out_kb:>8.1f} KB  (缩减了 {ratio:.0f}%)")
    print(f"  输出路径 : {output_file}")

    if out_kb > 150:
        print()
        print(f"  ⚠ 样本仍超过 150 KB，建议重新运行并减小 row 保留数：")
        print(f"    python {os.path.basename(__file__)} \"{input_file}\" "
              f"--body-rows 1 --header-rows 0 --footer-rows 0")

    print()
    print("── 发给 Claude 的话术 ─────────────────────────────────────────")
    print("  '这是一个 LabsCare .xmreport 模板的结构样本。")
    print("   header / body / footer 均为 children > grid > rows > row 结构，")
    print("   每区域已保留少量 row 展示格式，其余已省略。")
    print("   请学习这个 XML 格式，我接下来会发给你 docx 内容，")
    print("   请将其转换为相同格式的完整 .xmreport 文件。'")

    return output_file, out_kb


if __name__ == "__main__":
    parser = argparse.ArgumentParser(
        description="LabsCare xmreport 结构样本提取工具 v5",
        formatter_class=argparse.RawTextHelpFormatter,
        epilog=(
            "示例：\n"
            "  # 默认：body保留3行，header/footer各保留1行\n"
            "  python extract_sample.py report.xmreport\n\n"
            "  # 样本仍过大时，全部只保留1行\n"
            "  python extract_sample.py report.xmreport --body-rows 1 --header-rows 1 --footer-rows 1\n\n"
            "  # 极限压缩：header/footer不保留任何row\n"
            "  python extract_sample.py report.xmreport --body-rows 2 --header-rows 0 --footer-rows 0\n"
        )
    )
    parser.add_argument("input", help=".xmreport 文件路径")
    parser.add_argument("--body-rows",   type=int, default=3, help="body 保留的 row 数（默认3）")
    parser.add_argument("--header-rows", type=int, default=1, help="header 保留的 row 数（默认1）")
    parser.add_argument("--footer-rows", type=int, default=1, help="footer 保留的 row 数（默认1）")
    parser.add_argument("-o", "--output", default=None, help="输出文件路径")
    args = parser.parse_args()

    if not os.path.isfile(args.input):
        print(f"[错误] 文件不存在: {args.input}")
        sys.exit(1)

    extract_sample(
        args.input, args.output,
        body_rows=args.body_rows,
        header_rows=args.header_rows,
        footer_rows=args.footer_rows,
    )