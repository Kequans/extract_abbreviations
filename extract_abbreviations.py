#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Kequan
缩略词提取工具
从Word文档中提取格式为"中文（英文全称, 英文缩写）"的缩略词
"""

import re
from docx import Document
import sys


def extract_abbreviations(docx_path):
    """
    从Word文档中提取缩略词

    Args:
        docx_path: Word文档路径

    Returns:
        list: 包含(中文, 英文全称, 英文缩写)的元组列表
    """
    try:
        doc = Document(docx_path)
    except Exception as e:
        print(f"错误：无法打开文档 {docx_path}")
        print(f"详细信息：{e}")
        return []

    # 多个正则表达式匹配模式，覆盖不同格式
    patterns = [
        # 模式1: 中文（英文全称, 英文缩写）
        r'([\u4e00-\u9fa5]+)\s*[（(]\s*([A-Za-z\s\-]+)\s*[,，]\s*([A-Z][A-Za-z0-9]*)\s*[)）]',
        # 模式2: 英文+中文（英文全称, 英文缩写） - 如"旋转IoU（Rotated IoU, RIoU）"
        r'(?:[\u4e00-\u9fa5]*[A-Za-z]+[\u4e00-\u9fa5]*)\s*[（(]\s*([A-Za-z\s\-]+)\s*[,，]\s*([A-Z][A-Za-z0-9]*)\s*[)）]',
        # 模式3: 纯中文后跟括号 - 更宽松的匹配
        r'([\u4e00-\u9fa5]{2,})\s*[（(]\s*([A-Za-z][\w\s\-]*[A-Za-z])\s*[,，]\s*([A-Z][A-Za-z0-9]+)\s*[)）]',
    ]

    abbreviations = []
    seen = set()  # 用于去重

    def process_text(text):
        """处理文本，提取缩略词"""
        for pattern in patterns:
            matches = re.finditer(pattern, text)
            for match in matches:
                groups = match.groups()

                # 根据匹配的组数确定提取方式
                if len(groups) == 3:
                    chinese = groups[0].strip() if groups[0] else ""
                    english_full = groups[1].strip()
                    english_abbr = groups[2].strip()
                elif len(groups) == 2:
                    # 模式2的情况，需要从原文提取中文部分
                    english_full = groups[0].strip()
                    english_abbr = groups[1].strip()
                    # 提取括号前的中文部分，扩大搜索范围并改进匹配
                    start = match.start()
                    prefix = text[max(0, start-50):start]
                    # 匹配更完整的中文短语（包含标点符号前的内容）
                    chinese_match = re.search(r'([^\n\r。；，、！？]+[\u4e00-\u9fa5A-Za-z0-9]+)\s*$', prefix)
                    if chinese_match:
                        chinese = chinese_match.group(1).strip()
                        # 清理开头的标点和连接词
                        chinese = re.sub(r'^[的、，。；：\s]+', '', chinese)
                    else:
                        chinese = ""
                else:
                    continue

                # 清理和验证
                if not english_abbr or len(english_abbr) < 2:
                    continue

                # 去重（基于缩写和英文全称，允许同一缩写有不同中文描述）
                key = (english_abbr, english_full)
                if key not in seen:
                    seen.add(key)
                    abbreviations.append((english_abbr, english_full, chinese))

    # 遍历所有段落
    for para in doc.paragraphs:
        process_text(para.text)

    # 同时检查表格中的内容
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                process_text(cell.text)

    return abbreviations


def save_to_word(abbreviations, template_path, output_path):
    """
    将缩略词保存到Word文档

    Args:
        abbreviations: 缩略词列表
        template_path: 模板文档路径
        output_path: 输出文档路径
    """
    # 加载模板
    doc = Document(template_path)

    # 获取表格（假设模板中只有一个表格）
    table = doc.tables[0]

    # 按英文缩写字母顺序排序
    abbreviations_sorted = sorted(abbreviations, key=lambda x: x[0].upper())

    # 添加数据行
    for abbr, english_full, chinese in abbreviations_sorted:
        row = table.add_row()
        row.cells[0].text = abbr
        row.cells[1].text = english_full
        row.cells[2].text = chinese

    # 保存文档
    doc.save(output_path)


def main():
    if len(sys.argv) < 2:
        print("用法: python extract_abbreviations.py <word文档路径> [模板路径]")
        print("示例: python extract_abbreviations.py 论文.docx")
        print("      python extract_abbreviations.py 论文.docx 缩略词格式.docx")
        sys.exit(1)

    docx_path = sys.argv[1]
    template_path = sys.argv[2] if len(sys.argv) > 2 else "缩略词格式.docx"
    output_path = "缩略词列表.docx"

    print(f"正在提取 {docx_path} 中的缩略词...\n")

    abbreviations = extract_abbreviations(docx_path)

    if not abbreviations:
        print("未找到符合格式的缩略词")
        return

    print(f"共找到 {len(abbreviations)} 个缩略词：\n")
    print("-" * 80)
    print(f"{'序号':<6}{'英文缩写':<15}{'英文全称':<40}{'中文全称':<20}")
    print("-" * 80)

    # 按英文缩写排序显示
    abbreviations_sorted = sorted(abbreviations, key=lambda x: x[0].upper())
    for idx, (abbr, english_full, chinese) in enumerate(abbreviations_sorted, 1):
        print(f"{idx:<6}{abbr:<15}{english_full:<40}{chinese:<20}")

    print("-" * 80)

    # 保存到Word文档
    try:
        save_to_word(abbreviations, template_path, output_path)
        print(f"\n结果已保存到：{output_path}")
    except Exception as e:
        print(f"\n保存Word文档时出错：{e}")
        print("尝试保存为文本文件...")

        # 备用：保存为文本文件
        txt_output = "缩略词列表.txt"
        with open(txt_output, 'w', encoding='utf-8') as f:
            f.write(f"缩略词提取结果（共 {len(abbreviations)} 个）\n")
            f.write("=" * 80 + "\n\n")
            f.write(f"{'英文缩写':<15}{'英文全称':<40}{'中文全称':<20}\n")
            f.write("-" * 80 + "\n")
            for abbr, english_full, chinese in abbreviations_sorted:
                f.write(f"{abbr:<15}{english_full:<40}{chinese:<20}\n")
        print(f"已保存为文本文件：{txt_output}")


if __name__ == "__main__":
    main()
