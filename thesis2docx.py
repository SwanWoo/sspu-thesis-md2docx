#!/usr/bin/env python3
"""
将 final_paper.md 毕业论文 Markdown 转换为 Word 文档。
基于 md2docx.py 架构，扩展支持：封面注入、表格、图片、公式、代码块等。
模板结构（段落索引 P0-P151，2026届新版模板）：
  P0-P3:   封面页头（格式说明+校徽图片）
  P4-P13:  封面字段（af2样式，P5=英文题目，无ddList下拉框）
  P14:     分节符（封面→声明，continuous）
  P15-P29: 声明页
  P30:     分节符（声明→摘要，nextPage）
  P31:     中文标题
  P32:     "摘要"标签
  P33-P35: 中文摘要正文
  P36:     中文关键词
  P37:     关键词格式说明（需清空）
  P38:     空段落（格式说明，需清空）
  P39:     英文标题（af2样式，TNR+宋体，sz=36）
  P40:     "ABSTRACT"标签（af2样式，TNR+宋体）
  P41:     英文摘要正文（af2样式，TNR+宋体，b=0，sz=24）
  P42:     英文关键词（af2样式，TNR+宋体）
  P43:     分节符（摘要→目录，nextPage）
  P44:     目录标题
  P45-P55: 目录条目（TOC1/TOC2/TOC3样式）
  P65:     分节符（目录→正文，nextPage）
  P66+:    正文内容
"""

import re
import copy
import io
import os
import sys
import zipfile
import urllib.request
import ssl
import tempfile
import shutil
import subprocess
import xml.etree.ElementTree as ET
from latex2mathml import converter

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
M_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/math'
R_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
XML_NS = '{http://www.w3.org/XML/1998/namespace}'
ML_NS = '{http://www.w3.org/1998/Math/MathML}'

UPRIGHT_FUNCTIONS = {
    'sin', 'cos', 'tan', 'cot', 'sec', 'csc', 'arcsin', 'arccos', 'arctan',
    'sinh', 'cosh', 'tanh', 'log', 'ln', 'exp', 'lim', 'min', 'max',
    'sup', 'inf', 'det', 'dim', 'mod', 'gcd', 'deg', 'arg', 'hom',
    'ker', 'Im', 'Re', 'Pr',
}


# ============================================================
# 1.1 中文标点符号规范化
# ============================================================

def _normalize_table_punctuation(text):
    """表格单元格标点规范化：使用与正文相同的中文标点规则。
    冒号、逗号等在中文语境中统一使用全角标点。"""
    if not text:
        return text
    return _normalize_chinese_punctuation(text)


def _normalize_chinese_punctuation(text):
    """将中文语境中的英文标点转换为中文标点。
    规则：
      - 英文双引号 "" → 中文双引号""（成对匹配）
      - 中文之间的英文逗号 , → 中文逗号 ，
      - 中文之间的英文冒号 : → 中文冒号 ：
      - 中文之间的英文问号 ? → 中文问号 ？
    排除：代码标记(`...`)内的内容、URL、公式($...$)内的内容保持不变。
    """
    if not text:
        return text

    # 先保护代码标记和公式中的内容
    protected = []
    counter = [0]

    def _protect(m):
        placeholder = f'\x00PROT{counter[0]}\x00'
        protected.append((placeholder, m.group(0)))
        counter[0] += 1
        return placeholder

    # 保护行内代码 `...`
    text = re.sub(r'`[^`]+`', _protect, text)
    # 保护行内公式 $...$
    text = re.sub(r'(?<!\$)\$(?!\$)(.+?)(?<!\$)\$(?!\$)', _protect, text)
    # 保护 URL
    text = re.sub(r'https?://[^\s)）,，]+', _protect, text)
    # 保护 API路径 /api/...
    text = re.sub(r'/api/[^\s,，:：)）]+', _protect, text)

    # 1. 英文双引号 → 中文双引号（成对匹配）
    result = []
    quote_open = False
    for ch in text:
        if ch == '"':
            if not quote_open:
                result.append('“')  # "
                quote_open = True
            else:
                result.append('”')  # "
                quote_open = False
        else:
            result.append(ch)
    text = ''.join(result)

    # 2. 中文之间的英文逗号 → 中文逗号
    # 包括中文与数字之间的逗号（如 [16,17]）
    text = re.sub(r'([一-鿿　-〿＀-￯\d]),(?=[一-鿿　-〿＀-￯\d])',
                  r'\1，', text)
    # 方括号内数字间的逗号（如 [16,17] → [16，17]）
    text = re.sub(r'\[(\d+),(\d+)\]', r'[\1，\2]', text)

    # 3. 中文之间的英文冒号 → 中文冒号（但保护 "Agent：" 等已在Markdown中使用的全角冒号场景）
    # 也匹配行末冒号（如 "属性:" 行末无后续字符）
    text = re.sub(r'([一-鿿　-〿＀-￯]):(?=[一-鿿　-〿＀-￯\s]|$)',
                  r'\1：', text)

    # 3b. 中文语境中英文词汇后的冒号也转为全角（如 "Authorization: Bearer" 在中文段落中）
    if re.search(r'[一-鿿]', text):
        text = re.sub(r':(?=[\s一-鿿]|$)', '：', text)

    # 4. 中文之间的英文问号 → 中文问号
    text = re.sub(r'([一-鿿　-〿＀-￯])\?(?=[一-鿿　-〿＀-￯])',
                  r'\1？', text)

    # 4b. 中文语境中更广泛地问号转换
    if re.search(r'[一-鿿]', text):
        text = re.sub(r'\?(?=[一-鿿　-〿＀-￯])', '？', text)

    # 还原保护的内容
    for placeholder, original in protected:
        text = text.replace(placeholder, original)

    # 5. 还原后二次扫描：中文语境中剩余英文冒号转为全角（检测工具要求）
    # 保护：URL协议分隔符 :// 、端口号 :\d 、代码标记内内容
    if re.search(r'[一-鿿]', text):
        text = re.sub(r':(?!//|\d)', '：', text)

    return text


# ============================================================
# 1. 命名空间注册
# ============================================================

def register_namespaces():
    """注册所有需要的命名空间。注意：同一 URI 只保留最后一个注册的前缀。"""
    namespaces = {
        'wpc': 'http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas',
        'mc': 'http://schemas.openxmlformats.org/markup-compatibility/2006',
        'o': 'urn:schemas-microsoft-com:office:office',
        'm': 'http://schemas.openxmlformats.org/officeDocument/2006/math',
        'v': 'urn:schemas-microsoft-com:vml',
        'wp14': 'http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing',
        'wp': 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
        'w10': 'urn:schemas-microsoft-com:office:word',
        'w14': 'http://schemas.microsoft.com/office/word/2010/wordml',
        'w15': 'http://schemas.microsoft.com/office/word/2012/wordml',
        'wpg': 'http://schemas.microsoft.com/office/word/2010/wordprocessingGroup',
        'wpi': 'http://schemas.microsoft.com/office/word/2010/wordprocessingInk',
        'wne': 'http://schemas.microsoft.com/office/word/2006/wordml',
        'wps': 'http://schemas.microsoft.com/office/word/2010/wordprocessingShape',
        'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
        'pic': 'http://schemas.openxmlformats.org/drawingml/2006/picture',
    }
    for prefix, uri in namespaces.items():
        ET.register_namespace(prefix, uri)
    # r: 前缀用于关系引用（图片 embed 等），必须在最后注册
    ET.register_namespace('r', R_NS)
    # w: 作为主命名空间前缀（而非默认命名空间，避免与 w:document 根元素冲突）
    ET.register_namespace('w', W)


# ============================================================
# 1.5 LaTeX → OMML 转换（行内公式）
# ============================================================

def _mathml_children(elem):
    """获取 MathML 元素的直接子元素列表（跳过注释等非元素节点）。"""
    return [c for c in elem if isinstance(c.tag, str)]


def _convert_mathml_node(src, dest_parent):
    """递归地将 MathML 元素转换为 OMML 元素，追加到 dest_parent。"""
    tag = src.tag
    local = tag.replace(ML_NS, '') if tag.startswith('{') else tag

    if local == 'math' or local == 'mrow':
        for child in _mathml_children(src):
            _convert_mathml_node(child, dest_parent)
        return

    if local in ('mi', 'mn', 'mo', 'mtext'):
        r = ET.SubElement(dest_parent, f'{{{M_NS}}}r')
        text = (src.text or '').strip() or src.text or ''
        t = ET.SubElement(r, f'{{{M_NS}}}t')
        t.text = text
        t.set(f'{XML_NS}space', 'preserve')
        # 函数名 upright 样式
        if local == 'mi' and text in UPRIGHT_FUNCTIONS:
            rpr = ET.SubElement(r, f'{{{M_NS}}}rPr')
            sty = ET.SubElement(rpr, f'{{{M_NS}}}sty')
            sty.set(f'{{{M_NS}}}val', 'p')
        return

    if local == 'mfrac':
        f = ET.SubElement(dest_parent, f'{{{M_NS}}}f')
        children = _mathml_children(src)
        if len(children) >= 2:
            num = ET.SubElement(f, f'{{{M_NS}}}num')
            _convert_mathml_node(children[0], num)
            den = ET.SubElement(f, f'{{{M_NS}}}den')
            _convert_mathml_node(children[1], den)
        return

    if local == 'msup':
        s = ET.SubElement(dest_parent, f'{{{M_NS}}}sSup')
        children = _mathml_children(src)
        if len(children) >= 2:
            e = ET.SubElement(s, f'{{{M_NS}}}e')
            _convert_mathml_node(children[0], e)
            sup = ET.SubElement(s, f'{{{M_NS}}}sup')
            _convert_mathml_node(children[1], sup)
        return

    if local == 'msub':
        s = ET.SubElement(dest_parent, f'{{{M_NS}}}sSub')
        children = _mathml_children(src)
        if len(children) >= 2:
            e = ET.SubElement(s, f'{{{M_NS}}}e')
            _convert_mathml_node(children[0], e)
            sub = ET.SubElement(s, f'{{{M_NS}}}sub')
            _convert_mathml_node(children[1], sub)
        return

    if local == 'msubsup':
        s = ET.SubElement(dest_parent, f'{{{M_NS}}}sSubSup')
        children = _mathml_children(src)
        if len(children) >= 3:
            e = ET.SubElement(s, f'{{{M_NS}}}e')
            _convert_mathml_node(children[0], e)
            sub = ET.SubElement(s, f'{{{M_NS}}}sub')
            _convert_mathml_node(children[1], sub)
            sup = ET.SubElement(s, f'{{{M_NS}}}sup')
            _convert_mathml_node(children[2], sup)
        return

    if local == 'msqrt':
        rad = ET.SubElement(dest_parent, f'{{{M_NS}}}rad')
        deg = ET.SubElement(rad, f'{{{M_NS}}}deg')
        e = ET.SubElement(rad, f'{{{M_NS}}}e')
        for child in _mathml_children(src):
            _convert_mathml_node(child, e)
        return

    if local == 'mroot':
        rad = ET.SubElement(dest_parent, f'{{{M_NS}}}rad')
        children = _mathml_children(src)
        # mroot: [base, index] → OMML: m:e=base, m:deg=index
        deg = ET.SubElement(rad, f'{{{M_NS}}}deg')
        e = ET.SubElement(rad, f'{{{M_NS}}}e')
        if len(children) >= 1:
            _convert_mathml_node(children[0], e)
        if len(children) >= 2:
            _convert_mathml_node(children[1], deg)
        return

    if local == 'mover':
        children = _mathml_children(src)
        if len(children) >= 2:
            # 检查是否为重音符号（第二个子元素是单个字符的 mo）
            acc_char = ''
            acc_src = children[1]
            acc_tag = acc_src.tag.replace(ML_NS, '') if acc_src.tag.startswith('{') else acc_src.tag
            if acc_tag == 'mo' and acc_src.text and len(acc_src.text.strip()) <= 2:
                acc_char = acc_src.text.strip()
            if acc_char:
                acc = ET.SubElement(dest_parent, f'{{{M_NS}}}acc')
                accPr = ET.SubElement(acc, f'{{{M_NS}}}accPr')
                chr_el = ET.SubElement(accPr, f'{{{M_NS}}}chr')
                chr_el.set(f'{{{M_NS}}}val', acc_char)
                e = ET.SubElement(acc, f'{{{M_NS}}}e')
                _convert_mathml_node(children[0], e)
            else:
                limU = ET.SubElement(dest_parent, f'{{{M_NS}}}limUpp')
                e = ET.SubElement(limU, f'{{{M_NS}}}e')
                _convert_mathml_node(children[0], e)
                lim = ET.SubElement(limU, f'{{{M_NS}}}lim')
                _convert_mathml_node(children[1], lim)
        return

    if local == 'munder':
        children = _mathml_children(src)
        if len(children) >= 2:
            limL = ET.SubElement(dest_parent, f'{{{M_NS}}}limLow')
            e = ET.SubElement(limL, f'{{{M_NS}}}e')
            _convert_mathml_node(children[0], e)
            lim = ET.SubElement(limL, f'{{{M_NS}}}lim')
            _convert_mathml_node(children[1], lim)
        return

    if local == 'munderover':
        children = _mathml_children(src)
        if len(children) >= 3:
            # 尝试检测 nary (求和/积分/乘积)
            nary_chars = {'∑': '∑', '∏': '∏', '∫': '∫', '⋃': '⋃', '⋂': '⋂'}
            base_text = ''
            base_src = children[0]
            base_tag = base_src.tag.replace(ML_NS, '') if base_src.tag.startswith('{') else base_src.tag
            if base_tag == 'mo' and base_src.text:
                base_text = base_src.text.strip()
            if base_text in nary_chars:
                nary = ET.SubElement(dest_parent, f'{{{M_NS}}}nary')
                naryPr = ET.SubElement(nary, f'{{{M_NS}}}naryPr')
                chr_el = ET.SubElement(naryPr, f'{{{M_NS}}}chr')
                chr_el.set(f'{{{M_NS}}}val', nary_chars[base_text])
                limLoc = ET.SubElement(naryPr, f'{{{M_NS}}}limLoc')
                limLoc.set(f'{{{M_NS}}}val', 'subSup')
                sub = ET.SubElement(nary, f'{{{M_NS}}}sub')
                _convert_mathml_node(children[1], sub)
                sup = ET.SubElement(nary, f'{{{M_NS}}}sup')
                _convert_mathml_node(children[2], sup)
                e = ET.SubElement(nary, f'{{{M_NS}}}e')
            else:
                # 普通上下标
                nary = ET.SubElement(dest_parent, f'{{{M_NS}}}nary')
                naryPr = ET.SubElement(nary, f'{{{M_NS}}}naryPr')
                limLoc = ET.SubElement(naryPr, f'{{{M_NS}}}limLoc')
                limLoc.set(f'{{{M_NS}}}val', 'subSup')
                sub = ET.SubElement(nary, f'{{{M_NS}}}sub')
                _convert_mathml_node(children[1], sub)
                sup = ET.SubElement(nary, f'{{{M_NS}}}sup')
                _convert_mathml_node(children[2], sup)
                e = ET.SubElement(nary, f'{{{M_NS}}}e')
                _convert_mathml_node(children[0], e)
        return

    if local == 'mtable':
        m = ET.SubElement(dest_parent, f'{{{M_NS}}}m')
        for child in _mathml_children(src):
            _convert_mathml_node(child, m)
        return

    if local == 'mtr':
        mr = ET.SubElement(dest_parent, f'{{{M_NS}}}mr')
        for child in _mathml_children(src):
            _convert_mathml_node(child, mr)
        return

    if local == 'mtd':
        e = ET.SubElement(dest_parent, f'{{{M_NS}}}e')
        for child in _mathml_children(src):
            _convert_mathml_node(child, e)
        return

    if local == 'mo' and (src.get('fence') == 'true' or src.text in ('(', ')', '[', ']', '{', '}', '|', '‖')):
        # 括号元素
        _convert_mathml_node.__wrapped__(src, dest_parent) if False else None
        r = ET.SubElement(dest_parent, f'{{{M_NS}}}r')
        t = ET.SubElement(r, f'{{{M_NS}}}t')
        t.text = (src.text or '')
        t.set(f'{XML_NS}space', 'preserve')
        return

    # 未知元素：递归处理子元素
    for child in _mathml_children(src):
        _convert_mathml_node(child, dest_parent)


def latex_to_omml(latex_str):
    """将 LaTeX 行内公式转换为 OMML <m:oMath> 元素。失败返回 None。"""
    try:
        latex_str = latex_str.strip()
        if not latex_str:
            return None
        mathml_str = converter.convert(latex_str)
        mathml_elem = ET.fromstring(mathml_str)
        omath = ET.Element(f'{{{M_NS}}}oMath')
        _convert_mathml_node(mathml_elem, omath)
        return omath
    except Exception as e:
        print(f'  警告: 行内公式转换失败: {e}', file=sys.stderr)
        return None


# ============================================================
# 2. Markdown 解析
# ============================================================

def unescape_markdown(text):
    """去除 Markdown 转义符。"""
    text = text.replace(r'\*\*', '**')
    text = text.replace(r'\*', '*')
    text = text.replace(r'\-\>', '->')
    text = text.replace(r'\-', '-')
    text = text.replace(r'\(', '(')
    text = text.replace(r'\)', ')')
    text = text.replace(r'\#', '#')
    text = text.replace(r'\_', '_')
    text = text.replace(r'\\', '\\')
    return text


def parse_cover_table(md_text):
    """从 Markdown 开头提取封面信息表格。"""
    # 找到 </div> 之前的表格
    match = re.search(r'\|.*?\*\*题目\*\*.*?\|(.*?)\|', md_text, re.DOTALL)
    if not match:
        return {}

    cover = {}
    # 匹配所有 | **key** | value | 行
    lines = md_text.split('\n')
    in_table = False
    for line in lines:
        line = line.strip()
        if line.startswith('|') and '**' in line:
            in_table = True
            # 提取 key 和 value
            cells = [c.strip() for c in line.split('|')]
            cells = [c for c in cells if c]  # 去空
            if len(cells) >= 2:
                key = re.sub(r'\*\*', '', cells[0]).strip()
                value = re.sub(r'\*\*', '', cells[1]).strip()
                cover[key] = value
        elif in_table and not line.startswith('|'):
            break
        # 跳过分隔行
        if re.match(r'^\|[-:|]+\|$', line):
            continue
    return cover


def parse_markdown(md_text):
    """解析 Markdown 文本为结构化块列表。扩展支持表格、代码块、公式、图片等。"""
    lines = md_text.split('\n')
    blocks = []
    i = 0
    in_appendix = False  # 标记是否在附录部分

    while i < len(lines):
        line = lines[i]

        # 跳过空行
        if line.strip() == '':
            i += 1
            continue

        # 跳过 HTML div 和居中标签（页面分隔等）
        if re.match(r'^\s*<div\s', line.strip()):
            i += 1
            continue
        if re.match(r'^\s*</div>', line.strip()):
            i += 1
            continue

        # 水平分隔线 — 直接跳过，不在 docx 中生成
        if re.match(r'^---+\s*$', line.strip()):
            i += 1
            continue

        # 标题
        heading_match = re.match(r'^(#{1,6})\s+(.+)$', line)
        if heading_match:
            level = len(heading_match.group(1))
            text = unescape_markdown(heading_match.group(2).strip())
            blocks.append({'type': 'heading', 'level': level, 'text': text})
            # 检测是否进入附录部分
            if text == '附录':
                in_appendix = True
            i += 1
            continue

        # 代码块
        if re.match(r'^```', line.strip()):
            lang_match = re.match(r'^```(\w*)', line.strip())
            lang = lang_match.group(1) if lang_match else ''
            code_lines = []
            i += 1
            while i < len(lines) and not lines[i].strip().startswith('```'):
                code_lines.append(lines[i])
                i += 1
            i += 1  # 跳过结束的 ```

            if lang == 'mermaid':
                # mermaid 块：检查下一个非空行是否为图注（图X-X ...）
                caption = ''
                saved_i = i
                while i < len(lines) and lines[i].strip() == '':
                    i += 1
                if i < len(lines) and re.match(r'^图\d', lines[i].strip()):
                    caption = lines[i].strip()
                    i += 1
                blocks.append({'type': 'mermaid', 'code': '\n'.join(code_lines), 'caption': caption})
            else:
                # 正文和附录中的代码都转换为表格格式
                # 检查代码块前是否有描述注释
                description = ''
                # 向前查找最近的注释行
                for j in range(len(blocks)-1, -1, -1):
                    if blocks[j]['type'] == 'paragraph':
                        text = blocks[j]['text']
                        # 检查是否是描述注释
                        if text.startswith('<!-- 代码描述:') and text.endswith('-->'):
                            description = text[10:-3].strip()  # 移除<!-- 代码描述:和-->
                            # 移除这个注释段落
                            blocks.pop(j)
                            break

                # 如果没有找到注释描述，尝试查找代码块前面的普通段落作为描述
                if not description:
                    for j in range(len(blocks)-1, -1, -1):
                        if blocks[j]['type'] == 'paragraph':
                            text = blocks[j]['text'].strip()
                            # 跳过空段落和特殊格式
                            if text and not text.startswith('<!--') and not text.endswith('-->'):
                                description = text
                                # 不移除这个段落，因为它可能是正常的文本内容
                                break

                # 如果还是没有描述，使用默认描述
                if not description:
                    description = f'代码示例 ({lang if lang else "未知语言"})'

                blocks.append({'type': 'code_table', 'lang': lang, 'code': '\n'.join(code_lines), 'description': description, 'in_appendix': in_appendix})
            continue

        # 公式 $$...$$（独占一行）
        if line.strip().startswith('$$'):
            formula_lines = [line.strip()[2:]]
            if not formula_lines[0].endswith('$$'):
                i += 1
                while i < len(lines):
                    if lines[i].strip().endswith('$$'):
                        formula_lines.append(lines[i].strip()[:-2])
                        break
                    formula_lines.append(lines[i].strip())
                    i += 1
            else:
                formula_lines[0] = formula_lines[0][:-2]
            i += 1
            blocks.append({'type': 'formula', 'latex': '\n'.join(formula_lines).strip()})
            continue

        # 表格
        if line.strip().startswith('|') and i + 1 < len(lines) and re.match(r'^\|[\s\-:|]+\|$', lines[i + 1].strip()):
            headers = [c.strip() for c in line.strip().split('|')]
            headers = [c for c in headers if c]
            i += 2  # 跳过表头和分隔行
            rows = []
            while i < len(lines) and lines[i].strip().startswith('|'):
                cells = [c.strip() for c in lines[i].strip().split('|')]
                cells = [c for c in cells if c]
                rows.append(cells)
                i += 1
            blocks.append({'type': 'table', 'headers': headers, 'rows': rows})
            continue

        # 引用块
        if re.match(r'^>\s*(.*)$', line):
            quote_lines = []
            while i < len(lines) and re.match(r'^>\s*(.*)$', lines[i]):
                quote_lines.append(re.sub(r'^>\s*', '', lines[i]))
                i += 1
            blocks.append({'type': 'blockquote', 'text': unescape_markdown(' '.join(quote_lines))})
            continue

        # 图片行 ![alt](url)
        img_match = re.match(r'^!\[([^\]]*)\]\(([^)]+)\)\s*$', line.strip())
        if img_match:
            blocks.append({'type': 'image', 'alt': img_match.group(1), 'url': img_match.group(2).strip()})
            i += 1
            continue

        # 无序列表
        bullet_match = re.match(r'^(\s*)[-*]\s+(.+)$', line)
        if bullet_match:
            indent = len(bullet_match.group(1))
            text = unescape_markdown(bullet_match.group(2).strip())
            blocks.append({'type': 'bullet', 'text': text, 'indent': indent})
            i += 1
            continue

        # 中文括号编号列表 （1）（2）... 作为普通段落保留编号文本
        cn_num_match = re.match(r'^(\s*)[\(（](\d+)[\)）]\s*(.*)', line)
        if cn_num_match:
            indent = len(cn_num_match.group(1))
            num_text = cn_num_match.group(2)
            text = unescape_markdown(cn_num_match.group(3).strip())
            full_text = f'（{num_text}）{text}'
            blocks.append({'type': 'paragraph', 'text': full_text, 'indent': indent})
            i += 1
            continue

        # 阿拉伯数字有序列表 1. 2. ... 保留原始编号文本
        num_match = re.match(r'^(\s*)(\d+\.)\s+(.+)$', line)
        if num_match:
            indent = len(num_match.group(1))
            num_text = num_match.group(2)  # 如 "1."
            text = unescape_markdown(num_match.group(3).strip())
            blocks.append({'type': 'numbered', 'text': text, 'num_text': num_text, 'indent': indent})
            i += 1
            continue

        # 普通段落
        para_lines = [line]
        i += 1
        while i < len(lines):
            next_line = lines[i]
            if next_line.strip() == '':
                i += 1
                break
            if (re.match(r'^#{1,6}\s', next_line) or
                re.match(r'^[-*]\s', next_line) or
                re.match(r'^\s*[\(（]\d+[\)）]', next_line) or
                re.match(r'^\s*\d+\.\s', next_line) or
                re.match(r'^---+\s*$', next_line.strip()) or
                re.match(r'^```', next_line.strip()) or
                re.match(r'^\$\$', next_line.strip()) or
                re.match(r'^!\[', next_line.strip()) or
                re.match(r'^>\s', next_line) or
                (next_line.strip().startswith('|') and i + 1 < len(lines) and re.match(r'^\|[\s\-:|]+\|$', lines[i + 1].strip()))):
                break
            para_lines.append(next_line)
            i += 1
        blocks.append({'type': 'paragraph', 'text': unescape_markdown(' '.join(para_lines))})

    return blocks


def parse_inline(text):
    """解析行内格式，返回 run 列表。"""
    runs = []
    pattern = r'(\*\*\*(.+?)\*\*\*|\*\*(.+?)\*\*|\*(.+?)\*|`([^`]+)`)'
    last_end = 0
    for m in re.finditer(pattern, text):
        if m.start() > last_end:
            runs.append({'text': text[last_end:m.start()], 'bold': False, 'italic': False, 'code': False})
        if m.group(2):
            runs.append({'text': m.group(2), 'bold': True, 'italic': True, 'code': False})
        elif m.group(3):
            runs.append({'text': m.group(3), 'bold': True, 'italic': False, 'code': False})
        elif m.group(4):
            runs.append({'text': m.group(4), 'bold': False, 'italic': True, 'code': False})
        elif m.group(5):
            runs.append({'text': m.group(5), 'bold': False, 'italic': False, 'code': True})
        last_end = m.end()
    if last_end < len(text):
        runs.append({'text': text[last_end:], 'bold': False, 'italic': False, 'code': False})
    return runs


def strip_markdown_formatting(text):
    """去除 Markdown 格式标记，只保留纯文本。"""
    text = re.sub(r'(?<!\$)\$(?!\$)(.+?)(?<!\$)\$(?!\$)', r'\1', text)
    text = re.sub(r'\*\*\*(.+?)\*\*\*', r'\1', text)
    text = re.sub(r'\*\*(.+?)\*\*', r'\1', text)
    text = re.sub(r'\*(.+?)\*', r'\1', text)
    text = re.sub(r'`(.+?)`', r'\1', text)
    text = re.sub(r'\[(.+?)\]\(.+?\)', r'\1', text)
    return text.strip()


# ============================================================
# 3. XML 辅助函数
# ============================================================

def make_text_run(text, bold=False, italic=False, code=False, font_ascii=None, font_east_asia=None, hint=None, sz=None, unbold=False, skip_punct_norm=False, theme_fonts=False):
    """创建一个文本 run。自动对非代码文本进行中文标点规范化。
    skip_punct_norm=True 时使用表格标点规范（中文标点→英文标点）。
    theme_fonts=True 时添加 asciiTheme/hAnsiTheme/cstheme='minorHAnsi'（匹配模板摘要段落格式）。
    ASCII 字体默认 Calibri（数字/英文），东亚字体默认宋体（中文）。"""
    # 对非代码文本进行标点规范化
    if not code and text:
        if skip_punct_norm:
            text = _normalize_table_punctuation(text)
        else:
            text = _normalize_chinese_punctuation(text)
    r = ET.Element(f'{{{W}}}r')
    rpr = ET.SubElement(r, f'{{{W}}}rPr')

    has_fonts = False
    # 所有文本都设置中英文字体分离：ASCII=Calibri，东亚=宋体
    if True:
        fonts = ET.SubElement(rpr, f'{{{W}}}rFonts')
        has_fonts = True
        if code:
            fonts.set(f'{{{W}}}ascii', 'Calibri')
            fonts.set(f'{{{W}}}hAnsi', 'Calibri')
            fonts.set(f'{{{W}}}eastAsia', '宋体')
        else:
            effective_ascii = font_ascii or 'Calibri'
            effective_ea = font_east_asia or '宋体'
            fonts.set(f'{{{W}}}ascii', effective_ascii)
            fonts.set(f'{{{W}}}hAnsi', effective_ascii)
            fonts.set(f'{{{W}}}eastAsia', effective_ea)
        if hint:
            fonts.set(f'{{{W}}}hint', hint)
        if theme_fonts:
            fonts.set(f'{{{W}}}asciiTheme', 'minorHAnsi')
            fonts.set(f'{{{W}}}hAnsiTheme', 'minorHAnsi')
            fonts.set(f'{{{W}}}cstheme', 'minorHAnsi')

    if bold:
        ET.SubElement(rpr, f'{{{W}}}b')
    elif unbold:
        # 显式关闭加粗（覆盖样式继承的 bold）
        b = ET.SubElement(rpr, f'{{{W}}}b')
        b.set(f'{{{W}}}val', '0')
    if italic:
        ET.SubElement(rpr, f'{{{W}}}i')
        ET.SubElement(rpr, f'{{{W}}}iCs')

    if sz:
        s = ET.SubElement(rpr, f'{{{W}}}sz')
        s.set(f'{{{W}}}val', str(sz))
        scs = ET.SubElement(rpr, f'{{{W}}}szCs')
        scs.set(f'{{{W}}}val', str(sz))

    t = ET.SubElement(r, f'{{{W}}}t')
    t.text = text
    t.set(f'{XML_NS}space', 'preserve')
    return r


def add_text_to_paragraph(p, text, bold=False, italic=False, font_ascii=None, font_east_asia=None, sz=None, skip_punct_norm=False):
    """将文本添加到段落中，解析行内格式和行内公式 $...$。"""
    # 按 $...$ 分割（排除 $$），交替处理文本和公式
    math_pattern = r'(?<!\$)\$(?!\$)(.+?)(?<!\$)\$(?!\$)'
    last_end = 0

    for m in re.finditer(math_pattern, text):
        # 公式前的文本
        text_before = text[last_end:m.start()]
        if text_before:
            _add_text_runs(p, text_before, bold, italic, font_ascii, font_east_asia, sz, skip_punct_norm)

        # 行内公式
        latex = m.group(1)
        omml = latex_to_omml(latex)
        if omml is not None:
            p.append(omml)
        else:
            # 回退：斜体 Cambria Math 文本
            r = make_text_run(f'${latex}$', font_ascii='Cambria Math', font_east_asia='Cambria Math', italic=True)
            p.append(r)

        last_end = m.end()

    # 最后一段文本
    remaining = text[last_end:]
    if remaining:
        _add_text_runs(p, remaining, bold, italic, font_ascii, font_east_asia, sz, skip_punct_norm)


def _add_text_runs(p, text, bold=False, italic=False, font_ascii=None, font_east_asia=None, sz=None, skip_punct_norm=False):
    """将纯文本（不含公式）按行内格式解析后添加为 runs。"""
    inline_runs = parse_inline(text)
    # 检测整体文本是否包含中文（用于代码片段内的冒号转换）
    has_chinese_context = bool(re.search(r'[一-鿿]', text))
    for run_info in inline_runs:
        if not run_info['text']:
            continue
        run_text = run_info['text']
        # 代码片段在中文语境中：将所有英文冒号转为全角（检测工具要求）
        # 跳过 URL 协议分隔符 :// 和端口号 :\d
        if run_info.get('code') and has_chinese_context:
            run_text = re.sub(r':(?!//|\d)', '：', run_text)
        r = make_text_run(
            run_text,
            bold=run_info['bold'] or bold,
            italic=run_info['italic'] or italic,
            code=run_info.get('code', False),
            font_ascii=font_ascii or 'Calibri',
            font_east_asia=font_east_asia,
            hint='eastAsia',
            sz=sz,
            skip_punct_norm=skip_punct_norm,
        )
        p.append(r)


# ============================================================
# 4. 段落构建函数
# ============================================================

def _convert_heading_numbering(text, level):
    """将标题中的阿拉伯数字编号转换为中文编号格式。
    H1: "1 绪论" → "一、绪论"
    H2: "1.1 研究背景" → "（一）研究背景"
    H3: "1.1.1 系统需实现的目标" → "1. 系统需实现的目标"
    保留无编号标题（如"结论"、"致谢"等）不变。
    """
    clean = strip_markdown_formatting(text)

    # 无编号标题（结论、致谢、参考文献、附录等）不转换
    if any(clean == kw or clean.startswith(kw + ' ') for kw in _UNNUMBERED_HEADING_KEYWORDS):
        return text

    CN_NUMS = '一二三四五六七八九十'

    if level == 1:
        # "1 绪论" → "一、绪论"
        m = re.match(r'^(\d+)\s+(.+)$', clean)
        if m:
            num = int(m.group(1))
            if 1 <= num <= 10:
                return f'{CN_NUMS[num-1]}、{m.group(2)}'
            return text
    elif level == 2:
        # "1.1 研究背景" → "（一）研究背景"
        # 需要跟踪当前章节内的节号，这里简化处理：提取第二位数字
        m = re.match(r'^\d+\.(\d+)\s+(.+)$', clean)
        if m:
            num = int(m.group(1))
            if 1 <= num <= 26:
                # 使用中文括号编号
                cn_num_map = list('一二三四五六七八九十') + \
                             ['十一','十二','十三','十四','十五','十六','十七','十八','十九','二十', \
                              '二十一','二十二','二十三','二十四','二十五','二十六']
                return f'（{cn_num_map[num-1]}）{m.group(2)}'
            return text
    elif level == 3:
        # "1.1.1 系统需实现的目标" → "1. 系统需实现的目标"
        m = re.match(r'^\d+\.\d+\.(\d+)\s+(.+)$', clean)
        if m:
            num = int(m.group(1))
            return f'{num}. {m.group(2)}'

    return text


def build_heading(block, bookmark_name=None, bookmark_id=None):
    """构建标题段落。H1→style 12, H2→style afc, H3→style afa+outlineLvl。
    标题编号已转换为中文编号格式（一、（一）、1.）。"""
    level = block['level']
    text = block['text']
    # 转换标题编号格式
    text = _convert_heading_numbering(text, level)

    if level == 1:
        p = ET.Element(f'{{{W}}}p')
        ppr = ET.SubElement(p, f'{{{W}}}pPr')
        pstyle = ET.SubElement(ppr, f'{{{W}}}pStyle')
        pstyle.set(f'{{{W}}}val', '12')
        adj = ET.SubElement(ppr, f'{{{W}}}adjustRightInd')
        adj.set(f'{{{W}}}val', '0')
        snap = ET.SubElement(ppr, f'{{{W}}}snapToGrid')
        snap.set(f'{{{W}}}val', '0')
        # 范文 H1 无显式 jc/spacing，居中和行距由样式继承

    elif level == 2:
        p = ET.Element(f'{{{W}}}p')
        ppr = ET.SubElement(p, f'{{{W}}}pPr')
        pstyle = ET.SubElement(ppr, f'{{{W}}}pStyle')
        pstyle.set(f'{{{W}}}val', 'afc')
        adj = ET.SubElement(ppr, f'{{{W}}}adjustRightInd')
        adj.set(f'{{{W}}}val', '0')
        snap = ET.SubElement(ppr, f'{{{W}}}snapToGrid')
        snap.set(f'{{{W}}}val', '0')
        # H2 间距由样式继承，不显式设置（范文 H2 段落无 spacing 属性）

    else:
        p = ET.Element(f'{{{W}}}p')
        ppr = ET.SubElement(p, f'{{{W}}}pPr')
        pstyle = ET.SubElement(ppr, f'{{{W}}}pStyle')
        # 不用样式 'a'（绑定了 Wingdings 编号列表，WPS 兼容性差），改用正文样式
        pstyle.set(f'{{{W}}}val', 'afa')
        adj = ET.SubElement(ppr, f'{{{W}}}adjustRightInd')
        adj.set(f'{{{W}}}val', '0')
        snap = ET.SubElement(ppr, f'{{{W}}}snapToGrid')
        snap.set(f'{{{W}}}val', '0')
        spacing = ET.SubElement(ppr, f'{{{W}}}spacing')
        spacing.set(f'{{{W}}}line', '360')
        spacing.set(f'{{{W}}}lineRule', 'auto')
        ind = ET.SubElement(ppr, f'{{{W}}}ind')
        ind.set(f'{{{W}}}left', '0')
        ind.set(f'{{{W}}}firstLine', '0')
        # 大纲级别 2（三级标题），让目录能识别
        outlineLvl = ET.SubElement(ppr, f'{{{W}}}outlineLvl')
        outlineLvl.set(f'{{{W}}}val', '2')
        rpr_in_ppr = ET.SubElement(ppr, f'{{{W}}}rPr')
        fonts = ET.SubElement(rpr_in_ppr, f'{{{W}}}rFonts')
        fonts.set(f'{{{W}}}ascii', 'Calibri')
        fonts.set(f'{{{W}}}hAnsi', 'Calibri')
        fonts.set(f'{{{W}}}eastAsia', '黑体')
        # 字号：小四（12磅）
        sz = ET.SubElement(rpr_in_ppr, f'{{{W}}}sz')
        sz.set(f'{{{W}}}val', '24')
        szCs = ET.SubElement(rpr_in_ppr, f'{{{W}}}szCs')
        szCs.set(f'{{{W}}}val', '24')

    heading_sz = {1: None, 2: None, 3: 24}.get(level, 24)
    # H1/H2 不额外加粗（黑体本身视觉效果即为粗体），sz 由样式继承（范文 H1/H2 runs 无 sz）
    add_text_to_paragraph(p, text, bold=False, font_east_asia='黑体', sz=heading_sz)

    if bookmark_name:
        bm_id = bookmark_id if bookmark_id is not None else 0
        bm_start = ET.SubElement(p, f'{{{W}}}bookmarkStart')
        bm_start.set(f'{{{W}}}id', str(bm_id))
        bm_start.set(f'{{{W}}}name', bookmark_name)
        bm_end = ET.SubElement(p, f'{{{W}}}bookmarkEnd')
        bm_end.set(f'{{{W}}}id', str(bm_id))

    return p


def build_body_paragraph(block, no_indent=False):
    """构建正文段落（style aa, 首行缩进）。表格标题用 af4 样式。"""
    text = block['text']
    p = ET.Element(f'{{{W}}}p')
    ppr = ET.SubElement(p, f'{{{W}}}pPr')
    pstyle = ET.SubElement(ppr, f'{{{W}}}pStyle')

    # 表格标题（如 "**表2-1 xxx**" 或 "续表2-1 xxx"）使用图注样式
    if re.match(r'^\*{0,2}(续)?表\d', text):
        pstyle.set(f'{{{W}}}val', 'a4')
        jc = ET.SubElement(ppr, f'{{{W}}}jc')
        jc.set(f'{{{W}}}val', 'center')
        adj = ET.SubElement(ppr, f'{{{W}}}adjustRightInd')
        adj.set(f'{{{W}}}val', '0')
        snap = ET.SubElement(ppr, f'{{{W}}}snapToGrid')
        snap.set(f'{{{W}}}val', '0')
        # 范文表注行距 1.5 倍
        cap_spacing = ET.SubElement(ppr, f'{{{W}}}spacing')
        cap_spacing.set(f'{{{W}}}line', '360')
        cap_spacing.set(f'{{{W}}}lineRule', 'auto')
        # 去除 Markdown 加粗标记和末尾标点
        clean_text = strip_markdown_formatting(text)
        clean_text = re.sub(r'[。,，.]+$', '', clean_text)
        # 范文表注格式：asciiTheme='minorHAnsi' + eastAsia='黑体'，无显式 ascii
        r = ET.SubElement(p, f'{{{W}}}r')
        rpr = ET.SubElement(r, f'{{{W}}}rPr')
        rfonts = ET.SubElement(rpr, f'{{{W}}}rFonts')
        rfonts.set(f'{{{W}}}asciiTheme', 'minorHAnsi')
        rfonts.set(f'{{{W}}}hAnsiTheme', 'minorHAnsi')
        rfonts.set(f'{{{W}}}cstheme', 'minorHAnsi')
        rfonts.set(f'{{{W}}}eastAsia', '黑体')
        rsz = ET.SubElement(rpr, f'{{{W}}}sz')
        rsz.set(f'{{{W}}}val', '21')
        rszcs = ET.SubElement(rpr, f'{{{W}}}szCs')
        rszcs.set(f'{{{W}}}val', '21')
        rt = ET.SubElement(r, f'{{{W}}}t')
        rt.text = clean_text
        rt.set(f'{XML_NS}space', 'preserve')
        return p

    # 图注（如 "图2-1 xxx"）使用图注样式
    if re.match(r'^图\d', text):
        pstyle.set(f'{{{W}}}val', 'a4')
        jc = ET.SubElement(ppr, f'{{{W}}}jc')
        jc.set(f'{{{W}}}val', 'center')
        adj = ET.SubElement(ppr, f'{{{W}}}adjustRightInd')
        adj.set(f'{{{W}}}val', '0')
        snap = ET.SubElement(ppr, f'{{{W}}}snapToGrid')
        snap.set(f'{{{W}}}val', '0')
        cap_spacing = ET.SubElement(ppr, f'{{{W}}}spacing')
        cap_spacing.set(f'{{{W}}}line', '360')
        cap_spacing.set(f'{{{W}}}lineRule', 'auto')
        clean_text = strip_markdown_formatting(text)
        clean_text = re.sub(r'[。,，.]+$', '', clean_text)
        r = ET.SubElement(p, f'{{{W}}}r')
        rpr = ET.SubElement(r, f'{{{W}}}rPr')
        rfonts = ET.SubElement(rpr, f'{{{W}}}rFonts')
        rfonts.set(f'{{{W}}}asciiTheme', 'minorHAnsi')
        rfonts.set(f'{{{W}}}hAnsiTheme', 'minorHAnsi')
        rfonts.set(f'{{{W}}}cstheme', 'minorHAnsi')
        rfonts.set(f'{{{W}}}eastAsia', '黑体')
        rsz = ET.SubElement(rpr, f'{{{W}}}sz')
        rsz.set(f'{{{W}}}val', '21')
        rszcs = ET.SubElement(rpr, f'{{{W}}}szCs')
        rszcs.set(f'{{{W}}}val', '21')
        rt = ET.SubElement(r, f'{{{W}}}t')
        rt.text = clean_text
        rt.set(f'{XML_NS}space', 'preserve')
        return p

    pstyle.set(f'{{{W}}}val', 'afa')
    adj = ET.SubElement(ppr, f'{{{W}}}adjustRightInd')
    adj.set(f'{{{W}}}val', '0')
    snap = ET.SubElement(ppr, f'{{{W}}}snapToGrid')
    snap.set(f'{{{W}}}val', '0')
    ind = ET.SubElement(ppr, f'{{{W}}}ind')
    if no_indent:
        ind.set(f'{{{W}}}firstLineChars', '0')
        ind.set(f'{{{W}}}firstLine', '0')
    else:
        ind.set(f'{{{W}}}firstLineChars', '200')
        ind.set(f'{{{W}}}firstLine', '480')
    add_text_to_paragraph(p, text)
    return p


def build_bullet_paragraph(block):
    """构建无序列表段落。首行缩进2字符，项目符号由numPr自动生成。"""
    text = block['text']
    p = ET.Element(f'{{{W}}}p')
    ppr = ET.SubElement(p, f'{{{W}}}pPr')
    pstyle = ET.SubElement(ppr, f'{{{W}}}pStyle')
    pstyle.set(f'{{{W}}}val', 'afa')
    adj = ET.SubElement(ppr, f'{{{W}}}adjustRightInd')
    adj.set(f'{{{W}}}val', '0')
    snap = ET.SubElement(ppr, f'{{{W}}}snapToGrid')
    snap.set(f'{{{W}}}val', '0')
    numpr = ET.SubElement(ppr, f'{{{W}}}numPr')
    ilvl = ET.SubElement(numpr, f'{{{W}}}ilvl')
    ilvl.set(f'{{{W}}}val', '0')
    numid = ET.SubElement(numpr, f'{{{W}}}numId')
    numid.set(f'{{{W}}}val', '2')
    ind = ET.SubElement(ppr, f'{{{W}}}ind')
    ind.set(f'{{{W}}}firstLineChars', '200')
    ind.set(f'{{{W}}}firstLine', '480')
    add_text_to_paragraph(p, text)
    return p


def build_numbered_paragraph(block):
    """构建编号列表段落。编号文本作为正文一部分，首行缩进2字符。"""
    text = block['text']
    num_text = block.get('num_text', '1.')
    p = ET.Element(f'{{{W}}}p')
    ppr = ET.SubElement(p, f'{{{W}}}pPr')
    pstyle = ET.SubElement(ppr, f'{{{W}}}pStyle')
    pstyle.set(f'{{{W}}}val', 'afa')
    adj = ET.SubElement(ppr, f'{{{W}}}adjustRightInd')
    adj.set(f'{{{W}}}val', '0')
    snap = ET.SubElement(ppr, f'{{{W}}}snapToGrid')
    snap.set(f'{{{W}}}val', '0')
    ind = ET.SubElement(ppr, f'{{{W}}}ind')
    ind.set(f'{{{W}}}firstLineChars', '200')
    ind.set(f'{{{W}}}firstLine', '480')
    # 将 "1." 转为 "（1）" 避免与三级标题编号混淆
    m = re.match(r'^(\d+)\.$', num_text)
    if m:
        display_text = f'（{m.group(1)}）{text}'
    else:
        display_text = f'{num_text} {text}'
    add_text_to_paragraph(p, display_text)
    return p


def build_blockquote(block):
    """构建引用块段落（图注等，style a4, 居中，行距1.5倍）。"""
    text = block['text']
    p = ET.Element(f'{{{W}}}p')
    ppr = ET.SubElement(p, f'{{{W}}}pPr')
    pstyle = ET.SubElement(ppr, f'{{{W}}}pStyle')
    pstyle.set(f'{{{W}}}val', 'a4')
    jc = ET.SubElement(ppr, f'{{{W}}}jc')
    jc.set(f'{{{W}}}val', 'center')
    sp = ET.SubElement(ppr, f'{{{W}}}spacing')
    sp.set(f'{{{W}}}line', '360')
    sp.set(f'{{{W}}}lineRule', 'auto')
    add_text_to_paragraph(p, text)
    return p


def render_mermaid_to_png(mermaid_code):
    """用 mmdc CLI 将 mermaid 源码渲染为 PNG 图片。返回 (img_bytes, width_px, height_px) 或 None。
    策略：先用 -w 1200 渲染探测实际宽高比，根据比例选择最佳渲染参数。
    - 宽高比 < 0.5（甘特图/时序图等超宽矮图）：用较小宽度渲染，避免缩放后文字太小
    - 宽高比 >= 0.5 的正常图：用 3200+scale 保证清晰
    - 自动提升 fontSize 以确保文字在 15cm 宽度下清晰可读
    """
    # 优先使用系统已安装的 Chrome，避免 Puppeteer 下载版本不匹配
    chrome_path = '/Applications/Google Chrome.app/Contents/MacOS/Google Chrome'
    if os.path.exists(chrome_path) and 'PUPPETEER_EXECUTABLE_PATH' not in os.environ:
        os.environ['PUPPETEER_EXECUTABLE_PATH'] = chrome_path

    tmpdir = tempfile.mkdtemp()
    mmd_path = os.path.join(tmpdir, 'mermaid_input.mmd')
    png_path = os.path.join(tmpdir, 'mermaid_output.png')

    try:
        # 根据图的宽高比按比例调整 fontSize：越扁的图文字越大，确保 15cm 宽度下清晰
        # 先用基础 fontSize 渲染一次探测宽高比，再决定最终 fontSize
        base_fontsize = 22
        enhanced_code = mermaid_code
        # 替换已有的 fontSize（如果存在）
        if 'fontSize' in enhanced_code:
            enhanced_code = re.sub(r"""['"]fontSize['"]:\s*['"]\d+px['"]""", f"'fontSize': '{base_fontsize}px'", enhanced_code)
        # 确保 fontFamily 被设置（注入或替换）
        if 'fontFamily' in enhanced_code:
            enhanced_code = re.sub(r"""['"]fontFamily['"]:\s*['"][^'"]*['"]""", "'fontFamily': 'SimSun, Times New Roman, serif'", enhanced_code)
        elif 'themeVariables' in enhanced_code:
            # 在 themeVariables 块中追加 fontFamily + fontSize（如果上面没替换到）
            insert_vars = ''
            if 'fontSize' not in enhanced_code:
                insert_vars += f"'fontSize': '{base_fontsize}px',\n      "
            insert_vars += "'fontFamily': 'SimSun, Times New Roman, serif',"
            enhanced_code = enhanced_code.replace("'themeVariables': {", f"'themeVariables': {{\n      {insert_vars}", 1)
        if '%%{' not in enhanced_code and 'init' not in enhanced_code:
            font_config = f"""%%{{
  init: {{
    'theme': 'default',
    'themeVariables': {{
      'fontSize': '{base_fontsize}px',
      'fontFamily': 'SimSun, Times New Roman, serif',
      'primaryTextColor': '#000000',
      'primaryBorderColor': '#000000',
      'lineColor': '#000000',
      'secondaryColor': '#0066cc',
      'tertiaryColor': '#ffffff'
    }}
  }}
}}%%
"""
            enhanced_code = font_config + enhanced_code

        with open(mmd_path, 'w', encoding='utf-8') as f:
            f.write(enhanced_code)

        # 探测渲染：用较小宽度快速获取图的宽高比
        result = subprocess.run(
            ['mmdc', '-i', mmd_path, '-o', png_path, '-w', '1200', '-b', 'white'],
            capture_output=True, timeout=120
        )

        if not os.path.exists(png_path):
            print(f'  警告: mermaid 渲染失败: {result.stderr.decode("utf-8", errors="replace")}', file=sys.stderr)
            return None

        with open(png_path, 'rb') as f:
            img_bytes = f.read()

        w = int.from_bytes(img_bytes[16:20], 'big')
        h = int.from_bytes(img_bytes[20:24], 'big')
        aspect = h / w if w > 0 else 1

        # 根据宽高比调整 fontSize：扁图（低 aspect）需要更大字号
        # aspect >= 1.0 → 22px, aspect 0.5~1.0 → 24px, aspect 0.3~0.5 → 28px, aspect < 0.3 → 32px
        if aspect >= 1.0:
            final_fontsize = 22
        elif aspect >= 0.5:
            final_fontsize = 24
        elif aspect >= 0.3:
            final_fontsize = 28
        else:
            final_fontsize = 32

        if final_fontsize != base_fontsize:
            # 需要用更大的 fontSize 重新渲染
            enhanced_code = re.sub(r"""['"]fontSize['"]:\s*['"]\d+px['"]""", f"'fontSize': '{final_fontsize}px'", enhanced_code)
            with open(mmd_path, 'w', encoding='utf-8') as f:
                f.write(enhanced_code)

        if aspect < 0.5:
            # 超宽矮图：用高 scale 渲染
            os.remove(png_path)
            result = subprocess.run(
                ['mmdc', '-i', mmd_path, '-o', png_path, '-w', '2400', '-b', 'white', '--scale', '3'],
                capture_output=True, timeout=120
            )
            if os.path.exists(png_path):
                with open(png_path, 'rb') as f:
                    img_bytes = f.read()
                w = int.from_bytes(img_bytes[16:20], 'big')
                h = int.from_bytes(img_bytes[20:24], 'big')
        elif w < 1200 and w > 0:
            # 小图：用更高分辨率渲染
            os.remove(png_path)
            result = subprocess.run(
                ['mmdc', '-i', mmd_path, '-o', png_path, '-w', '3200', '-b', 'white', '--scale', '2.5'],
                capture_output=True, timeout=120
            )
            if os.path.exists(png_path):
                with open(png_path, 'rb') as f:
                    img_bytes = f.read()
                w = int.from_bytes(img_bytes[16:20], 'big')
                h = int.from_bytes(img_bytes[20:24], 'big')
        else:
            # 正常图：用 3200 + scale 渲染
            os.remove(png_path)
            result = subprocess.run(
                ['mmdc', '-i', mmd_path, '-o', png_path, '-w', '3200', '-b', 'white', '--scale', '2'],
                capture_output=True, timeout=120
            )
            if os.path.exists(png_path):
                with open(png_path, 'rb') as f:
                    img_bytes = f.read()
                w = int.from_bytes(img_bytes[16:20], 'big')
                h = int.from_bytes(img_bytes[20:24], 'big')

        return (img_bytes, w, h)
    except Exception as e:
        print(f'  警告: mermaid 渲染异常: {e}', file=sys.stderr)
        return None
    finally:
        # 清理临时文件
        shutil.rmtree(tmpdir, ignore_errors=True)


def build_mermaid_block(block, media_files_dict=None, rels_counter=None):
    """将 mermaid 代码渲染为 PNG 并嵌入 docx。返回段落列表（图片+图注）。"""
    mermaid_code = block.get('code', '')
    caption = block.get('caption', '')

    result = render_mermaid_to_png(mermaid_code)
    if result is None:
        # 渲染失败，回退到代码块输出
        return build_codeblock(block)

    img_bytes, width_px, height_px = result

    # 保存到 media 目录
    img_filename = f'mermaid_{hash(mermaid_code) & 0xFFFFFFFF:08x}.png'
    media_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'template', 'word', 'media')
    os.makedirs(media_dir, exist_ok=True)
    media_path = os.path.join(media_dir, img_filename)
    with open(media_path, 'wb') as f:
        f.write(img_bytes)

    if media_files_dict is not None:
        media_files_dict[f'word/media/{img_filename}'] = f'word/media/{img_filename}'

    # 统一宽度 15cm，保持宽高比
    aspect = height_px / width_px if width_px > 0 else 1
    cx = 15 * 360000  # 15cm = 5,400,000 EMU
    cy = int(cx * aspect)

    r_id, cx, cy = _make_image_rel(media_files_dict, rels_counter, img_filename, img_bytes,
                                    cx / 9525, cy / 9525, no_scale=True)

    # 构建图片段落
    p_img = _build_image_paragraph(r_id, cx, cy)
    elements = [p_img]

    # 添加图注段落（af4 样式，居中，黑体五号）
    if caption:
        p_cap = ET.Element(f'{{{W}}}p')
        p_cap_ppr = ET.SubElement(p_cap, f'{{{W}}}pPr')
        p_cap_style = ET.SubElement(p_cap_ppr, f'{{{W}}}pStyle')
        p_cap_style.set(f'{{{W}}}val', 'a4')
        p_cap_jc = ET.SubElement(p_cap_ppr, f'{{{W}}}jc')
        p_cap_jc.set(f'{{{W}}}val', 'center')
        p_cap_adj = ET.SubElement(p_cap_ppr, f'{{{W}}}adjustRightInd')
        p_cap_adj.set(f'{{{W}}}val', '0')
        p_cap_snap = ET.SubElement(p_cap_ppr, f'{{{W}}}snapToGrid')
        p_cap_snap.set(f'{{{W}}}val', '0')
        p_cap_spacing = ET.SubElement(p_cap_ppr, f'{{{W}}}spacing')
        p_cap_spacing.set(f'{{{W}}}line', '360')
        p_cap_spacing.set(f'{{{W}}}lineRule', 'auto')
        cap_run = ET.SubElement(p_cap, f'{{{W}}}r')
        cap_rpr = ET.SubElement(cap_run, f'{{{W}}}rPr')
        cap_fonts = ET.SubElement(cap_rpr, f'{{{W}}}rFonts')
        cap_fonts.set(f'{{{W}}}asciiTheme', 'minorHAnsi')
        cap_fonts.set(f'{{{W}}}hAnsiTheme', 'minorHAnsi')
        cap_fonts.set(f'{{{W}}}cstheme', 'minorHAnsi')
        cap_fonts.set(f'{{{W}}}eastAsia', '黑体')
        cap_sz = ET.SubElement(cap_rpr, f'{{{W}}}sz')
        cap_sz.set(f'{{{W}}}val', '21')
        cap_scs = ET.SubElement(cap_rpr, f'{{{W}}}szCs')
        cap_scs.set(f'{{{W}}}val', '21')
        cap_t = ET.SubElement(cap_run, f'{{{W}}}t')
        cap_t.text = caption
        cap_t.set(f'{XML_NS}space', 'preserve')
        elements.append(p_cap)

    return elements


def build_codeblock(block):
    """构建代码块（等宽字体，无缩进）。"""
    code = block['code']
    lines = code.split('\n')

    # 每行创建一个独立段落
    paragraphs = []
    for line in lines:
        p = ET.Element(f'{{{W}}}p')
        ppr = ET.SubElement(p, f'{{{W}}}pPr')
        pstyle = ET.SubElement(ppr, f'{{{W}}}pStyle')
        pstyle.set(f'{{{W}}}val', 'afa')
        adj = ET.SubElement(ppr, f'{{{W}}}adjustRightInd')
        adj.set(f'{{{W}}}val', '0')
        snap = ET.SubElement(ppr, f'{{{W}}}snapToGrid')
        snap.set(f'{{{W}}}val', '0')
        ind = ET.SubElement(ppr, f'{{{W}}}ind')
        ind.set(f'{{{W}}}firstLineChars', '0')
        ind.set(f'{{{W}}}firstLine', '0')

        r = make_text_run(line or ' ', code=True)
        p.append(r)
        paragraphs.append(p)

    return paragraphs  # 返回段落列表


def build_code_table(block):
    """构建代码表格：单行一列，仅包含代码。表注由调用方在表格上方单独渲染。"""
    code = block['code']

    # 创建表格：一列
    tbl = ET.Element(f'{{{W}}}tbl')

    # 表格属性
    tblPr = ET.SubElement(tbl, f'{{{W}}}tblPr')
    tblW = ET.SubElement(tblPr, f'{{{W}}}tblW')
    tblW.set(f'{{{W}}}w', '0')
    tblW.set(f'{{{W}}}type', 'auto')
    jc = ET.SubElement(tblPr, f'{{{W}}}jc')
    jc.set(f'{{{W}}}val', 'center')
    tblBorders = ET.SubElement(tblPr, f'{{{W}}}tblBorders')
    for border_name in ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']:
        b = ET.SubElement(tblBorders, f'{{{W}}}{border_name}')
        b.set(f'{{{W}}}val', 'single')
        b.set('sz', '4')
        b.set('space', '0')
        b.set('color', 'auto')
    tblLayout = ET.SubElement(tblPr, f'{{{W}}}tblLayout')
    tblLayout.set(f'{{{W}}}type', 'fixed')

    # 列定义：一列，全宽度
    total_width = 8312
    col_width = total_width
    tblGrid = ET.SubElement(tbl, f'{{{W}}}tblGrid')
    gc = ET.SubElement(tblGrid, f'{{{W}}}gridCol')
    gc.set(f'{{{W}}}w', str(col_width))

    # 判断代码是否包含中文语境（代码含中文注释）
    code_has_chinese = bool(re.search(r'[一-鿿]', code))

    # 单行：代码
    tr = ET.SubElement(tbl, f'{{{W}}}tr')
    tc = ET.SubElement(tr, f'{{{W}}}tc')
    tcPr = ET.SubElement(tc, f'{{{W}}}tcPr')
    tcW = ET.SubElement(tcPr, f'{{{W}}}tcW')
    tcW.set(f'{{{W}}}w', str(col_width))
    tcW.set(f'{{{W}}}type', 'dxa')
    vAlign = ET.SubElement(tcPr, f'{{{W}}}vAlign')
    vAlign.set(f'{{{W}}}val', 'center')

    # 代码使用等宽字体，每行一个独立段落，五号字，1.5倍行距
    code_lines = code.split('\n')
    for line_num, line in enumerate(code_lines):
        if line.strip() == '':
            line = ' '  # 空行用空格代替

        p = ET.SubElement(tc, f'{{{W}}}p')
        ppr = ET.SubElement(p, f'{{{W}}}pPr')
        pstyle = ET.SubElement(ppr, f'{{{W}}}pStyle')
        pstyle.set(f'{{{W}}}val', 'afa')
        adj = ET.SubElement(ppr, f'{{{W}}}adjustRightInd')
        adj.set(f'{{{W}}}val', '0')
        snap = ET.SubElement(ppr, f'{{{W}}}snapToGrid')
        snap.set(f'{{{W}}}val', '0')
        spacing = ET.SubElement(ppr, f'{{{W}}}spacing')
        spacing.set(f'{{{W}}}line', '360')
        spacing.set(f'{{{W}}}lineRule', 'auto')
        ind = ET.SubElement(ppr, f'{{{W}}}ind')
        ind.set(f'{{{W}}}firstLineChars', '0')
        ind.set(f'{{{W}}}firstLine', '0')

        # 中文语境中代码行的英文冒号也转为全角（跳过 :// URL协议分隔符 和端口号 :\d）
        if code_has_chinese:
            line = re.sub(r':(?!//|\d)', '：', line)
        r = make_text_run(line, code=True, sz=21)
        p.append(r)

    return tbl


def _render_formula_with_xelatex(latex):
    """用 xelatex 将 LaTeX 公式编译为 PNG 图片。\tag{} 编号直接在 LaTeX 中渲染。"""
    formula_text = latex.strip().strip('$').strip()

    # 有 \tag{} 的公式必须放在 equation 环境中（$...$ 内联模式不支持 \tag）
    has_tag = r'\tag{' in formula_text

    # scalebox 1.6 配合 equation* 使公式字号接近小四号，且最长公式不溢出 A4 正文区
    FORMULA_SCALE = 1.6
    doc = r"\documentclass[preview,border=2pt]{standalone}" + "\n"
    doc += r"\usepackage{amsmath,amssymb}" + "\n"
    doc += r"\usepackage{unicode-math}" + "\n"
    doc += r"\usepackage{graphicx}" + "\n"
    doc += r"\begin{document}" + "\n"
    if has_tag:
        doc += r"\scalebox{" + str(FORMULA_SCALE) + r"}{\vbox{" + "\n"
        doc += r"\begin{equation*}" + "\n"
        doc += formula_text + "\n"
        doc += r"\end{equation*}" + "\n"
        doc += r"}}" + "\n"
    else:
        doc += r"\scalebox{" + str(FORMULA_SCALE) + r"}{$\displaystyle " + formula_text + r"$}" + "\n"
    doc += r"\end{document}"

    # 在临时目录中编译
    tmpdir = tempfile.mkdtemp()
    try:
        tex_path = os.path.join(tmpdir, 'formula.tex')
        with open(tex_path, 'w') as f:
            f.write(doc)

        # xelatex 编译
        result = subprocess.run(
            ['xelatex', '-interaction=nonstopmode', '-output-directory', tmpdir, tex_path],
            capture_output=True, timeout=30
        )

        # 检查 pdf 是否生成
        pdf_path = os.path.join(tmpdir, 'formula.pdf')
        if not os.path.exists(pdf_path):
            raise RuntimeError(f'xelatex 编译失败: {result.stderr.decode()}')

        # pdf 转 png（使用 PyMuPDF 以 1200 DPI 渲染，保证字体清晰锐利）
        import pymupdf
        pdfdoc = pymupdf.open(pdf_path)
        page = pdfdoc[0]
        zoom = 1200 / 72  # 1200 DPI
        mat = pymupdf.Matrix(zoom, zoom)
        pix = page.get_pixmap(matrix=mat, alpha=False)
        img_bytes = pix.tobytes('png')
        pdfdoc.close()
        return img_bytes
    finally:
        shutil.rmtree(tmpdir, ignore_errors=True)


def build_formula_paragraph(block, media_files_dict=None, rels_counter=None):
    """构建公式段落：将 LaTeX 渲染为 PNG 图片并嵌入 docx。\tag{} 编号直接由 LaTeX 渲染。"""
    latex = block['latex']
    try:
        img_bytes = _render_formula_with_xelatex(latex)

        # 保存到 word/media/
        img_filename = f'formula_{hash(latex) & 0xFFFFFFFF:08x}.png'
        media_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'template', 'word', 'media', img_filename)
        with open(media_path, 'wb') as f:
            f.write(img_bytes)

        if media_files_dict is not None:
            media_files_dict[f'word/media/{img_filename}'] = f'word/media/{img_filename}'

        # 公式图片按实际像素比例显示（1200 DPI 渲染）
        # EMU = pixels / DPI * 72 * 9525
        w_px = int.from_bytes(img_bytes[16:20], 'big')
        h_px = int.from_bytes(img_bytes[20:24], 'big')
        cx = round(w_px / 1200 * 72 * 9525)
        cy = round(h_px / 1200 * 72 * 9525)

        # A4 正文区最大宽度约 14.7cm = 5274315 EMU，超限时等比缩放
        MAX_CX = 5274315
        if cx > MAX_CX:
            scale = MAX_CX / cx
            cx = MAX_CX
            cy = round(cy * scale)

        r_id = None
        if media_files_dict is not None and '_rels_entries' in rels_counter:
            rels_counter['count'] += 1
            r_id = f'rIdImage{rels_counter["count"]}'
            rels_counter['_rels_entries'].append((r_id, f'media/{img_filename}'))
        else:
            r_id = 'rIdImage1'

        # 构建段落：公式图片居中（\tag{} 编号已在 LaTeX 图片中渲染）
        PIC_NS = 'http://schemas.openxmlformats.org/drawingml/2006/picture'
        WP_NS = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing'
        A_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main'

        p = ET.Element(f'{{{W}}}p')
        ppr = ET.SubElement(p, f'{{{W}}}pPr')
        p_adj = ET.SubElement(ppr, f'{{{W}}}adjustRightInd')
        p_adj.set(f'{{{W}}}val', '0')
        p_snap = ET.SubElement(ppr, f'{{{W}}}snapToGrid')
        p_snap.set(f'{{{W}}}val', '0')
        ET.SubElement(ppr, f'{{{W}}}keepNext')
        ET.SubElement(ppr, f'{{{W}}}keepLines')
        p_jc = ET.SubElement(ppr, f'{{{W}}}jc')
        p_jc.set(f'{{{W}}}val', 'center')

        # 公式图片 run
        r = ET.SubElement(p, f'{{{W}}}r')
        rpr = ET.SubElement(r, f'{{{W}}}rPr')
        ET.SubElement(rpr, f'{{{W}}}noProof')

        drawing = ET.SubElement(r, f'{{{W}}}drawing')
        inline = ET.SubElement(drawing, f'{{{WP_NS}}}inline')
        ET.SubElement(inline, f'{{{WP_NS}}}extent', attrib={'cx': str(cx), 'cy': str(cy)})
        ET.SubElement(inline, f'{{{WP_NS}}}effectExtent', attrib={'l': '0', 't': '0', 'r': '0', 'b': '0'})
        ET.SubElement(inline, f'{{{WP_NS}}}docPr', attrib={'id': str(next_drawing_id()), 'name': 'Picture'})
        cNvGFP = ET.SubElement(inline, f'{{{PIC_NS}}}cNvGraphicFramePr')
        graphic = ET.SubElement(inline, f'{{{A_NS}}}graphic')
        graphicData = ET.SubElement(graphic, f'{{{A_NS}}}graphicData',
                                    attrib={'uri': 'http://schemas.openxmlformats.org/drawingml/2006/picture'})
        pic = ET.SubElement(graphicData, f'{{{PIC_NS}}}pic')
        nvPicPr = ET.SubElement(pic, f'{{{PIC_NS}}}nvPicPr')
        ET.SubElement(nvPicPr, f'{{{PIC_NS}}}cNvPr', attrib={'id': '0', 'name': 'Picture'})
        ET.SubElement(nvPicPr, f'{{{PIC_NS}}}cNvPicPr')
        blipFill = ET.SubElement(pic, f'{{{PIC_NS}}}blipFill')
        ET.SubElement(blipFill, f'{{{A_NS}}}blip',
                      attrib={f'{{{R_NS}}}embed': r_id})
        stretch = ET.SubElement(blipFill, f'{{{A_NS}}}stretch')
        ET.SubElement(stretch, f'{{{A_NS}}}fillRect')
        spPr = ET.SubElement(pic, f'{{{PIC_NS}}}spPr')
        xfrm = ET.SubElement(spPr, f'{{{A_NS}}}xfrm')
        ET.SubElement(xfrm, f'{{{A_NS}}}off', attrib={'x': '0', 'y': '0'})
        ET.SubElement(xfrm, f'{{{A_NS}}}ext', attrib={'cx': str(cx), 'cy': str(cy)})
        prstGeom = ET.SubElement(spPr, f'{{{A_NS}}}prstGeom')
        prstGeom.set('prst', 'rect')
        ET.SubElement(prstGeom, f'{{{A_NS}}}avLst')

        return [p]

    except Exception as e:
        print(f'  警告: LaTeX 公式渲染失败: {e}', file=sys.stderr)
        # 回退：居中显示 LaTeX 文本
        p = ET.Element(f'{{{W}}}p')
        ppr = ET.SubElement(p, f'{{{W}}}pPr')
        pstyle = ET.SubElement(ppr, f'{{{W}}}pStyle')
        pstyle.set(f'{{{W}}}val', 'afa')
        jc = ET.SubElement(ppr, f'{{{W}}}jc')
        jc.set(f'{{{W}}}val', 'center')
        adj = ET.SubElement(ppr, f'{{{W}}}adjustRightInd')
        adj.set(f'{{{W}}}val', '0')
        snap = ET.SubElement(ppr, f'{{{W}}}snapToGrid')
        snap.set(f'{{{W}}}val', '0')
        ind = ET.SubElement(ppr, f'{{{W}}}ind')
        ind.set(f'{{{W}}}firstLineChars', '0')
        ind.set(f'{{{W}}}firstLine', '0')
        r = make_text_run(latex, font_ascii='Cambria Math', font_east_asia='Cambria Math', italic=True)
        p.append(r)
        return [p]


def build_image_placeholder(block):
    """图片下载失败时的占位段落。"""
    alt = block.get('alt', '')
    url = block.get('url', '')
    p = ET.SubElement(ET.Element('dummy'), f'{{{W}}}p')
    ppr = ET.SubElement(p, f'{{{W}}}pPr')
    pstyle = ET.SubElement(ppr, f'{{{W}}}pStyle')
    pstyle.set(f'{{{W}}}val', 'a4')
    r = ET.SubElement(p, f'{{{W}}}r')
    rt = ET.SubElement(r, f'{{{W}}}t')
    rt.text = f'[图片: {alt}]' if alt else '[图片]'
    return p


def build_image_block(block, media_files_dict=None, rels_counter=None):
    """下载图片并嵌入 docx。返回段落列表。"""
    url = block.get('url', '').strip()
    alt = block.get('alt', '')

    if not url:
        return [build_image_placeholder(block)]

    try:
        # 下载图片（优先本地文件 → 缓存 → URL直连 → 代理重试）
        img_filename_prefix = f'image_{hash(url) & 0xFFFFFFFF:08x}'
        media_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'template', 'word', 'media')
        cached = [f for f in os.listdir(media_dir) if f.startswith(img_filename_prefix + '.')] if os.path.isdir(media_dir) else []
        if cached:
            cached_path = os.path.join(media_dir, cached[0])
            with open(cached_path, 'rb') as _f:
                img_bytes = _f.read()
            content_type = 'image/png' if cached[0].endswith('.png') else 'image/jpeg'
        else:
            # 本地文件回退：支持相对路径（相对项目根）和 file:// URI
            local_path = None
            if url.startswith('file://'):
                local_path = url[7:]
            elif not url.startswith(('http://', 'https://')):
                project_root = os.path.dirname(os.path.abspath(__file__))
                local_path = os.path.normpath(os.path.join(project_root, url))
            if local_path and os.path.isfile(local_path):
                with open(local_path, 'rb') as _f:
                    img_bytes = _f.read()
                ext = os.path.splitext(local_path)[1].lower().lstrip('.')
                content_type = {'png': 'image/png', 'gif': 'image/gif', 'bmp': 'image/bmp'}.get(ext, 'image/jpeg')
            else:
                req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
                response = None
                # 先直连尝试
                try:
                    _opener = urllib.request.build_opener()
                    response = _opener.open(req, timeout=15)
                except Exception:
                    pass
                # 直连失败，挂代理重试
                if response is None:
                    proxy_url = os.environ.get('https_proxy') or os.environ.get('HTTPS_PROXY')
                    if proxy_url:
                        _proxy_handler = urllib.request.ProxyHandler({'https': proxy_url, 'http': proxy_url})
                        _ssl_ctx = ssl.create_default_context()
                        _https_handler = urllib.request.HTTPSHandler(context=_ssl_ctx)
                        _opener = urllib.request.build_opener(_proxy_handler, _https_handler)
                    else:
                        _ssl_ctx = ssl.create_default_context()
                        _opener = urllib.request.build_opener(urllib.request.HTTPSHandler(context=_ssl_ctx))
                    req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
                    response = _opener.open(req, timeout=30)
                    print(f'  图片通过代理下载: {url[:60]}...', file=sys.stderr)
                img_bytes = response.read()
                content_type = response.headers.get('Content-Type', '')

        # 确定文件扩展名
        if 'png' in content_type:
            ext = 'png'
        elif 'gif' in content_type:
            ext = 'gif'
        elif 'bmp' in content_type:
            ext = 'bmp'
        else:
            ext = 'jpeg'

        img_filename = f'image_{hash(url) & 0xFFFFFFFF:08x}.{ext}'
        media_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'template', 'word', 'media', img_filename)
        with open(media_path, 'wb') as f:
            f.write(img_bytes)

        if media_files_dict is not None:
            media_files_dict[f'word/media/{img_filename}'] = f'word/media/{img_filename}'

        # 统一宽度 15cm，保持宽高比
        r_id, cx, cy = _make_image_rel(media_files_dict, rels_counter, img_filename, img_bytes,
                                        15 * 360000 / 9525, 10 * 360000 / 9525)

        p = _build_image_paragraph(r_id, cx, cy)
        result = [p]

        # 添加图注段落（style=a4, jc=center, rFonts={ascii:'黑体', hAnsi:'黑体'}, szCs=21 五号）
        if alt:
            p_cap = ET.Element(f'{{{W}}}p')
            p_cap_ppr = ET.SubElement(p_cap, f'{{{W}}}pPr')
            p_cap_style = ET.SubElement(p_cap_ppr, f'{{{W}}}pStyle')
            p_cap_style.set(f'{{{W}}}val', 'a4')
            p_cap_jc = ET.SubElement(p_cap_ppr, f'{{{W}}}jc')
            p_cap_jc.set(f'{{{W}}}val', 'center')
            p_cap_adj = ET.SubElement(p_cap_ppr, f'{{{W}}}adjustRightInd')
            p_cap_adj.set(f'{{{W}}}val', '0')
            p_cap_snap = ET.SubElement(p_cap_ppr, f'{{{W}}}snapToGrid')
            p_cap_snap.set(f'{{{W}}}val', '0')
            p_cap_spacing = ET.SubElement(p_cap_ppr, f'{{{W}}}spacing')
            p_cap_spacing.set(f'{{{W}}}line', '360')
            p_cap_spacing.set(f'{{{W}}}lineRule', 'auto')
            cap_run = ET.SubElement(p_cap, f'{{{W}}}r')
            cap_rpr = ET.SubElement(cap_run, f'{{{W}}}rPr')
            cap_fonts = ET.SubElement(cap_rpr, f'{{{W}}}rFonts')
            cap_fonts.set(f'{{{W}}}asciiTheme', 'minorHAnsi')
            cap_fonts.set(f'{{{W}}}hAnsiTheme', 'minorHAnsi')
            cap_fonts.set(f'{{{W}}}cstheme', 'minorHAnsi')
            cap_fonts.set(f'{{{W}}}eastAsia', '黑体')
            cap_sz = ET.SubElement(cap_rpr, f'{{{W}}}sz')
            cap_sz.set(f'{{{W}}}val', '21')
            cap_scs = ET.SubElement(cap_rpr, f'{{{W}}}szCs')
            cap_scs.set(f'{{{W}}}val', '21')
            cap_t = ET.SubElement(cap_run, f'{{{W}}}t')
            cap_t.text = alt
            cap_t.set(f'{XML_NS}space', 'preserve')
            result.append(p_cap)

        return result

    except Exception as e:
        print(f'  警告: 图片下载失败 ({url}): {e}', file=sys.stderr)
        return [build_image_placeholder(block)]


def _make_image_rel(media_files_dict, rels_counter, img_filename, img_bytes, default_cx, default_cy, no_scale=False):
    """为图片创建关系 ID 并计算尺寸。返回 (r_id, cx_emu, cy_emu)。
    no_scale=True 时直接使用 default_cx/default_cy（单位 pt），不做 max_w_emu 缩放。"""
    if rels_counter is not None:
        rels_counter['count'] += 1
        r_id = f'rIdImage{rels_counter["count"]}'
    else:
        r_id = 'rIdImage1'

    # 记录 r_id 和文件名的映射到 rels_entries（Target 相对于 word/_rels/ 目录）
    if rels_counter is not None and '_rels_entries' in rels_counter:
        rels_counter['_rels_entries'].append((r_id, f'media/{img_filename}'))

    # 统一宽度 15cm，保持宽高比
    target_cx = 15 * 360000  # 15cm = 5,400,000 EMU
    cx, cy = default_cx * 9525, default_cy * 9525  # pt → EMU (fallback)
    try:
        if img_bytes[:8] == b'\x89PNG\r\n\x1a\n':
            # PNG
            w = int.from_bytes(img_bytes[16:20], 'big')
            h = int.from_bytes(img_bytes[20:24], 'big')
            if w > 0 and h > 0:
                aspect = h / w
                cx = target_cx
                cy = int(cx * aspect)
        elif img_bytes[:2] == b'\xff\xd8':
            # JPEG
            i = 2
            while i < len(img_bytes) - 1:
                if img_bytes[i] != 0xFF:
                    break
                marker = img_bytes[i + 1]
                if marker == 0xC0 or marker == 0xC2:
                    h = int.from_bytes(img_bytes[i + 5:i + 7], 'big')
                    w = int.from_bytes(img_bytes[i + 7:i + 9], 'big')
                    if w > 0 and h > 0:
                        aspect = h / w
                        cx = target_cx
                        cy = int(cx * aspect)
                    break
                elif marker == 0xD9 or marker == 0xDA:
                    break
                else:
                    length = int.from_bytes(img_bytes[i + 2:i + 4], 'big')
                    i += 2 + length
    except Exception:
        pass

    return r_id, cx, cy


def _build_image_paragraph(r_id, cx, cy):
    """构建包含嵌入式图片的居中段落。"""
    PIC_NS = 'http://schemas.openxmlformats.org/drawingml/2006/picture'

    p = ET.Element(f'{{{W}}}p')
    ppr = ET.SubElement(p, f'{{{W}}}pPr')
    jc = ET.SubElement(ppr, f'{{{W}}}jc')
    jc.set(f'{{{W}}}val', 'center')
    adj = ET.SubElement(ppr, f'{{{W}}}adjustRightInd')
    adj.set(f'{{{W}}}val', '0')
    snap = ET.SubElement(ppr, f'{{{W}}}snapToGrid')
    snap.set(f'{{{W}}}val', '0')
    ET.SubElement(ppr, f'{{{W}}}keepNext')
    ET.SubElement(ppr, f'{{{W}}}keepLines')

    r = ET.SubElement(p, f'{{{W}}}r')
    rpr = ET.SubElement(r, f'{{{W}}}rPr')
    ET.SubElement(rpr, f'{{{W}}}noProof')

    drawing = ET.SubElement(r, f'{{{W}}}drawing')

    # wp:inline
    WP_NS = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing'
    A_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main'

    inline = ET.SubElement(drawing, f'{{{WP_NS}}}inline')
    ET.SubElement(inline, f'{{{WP_NS}}}extent', attrib={'cx': str(cx), 'cy': str(cy)})
    ET.SubElement(inline, f'{{{WP_NS}}}effectExtent', attrib={'l': '0', 't': '0', 'r': '0', 'b': '0'})
    docPr = ET.SubElement(inline, f'{{{WP_NS}}}docPr', attrib={'id': str(next_drawing_id()), 'name': 'Picture'})

    # cNvGraphicFramePr
    cNvGFP = ET.SubElement(inline, f'{{{PIC_NS}}}cNvGraphicFramePr')

    # graphic
    graphic = ET.SubElement(inline, f'{{{A_NS}}}graphic')
    graphicData = ET.SubElement(graphic, f'{{{A_NS}}}graphicData',
                                attrib={'uri': 'http://schemas.openxmlformats.org/drawingml/2006/picture'})

    # pic:pic
    pic = ET.SubElement(graphicData, f'{{{PIC_NS}}}pic')
    nvPicPr = ET.SubElement(pic, f'{{{PIC_NS}}}nvPicPr')
    cNvPr = ET.SubElement(nvPicPr, f'{{{PIC_NS}}}cNvPr', attrib={'id': '0', 'name': 'Picture'})
    cNvPicPr = ET.SubElement(nvPicPr, f'{{{PIC_NS}}}cNvPicPr')

    blipFill = ET.SubElement(pic, f'{{{PIC_NS}}}blipFill')
    blip = ET.SubElement(blipFill, f'{{{A_NS}}}blip',
                          attrib={f'{{{R_NS}}}embed': r_id})
    stretch = ET.SubElement(blipFill, f'{{{A_NS}}}stretch')
    ET.SubElement(stretch, f'{{{A_NS}}}fillRect')

    spPr = ET.SubElement(pic, f'{{{PIC_NS}}}spPr')
    xfrm = ET.SubElement(spPr, f'{{{A_NS}}}xfrm')
    ET.SubElement(xfrm, f'{{{A_NS}}}off', attrib={'x': '0', 'y': '0'})
    ET.SubElement(xfrm, f'{{{A_NS}}}ext', attrib={'cx': str(cx), 'cy': str(cy)})
    prstGeom = ET.SubElement(spPr, f'{{{A_NS}}}prstGeom', attrib={'prst': 'rect'})
    ET.SubElement(prstGeom, f'{{{A_NS}}}avLst')

    return p


def _split_tall_image_snake(img_bytes, max_h_cm=20):
    """将过高的图片按页面高度切成多段，蛇形排列到 Word 表格中。
    返回 (strip_filenames, col_h_cm) 或 (None, original_h_cm)（无需分割时）。
    蛇形排列：多列并排，每段宽度 = 15cm / cols。
    """
    import struct as _struct

    is_png = img_bytes[:8] == b'\x89PNG\r\n\x1a\n'
    is_jpeg = img_bytes[:2] == b'\xff\xd8'

    w_px, h_px = 0, 0

    if is_png:
        w_px = _struct.unpack('>I', img_bytes[16:20])[0]
        h_px = _struct.unpack('>I', img_bytes[20:24])[0]
    elif is_jpeg:
        i = 2
        while i < len(img_bytes) - 1:
            if img_bytes[i] != 0xFF:
                break
            marker = img_bytes[i + 1]
            if marker == 0xC0 or marker == 0xC2:
                h_px = _struct.unpack('>H', img_bytes[i + 5:i + 7])[0]
                w_px = _struct.unpack('>H', img_bytes[i + 7:i + 9])[0]
                break
            elif marker == 0xD9 or marker == 0xDA:
                break
            else:
                length = _struct.unpack('>H', img_bytes[i + 2:i + 4])[0]
                i += 2 + length
    else:
        return None, 0

    if w_px == 0 or h_px == 0:
        return None, 0

    aspect = h_px / w_px
    img_h_cm = 15 * aspect

    if img_h_cm <= max_h_cm:
        return None, img_h_cm

    # 计算需要几列才能让每段高度 <= max_h_cm
    num_cols = int(img_h_cm / max_h_cm) + (1 if img_h_cm % max_h_cm > 0 else 0)
    num_cols = max(2, min(num_cols, 3))

    strip_h_px = h_px // num_cols
    if strip_h_px == 0:
        return None, img_h_cm

    col_w_cm = 15.0 / num_cols
    col_h_cm = col_w_cm * strip_h_px / w_px

    # 用 ImageMagick convert 切割
    tmpdir = tempfile.mkdtemp()
    try:
        ext = 'png' if is_png else 'jpg'
        src_path = os.path.join(tmpdir, f'src.{ext}')
        with open(src_path, 'wb') as f:
            f.write(img_bytes)

        strips = []
        for i in range(num_cols):
            offset_y = i * strip_h_px
            actual_h = strip_h_px if i < num_cols - 1 else (h_px - offset_y)
            strip_path = os.path.join(tmpdir, f'strip_{i}.png')
            result = subprocess.run(
                ['convert', src_path, '-crop', f'{w_px}x{actual_h}+0+{offset_y}', '+repage', strip_path],
                capture_output=True, timeout=30
            )
            if not os.path.exists(strip_path):
                return None, img_h_cm
            with open(strip_path, 'rb') as f:
                strips.append(f.read())

        # 保存条带图片到 media 目录
        media_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'template', 'word', 'media')
        os.makedirs(media_dir, exist_ok=True)

        strip_filenames = []
        for i, strip_data in enumerate(strips):
            fname = f'snake_{hash(img_bytes) & 0xFFFFFFFF:08x}_{i}.png'
            fpath = os.path.join(media_dir, fname)
            with open(fpath, 'wb') as f:
                f.write(strip_data)
            strip_filenames.append(fname)

        return strip_filenames, col_h_cm

    finally:
        shutil.rmtree(tmpdir, ignore_errors=True)


def _build_snake_table(strip_filenames, col_w_emu, col_h_emu, media_files_dict, rels_counter):
    """构建蛇形排列的 Word 表格：单行多列，每列一个图片条带。"""
    num_cols = len(strip_filenames)

    tbl = ET.Element(f'{{{W}}}tbl')

    # 表格属性
    tblPr = ET.SubElement(tbl, f'{{{W}}}tblPr')
    tblW = ET.SubElement(tblPr, f'{{{W}}}tblW')
    tblW.set(f'{{{W}}}w', '0')
    tblW.set(f'{{{W}}}type', 'auto')
    jc = ET.SubElement(tblPr, f'{{{W}}}jc')
    jc.set(f'{{{W}}}val', 'center')
    tblBorders = ET.SubElement(tblPr, f'{{{W}}}tblBorders')
    for border_name in ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']:
        b = ET.SubElement(tblBorders, f'{{{W}}}{border_name}')
        b.set(f'{{{W}}}val', 'none')
        b.set(f'{{{W}}}sz', '0')
        b.set(f'{{{W}}}space', '0')
        b.set(f'{{{W}}}color', 'auto')
    tblLayout = ET.SubElement(tblPr, f'{{{W}}}tblLayout')
    tblLayout.set(f'{{{W}}}type', 'fixed')

    # 列定义
    tblGrid = ET.SubElement(tbl, f'{{{W}}}tblGrid')
    for _ in range(num_cols):
        gc = ET.SubElement(tblGrid, f'{{{W}}}gridCol')
        gc.set(f'{{{W}}}w', str(col_w_emu))

    # 单行：每列一个条带图片
    tr = ET.SubElement(tbl, f'{{{W}}}tr')
    for i, fname in enumerate(strip_filenames):
        tc = ET.SubElement(tr, f'{{{W}}}tc')
        tcPr = ET.SubElement(tc, f'{{{W}}}tcPr')
        tcW = ET.SubElement(tcPr, f'{{{W}}}tcW')
        tcW.set(f'{{{W}}}w', str(col_w_emu))
        tcW.set(f'{{{W}}}type', 'dxa')

        # 计算此条带的实际尺寸
        strip_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'template', 'word', 'media', fname)
        sw_px, sh_px = 0, 0
        try:
            with open(strip_path, 'rb') as f:
                data = f.read(24)
                if data[:8] == b'\x89PNG\r\n\x1a\n':
                    sw_px = int.from_bytes(data[16:20], 'big')
                    sh_px = int.from_bytes(data[20:24], 'big')
        except Exception:
            pass

        # 图片段落的居中
        p = ET.SubElement(tc, f'{{{W}}}p')
        ppr = ET.SubElement(p, f'{{{W}}}pPr')
        jc_p = ET.SubElement(ppr, f'{{{W}}}jc')
        jc_p.set(f'{{{W}}}val', 'center')

        if media_files_dict is not None:
            media_files_dict[f'word/media/{fname}'] = f'word/media/{fname}'

        if rels_counter is not None:
            rels_counter['count'] += 1
            r_id = f'rIdImage{rels_counter["count"]}'
            if '_rels_entries' in rels_counter:
                rels_counter['_rels_entries'].append((r_id, f'media/{fname}'))
        else:
            r_id = 'rIdImage1'

        # 计算条带在列宽下的显示尺寸
        if sw_px > 0 and sh_px > 0:
            cx = col_w_emu
            cy = int(col_w_emu * sh_px / sw_px)
        else:
            cx = col_w_emu
            cy = col_h_emu

        # 构建图片 run
        PIC_NS = 'http://schemas.openxmlformats.org/drawingml/2006/picture'
        WP_NS_I = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing'
        A_NS_I = 'http://schemas.openxmlformats.org/drawingml/2006/main'

        r = ET.SubElement(p, f'{{{W}}}r')
        rpr = ET.SubElement(r, f'{{{W}}}rPr')
        ET.SubElement(rpr, f'{{{W}}}noProof')

        drawing = ET.SubElement(r, f'{{{W}}}drawing')
        inline = ET.SubElement(drawing, f'{{{WP_NS_I}}}inline')
        ET.SubElement(inline, f'{{{WP_NS_I}}}extent', attrib={'cx': str(cx), 'cy': str(cy)})
        ET.SubElement(inline, f'{{{WP_NS_I}}}effectExtent', attrib={'l': '0', 't': '0', 'r': '0', 'b': '0'})
        ET.SubElement(inline, f'{{{WP_NS_I}}}docPr', attrib={'id': str(i + 1), 'name': f'SnakeStrip{i}'})
        cNvGFP = ET.SubElement(inline, f'{{{PIC_NS}}}cNvGraphicFramePr')
        graphic = ET.SubElement(inline, f'{{{A_NS_I}}}graphic')
        graphicData = ET.SubElement(graphic, f'{{{A_NS_I}}}graphicData',
                                    attrib={'uri': 'http://schemas.openxmlformats.org/drawingml/2006/picture'})
        pic = ET.SubElement(graphicData, f'{{{PIC_NS}}}pic')
        nvPicPr = ET.SubElement(pic, f'{{{PIC_NS}}}nvPicPr')
        ET.SubElement(nvPicPr, f'{{{PIC_NS}}}cNvPr', attrib={'id': '0', 'name': f'Strip{i}'})
        ET.SubElement(nvPicPr, f'{{{PIC_NS}}}cNvPicPr')
        blipFill = ET.SubElement(pic, f'{{{PIC_NS}}}blipFill')
        ET.SubElement(blipFill, f'{{{A_NS_I}}}blip',
                      attrib={f'{{{R_NS}}}embed': r_id})
        stretch = ET.SubElement(blipFill, f'{{{A_NS_I}}}stretch')
        ET.SubElement(stretch, f'{{{A_NS_I}}}fillRect')
        xfrm = ET.SubElement(pic, f'{{{PIC_NS}}}shapePr')
        xfrm_el = ET.SubElement(xfrm, f'{{{A_NS_I}}}xfrm')
        off = ET.SubElement(xfrm_el, f'{{{A_NS_I}}}off', attrib={'x': '0', 'y': '0'})
        ext = ET.SubElement(xfrm_el, f'{{{A_NS_I}}}ext', attrib={'cx': str(cx), 'cy': str(cy)})
        prstGeom = ET.SubElement(xfrm, f'{{{A_NS_I}}}prstGeom', attrib={'prst': 'rect'})
        ET.SubElement(prstGeom, f'{{{A_NS_I}}}avLst')

    return tbl
    """构建图片占位段落（显示图片说明文字，居中）。"""
    alt = block.get('alt', block.get('url', ''))
    p = ET.Element(f'{{{W}}}p')
    ppr = ET.SubElement(p, f'{{{W}}}pPr')
    pstyle = ET.SubElement(ppr, f'{{{W}}}pStyle')
    pstyle.set(f'{{{W}}}val', 'a4')
    add_text_to_paragraph(p, f'[{alt}]')
    return p


def _build_single_table(headers, rows, col_width, num_cols, include_header=True):
    """构建单个 Word 表格 XML 元素。"""
    tbl = ET.Element(f'{{{W}}}tbl')

    # 表格属性
    tblPr = ET.SubElement(tbl, f'{{{W}}}tblPr')
    tblW = ET.SubElement(tblPr, f'{{{W}}}tblW')
    tblW.set(f'{{{W}}}w', '0')
    tblW.set(f'{{{W}}}type', 'auto')
    jc = ET.SubElement(tblPr, f'{{{W}}}jc')
    jc.set(f'{{{W}}}val', 'center')
    tblBorders = ET.SubElement(tblPr, f'{{{W}}}tblBorders')
    for border_name in ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']:
        b = ET.SubElement(tblBorders, f'{{{W}}}{border_name}')
        b.set(f'{{{W}}}val', 'single')
        b.set('sz', '4')
        b.set('space', '0')
        b.set('color', 'auto')
    tblLayout = ET.SubElement(tblPr, f'{{{W}}}tblLayout')
    tblLayout.set(f'{{{W}}}type', 'fixed')

    # 列定义
    tblGrid = ET.SubElement(tbl, f'{{{W}}}tblGrid')
    for _ in range(num_cols):
        gc = ET.SubElement(tblGrid, f'{{{W}}}gridCol')
        gc.set(f'{{{W}}}w', str(col_width))

    # 构建行的辅助函数
    def _make_row(cells_data, is_header=False):
        tr = ET.SubElement(tbl, f'{{{W}}}tr')
        if is_header:
            trPr = ET.SubElement(tr, f'{{{W}}}trPr')
            ET.SubElement(trPr, f'{{{W}}}tblHeader')
        for cell_text in cells_data:
            tc = ET.SubElement(tr, f'{{{W}}}tc')
            tcPr = ET.SubElement(tc, f'{{{W}}}tcPr')
            tcW = ET.SubElement(tcPr, f'{{{W}}}tcW')
            tcW.set(f'{{{W}}}w', str(col_width))
            tcW.set(f'{{{W}}}type', 'dxa')
            vAlign = ET.SubElement(tcPr, f'{{{W}}}vAlign')
            vAlign.set(f'{{{W}}}val', 'center')

            p = ET.SubElement(tc, f'{{{W}}}p')
            ppr = ET.SubElement(p, f'{{{W}}}pPr')
            pstyle = ET.SubElement(ppr, f'{{{W}}}pStyle')
            pstyle.set(f'{{{W}}}val', 'afa')
            # 范文表格单元格居中
            jc_cell = ET.SubElement(ppr, f'{{{W}}}jc')
            jc_cell.set(f'{{{W}}}val', 'center')
            adj = ET.SubElement(ppr, f'{{{W}}}adjustRightInd')
            adj.set(f'{{{W}}}val', '0')
            snap = ET.SubElement(ppr, f'{{{W}}}snapToGrid')
            snap.set(f'{{{W}}}val', '0')
            ind = ET.SubElement(ppr, f'{{{W}}}ind')
            ind.set(f'{{{W}}}firstLineChars', '0')
            ind.set(f'{{{W}}}firstLine', '0')

            if is_header:
                # 表头，五号字，不加粗（范文表格表头无 bold）
                r = make_text_run(cell_text, font_east_asia='宋体', sz=21, skip_punct_norm=True)
                p.append(r)
            else:
                # 数据行，五号字
                add_text_to_paragraph(p, cell_text, font_east_asia='宋体', sz=21, skip_punct_norm=True)

    # 表头
    if include_header:
        _make_row(headers, is_header=True)
    # 数据行
    for row in rows:
        # 补齐列数
        while len(row) < num_cols:
            row.append('')
        _make_row(row)

    return tbl


def _build_continued_caption(table_label):
    """构建续表标题段落（如 '续表3-1'，右对齐，黑体，五号）。
    格式对照模板 document.xml index 114 的续表段落。"""
    p = ET.Element(f'{{{W}}}p')
    ppr = ET.SubElement(p, f'{{{W}}}pPr')
    spacing = ET.SubElement(ppr, f'{{{W}}}spacing')
    spacing.set(f'{{{W}}}line', '360')
    spacing.set(f'{{{W}}}lineRule', 'auto')
    jc = ET.SubElement(ppr, f'{{{W}}}jc')
    jc.set(f'{{{W}}}val', 'right')
    # 段落级默认 run 属性
    pPr_rPr = ET.SubElement(ppr, f'{{{W}}}rPr')
    pPr_fonts = ET.SubElement(pPr_rPr, f'{{{W}}}rFonts')
    pPr_fonts.set(f'{{{W}}}asciiTheme', 'minorHAnsi')
    pPr_fonts.set(f'{{{W}}}hAnsiTheme', 'minorHAnsi')
    pPr_fonts.set(f'{{{W}}}cstheme', 'minorHAnsi')
    pPr_fonts.set(f'{{{W}}}eastAsia', '黑体')
    pPr_sz = ET.SubElement(pPr_rPr, f'{{{W}}}sz')
    pPr_sz.set(f'{{{W}}}val', '21')
    pPr_szCs = ET.SubElement(pPr_rPr, f'{{{W}}}szCs')
    pPr_szCs.set(f'{{{W}}}val', '21')
    # 文本 run
    r = ET.SubElement(p, f'{{{W}}}r')
    rpr = ET.SubElement(r, f'{{{W}}}rPr')
    r_fonts = ET.SubElement(rpr, f'{{{W}}}rFonts')
    r_fonts.set(f'{{{W}}}asciiTheme', 'minorHAnsi')
    r_fonts.set(f'{{{W}}}hAnsiTheme', 'minorHAnsi')
    r_fonts.set(f'{{{W}}}cstheme', 'minorHAnsi')
    r_fonts.set(f'{{{W}}}eastAsia', '黑体')
    r_sz = ET.SubElement(rpr, f'{{{W}}}sz')
    r_sz.set(f'{{{W}}}val', '21')
    r_szCs = ET.SubElement(rpr, f'{{{W}}}szCs')
    r_szCs.set(f'{{{W}}}val', '21')
    t = ET.SubElement(r, f'{{{W}}}t')
    t.text = f'续表{table_label}'
    t.set(f'{XML_NS}space', 'preserve')
    return p


def build_table(block, caption_text=None):
    """构建 Word 表格 XML。支持大表格自动拆分并在续页添加续表标题。
    返回元素列表（可能是单个表格，或多个表格+续表标题）。

    拆分触发条件（满足任一）：
    1. 数据行数超过 SPLIT_THRESHOLD
    2. 数据行中包含 <!-- split --> 标记（手动指定拆分位置）
    """
    headers = block['headers']
    rows = block['rows']
    num_cols = len(headers)

    # 均分列宽
    total_width = 8312
    col_width = total_width // num_cols

    # 从表注中提取表序（如 "3-1"），用于生成续表标题
    table_label = None
    if caption_text:
        clean_cap = strip_markdown_formatting(caption_text)
        m = re.match(r'(?:续)?表(\d[\w-]*\d)\s', clean_cap)
        if m:
            table_label = m.group(1)

    # 检测手动拆分标记：行中包含 <!-- split -->
    split_indices = []
    clean_rows = []
    for idx, row in enumerate(rows):
        row_text = ' '.join(row)
        if '<!-- split -->' in row_text:
            split_indices.append(len(clean_rows))
        else:
            clean_rows.append(row)
    rows = clean_rows

    # 自动拆分阈值（数据行数，不含表头）
    SPLIT_THRESHOLD = 20

    # 如果没有手动拆分标记且行数未超过阈值，不拆分
    if not split_indices and len(rows) <= SPLIT_THRESHOLD:
        tbl = _build_single_table(headers, rows, col_width, num_cols, include_header=True)
        return [tbl]

    # 确定拆分点：合并手动标记和自动阈值
    split_points = set(split_indices)
    if len(rows) > SPLIT_THRESHOLD:
        for pos in range(SPLIT_THRESHOLD, len(rows), SPLIT_THRESHOLD):
            split_points.add(pos)

    # 构建拆分后的元素列表
    elements = []
    prev = 0
    for sp in sorted(split_points):
        if sp <= prev:
            continue
        if prev == 0:
            # 第一个表格：正常表头
            tbl = _build_single_table(headers, rows[prev:sp], col_width, num_cols, include_header=True)
            elements.append(tbl)
        else:
            # 续表：先插续表标题，再插带表头的子表格
            if table_label:
                elements.append(_build_continued_caption(table_label))
            tbl = _build_single_table(headers, rows[prev:sp], col_width, num_cols, include_header=True)
            elements.append(tbl)
        prev = sp

    # 最后一个 chunk
    if prev < len(rows):
        if prev == 0:
            tbl = _build_single_table(headers, rows[prev:], col_width, num_cols, include_header=True)
            elements.append(tbl)
        else:
            if table_label:
                elements.append(_build_continued_caption(table_label))
            tbl = _build_single_table(headers, rows[prev:], col_width, num_cols, include_header=True)
            elements.append(tbl)

    return elements


def build_hr_paragraph():
    """构建分隔线段落。"""
    p = ET.Element(f'{{{W}}}p')
    ppr = ET.SubElement(p, f'{{{W}}}pPr')
    pstyle = ET.SubElement(ppr, f'{{{W}}}pStyle')
    pstyle.set(f'{{{W}}}val', 'afa')
    adj = ET.SubElement(ppr, f'{{{W}}}adjustRightInd')
    adj.set(f'{{{W}}}val', '0')
    snap = ET.SubElement(ppr, f'{{{W}}}snapToGrid')
    snap.set(f'{{{W}}}val', '0')
    pBdr = ET.SubElement(ppr, f'{{{W}}}pBdr')
    bottom = ET.SubElement(pBdr, f'{{{W}}}bottom')
    bottom.set(f'{{{W}}}val', 'single')
    bottom.set('sz', '6')
    bottom.set('space', '1')
    bottom.set('color', 'auto')
    return p


def build_page_break_paragraph():
    """构建分页符段落。"""
    p = ET.Element(f'{{{W}}}p')
    r = ET.SubElement(p, f'{{{W}}}r')
    br = ET.SubElement(r, f'{{{W}}}br')
    br.set(f'{{{W}}}type', 'page')
    return p


# ============================================================
# 5. 目录生成
# ============================================================

_bookmark_counter = 0
_drawing_id_counter = 0


def next_drawing_id():
    global _drawing_id_counter
    _drawing_id_counter += 1
    return _drawing_id_counter


def next_bookmark_name():
    global _bookmark_counter
    _bookmark_counter += 1
    return f'_Toc{999999999 + _bookmark_counter}'


def build_toc_entry(number, title, toc_level, bookmark_name='_Toc1000000000'):
    """构建目录条目段落。"""
    p = ET.Element(f'{{{W}}}p')
    ppr = ET.SubElement(p, f'{{{W}}}pPr')

    if toc_level == 1:
        pstyle = ET.SubElement(ppr, f'{{{W}}}pStyle')
        pstyle.set(f'{{{W}}}val', 'TOC1')
        tabs = ET.SubElement(ppr, f'{{{W}}}tabs')
        tab2 = ET.SubElement(tabs, f'{{{W}}}tab')
        tab2.set(f'{{{W}}}val', 'right')
        tab2.set(f'{{{W}}}leader', 'dot')
        tab2.set(f'{{{W}}}pos', '8302')
        rpr_in_ppr = ET.SubElement(ppr, f'{{{W}}}rPr')
        fonts = ET.SubElement(rpr_in_ppr, f'{{{W}}}rFonts')
        fonts.set(f'{{{W}}}ascii', 'Calibri')
        fonts.set(f'{{{W}}}eastAsia', '黑体')
        no_proof = ET.SubElement(rpr_in_ppr, f'{{{W}}}noProof')
        sz = ET.SubElement(rpr_in_ppr, f'{{{W}}}sz')
        sz.set(f'{{{W}}}val', '28')
        szCs = ET.SubElement(rpr_in_ppr, f'{{{W}}}szCs')
        szCs.set(f'{{{W}}}val', '28')
    elif toc_level == 2:
        pstyle = ET.SubElement(ppr, f'{{{W}}}pStyle')
        pstyle.set(f'{{{W}}}val', 'TOC2')
        # 范文TOC2首行缩进480（2字符）
        ind2 = ET.SubElement(ppr, f'{{{W}}}ind')
        ind2.set(f'{{{W}}}firstLine', '480')
        tabs = ET.SubElement(ppr, f'{{{W}}}tabs')
        tab2 = ET.SubElement(tabs, f'{{{W}}}tab')
        tab2.set(f'{{{W}}}val', 'right')
        tab2.set(f'{{{W}}}leader', 'dot')
        tab2.set(f'{{{W}}}pos', '8302')
        rpr_in_ppr = ET.SubElement(ppr, f'{{{W}}}rPr')
        fonts = ET.SubElement(rpr_in_ppr, f'{{{W}}}rFonts')
        fonts.set(f'{{{W}}}ascii', 'Calibri')
        fonts.set(f'{{{W}}}eastAsia', '宋体')
        no_proof = ET.SubElement(rpr_in_ppr, f'{{{W}}}noProof')
        sz = ET.SubElement(rpr_in_ppr, f'{{{W}}}sz')
        sz.set(f'{{{W}}}val', '24')
    else:
        pstyle = ET.SubElement(ppr, f'{{{W}}}pStyle')
        pstyle.set(f'{{{W}}}val', 'TOC3')
        # 范文TOC3首行缩进960（4字符）
        ind3 = ET.SubElement(ppr, f'{{{W}}}ind')
        ind3.set(f'{{{W}}}firstLine', '960')
        tabs = ET.SubElement(ppr, f'{{{W}}}tabs')
        tab2 = ET.SubElement(tabs, f'{{{W}}}tab')
        tab2.set(f'{{{W}}}val', 'right')
        tab2.set(f'{{{W}}}leader', 'dot')
        tab2.set(f'{{{W}}}pos', '8302')
        rpr_in_ppr = ET.SubElement(ppr, f'{{{W}}}rPr')
        fonts = ET.SubElement(rpr_in_ppr, f'{{{W}}}rFonts')
        fonts.set(f'{{{W}}}ascii', 'Calibri')
        fonts.set(f'{{{W}}}eastAsia', '楷体')
        no_proof = ET.SubElement(rpr_in_ppr, f'{{{W}}}noProof')
        sz = ET.SubElement(rpr_in_ppr, f'{{{W}}}sz')
        sz.set(f'{{{W}}}val', '24')

    def _r(text_content=None):
        """创建一个 run。"""
        r = ET.SubElement(p, f'{{{W}}}r')
        rpr = ET.SubElement(r, f'{{{W}}}rPr')
        fonts = ET.SubElement(rpr, f'{{{W}}}rFonts')
        fonts.set(f'{{{W}}}ascii', 'Calibri')
        fonts.set(f'{{{W}}}hAnsi', 'Calibri')
        no_proof = ET.SubElement(rpr, f'{{{W}}}noProof')
        if text_content is not None:
            t = ET.SubElement(r, f'{{{W}}}t')
            t.text = text_content
            t.set(f'{XML_NS}space', 'preserve')
        return r

    # 编号+标题合并为一个 run（目录条目格式：编号标题 + tab + 页码）
    if number:
        display_text = f'{number}{title}'
    else:
        display_text = title
    _r(display_text)

    # tab before page number
    tab_run = ET.SubElement(p, f'{{{W}}}r')
    rpr = ET.SubElement(tab_run, f'{{{W}}}rPr')
    fonts = ET.SubElement(rpr, f'{{{W}}}rFonts')
    fonts.set(f'{{{W}}}ascii', 'Calibri')
    fonts.set(f'{{{W}}}hAnsi', 'Calibri')
    no_proof = ET.SubElement(rpr, f'{{{W}}}noProof')
    ET.SubElement(tab_run, f'{{{W}}}tab')

    # PAGEREF field: begin
    r_begin = ET.SubElement(p, f'{{{W}}}r')
    rpr = ET.SubElement(r_begin, f'{{{W}}}rPr')
    fonts = ET.SubElement(rpr, f'{{{W}}}rFonts')
    fonts.set(f'{{{W}}}ascii', 'Calibri')
    fonts.set(f'{{{W}}}hAnsi', 'Calibri')
    no_proof = ET.SubElement(rpr, f'{{{W}}}noProof')
    fld_begin = ET.SubElement(r_begin, f'{{{W}}}fldChar')
    fld_begin.set(f'{{{W}}}fldCharType', 'begin')

    # PAGEREF instrText
    r_instr = ET.SubElement(p, f'{{{W}}}r')
    rpr = ET.SubElement(r_instr, f'{{{W}}}rPr')
    fonts = ET.SubElement(rpr, f'{{{W}}}rFonts')
    fonts.set(f'{{{W}}}ascii', 'Calibri')
    fonts.set(f'{{{W}}}hAnsi', 'Calibri')
    no_proof = ET.SubElement(rpr, f'{{{W}}}noProof')
    instr = ET.SubElement(r_instr, f'{{{W}}}instrText')
    instr.set(f'{XML_NS}space', 'preserve')
    instr.text = f' PAGEREF {bookmark_name} \\h '

    # empty run
    ET.SubElement(p, f'{{{W}}}r')

    # PAGEREF field: separate
    r_sep = ET.SubElement(p, f'{{{W}}}r')
    rpr = ET.SubElement(r_sep, f'{{{W}}}rPr')
    fonts = ET.SubElement(rpr, f'{{{W}}}rFonts')
    fonts.set(f'{{{W}}}ascii', 'Calibri')
    fonts.set(f'{{{W}}}hAnsi', 'Calibri')
    no_proof = ET.SubElement(rpr, f'{{{W}}}noProof')
    fld_sep = ET.SubElement(r_sep, f'{{{W}}}fldChar')
    fld_sep.set(f'{{{W}}}fldCharType', 'separate')

    # 页码占位
    r_page = ET.SubElement(p, f'{{{W}}}r')
    rpr = ET.SubElement(r_page, f'{{{W}}}rPr')
    fonts = ET.SubElement(rpr, f'{{{W}}}rFonts')
    fonts.set(f'{{{W}}}ascii', 'Calibri')
    fonts.set(f'{{{W}}}hAnsi', 'Calibri')
    no_proof = ET.SubElement(rpr, f'{{{W}}}noProof')
    t = ET.SubElement(r_page, f'{{{W}}}t')
    t.text = '1'

    # PAGEREF field: end
    r_end = ET.SubElement(p, f'{{{W}}}r')
    rpr = ET.SubElement(r_end, f'{{{W}}}rPr')
    fonts = ET.SubElement(rpr, f'{{{W}}}rFonts')
    fonts.set(f'{{{W}}}ascii', 'Calibri')
    fonts.set(f'{{{W}}}hAnsi', 'Calibri')
    no_proof = ET.SubElement(rpr, f'{{{W}}}noProof')
    fld_end = ET.SubElement(r_end, f'{{{W}}}fldChar')
    fld_end.set(f'{{{W}}}fldCharType', 'end')

    return p


# 无编号标题关键词：匹配到的 H1 标题不分配章节编号
_UNNUMBERED_HEADING_KEYWORDS = ('结论', '致谢', '参考文献', '附录', '附录A', '附录B', '附录C')


def _number_to_chinese(number_str, toc_level):
    """将阿拉伯数字编号转换为中文编号格式。
    toc_level 1: "1" → "一",  "2" → "二"
    toc_level 2: "1.1" → "（一）", "1.2" → "（二）"
    toc_level 3: "1.1.1" → "1.", "1.1.2" → "2."
    """
    CN_NUMS = '一二三四五六七八九十'
    cn_num_map = list(CN_NUMS) + \
                 ['十一','十二','十三','十四','十五','十六','十七','十八','十九','二十',
                  '二十一','二十二','二十三','二十四','二十五','二十六']

    if toc_level == 1:
        try:
            num = int(number_str)
            if 1 <= num <= len(cn_num_map):
                return f'{cn_num_map[num - 1]}、'
        except ValueError:
            pass
        return number_str
    elif toc_level == 2:
        # 从 "1.1" 中提取第二位数字
        parts = number_str.split('.')
        if len(parts) >= 2:
            try:
                num = int(parts[-1])
                if 1 <= num <= len(cn_num_map):
                    return f'（{cn_num_map[num - 1]}）'
            except ValueError:
                pass
        return number_str
    elif toc_level == 3:
        # 从 "1.1.1" 中提取第三位数字
        parts = number_str.split('.')
        if len(parts) >= 3:
            try:
                num = int(parts[-1])
                return f'{num}.'
            except ValueError:
                pass
        return number_str
    return number_str


def _is_reference_paragraph(text):
    """判断段落是否为参考文献条目（以 [N] 开头）。"""
    return bool(re.match(r'^\[(\d+)\]\s*', text))


def _fix_reference_format(text):
    """修正参考文献条目格式：移除 [N] 后的空格。"""
    return re.sub(r'^\[(\d+)\]\s+', r'[\1]', text)


def build_reference_paragraph(block):
    """构建参考文献段落。五号字（sz=21），无首行缩进，[N]后无空格。
    范文参考文献不使用样式，宋体+Calibri，左对齐。"""
    text = block['text']
    text = _fix_reference_format(text)
    p = ET.Element(f'{{{W}}}p')
    ppr = ET.SubElement(p, f'{{{W}}}pPr')
    # 范文参考文献无段落样式，左对齐
    jc = ET.SubElement(ppr, f'{{{W}}}jc')
    jc.set(f'{{{W}}}val', 'left')
    # 参考文献使用五号字（sz=21），宋体
    add_text_to_paragraph(p, text, font_ascii='Calibri', font_east_asia='宋体', sz=21)
    return p


def build_toc_entries(blocks):
    """从 Markdown blocks 中提取标题，生成目录条目列表。"""
    global _bookmark_counter
    entries = []
    h1_count = 0
    h2_count = 0
    h3_count = 0
    real_h1_count = 0  # 真正的 H1 标题数，不受裸数字提升影响
    in_unnumbered_section = False  # 标记是否进入了无编号后记部分

    for block in blocks:
        if block['type'] != 'heading':
            continue
        level = block['level']
        if level > 3:
            level = 3
        skip_toc = False

        clean_title_early = strip_markdown_formatting(block['text'])

        # 检测无编号后记标题（参考文献、致谢、附录、结论等），不分配编号
        is_unnumbered = any(clean_title_early == kw or clean_title_early.startswith(kw + ' ')
                            for kw in _UNNUMBERED_HEADING_KEYWORDS)

        if is_unnumbered:
            in_unnumbered_section = True
            number = ''
            toc_level = 1 if level == 1 else 2
            skip_toc = False  # 后记顶级标题（致谢/参考文献/附录）仍进目录
        elif in_unnumbered_section:
            # 无编号后记部分的子标题（如 附录A下的二级/三级标题），不进目录
            number = ''
            toc_level = 2
            skip_toc = True
            block['_plain_heading'] = True  # 正文用普通段落样式，不用标题样式
        elif level == 1:
            h1_count += 1
            real_h1_count += 1
            h2_count = 0
            h3_count = 0
            number = str(h1_count)
            toc_level = 1
        elif level == 2:
            clean_title_early = strip_markdown_formatting(block['text'])
            # 检测裸数字标题（如 "4 多Agent辩论式协作框架设计"）
            # 排除通用标题如 "7 本章小结"
            bare_num_match = re.match(r'^(\d+)\s+\S', clean_title_early)
            is_generic = re.match(r'^\d+\s+(本章小结|小结|引言|结论)$', clean_title_early)
            if bare_num_match and not is_generic:
                bare_num = int(bare_num_match.group(1))
                # 仅当裸数字紧接在最后一个真正 H1 之后（= real_h1_count + 1）才提升
                # 这处理了多个章节合并为一个时第一个降级标题为新章的情况
                if bare_num == real_h1_count + 1:
                    h1_count = bare_num
                    # 不更新 real_h1_count：它只追踪真正的 H1 标题
                    h2_count = 0
                    h3_count = 0
                    number = str(h1_count)
                    toc_level = 1
                    block['_force_h1'] = True
                else:
                    # 作为当前章的子节，重新编号为 h1_count.N
                    h2_count += 1
                    h3_count = 0
                    number = f'{h1_count}.{h2_count}'
                    toc_level = 2
            else:
                h2_count += 1
                h3_count = 0
                number = f'{h1_count}.{h2_count}'
                toc_level = 2
        else:
            # H3 标题：若 h2_count==0（紧跟被提升的 H1），视为 H2 级别
            if h2_count == 0:
                h2_count += 1
                number = f'{h1_count}.{h2_count}'
                toc_level = 2
            else:
                h3_count += 1
                number = f'{h1_count}.{h2_count}.{h3_count}'
                toc_level = 3

        clean_title = strip_markdown_formatting(block['text'])
        # 从标题中分离编号和标题文字（标题可能已包含编号如 "1 绪论"）
        title_without_number = clean_title
        # 尝试匹配任意深度的编号前缀：如 "1"、"4.1"、"4.1.1"
        m = re.match(r'^(\d+(?:\.\d+)*)\s+(.+)$', clean_title)
        if m:
            title_without_number = m.group(2)

        bookmark = next_bookmark_name()
        block['_bookmark'] = bookmark
        block['_bookmark_id'] = _bookmark_counter
        # 将编号转换为中文格式
        display_number = _number_to_chinese(number, toc_level) if number else ''
        if not skip_toc:
            entries.append({
                'number': display_number,
                'title': title_without_number,
                'toc_level': toc_level,
                'bookmark': bookmark,
            })

    return entries


# ============================================================
# 6. Bookmark 管理
# ============================================================

def find_max_bookmark_id(body):
    """扫描 body 中所有 bookmarkStart 的最大 id。"""
    max_id = -1
    for bm in body.iter(f'{{{W}}}bookmarkStart'):
        bid = bm.get(f'{{{W}}}id')
        if bid is not None:
            try:
                max_id = max(max_id, int(bid))
            except ValueError:
                pass
    return max_id


def _remove_callout_annotations(body):
    """移除模板中所有圆角矩形标注框（wedgeRoundRectCallout）。
    这些是模板自带的红色格式说明批注，不应出现在最终文档中。"""
    MC = 'http://schemas.openxmlformats.org/markup-compatibility/2006'
    A_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main'
    WPS = 'http://schemas.microsoft.com/office/word/2010/wordprocessingShape'
    removed = 0

    # 收集需要删除的 AlternateContent 元素（不能在迭代中修改树）
    to_remove = []
    for ac in body.iter(f'{{{MC}}}AlternateContent'):
        has_callout = False
        for pg in ac.iter(f'{{{A_NS}}}prstGeom'):
            if pg.get('prst') == 'wedgeRoundRectCallout':
                has_callout = True
                break
        if has_callout:
            to_remove.append(ac)

    for ac in to_remove:
        parent_map = {c: p for p in body.iter() for c in p}
        parent = parent_map.get(ac)
        if parent is not None:
            parent.remove(ac)
            removed += 1

    if removed:
        print(f'已清除 {removed} 个圆角矩形标注框')


def remove_old_bookmarks(body):
    """移除整个 body 中所有 _Toc bookmarkStart 及配对的 bookmarkEnd。"""
    parent_map = {c: p for p in body.iter() for c in p}

    for bm in list(body.iter(f'{{{W}}}bookmarkStart')):
        name = bm.get(f'{{{W}}}name', '')
        if name.startswith('_Toc'):
            parent = parent_map.get(bm)
            if parent is not None:
                parent.remove(bm)

    parent_map = {c: p for p in body.iter() for c in p}

    for bm in list(body.iter(f'{{{W}}}bookmarkEnd')):
        bid = bm.get(f'{{{W}}}id', '')
        has_start = any(b.get(f'{{{W}}}id') == bid for b in body.iter(f'{{{W}}}bookmarkStart'))
        if not has_start:
            parent = parent_map.get(bm)
            if parent is not None:
                parent.remove(bm)


# ============================================================
# 7. 封面注入
# ============================================================

def _find_label_end_run(paragraph, label_patterns):
    """
    在段落中找到标签结束的位置。
    label_patterns: 标签文字列表（如 ['题    目', '学    号']）
    返回标签结束后的第一个 run 索引。
    """
    runs = paragraph.findall(f'{{{W}}}r')
    accumulated = ''
    for i, r in enumerate(runs):
        t = r.find(f'{{{W}}}t')
        if t is not None and t.text:
            accumulated += t.text
        # 检查累积文本是否匹配某个标签
        for pat in label_patterns:
            if pat in accumulated:
                return i + 1
    return 1  # 默认跳过第一个 run


def _make_cover_run(text, is_label=False, sz_value=None, underline=True):
    """创建封面字段 run，严格对照2026届新模板格式。
    模板 P4-P13 的 run 格式（统一使用 theme font）：
      - 所有 run：rFonts={asciiTheme:'minorHAnsi', hAnsiTheme:'minorHAnsi', cstheme:'minorHAnsi'}
      - 标签 run：b={val:'0'}, sz={val:'44'}, 无下划线
      - 值 run：b={val:'0'}, sz 按字段不同, u={val:'single'}
    is_label=True: 标签 run，无下划线
    is_label=False: 值 run，有下划线
    sz_value: 字号覆盖（半点），默认 44
    """
    r = ET.Element(f'{{{W}}}r')
    rpr = ET.SubElement(r, f'{{{W}}}rPr')
    fonts = ET.SubElement(rpr, f'{{{W}}}rFonts')
    fonts.set(f'{{{W}}}asciiTheme', 'minorHAnsi')
    fonts.set(f'{{{W}}}hAnsiTheme', 'minorHAnsi')
    fonts.set(f'{{{W}}}cstheme', 'minorHAnsi')
    b = ET.SubElement(rpr, f'{{{W}}}b')
    b.set(f'{{{W}}}val', '0')

    s = ET.SubElement(rpr, f'{{{W}}}sz')
    s.set(f'{{{W}}}val', str(sz_value or 44))

    if not is_label and underline:
        u = ET.SubElement(rpr, f'{{{W}}}u')
        u.set(f'{{{W}}}val', 'single')

    scs = ET.SubElement(rpr, f'{{{W}}}szCs')
    scs.set(f'{{{W}}}val', '44')
    t = ET.SubElement(r, f'{{{W}}}t')
    t.text = text
    t.set(f'{XML_NS}space', 'preserve')
    return r


def _display_width(text):
    """计算文本的显示宽度（中文字符占2，ASCII占1）。"""
    width = 0
    for ch in text:
        if '一' <= ch <= '鿿' or '　' <= ch <= '〿' or '＀' <= ch <= '￯':
            width += 2
        else:
            width += 1
    return width


# 封面有效行容量（显示单位，sz=44基准）
# A4页面: w=11906 twips, margins left=1797 right=1797 → text area=8312 twips
# At sz=44, half-width char = 11pt = 220 twips
# P4-P5: ind left=3250 right=1050 → (8312-4300)/220 = 18
# P6-P13: ind left=1050 right=1050 → (8312-2100)/220 = 28
_COVER_CAPACITY_WIDE = 18   # P4-P5 (题目/英文题目，左侧缩进较大)
_COVER_CAPACITY_NORMAL = 28 # P6-P13 (其他字段)


def _make_cover_centered_value(value, label_display_width, sz_value=None, underline=True, capacity=_COVER_CAPACITY_NORMAL):
    """创建封面居中值：前导空格 + 值 + 尾随空格。
    空格始终使用 base_sz=44，值使用 sz_value（仅冒号后内容）。
    混合字号时按 twips 精确计算剩余宽度。"""
    base_sz = 44
    val_sz = sz_value or base_sz
    # 可用宽度（twips）：(capacity - label_dw) * 每单位twips（base_sz下220twips/单位）
    available_twips = (capacity - label_display_width) * (base_sz * 5)
    # 值宽度（twips）：display_width * val_sz下每单位twips
    val_twips = _display_width(value) * (val_sz * 5)
    # 空格每个占 base_sz 下 1 单位 = base_sz*5 twips
    space_unit_twips = base_sz * 5  # 220
    remaining_twips = max(0, available_twips - val_twips)
    total_spaces = remaining_twips // space_unit_twips
    # 前导空格向上取整，避免值偏右
    leading = (total_spaces + 1) // 2
    trailing = total_spaces - leading

    def _make_space_run(count, sz, under):
        r = ET.Element(f'{{{W}}}r')
        rpr = ET.SubElement(r, f'{{{W}}}rPr')
        fonts = ET.SubElement(rpr, f'{{{W}}}rFonts')
        fonts.set(f'{{{W}}}asciiTheme', 'minorHAnsi')
        fonts.set(f'{{{W}}}hAnsiTheme', 'minorHAnsi')
        fonts.set(f'{{{W}}}cstheme', 'minorHAnsi')
        b_el = ET.SubElement(rpr, f'{{{W}}}b')
        b_el.set(f'{{{W}}}val', '0')
        if under:
            u_el = ET.SubElement(rpr, f'{{{W}}}u')
            u_el.set(f'{{{W}}}val', 'single')
        s_el = ET.SubElement(rpr, f'{{{W}}}sz')
        s_el.set(f'{{{W}}}val', str(sz))
        scs = ET.SubElement(rpr, f'{{{W}}}szCs')
        scs.set(f'{{{W}}}val', str(sz))
        t_el = ET.SubElement(r, f'{{{W}}}t')
        t_el.text = ' ' * count
        t_el.set(f'{XML_NS}space', 'preserve')
        return r

    runs = []
    if leading > 0:
        runs.append(_make_space_run(leading, base_sz, underline))
    runs.append(_make_cover_run(value, sz_value=sz_value, underline=underline))
    if trailing > 0:
        runs.append(_make_space_run(trailing, base_sz, underline))
    return runs


def inject_cover_data(paragraphs, cover_data):
    """
    将 final_paper.md 中的封面信息注入到模板封面段落（P4-P13）。
    2026届新模板：所有封面字段使用 af2 (Title) 样式，无 ddList 下拉框。
    P5 英文题目和 P10 学部(院) 的字号保持模板原样不修改。
    """
    # 段落标签文字和对应字段
    field_map = {
        4: ('题目', '题    目：'),
        5: ('英文题目', '英文题目：'),
        6: ('学号', '学    号：'),
        7: ('姓名', '姓    名：'),
        8: ('班级', '班    级：'),
        9: ('专业', '专    业：'),
        10: ('学部(院)', '学 部(院)：'),
        11: ('入学时间', '入学时间：'),
        12: ('指导教师', '指导教师：'),
        13: ('日期', '日    期：'),
    }

    for p_idx, (key, label) in field_map.items():
        if key not in cover_data:
            continue
        value = cover_data[key]
        if p_idx >= len(paragraphs):
            continue

        p = paragraphs[p_idx]
        runs = p.findall(f'{{{W}}}r')

        # 移除所有非 pPr 子元素
        for child in list(p):
            if child.tag != f'{{{W}}}pPr':
                p.remove(child)

        # 标签（无下划线）
        p.append(_make_cover_run(label, is_label=True))

        # 自动计算标签显示宽度，用于值居中对齐
        label_dw = _display_width(label)

        # 值（居中，带前导/尾随空格和下划线，自动对齐行宽）
        if p_idx == 4:
            # P4 题目：值较长，左对齐 + 尾部下划线空格填充到行尾
            p.append(_make_cover_run(value))
            val_dw = _display_width(value)
            remaining = max(0, _COVER_CAPACITY_WIDE - label_dw - val_dw)
            if remaining > 0:
                # 尾部填充空格
                r_trail = ET.Element(f'{{{W}}}r')
                rpr = ET.SubElement(r_trail, f'{{{W}}}rPr')
                fonts = ET.SubElement(rpr, f'{{{W}}}rFonts')
                fonts.set(f'{{{W}}}asciiTheme', 'minorHAnsi')
                fonts.set(f'{{{W}}}hAnsiTheme', 'minorHAnsi')
                fonts.set(f'{{{W}}}cstheme', 'minorHAnsi')
                b_el = ET.SubElement(rpr, f'{{{W}}}b')
                b_el.set(f'{{{W}}}val', '0')
                u_el = ET.SubElement(rpr, f'{{{W}}}u')
                u_el.set(f'{{{W}}}val', 'single')
                s_el = ET.SubElement(rpr, f'{{{W}}}sz')
                s_el.set(f'{{{W}}}val', '44')
                scs = ET.SubElement(rpr, f'{{{W}}}szCs')
                scs.set(f'{{{W}}}val', '44')
                t_el = ET.SubElement(r_trail, f'{{{W}}}t')
                t_el.text = ' ' * remaining
                t_el.set(f'{XML_NS}space', 'preserve')
                p.append(r_trail)
        elif p_idx == 5:
            # P5 英文题目：值 sz=24（小四）
            for r in _make_cover_centered_value(value, label_dw, sz_value=24, capacity=_COVER_CAPACITY_WIDE):
                p.append(r)
        elif p_idx == 10:
            # P10 学部(院)：值 sz=40（20pt）
            for r in _make_cover_centered_value(value, label_dw, sz_value=40, capacity=_COVER_CAPACITY_NORMAL):
                p.append(r)
        else:
            # 其他字段：sz=44
            for r in _make_cover_centered_value(value, label_dw, capacity=_COVER_CAPACITY_NORMAL):
                p.append(r)


# ============================================================
# 8. 摘要替换
# ============================================================

def _extract_title_from_blocks(blocks, chinese=True):
    """从 blocks 中提取摘要标题（中文/英文）。"""
    for block in blocks:
        if block['type'] == 'heading':
            text = strip_markdown_formatting(block['text'])
            if chinese:
                # 中文摘要标题：居中的 H1 级标题（不含"摘要"二字）
                if re.match(r'^[\u4e00-\u9fff]', text) and len(text) > 4 and '摘要' not in text:
                    return text
            else:
                # 英文摘要标题：ABSTRACT 之前的标题（大写英文）
                if re.match(r'^[A-Z]', text) and 'ABSTRACT' not in text.upper():
                    return text
    return None


def replace_abstracts(paragraphs, blocks, parent_body=None, cover_data=None):
    """
    替换中英文摘要段落。
    2026届新模板结构（基于 document.xml 实际段落索引）：
      P31: 中文标题
      P32: "摘要"标签
      P33-P35: 中文摘要正文
      P36: 中文关键词
      P37: 关键词格式说明（需清空）
      P38: 空段落/格式说明（需清空，插入分页符分隔中英文摘要）
      P39: 英文标题（af2样式，TNR+宋体，sz=36）
      P40: "ABSTRACT"标签（af2样式，TNR+宋体）
      P41: 英文摘要正文（af2样式，TNR+宋体，b=0，sz=24）
      P42: 英文关键词（af2样式，TNR+宋体）
      P43: 分节符（摘要→目录，nextPage）
    """
    # 优先从封面数据提取标题，回退到 Markdown blocks
    cn_title = (cover_data or {}).get('题目') or _extract_title_from_blocks(blocks, chinese=True) or '毕业设计（论文）题目'
    en_title = (cover_data or {}).get('英文题目') or _extract_title_from_blocks(blocks, chinese=False) or 'THESIS TITLE'

    # === 清理 P31 格式说明文字（保留模板 pPr 及 pPr/rPr，清除 AlternateContent 和旧 runs） ===
    _clean_paragraph_keep_ppr(paragraphs[31], keep_rpr=True)
    # 重新添加标题 run（模板格式：居中，黑体，sz=36，theme fonts）
    title_run = make_text_run(cn_title, font_east_asia='黑体', sz=36, theme_fonts=True)
    paragraphs[31].append(title_run)

    # === 清理 P32 格式说明文字，只保留"摘要" ===
    _clean_paragraph_keep_ppr(paragraphs[32], keep_rpr=True)
    title_run2 = make_text_run('摘要', font_east_asia='黑体', bold=True,
                                sz=32, theme_fonts=True)
    paragraphs[32].append(title_run2)

    # === 清理 P37（关键词格式说明段落）— 清空内容 ===
    _clean_paragraph_keep_ppr(paragraphs[37], keep_rpr=True)

    # === 清理 P38（格式说明段落）— 插入分页符（中英文摘要之间） ===
    _clean_paragraph_keep_ppr(paragraphs[38], keep_rpr=True)
    # 新版模板在中文摘要末尾有分页符（P132: <w:br type="page"/>），中英文摘要分页显示
    br_run = ET.Element(f'{{{W}}}r')
    br_el = ET.SubElement(br_run, f'{{{W}}}br')
    br_el.set(f'{{{W}}}type', 'page')
    paragraphs[38].append(br_run)

    # === 清理 P39（英文标题）— af2样式，模板 pPr/rPr 已设 sz=36（小二号）===
    _clean_paragraph_keep_ppr(paragraphs[39], keep_rpr=True)
    _add_jc_to_paragraph(paragraphs[39], 'center')
    # 英文标题 run: sz=36 显式覆盖 af2 样式的 sz=32
    en_title_run = make_text_run(en_title, font_east_asia='宋体', sz=36, theme_fonts=True)
    paragraphs[39].append(en_title_run)

    # === 清理 P40（ABSTRACT 标签）— af2样式，模板 pPr/rPr 已设 bCs=0, szCs=32 ===
    _clean_paragraph_keep_ppr(paragraphs[40], keep_rpr=True)
    _add_jc_to_paragraph(paragraphs[40], 'center')
    # 范文 ABSTRACT run: 仅设 asciiTheme + eastAsia，无 sz、无 bold
    abstract_run = make_text_run('ABSTRACT', font_east_asia='宋体', theme_fonts=True)
    paragraphs[40].append(abstract_run)

    # === 清理 P42（英文关键词格式说明）===
    _clean_paragraph_keep_ppr(paragraphs[42], keep_rpr=True)

    # 找到中文摘要正文块（"摘要"标题之后，"ABSTRACT"标题之前的内容块）
    cn_abstract_blocks = []
    en_abstract_blocks = []
    in_cn_abstract = False
    in_en_abstract = False

    for block in blocks:
        if block['type'] == 'heading':
            text = strip_markdown_formatting(block['text'])
            if '摘要' in text and 'ABSTRACT' not in text.upper():
                in_cn_abstract = True
                in_en_abstract = False
                continue
            elif 'ABSTRACT' in text.upper():
                in_cn_abstract = False
                in_en_abstract = True
                continue
            else:
                in_cn_abstract = False
                in_en_abstract = False
                continue

        if block['type'] == 'hr':
            in_cn_abstract = False
            in_en_abstract = False
            continue

        if block['type'] in ('paragraph', 'bullet', 'numbered'):
            # 排除关键词段落（单独处理）
            text = block.get('text', '')
            if '关键词' in text or 'key words' in text.lower() or 'keywords' in text.lower():
                continue
            if in_cn_abstract:
                cn_abstract_blocks.append(block)
            elif in_en_abstract:
                en_abstract_blocks.append(block)

    # 替换中文摘要正文（P33-P35 → 使用摘要专用格式）
    _replace_abstract_body(paragraphs, 33, 35, cn_abstract_blocks, chinese=True,
                           parent_body=parent_body, insert_before_idx=36)

    # 清理未使用的摘要占位段落（P34/P35），移除缩进避免多余空白
    for i in range(33 + len(cn_abstract_blocks), 36):
        if i < len(paragraphs):
            p = paragraphs[i]
            # 检查段落是否为空（无文本内容）
            has_text = any(t.text and t.text.strip() for t in p.findall(f'.//{{{W}}}t'))
            if not has_text:
                ppr = p.find(f'{{{W}}}pPr')
                if ppr is not None:
                    ind = ppr.find(f'{{{W}}}ind')
                    if ind is not None:
                        ppr.remove(ind)

    # 替换英文摘要正文（P41 → 合并多段为一个段落，避免插入破坏后续索引）
    # 2026届新模板 P41 格式：style=af2, TNR+宋体, b=0, sz=24, jc=both, theme fonts
    if en_abstract_blocks:
        _clean_paragraph_keep_ppr(paragraphs[41], keep_rpr=True)
        _add_jc_to_paragraph(paragraphs[41], 'both')
        # 将所有英文摘要正文合并为一段
        merged_text = ' '.join(b['text'] for b in en_abstract_blocks)
        inline_runs = parse_inline(merged_text)
        for run_info in inline_runs:
            if not run_info['text']:
                continue
            r = ET.Element(f'{{{W}}}r')
            rpr = ET.SubElement(r, f'{{{W}}}rPr')
            fonts = ET.SubElement(rpr, f'{{{W}}}rFonts')
            fonts.set(f'{{{W}}}ascii', 'Calibri')
            fonts.set(f'{{{W}}}eastAsia', '宋体')
            fonts.set(f'{{{W}}}hAnsi', 'Calibri')
            fonts.set(f'{{{W}}}asciiTheme', 'minorHAnsi')
            fonts.set(f'{{{W}}}hAnsiTheme', 'minorHAnsi')
            fonts.set(f'{{{W}}}cstheme', 'minorHAnsi')
            b_el = ET.SubElement(rpr, f'{{{W}}}b')
            b_el.set(f'{{{W}}}val', '0')
            bcs = ET.SubElement(rpr, f'{{{W}}}bCs')
            bcs.set(f'{{{W}}}val', '0')
            s = ET.SubElement(rpr, f'{{{W}}}sz')
            s.set(f'{{{W}}}val', '24')
            t_el = ET.SubElement(r, f'{{{W}}}t')
            t_el.text = run_info['text']
            t_el.set(f'{XML_NS}space', 'preserve')
            paragraphs[41].append(r)

    # 替换中文关键词（P36）
    cn_keywords = extract_keywords(cn_abstract_blocks, blocks, chinese=True)
    if cn_keywords:
        replace_keyword_paragraph(paragraphs, 36, cn_keywords, chinese=True)

    # 替换英文关键词（P42）— TNR+宋体
    en_keywords = extract_keywords(en_abstract_blocks, blocks, chinese=False)
    if en_keywords:
        replace_keyword_paragraph(paragraphs, 42, en_keywords, chinese=False)


def _clean_paragraph_keep_ppr(p, keep_rpr=False):
    """清除段落中所有子元素（保留 pPr），包括 runs、AlternateContent 等。
    keep_rpr=True 时保留 pPr 中的 rPr（段落默认 run 属性，范文/模板均使用）。"""
    ppr = p.find(f'{{{W}}}pPr')
    children = list(p)
    for child in children:
        if child is not ppr:
            p.remove(child)
    if ppr is not None and not keep_rpr:
        for rpr in ppr.findall(f'{{{W}}}rPr'):
            ppr.remove(rpr)


def _add_jc_to_paragraph(p, jc_val):
    """给段落添加或更新 jc（对齐方式）。"""
    ppr = p.find(f'{{{W}}}pPr')
    if ppr is None:
        ppr = ET.SubElement(p, f'{{{W}}}pPr')
    existing_jc = ppr.find(f'{{{W}}}jc')
    if existing_jc is not None:
        existing_jc.set(f'{{{W}}}val', jc_val)
    else:
        jc = ET.SubElement(ppr, f'{{{W}}}jc')
        jc.set(f'{{{W}}}val', jc_val)


def _replace_abstract_body(paragraphs, start_idx, end_idx, content_blocks, chinese=True,
                           parent_body=None, insert_before_idx=None):
    """替换摘要正文段落，使用模板中的摘要格式（firstLineChars=200, firstLine=480, sz=24）。
    insert_before_idx: 溢出段落插入到此索引的段落之前。"""
    if not content_blocks:
        return

    for i in range(start_idx, end_idx + 1):
        if i < len(paragraphs):
            _clean_paragraph_keep_ppr(paragraphs[i], keep_rpr=True)

    # 将内容写入第一个段落
    p = paragraphs[start_idx]
    first_block = content_blocks[0]
    inline_runs = parse_inline(first_block['text'])
    for run_info in inline_runs:
        if not run_info['text']:
            continue
        if chinese:
            r = make_text_run(run_info['text'],
                              bold=run_info['bold'],
                              italic=run_info['italic'],
                              code=run_info.get('code', False),
                              font_east_asia='宋体',
                              sz=24, theme_fonts=True)
        else:
            r = make_text_run(run_info['text'],
                              bold=run_info['bold'],
                              italic=run_info['italic'],
                              code=run_info.get('code', False),
                              sz=24)
        p.append(r)

    # 后续段落（2~N）写入 P34, P35...
    for block_i, block in enumerate(content_blocks[1:], start=1):
        para_idx = start_idx + block_i
        if para_idx <= end_idx and para_idx < len(paragraphs):
            p = paragraphs[para_idx]
        else:
            # 超出模板段落数，创建新段落插入到 body 中
            p = ET.Element(f'{{{W}}}p')
            ppr = ET.SubElement(p, f'{{{W}}}pPr')
            if not chinese:
                # 英文摘要使用 af2 (Title) 样式
                pstyle = ET.SubElement(ppr, f'{{{W}}}pStyle')
                pstyle.set(f'{{{W}}}val', 'af2')
            adj = ET.SubElement(ppr, f'{{{W}}}adjustRightInd')
            adj.set(f'{{{W}}}val', '0')
            snap = ET.SubElement(ppr, f'{{{W}}}snapToGrid')
            snap.set(f'{{{W}}}val', '0')
            spacing = ET.SubElement(ppr, f'{{{W}}}spacing')
            spacing.set(f'{{{W}}}line', '360')
            spacing.set(f'{{{W}}}lineRule', 'auto')
            ind = ET.SubElement(ppr, f'{{{W}}}ind')
            ind.set(f'{{{W}}}firstLineChars', '200')
            ind.set(f'{{{W}}}firstLine', '480')
            # 插入到关键词段落之前（避免打乱后续段落顺序）
            if parent_body is not None and insert_before_idx is not None:
                before_p = paragraphs[insert_before_idx]
                idx = list(parent_body).index(before_p)
                parent_body.insert(idx, p)
            else:
                paragraphs.append(p)
                continue

        # 添加 runs（使用 pPr 中已有的格式，不再重新设置）
        inline_runs = parse_inline(block['text'])
        for run_info in inline_runs:
            if not run_info['text']:
                continue
            r = make_text_run(run_info['text'],
                              bold=run_info['bold'],
                              italic=run_info['italic'],
                              font_east_asia='宋体' if chinese else None,
                              sz=24, theme_fonts=True)
            p.append(r)


def body_paragraph_builder(block):
    """将 block 转换为正文段落。"""
    p = ET.Element(f'{{{W}}}p')
    ppr = ET.SubElement(p, f'{{{W}}}pPr')
    pstyle = ET.SubElement(ppr, f'{{{W}}}pStyle')
    pstyle.set(f'{{{W}}}val', 'afa')
    adj = ET.SubElement(ppr, f'{{{W}}}adjustRightInd')
    adj.set(f'{{{W}}}val', '0')
    snap = ET.SubElement(ppr, f'{{{W}}}snapToGrid')
    snap.set(f'{{{W}}}val', '0')
    ind = ET.SubElement(ppr, f'{{{W}}}ind')
    ind.set(f'{{{W}}}firstLineChars', '200')
    ind.set(f'{{{W}}}firstLine', '480')
    add_text_to_paragraph(p, block['text'])
    return p


def replace_paragraphs_range(paragraphs, start_idx, end_idx, content_blocks, builder_fn):
    """替换段落范围内的内容。用新块替换 P[start_idx] 到 P[end_idx] 的子 runs。"""
    if not content_blocks:
        return

    # 收集所有要替换的段落
    parents = []
    for i in range(start_idx, end_idx + 1):
        if i < len(paragraphs):
            parents.append(paragraphs[i])

    if not parents:
        return

    # 替换第一个段落的内容
    p = parents[0]
    # 移除所有 runs
    for r in list(p.findall(f'{{{W}}}r')):
        p.remove(r)
    # 添加新内容的 runs
    new_p = builder_fn(content_blocks[0])
    for r in new_p.findall(f'{{{W}}}r'):
        p.append(copy.deepcopy(r))

    # 清空多余的段落
    for i in range(1, len(parents)):
        p_empty = parents[i]
        for r in list(p_empty.findall(f'{{{W}}}r')):
            p_empty.remove(r)


def extract_keywords(abstract_blocks, all_blocks, chinese=True):
    """从所有 blocks 中提取关键词。"""
    for block in all_blocks:
        if block['type'] != 'paragraph':
            continue
        text = block['text']
        if chinese and '关键词' in text:
            # 提取 "关键词：xxx；yyy" 中的内容
            match = re.search(r'关键词[：:]\s*(.+)', text)
            if match:
                raw = match.group(1).strip().rstrip('；;')
                cleaned = strip_markdown_formatting(raw)
                # 清除可能残留的前后 ** 标记
                cleaned = cleaned.strip('*').strip()
                return cleaned
        elif not chinese and ('Key words' in text or 'Keywords' in text.lower()):
            match = re.search(r'(?:Key\s*words|Keywords)[：:]\s*(.+)', text, re.IGNORECASE)
            if match:
                raw = match.group(1).strip().rstrip(';,；，')
                cleaned = strip_markdown_formatting(raw)
                # 英文逗号分隔转为中文分号分隔（模板规范：关键词之间用分号间隔）
                cleaned = re.sub(r'\s*,\s*', '；', cleaned)
                cleaned = cleaned.strip('*').strip()
                return cleaned
    return None


def replace_keyword_paragraph(paragraphs, p_idx, keywords_text, chinese=True):
    """替换关键词段落。如果段落已被清理则重建完整内容，否则替换冒号后的内容。"""
    if p_idx >= len(paragraphs):
        return

    p = paragraphs[p_idx]
    runs = p.findall(f'{{{W}}}r')

    # 清除所有 runs 并重建
    for r in list(runs):
        p.remove(r)

    label = '关键词：' if chinese else 'Key words: '
    ea_font = '黑体' if chinese else '宋体'
    # 英文关键词标签不应加粗（范文Key words: 无bold），需用unbold覆盖af2样式继承
    label_run = make_text_run(label, font_east_asia=ea_font,
                               bold=False, unbold=(not chinese), sz=24, theme_fonts=True)
    p.append(label_run)

    # 用分号分隔的关键词列表
    if chinese:
        kw_parts = re.split(r'[；;]', keywords_text)
    else:
        kw_parts = re.split(r'[;；]', keywords_text)

    for i, kw in enumerate(kw_parts):
        kw = kw.strip()
        if not kw:
            continue
        if i > 0:
            sep_r = make_text_run('；', font_east_asia='宋体', sz=24,
                                  unbold=(not chinese), theme_fonts=True)
            p.append(sep_r)
        new_r = make_text_run(kw, font_east_asia='宋体',
                              sz=24, unbold=(not chinese), theme_fonts=True)
        p.append(new_r)


# ============================================================
# 9. 正文段落提取（区分标题和正文）
# ============================================================

def extract_body_blocks(blocks):
    """
    从 blocks 中提取正文部分（"1 绪论"开始到末尾）。
    跳过封面、声明、摘要等前置部分。
    """
    body_blocks = []
    in_body = False

    for block in blocks:
        if block['type'] == 'heading':
            text = strip_markdown_formatting(block['text'])
            # 正文从 "1 绪论" 或第一个带数字编号的 H1 开始
            if re.match(r'^1\s+\S', text) or re.match(r'^第[一二三四五六七八九十]', text):
                in_body = True
            # 也检测 "结论"、"致谢"、"参考文献"、"附录" 等 H1
            if text in ('结论', '致谢', '参考文献', '附录'):
                in_body = True

        if in_body:
            body_blocks.append(block)

    return body_blocks


# ============================================================
# 10. 主函数
# ============================================================

import argparse

def main():
    global _bookmark_counter, _drawing_id_counter
    register_namespaces()

    parser = argparse.ArgumentParser(description='将 Markdown 毕业论文转换为 Word 文档')
    parser.add_argument('input', nargs='?', default=None,
                        help='输入的 Markdown 文件路径 (默认: input/ 目录下第一个 .md 文件)')
    args = parser.parse_args()

    project_dir = os.path.dirname(os.path.abspath(__file__))
    template_dir = os.path.join(project_dir, 'template')
    input_dir = os.path.join(project_dir, 'input')
    output_dir = os.path.join(project_dir, 'output')
    os.makedirs(output_dir, exist_ok=True)
    source_xml = os.path.join(template_dir, 'word', 'document.xml')

    if args.input:
        md_file = os.path.abspath(args.input)
    else:
        md_files = sorted(f for f in os.listdir(input_dir) if f.endswith('.md'))
        if not md_files:
            print('错误: input/ 目录下没有 .md 文件', file=sys.stderr)
            sys.exit(1)
        md_file = os.path.join(input_dir, md_files[0])
        print(f'自动选择输入文件: {md_file}')

    from datetime import datetime
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    md_basename = os.path.splitext(os.path.basename(md_file))[0]
    output_docx = os.path.join(output_dir, f'{md_basename}_{timestamp}.docx')

    # 读取 Markdown
    with open(md_file, 'r', encoding='utf-8') as f:
        md_text = f.read()

    # 解析 Markdown
    blocks = parse_markdown(md_text)

    # 调试：统计不同类型的块
    block_types = {}
    for block in blocks:
        btype = block['type']
        block_types[btype] = block_types.get(btype, 0) + 1
    print(f'块类型统计: {block_types}')

    # 提取封面信息
    cover_data = parse_cover_table(md_text)
    print(f'封面信息: {cover_data}')

    # 解析 document.xml（从解压的毕业论文模板 XML 中读取）
    tree = ET.parse(source_xml)
    root = tree.getroot()
    body = root.find(f'{{{W}}}body')

    paragraphs = body.findall(f'{{{W}}}p')
    print(f'模板段落总数: {len(paragraphs)}')

    # 清除圆角矩形标注框（模板中的格式说明批注）
    _remove_callout_annotations(body)

    # 清除旧 _Toc bookmarks
    remove_old_bookmarks(body)
    max_existing_id = find_max_bookmark_id(body)
    _bookmark_counter = max_existing_id + 1
    print(f'Bookmark 起始 ID: {_bookmark_counter}')

    # 初始化 drawing ID 计数器（扫描模板中已有的最大 docPr id）
    max_docPr_id = 0
    wp_ns = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing'
    for dp in root.iter(f'{{{wp_ns}}}docPr'):
        try:
            max_docPr_id = max(max_docPr_id, int(dp.get('id', '0')))
        except ValueError:
            pass
    _drawing_id_counter = max_docPr_id
    print(f'Drawing ID 起始值: {_drawing_id_counter}')

    # === 1. 注入封面信息 ===
    inject_cover_data(paragraphs, cover_data)
    print('封面信息已注入')

    # === 2. 替换摘要（含清理格式说明文字，修正段落索引） ===
    current_children = list(body)
    replace_abstracts(paragraphs, blocks, parent_body=body, cover_data=cover_data)
    print('摘要已替换')

    # === 3. 生成 TOC 条目 ===
    body_blocks = extract_body_blocks(blocks)
    toc_entries = build_toc_entries(body_blocks)
    print(f'TOC 条目数: {len(toc_entries)}')

    # === 4. 替换 TOC 标题和条目 (P44=TOC标题, P45+=TOC条目) ===
    # 清理 TOC 标题段落 (P44)：只保留 "目录" 文字，清除格式说明
    toc_title_p = paragraphs[44]
    _clean_paragraph_keep_ppr(toc_title_p, keep_rpr=True)
    new_title_r = make_text_run('目录', bold=False, font_east_asia='黑体', sz=36)
    toc_title_p.append(new_title_r)

    # 重新获取当前 children
    current_children = list(body)
    toc_title_idx = current_children.index(toc_title_p)

    # 移除旧的 TOC 条目（TOC标题之后到正文分节符之前的所有元素）
    toc_remove = []
    body_sect_break_p = None
    for i in range(toc_title_idx + 1, len(current_children)):
        elem = current_children[i]
        tag = elem.tag.split('}')[-1] if '}' in elem.tag else elem.tag
        if tag == 'sectPr' or tag == 'p':
            # 检查是否是正文分节符（P65 附近，2026届新模板）
            if tag == 'p':
                ppr = elem.find(f'{{{W}}}pPr')
                if ppr is not None:
                    sect = ppr.find(f'{{{W}}}sectPr')
                    if sect is not None:
                        body_sect_break_p = elem
                        break
        if elem is not toc_title_p:
            toc_remove.append(elem)

    for elem in toc_remove:
        if elem in list(body):
            body.remove(elem)

    # 重新获取 children
    current_children = list(body)
    toc_title_idx = current_children.index(toc_title_p)

    # 插入新 TOC 条目
    for j, entry in enumerate(toc_entries):
        toc_p = build_toc_entry(entry['number'], entry['title'], entry['toc_level'], entry['bookmark'])
        body.insert(toc_title_idx + 1 + j, toc_p)

    print('TOC 已替换')

    # === 5. 替换正文内容 ===
    # 找到正文分节符（最后一个 sectPr 段落中的嵌套分节符）
    # 在 P65 附近查找包含 sectPr 的段落（2026届新模板）
    current_children = list(body)
    body_sect_idx = len(current_children) - 1
    for i, child in enumerate(current_children):
        tag = child.tag.split('}')[-1] if '}' in child.tag else child.tag
        if tag == 'p':
            ppr = child.find(f'{{{W}}}pPr')
            if ppr is not None:
                sect = ppr.find(f'{{{W}}}sectPr')
                if sect is not None:
                    # 检查 sectPr 的 type 是否为 nextPage
                    sect_type = sect.get(f'{{{W}}}type', '')
                    # 取最后一个分节符作为正文开始点
                    body_sect_idx = list(body).index(child)

    # 找到最后一个 body > sectPr
    last_sectPr = body.find(f'{{{W}}}sectPr')

    # 移除正文分节符之后、last_sectPr 之前的所有元素
    current_children = list(body)
    to_remove = []
    for i, child in enumerate(current_children):
        if child is last_sectPr:
            continue
        if i > body_sect_idx:
            to_remove.append(child)

    for elem in to_remove:
        body.remove(elem)

    # 用于跟踪新增的图片/公式媒体文件和关系
    media_files_dict = {}
    rels_counter = {'count': 0, '_rels_entries': []}

    # === 预扫描：为所有表格（markdown 表 + 代码表）分配章节内连续编号 ===
    _table_num_map = {}   # {block_index: (chapter_prefix, table_number)}
    _pre_ch = 0
    _pre_ch_letter = ''   # 附录章节字母（A, B, C...）
    _pre_cnt = {}
    for _si, _sb in enumerate(body_blocks):
        if _sb['type'] == 'heading':
            _st = strip_markdown_formatting(_sb['text'])
            _m = re.match(r'^(\d+)', _st)
            if _m:
                _pre_ch = int(_m.group(1))
                _pre_ch_letter = ''
            else:
                # 检测附录章节：附录A、附录B...
                _am = re.match(r'^附录([A-Z])', _st)
                if _am:
                    _pre_ch_letter = _am.group(1)
                    _pre_ch = 0
        elif _sb['type'] == 'table':
            if _pre_ch > 0:
                _pre_cnt[_pre_ch] = _pre_cnt.get(_pre_ch, 0) + 1
                _table_num_map[_si] = (str(_pre_ch), _pre_cnt[_pre_ch])
            elif _pre_ch_letter:
                _key = f'_{_pre_ch_letter}'
                _pre_cnt[_key] = _pre_cnt.get(_key, 0) + 1
                _table_num_map[_si] = (_pre_ch_letter, _pre_cnt[_key])
        elif _sb['type'] == 'code_table':
            # code_table 统一参与编号，表注由 code_table 处理器在表格上方渲染
            if _pre_ch > 0:
                _pre_cnt[_pre_ch] = _pre_cnt.get(_pre_ch, 0) + 1
                _table_num_map[_si] = (str(_pre_ch), _pre_cnt[_pre_ch])
            elif _pre_ch_letter:
                _key = f'_{_pre_ch_letter}'
                _pre_cnt[_key] = _pre_cnt.get(_key, 0) + 1
                _table_num_map[_si] = (_pre_ch_letter, _pre_cnt[_key])

    # 插入新的正文内容
    first_h1_done = False
    in_no_indent_section = False  # 参考文献、致谢等不需要首行缩进的 section
    in_reference_section = False  # 是否在参考文献部分
    current_chapter = 0           # 当前章节号
    current_chapter_letter = ''   # 当前附录章节字母
    _eq_counter = {}              # {chapter: count} 公式按章编号
    _fig_counter = {}             # {chapter: count} 图按章编号
    for idx, block in enumerate(body_blocks):
        btype = block['type']
        if btype == 'heading':
            heading_text = strip_markdown_formatting(block['text'])
            # 检测是否进入/离开不需要首行缩进的 section
            in_no_indent_section = heading_text == '参考文献'
            in_reference_section = heading_text == '参考文献'

            # _force_h1: 裸数字标题（如 "4 多Agent..."）提升为 H1 级别渲染
            orig_level = block['level']

            # 附录等后记部分的子标题：用普通段落样式渲染，不用标题样式
            if block.get('_plain_heading'):
                p = ET.Element(f'{{{W}}}p')
                ppr = ET.SubElement(p, f'{{{W}}}pPr')
                pstyle = ET.SubElement(ppr, f'{{{W}}}pStyle')
                pstyle.set(f'{{{W}}}val', 'afa')
                p_jc = ET.SubElement(ppr, f'{{{W}}}jc')
                p_jc.set(f'{{{W}}}val', 'center')
                r = make_text_run(heading_text, bold=True, font_east_asia='黑体', sz=24)
                p.append(r)
                body.append(p)
                continue

            if block.get('_force_h1'):
                block['level'] = 1
            # H1 标题：提取章节号或附录字母
            if block['level'] == 1:
                m = re.match(r'^(\d+)', heading_text)
                if m:
                    current_chapter = int(m.group(1))
                    current_chapter_letter = ''
                else:
                    am = re.match(r'^附录([A-Z])', heading_text)
                    if am:
                        current_chapter_letter = am.group(1)
                        current_chapter = 0
            # H2 标题：检测附录子章节（如 "附录A ..." 在 H2 级别）
            elif block['level'] == 2:
                am = re.match(r'^附录([A-Z])', heading_text)
                if am:
                    current_chapter_letter = am.group(1)
                    current_chapter = 0
            # H1 之前插入分页符 + 空行
            if block['level'] == 1 and first_h1_done:
                p_break = build_page_break_paragraph()
                body.append(p_break)
            # H1 标题前空一行
            if block['level'] == 1:
                p_blank_before = ET.Element(f'{{{W}}}p')
                p_blank_before_ppr = ET.SubElement(p_blank_before, f'{{{W}}}pPr')
                p_blank_before_style = ET.SubElement(p_blank_before_ppr, f'{{{W}}}pStyle')
                p_blank_before_style.set(f'{{{W}}}val', 'afa')
                # 范文 H1 前空行居中（P81, P114）
                p_blank_before_jc = ET.SubElement(p_blank_before_ppr, f'{{{W}}}jc')
                p_blank_before_jc.set(f'{{{W}}}val', 'center')
                p_blank_before_spacing = ET.SubElement(p_blank_before_ppr, f'{{{W}}}spacing')
                p_blank_before_spacing.set(f'{{{W}}}line', '360')
                p_blank_before_spacing.set(f'{{{W}}}lineRule', 'auto')
                p_blank_before_rpr = ET.SubElement(p_blank_before_ppr, f'{{{W}}}rPr')
                p_blank_before_sz = ET.SubElement(p_blank_before_rpr, f'{{{W}}}sz')
                p_blank_before_sz.set(f'{{{W}}}val', '24')
                p_blank_before_szcs = ET.SubElement(p_blank_before_rpr, f'{{{W}}}szCs')
                p_blank_before_szcs.set(f'{{{W}}}val', '24')
                body.append(p_blank_before)
            if block['level'] == 1:
                first_h1_done = True
            p = build_heading(block,
                              bookmark_name=block.get('_bookmark'),
                              bookmark_id=block.get('_bookmark_id'))
            body.append(p)
            # H1 标题后空一行（范文 P67: style=afc 空行, P81: style=afa jc=center 空行）
            if block['level'] == 1:
                p_blank_after = ET.Element(f'{{{W}}}p')
                p_blank_after_ppr = ET.SubElement(p_blank_after, f'{{{W}}}pPr')
                p_blank_after_style = ET.SubElement(p_blank_after_ppr, f'{{{W}}}pStyle')
                p_blank_after_style.set(f'{{{W}}}val', 'afa')
                p_blank_after_jc = ET.SubElement(p_blank_after_ppr, f'{{{W}}}jc')
                p_blank_after_jc.set(f'{{{W}}}val', 'center')
                body.append(p_blank_after)
            # 恢复原始 level（_force_h1 临时提升）
            block['level'] = orig_level
        elif btype == 'paragraph':
            # 如果当前段落是表格标题且下一个块是 markdown 表格，跳过（由表格处理器统一渲染表注）
            _next_block = body_blocks[idx + 1] if idx + 1 < len(body_blocks) else None
            if (_next_block and _next_block['type'] == 'table'
                    and re.match(r'^\*{0,2}(续)?表\d', block['text'])):
                continue
            # 如果当前段落是 code_table 的描述或正式表注，跳过（由 code_table 处理器统一渲染表注）
            if _next_block and _next_block['type'] == 'code_table':
                _ct_desc = _next_block.get('description', '')
                _curr_text = strip_markdown_formatting(block['text'])
                # 跳过正式表注（"表X-Y ..."）或与 code_table description 匹配的段落
                if re.match(r'^\*{0,2}(续)?表\d', block['text']):
                    continue
                if _ct_desc and _ct_desc == _curr_text:
                    continue
            # 参考文献条目特殊处理
            if in_reference_section and _is_reference_paragraph(block['text']):
                p = build_reference_paragraph(block)
            else:
                p = build_body_paragraph(block, no_indent=in_no_indent_section)
            body.append(p)
        elif btype == 'bullet':
            p = build_bullet_paragraph(block)
            body.append(p)
        elif btype == 'numbered':
            p = build_numbered_paragraph(block)
            body.append(p)
        elif btype == 'blockquote':
            p = build_blockquote(block)
            body.append(p)
        elif btype == 'codeblock':
            # 检查是否在附录中
            # 后备处理：代码块作为普通段落
            code_paras = build_codeblock(block)
            if isinstance(code_paras, list):
                for cp in code_paras:
                    body.append(cp)
            else:
                body.append(code_paras)
        elif btype == 'code_table':
            # 生成表注（使用统一编号：表X-Y + 代码描述）
            description = block.get('description', '')
            _tbl_num = _table_num_map.get(idx)
            if description and _tbl_num:
                _ch_prefix, _tbl_seq = _tbl_num
                if _ch_prefix and _tbl_seq:
                    caption_text = f'表{_ch_prefix}-{_tbl_seq} {description}'
                p_cap = ET.Element(f'{{{W}}}p')
                p_cap_ppr = ET.SubElement(p_cap, f'{{{W}}}pPr')
                p_cap_style = ET.SubElement(p_cap_ppr, f'{{{W}}}pStyle')
                p_cap_style.set(f'{{{W}}}val', 'a4')
                p_cap_jc = ET.SubElement(p_cap_ppr, f'{{{W}}}jc')
                p_cap_jc.set(f'{{{W}}}val', 'center')
                p_cap_adj = ET.SubElement(p_cap_ppr, f'{{{W}}}adjustRightInd')
                p_cap_adj.set(f'{{{W}}}val', '0')
                p_cap_snap = ET.SubElement(p_cap_ppr, f'{{{W}}}snapToGrid')
                p_cap_snap.set(f'{{{W}}}val', '0')
                p_cap_spacing = ET.SubElement(p_cap_ppr, f'{{{W}}}spacing')
                p_cap_spacing.set(f'{{{W}}}line', '360')
                p_cap_spacing.set(f'{{{W}}}lineRule', 'auto')
                # 范文表注格式：asciiTheme + eastAsia='黑体'
                cap_r = ET.SubElement(p_cap, f'{{{W}}}r')
                cap_rpr = ET.SubElement(cap_r, f'{{{W}}}rPr')
                cap_rfonts = ET.SubElement(cap_rpr, f'{{{W}}}rFonts')
                cap_rfonts.set(f'{{{W}}}asciiTheme', 'minorHAnsi')
                cap_rfonts.set(f'{{{W}}}hAnsiTheme', 'minorHAnsi')
                cap_rfonts.set(f'{{{W}}}cstheme', 'minorHAnsi')
                cap_rfonts.set(f'{{{W}}}eastAsia', '黑体')
                cap_rsz = ET.SubElement(cap_rpr, f'{{{W}}}sz')
                cap_rsz.set(f'{{{W}}}val', '21')
                cap_rszcs = ET.SubElement(cap_rpr, f'{{{W}}}szCs')
                cap_rszcs.set(f'{{{W}}}val', '21')
                cap_rt = ET.SubElement(cap_r, f'{{{W}}}t')
                cap_rt.text = caption_text
                cap_rt.set(f'{XML_NS}space', 'preserve')
                body.append(p_cap)
            table = build_code_table(block)
            body.append(table)
        elif btype == 'mermaid':
            # 图按章编号：重写 caption 中的图序号
            _ch_key = current_chapter_letter if current_chapter_letter else str(current_chapter)
            if _ch_key and block.get('caption'):
                _fig_counter[_ch_key] = _fig_counter.get(_ch_key, 0) + 1
                _fig_label = f'图{_ch_key}-{_fig_counter[_ch_key]}'
                block = dict(block)
                # 替换已有 "图X-X" 或添加编号前缀
                _cap = block['caption']
                _cap_new = re.sub(r'^图[\dA-Z]+-\d+', _fig_label, _cap)
                if _cap_new == _cap:
                    _cap_new = _fig_label + ' ' + _cap
                block['caption'] = _cap_new
            mermaid_paras = build_mermaid_block(block, media_files_dict, rels_counter)
            if isinstance(mermaid_paras, list):
                for mp in mermaid_paras:
                    body.append(mp)
            else:
                body.append(mermaid_paras)
        elif btype == 'formula':
            # 公式按章编号：注入 \tag{章-序号}
            _ch_key = current_chapter_letter if current_chapter_letter else str(current_chapter)
            if _ch_key:
                _eq_counter[_ch_key] = _eq_counter.get(_ch_key, 0) + 1
                _eq_label = f'{_ch_key}-{_eq_counter[_ch_key]}'
                # 如果 LaTeX 中没有 \tag{}，自动添加
                _latex = block['latex'].strip()
                if r'\tag{' not in _latex:
                    block = dict(block)  # shallow copy to avoid mutating original
                    block['latex'] = _latex.rstrip('$').rstrip() + r' \tag{' + _eq_label + '}'
            formula_paras = build_formula_paragraph(block, media_files_dict, rels_counter)
            if isinstance(formula_paras, list):
                for fp in formula_paras:
                    body.append(fp)
            else:
                body.append(formula_paras)
        elif btype == 'image':
            # 图按章编号：重写 alt 中的图序号
            _ch_key = current_chapter_letter if current_chapter_letter else str(current_chapter)
            if _ch_key and block.get('alt'):
                _fig_counter[_ch_key] = _fig_counter.get(_ch_key, 0) + 1
                _fig_label = f'图{_ch_key}-{_fig_counter[_ch_key]}'
                block = dict(block)
                _alt = block['alt']
                _alt_new = re.sub(r'^图[\dA-Z]+-\d+', _fig_label, _alt)
                if _alt_new == _alt:
                    _alt_new = _fig_label + ' ' + _alt
                block['alt'] = _alt_new
            # 下载图片并嵌入
            image_paras = build_image_block(block, media_files_dict, rels_counter)
            if isinstance(image_paras, list):
                for ip in image_paras:
                    body.append(ip)
            else:
                body.append(image_paras)
        elif btype == 'table':
            # 检查上一个块是否为表格标题（表格标题在表格上方，段落处理器已跳过渲染）
            caption_block = None
            caption_text = None
            if idx - 1 >= 0:
                prev_block = body_blocks[idx - 1]
                if prev_block['type'] == 'paragraph':
                    prev_text = prev_block.get('text', '')
                    if re.match(r'^\*{0,2}(续)?表\d', prev_text):
                        caption_block = prev_block
                        caption_text = prev_text
            # 使用统一编号重写表注中的序号
            _tbl_num = _table_num_map.get(idx)
            if _tbl_num:
                _ch_prefix, _tbl_seq = _tbl_num
                _new_label = f'表{_ch_prefix}-{_tbl_seq}'
                if caption_text:
                    caption_text = re.sub(r'表[\dA-Z]+-\d+', _new_label, caption_text)
            if caption_block:
                # 更新 caption_block 的文本以使用新编号
                if _tbl_num:
                    _ch_prefix, _tbl_seq = _tbl_num
                    _new_label = f'表{_ch_prefix}-{_tbl_seq}'
                    caption_block = copy.deepcopy(caption_block)
                    caption_block['text'] = re.sub(r'表[\dA-Z]+-\d+', _new_label, caption_block['text'])
                p_cap = build_body_paragraph(caption_block)
                body.append(p_cap)
            tbl_elements = build_table(block, caption_text=caption_text)
            for elem in tbl_elements:
                body.append(elem)
        # hr blocks 已在 parse_markdown 中跳过
        else:
            continue

    # 确保 last_sectPr 在最后
    if last_sectPr is not None:
        if last_sectPr in list(body):
            body.remove(last_sectPr)
        body.append(last_sectPr)

    print(f'正文内容已替换 ({len(body_blocks)} 个块)')

    # === 6. 输出 docx ===
    output_xml = os.path.join(template_dir, 'word', 'document_new.xml')
    tree.write(output_xml, encoding='UTF-8', xml_declaration=True)

    # 从解压的文件重新构建 docx ZIP 容器
    docx_files = {
        '[Content_Types].xml': '[Content_Types].xml',
        '_rels/.rels': '_rels/.rels',
        'customXml/item1.xml': 'customXml/item1.xml',
        'customXml/item2.xml': 'customXml/item2.xml',
        'customXml/itemProps1.xml': 'customXml/itemProps1.xml',
        'customXml/itemProps2.xml': 'customXml/itemProps2.xml',
        'customXml/_rels/item1.xml.rels': 'customXml/_rels/item1.xml.rels',
        'customXml/_rels/item2.xml.rels': 'customXml/_rels/item2.xml.rels',
        'docProps/app.xml': 'docProps/app.xml',
        'docProps/core.xml': 'docProps/core.xml',
        'docProps/custom.xml': 'docProps/custom.xml',
        'word/document.xml': output_xml,
        'word/_rels/document.xml.rels': 'word/_rels/document.xml.rels',
        'word/endnotes.xml': 'word/endnotes.xml',
        'word/fontTable.xml': 'word/fontTable.xml',
        'word/footer1.xml': 'word/footer1.xml',
        'word/footer2.xml': 'word/footer2.xml',
        'word/footer3.xml': 'word/footer3.xml',
        'word/footnotes.xml': 'word/footnotes.xml',
        'word/header1.xml': 'word/header1.xml',
        'word/header2.xml': 'word/header2.xml',
        'word/header3.xml': 'word/header3.xml',
        'word/header4.xml': 'word/header4.xml',
        'word/header5.xml': 'word/header5.xml',
        'word/header6.xml': 'word/header6.xml',
        'word/header7.xml': 'word/header7.xml',
        'word/_rels/header3.xml.rels': 'word/_rels/header3.xml.rels',
        'word/_rels/header5.xml.rels': 'word/_rels/header5.xml.rels',
        'word/_rels/header6.xml.rels': 'word/_rels/header6.xml.rels',
        'word/_rels/header7.xml.rels': 'word/_rels/header7.xml.rels',
        'word/media/image1.jpeg': 'word/media/image1.jpeg',
        'word/media/image2.png': 'word/media/image2.png',
        'word/media/image3.png': 'word/media/image3.png',
        'word/media/image4.emf': 'word/media/image4.emf',
        'word/media/image5.png': 'word/media/image5.png',
        'word/media/image6.wmf': 'word/media/image6.wmf',
        'word/numbering.xml': 'word/numbering.xml',
        'word/settings.xml': 'word/settings.xml',
        'word/styles.xml': 'word/styles.xml',
        'word/theme/theme1.xml': 'word/theme/theme1.xml',
        'word/webSettings.xml': 'word/webSettings.xml',
        'word/embeddings/oleObject1.bin': 'word/embeddings/oleObject1.bin',
        'word/embeddings/Microsoft_Visio_Drawing.vsdx': 'word/embeddings/Microsoft_Visio_Drawing.vsdx',
    }

    # 添加动态生成的图片/公式媒体文件
    for zip_path, local_path in media_files_dict.items():
        if zip_path not in docx_files:
            docx_files[zip_path] = local_path

    # 生成更新后的 document.xml.rels（添加新图片关系）
    _update_rels_file(template_dir, media_files_dict, rels_counter)

    with zipfile.ZipFile(output_docx, 'w', zipfile.ZIP_DEFLATED) as zf_out:
        for zip_path, local_path in docx_files.items():
            full_path = os.path.join(template_dir, local_path)
            if os.path.exists(full_path):
                data = open(full_path, 'rb').read()
                if zip_path == 'word/settings.xml':
                    try:
                        settings_str = data.decode('utf-8')
                        if 'updateFields' not in settings_str:
                            settings_str = settings_str.replace(
                                '</w:settings>',
                                '<w:updateFields w:val="true"/></w:settings>'
                            )
                        data = settings_str.encode('utf-8')
                    except Exception:
                        pass
                zf_out.writestr(zip_path, data)

    # 清理临时文件
    if os.path.exists(output_xml):
        os.remove(output_xml)

    # 清理 media 目录中动态生成的文件（mermaid_*/formula_*/image_*/snake_*）
    # 保留模板原始文件（image1.jpeg ~ image6.wmf）
    media_dir = os.path.join(template_dir, 'word', 'media')
    _TEMPLATE_MEDIA_FILES = {
        'image1.jpeg', 'image2.png', 'image3.png',
        'image4.emf', 'image5.png', 'image6.wmf',
    }
    if os.path.isdir(media_dir):
        removed = 0
        for fname in os.listdir(media_dir):
            if fname in _TEMPLATE_MEDIA_FILES:
                continue
            if fname.startswith(('mermaid_', 'formula_', 'image_', 'snake_')):
                os.remove(os.path.join(media_dir, fname))
                removed += 1
        if removed:
            print(f'已清理 {removed} 个临时媒体文件')

    print(f'\n生成文件: {output_docx}')
    print(f'转换了 {len(body_blocks)} 个正文块')
    print(f'生成了 {len(toc_entries)} 个目录条目')


def _update_rels_file(template_dir, media_files_dict, rels_counter):
    """更新 document.xml.rels，为新增的图片添加关系条目。"""
    rels_path = os.path.join(template_dir, 'word', '_rels', 'document.xml.rels')
    if not os.path.exists(rels_path):
        return

    try:
        with open(rels_path, 'r', encoding='utf-8') as f:
            rels_content = f.read()

        # 移除旧的 rIdImage* 条目（每次重新生成）
        import re as _re
        rels_content = _re.sub(r'<Relationship\s+Id="rIdImage\d+"[^>]*/>\s*', '', rels_content)

        # 重新添加所有图片关系
        rels_entries = rels_counter.get('_rels_entries', [])
        for r_id, target in rels_entries:
            rel_entry = (f'<Relationship Id="{r_id}" '
                         f'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" '
                         f'Target="{target}"/>')
            rels_content = rels_content.replace('</Relationships>', rel_entry + '</Relationships>')

        with open(rels_path, 'w', encoding='utf-8') as f:
            f.write(rels_content)

    except Exception as e:
        print(f'  警告: 更新 rels 文件失败: {e}', file=sys.stderr)


if __name__ == '__main__':
    main()
