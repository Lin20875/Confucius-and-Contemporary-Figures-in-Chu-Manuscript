# -*- coding: utf-8 -*-
"""
将"相關文獻匯編"系列 docx 文件转换为结构化 HTML 页面
v4: 真实表格渲染 + 标题识别 + 去掉逐段方框背景 + Word 自动编号前缀 + 域代码圆圈数字
"""
import os
import sys
import re
from zipfile import ZipFile
import xml.etree.ElementTree as ET

if sys.platform == 'win32':
    sys.stdout.reconfigure(encoding='utf-8')

NS = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'

DOCX_DIR = r'相關文獻匯編\相關文獻匯編'
OUTPUT_DIR = r'articles\huibian'

ARTICLE_DOCX_MAP = {
    "民之父母": "《民之父母》相關文獻彙編.docx",
    "窮達以時": "《窮達以時》相關文獻彙編.docx",
    "仲弓": "仲弓 相關文獻彙編 20240806.docx",
    "史蒥問於夫子": "史蒥問於夫子 相關文獻彙編 20250210.docx",
    "君子為禮": "君子為禮 相關文獻彙編 20240819.docx",
    "子羔": "子羔 相關文獻彙編 20250218.docx",
    "孔子見季桓子": "孔子見季桓子 相關文獻彙編 20250122.docx",
    "季庚子問於孔子": "季庚子問於孔子20250228（修改版）.docx",
    "弟子問": "弟子問 相關文獻彙編 20240921.docx",
    "相邦之道": "相邦之道 相關文獻彙編 20250203.docx",
    "邦家之政": "邦家之政 相關文獻彙編 20250114.docx",
    "顏淵問於孔子": "顏淵問於孔子 相關文獻彙編 20240806.docx",
    "魯邦大旱": "魯邦大旱 相關文獻彙編 20240721.docx",
}

SKIP_REGEN = {"民之父母", "窮達以時"}

# ─── 编号系统 ───

_CHINESE_NUMS = '零一二三四五六七八九十'

def _decimal_to_chinese(n):
    if n <= 10:
        return _CHINESE_NUMS[n]
    if n < 20:
        return '十' + (_CHINESE_NUMS[n - 10] if n > 10 else '')
    t, u = divmod(n, 10)
    return _CHINESE_NUMS[t] + '十' + (_CHINESE_NUMS[u] if u else '')

_TIANGAN = '甲乙丙丁戊己庚辛壬癸'

def _format_number(n, fmt):
    if fmt == 'decimal':
        return str(n)
    if fmt in ('taiwaneseCountingThousand', 'chineseCountingThousand'):
        return _decimal_to_chinese(n)
    if fmt == 'ideographTraditional':
        return _TIANGAN[(n - 1) % 10] if n >= 1 else '?'
    if fmt == 'upperLetter':
        return chr(64 + n) if 1 <= n <= 26 else str(n)
    if fmt == 'lowerLetter':
        return chr(96 + n) if 1 <= n <= 26 else str(n)
    if fmt == 'lowerRoman':
        vals = [(1000,'m'),(900,'cm'),(500,'d'),(400,'cd'),(100,'c'),(90,'xc'),
                (50,'l'),(40,'xl'),(10,'x'),(9,'ix'),(5,'v'),(4,'iv'),(1,'i')]
        r = ''
        for v, s in vals:
            while n >= v:
                r += s; n -= v
        return r
    if fmt == 'upperRoman':
        return _format_number(n, 'lowerRoman').upper()
    return str(n)


def parse_numbering(zf):
    """解析 numbering.xml，返回 numId -> {ilvl: {fmt, text}} 映射"""
    if 'word/numbering.xml' not in zf.namelist():
        return {}
    tree = ET.parse(zf.open('word/numbering.xml'))
    root = tree.getroot()

    abstract_defs = {}
    for abnum in root.findall('.//w:abstractNum', NS):
        abid = abnum.get(f'{{{W}}}abstractNumId')
        levels = {}
        for lvl in abnum.findall('w:lvl', NS):
            ilvl = int(lvl.get(f'{{{W}}}ilvl', '0'))
            nfmt = lvl.find('w:numFmt', NS)
            lvltext = lvl.find('w:lvlText', NS)
            fmt_val = nfmt.get(f'{{{W}}}val', 'decimal') if nfmt is not None else 'decimal'
            txt_val = lvltext.get(f'{{{W}}}val', '') if lvltext is not None else ''
            levels[ilvl] = {'fmt': fmt_val, 'text': txt_val}
        abstract_defs[abid] = levels

    num_map = {}
    for num_el in root.findall('.//w:num', NS):
        nid = num_el.get(f'{{{W}}}numId')
        abref = num_el.find('w:abstractNumId', NS)
        if abref is not None:
            abid = abref.get(f'{{{W}}}val')
            if abid in abstract_defs:
                num_map[nid] = abstract_defs[abid]
    return num_map


def compute_prefix(num_map, counters, numId, ilvl):
    """根据 numId 和 ilvl 计算编号前缀文本"""
    if numId not in num_map:
        return ''
    levels = num_map[numId]
    if ilvl not in levels:
        return ''
    lvl = levels[ilvl]
    fmt = lvl['fmt']
    tmpl = lvl['text']

    if fmt == 'bullet':
        ch = tmpl.strip()
        if ch in ('-', '\u2013', '\u2014', '\uFF0D'):
            return ch + ' '
        return '\u00b7 '

    key = (numId, ilvl)
    counters[key] = counters.get(key, 0) + 1
    n = counters[key]

    result = tmpl
    placeholder = f'%{ilvl + 1}'
    result = result.replace(placeholder, _format_number(n, fmt))
    return result if result else ''


_RE_CIRCLE_NUM = re.compile(r'eq\s+\\o\\ac\(○\s*,\s*(\d+)\)')


def _field_to_char(instr_text):
    """将 Word 域代码转换为对应字符，如 eq \\o\\ac(○,1) → ①"""
    m = _RE_CIRCLE_NUM.search(instr_text)
    if m:
        n = int(m.group(1))
        if 1 <= n <= 20:
            return chr(0x2460 + n - 1)
        if 21 <= n <= 35:
            return chr(0x3251 + n - 21)
        if 36 <= n <= 50:
            return chr(0x32B1 + n - 36)
        return f'({n})'
    return None


def _get_para_data(para_el, num_map=None, counters=None):
    """从一个 w:p 元素提取文本、样式、粗体信息、编号前缀和域代码字符"""
    pPr = para_el.find('w:pPr', NS)
    style_id = ''
    num_prefix = ''
    if pPr is not None:
        pStyle = pPr.find('w:pStyle', NS)
        if pStyle is not None:
            style_id = pStyle.get(f'{{{W}}}val', '')
        if num_map is not None and counters is not None:
            numPr = pPr.find('w:numPr', NS)
            if numPr is not None:
                numId_el = numPr.find('w:numId', NS)
                ilvl_el = numPr.find('w:ilvl', NS)
                if numId_el is not None:
                    numId = numId_el.get(f'{{{W}}}val', '0')
                    ilvl = int(ilvl_el.get(f'{{{W}}}val', '0')) if ilvl_el is not None else 0
                    if numId != '0':
                        num_prefix = compute_prefix(num_map, counters, numId, ilvl)

    runs_data = []
    in_field = False
    for run in para_el.findall('.//w:r', NS):
        fldChar = run.find('w:fldChar', NS)
        if fldChar is not None:
            ftype = fldChar.get(f'{{{W}}}fldCharType', '')
            if ftype == 'begin':
                in_field = True
            elif ftype in ('end', 'separate'):
                if ftype == 'end':
                    in_field = False
            continue

        if in_field:
            instrText = run.find('w:instrText', NS)
            if instrText is not None and instrText.text:
                ch = _field_to_char(instrText.text)
                if ch:
                    runs_data.append((ch, False))
            continue

        rPr = run.find('w:rPr', NS)
        is_bold = False
        if rPr is not None:
            b = rPr.find('w:b', NS)
            if b is not None:
                val = b.get(f'{{{W}}}val')
                is_bold = val is None or val != '0'
        text_parts = []
        for t in run.findall('.//w:t', NS):
            if t.text:
                text_parts.append(t.text)
        if text_parts:
            runs_data.append((''.join(text_parts), is_bold))

    full_text = ''.join(r[0] for r in runs_data).strip()
    has_bold = any(r[1] for r in runs_data if r[0].strip())
    all_bold = all(r[1] for r in runs_data if r[0].strip()) if runs_data else False

    return {
        'text': full_text,
        'style': style_id,
        'has_bold': has_bold,
        'all_bold': all_bold,
        'runs': runs_data,
        'num_prefix': num_prefix,
    }


def _get_table_data(tbl_el):
    """从一个 w:tbl 元素提取表格数据，保留粗体"""
    rows = []
    for tr in tbl_el.findall('w:tr', NS):
        cells = []
        for tc in tr.findall('w:tc', NS):
            cell_runs = []
            for p in tc.findall('w:p', NS):
                for run in p.findall('.//w:r', NS):
                    rPr = run.find('w:rPr', NS)
                    is_bold = False
                    if rPr is not None:
                        b = rPr.find('w:b', NS)
                        if b is not None:
                            val = b.get(f'{{{W}}}val')
                            is_bold = val is None or val != '0'
                    text_parts = []
                    for t in run.findall('.//w:t', NS):
                        if t.text:
                            text_parts.append(t.text)
                    if text_parts:
                        cell_runs.append((''.join(text_parts), is_bold))
            cells.append(cell_runs)
        if any(any(r[0].strip() for r in cell) for cell in cells if cell):
            rows.append(cells)
    return rows


def extract_body_elements(docx_path):
    """从 docx 按顺序提取段落和表格，返回 list of ('para', data) | ('table', data)"""
    elements = []
    with ZipFile(docx_path) as z:
        num_map = parse_numbering(z)
        counters = {}

        with z.open('word/document.xml') as f:
            tree = ET.parse(f)
            root = tree.getroot()
            body = root.find('.//w:body', NS)

            for child in body:
                tag = child.tag.split('}')[-1] if '}' in child.tag else child.tag
                if tag == 'p':
                    para = _get_para_data(child, num_map, counters)
                    if para['text']:
                        elements.append(('para', para))
                elif tag == 'tbl':
                    table = _get_table_data(child)
                    if table:
                        elements.append(('table', table))
    return elements


# ── heading / content-type patterns ──

_RE_CHINESE_NUM_HEADING = re.compile(
    r'^[一二三四五六七八九十百]+[、．.\s]')
_RE_NUMBERED_SOURCE = re.compile(
    r'^\d+[、．.\s]')
_RE_DASH_HEADING = re.compile(
    r'^[-\-－—]\s*\S')
_RE_SECTION_INTRO = re.compile(
    r'^[-\-－—]?\s*[〈《].*[〉》]\s*相[應应][內内]容')
_RE_STAR_HEADING = re.compile(
    r'^\*')


def classify_paragraph(para):
    """根据样式和内容模式分类段落"""
    style = para['style']
    text = para['text']
    tlen = len(text)

    if style in ('1', 'Heading1'):
        return 'h2'
    if style in ('2', 'Heading2'):
        return 'h3'
    if style in ('3', 'Heading3'):
        return 'h4'
    if style in ('4', 'Heading4'):
        return 'h5'

    if _RE_SECTION_INTRO.match(text) and tlen < 40:
        return 'section_intro'

    if para['all_bold'] and tlen < 100:
        if _RE_CHINESE_NUM_HEADING.match(text):
            return 'h2'
        if _RE_NUMBERED_SOURCE.match(text):
            return 'h4'
        if _RE_DASH_HEADING.match(text) and tlen < 60:
            return 'sub_heading'
        if _RE_STAR_HEADING.match(text) and tlen < 80:
            return 'sub_heading'
        return 'bold_heading'

    if not para['all_bold'] and tlen < 100:
        if _RE_CHINESE_NUM_HEADING.match(text):
            if para['has_bold'] or style in ('', 'HTML', 'Web', 'a3', 'a9'):
                return 'h2'
        if _RE_NUMBERED_SOURCE.match(text) and tlen < 60:
            return 'h4'
        if _RE_DASH_HEADING.match(text) and tlen < 40:
            return 'sub_heading'

    if text.startswith('簡文') and (text.startswith('簡文：') or text.startswith('簡文:')) and tlen > 30:
        return 'bamboo'

    return 'normal'


def escape_html(text):
    return text.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')


def render_runs_html(runs):
    parts = []
    for text, is_bold in runs:
        escaped = escape_html(text)
        if is_bold and text.strip():
            parts.append(f'<strong>{escaped}</strong>')
        else:
            parts.append(escaped)
    return ''.join(parts)


def render_cell_html(cell_runs):
    if not cell_runs:
        return ''
    return render_runs_html(cell_runs)


def _is_header_row(row):
    """判断表格首行是否为表头（所有单元格文字较短，像列标签）"""
    for cell in row:
        text = ''.join(r[0] for r in cell).strip()
        if len(text) > 20:
            return False
    return True


def render_table_html(table_rows):
    """将表格数据渲染为 HTML <table>"""
    lines = ['<div class="hb-table-wrap"><table class="hb-table">']
    has_header = len(table_rows) > 1 and _is_header_row(table_rows[0])
    for ri, row in enumerate(table_rows):
        lines.append('  <tr>')
        tag = 'th' if (ri == 0 and has_header) else 'td'
        for cell in row:
            content = render_cell_html(cell)
            if not content.strip():
                content = '&mdash;'
            lines.append(f'    <{tag}>{content}</{tag}>')
        lines.append('  </tr>')
    lines.append('</table></div>')
    return '\n'.join(lines)


def _render_with_prefix(tag, cls, text_html, prefix):
    """渲染带编号前缀的 HTML 元素"""
    pfx = escape_html(prefix) if prefix else ''
    if pfx:
        pfx_html = f'<span class="num-pfx">{pfx}</span>'
        if tag in ('h2', 'h3', 'h4', 'h5'):
            return f'<{tag} class="{cls}">{pfx_html}{text_html}</{tag}>'
        return f'<p class="{cls} hb-listed">{pfx_html}{text_html}</p>'
    if tag == 'p':
        return f'<p class="{cls}">{text_html}</p>'
    return f'<{tag} class="{cls}">{text_html}</{tag}>'


def build_html(title, elements):
    body_parts = []

    for etype, data in elements:
        if etype == 'table':
            body_parts.append(render_table_html(data))
            continue

        para = data
        cat = classify_paragraph(para)
        text_html = render_runs_html(para['runs'])
        plain = para['text']
        pfx = para.get('num_prefix', '')

        if plain == title or plain in (
            f'〈{title}〉相關文獻彙編',
            f'《{title}》相關文獻彙編',
            f'〈{title}〉 相關文獻彙編',
        ):
            continue

        if cat == 'h2':
            body_parts.append(_render_with_prefix('h2', 'hb-section', text_html, pfx))
        elif cat == 'h3':
            body_parts.append(_render_with_prefix('h3', 'hb-topic', text_html, pfx))
        elif cat == 'h4':
            body_parts.append(_render_with_prefix('h4', 'hb-source', text_html, pfx))
        elif cat == 'h5':
            body_parts.append(_render_with_prefix('h5', 'hb-subsource', text_html, pfx))
        elif cat == 'section_intro':
            body_parts.append(_render_with_prefix('h3', 'hb-intro', text_html, pfx))
        elif cat == 'sub_heading':
            body_parts.append(_render_with_prefix('h4', 'hb-subhead', text_html, pfx))
        elif cat == 'bold_heading':
            body_parts.append(_render_with_prefix('h4', 'hb-subhead', text_html, pfx))
        elif cat == 'bamboo':
            body_parts.append(f'<div class="hb-bamboo"><p>{text_html}</p></div>')
        else:
            body_parts.append(_render_with_prefix('p', 'hb-text', text_html, pfx))

    body_content = '\n    '.join(body_parts)

    return f'''<!DOCTYPE html>
<html lang="zh-Hant">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width,initial-scale=1">
  <title>{escape_html(title)} — 相關文獻彙編</title>
  <style>
    :root {{
      --bg: #faf8f2;
      --fg: #1a1a1a;
      --accent: #2e6b4f;
      --accent2: #8c6d1f;
      --muted: #666;
      --border: #d5cdb8;
      --section-bg: #f0ece0;
      --bamboo-bg: #e8f0e8;
      --table-header-bg: #e8e2d0;
      --table-stripe: #f5f2ea;
    }}
    * {{ box-sizing: border-box; }}
    body {{
      margin: 0; padding: 0;
      font-family: "Noto Serif CJK TC","PingFang TC","Microsoft JhengHei",serif;
      line-height: 1.8;
      background: var(--bg);
      color: var(--fg);
    }}
    header, main, footer {{
      max-width: 960px;
      margin: 0 auto;
      padding: clamp(1.2rem,3vw,2.5rem) clamp(1rem,3vw,2rem);
    }}
    header {{
      border-bottom: 2px solid var(--accent);
      background: linear-gradient(135deg, rgba(46,107,79,.08), transparent);
    }}
    header h1 {{
      font-size: clamp(1.6rem,3vw,2.4rem);
      margin: 0 0 .3em;
      color: var(--accent);
      letter-spacing: .06em;
    }}
    header .subtitle {{
      margin: 0;
      font-size: .95rem;
      color: var(--muted);
    }}

    /* ── 大段标题：带背景条 ── */
    .hb-section {{
      font-size: clamp(1.25rem,2.2vw,1.7rem);
      font-weight: bold;
      color: var(--accent);
      margin: 2.5rem 0 1rem;
      padding: .4em .7em;
      background: var(--section-bg);
      border-left: 4px solid var(--accent);
      border-radius: 0 .5rem .5rem 0;
    }}

    /* ── 中级标题：无背景 ── */
    .hb-topic {{
      font-size: clamp(1.1rem,1.8vw,1.35rem);
      font-weight: bold;
      color: var(--accent2);
      margin: 2rem 0 .6rem;
      padding: .2em 0;
      border-bottom: 1px dashed var(--border);
    }}

    /* ── 篇名引导 ── */
    .hb-intro {{
      font-size: clamp(1.05rem,1.6vw,1.2rem);
      font-weight: bold;
      color: var(--accent);
      margin: 1.8rem 0 .5rem;
      padding: 0;
    }}

    /* ── 小标题 / 来源标题 ── */
    .hb-source {{
      font-size: clamp(1rem,1.5vw,1.1rem);
      font-weight: bold;
      color: #444;
      margin: 1.5rem 0 .4rem;
      padding-left: .5em;
      border-left: 3px solid var(--accent2);
    }}
    .hb-subsource {{
      font-size: 1rem;
      font-weight: bold;
      color: #555;
      margin: 1.2rem 0 .3rem;
      padding-left: 1em;
    }}
    .hb-subhead {{
      font-size: clamp(1.02rem,1.5vw,1.12rem);
      font-weight: bold;
      color: var(--accent2);
      margin: 1.5rem 0 .5rem;
      padding: 0;
    }}

    /* ── 正文段落：无背景框 ── */
    .hb-text {{
      margin: .6em 0;
      font-size: .95rem;
      text-indent: 2em;
    }}

    /* ── 带编号前缀的段落 ── */
    .hb-listed {{
      text-indent: 0;
      padding-left: 2.8em;
      position: relative;
    }}
    .num-pfx {{
      position: absolute;
      left: 0;
      display: inline-block;
      min-width: 2.5em;
      text-align: right;
      padding-right: .3em;
      color: var(--accent2);
      font-weight: bold;
    }}
    h2 .num-pfx, h3 .num-pfx, h4 .num-pfx {{
      position: static;
      min-width: auto;
      text-align: left;
      padding-right: .2em;
    }}

    /* ── 简文区块：保留背景框 ── */
    .hb-bamboo {{
      margin: 1em 0;
      padding: .8em 1.2em;
      background: var(--bamboo-bg);
      border: 1px solid #b5cfb5;
      border-radius: .5rem;
      font-size: .95rem;
      line-height: 1.9;
    }}
    .hb-bamboo p {{ margin: 0; }}

    /* ── 真实表格 ── */
    .hb-table-wrap {{
      margin: 1.2em 0;
      overflow-x: auto;
    }}
    .hb-table {{
      width: 100%;
      border-collapse: collapse;
      font-size: .9rem;
      line-height: 1.7;
    }}
    .hb-table th,
    .hb-table td {{
      border: 1px solid var(--border);
      padding: .5em .7em;
      vertical-align: top;
      text-align: left;
    }}
    .hb-table th {{
      background: var(--table-header-bg);
      font-weight: bold;
      color: var(--accent);
      white-space: nowrap;
    }}
    .hb-table tr:nth-child(even) td {{
      background: var(--table-stripe);
    }}

    footer {{
      border-top: 1px solid var(--border);
      font-size: .85rem;
      color: var(--muted);
      padding-top: 1.5rem;
      padding-bottom: 2rem;
      margin-top: 3rem;
    }}
    footer p {{ margin: .3em 0; }}

    @media(max-width:768px) {{
      .hb-section {{ font-size: 1.15rem; }}
      header, main, footer {{
        padding: clamp(1rem,5vw,1.5rem) clamp(.8rem,4vw,1.2rem);
      }}
      .hb-table {{ font-size: .82rem; }}
      .hb-table th, .hb-table td {{ padding: .35em .5em; }}
    }}
  </style>
</head>
<body>
  <header>
    <h1>{escape_html(title)} — 相關文獻彙編</h1>
    <p class="subtitle">與讀本對照之相關傳世文獻及學者研究匯集</p>
  </header>
  <main>
    {body_content}
  </main>
  <footer>
    <p>本頁內容整理自《{escape_html(title)}》相關文獻彙編</p>
  </footer>
</body>
</html>'''


def main():
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    converted = []
    skipped = []
    for article_name, docx_name in ARTICLE_DOCX_MAP.items():
        if article_name in SKIP_REGEN:
            skipped.append(article_name)
            print(f'  [SKIP-MANUAL] {article_name} (已手动修改，跳过)')
            continue

        docx_path = os.path.join(DOCX_DIR, docx_name)
        if not os.path.exists(docx_path):
            print(f'  [SKIP] {docx_name} not found')
            continue

        print(f'Converting: {docx_name} ...')
        elements = extract_body_elements(docx_path)

        if elements and elements[0][0] == 'para':
            first = elements[0][1]
            if '相關文獻' in first['text'] or first['text'].strip() in (
                f'〈{article_name}〉相關文獻彙編',
                f'《{article_name}》相關文獻彙編',
                f'〈{article_name}〉 相關文獻彙編',
            ):
                elements = elements[1:]

        html = build_html(article_name, elements)
        output_path = os.path.join(OUTPUT_DIR, f'{article_name}_匯編.html')
        with open(output_path, 'w', encoding='utf-8') as fout:
            fout.write(html)
        n_para = sum(1 for e in elements if e[0] == 'para')
        n_tbl = sum(1 for e in elements if e[0] == 'table')
        n_pfx = sum(1 for e in elements if e[0] == 'para' and e[1].get('num_prefix'))
        converted.append(article_name)
        print(f'  -> {output_path} ({n_para} paragraphs, {n_tbl} tables, {n_pfx} numbered)')

    print(f'\nDone! Converted {len(converted)} files, skipped {len(skipped)} manually edited.')
    print('Converted:', converted)
    if skipped:
        print('Skipped (manual edits preserved):', skipped)


if __name__ == '__main__':
    main()
