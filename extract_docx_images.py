import os
from zipfile import ZipFile
import xml.etree.ElementTree as ET

input_path = r"C:\Users\lyue\Desktop\网页\君子為禮 廣義讀本 20240520.docx"
output_dir = os.path.join(os.path.dirname(input_path), 'images_君子為禮_20251201')
os.makedirs(output_dir, exist_ok=True)

with ZipFile(input_path, 'r') as zf:
    # 需要扫描的部件（按常见阅读顺序）：正文 -> 页眉 -> 页脚 -> 脚注 -> 尾注 -> 批注
    parts_in_order = [
        'word/document.xml',
        # headers/footers（按文件名排序保证固定顺序）
        *sorted([n for n in zf.namelist() if n.startswith('word/header') and n.endswith('.xml')]),
        *sorted([n for n in zf.namelist() if n.startswith('word/footer') and n.endswith('.xml')]),
        # notes/comments
        *[n for n in ['word/footnotes.xml', 'word/endnotes.xml', 'word/comments.xml'] if n in zf.namelist()],
    ]

    # XML 命名空间
    ns = {
        'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
        'wp': 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
        'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
        'pic': 'http://schemas.openxmlformats.org/drawingml/2006/picture',
        'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    }
    ns_v = {'v': 'urn:schemas-microsoft-com:vml'}
    ns_rels = {'r': 'http://schemas.openxmlformats.org/package/2006/relationships'}

    # 全部目标文件（按出现顺序，不去重，相同图片也按顺序提取）
    ordered_targets = []

    for part_xml_path in parts_in_order:
        try:
            xml_bytes = zf.read(part_xml_path)
        except KeyError:
            continue

        # 对应的 rels 文件路径，如 word/_rels/document.xml.rels
        dirname, filename = os.path.split(part_xml_path)
        rels_path = f"{dirname}/_rels/{filename}.rels"

        rid_to_target = {}
        if rels_path in zf.namelist():
            rels_xml = zf.read(rels_path)
            rels_root = ET.fromstring(rels_xml)
            for rel in rels_root.findall('r:Relationship', ns_rels):
                rId = rel.attrib.get('Id')
                target = rel.attrib.get('Target')  # 可能是 'media/imageX.png'
                if rId and target:
                    # 只处理内嵌媒体文件
                    if target.startswith('media/'):
                        rid_to_target[rId] = f"word/{target}"

        # 解析该部件 XML，按出现顺序收集 r:embed 与 v:imagedata
        root = ET.fromstring(xml_bytes)

        # a:blip r:embed - 按出现顺序添加，不去重
        for blip in root.findall('.//a:blip', ns):
            rId = blip.attrib.get('{%s}embed' % ns['r'])
            if not rId:
                continue
            target = rid_to_target.get(rId)
            if not target:
                continue
            # 不去重，每次出现都按顺序添加
            if target in zf.namelist():
                ordered_targets.append(target)

        # v:imagedata r:id - 按出现顺序添加，不去重
        for imd in root.findall('.//v:imagedata', ns_v):
            rId = imd.attrib.get('{%s}id' % ns['r'])
            if not rId:
                continue
            target = rid_to_target.get(rId)
            if not target:
                continue
            # 不去重，每次出现都按顺序添加
            if target in zf.namelist():
                ordered_targets.append(target)

    # 写出文件
    saved = []
    for idx, target in enumerate(ordered_targets, start=1):
        data = zf.read(target)
        _, ext = os.path.splitext(target)
        if not ext:
            ext = '.bin'
        filename = f"{idx:03d}{ext.lower()}"
        out_path = os.path.join(output_dir, filename)
        with open(out_path, 'wb') as f:
            f.write(data)
        saved.append(filename)

print(output_dir)
print("EXTRACTED=" + str(len(saved)))
for name in saved:
    print(name)