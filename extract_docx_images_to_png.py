# -*- coding: utf-8 -*-
"""
从DOCX文件中按模块顺序提取文字和图片
改进版：更全面地识别所有图片引用，并按Word模块组织输出
"""
import os
import sys
import subprocess
from zipfile import ZipFile
import xml.etree.ElementTree as ET
from pathlib import Path
import re
from datetime import datetime

# 设置输出编码为UTF-8
if sys.platform == 'win32':
    try:
        sys.stdout.reconfigure(encoding='utf-8')
    except:
        pass

def convert_emf_to_png_powershell(emf_path, png_path):
    """使用PowerShell和.NET转换EMF到PNG"""
    emf_path_escaped = str(emf_path).replace('\\', '\\\\')
    png_path_escaped = str(png_path).replace('\\', '\\\\')
    
    ps_script = f'''
try {{
    Add-Type -AssemblyName System.Drawing
    $emf = New-Object System.Drawing.Imaging.Metafile("{emf_path_escaped}")
    $bitmap = New-Object System.Drawing.Bitmap($emf.Width, $emf.Height)
    $graphics = [System.Drawing.Graphics]::FromImage($bitmap)
    $graphics.Clear([System.Drawing.Color]::White)
    $graphics.DrawImage($emf, 0, 0, $emf.Width, $emf.Height)
    $bitmap.Save("{png_path_escaped}", [System.Drawing.Imaging.ImageFormat]::Png)
    $graphics.Dispose()
    $bitmap.Dispose()
    $emf.Dispose()
    Write-Output "SUCCESS"
}} catch {{
    Write-Output "ERROR: $($_.Exception.Message)"
    exit 1
}}
'''
    
    try:
        result = subprocess.run(
            ['powershell', '-ExecutionPolicy', 'Bypass', '-Command', ps_script],
            capture_output=True,
            text=True,
            timeout=30
        )
        return result.returncode == 0 and "SUCCESS" in result.stdout
    except Exception as e:
        return False

def convert_to_png(input_path, output_path):
    """将图片转换为PNG格式"""
    input_path = Path(input_path)
    output_path = Path(output_path)
    
    # 如果已经是PNG，直接复制
    if input_path.suffix.lower() == '.png':
        import shutil
        shutil.copy2(input_path, output_path)
        return True
    
    # 如果是EMF，使用PowerShell转换
    if input_path.suffix.lower() == '.emf':
        return convert_emf_to_png_powershell(input_path, output_path)
    
    # 其他格式，尝试使用PIL转换
    try:
        from PIL import Image
        img = Image.open(input_path)
        # 如果是RGBA模式，保持透明度；否则转换为RGB
        if img.mode in ('RGBA', 'LA'):
            img.save(output_path, 'PNG')
        else:
            img.convert('RGB').save(output_path, 'PNG')
        return True
    except ImportError:
        print(f"  警告: PIL未安装，无法转换 {input_path.suffix} 格式")
        return False
    except Exception as e:
        print(f"  警告: 转换失败 {input_path.suffix}: {e}")
        return False

def resolve_media_path(target, rels_dir, zf_namelist_func):
    """解析media文件的完整路径"""
    # 确保 zf_namelist_func 是可调用的
    if callable(zf_namelist_func):
        zf_namelist = zf_namelist_func()
    else:
        zf_namelist = zf_namelist_func
    
    # 处理相对路径
    if target.startswith('../'):
        parent_dir = os.path.dirname(rels_dir)
        full_target = os.path.normpath(os.path.join(parent_dir, target)).replace('\\', '/')
    elif target.startswith('word/'):
        full_target = target
    elif 'media' in target.lower():
        # 如果已经是完整路径
        if target in zf_namelist:
            return target
        # 尝试添加word/前缀
        if not target.startswith('word/'):
            full_target = f"word/{target}"
        else:
            full_target = target
    else:
        # 相对路径，需要计算
        if rels_dir == 'word/_rels':
            full_target = f"word/{target}"
        else:
            parent_dir = os.path.dirname(rels_dir)
            full_target = os.path.normpath(os.path.join(parent_dir, target)).replace('\\', '/')
    
    # 尝试多种可能的路径
    possible_paths = [
        full_target,
        f"word/{target}" if not target.startswith('word/') else target,
        target if target.startswith('word/') else f"word/{target}",
    ]
    
    for path in possible_paths:
        if path in zf_namelist:
            return path
    
    return None

def extract_image_references_from_xml(xml_bytes, all_rid_to_target, ns, ns_rels):
    """从XML中提取所有图片引用，返回按文档顺序的图片路径列表（保留重复）"""
    ordered_targets = []
    root = ET.fromstring(xml_bytes)
    
    # 定义所有可能的关系ID属性
    r_embed_attr = '{%s}embed' % ns['r']
    r_id_attr = '{%s}id' % ns['r']
    r_link_attr = '{%s}link' % ns['r']
    
    # 用于跟踪当前元素是否已添加图片，避免同一元素被多个方法重复匹配
    processed_elements = set()
    
    # 按文档顺序遍历所有元素
    for elem in root.iter():
        # 为每个元素生成唯一标识符
        elem_id = id(elem)
        if elem_id in processed_elements:
            continue
        
        found_target = None
        
        # 方法1: 检查r:embed属性（DrawingML格式，如a:blip）
        rId = elem.attrib.get(r_embed_attr)
        if rId and rId in all_rid_to_target:
            found_target = all_rid_to_target[rId]
        
        # 方法2: 检查r:id属性（VML格式，如v:imagedata）
        if not found_target:
            rId = elem.attrib.get(r_id_attr)
            if rId and rId in all_rid_to_target:
                found_target = all_rid_to_target[rId]
        
        # 方法3: 检查r:link属性
        if not found_target:
            rId = elem.attrib.get(r_link_attr)
            if rId and rId in all_rid_to_target:
                found_target = all_rid_to_target[rId]
        
        # 方法4: 检查所有属性值，查找可能的关系ID
        if not found_target:
            for attr_name, attr_value in elem.attrib.items():
                if attr_value and attr_value in all_rid_to_target:
                    found_target = all_rid_to_target[attr_value]
                    break
        
        # 如果找到目标，添加到列表（保留重复，因为同一个图片可能在文档中出现多次）
        if found_target:
            ordered_targets.append(found_target)
            processed_elements.add(elem_id)  # 标记当前元素已处理，避免同一元素被多个方法重复匹配
    
    return ordered_targets

def extract_text_from_xml(xml_bytes, all_rid_to_target, ns, ns_rels, image_counter):
    """从XML中提取文字内容，返回段落列表和图片引用"""
    root = ET.fromstring(xml_bytes)
    paragraphs = []
    
    # 定义所有可能的关系ID属性
    r_embed_attr = '{%s}embed' % ns['r']
    r_id_attr = '{%s}id' % ns['r']
    ns_v = {'v': 'urn:schemas-microsoft-com:vml'}
    
    # 遍历所有段落
    for para in root.findall('.//w:p', ns):
        para_text = []
        para_images = []
        
        # 按顺序收集段落中的所有运行（runs）
        runs = para.findall('.//w:r', ns)
        
        for run in runs:
            # 检查是否有脚注引用
            footnote_ref = run.find('.//w:footnoteReference', ns)
            if footnote_ref is not None:
                footnote_id = footnote_ref.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}id')
                if footnote_id:
                    para_text.append(f'[脚注{footnote_id}]')
                continue
            
            # 检查是否有图片
            found_image = None
            
            # 检查a:blip
            blip = run.find('.//a:blip', ns)
            if blip is not None:
                rId = blip.attrib.get(r_embed_attr)
                if rId and rId in all_rid_to_target:
                    found_image = all_rid_to_target[rId]
            
            # 检查v:imagedata
            if not found_image:
                imd = run.find('.//v:imagedata', ns_v)
                if imd is not None:
                    rId = imd.attrib.get(r_id_attr)
                    if rId and rId in all_rid_to_target:
                        found_image = all_rid_to_target[rId]
            
            if found_image:
                image_counter[0] += 1
                para_text.append(f'[圖字{image_counter[0]:03d}]')
                para_images.append((image_counter[0], found_image))
                continue
            
            # 提取文本
            text_elements = run.findall('.//w:t', ns)
            for text_elem in text_elements:
                if text_elem.text:
                    para_text.append(text_elem.text)
        
        # 如果没有找到运行元素，尝试直接提取段落文本
        if not para_text:
            text_elements = para.findall('.//w:t', ns)
            for text_elem in text_elements:
                if text_elem.text:
                    para_text.append(text_elem.text)
        
        if para_text or para_images:
            paragraphs.append({
                'text': ''.join(para_text),
                'images': para_images
            })
    
    return paragraphs

def get_module_name(part_path):
    """根据XML路径获取模块名称"""
    if part_path == 'word/document.xml':
        return '正文'
    elif part_path.startswith('word/header'):
        match = re.search(r'header(\d+)', part_path)
        if match:
            return f'页眉{match.group(1)}'
        return '页眉'
    elif part_path.startswith('word/footer'):
        match = re.search(r'footer(\d+)', part_path)
        if match:
            return f'页脚{match.group(1)}'
        return '页脚'
    elif part_path == 'word/footnotes.xml':
        return '脚注'
    elif part_path == 'word/endnotes.xml':
        return '尾注'
    elif part_path == 'word/comments.xml':
        return '批注'
    elif part_path == 'word/numbering.xml':
        return '编号'
    else:
        return os.path.basename(part_path).replace('.xml', '')

def main():
    input_path = r"C:\Users\lyue\Desktop\出土文献读本网页\articles\民之父母.docx"
    
    # 图片输出目录（用户指定）
    images_output_dir = Path(r"C:\Users\lyue\Desktop\出土文献读本网页\articles\images_民之父母_20260112")
    
    # 文字输出目录（在同一位置创建文字文件夹）
    base_output_dir = Path(r"C:\Users\lyue\Desktop\出土文献读本网页\articles")
    timestamp = datetime.now().strftime('%Y%m%d')
    text_output_dir = base_output_dir / f'民之父母_提取_{timestamp}'
    
    if not os.path.exists(input_path):
        print(f'❌ DOCX文件不存在: {input_path}')
        sys.exit(1)
    
    # 创建输出目录
    if images_output_dir.exists():
        print(f'⚠️  图片输出目录已存在，将清空并替换现有文件')
        # 清空目录中的旧文件
        import shutil
        for file in images_output_dir.iterdir():
            if file.is_file():
                file.unlink()
            elif file.is_dir():
                shutil.rmtree(file)
    
    images_output_dir.mkdir(parents=True, exist_ok=True)
    text_output_dir.mkdir(parents=True, exist_ok=True)
    
    # 创建文字子目录
    text_dir = text_output_dir / '文字'
    text_dir.mkdir(exist_ok=True)
    
    print('=' * 60)
    print('从DOCX按模块顺序提取文字和图片')
    print('=' * 60)
    print(f'输入文件: {input_path}')
    print(f'图片输出目录: {images_output_dir}')
    print(f'文字输出目录: {text_output_dir}')
    print()
    
    with ZipFile(input_path, 'r') as zf:
        # 需要扫描的部件（按文档顺序）
        parts_in_order = [
            'word/document.xml',
            *sorted([n for n in zf.namelist() if n.startswith('word/header') and n.endswith('.xml')]),
            *sorted([n for n in zf.namelist() if n.startswith('word/footer') and n.endswith('.xml')]),
            *[n for n in ['word/footnotes.xml', 'word/endnotes.xml', 'word/comments.xml', 'word/numbering.xml'] if n in zf.namelist()],
        ]
        
        # XML 命名空间
        ns = {
            'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
            'wp': 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
            'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
            'pic': 'http://schemas.openxmlformats.org/drawingml/2006/picture',
            'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
            'v': 'urn:schemas-microsoft-com:vml',  # 添加VML命名空间
        }
        ns_rels = {'r': 'http://schemas.openxmlformats.org/package/2006/relationships'}
        ns_v = {'v': 'urn:schemas-microsoft-com:vml'}
        
        # 第一步：收集所有 rels 文件中的图片关系
        all_rid_to_target = {}  # 全局的关系映射
        
        print('📋 步骤1: 扫描所有关系文件...')
        for name in zf.namelist():
            if '/_rels/' in name and name.endswith('.rels'):
                try:
                    rels_xml = zf.read(name)
                    rels_root = ET.fromstring(rels_xml)
                    rels_dir = os.path.dirname(name)
                    
                    for rel in rels_root.findall('r:Relationship', ns_rels):
                        rId = rel.attrib.get('Id')
                        target = rel.attrib.get('Target')
                        if rId and target and isinstance(target, str):
                            # 检查是否是图片文件
                            target_lower = target.lower()
                            if 'media' in target_lower or any(target_lower.endswith(ext) for ext in ['.png', '.jpg', '.jpeg', '.gif', '.bmp', '.emf', '.wmf', '.tiff', '.tif']):
                                resolved_path = resolve_media_path(target, rels_dir, zf.namelist)
                                if resolved_path:
                                    all_rid_to_target[rId] = resolved_path
                                    all_rid_to_target[resolved_path] = resolved_path
                except Exception as e:
                    print(f"  警告: 解析 {name} 时出错: {e}")
                    continue
        
        print(f'  ✓ 从 rels 文件找到 {len([k for k in all_rid_to_target.keys() if not k.startswith("word/")])} 个图片关系')
        
        # 第二步：按模块顺序提取文字和图片
        print('📋 步骤2: 按模块顺序提取文字和图片...')
        print()
        
        global_image_counter = [0]  # 全局图片计数器
        all_extracted_images = {}  # 存储所有提取的图片 {image_path: (module_name, local_index)}
        
        for part_xml_path in parts_in_order:
            try:
                xml_bytes = zf.read(part_xml_path)
            except KeyError:
                continue
            
            module_name = get_module_name(part_xml_path)
            print(f'📄 处理模块: {module_name}')
            
            # 创建模块文字文件夹（图片直接保存到根目录）
            module_text_dir = text_dir / module_name
            module_text_dir.mkdir(exist_ok=True)
            
            # 对应的 rels 文件
            dirname, filename = os.path.split(part_xml_path)
            rels_path = f"{dirname}/_rels/{filename}.rels"
            
            part_rid_to_target = {}
            if rels_path in zf.namelist():
                try:
                    rels_xml = zf.read(rels_path)
                    rels_root = ET.fromstring(rels_xml)
                    for rel in rels_root.findall('r:Relationship', ns_rels):
                        rId = rel.attrib.get('Id')
                        target = rel.attrib.get('Target')
                        if rId and target and isinstance(target, str):
                            target_lower = target.lower()
                            if 'media' in target_lower or any(target_lower.endswith(ext) for ext in ['.png', '.jpg', '.jpeg', '.gif', '.bmp', '.emf', '.wmf', '.tiff', '.tif']):
                                resolved_path = resolve_media_path(target, os.path.dirname(rels_path), zf.namelist)
                                if resolved_path:
                                    part_rid_to_target[rId] = resolved_path
                except Exception as e:
                    pass
            
            # 合并关系映射
            combined_rid_to_target = {**all_rid_to_target, **part_rid_to_target}
            
            # 提取文字和图片引用
            paragraphs = extract_text_from_xml(xml_bytes, combined_rid_to_target, ns, ns_rels, global_image_counter)
            
            # 保存文字内容
            text_file = module_text_dir / f'{module_name}.txt'
            with open(text_file, 'w', encoding='utf-8') as f:
                for para in paragraphs:
                    if para['text'].strip():
                        f.write(para['text'] + '\n\n')
                    # 记录图片引用
                    for img_idx, img_path in para['images']:
                        all_extracted_images[img_path] = (module_name, img_idx)
            
            # 提取并保存图片（直接保存到指定目录，使用全局计数器）
            module_image_count = 0
            for para in paragraphs:
                for img_idx, img_path in para['images']:
                    module_image_count += 1
                    try:
                        data = zf.read(img_path)
                        _, ext = os.path.splitext(img_path)
                        if not ext:
                            ext = '.bin'
                        
                        temp_path = images_output_dir / f"temp_{img_idx:03d}{ext.lower()}"
                        png_path = images_output_dir / f"{img_idx:03d}.png"
                        
                        with open(temp_path, 'wb') as f:
                            f.write(data)
                        
                        if convert_to_png(temp_path, png_path):
                            if temp_path.exists():
                                temp_path.unlink()
                            print(f'  ✓ 图片 {img_idx:03d}: {os.path.basename(img_path)} -> {png_path.name}')
                        else:
                            if temp_path.exists():
                                if ext.lower() == '.png':
                                    temp_path.rename(png_path)
                                    print(f'  ✓ 图片 {img_idx:03d}: {os.path.basename(img_path)} -> {png_path.name} (直接复制)')
                                else:
                                    temp_path.unlink()
                    except Exception as e:
                        print(f'  ❌ 提取图片失败 {img_path}: {e}')
            
            print(f'  ✓ 文字段落: {len(paragraphs)} 个')
            print(f'  ✓ 图片: {module_image_count} 个')
            print()
        
        # 第三步：检查是否有未引用的图片
        print('📋 步骤3: 检查所有media文件...')
        all_media_list = sorted([n for n in zf.namelist() if 'media/' in n and not n.endswith('/')])
        unreferenced_files = [f for f in all_media_list if f not in all_extracted_images]
        
        if unreferenced_files:
            print(f'  ⚠️  发现 {len(unreferenced_files)} 个未被引用的media文件')
            
            for idx, img_path in enumerate(unreferenced_files, start=1):
                try:
                    data = zf.read(img_path)
                    _, ext = os.path.splitext(img_path)
                    if not ext:
                        ext = '.bin'
                    
                    # 使用全局计数器之后的编号
                    unref_idx = global_image_counter[0] + idx
                    temp_path = images_output_dir / f"temp_{unref_idx:03d}{ext.lower()}"
                    png_path = images_output_dir / f"{unref_idx:03d}.png"
                    
                    with open(temp_path, 'wb') as f:
                        f.write(data)
                    
                    if convert_to_png(temp_path, png_path):
                        if temp_path.exists():
                            temp_path.unlink()
                        print(f'  ✓ 未引用图片 {unref_idx:03d}: {os.path.basename(img_path)}')
                except Exception as e:
                    print(f'  ❌ 提取未引用图片失败 {img_path}: {e}')
        
        print()
        print('=' * 60)
        print('提取完成！')
        print(f'   文字模块: {len(parts_in_order)} 个')
        print(f'   图片总数: {global_image_counter[0]} 个')
        if unreferenced_files:
            print(f'   未引用图片: {len(unreferenced_files)} 个')
        print('=' * 60)
        print()
        print(f'图片输出目录: {images_output_dir}')
        print(f'文字输出目录: {text_output_dir}')

if __name__ == '__main__':
    main()
