# -*- coding: utf-8 -*-
"""
从DOCX文件中提取所有图片并转换为PNG格式
"""
import os
import sys
import subprocess
from zipfile import ZipFile
import xml.etree.ElementTree as ET
from pathlib import Path

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

def main():
    input_path = r"C:\Users\lyue\Desktop\网页\君子為禮 廣義讀本 20240520.docx"
    output_dir = Path(r"C:\Users\lyue\Desktop\网页\articles\images_君子為禮_20251201")
    
    if not os.path.exists(input_path):
        print(f'❌ DOCX文件不存在: {input_path}')
        sys.exit(1)
    
    output_dir.mkdir(parents=True, exist_ok=True)
    
    print('=' * 60)
    print('从DOCX提取图片并转换为PNG')
    print('=' * 60)
    print(f'输入文件: {input_path}')
    print(f'输出目录: {output_dir}')
    print()
    
    with ZipFile(input_path, 'r') as zf:
        # 需要扫描的部件
        parts_in_order = [
            'word/document.xml',
            *sorted([n for n in zf.namelist() if n.startswith('word/header') and n.endswith('.xml')]),
            *sorted([n for n in zf.namelist() if n.startswith('word/footer') and n.endswith('.xml')]),
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
        
        # 收集所有图片（按出现顺序，不去重）
        ordered_targets = []
        
        for part_xml_path in parts_in_order:
            try:
                xml_bytes = zf.read(part_xml_path)
            except KeyError:
                continue
            
            # 对应的 rels 文件
            dirname, filename = os.path.split(part_xml_path)
            rels_path = f"{dirname}/_rels/{filename}.rels"
            
            rid_to_target = {}
            if rels_path in zf.namelist():
                rels_xml = zf.read(rels_path)
                rels_root = ET.fromstring(rels_xml)
                for rel in rels_root.findall('r:Relationship', ns_rels):
                    rId = rel.attrib.get('Id')
                    target = rel.attrib.get('Target')
                    if rId and target and target.startswith('media/'):
                        rid_to_target[rId] = f"word/{target}"
            
            # 解析XML
            root = ET.fromstring(xml_bytes)
            
            # a:blip r:embed - 按出现顺序添加，不去重
            for blip in root.findall('.//a:blip', ns):
                rId = blip.attrib.get('{%s}embed' % ns['r'])
                if rId and rId in rid_to_target:
                    target = rid_to_target[rId]
                    if target in zf.namelist():
                        # 不去重，每次出现都添加
                        ordered_targets.append(target)
            
            # v:imagedata r:id - 按出现顺序添加，不去重
            for imd in root.findall('.//v:imagedata', ns_v):
                rId = imd.attrib.get('{%s}id' % ns['r'])
                if rId and rId in rid_to_target:
                    target = rid_to_target[rId]
                    if target in zf.namelist():
                        # 不去重，每次出现都添加
                        ordered_targets.append(target)
        
        # 提取并转换图片
        print(f'找到 {len(ordered_targets)} 个图片')
        print()
        
        saved = []
        converted = 0
        failed = 0
        
        for idx, target in enumerate(ordered_targets, start=1):
            try:
                # 读取原始图片数据
                data = zf.read(target)
                
                # 临时保存原始文件
                _, ext = os.path.splitext(target)
                if not ext:
                    ext = '.bin'
                
                temp_path = output_dir / f"temp_{idx:03d}{ext.lower()}"
                png_path = output_dir / f"{idx:03d}.png"
                
                # 写入临时文件
                with open(temp_path, 'wb') as f:
                    f.write(data)
                
                # 转换为PNG
                print(f'[{idx:3d}/{len(ordered_targets)}] {target}', end=' ... ')
                
                if convert_to_png(temp_path, png_path):
                    # 删除临时文件
                    if temp_path.exists():
                        temp_path.unlink()
                    
                    if png_path.exists() and png_path.stat().st_size > 0:
                        saved.append(f"{idx:03d}.png")
                        converted += 1
                        print('✅ PNG')
                    else:
                        print('❌ PNG文件创建失败')
                        failed += 1
                else:
                    # 转换失败，尝试直接复制
                    if temp_path.exists():
                        # 如果转换失败但文件存在，尝试重命名
                        if ext.lower() == '.png':
                            temp_path.rename(png_path)
                            saved.append(f"{idx:03d}.png")
                            converted += 1
                            print('✅ (直接复制)')
                        else:
                            temp_path.unlink()
                            print('❌ 转换失败')
                            failed += 1
                    else:
                        print('❌ 提取失败')
                        failed += 1
                        
            except Exception as e:
                print(f'❌ 错误: {e}')
                failed += 1
        
        print()
        print('=' * 60)
        print('提取完成！')
        print(f'   成功: {converted} 个 (已转换为PNG)')
        if failed > 0:
            print(f'   失败: {failed} 个')
        print('=' * 60)
        print()
        print(f'输出目录: {output_dir}')
        print(f'文件列表:')
        for name in saved:
            print(f'  {name}')

if __name__ == '__main__':
    main()

