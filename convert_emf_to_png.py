# -*- coding: utf-8 -*-
"""
将EMF格式图片转换为PNG格式
使用Windows PowerShell + .NET
"""
import os
import sys
import subprocess
from pathlib import Path

def convert_emf_to_png(emf_path, png_path):
    """使用PowerShell和.NET转换EMF到PNG"""
    # 转义路径中的反斜杠
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
    except subprocess.TimeoutExpired:
        print("  ⚠️  转换超时")
        return False
    except Exception as e:
        print(f"  ⚠️  执行错误: {e}")
        return False

def main():
    image_dir = Path(r'C:\Users\lyue\Desktop\网页\articles\images_君子為禮_20251201')
    
    if not image_dir.exists():
        print(f'❌ 目录不存在: {image_dir}')
        sys.exit(1)
    
    # 查找所有EMF文件
    emf_files = list(image_dir.glob('*.emf'))
    
    if not emf_files:
        print('✅ 没有找到EMF文件，无需转换')
        return
    
    print(f'找到 {len(emf_files)} 个EMF文件需要转换')
    print('使用 Windows PowerShell + .NET 进行转换...')
    print()
    
    converted = 0
    failed = 0
    skipped = 0
    
    for emf_file in sorted(emf_files):
        png_file = emf_file.with_suffix('.png')
        
        # 检查PNG是否已存在
        if png_file.exists():
            print(f'跳过: {emf_file.name} (PNG已存在)')
            skipped += 1
            continue
        
        print(f'转换: {emf_file.name} -> {png_file.name}', end=' ... ', flush=True)
        
        try:
            if convert_emf_to_png(emf_file, png_file):
                # 验证PNG文件是否创建成功
                if png_file.exists() and png_file.stat().st_size > 0:
                    # 删除原EMF文件
                    try:
                        emf_file.unlink()
                        print('✅')
                        converted += 1
                    except Exception as e:
                        print(f'✅ (但无法删除原EMF文件: {e})')
                        converted += 1
                else:
                    print('❌ PNG文件创建失败')
                    failed += 1
            else:
                print('❌ 转换失败')
                failed += 1
        except Exception as e:
            print(f'❌ 错误: {e}')
            failed += 1
    
    print()
    print('=' * 50)
    print(f'转换完成！')
    print(f'   成功: {converted} 个')
    if skipped > 0:
        print(f'   跳过: {skipped} 个 (PNG已存在)')
    if failed > 0:
        print(f'   失败: {failed} 个')
    print('=' * 50)
    
    if failed > 0:
        print()
        print('如果转换失败，可以尝试：')
        print('  1. 安装 ImageMagick: https://imagemagick.org/script/download.php')
        print('  2. 然后运行: pip install Wand')
        print('  3. 重新运行此脚本')

if __name__ == '__main__':
    main()
