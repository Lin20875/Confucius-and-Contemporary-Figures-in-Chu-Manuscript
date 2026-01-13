# -*- coding: utf-8 -*-
"""
从《季庚子問於孔子 廣義讀本 20250228.docx》文件提取文本内容并生成HTML网页
"""
import os
import sys
import re
from zipfile import ZipFile
import xml.etree.ElementTree as ET
from datetime import datetime

# 设置输出编码为UTF-8
if sys.platform == 'win32':
    sys.stdout.reconfigure(encoding='utf-8')

# 导入convert_docx_to_html.py中的函数
from convert_docx_to_html import extract_text_from_docx, extract_footnotes_from_docx, create_html_template

def main():
    docx_path = r"C:\Users\lyue\Desktop\出土文献读本网页\articles\季庚子問於孔子 廣義讀本 20250228.docx"
    output_html = r"C:\Users\lyue\Desktop\出土文献读本网页\articles\季庚子問於孔子.html"
    image_folder = "images_季庚子問於孔子_20260112"
    
    # 实际提取的图片数量（根据extract_docx_images.py的输出）
    actual_image_count = 440
    
    print(f"正在提取DOCX内容: {docx_path}")
    
    # 提取文本内容和脚注引用
    paragraphs, detected_image_count, footnote_refs, main_text_image_count = extract_text_from_docx(docx_path)
    
    # 提取脚注内容
    print("正在提取脚注...")
    footnotes_dict = extract_footnotes_from_docx(docx_path, main_text_image_count)
    
    print(f"提取了 {len(paragraphs)} 个段落")
    print(f"检测到 {detected_image_count} 个图片占位符")
    print(f"检测到 {len(footnote_refs)} 个脚注引用")
    print(f"提取了 {len(footnotes_dict)} 个脚注")
    print(f"使用实际图片数量: {actual_image_count}")
    
    # 使用实际提取的图片数量
    image_count = actual_image_count
    
    # 分离各个部分
    bianlian_paragraphs = []  # 編聯部分
    bianlian_shuoming_paragraphs = []  # 本文編聯説明
    transcription_paragraphs = []  # 釋文部分
    
    current_section = 'bianlian'  # 当前部分：bianlian, shuoming, transcription
    
    for para in paragraphs:
        if para.strip():
            # 检查是否是"本文編聯説明"标题
            if '本文編聯説明' in para.strip():
                current_section = 'shuoming'
                # 提取说明内容（去掉标题）
                content = para.strip().replace('本文編聯説明：', '').replace('本文編聯説明', '').strip()
                if content:
                    bianlian_shuoming_paragraphs.append(content)
                continue
            
            # 检查是否是"釋文"标题
            if para.strip() == '釋文' or para.strip().startswith('釋文'):
                current_section = 'transcription'
                continue
            
            # 检查是否是"本篇竹簡編聯"标题
            if '本篇竹簡編聯' in para.strip():
                current_section = 'bianlian'
                continue
            
            # 检查是否是标题"季庚子問於孔子"
            if para.strip() == '季庚子問於孔子':
                continue
            
            # 根据当前部分添加到相应的列表
            if current_section == 'bianlian':
                bianlian_paragraphs.append(para)
            elif current_section == 'shuoming':
                bianlian_shuoming_paragraphs.append(para)
            elif current_section == 'transcription':
                transcription_paragraphs.append(para)
    
    # 构建内容部分
    content_sections = ''
    
    # 添加"本篇竹簡編聯"章节
    if bianlian_paragraphs:
        content_sections += '''    <section class="doc-section" id="bianlian">
      <h2>本篇竹簡編聯</h2>
      <ul>
'''
        for para in bianlian_paragraphs:
            if para.strip():
                # 转义HTML特殊字符
                para_escaped = para.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
                content_sections += f'        <li>{para_escaped}</li>\n'
        
        content_sections += '''      </ul>
'''
        
        # 添加"本文編聯説明"段落（如果有）
        if bianlian_shuoming_paragraphs:
            for para in bianlian_shuoming_paragraphs:
                if para.strip():
                    para_escaped = para.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
                    content_sections += f'      <p><strong>本文編聯説明</strong>：{para_escaped}</p>\n'
        
        content_sections += '''    </section>

'''
    
    # 添加"釋文"章节
    content_sections += '''    <section class="doc-section" id="transcription">
      <h2 id="transcription-title">釋文</h2>
      <div class="transcription-block">
'''
    
    for para in transcription_paragraphs:
        if para.strip():
            # 转义HTML特殊字符
            para_escaped = para.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
            content_sections += f'        <p>{para_escaped}</p>\n'
    
    content_sections += '''      </div>
    </section>

    <section class="doc-section" id="annotations">
      <h2>注釋</h2>
      <ol class="footnotes">
'''
    
    # 生成注释列表
    sorted_footnote_ids = sorted(footnote_refs, key=lambda x: int(x))
    
    for footnote_id in sorted_footnote_ids:
        footnote_text = footnotes_dict.get(footnote_id, '')
        if footnote_text:
            footnote_text_escaped = footnote_text.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
            content_sections += f'        <li id="fn-{footnote_id}">{footnote_text_escaped}</li>\n'
        else:
            content_sections += f'        <li id="fn-{footnote_id}">注釋內容待補充</li>\n'
    
    if not sorted_footnote_ids:
        content_sections += '        <li id="fn-1">注釋內容待補充</li>\n'
    
    content_sections += '''      </ol>
    </section>
'''
    
    # 创建HTML
    html_content = create_html_template(
        title="季庚子問於孔子",
        subtitle="《季庚子問於孔子 廣義讀本》",
        content_sections=content_sections,
        image_folder=image_folder,
        image_count=image_count
    )
    
    # 写入文件
    with open(output_html, 'w', encoding='utf-8') as f:
        f.write(html_content)
    
    print(f"HTML文件已生成: {output_html}")
    if sorted_footnote_ids:
        print(f"已添加 {len(sorted_footnote_ids)} 条注释")

if __name__ == '__main__':
    main()


