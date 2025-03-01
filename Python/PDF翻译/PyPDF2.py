import Python.PDF翻译.PyPDF2 as PyPDF2
from googletrans import Translator
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter, A4
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import Paragraph, SimpleDocTemplate, Table, TableStyle, Spacer
from reportlab.lib import colors
import time
import os
import concurrent.futures
import threading
import re
import random
import json
import pickle

# 创建线程本地存储的翻译器
thread_local = threading.local()

def get_translator():
    """获取或创建线程本地的翻译器实例"""
    if not hasattr(thread_local, "translator"):
        thread_local.translator = Translator(service_urls=['translate.google.com'])
    return thread_local.translator

def translate_batch(batch):
    """批量翻译一组单词"""
    results = []
    for word, pos in batch:
        try:
            translation = translate_word(word)
            results.append((word, pos, translation))
            print(f"成功翻译: {word} -> {translation}")
        except Exception as e:
            print(f"批量翻译中单词 {word} 失败: {str(e)}")
            results.append((word, pos, f"[未翻译:{word}]"))
    return results

def translate_word(word, max_retries=3):
    """使用谷歌翻译将单词翻译成中文，带有重试机制"""
    if not word.strip():
        return word
        
    translator = get_translator()
    retry_count = 0
    last_exception = None
    
    while retry_count < max_retries:
        try:
            # 添加随机延迟，避免请求过于频繁
            time.sleep(0.5 + random.random())
            result = translator.translate(word, dest='zh-cn')
            return result.text
        except Exception as e:
            last_exception = e
            retry_count += 1
            print(f"翻译 '{word}' 失败 (尝试 {retry_count}/{max_retries}): {str(e)}")
            if retry_count < max_retries:
                # 重试前重新创建翻译器实例
                thread_local.translator = Translator(service_urls=['translate.google.com'])
                # 指数退避，等待时间随重试次数增加
                time.sleep(2 ** retry_count)
            
    print(f"翻译失败 ({word}): {str(last_exception)}")
    return f"[未翻译:{word}]"  # 返回未翻译标记而不是原词

def save_translations(translations, filename):
    """保存翻译结果到文件"""
    with open(filename, 'wb') as f:
        pickle.dump(translations, f)
    print(f"翻译结果已保存到: {filename}")

def load_translations(filename):
    """从文件加载翻译结果"""
    if os.path.exists(filename):
        with open(filename, 'rb') as f:
            return pickle.load(f)
    return []

def translate_pdf(input_pdf, output_pdf, cache_file=None, max_words=30, only_first_page=True):
    """翻译PDF文件中的单词并生成新的PDF"""
    print("开始处理PDF文件...")
    
    # 如果提供了缓存文件并且存在，则加载已有的翻译结果
    all_translations = []
    if cache_file and os.path.exists(cache_file):
        print(f"从缓存文件加载翻译结果: {cache_file}")
        all_translations = load_translations(cache_file)
        print(f"已加载 {len(all_translations)} 个翻译结果")
        
        # 如果缓存中的翻译结果超过了max_words，则截断
        if max_words and len(all_translations) > max_words:
            all_translations = all_translations[:max_words]
            print(f"截断翻译结果至 {max_words} 个单词")
    
    # 如果没有缓存或缓存为空，则进行翻译
    if not all_translations:
        # 读取PDF文件
        try:
            reader = PyPDF2.PdfReader(input_pdf)
            print(f"成功打开PDF文件: {input_pdf}")
        except Exception as e:
            print(f"打开PDF文件失败: {str(e)}")
            return
        
        # 注册中文字体
        try:
            # 尝试注册多种常见中文字体
            font_paths = [
                '/Library/Fonts/Arial Unicode.ttf',
                '/System/Library/Fonts/AppleGothic.ttf',
                '/System/Library/Fonts/STHeiti Light.ttc',
                '/System/Library/Fonts/STHeiti Medium.ttc',
                '/Library/Fonts/Microsoft/SimHei.ttf',
                '/Library/Fonts/Microsoft/SimSun.ttf',
                '/Library/Fonts/Songti.ttc',
                '/System/Library/Fonts/PingFang.ttc',
                '/System/Library/Fonts/Hiragino Sans GB.ttc',
                '/System/Library/Fonts/Heiti.ttc',
                '/System/Library/Fonts/PingFang.ttc',
            ]
            
            # 尝试使用系统字体目录
            system_font_dirs = [
                '/Library/Fonts/',
                '/System/Library/Fonts/',
                '/Users/henry/Library/Fonts/',
            ]
            
            # 查找所有可能的中文字体
            for font_dir in system_font_dirs:
                if os.path.exists(font_dir):
                    for file in os.listdir(font_dir):
                        if file.endswith(('.ttf', '.ttc', '.otf')) and ('song' in file.lower() or 'hei' in file.lower() or 'ming' in file.lower() or 'gothic' in file.lower() or 'kai' in file.lower()):
                            font_paths.append(os.path.join(font_dir, file))
            
            font_registered = False
            for font_path in font_paths:
                try:
                    if os.path.exists(font_path):
                        font_name = os.path.basename(font_path).split('.')[0].replace(' ', '')
                        pdfmetrics.registerFont(TTFont(font_name, font_path))
                        print(f"成功注册字体: {font_name}")
                        font_registered = True
                        break
                except Exception as e:
                    print(f"注册字体 {font_path} 失败: {str(e)}")
                    
            if not font_registered:
                # 尝试使用reportlab内置的中文字体
                try:
                    from reportlab.pdfbase.cidfonts import UnicodeCIDFont
                    pdfmetrics.registerFont(UnicodeCIDFont('STSong-Light'))
                    print("成功注册内置字体: STSong-Light")
                    font_name = 'STSong-Light'
                    font_registered = True
                except:
                    print("警告: 无法找到合适的中文字体，将使用默认字体")
                    font_name = 'Helvetica'
        except Exception as e:
            print(f"注册字体失败: {str(e)}")
            font_name = 'Helvetica'
        
        # 收集所有页面的文本
        all_text = ""
        # 如果只处理第一页
        if only_first_page:
            page_range = [0]  # 只处理第一页
        else:
            page_range = range(len(reader.pages))
            
        for page_num in page_range:
            print(f"正在处理第 {page_num + 1} 页...")
            page = reader.pages[page_num]
            page_text = page.extract_text()
            all_text += page_text
            
            # 打印前几行文本，用于调试
            if page_num == 0:
                print("页面文本示例（前5行）:")
                lines = page_text.split('\n')
                for i, line in enumerate(lines[:5]):
                    print(f"行 {i+1}: {line}")
        
        # 使用正则表达式提取单词和词性
        # 根据页面文本示例，我们需要匹配每行开头的单词和词性
        # 例如: "abandon v.  B2" 中的 "abandon" 和 "v."
        pattern = r'(\w+)\s+([a-z]+\.\s*(?:,\s*[a-z]+\.\s*)*)\s*([AB][12])'
        matches = re.findall(pattern, all_text, re.MULTILINE)
        
        # 打印匹配结果，用于调试
        print("匹配结果示例（前5个）:")
        for i, match in enumerate(matches[:5]):
            print(f"匹配 {i+1}: 单词={match[0]}, 词性={match[1]}, 难度={match[2]}")
        
        # 收集所有匹配的单词和词性
        all_words = []
        for match in matches:
            word = match[0]
            pos = match[1].strip()
            all_words.append((word, pos))
            # 如果设置了最大单词数，则只处理指定数量的单词
            if max_words and len(all_words) >= max_words:
                break
        
        print(f"找到 {len(all_words)} 个单词")
        
        # 将单词分成多个批次，减小批次大小以提高成功率
        batch_size = 15  # 每批处理15个单词，原来是5个
        batches = [all_words[i:i + batch_size] 
                  for i in range(0, len(all_words), batch_size)]
        
        # 使用线程池并行处理所有批次
        all_translations = []
        with concurrent.futures.ThreadPoolExecutor(max_workers=8) as executor:  # 增加线程数，原来是2个
            # 提交所有批次的翻译任务
            future_to_batch = {executor.submit(translate_batch, batch): i 
                             for i, batch in enumerate(batches)}
            
            # 收集翻译结果
            for future in concurrent.futures.as_completed(future_to_batch):
                batch_index = future_to_batch[future]
                try:
                    results = future.result()
                    all_translations.extend(results)
                    print(f"完成批次 {batch_index + 1}/{len(batches)}")
                    
                    # 每完成10个批次保存一次中间结果
                    if batch_index % 10 == 0 and batch_index > 0 and cache_file:
                        save_translations(all_translations, cache_file)
                        
                except Exception as e:
                    print(f"批次 {batch_index + 1} 处理失败: {str(e)}")
        
        # 保存最终翻译结果
        if cache_file:
            save_translations(all_translations, cache_file)
    
    # 创建一个新的PDF文档
    doc = SimpleDocTemplate(output_pdf, pagesize=A4)
    
    # 创建段落样式
    styles = getSampleStyleSheet()
    normal_style = styles['Normal']
    normal_style.fontName = font_name
    normal_style.fontSize = 12
    normal_style.leading = 16  # 行间距
    
    # 创建文档元素列表
    elements = []
    
    # 添加标题
    title_style = styles['Title']
    title_style.fontName = font_name
    elements.append(Paragraph("牛津3000核心词汇表（带中文翻译）", title_style))
    elements.append(Spacer(1, 20))  # 添加空间
    
    # 按三列排版
    columns_per_row = 3
    column_width = (A4[0] - 40) / columns_per_row  # 页面宽度减去左右边距
    
    # 创建三列布局
    print("开始生成PDF...")
    
    # 将翻译结果分成三列
    column1 = []
    column2 = []
    column3 = []
    
    for i, (word, pos, translation) in enumerate(all_translations):
        # 创建一个条目：单词 词性 中文翻译
        # 确保格式为 "about prep., adv. 关于"
        # 清理词性中的多余空格
        pos = pos.strip()
        entry = f"<b>{word}</b> {pos} {translation}"
        
        # 根据索引分配到不同列
        if i % 3 == 0:
            column1.append(Paragraph(entry, normal_style))
        elif i % 3 == 1:
            column2.append(Paragraph(entry, normal_style))
        else:
            column3.append(Paragraph(entry, normal_style))
    
    # 确定最长的列
    max_length = max(len(column1), len(column2), len(column3))
    
    # 创建表格数据，每行包含三列
    table_data = []
    for i in range(max_length):
        row = []
        # 第一列
        if i < len(column1):
            row.append(column1[i])
        else:
            row.append("")
        
        # 第二列
        if i < len(column2):
            row.append(column2[i])
        else:
            row.append("")
        
        # 第三列
        if i < len(column3):
            row.append(column3[i])
        else:
            row.append("")
        
        table_data.append(row)
    
    # 创建表格，但不显示边框
    col_widths = [column_width] * columns_per_row
    table = Table(table_data, colWidths=col_widths)
    
    # 设置表格样式，不显示边框
    table.setStyle(TableStyle([
        ('FONT', (0, 0), (-1, -1), font_name, 12),
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ('LEFTPADDING', (0, 0), (-1, -1), 5),
        ('RIGHTPADDING', (0, 0), (-1, -1), 5),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 8),
        ('TOPPADDING', (0, 0), (-1, -1), 8),
        ('ALIGNMENT', (0, 0), (-1, -1), 'LEFT'),
    ]))
    
    elements.append(table)
    
    # 构建文档
    doc.build(elements)
    
    print(f"翻译完成,已保存到: {output_pdf}")

# 使用示例
input_pdf = "/Users/henry/Desktop/The_Oxford_3000.pdf"
# 添加时间戳到输出文件名
timestamp = time.strftime("%Y%m%d_%H%M%S")
output_pdf = f"/Users/henry/Desktop/output_with_translations_{timestamp}.pdf"
cache_file = f"/Users/henry/Desktop/translations_cache_{timestamp}.pkl"

# 直接开始翻译过程，翻译第一页的前30个单词
translate_pdf(input_pdf, output_pdf, cache_file, max_words=30000, only_first_page=False)
 