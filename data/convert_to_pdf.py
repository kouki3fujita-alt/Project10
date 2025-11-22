#!/usr/bin/env python
# -*- coding: utf-8 -*-
import markdown
from bs4 import BeautifulSoup
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import cm
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak, Preformatted
from reportlab.lib import colors
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.enums import TA_LEFT, TA_CENTER, TA_RIGHT
from pathlib import Path
import re

# 日本語フォントの登録（Windows標準フォントを使用）
try:
    # MSゴシックを試す
    pdfmetrics.registerFont(TTFont('Japanese', 'C:/Windows/Fonts/msgothic.ttc', subfontIndex=0))
    japanese_font = 'Japanese'
except:
    try:
        # メイリオを試す
        pdfmetrics.registerFont(TTFont('Japanese', 'C:/Windows/Fonts/meiryo.ttc', subfontIndex=0))
        japanese_font = 'Japanese'
    except:
        # フォールバック（デフォルトフォント）
        japanese_font = 'Helvetica'
        print("警告: 日本語フォントが見つかりません。デフォルトフォントを使用します。")

# Markdownファイルを読み込む
md_file = Path("検証レポート.md")
with open(md_file, "r", encoding="utf-8") as f:
    md_content = f.read()

# MarkdownをHTMLに変換
html_content = markdown.markdown(
    md_content,
    extensions=['tables', 'fenced_code', 'nl2br']
)

# HTMLをパース
soup = BeautifulSoup(html_content, 'html.parser')

# スタイルの設定
styles = getSampleStyleSheet()
styles.add(ParagraphStyle(
    name='JapaneseTitle',
    parent=styles['Heading1'],
    fontName=japanese_font,
    fontSize=24,
    spaceAfter=12,
    textColor=colors.black
))
styles.add(ParagraphStyle(
    name='JapaneseHeading1',
    parent=styles['Heading1'],
    fontName=japanese_font,
    fontSize=18,
    spaceAfter=12,
    textColor=colors.black
))
styles.add(ParagraphStyle(
    name='JapaneseHeading2',
    parent=styles['Heading2'],
    fontName=japanese_font,
    fontSize=14,
    spaceAfter=10,
    textColor=colors.black
))
styles.add(ParagraphStyle(
    name='JapaneseBody',
    parent=styles['BodyText'],
    fontName=japanese_font,
    fontSize=11,
    spaceAfter=6,
    textColor=colors.black,
    leading=16
))

# HTML要素をreportlabの要素に変換する関数
def html_to_reportlab(element, story):
    """HTML要素をreportlabの要素に変換"""
    if element.name is None:  # テキストノード
        text = str(element.string) if element.string else ''
        if text.strip():
            # HTMLエンティティをデコード
            text = text.replace('&lt;', '<').replace('&gt;', '>').replace('&amp;', '&')
            return text
        return ''
    
    tag = element.name.lower()
    
    if tag == 'h1':
        text = element.get_text()
        story.append(Paragraph(text, styles['JapaneseTitle']))
        story.append(Spacer(1, 0.5*cm))
    elif tag == 'h2':
        text = element.get_text()
        story.append(Paragraph(text, styles['JapaneseHeading1']))
        story.append(Spacer(1, 0.3*cm))
    elif tag == 'h3':
        text = element.get_text()
        story.append(Paragraph(text, styles['JapaneseHeading2']))
        story.append(Spacer(1, 0.2*cm))
    elif tag == 'p':
        # 段落内のHTMLタグを処理
        text = ''
        for child in element.children:
            if hasattr(child, 'name'):
                if child.name == 'strong' or child.name == 'b':
                    text += f'<b>{child.get_text()}</b>'
                elif child.name == 'code':
                    text += f'<font name="Courier" size="9">{child.get_text()}</font>'
                elif child.name == 'em' or child.name == 'i':
                    text += f'<i>{child.get_text()}</i>'
                else:
                    text += child.get_text()
            else:
                text += str(child)
        if text.strip():
            # HTMLエンティティをエスケープ
            text = text.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
            # reportlabのタグを戻す
            text = text.replace('&lt;b&gt;', '<b>').replace('&lt;/b&gt;', '</b>')
            text = text.replace('&lt;i&gt;', '<i>').replace('&lt;/i&gt;', '</i>')
            text = text.replace('&lt;font name="Courier" size="9"&gt;', '<font name="Courier" size="9">')
            text = text.replace('&lt;/font&gt;', '</font>')
            story.append(Paragraph(text, styles['JapaneseBody']))
    elif tag == 'ul' or tag == 'ol':
        for li in element.find_all('li', recursive=False):
            text = ''
            for child in li.children:
                if hasattr(child, 'name'):
                    if child.name == 'strong' or child.name == 'b':
                        text += f'<b>{child.get_text()}</b>'
                    elif child.name == 'code':
                        text += f'<font name="Courier" size="9">{child.get_text()}</font>'
                    elif child.name == 'em' or child.name == 'i':
                        text += f'<i>{child.get_text()}</i>'
                    else:
                        text += child.get_text()
                else:
                    text += str(child)
            if text.strip():
                # HTMLエンティティをエスケープ
                text = text.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
                # reportlabのタグを戻す
                text = text.replace('&lt;b&gt;', '<b>').replace('&lt;/b&gt;', '</b>')
                text = text.replace('&lt;i&gt;', '<i>').replace('&lt;/i&gt;', '</i>')
                text = text.replace('&lt;font name="Courier" size="9"&gt;', '<font name="Courier" size="9">')
                text = text.replace('&lt;/font&gt;', '</font>')
                bullet = '• ' if tag == 'ul' else ''
                story.append(Paragraph(f'{bullet}{text}', styles['JapaneseBody']))
        story.append(Spacer(1, 0.2*cm))
    elif tag == 'table':
        # テーブルを処理
        table_data = []
        for tr in element.find_all('tr'):
            row = []
            for td in tr.find_all(['td', 'th']):
                cell_text = td.get_text().strip()
                row.append(cell_text)
            if row:
                table_data.append(row)
        
        if table_data:
            # テーブルスタイル
            table_style = TableStyle([
                ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
                ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
                ('FONTNAME', (0, 0), (-1, 0), japanese_font),
                ('FONTNAME', (0, 1), (-1, -1), japanese_font),
                ('FONTSIZE', (0, 0), (-1, 0), 11),
                ('FONTSIZE', (0, 1), (-1, -1), 10),
                ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
                ('GRID', (0, 0), (-1, -1), 1, colors.black),
                ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ])
            
            # テーブルを作成
            pdf_table = Table(table_data)
            pdf_table.setStyle(table_style)
            story.append(pdf_table)
            story.append(Spacer(1, 0.3*cm))
    elif tag == 'pre' or tag == 'code':
        code_text = element.get_text()
        story.append(Preformatted(code_text, styles['Code']))
        story.append(Spacer(1, 0.3*cm))
    elif tag == 'hr':
        story.append(Spacer(1, 0.5*cm))
    else:
        # その他の要素は子要素を再帰的に処理
        for child in element.children:
            if hasattr(child, 'name'):
                html_to_reportlab(child, story)
            elif str(child).strip():
                text = str(child).strip()
                if text:
                    story.append(Paragraph(text, styles['JapaneseBody']))

# PDFドキュメントの作成
pdf_file = Path("検証レポート.pdf")
doc = SimpleDocTemplate(
    str(pdf_file),
    pagesize=A4,
    rightMargin=2*cm,
    leftMargin=2*cm,
    topMargin=2*cm,
    bottomMargin=2*cm
)

# ストーリー（PDFの内容）を構築
story = []

# HTMLの各要素を処理
for element in soup.children:
    if hasattr(element, 'name'):
        html_to_reportlab(element, story)
    elif str(element).strip():
        text = str(element).strip()
        if text:
            story.append(Paragraph(text, styles['JapaneseBody']))

# PDFを生成
doc.build(story)
print(f"PDFファイルが作成されました: {pdf_file}")
