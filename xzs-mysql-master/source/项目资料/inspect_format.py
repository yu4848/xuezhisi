from docx import Document
import json, os

base = os.path.dirname(os.path.abspath(__file__))
fmt_path = os.path.join(base, '论文格式.docx')

doc = Document(fmt_path)

# 列出所有命名样式
print("=== 文档样式列表 ===")
for style in doc.styles:
    if style.type.name in ('PARAGRAPH', 'CHARACTER'):
        print(f"  [{style.type.name}] {style.name}")

print("\n=== 段落详情 ===")
for i, para in enumerate(doc.paragraphs[:100]):
    if not para.text.strip():
        continue
    pf = para.paragraph_format
    run = para.runs[0] if para.runs else None
    print(json.dumps({
        'idx': i,
        'style': para.style.name,
        'text': para.text[:50],
        'alignment': str(para.alignment),
        'space_before': str(pf.space_before),
        'space_after': str(pf.space_after),
        'line_spacing': str(pf.line_spacing),
        'line_spacing_rule': str(pf.line_spacing_rule),
        'first_line_indent': str(pf.first_line_indent),
        'left_indent': str(pf.left_indent),
        'run_bold': run.bold if run else None,
        'run_font': run.font.name if run else None,
        'run_size': str(run.font.size) if run else None,
        'run_color': str(run.font.color.rgb) if run and run.font.color and run.font.color.type else None,
    }, ensure_ascii=False))
