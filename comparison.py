import streamlit as st
from docx import Document
from docx.shared import RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from io import BytesIO
from difflib import SequenceMatcher, get_close_matches
import difflib , re
from zipfile import ZipFile
from lxml import etree

# === UTILS ===
# Add a paragraph to the output document with a colored label and content
def add_colored_paragraph(doc, label, content, color):
    p = doc.add_paragraph()
    run = p.add_run(f"{label} {content}")
    run.font.color.rgb = RGBColor(*color)

def get_alignment(para):
    return {
        WD_ALIGN_PARAGRAPH.LEFT: "LEFT",
        WD_ALIGN_PARAGRAPH.CENTER: "CENTER",
        WD_ALIGN_PARAGRAPH.RIGHT: "RIGHT",
        WD_ALIGN_PARAGRAPH.JUSTIFY: "JUSTIFY",
        None: "UNKNOWN"
    }.get(para.alignment, "UNKNOWN")

def get_indent(para):
    l = para.paragraph_format.left_indent
    r = para.paragraph_format.right_indent
    return (round(l.cm, 2) if l else 0.0, round(r.cm, 2) if r else 0.0)

def get_line_spacing(para):
    return round(para.paragraph_format.line_spacing, 2) if para.paragraph_format.line_spacing else "Default"

def get_paragraph_info(para):
    run = para.runs[0] if para.runs else None
    return {
        "text": para.text.strip(),
        "font_name": run.font.name if run and run.font.name else "Default",
        "font_size": run.font.size.pt if run and run.font.size else "Default",
        "bold": run.bold if run else False,
        "italic": run.italic if run else False,
        "alignment": get_alignment(para),
        "spacing": get_line_spacing(para),
        "left_indent": get_indent(para)[0],
        "right_indent": get_indent(para)[1]
    }

def get_word_diff(old, new):
    old_words, new_words = old.split(), new.split()
    matcher = difflib.SequenceMatcher(None, old_words, new_words)
    diffs = []
    for tag, i1, i2, j1, j2 in matcher.get_opcodes():
        if tag in ['replace', 'insert', 'delete']:
            ow = " ".join(old_words[i1:i2])
            nw = " ".join(new_words[j1:j2])
            if ow != nw:
                diffs.append(f"[Word Changed] {ow} → {nw}")
    return diffs

def get_cell_style(cell):
    style = {
        "text": cell.text.strip(),
        "font_name": "Default",
        "font_size": "Default",
        "bold": False,
        "italic": False,
        "alignment": "UNKNOWN",
        "spacing": "Default",
        "left_indent": 0.0,
        "right_indent": 0.0,
        "letter_spacing": "Default"
    }
    if cell.paragraphs:
        para = cell.paragraphs[0]
        run = para.runs[0] if para.runs else None
        pf = para.paragraph_format
        if run:
            spacing_val = get_letter_spacing_from_run(run)
            style.update({
                "font_name": run.font.name or "Default",
                "font_size": run.font.size.pt if run.font.size else "Default",
                "bold": run.bold or False,
                "italic": run.italic or False,
                "alignment": get_alignment(para),
                "spacing": round(pf.line_spacing, 2) if pf.line_spacing else "Default",
                "left_indent": round(pf.left_indent.cm, 2) if pf.left_indent else 0.0,
                "right_indent": round(pf.right_indent.cm, 2) if pf.right_indent else 0.0,
                "letter_spacing": spacing_val
            })
    return style

def get_letter_spacing_from_run(run):
    try:
        xml = run._element
        spacing_elem = xml.find('.//w:spacing', namespaces={'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'})
        if spacing_elem is not None and 'w:val' in spacing_elem.attrib:
            return spacing_elem.attrib['w:val']
    except:
        pass
    return "Default"

def detect_spacing_issues(old, new):
    issues = []
    if re.search(r"\s{2,}", old) and not re.search(r"\s{2,}", new):
        issues.append(f"[Extra Spaces Detected] \"{old}\" → \"{new}\"")
    if re.sub(r'\s+', '', old) == re.sub(r'\s+', '', new) and old != new:
        issues.append(f"[Letter Spacing Issue] \"{old}\" → \"{new}\"")
    return issues

def fuzzy_compare_lists(list1, list2, label, output_doc):
    matcher = difflib.SequenceMatcher(None, list1, list2)
    table_count = 0
    for tag, i1, i2, j1, j2 in matcher.get_opcodes():
        if tag == 'equal':
            continue
        elif tag == 'replace':
            for i, j in zip(range(i1, i2), range(j1, j2)):
                preview = list2[j][:80] + ("..." if len(list2[j]) > 80 else "")
                heading_label = f"{label} {j+1}: \"{preview}\""
                if label == "TableRow":
                    table_count += 1
                    heading_label = f"Table {table_count} | Preview: \"{preview}\""
                output_doc.add_paragraph(heading_label, style="Heading 3")
                add_colored_paragraph(output_doc, f"[Old {label} {i+1}]", list1[i], (255, 0, 0))
                add_colored_paragraph(output_doc, f"[New {label} {j+1}]", list2[j], (0, 128, 0))
                for diff in get_word_diff(list1[i], list2[j]) + detect_spacing_issues(list1[i], list2[j]):
                    add_colored_paragraph(output_doc, "", diff, (255, 165, 0))
        elif tag == 'delete':
            for i in range(i1, i2):
                output_doc.add_paragraph(f"Removed {label} (Line {i+1})", style="Heading 3")
                add_colored_paragraph(output_doc, f"[Removed {label}]", list1[i], (255, 0, 0))
        elif tag == 'insert':
            for j in range(j1, j2):
                preview = list2[j][:80] + ("..." if len(list2[j]) > 80 else "")
                heading_label = f"Added {label} (Line {j+1})"
                if label == "TableRow":
                    table_count += 1
                    heading_label = f"Table {table_count} | Preview: \"{preview}\""
                output_doc.add_paragraph(heading_label, style="Heading 3")
                add_colored_paragraph(output_doc, f"[Added {label}]", list2[j], (0, 128, 0))

def estimate_paragraph_pages(doc):
    avg_paragraphs_per_page = 25
    para_map = []
    for i, para in enumerate(doc.paragraphs):
        page = (i // avg_paragraphs_per_page) + 1
        para_map.append(page)
    return para_map

def compare_paragraphs(doc1, doc2, output_doc):
    output_doc.add_heading("Paragraph Comparison", level=1)
    paras1 = [p.text.strip() for p in doc1.paragraphs if p.text.strip()]
    paras2 = [p.text.strip() for p in doc2.paragraphs if p.text.strip()]
    page_map = estimate_paragraph_pages(doc2)
    matcher = difflib.SequenceMatcher(None, paras1, paras2)
    for tag, i1, i2, j1, j2 in matcher.get_opcodes():
        if tag == 'equal':
            continue
        elif tag == 'replace':
            for i, j in zip(range(i1, i2), range(j1, j2)):
                preview = paras2[j][:80] + ("..." if len(paras2[j]) > 80 else "")
                page = page_map[j] if j < len(page_map) else "?"
                output_doc.add_paragraph(f"Paragraph {j+1} | Preview: \"{preview}\"", style="Heading 3")
                add_colored_paragraph(output_doc, f"[Old Paragraph]", paras1[i], (255, 0, 0))
                add_colored_paragraph(output_doc, f"[New Paragraph]", paras2[j], (0, 128, 0))
                for diff in get_word_diff(paras1[i], paras2[j]) + detect_spacing_issues(paras1[i], paras2[j]):
                    add_colored_paragraph(output_doc, "", diff, (255, 165, 0))
                
                # Compare formatting
                pre_info = get_paragraph_info(doc1.paragraphs[i])
                post_info = get_paragraph_info(doc2.paragraphs[j])
                for key in ['font_name','font_size','bold','italic','alignment','spacing','left_indent','right_indent']:
                    if pre_info[key] != post_info[key]:
                        add_colored_paragraph(output_doc, f"[{key.replace('_',' ').title()} Changed]", f"{pre_info[key]} → {post_info[key]}", (0, 0, 255))
        elif tag == 'delete':
            for i in range(i1, i2):
                output_doc.add_paragraph(f"Removed Paragraph (Line {i+1})", style="Heading 3")
                add_colored_paragraph(output_doc, f"[Removed Paragraph]", paras1[i], (255, 0, 0))
        elif tag == 'insert':
            for j in range(j1, j2):
                preview = paras2[j][:80] + ("..." if len(paras2[j]) > 80 else "")
                page = page_map[j] if j < len(page_map) else "?"
                output_doc.add_paragraph(f"Added Paragraph {j+1} | Preview: \"{preview}\"", style="Heading 3")
                add_colored_paragraph(output_doc, f"[Added Paragraph]", paras2[j], (0, 128, 0))

def extract_textbox_paragraphs_with_pages(docx_path):
    paragraphs_with_pages = []
    try:
        with ZipFile(docx_path) as docx:
            xml_content = docx.read("word/document.xml")
            tree = etree.fromstring(xml_content)
            namespaces = {
                'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
                'wps': 'http://schemas.microsoft.com/office/word/2010/wordprocessingShape'
            }
            page_number = 1
            count = 0
            for para in tree.xpath('.//w:drawing//wps:txbx//w:p', namespaces=namespaces):
                texts = para.xpath('.//w:t', namespaces=namespaces)
                text_content = "".join([t.text for t in texts if t.text])
                if text_content.strip():
                    count += 1
                    page_number = (count // 5) + 1  # assume 5 textboxes per page avg
                    paragraphs_with_pages.append((text_content.strip(), page_number))
    except:
        pass
    return paragraphs_with_pages

def compare_textboxes(doc1, doc2, output_doc):
    output_doc.add_heading("Text Box Comparison", level=1)
    tb1 = extract_textbox_paragraphs_with_pages("Pre Results\MergeDataLetter_old.docx")
    tb2 = extract_textbox_paragraphs_with_pages("Post Results\MergeDataLetter_new.docx")
    texts1 = [t[0] for t in tb1]
    texts2 = [t[0] for t in tb2]
    pages2 = [t[1] for t in tb2]
    matcher = difflib.SequenceMatcher(None, texts1, texts2)
    for tag, i1, i2, j1, j2 in matcher.get_opcodes():
        if tag == 'equal':
            continue
        elif tag == 'replace':
            for i, j in zip(range(i1, i2), range(j1, j2)):
                preview = texts2[j][:80] + ("..." if len(texts2[j]) > 80 else "")
                page = pages2[j] if j < len(pages2) else "?"
                output_doc.add_paragraph(f"TextBox {j+1} | Page {page} | Preview: \"{preview}\"", style="Heading 3")
                add_colored_paragraph(output_doc, "[Old]", texts1[i], (255, 0, 0))
                add_colored_paragraph(output_doc, "[New]", texts2[j], (0, 128, 0))
                for diff in get_word_diff(texts1[i], texts2[j]) + detect_spacing_issues(texts1[i], texts2[j]):
                    add_colored_paragraph(output_doc, "", diff, (255, 165, 0))

                # Compare formatting if available via doc1/2 (fallback to first match)
                para1 = next((p for p in doc1.paragraphs if p.text.strip() == texts1[i]), None)
                para2 = next((p for p in doc2.paragraphs if p.text.strip() == texts2[j]), None)
                if para1 and para2:
                    pre_info = get_paragraph_info(para1)
                    post_info = get_paragraph_info(para2)
                    for key in ['font_name','font_size','bold','italic','alignment','spacing','left_indent','right_indent']:
                        if pre_info[key] != post_info[key]:
                            add_colored_paragraph(output_doc, f"[{key.replace('_',' ').title()} Changed]", f"{pre_info[key]} → {post_info[key]}", (0, 0, 255))
        elif tag == 'delete':
            for i in range(i1, i2):
                output_doc.add_paragraph(f"TextBox Removed (Line {i+1})", style="Heading 3")
                add_colored_paragraph(output_doc, "[Removed]", texts1[i], (255, 0, 0))
        elif tag == 'insert':
            for j in range(j1, j2):
                preview = texts2[j][:80] + ("..." if len(texts2[j]) > 80 else "")
                page = pages2[j] if j < len(pages2) else "?"
                output_doc.add_paragraph(f"TextBox Added {j+1} | Page {page} | Preview: \"{preview}\"", style="Heading 3")
                add_colored_paragraph(output_doc, "[Added]", texts2[j], (0, 128, 0))

def compare_headers_footers(doc1, doc2, output_doc):
    output_doc.add_heading("Header/Footer Comparison", level=1)
    
    def get_section_paragraphs(doc, attr):
        items = []
        for idx, sec in enumerate(doc.sections):
            paras = getattr(sec, attr).paragraphs
            for p in paras:
                if p.text.strip():
                    items.append((idx + 1, p.text.strip()))
        return items

    headers1 = get_section_paragraphs(doc1, 'header')
    headers2 = get_section_paragraphs(doc2, 'header')
    footers1 = get_section_paragraphs(doc1, 'footer')
    footers2 = get_section_paragraphs(doc2, 'footer')

    def compare_parts(part1, part2, label):
        texts1 = [t[1] for t in part1]
        texts2 = [t[1] for t in part2]
        matcher = difflib.SequenceMatcher(None, texts1, texts2)
        for tag, i1, i2, j1, j2 in matcher.get_opcodes():
            if tag == 'equal':
                continue
            elif tag == 'replace':
                for i, j in zip(range(i1, i2), range(j1, j2)):
                    sec1 = part1[i][0]
                    sec2 = part2[j][0]
                    preview = texts2[j][:80] + ("..." if len(texts2[j]) > 80 else "")
                    output_doc.add_paragraph(f"{label} - Section {sec2} | Preview: \"{preview}\"", style="Heading 3")
                    add_colored_paragraph(output_doc, f"[Old {label}]", texts1[i], (255, 0, 0))
                    add_colored_paragraph(output_doc, f"[New {label}]", texts2[j], (0, 128, 0))
                    for diff in get_word_diff(texts1[i], texts2[j]) + detect_spacing_issues(texts1[i], texts2[j]):
                        add_colored_paragraph(output_doc, "", diff, (255, 165, 0))
            elif tag == 'delete':
                removed = {}
                for i in range(i1, i2):
                    txt = texts1[i]
                    sec = part1[i][0]
                    removed.setdefault(txt, []).append(str(sec))
                for txt, secs in removed.items():
                    output_doc.add_paragraph(
    f"[Removed {label}] {txt}→ Removed from Sections: {', '.join(secs)}",
    style="Heading 3"
)
            elif tag == 'insert':
                for j in range(j1, j2):
                    sec = part2[j][0]
                    preview = texts2[j][:80] + ("..." if len(texts2[j]) > 80 else "")
                    output_doc.add_paragraph(f"Added {label} - Section {sec} | Preview: \"{preview}\"", style="Heading 3")
                    add_colored_paragraph(output_doc, f"[Added {label}]", texts2[j], (0, 128, 0))

    compare_parts(headers1, headers2, "Header")
    compare_parts(footers1, footers2, "Footer")

def estimate_pages_with_breaks(doc):
    page_number = 1
    block_page_map = []
    for block in doc.element.body.iter():
        if block.tag.endswith("}br") and block.attrib.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}type') == 'page':
            page_number += 1
        elif block.tag.endswith("}p"):
            block_page_map.append(("paragraph", page_number))
        elif block.tag.endswith("}tbl"):
            block_page_map.append(("table", page_number))
    return block_page_map

def get_table_estimated_page(doc, table_index):
    pages = estimate_pages_with_breaks(doc)
    table_count = -1
    for block_type, page in pages:
        if block_type == "table":
            table_count += 1
            if table_count == table_index:
                return page
    return 1

def compare_tables(doc1, doc2, output_doc):
    output_doc.add_heading("Table Comparison", level=1)
    for t_index, (t1, t2) in enumerate(zip(doc1.tables, doc2.tables)):
        page_est = get_table_estimated_page(doc2, t_index)
        preview = t2.cell(0, 0).text.strip()[:80] if t2.rows and t2.rows[0].cells else ""
        has_changes = False
        changes_buffer = []
        for r_idx, (r1, r2) in enumerate(zip(t1.rows, t2.rows)):
            for c_idx, (c1, c2) in enumerate(zip(r1.cells, r2.cells)):
                text1 = c1.text.strip()
                text2 = c2.text.strip()
                if text1 != text2:
                    has_changes = True
                    changes_buffer.append((r_idx, c_idx, text1, text2))
        if not has_changes:
            continue
        output_doc.add_paragraph(f"Table {t_index+1} | Page {page_est} | Preview: \"{preview}\"", style="Heading 3")

        for r_idx, (r1, r2) in enumerate(zip(t1.rows, t2.rows)):
            for c_idx, (c1, c2) in enumerate(zip(r1.cells, r2.cells)):
                text1 = c1.text.strip()
                text2 = c2.text.strip()
                if text1 != text2:
                    add_colored_paragraph(output_doc, f"[Cell ({r_idx+1},{c_idx+1}) Old]", text1, (255, 0, 0))
                    add_colored_paragraph(output_doc, f"[Cell ({r_idx+1},{c_idx+1}) New]", text2, (0, 128, 0))
                    for diff in get_word_diff(text1, text2) + detect_spacing_issues(text1, text2):
                        add_colored_paragraph(output_doc, "", diff, (255, 165, 0))



def compare_images(doc1, doc2, output_doc):
    output_doc.add_heading("Image Comparison", level=1)
    imgs1 = [(img.width, img.height) for img in doc1.inline_shapes]
    imgs2 = [(img.width, img.height) for img in doc2.inline_shapes]
    matcher = difflib.SequenceMatcher(None, imgs1, imgs2)
    for tag, i1, i2, j1, j2 in matcher.get_opcodes():
        if tag == 'equal':
            continue
        if tag in ('replace', 'delete'):
            for i in range(i1, i2):
                if i < len(imgs1):
                    add_colored_paragraph(output_doc, f"[Removed Image {i+1}]", f"Size: {imgs1[i]}", (255, 0, 0))
        if tag in ('replace', 'insert'):
            for j in range(j1, j2):
                if j < len(imgs2):
                    add_colored_paragraph(output_doc, f"[Added Image {j+1}]", f"Size: {imgs2[j]}", (0, 128, 0))

# ====================== SREAMLIT UI ================================
from comparison import compare_paragraphs, compare_textboxes, compare_headers_footers, compare_tables, compare_images

st.set_page_config(page_title="Word File Comparator", layout="centered")
st.title("📄 Word Document Comparator")
st.markdown("Upload your **Pre** and **Post** Word documents to compare changes like **text, formatting, spacing, fonts, images, tables, headers/footers, etc.**")

# Uploaders
st.subheader("📤 Upload Pre Document (.docx)")
uploaded_pre = st.file_uploader("Upload the PRE document", type=["docx"], key="pre")

st.subheader("📥 Upload Post Document (.docx)")
uploaded_post = st.file_uploader("Upload the POST document", type=["docx"], key="post")

if uploaded_pre and uploaded_post:
    if st.button("🔍 Compare Documents"):
        with st.spinner("Comparing documents, please wait..."):
            doc1 = Document(uploaded_pre)
            doc2 = Document(uploaded_post)
            output_doc = Document()

            compare_paragraphs(doc1, doc2, output_doc)
            compare_textboxes(doc1, doc2, output_doc)
            compare_headers_footers(doc1, doc2, output_doc)
            compare_tables(doc1, doc2, output_doc)
            compare_images(doc1, doc2, output_doc)

            buffer = BytesIO()
            output_doc.save(buffer)
            buffer.seek(0)

            pre_name = uploaded_pre.name.rsplit('.', 1)[0]
            post_name = uploaded_post.name.rsplit('.', 1)[0]
            dynamic_filename = f"Comparison_{pre_name}_vs_{post_name}.docx"

        st.success("✅ Comparison Complete!")

        # Optional: Summary Section placeholder
        st.markdown("### 📊 Summary of Changes")
        st.write("- Paragraphs compared ✅")
        st.write("- Tables compared ✅")
        st.write("- Images compared ✅")
        st.write("- Text boxes compared ✅")
        st.write("- Headers/Footers compared ✅")

        # Download button
        st.download_button(
            "⬇️ Download Comparison Report",
            data=buffer.getvalue(),
            file_name=dynamic_filename,
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
else:
    st.info("📌 Please upload both **Pre** and **Post** documents to begin comparison.")
