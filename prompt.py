import base64
import zipfile
import io
import re
from xml.etree import ElementTree as ET

NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
}

def get_text_from_paragraph(p):
    # Collect all w:t nodes in the paragraph in order
    texts = []
    for t in p.findall(".//w:t", NS):
        if t.text:
            texts.append(t.text)
    text = "".join(texts)
    # normalise whitespace lightly
    text = re.sub(r"\s+", " ", text).strip()
    return text

def get_paragraph_style(p):
    pPr = p.find("./w:pPr", NS)
    if pPr is None:
        return ""
    pStyle = pPr.find("./w:pStyle", NS)
    if pStyle is None:
        return ""
    return pStyle.attrib.get(f"{{{NS['w']}}}val", "")

def heading_level_from_style(style):
    if not style:
        return 0
    s = style.lower()
    if "heading" not in s:
        return 0
    m = re.search(r"heading\s*([0-9]+)", s)
    if not m:
        return 0
    try:
        return int(m.group(1))
    except:
        return 0

def update_heading_stack(stack, level, text):
    # level is 1 based
    while len(stack) < level:
        stack.append("")
    stack[level - 1] = text
    # truncate deeper levels
    return stack[:level]

def heading_path(stack):
    parts = [s.strip() for s in stack if s and s.strip()]
    return " > ".join(parts)

def parse_table(tbl):
    # Return rows as list of list of cell text
    rows = []
    for tr in tbl.findall("./w:tr", NS):
        row = []
        for tc in tr.findall("./w:tc", NS):
            cell_texts = []
            for p in tc.findall(".//w:p", NS):
                t = get_text_from_paragraph(p)
                if t:
                    cell_texts.append(t)
            cell_text = " ".join(cell_texts)
            cell_text = re.sub(r"\s+", " ", cell_text).strip()
            row.append(cell_text)
        rows.append(row)
    return rows

def load_document_xml_from_docx(binary_data_b64):
    buf = base64.b64decode(binary_data_b64)
    z = zipfile.ZipFile(io.BytesIO(buf))
    xml_bytes = z.read("word/document.xml")
    return xml_bytes.decode("utf-8", errors="ignore")

def docx_to_blocks(doc_xml):
    root = ET.fromstring(doc_xml)
    body = root.find("./w:body", NS)
    if body is None:
        return []

    blocks = []
    hstack = []
    idx = 1

    # iterate direct children of body in order: paragraphs and tables
    for child in list(body):
        tag = child.tag

        if tag == f"{{{NS['w']}}}p":
            style = get_paragraph_style(child)
            text = get_text_from_paragraph(child)

            lvl = heading_level_from_style(style)
            if lvl and text:
                hstack = update_heading_stack(hstack, lvl, text)

            if text:
                blocks.append({
                    "type": "paragraph",
                    "index": idx,
                    "headingPath": heading_path(hstack),
                    "style": style,
                    "text": text
                })
                idx += 1

        elif tag == f"{{{NS['w']}}}tbl":
            rows = parse_table(child)
            blocks.append({
                "type": "table",
                "index": idx,
                "headingPath": heading_path(hstack),
                "rows": rows
            })
            idx += 1

    return blocks

# n8n input
item = items[0]
binary = item.get("binary", {})

if "oldDoc" not in binary:
    raise Exception("Expected binary.oldDoc to exist. Check your previous node standardised the key names.")

old_b64 = binary["oldDoc"]["data"]
old_name = binary["oldDoc"].get("fileName", "")

doc_xml = load_document_xml_from_docx(old_b64)
old_blocks = docx_to_blocks(doc_xml)

return [{
    "json": {
        "oldFileName": old_name,
        "oldBlocks": old_blocks
    },
    "binary": binary
}]
