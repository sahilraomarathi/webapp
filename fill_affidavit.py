#!/usr/bin/env python3
"""
Fill Affidavit_For_Change_Of_Name.docx, save as .docx or convert to PDF via pandoc.
Usage: python3 fill_affidavit.py <json_input> <output_path> <format: docx|pdf>
"""
import sys
import json
import os
import tempfile
import subprocess
from datetime import datetime
from docx import Document

def replace_in_paragraph(para, old, new):
    full_text = ''.join(run.text for run in para.runs)
    if old not in full_text:
        return False
    new_text = full_text.replace(old, new)
    if para.runs:
        para.runs[0].text = new_text
        for run in para.runs[1:]:
            run.text = ''
    return True

def replace_in_doc(doc, old, new):
    for para in doc.paragraphs:
        replace_in_paragraph(para, old, new)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    replace_in_paragraph(para, old, new)

def get_day_name(date_str):
    parts = date_str.split('/')
    if len(parts[2]) == 2:
        parts[2] = '20' + parts[2]
    dt = datetime.strptime(f"{parts[0]}/{parts[1]}/{parts[2]}", "%d/%m/%Y")
    return dt.strftime("%A")

def convert_to_pdf(docx_path, pdf_path):
    """Convert docx to PDF using pandoc + wkhtmltopdf (no root/LibreOffice needed)."""
    result = subprocess.run(
        ['pandoc', docx_path, '-o', pdf_path,
         '--pdf-engine=wkhtmltopdf',
         '--metadata', 'title=Document'],
        capture_output=True, text=True, timeout=45
    )
    if not os.path.exists(pdf_path):
        raise RuntimeError(f"PDF conversion failed: {result.stderr}")

def fill_affidavit(data, template_path, out_path, fmt='docx'):
    doc = Document(template_path)

    mag_date   = data['magistrateDate']
    title      = data['applicantTitle']
    full_name  = data['applicantFullName']
    age        = data['applicantAge']
    occupation = data['occupation']
    address    = data['applicantAddress']
    relation   = data['deceasedRelation']
    dec_title  = data['deceasedTitle']
    dec_name   = data['deceasedFullName']
    sv_number  = data['svNumber']
    sv_date    = data['svDate']
    exp_date   = data['expiredDate']
    exp_place  = data['expiredPlace']
    today_date = data['todayDate']

    day_name = get_day_name(today_date)

    old_app_block = "Mr/Mrs/Miss. Gajanan Damodar Mahadik, Age-74 years, Occupation- Farmer, Resident Of Adhi Dongroli Mahad Raigad 402103"
    new_app_block = f"{title}. {full_name}, Age-{age} years, Occupation- {occupation}, Resident Of {address}"

    old_dec_full_paren = "My Wife (Mrs. Saraswati Gajanan Mahadik)"
    new_dec_full_paren = f"My {relation} ({dec_title}. {dec_name})"

    old_dec_sv = "My Wife Mrs. Saraswati Gajanan Mahadik has expired on 01/10/2025 at Dongroli (Adhi)"
    new_dec_sv = f"My {relation} {dec_title}. {dec_name} has expired on {exp_date} at {exp_place}"

    old_today_day = "Date:- 25/05/26 Day:- Monday"
    new_today_day = f"Date:- {today_date} Day:- {day_name}"

    replace_in_doc(doc, old_app_block, new_app_block)
    replace_in_doc(doc, old_dec_full_paren, new_dec_full_paren)
    replace_in_doc(doc, old_dec_sv, new_dec_sv)
    replace_in_doc(doc, "1017214", sv_number)
    replace_in_doc(doc, "11/12/2018", sv_date)
    replace_in_doc(doc, "09/04/2026", mag_date)
    replace_in_doc(doc, old_today_day, new_today_day)
    replace_in_doc(doc, "Mrs. Saraswati Gajanan Mahadik", f"{dec_title}. {dec_name}")
    replace_in_doc(doc, "Mr. Gajanan Damodar Mahadik", f"{title}. {full_name}")
    replace_in_doc(doc, "25/05/26", today_date)

    if fmt == 'pdf':
        with tempfile.NamedTemporaryFile(suffix='.docx', delete=False) as tmp:
            tmp_docx = tmp.name
        doc.save(tmp_docx)
        try:
            convert_to_pdf(tmp_docx, out_path)
        finally:
            os.unlink(tmp_docx)
    else:
        doc.save(out_path)

if __name__ == "__main__":
    data = json.loads(sys.argv[1])
    out_path = sys.argv[2]
    fmt = sys.argv[3] if len(sys.argv) > 3 else 'docx'
    template = os.path.join(os.path.dirname(__file__), "templates", "Affidavit_For_Change_Of_Name.docx")
    fill_affidavit(data, template, out_path, fmt)
    print("OK")
