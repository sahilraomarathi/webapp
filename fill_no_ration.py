#!/usr/bin/env python3
"""
Fill No_RationCard.docx template and save as .docx (no PDF conversion).
Usage: python3 fill_no_ration.py <json_input> <output_docx_path>
"""
import sys
import json
import os
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

def fill_no_ration(data, template_path, out_docx):
    doc = Document(template_path)

    title       = data['applicantTitle']
    full_name   = data['applicantFullName']
    father_name = data['fatherHusbandName']
    age         = data['age']
    dob         = data['dob']
    address     = data['address']
    today_date  = data['todayDate']

    old_p6 = "Mr. Pradeep Sampat Katkar son/Daughter/wife of Sampat Katkar, Age 34 Years resident of At Nangalwadi Phata Mahad Dist Raigad M.I.D.C 402301"
    new_p6 = f"{title}. {full_name} son/Daughter/wife of {father_name}, Age {age} Years resident of At {address}"
    replace_in_doc(doc, old_p6, new_p6)

    replace_in_doc(doc, "22/11/1992", dob)
    replace_in_doc(doc, "25/05/26", today_date)
    replace_in_doc(doc, "Mr. Pradeep Sampat Katkar", f"{title}. {full_name}")

    doc.save(out_docx)

if __name__ == "__main__":
    data = json.loads(sys.argv[1])
    out_docx = sys.argv[2]
    template = os.path.join(os.path.dirname(__file__), "templates", "No_RationCard.docx")
    fill_no_ration(data, template, out_docx)
    print("OK")
