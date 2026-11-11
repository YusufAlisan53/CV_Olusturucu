import argparse
import json
import os
import subprocess
import sys
from pathlib import Path

import docx
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING, WD_TAB_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor


FONT_NAME = "Times New Roman"
BLACK = RGBColor(0, 0, 0)
RIGHT_TAB_POS = Inches(7.5)  # 8.5" page width - 0.5" left margin - 0.5" right margin


def set_run_font(run, size_pt, bold=False, italic=False, small_caps=False, underline=False):
    """Apply consistent Times New Roman styling (including EastAsia/CS for Turkish chars) to a run."""
    run.font.name = FONT_NAME
    run.font.size = Pt(size_pt)
    run.font.bold = bold
    run.font.italic = italic
    run.font.small_caps = small_caps
    run.font.underline = underline
    run.font.color.rgb = BLACK

    rPr = run._r.get_or_add_rPr()
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rPr.insert(0, rFonts)
    rFonts.set(qn("w:ascii"), FONT_NAME)
    rFonts.set(qn("w:hAnsi"), FONT_NAME)
    rFonts.set(qn("w:cs"), FONT_NAME)
    rFonts.set(qn("w:eastAsia"), FONT_NAME)


def add_hyperlink(paragraph, text, url, size_pt=10):
    """Add a black, underlined clickable hyperlink run to a paragraph."""
    part = paragraph.part
    r_id = part.relate_to(
        url,
        docx.opc.constants.RELATIONSHIP_TYPE.HYPERLINK,
        is_external=True,
    )

    hyperlink = OxmlElement("w:hyperlink")
    hyperlink.set(qn("r:id"), r_id)

    new_run = OxmlElement("w:r")
    rPr = OxmlElement("w:rPr")

    rFonts = OxmlElement("w:rFonts")
    rFonts.set(qn("w:ascii"), FONT_NAME)
    rFonts.set(qn("w:hAnsi"), FONT_NAME)
    rFonts.set(qn("w:cs"), FONT_NAME)
    rFonts.set(qn("w:eastAsia"), FONT_NAME)
    rPr.append(rFonts)

    color = OxmlElement("w:color")
    color.set(qn("w:val"), "000000")
    rPr.append(color)

    u = OxmlElement("w:u")
    u.set(qn("w:val"), "single")
    rPr.append(u)

    sz = OxmlElement("w:sz")
    sz.set(qn("w:val"), str(int(size_pt * 2)))
    rPr.append(sz)

    szCs = OxmlElement("w:szCs")
    szCs.set(qn("w:val"), str(int(size_pt * 2)))
    rPr.append(szCs)

    new_run.append(rPr)
    t = OxmlElement("w:t")
    t.text = text
    new_run.append(t)

    hyperlink.append(new_run)
    paragraph._p.append(hyperlink)
    return hyperlink


def add_bottom_border(paragraph, sz_eighths=6, space_pt=1, color_hex="000000"):
    """Add a horizontal bottom rule under a section header paragraph."""
    pPr = paragraph._p.get_or_add_pPr()
    pBdr = pPr.find(qn("w:pBdr"))
    if pBdr is None:
        pBdr = OxmlElement("w:pBdr")
        pPr.append(pBdr)
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), str(sz_eighths))
    bottom.set(qn("w:space"), str(space_pt))
    bottom.set(qn("w:color"), color_hex)
    pBdr.append(bottom)


def configure_paragraph(
    paragraph,
    space_before_pt=0,
    space_after_pt=0,
    alignment=None,
    left_indent_pt=None,
    first_line_indent_pt=None,
    add_right_tab=False,
):
    """Configure spacing, indentation, alignment, and optional right-aligned tab stop."""
    pf = paragraph.paragraph_format
    pf.space_before = Pt(space_before_pt)
    pf.space_after = Pt(space_after_pt)
    pf.line_spacing_rule = WD_LINE_SPACING.SINGLE
    if alignment is not None:
        paragraph.alignment = alignment
    if left_indent_pt is not None:
        pf.left_indent = Pt(left_indent_pt)
    if first_line_indent_pt is not None:
        pf.first_line_indent = Pt(first_line_indent_pt)
    if add_right_tab:
        pf.tab_stops.add_tab_stop(RIGHT_TAB_POS, WD_TAB_ALIGNMENT.RIGHT)


def add_section_header(doc, title):
    """Add a Jake's Resume style section header (12pt Bold Small-Caps + bottom border)."""
    p = doc.add_paragraph()
    configure_paragraph(p, space_before_pt=4, space_after_pt=2)
    add_bottom_border(p, sz_eighths=6, space_pt=1, color_hex="000000")
    run = p.add_run(title)
    set_run_font(run, size_pt=12, bold=True, small_caps=True)
    return p


def add_bullet(doc, text):
    """Add a hanging-indent bullet point (10pt Times New Roman, '•  ' prefix)."""
    p = doc.add_paragraph()
    configure_paragraph(
        p,
        space_before_pt=0,
        space_after_pt=0,
        left_indent_pt=10.8,
        first_line_indent_pt=-9.0,
    )
    run = p.add_run(f"•  {text}")
    set_run_font(run, size_pt=10)
    return p


def build_docx(data: dict, output_docx_path: str):
    """Build the Word (.docx) CV from structured JSON data."""
    doc = Document()

    # Configure US Letter (8.5 x 11 in) and 0.5 in margins on all sides
    for section in doc.sections:
        section.page_width = Inches(8.5)
        section.page_height = Inches(11.0)
        section.top_margin = Inches(0.5)
        section.bottom_margin = Inches(0.5)
        section.left_margin = Inches(0.5)
        section.right_margin = Inches(0.5)

    # Configure Normal style
    normal_style = doc.styles["Normal"]
    normal_style.font.name = FONT_NAME
    normal_style.font.size = Pt(11)
    normal_style.font.color.rgb = BLACK
    rPr = normal_style.element.get_or_add_rPr()
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rPr.insert(0, rFonts)
    rFonts.set(qn("w:ascii"), FONT_NAME)
    rFonts.set(qn("w:hAnsi"), FONT_NAME)
    rFonts.set(qn("w:cs"), FONT_NAME)
    rFonts.set(qn("w:eastAsia"), FONT_NAME)

    normal_pf = normal_style.paragraph_format
    normal_pf.space_before = Pt(0)
    normal_pf.space_after = Pt(0)
    normal_pf.line_spacing_rule = WD_LINE_SPACING.SINGLE

    # 1. Header: Full Name (20pt Bold Small Caps, Centered)
    name_p = doc.add_paragraph()
    configure_paragraph(
        name_p,
        space_before_pt=0,
        space_after_pt=2,
        alignment=WD_ALIGN_PARAGRAPH.CENTER,
    )
    name_run = name_p.add_run(data.get("name", ""))
    set_run_font(name_run, size_pt=20, bold=True, small_caps=True)

    # 2. Sub-header: Contact Info (10pt Centered, pipe-separated, clickable links)
    contact_items = data.get("contact", [])
    if contact_items:
        contact_p = doc.add_paragraph()
        configure_paragraph(
            contact_p,
            space_before_pt=0,
            space_after_pt=0,
            alignment=WD_ALIGN_PARAGRAPH.CENTER,
        )
        for idx, item in enumerate(contact_items):
            if idx > 0:
                sep_run = contact_p.add_run(" | ")
                set_run_font(sep_run, size_pt=10)
            text = item.get("text", "")
            url = item.get("url")
            if url:
                add_hyperlink(contact_p, text, url, size_pt=10)
            else:
                txt_run = contact_p.add_run(text)
                set_run_font(txt_run, size_pt=10)

    titles = data.get("section_titles", {})

    # 3. Education Section
    education_list = data.get("education", [])
    if education_list:
        add_section_header(doc, titles.get("education", "Education"))
        for edu in education_list:
            # Line 1: School (11pt Bold) + Right Tab + Location (11pt Regular)
            p1 = doc.add_paragraph()
            configure_paragraph(p1, space_before_pt=2, space_after_pt=0, add_right_tab=True)
            r_school = p1.add_run(edu.get("school", ""))
            set_run_font(r_school, size_pt=11, bold=True)
            if edu.get("location"):
                r_loc = p1.add_run(f"\t{edu['location']}")
                set_run_font(r_loc, size_pt=11, bold=False)

            # Line 2: Degree (10pt Italic) + Right Tab + Date (10pt Italic)
            p2 = doc.add_paragraph()
            configure_paragraph(p2, space_before_pt=0, space_after_pt=0, add_right_tab=True)
            r_deg = p2.add_run(edu.get("degree", ""))
            set_run_font(r_deg, size_pt=10, italic=True)
            if edu.get("date"):
                r_date = p2.add_run(f"\t{edu['date']}")
                set_run_font(r_date, size_pt=10, italic=True)

            # Optional bullets under education
            for bullet in edu.get("bullets", []):
                add_bullet(doc, bullet)

    # 4. Experience Section
    experience_list = data.get("experience", [])
    if experience_list:
        add_section_header(doc, titles.get("experience", "Experience"))
        for exp in experience_list:
            # Line 1: Role Title (11pt Bold) + Right Tab + Date (11pt Regular)
            p1 = doc.add_paragraph()
            configure_paragraph(p1, space_before_pt=2, space_after_pt=0, add_right_tab=True)
            r_title = p1.add_run(exp.get("title", ""))
            set_run_font(r_title, size_pt=11, bold=True)
            if exp.get("date"):
                r_date = p1.add_run(f"\t{exp['date']}")
                set_run_font(r_date, size_pt=11, bold=False)

            # Line 2: Company (10pt Italic) + Right Tab + Location (10pt Italic)
            p2 = doc.add_paragraph()
            configure_paragraph(p2, space_before_pt=0, space_after_pt=0, add_right_tab=True)
            r_comp = p2.add_run(exp.get("company", ""))
            set_run_font(r_comp, size_pt=10, italic=True)
            if exp.get("location"):
                r_loc = p2.add_run(f"\t{exp['location']}")
                set_run_font(r_loc, size_pt=10, italic=True)

            # Bullets
            for bullet in exp.get("bullets", []):
                add_bullet(doc, bullet)

    # 5. Projects Section
    projects_list = data.get("projects", [])
    if projects_list:
        add_section_header(doc, titles.get("projects", "Projects"))
        for proj in projects_list:
            p = doc.add_paragraph()
            configure_paragraph(
                p,
                space_before_pt=2,
                space_after_pt=0,
                add_right_tab=bool(proj.get("date")),
            )
            r_name = p.add_run(proj.get("name", ""))
            set_run_font(r_name, size_pt=11, bold=True)

            tech = proj.get("technologies", "")
            if tech:
                r_sep = p.add_run(" | ")
                set_run_font(r_sep, size_pt=11)
                r_tech = p.add_run(tech)
                set_run_font(r_tech, size_pt=11, italic=True)

            if proj.get("date"):
                r_date = p.add_run(f"\t{proj['date']}")
                set_run_font(r_date, size_pt=11)

            for bullet in proj.get("bullets", []):
                add_bullet(doc, bullet)

    # 6. Technical Skills Section
    skills_list = data.get("skills", [])
    if skills_list:
        add_section_header(doc, titles.get("skills", "Technical Skills"))
        for skill in skills_list:
            p = doc.add_paragraph()
            configure_paragraph(p, space_before_pt=0, space_after_pt=0)
            cat = skill.get("category", "")
            items = skill.get("items", "")
            r_cat = p.add_run(f"{cat}: ")
            set_run_font(r_cat, size_pt=11, bold=True)
            r_items = p.add_run(items)
            set_run_font(r_items, size_pt=11, bold=False)

    # 7. Optional Additional Custom Sections (if user adds e.g. Certifications, Awards, Languages)
    for extra_sec in data.get("custom_sections", []):
        sec_title = extra_sec.get("title", "")
        if not sec_title:
            continue
        add_section_header(doc, sec_title)
        for item in extra_sec.get("items", []):
            if isinstance(item, str):
                add_bullet(doc, item)
            elif isinstance(item, dict):
                if "category" in item and "items" in item:
                    p = doc.add_paragraph()
                    configure_paragraph(p, space_before_pt=0, space_after_pt=0)
                    r_cat = p.add_run(f"{item['category']}: ")
                    set_run_font(r_cat, size_pt=11, bold=True)
                    r_items = p.add_run(item["items"])
                    set_run_font(r_items, size_pt=11, bold=False)
                else:
                    p1 = doc.add_paragraph()
                    configure_paragraph(p1, space_before_pt=2, space_after_pt=0, add_right_tab=True)
                    r_left = p1.add_run(item.get("title", ""))
                    set_run_font(r_left, size_pt=11, bold=True)
                    if item.get("subtitle"):
                        r_sep = p1.add_run(" | ")
                        set_run_font(r_sep, size_pt=11)
                        r_sub = p1.add_run(item["subtitle"])
                        set_run_font(r_sub, size_pt=11, italic=True)
                    if item.get("right"):
                        r_right = p1.add_run(f"\t{item['right']}")
                        set_run_font(r_right, size_pt=11)
                    for b in item.get("bullets", []):
                        add_bullet(doc, b)

    doc.save(output_docx_path)
    return output_docx_path


def convert_docx_to_pdf(docx_path: str, pdf_path: str) -> bool:
    """Convert .docx to .pdf using Microsoft Word COM automation on Windows."""
    abs_docx = str(Path(docx_path).resolve())
    abs_pdf = str(Path(pdf_path).resolve())

    # Try docx2pdf first if installed
    try:
        from docx2pdf import convert

        convert(abs_docx, abs_pdf)
        if os.path.exists(abs_pdf):
            return True
    except ImportError:
        pass
    except Exception:
        pass

    # Use PowerShell + Microsoft Word COM Automation (wdFormatPDF = 17)
    ps_script = f"""
$ErrorActionPreference = 'Stop'
$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0
try {{
    $doc = $word.Documents.Open('{abs_docx.replace("'", "''")}', $false, $true)
    $doc.SaveAs([ref] '{abs_pdf.replace("'", "''")}', [ref] 17)
    $doc.Close(0)
}} finally {{
    $word.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null
}}
"""
    result = subprocess.run(
        ["powershell", "-NoProfile", "-NonInteractive", "-Command", ps_script],
        capture_output=True,
        text=True,
    )
    if result.returncode != 0:
        print(f"[!] Word PDF dönüştürme hatası: {result.stderr.strip()}", file=sys.stderr)
        return False
    return os.path.exists(abs_pdf)


def main():
    parser = argparse.ArgumentParser(
        description="Jake's Resume (Word -> PDF) CV Oluşturucu"
    )
    parser.add_argument(
        "--data",
        "-d",
        default="cv_data.json",
        help="CV verilerini içeren JSON dosyası (varsayılan: cv_data.json)",
    )
    parser.add_argument(
        "--output",
        "-o",
        default=None,
        help="Çıktı dosya adı (uzantısız veya .pdf/.docx). Belirtilmezse isimden üretilir.",
    )
    parser.add_argument(
        "--no-pdf",
        action="store_true",
        help="Sadece Word (.docx) oluştur, PDF'e çevirme",
    )
    args = parser.parse_args()

    data_path = Path(args.data)
    if not data_path.exists():
        print(f"[!] Veri dosyası bulunamadı: {data_path}", file=sys.stderr)
        sys.exit(1)

    with open(data_path, "r", encoding="utf-8") as f:
        data = json.load(f)

    if args.output:
        base_out = Path(args.output)
        if base_out.suffix.lower() in (".pdf", ".docx"):
            base_out = base_out.with_suffix("")
    else:
        raw_name = data.get("name", "CV").strip().replace(" ", "_")
        base_out = Path(f"{raw_name}_CV")

    docx_out = str(base_out.with_suffix(".docx"))
    pdf_out = str(base_out.with_suffix(".pdf"))

    print(f"[*] Word belgesi oluşturuluyor: {docx_out}")
    build_docx(data, docx_out)
    print(f"[+] Word belgesi hazır: {docx_out}")

    if not args.no_pdf:
        print(f"[*] PDF'e dönüştürülüyor: {pdf_out}")
        ok = convert_docx_to_pdf(docx_out, pdf_out)
        if ok:
            print(f"[+] PDF belgesi hazır: {pdf_out}")
        else:
            print("[!] PDF oluşturulamadı, ancak .docx dosyası kullanıma hazır.")


if __name__ == "__main__":
    main()
