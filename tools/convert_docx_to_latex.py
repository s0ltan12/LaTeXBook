import argparse
import re
from pathlib import Path

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml.table import CT_Tbl
from docx.oxml.text.paragraph import CT_P
from docx.shared import Emu
from docx.table import Table
from docx.text.paragraph import Paragraph


ALIGNMENTS = {
    WD_ALIGN_PARAGRAPH.LEFT: "raggedright",
    WD_ALIGN_PARAGRAPH.CENTER: "centering",
    WD_ALIGN_PARAGRAPH.RIGHT: "raggedleft",
    WD_ALIGN_PARAGRAPH.JUSTIFY: "justifying",
}


def iter_block_items(doc):
    body = doc.element.body
    for child in body.iterchildren():
        if isinstance(child, CT_P):
            yield Paragraph(child, doc)
        elif isinstance(child, CT_Tbl):
            yield Table(child, doc)


def escape_latex(text):
    replacements = {
        "\\": r"\textbackslash{}",
        "&": r"\&",
        "%": r"\%",
        "$": r"\$",
        "#": r"\#",
        "_": r"\_",
        "{": r"\{",
        "}": r"\}",
        "~": r"\textasciitilde{}",
        "^": r"\textasciicircum{}",
    }
    return "".join(replacements.get(ch, ch) for ch in text)


def length_pt(value, default=None):
    if value is None:
        return default
    return value.pt


def emu_pt(value):
    return Emu(int(value)).pt


def attr_int(element, attr_name, default=0):
    value = element.get(qn(attr_name))
    return int(value) if value is not None else default


def para_spacing(paragraph):
    fmt = paragraph.paragraph_format
    before = length_pt(fmt.space_before, 0) or 0
    after = length_pt(fmt.space_after, 0) or 0
    line = fmt.line_spacing
    line_pt = None
    if hasattr(line, "pt"):
        line_pt = line.pt
    elif isinstance(line, (int, float)):
        line_pt = 12 * float(line)
    return before, after, line_pt


def para_indent(paragraph):
    fmt = paragraph.paragraph_format
    left = length_pt(fmt.left_indent, 0) or 0
    first = length_pt(fmt.first_line_indent, 0) or 0
    right = length_pt(fmt.right_indent, 0) or 0
    return left, first, right


def paragraph_env_start(paragraph):
    before, after, line_pt = para_spacing(paragraph)
    left, first, right = para_indent(paragraph)
    align = ALIGNMENTS.get(paragraph.alignment)
    settings = [
        rf"\setlength{{\parskip}}{{{after:.2f}pt}}",
        rf"\setlength{{\leftskip}}{{{left:.2f}pt}}",
        rf"\setlength{{\rightskip}}{{{right:.2f}pt}}",
        rf"\setlength{{\parindent}}{{{max(first, 0):.2f}pt}}",
    ]
    if before:
        settings.insert(0, rf"\vspace*{{{before:.2f}pt}}")
    if line_pt:
        settings.append(rf"\fontsize{{12}}{{{line_pt:.2f}}}\selectfont")
    if align:
        settings.append("\\" + align)
    return r"\begingroup " + " ".join(settings)


def run_size_pt(run):
    if run.font.size:
        return run.font.size.pt
    style_font = getattr(run.style, "font", None)
    if style_font and style_font.size:
        return style_font.size.pt
    return None


def wrap_run_format(run, text):
    size = run_size_pt(run)
    if size:
        text = rf"{{\fontsize{{{size:.2f}}}{{{size * 1.2:.2f}}}\selectfont {text}}}"
    if run.bold:
        text = rf"\textbf{{{text}}}"
    if run.italic:
        text = rf"\textit{{{text}}}"
    if run.underline:
        text = rf"\underline{{{text}}}"
    if run.font.color and run.font.color.rgb:
        color = str(run.font.color.rgb)
        text = rf"\textcolor[HTML]{{{color}}}{{{text}}}"
    return text


def extract_run_images(run, paragraph, image_dir, image_counter):
    parts = []
    drawings = run.element.xpath(".//w:drawing")
    for drawing in drawings:
        blips = drawing.xpath(".//a:blip")
        extents = drawing.xpath(".//wp:extent")
        for blip in blips:
            r_id = blip.get(qn("r:embed"))
            if not r_id:
                continue
            part = paragraph.part.related_parts[r_id]
            ext = Path(part.partname).suffix.lower() or ".png"
            image_counter[0] += 1
            filename = f"docx_img_{image_counter[0]:03d}{ext}"
            output_path = image_dir / filename
            output_path.write_bytes(part.blob)
            rel_path = output_path.as_posix()
            if extents:
                width = emu_pt(extents[0].get("cx"))
                height = emu_pt(extents[0].get("cy"))
                parts.append(
                    rf"\includegraphics[width={width:.2f}pt,height={height:.2f}pt,keepaspectratio]{{{rel_path}}}"
                )
            else:
                parts.append(rf"\includegraphics{{{rel_path}}}")
    return parts


def paragraph_has_page_break(paragraph):
    return bool(paragraph._element.xpath(".//w:br[@w:type='page']"))


def convert_paragraph(paragraph, image_dir, image_counter):
    if paragraph.style and paragraph.style.name == "Header and Footer":
        return []

    parts = []
    for run in paragraph.runs:
        parts.extend(extract_run_images(run, paragraph, image_dir, image_counter))
        for piece in re.split(r"(\n|\t)", run.text or ""):
            if piece == "\n":
                parts.append(r"\\")
            elif piece == "\t":
                parts.append(r"\hspace*{2em}")
            elif piece:
                parts.append(wrap_run_format(run, escape_latex(piece)))

    lines = []
    if paragraph_has_page_break(paragraph):
        lines.append(r"\newpage")
    content = "".join(parts).strip()
    if not content:
        lines.append(r"\vspace{\baselineskip}")
        return lines
    lines.append(paragraph_env_start(paragraph))
    lines.append(content + r"\par")
    lines.append(r"\endgroup")
    return lines


def cell_widths_pt(table):
    grid = table._tbl.tblGrid
    if grid is None:
        return []
    widths = []
    for col in grid.gridCol_lst:
        widths.append(attr_int(col, "w:w") / 20)
    return widths


def convert_cell(cell):
    paras = []
    for para in cell.paragraphs:
        text = ""
        for run in para.runs:
            if run.text:
                text += wrap_run_format(run, escape_latex(run.text))
        if text.strip():
            paras.append(text.strip())
    return r"\par ".join(paras)


def convert_table(table):
    widths = cell_widths_pt(table)
    col_count = len(table.columns)
    if not widths or len(widths) != col_count:
        widths = [420 / max(col_count, 1)] * col_count
    spec = "|".join(rf"p{{{width:.2f}pt}}" for width in widths)
    lines = [
        r"\begingroup",
        r"\small",
        r"\setlength{\tabcolsep}{3pt}",
        r"\renewcommand{\arraystretch}{1.15}",
        rf"\begin{{longtable}}{{|{spec}|}}",
        r"\hline",
    ]
    for row in table.rows:
        cells = [convert_cell(cell) for cell in row.cells[:col_count]]
        lines.append(" & ".join(cells) + r" \\ \hline")
    lines.extend([r"\end{longtable}", r"\endgroup"])
    return lines


def section_geometry(section):
    return {
        "paperwidth": section.page_width.pt,
        "paperheight": section.page_height.pt,
        "left": section.left_margin.pt,
        "right": section.right_margin.pt,
        "top": section.top_margin.pt,
        "bottom": section.bottom_margin.pt,
    }


def convert_docx(input_path, output_tex, image_dir):
    doc = Document(input_path)
    image_dir.mkdir(parents=True, exist_ok=True)
    image_counter = [0]
    geometry = section_geometry(doc.sections[0])

    body = []
    for block in iter_block_items(doc):
        if isinstance(block, Paragraph):
            body.extend(convert_paragraph(block, image_dir, image_counter))
        elif isinstance(block, Table):
            body.extend(convert_table(block))
        body.append("")

    lines = [
        "% !TEX program = xelatex",
        r"\documentclass[12pt]{article}",
        r"\usepackage{fontspec}",
        r"\usepackage{polyglossia}",
        r"\usepackage{geometry}",
        r"\usepackage{graphicx}",
        r"\usepackage[table]{xcolor}",
        r"\usepackage{array}",
        r"\usepackage{longtable}",
        r"\usepackage{ragged2e}",
        r"\usepackage{ulem}",
        r"\usepackage[hidelinks]{hyperref}",
        r"\setmainlanguage{english}",
        r"\setotherlanguage{arabic}",
        r"\setmainfont{Times New Roman}",
        r"\newfontfamily\arabicfont[Script=Arabic]{Times New Roman}",
        rf"\geometry{{paperwidth={geometry['paperwidth']:.2f}pt,paperheight={geometry['paperheight']:.2f}pt,left={geometry['left']:.2f}pt,right={geometry['right']:.2f}pt,top={geometry['top']:.2f}pt,bottom={geometry['bottom']:.2f}pt}}",
        r"\setlength{\parindent}{0pt}",
        r"\setlength{\parskip}{0pt}",
        r"\pagestyle{plain}",
        r"\begin{document}",
        "",
        *body,
        r"\end{document}",
        "",
    ]
    output_tex.write_text("\n".join(lines), encoding="utf-8")
    return image_counter[0]


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--input", default="Graduation Project.docx")
    parser.add_argument("--output", default="main.tex")
    parser.add_argument("--image-dir", default="images")
    args = parser.parse_args()

    count = convert_docx(Path(args.input), Path(args.output), Path(args.image_dir))
    print(f"Wrote {args.output} and extracted {count} images to {args.image_dir}")


if __name__ == "__main__":
    main()
