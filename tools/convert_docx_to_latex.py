import argparse
import os
import re
from pathlib import Path

from docx import Document
from docx.oxml.ns import qn
from docx.oxml.table import CT_Tbl
from docx.oxml.text.paragraph import CT_P
from docx.table import Table
from docx.text.paragraph import Paragraph


def iter_block_items(doc):
    body = doc.element.body
    for child in body.iterchildren():
        if isinstance(child, CT_P):
            yield Paragraph(child, doc)
        elif isinstance(child, CT_Tbl):
            yield Table(child, doc)


def escape_latex(text):
    if not text:
        return ""
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
    result = []
    for ch in text:
        result.append(replacements.get(ch, ch))
    return "".join(result)


def normalize_text(text):
    return " ".join(text.split()).strip()


def convert_runs(paragraph):
    parts = []
    for run in paragraph.runs:
        if not run.text:
            continue
        if run.text.isspace():
            parts.append(" ")
            continue
        text = escape_latex(run.text)
        if run.bold and run.italic:
            text = f"\\textbf{{\\textit{{{text}}}}}"
        elif run.bold:
            text = f"\\textbf{{{text}}}"
        elif run.italic:
            text = f"\\textit{{{text}}}"
        parts.append(text)
    joined = "".join(parts)
    return normalize_text(joined)


def extract_images(paragraph, image_dir, image_counter):
    image_paths = []
    for run in paragraph.runs:
        blips = run.element.xpath(".//a:blip")
        for blip in blips:
            r_id = blip.get(qn("r:embed"))
            part = paragraph.part.related_parts[r_id]
            ext = Path(part.partname).suffix.lower() or ".png"
            image_counter[0] += 1
            filename = f"img_{image_counter[0]:03d}{ext}"
            output_path = image_dir / filename
            with output_path.open("wb") as f:
                f.write(part.blob)
            image_paths.append(output_path)
    return image_paths


def convert_table(table):
    rows = []
    max_cols = 0
    for row in table.rows:
        row_cells = [normalize_text(cell.text) for cell in row.cells]
        max_cols = max(max_cols, len(row_cells))
        rows.append(row_cells)
    if max_cols == 0:
        return []
    col_spec = "|" + "|".join(["l"] * max_cols) + "|"
    lines = ["\\begin{table}[H]", "\\centering", f"\\begin{{tabular}}{{{col_spec}}}", "\\hline"]
    for row in rows:
        padded = row + [""] * (max_cols - len(row))
        escaped = [escape_latex(cell) for cell in padded]
        lines.append(" & ".join(escaped) + r" \\"
        )
        lines.append("\\hline")
    lines.append("\\end{tabular}")
    lines.append("\\end{table}")
    return lines


def find_chapter_titles(doc):
    titles = {}
    for p in doc.paragraphs:
        text = normalize_text(p.text)
        match = re.match(r"^(\d+)\.\s*Chapter\s+\d+\s+(.+)$", text, re.IGNORECASE)
        if match:
            num = int(match.group(1))
            title = re.sub(r"\.{2,}.*$", "", match.group(2)).strip()
            if title:
                titles[num] = title
    return titles


def is_toc_or_list_line(text):
    if not text:
        return False
    if text.lower() in {"table of contents", "list of figures", "list of tables"}:
        return True
    if re.match(r"^\d+\.\s*Chapter\s+\d+\b", text, re.IGNORECASE):
        return True
    if "...." in text:
        return True
    return False


def heading_from_number(number_str, title):
    level = number_str.count(".")
    title = escape_latex(title)
    if level == 1:
        return f"\\section{{{title}}}"
    if level == 2:
        return f"\\subsection{{{title}}}"
    if level == 3:
        return f"\\subsubsection{{{title}}}"
    return f"\\paragraph{{{title}}}"


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--input", default="Graduation Project.docx")
    parser.add_argument("--out", default=".")
    args = parser.parse_args()

    doc = Document(args.input)
    out_dir = Path(args.out)
    chapters_dir = out_dir / "chapters"
    images_dir = out_dir / "images"
    chapters_dir.mkdir(parents=True, exist_ok=True)
    images_dir.mkdir(parents=True, exist_ok=True)

    chapter_titles = find_chapter_titles(doc)

    frontmatter_lines = []
    chapter_lines = {}
    current_chapter = None
    in_list = False
    image_counter = [0]
    skip_mode = None
    frontmatter_headings = {"declaration", "acknowledgements", "abstract"}

    section_re = re.compile(r"^(\d+(?:\.\d+)+)\s+(.+)$")

    def ensure_chapter(chapter_num):
        nonlocal current_chapter
        if current_chapter != chapter_num:
            current_chapter = chapter_num
            if chapter_num not in chapter_lines:
                title = chapter_titles.get(chapter_num, f"Chapter {chapter_num}")
                chapter_lines[chapter_num] = [f"\\chapter{{{escape_latex(title)}}}", ""]

    def append_lines(lines):
        target = frontmatter_lines if current_chapter is None else chapter_lines[current_chapter]
        target.extend(lines)

    for block in iter_block_items(doc):
        if isinstance(block, Paragraph):
            if block.style and block.style.name == "Header and Footer":
                continue
            text = normalize_text(block.text)
            images = extract_images(block, images_dir, image_counter)

            if text:
                lowered = text.lower()
                if lowered == "table of contents":
                    skip_mode = "toc"
                    continue
                if lowered in {"list of figures", "list of tables"}:
                    skip_mode = "list"
                    continue

            if skip_mode:
                if text.upper() == "ABSTRACT":
                    skip_mode = None
                else:
                    section_match = section_re.match(text)
                    if section_match and not is_toc_or_list_line(text):
                        skip_mode = None
                    else:
                        continue

            if text and is_toc_or_list_line(text):
                continue

            section_match = section_re.match(text)
            if section_match:
                if in_list:
                    append_lines(["\\end{itemize}", ""])
                    in_list = False
                number_str = section_match.group(1)
                title = section_match.group(2)
                chapter_num = int(number_str.split(".")[0])
                ensure_chapter(chapter_num)
                append_lines([heading_from_number(number_str, title), ""])
                continue

            if text and text.lower() in frontmatter_headings and current_chapter is None:
                if in_list:
                    append_lines(["\\end{itemize}", ""])
                    in_list = False
                append_lines([f"\\chapter*{{{escape_latex(text.title())}}}", ""])
                continue

            paragraph_text = convert_runs(block)

            if block.style and block.style.name == "List Paragraph" and paragraph_text:
                if not in_list:
                    append_lines(["\\begin{itemize}"])
                    in_list = True
                append_lines([f"\\item {paragraph_text}"])
            else:
                if in_list:
                    append_lines(["\\end{itemize}", ""])
                    in_list = False
                if paragraph_text:
                    append_lines([paragraph_text, ""])

            if images:
                if in_list:
                    append_lines(["\\end{itemize}", ""])
                    in_list = False
                for path in images:
                    rel_path = path.relative_to(out_dir).as_posix()
                    append_lines([
                        "\\begin{figure}[H]",
                        "\\centering",
                        f"\\includegraphics[width=0.95\\linewidth]{{{rel_path}}}",
                        "\\end{figure}",
                        "",
                    ])
        else:
            if in_list:
                append_lines(["\\end{itemize}", ""])
                in_list = False
            table_lines = convert_table(block)
            if table_lines:
                append_lines(table_lines + [""])

    if in_list:
        append_lines(["\\end{itemize}", ""])

    for chapter_num, lines in chapter_lines.items():
        chapter_path = chapters_dir / f"chapter{chapter_num:02d}.tex"
        chapter_path.write_text("\n".join(lines).strip() + "\n", encoding="utf-8")

    frontmatter_path = out_dir / "frontmatter.tex"
    if frontmatter_lines:
        frontmatter_path.write_text("\n".join(frontmatter_lines).strip() + "\n", encoding="utf-8")

    main_lines = [
        "\\documentclass[12pt]{report}",
        "\\usepackage[a4paper,margin=1in]{geometry}",
        "\\usepackage{graphicx}",
        "\\usepackage{longtable}",
        "\\usepackage{array}",
        "\\usepackage{hyperref}",
        "\\usepackage{float}",
        "",
        "\\begin{document}",
    ]

    if frontmatter_lines:
        main_lines.append("\\include{frontmatter}")

    for chapter_num in sorted(chapter_lines.keys()):
        main_lines.append(f"\\include{{chapters/chapter{chapter_num:02d}}}")

    main_lines.extend(["", "\\end{document}"])
    (out_dir / "main.tex").write_text("\n".join(main_lines) + "\n", encoding="utf-8")


if __name__ == "__main__":
    main()
