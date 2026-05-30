"""
thesis_corrector.py
Core engine: loads norm_control_rules.json and applies formatting rules
to a student thesis .docx file using python-docx.
"""

import json
import re
import copy
from pathlib import Path
from dataclasses import dataclass, field
from typing import Optional

from docx import Document
from docx.shared import Pt, Mm, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import lxml.etree as etree


# ──────────────────────────────────────────────
# Data structures
# ──────────────────────────────────────────────

@dataclass
class Violation:
    location: str          # e.g. "Para 12" or "Table 3 caption"
    rule_ref: str          # e.g. "7.1.10"
    description: str
    original: str
    corrected: Optional[str] = None
    auto_fixed: bool = False


@dataclass
class CorrectionReport:
    violations: list = field(default_factory=list)
    auto_fixed_count: int = 0
    manual_review_count: int = 0

    def add(self, v: Violation):
        self.violations.append(v)
        if v.auto_fixed:
            self.auto_fixed_count += 1
        else:
            self.manual_review_count += 1

    def summary(self) -> str:
        lines = [
            f"=== Отчёт нормоконтроля ===",
            f"Всего нарушений: {len(self.violations)}",
            f"  Исправлено автоматически: {self.auto_fixed_count}",
            f"  Требуют ручной проверки:  {self.manual_review_count}",
            "",
        ]
        for i, v in enumerate(self.violations, 1):
            status = "✅ ИСПРАВЛЕНО" if v.auto_fixed else "⚠️  ПРОВЕРИТЬ"
            lines.append(f"{i}. [{status}] {v.location} | Правило {v.rule_ref}")
            lines.append(f"   {v.description}")
            if v.original:
                lines.append(f"   Было:  {v.original[:120]}")
            if v.corrected:
                lines.append(f"   Стало: {v.corrected[:120]}")
            lines.append("")
        return "\n".join(lines)


# ──────────────────────────────────────────────
# Helper utilities
# ──────────────────────────────────────────────

def emu_to_mm(emu: int) -> float:
    return emu / 914400 * 25.4


def pt_to_half_pt(pt: float) -> int:
    """python-docx stores font size in half-points (twips)."""
    return int(pt * 2)


def get_paragraph_text(para) -> str:
    return "".join(run.text for run in para.runs)


def is_heading(para) -> bool:
    return para.style.name.lower().startswith("heading")


def heading_level(para) -> int:
    name = para.style.name.lower()
    if "heading 1" in name:
        return 1
    if "heading 2" in name:
        return 2
    if "heading 3" in name:
        return 3
    return 0


def detect_section_heading(text: str) -> bool:
    """True if paragraph looks like a numbered section heading (1, 1.1, 1.1.1)."""
    return bool(re.match(r'^\d+(\.\d+)*\s+\S', text.strip()))


def detect_structural_heading(text: str, structural_headings: list) -> bool:
    clean = text.strip().rstrip(".")
    return any(clean.lower() == h.lower() for h in structural_headings)


# ──────────────────────────────────────────────
# Correction modules
# ──────────────────────────────────────────────

class FontCorrector:
    """Rule 5.3 / typography section: font, size, line spacing."""

    def __init__(self, rules: dict, report: CorrectionReport):
        self.rules = rules["typography"]["main_body"]
        self.table_rules = rules["typography"]["tables"]
        self.report = report

    def fix_paragraph(self, para, location: str):
        required_font = self.rules["font_family"]
        required_size = self.rules["font_size_pt"]

        for i, run in enumerate(para.runs):
            changed = False
            original_font = run.font.name
            original_size = run.font.size

            # Font family
            if run.font.name and run.font.name != required_font:
                run.font.name = required_font
                changed = True

            # Font size
            if run.font.size and run.font.size != Pt(required_size):
                orig_pt = run.font.size.pt if run.font.size else "?"
                run.font.size = Pt(required_size)
                changed = True

            if changed:
                self.report.add(Violation(
                    location=f"{location} / run {i}",
                    rule_ref="5.3",
                    description=f"Шрифт: {original_font or '?'} {orig_pt if original_size else '?'}pt → {required_font} {required_size}pt",
                    original=f"{original_font} {original_size}",
                    corrected=f"{required_font} {required_size}pt",
                    auto_fixed=True
                ))


class PageLayoutCorrector:
    """Rule 5.5: margins."""

    def __init__(self, rules: dict, report: CorrectionReport):
        self.margins = rules["page_layout"]["margins_mm"]
        self.report = report

    def fix_document(self, doc: Document):
        required = {
            "left":   Mm(self.margins["left"]),
            "top":    Mm(self.margins["top"]),
            "right":  Mm(self.margins["right"]),
            "bottom": Mm(self.margins["bottom"]),
        }
        for i, section in enumerate(doc.sections):
            changed = []
            if abs(section.left_margin - required["left"]) > 1000:
                section.left_margin = required["left"]
                changed.append(f"left→{self.margins['left']}мм")
            if abs(section.top_margin - required["top"]) > 1000:
                section.top_margin = required["top"]
                changed.append(f"top→{self.margins['top']}мм")
            if abs(section.right_margin - required["right"]) > 1000:
                section.right_margin = required["right"]
                changed.append(f"right→{self.margins['right']}мм")
            if abs(section.bottom_margin - required["bottom"]) > 1000:
                section.bottom_margin = required["bottom"]
                changed.append(f"bottom→{self.margins['bottom']}мм")
            if changed:
                self.report.add(Violation(
                    location=f"Section {i+1}",
                    rule_ref="5.5",
                    description="Поля документа исправлены: " + ", ".join(changed),
                    original="см. оригинал",
                    corrected=str(self.margins),
                    auto_fixed=True
                ))


class ParagraphFormatCorrector:
    """Rule 5.5 / paragraph_formatting: indent, alignment, spacing."""

    def __init__(self, rules: dict, report: CorrectionReport):
        self.rules = rules["paragraph_formatting"]
        self.report = report
        self.skip_styles = {"heading 1", "heading 2", "heading 3",
                            "title", "subtitle", "caption"}

    def fix_paragraph(self, para, location: str):
        style_name = para.style.name.lower()
        if any(s in style_name for s in self.skip_styles):
            return
        if not para.text.strip():
            return

        pf = para.paragraph_format
        required_indent = Mm(self.rules["first_line_indent_mm"])
        required_align = WD_ALIGN_PARAGRAPH.JUSTIFY

        changed = []

        # First line indent
        if pf.first_line_indent is None or abs(pf.first_line_indent - required_indent) > 5000:
            pf.first_line_indent = required_indent
            changed.append(f"абзацный отступ → {self.rules['first_line_indent_mm']}мм")

        # Alignment
        if pf.alignment not in (WD_ALIGN_PARAGRAPH.JUSTIFY, None):
            old = pf.alignment
            pf.alignment = required_align
            changed.append(f"выравнивание {old} → JUSTIFY")

        if changed:
            self.report.add(Violation(
                location=location,
                rule_ref="5.5",
                description="; ".join(changed),
                original=get_paragraph_text(para)[:80],
                corrected=None,
                auto_fixed=True
            ))


class HeadingChecker:
    """Rules 7.1.2, 7.1.10, 7.1.11: section numbering, no dot at end, no hyphenation."""

    def __init__(self, rules: dict, report: CorrectionReport):
        self.structural = rules["document_structure"]["structural_headings_no_numbering"]
        self.report = report

    def check_paragraph(self, para, location: str):
        text = get_paragraph_text(para).strip()
        if not text:
            return

        is_struct = detect_structural_heading(text, self.structural)
        is_sect = detect_section_heading(text)

        if not is_struct and not is_sect:
            return

        # Check: no dot at end (7.1.10)
        if text.endswith(".") and not text.endswith("..."):
            self.report.add(Violation(
                location=location,
                rule_ref="7.1.10",
                description="Заголовок заканчивается точкой (не допускается)",
                original=text,
                corrected=text.rstrip("."),
                auto_fixed=False  # Needs human confirmation
            ))

        # Check: no hyphenation (переносы не допускаются — hard to check, just flag if contains soft hyphen)
        if "\u00ad" in text:
            self.report.add(Violation(
                location=location,
                rule_ref="7.1.10",
                description="В заголовке присутствует перенос (мягкий дефис)",
                original=text,
                auto_fixed=False
            ))


class FigureCaptionChecker:
    """Rule 7.3.1: figure captions format — «Рисунок N – Наименование»."""

    CAPTION_PATTERN = re.compile(
        r'^(Рисунок|Рис\.?)\s*(\d+[\.\d]*)\s*[-–—]\s*(.+)$', re.IGNORECASE
    )
    WRONG_PREFIX = re.compile(r'^(Рис\.)\s', re.IGNORECASE)

    def __init__(self, report: CorrectionReport):
        self.report = report

    def check_paragraph(self, para, location: str):
        text = get_paragraph_text(para).strip()
        if not text.startswith(("Рисунок", "Рис")):
            return

        match = self.CAPTION_PATTERN.match(text)
        if not match:
            self.report.add(Violation(
                location=location,
                rule_ref="7.3.1",
                description="Подпись рисунка не соответствует формату «Рисунок N – Наименование»",
                original=text,
                corrected="Рисунок N – Наименование",
                auto_fixed=False
            ))
            return

        prefix, num, name = match.group(1), match.group(2), match.group(3)
        if prefix.lower() != "рисунок":
            self.report.add(Violation(
                location=location,
                rule_ref="7.3.1",
                description=f"Используется сокращение «{prefix}» вместо «Рисунок»",
                original=text,
                corrected=f"Рисунок {num} – {name}",
                auto_fixed=False
            ))


class TableCaptionChecker:
    """Rule 7.5.1, 7.5.3: table captions format — «Таблица N – Наименование»."""

    CAPTION_PATTERN = re.compile(
        r'^(Таблица|Табл\.?)\s*(\d+[\.\d]*)\s*[-–—]?\s*(.*)$', re.IGNORECASE
    )

    def __init__(self, report: CorrectionReport):
        self.report = report

    def check_paragraph(self, para, location: str):
        text = get_paragraph_text(para).strip()
        if not text.startswith(("Таблица", "Табл")):
            return

        match = self.CAPTION_PATTERN.match(text)
        if not match:
            self.report.add(Violation(
                location=location,
                rule_ref="7.5.1",
                description="Заголовок таблицы не соответствует формату «Таблица N – Наименование»",
                original=text,
                auto_fixed=False
            ))


class BibliographyChecker:
    """Rule 7.1.19: bibliography citation format [N] and ordering."""

    # In-text citation: should be [26] not (26) or 26
    WRONG_CITE_PAREN = re.compile(r'\((\d+)\)')
    BRACKET_CITE = re.compile(r'\[(\d+)\]')

    def __init__(self, report: CorrectionReport):
        self.report = report

    def check_paragraph(self, para, location: str):
        text = get_paragraph_text(para)

        # Check for round-bracket citations — should be square brackets
        wrong = self.WRONG_CITE_PAREN.findall(text)
        if wrong:
            self.report.add(Violation(
                location=location,
                rule_ref="7.1.19.2",
                description=f"Ссылки на литературу оформлены в круглых скобках: {wrong[:5]}. Требуются квадратные скобки [N].",
                original=text[:120],
                auto_fixed=False
            ))


class TextStyleChecker:
    """Rule 7.2.3/7.2.4: forbidden patterns in text."""

    MINUS_SIGN = re.compile(r'(?<!\w)–\s*\d')        # dash before number (could be minus)
    DIAMETER_SYMBOL = re.compile(r'[ØøΦ]')            # diameter symbol in body text
    PERCENT_WITHOUT_NUM = re.compile(r'(?<!\d)\s*%')  # % without preceding number

    def __init__(self, report: CorrectionReport):
        self.report = report

    def check_paragraph(self, para, location: str):
        text = get_paragraph_text(para)
        if not text.strip():
            return

        if self.DIAMETER_SYMBOL.search(text):
            self.report.add(Violation(
                location=location,
                rule_ref="7.2.4",
                description="Символ диаметра Ø в тексте — следует писать слово «диаметр»",
                original=text[:120],
                auto_fixed=False
            ))


# ──────────────────────────────────────────────
# Main corrector
# ──────────────────────────────────────────────

class ThesisCorrector:
    """
    Loads rules from norm_control_rules.json and applies all
    correction modules to a .docx thesis file.
    """

    def __init__(self, rules_path: str):
        with open(rules_path, encoding="utf-8") as f:
            self.rules = json.load(f)

    def correct(self, input_path: str, output_path: str) -> CorrectionReport:
        doc = Document(input_path)
        report = CorrectionReport()

        # Instantiate correctors
        font_corrector = FontCorrector(self.rules, report)
        layout_corrector = PageLayoutCorrector(self.rules, report)
        para_corrector = ParagraphFormatCorrector(self.rules, report)
        heading_checker = HeadingChecker(self.rules, report)
        figure_checker = FigureCaptionChecker(report)
        table_checker = TableCaptionChecker(report)
        bib_checker = BibliographyChecker(report)
        text_checker = TextStyleChecker(report)

        # 1. Page layout (margins)
        layout_corrector.fix_document(doc)

        # 2. Paragraph-level corrections
        for i, para in enumerate(doc.paragraphs):
            loc = f"Абзац {i+1}"
            font_corrector.fix_paragraph(para, loc)
            para_corrector.fix_paragraph(para, loc)
            heading_checker.check_paragraph(para, loc)
            figure_checker.check_paragraph(para, loc)
            table_checker.check_paragraph(para, loc)
            bib_checker.check_paragraph(para, loc)
            text_checker.check_paragraph(para, loc)

        # 3. Tables
        for t_idx, table in enumerate(doc.tables):
            for r_idx, row in enumerate(table.rows):
                for c_idx, cell in enumerate(row.cells):
                    for p_idx, para in enumerate(cell.paragraphs):
                        loc = f"Таблица {t_idx+1} / строка {r_idx+1} / ячейка {c_idx+1} / абз. {p_idx+1}"
                        font_corrector.fix_paragraph(para, loc)

        # Save corrected document
        doc.save(output_path)
        return report


# ──────────────────────────────────────────────
# CLI usage
# ──────────────────────────────────────────────

if __name__ == "__main__":
    import sys

    if len(sys.argv) < 3:
        print("Usage: python thesis_corrector.py <input.docx> <output.docx> [rules.json]")
        sys.exit(1)

    input_path = sys.argv[1]
    output_path = sys.argv[2]
    rules_path = sys.argv[3] if len(sys.argv) > 3 else "data/norm_control_rules.json"

    corrector = ThesisCorrector(rules_path)
    report = corrector.correct(input_path, output_path)

    print(report.summary())

    report_path = output_path.replace(".docx", "_report.txt")
    with open(report_path, "w", encoding="utf-8") as f:
        f.write(report.summary())
    print(f"\nОтчёт сохранён: {report_path}")
    print(f"Исправленный файл: {output_path}")
