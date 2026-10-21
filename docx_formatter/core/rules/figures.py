from __future__ import annotations

from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.shared import Pt
from docx.text.paragraph import Paragraph
from docx.text.run import Run

from docx_formatter.core.context import DocContext
from docx_formatter.core.oxml import iter_block_items, set_run_font
from docx_formatter.core.profile import Profile
from docx_formatter.core.report import Change
from docx_formatter.core.units import lengths_equal, pt

_ALIGNMENTS = {
    "left": WD_ALIGN_PARAGRAPH.LEFT,
    "right": WD_ALIGN_PARAGRAPH.RIGHT,
    "center": WD_ALIGN_PARAGRAPH.CENTER,
    "justify": WD_ALIGN_PARAGRAPH.JUSTIFY,
}

_FONT_ATTRS = ("w:ascii", "w:hAnsi", "w:cs", "w:eastAsia")


def _run_matches_font(run: Run, font_name: str, size: Pt) -> bool:
    if run.font.size is None or not lengths_equal(run.font.size, size):
        return False
    rPr = run._element.find(qn("w:rPr"))
    if rPr is None:
        return False
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        return False
    return all(rFonts.get(qn(attr)) == font_name for attr in _FONT_ATTRS)


def _replace_paragraph_text(paragraph: Paragraph, text: str) -> None:
    """Переписать текст абзаца, сохранив оформление первого run'а.

    Остальные run'ы удаляются из XML (а не просто .text = "") — иначе на
    повторном прогоне в документе копились бы пустые <w:r/>.
    """
    runs = paragraph.runs
    if not runs:
        paragraph.add_run(text)
        return
    runs[0].text = text
    for run in runs[1:]:
        run._element.getparent().remove(run._element)


class CaptionsRule:
    """Подписи к рисункам и таблицам: приведение к шаблону профиля
    (captions.figure/captions.table), шрифт тела (сами подписи
    исключены из body_text — см. DocContext.caption_matches, иначе два
    правила начинают откатывать изменения друг друга на втором
    прогоне) и, если явно включено, сквозная перенумерация. По
    умолчанию (renumber=false) авторский номер не трогаем — иначе
    ломаются ручные ссылки вида «см. рисунок 3» в тексте отчёта
    (решение из плана).
    """

    id = "captions"

    def enabled_for(self, profile: Profile) -> bool:
        return True

    def apply(self, context: DocContext) -> list[Change]:
        profile = context.profile
        body = profile.body
        size = pt(body.size_pt)
        items = list(iter_block_items(context.document))
        changes: list[Change] = []

        for caption_index, match in context.caption_matches.items():
            caption_item = items[caption_index]
            style = profile.captions.figure if match.style_key == "figure" else profile.captions.table

            number = str(match.sequence) if style.renumber else match.parsed.number
            desired_text = style.template.format(n=number, text=match.parsed.description)

            changed = False
            if caption_item.text != desired_text:
                _replace_paragraph_text(caption_item, desired_text)
                changed = True

            for run in caption_item.runs:
                if not _run_matches_font(run, body.font, size):
                    set_run_font(run, body.font)
                    run.font.size = size
                    changed = True

            alignment = _ALIGNMENTS[style.alignment]
            if caption_item.paragraph_format.alignment != alignment:
                caption_item.paragraph_format.alignment = alignment
                changed = True

            if changed:
                kind = "рисунок" if match.style_key == "figure" else "таблица"
                changes.append(
                    Change(
                        self.id,
                        "info",
                        f"{kind} у блока #{match.anchor_index + 1}",
                        f"подпись приведена к шаблону: {desired_text}",
                    )
                )

        return changes
