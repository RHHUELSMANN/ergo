"""Gestaltung des Versicherungsangebots mit der bestehenden Word-Vorlage.

Die Schnittstelle export_doc(template_path, output_path, daten) und die
Platzhalterzuordnung bleiben unverändert. Es werden keine Tarife berechnet.
"""

from docx import Document
from docx.shared import Pt, RGBColor, Inches, Cm
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.shared import qn
from docx.oxml import OxmlElement
from docx.text.paragraph import Paragraph


FONT = "Calibri"
FONT_SIZE = 11
TEXT = "000000"
MUTED = "4A4A4A"
BORDER = "B8B8B8"


def _font(run, bold=False, color=TEXT, size=FONT_SIZE):
    run.font.name = FONT
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.italic = False
    run.font.color.rgb = RGBColor.from_string(color)
    fonts = run._element.get_or_add_rPr().get_or_add_rFonts()
    for name in ("ascii", "hAnsi", "eastAsia", "cs"):
        fonts.set(qn("w:" + name), FONT)


def _shade(cell, fill):
    props = cell._tc.get_or_add_tcPr()
    for old in list(props.findall(qn("w:shd"))):
        props.remove(old)
    shade = OxmlElement("w:shd")
    shade.set(qn("w:fill"), fill)
    shade.set(qn("w:val"), "clear")
    props.append(shade)


def _style_table(table):
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    # Bestehende drei Spalten; einheitliche Breiten für alle Tarifblöcke.
    if len(table.columns) >= 3:
        for col, width in zip(table.columns, (4.5, 6.25, 6.25)):
            col.width = Cm(width)

    props = table._tbl.tblPr
    table_width = props.find(qn("w:tblW"))
    if table_width is not None:
        table_width.set(qn("w:type"), "dxa")
        table_width.set(qn("w:w"), str(round(17.0 / 2.54 * 1440)))
    for tag in ("tblInd", "tblCellSpacing"):
        for old in list(props.findall(qn("w:" + tag))):
            props.remove(old)
    for old in list(props.findall(qn("w:tblBorders"))):
        props.remove(old)
    borders = OxmlElement("w:tblBorders")
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        item = OxmlElement("w:" + edge)
        item.set(qn("w:val"), "single")
        item.set(qn("w:sz"), "4")
        item.set(qn("w:color"), BORDER)
        borders.append(item)
    props.append(borders)

    for row_index, row in enumerate(table.rows):
        # Keine festen Zeilenhöhen: Text darf bei Bedarf umbrechen.
        row_props = row._tr.get_or_add_trPr()
        for old in list(row_props.findall(qn("w:trHeight"))):
            row_props.remove(old)
        if row_props.find(qn("w:cantSplit")) is None:
            row_props.append(OxmlElement("w:cantSplit"))
        if row_index == 0 and row_props.find(qn("w:tblHeader")) is None:
            row_props.append(OxmlElement("w:tblHeader"))

        for col_index, cell in enumerate(row.cells):
            cell.width = table.columns[col_index].width
            cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
            cell_props = cell._tc.get_or_add_tcPr()
            # Beschriftungen erhalten eine Zeile, auch in schmaleren Vorschauen.
            no_wrap = cell_props.find(qn("w:noWrap"))
            if col_index == 0 and no_wrap is None:
                cell_props.append(OxmlElement("w:noWrap"))
            elif col_index > 0 and no_wrap is not None:
                cell_props.remove(no_wrap)
            # Alte Zellrahmen können sonst die neuen Tabellenrahmen überdecken.
            for old in list(cell_props.findall(qn("w:tcBorders"))):
                cell_props.remove(old)
            cell_borders = OxmlElement("w:tcBorders")
            for edge in ("top", "left", "bottom", "right"):
                edge_border = OxmlElement("w:" + edge)
                edge_border.set(qn("w:val"), "single")
                edge_border.set(qn("w:sz"), "4")
                edge_border.set(qn("w:color"), BORDER)
                cell_borders.append(edge_border)
            cell_props.append(cell_borders)
            margins = cell_props.find(qn("w:tcMar"))
            if margins is None:
                margins = OxmlElement("w:tcMar")
                cell_props.append(margins)
            for edge, value in (("top", "130"), ("bottom", "130"),
                                ("left", "120"), ("right", "120")):
                item = margins.find(qn("w:" + edge))
                if item is None:
                    item = OxmlElement("w:" + edge)
                    margins.append(item)
                item.set(qn("w:w"), value)
                item.set(qn("w:type"), "dxa")

            _shade(cell, "ECECEC" if row_index == 0 else
                   "F7F7F7" if row_index % 2 == 0 else "FFFFFF")
            for para in cell.paragraphs:
                para.alignment = (WD_ALIGN_PARAGRAPH.LEFT if col_index == 0
                                  else WD_ALIGN_PARAGRAPH.CENTER)
                pf = para.paragraph_format
                pf.space_before = Pt(0)
                pf.space_after = Pt(0)
                pf.line_spacing = 1.05
                pf.keep_with_next = row_index < len(table.rows) - 1
                pf.widow_control = True
                para_props = para._element.get_or_add_pPr()
                hyphens = para_props.find(qn("w:suppressAutoHyphens"))
                if hyphens is None:
                    hyphens = OxmlElement("w:suppressAutoHyphens")
                    para_props.append(hyphens)
                hyphens.set(qn("w:val"), "1")
                is_placeholder = "{" in para.text and "}" in para.text
                for run in para.runs:
                    if is_placeholder:
                        _font(run)
                    elif row_index == 0:
                        _font(run, bold=True, size=9.5)
                    elif col_index == 0:
                        _font(run, size=9.5)
                    else:
                        _font(run, size=10)
                    # Lange technische Felder bleiben bei 11 Punkt und werden
                    # nur horizontal enger gesetzt. Befüllte Preise bleiben normal.
                    if is_placeholder:
                        props = run._element.get_or_add_rPr()
                        scale = OxmlElement("w:w")
                        scale.set(qn("w:val"), "70")
                        props.append(scale)


def _style_notes(para, headings):
    value = para.text.strip()
    pf = para.paragraph_format
    pf.line_spacing = 1.16
    pf.space_before = Pt(0)
    pf.space_after = Pt(10)
    pf.keep_together = True
    pf.widow_control = True
    if value == "Hinweise zu Ihrem Reiseschutz":
        pf.line_spacing = 1.1
        pf.space_after = Pt(17)
        for run in para.runs:
            _font(run, bold=True)
    elif value in headings:
        pf.line_spacing = 1.1
        pf.space_before = Pt(23)
        pf.space_after = Pt(11)
        pf.keep_with_next = True
        pf.tab_stops.clear_all()
        pf.tab_stops.add_tab_stop(Cm(0.3))
        para.text = "\t" + value
        for run in para.runs:
            _font(run, bold=True)
        # Ein ruhiger grauer Balken trennt die drei Versicherungen deutlich.
        props = para._element.get_or_add_pPr()
        shade = props.find(qn("w:shd"))
        if shade is None:
            shade = OxmlElement("w:shd")
            props.append(shade)
        shade.set(qn("w:fill"), "EEEEEE")
        shade.set(qn("w:val"), "clear")
        borders = props.find(qn("w:pBdr"))
        if borders is None:
            borders = OxmlElement("w:pBdr")
            props.append(borders)
        for edge in ("top", "bottom"):
            line = borders.find(qn("w:" + edge))
            if line is None:
                line = OxmlElement("w:" + edge)
                borders.append(line)
            line.set(qn("w:val"), "single")
            line.set(qn("w:sz"), "8")
            line.set(qn("w:color"), "EEEEEE")
            line.set(qn("w:space"), "5")
    elif value in ("Weitere Hinweise", "Abschluss"):
        pf.space_before = Pt(22)
        pf.space_after = Pt(9)
        pf.keep_with_next = True
        for run in para.runs:
            _font(run, bold=True)
        # Eine feine Linie gibt den beiden Schlussabschnitten Halt im Druck.
        borders = para._element.get_or_add_pPr().find(qn("w:pBdr"))
        if borders is None:
            borders = OxmlElement("w:pBdr")
            para._element.get_or_add_pPr().append(borders)
        bottom = borders.find(qn("w:bottom"))
        if bottom is None:
            bottom = OxmlElement("w:bottom")
            borders.append(bottom)
        bottom.set(qn("w:val"), "single")
        bottom.set(qn("w:sz"), "4")
        bottom.set(qn("w:color"), "B8B8B8")
        bottom.set(qn("w:space"), "5")
    elif value.startswith("Jahresversicherungen verlängern sich"):
        pf.space_before = Pt(0)
        pf.space_after = Pt(10)
        for run in para.runs:
            _font(run)
    elif value.startswith("Prüfen Sie bei Auslandsreisen"):
        pf.space_after = Pt(7)
        for run in para.runs:
            _font(run)
    elif value.startswith("Für den Abschluss benötigen"):
        pf.space_before = Pt(0)
        pf.space_after = Pt(5)
        for run in para.runs:
            _font(run)
    elif value.startswith("Für Rückfragen stehen wir Ihnen gerne zur Verfügung"):
        pf.space_before = Pt(0)
        pf.space_after = Pt(5)
        for run in para.runs:
            _font(run)
    elif value == "Ihr Reisebüro Hülsmann":
        pf.space_before = Pt(13)
        pf.space_after = Pt(0)
        for run in para.runs:
            _font(run)
    else:
        for run in para.runs:
            _font(run)
        # Die drei Angaben lassen sich in jedem Produktblock schnell finden.
        if value.startswith(("Selbstbehalt:", "Laufzeit:", "Abschlussfrist:")):
            lead, rest = para.text.split(":", 1)
            para.clear()
            _font(para.add_run(lead + ":"), color=MUTED)
            _font(para.add_run("\t" + rest.lstrip()))
            # Hängender Einzug: auch Folgezeilen beginnen an derselben Textkante.
            pf.left_indent = Cm(3.1)
            pf.first_line_indent = Cm(-3.1)
            pf.tab_stops.clear_all()
            pf.tab_stops.add_tab_stop(Cm(3.1))


def export_doc(template_path, output_path, daten):
    doc = Document(template_path)
    # Auch geerbte Titellinien der Vorlage entfernen.
    for style_name in ("Normal", "Title", "Heading 1", "Heading 2"):
        style = doc.styles[style_name]
        style.font.name = FONT
        style.font.size = Pt(FONT_SIZE)
        style.font.color.rgb = RGBColor(0, 0, 0)
        fonts = style.element.get_or_add_rPr().get_or_add_rFonts()
        for name in ("ascii", "hAnsi", "eastAsia", "cs"):
            fonts.set(qn("w:" + name), FONT)
        props = style.element.find(qn("w:pPr"))
        if props is not None:
            for border in list(props.findall(qn("w:pBdr"))):
                props.remove(border)

    def ersetze(text):
        return text if not text else text.format(**daten)

    def einfuegen_danach(para, value):
        element = OxmlElement("w:p")
        para._p.addnext(element)
        new_para = Paragraph(element, para._parent)
        new_para.style = para.style
        new_para.text = value
        return new_para

    def entferne_leerzeilen_davor(para):
        previous = para._p.getprevious()
        while previous is not None and previous.tag == qn("w:p"):
            old_para = Paragraph(previous, para._parent)
            # Ein Seitenumbruch darf nicht versehentlich entfernt werden.
            if old_para.text.strip() or previous.xpath(".//w:br") or previous.xpath(".//w:pageBreakBefore"):
                break
            preceding = previous.getprevious()
            previous.getparent().remove(previous)
            previous = preceding

    # Bestehende Ersetzung unverändert: dieselben Felder und Zellen.
    for p in doc.paragraphs:
        p.text = ersetze(p.text)
        # Auch ältere Versionen der Word-Vorlage erhalten dieselben Zeilen.
        if p.text == "Einmalversicherungen: Beitrag für diese Reise. Jahresversicherungen: Jahresbeitrag.":
            p.text = "Einmalversicherungen: Beitrag für diese Reise."
            einfuegen_danach(p, "Jahresversicherungen: Jahresbeitrag.")
        elif p.text == "Für den Abschluss benötigen wir Ihre IBAN. Bei Fragen beraten wir Sie gerne.":
            p.text = "Für den Abschluss benötigen wir Ihre IBAN."
            einfuegen_danach(p, "Für Rückfragen stehen wir Ihnen gerne zur Verfügung.")
        elif p.text == "Bei Fragen beraten wir Sie gerne.":
            p.text = "Für Rückfragen stehen wir Ihnen gerne zur Verfügung."
    for p in doc.paragraphs:
        if p.text in ("Für den Abschluss benötigen wir Ihre IBAN.",
                      "Für Rückfragen stehen wir Ihnen gerne zur Verfügung.",
                      "Ihr Reisebüro Hülsmann"):
            entferne_leerzeilen_davor(p)
    # Die Schlussinformationen bekommen auch mit einer älteren Vorlage
    # dieselbe ruhige Gliederung.
    for p in list(doc.paragraphs):
        title = None
        if p.text.startswith("Jahresversicherungen verlängern sich"):
            title = "Weitere Hinweise"
        elif p.text.startswith("Für den Abschluss benötigen wir Ihre IBAN"):
            title = "Abschluss"
        if title:
            previous = p._p.getprevious()
            if previous is None or Paragraph(previous, p._parent).text.strip() != title:
                element = OxmlElement("w:p")
                p._p.addprevious(element)
                heading = Paragraph(element, p._parent)
                heading.text = title
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                cell.text = ersetze(cell.text)

    headings = {
        "Welche Versicherungen sind wichtig", "Abschlussfristen",
        "Reiserücktritts-Versicherung", "Reisekranken-Versicherung",
        "RundumSorglos-Schutz", "Selbstbeteiligung und Laufzeiten",
        "Reiserücktritt und Laufzeiten", "Reisekranken und Laufzeiten",
        "So schließen Sie Ihren Reiseschutz ab",
        "Wichtig für Ihre Reise", "Abschluss und Rückfragen",
    }
    notes_started = False
    for para in doc.paragraphs:
        value = para.text.strip()
        pf = para.paragraph_format
        if value.startswith("Versicherungsangebot für "):
            para.style = doc.styles["Title"]
            para.alignment = WD_ALIGN_PARAGRAPH.LEFT
            pf.space_before = Pt(0)
            pf.space_after = Pt(16)
            pf.keep_with_next = True
            for border in list(para._element.get_or_add_pPr().findall(qn("w:pBdr"))):
                para._element.get_or_add_pPr().remove(border)
            for run in para.runs:
                _font(run, bold=True, color="000000")
        elif value == "Hinweise zu Ihrem Reiseschutz":
            para.style = doc.styles["Heading 1"]
            pf.space_before = Pt(0)
            pf.space_after = Pt(12)
            pf.keep_with_next = True
            for run in para.runs:
                _font(run, bold=True, color="000000")
        elif value in headings:
            para.style = doc.styles["Heading 2"]
            para.alignment = WD_ALIGN_PARAGRAPH.LEFT
            pf.space_before = Pt(18)
            pf.space_after = Pt(7)
            pf.keep_with_next = True
            for run in para.runs:
                _font(run, bold=True, color="000000")
        elif value.startswith("Die ausgewiesenen Beiträge gelten"):
            pf.space_before = Pt(14)
            pf.space_after = Pt(5)
            pf.keep_with_next = True
            for run in para.runs:
                _font(run, bold=True)
        elif value.startswith("Einmalversicherungen:"):
            pf.space_after = Pt(2)
            pf.keep_with_next = True
            for run in para.runs:
                _font(run, color=MUTED)
        elif value.startswith("Jahresversicherungen:"):
            pf.space_after = Pt(10)
            for run in para.runs:
                _font(run, color=MUTED)
        elif value.startswith("Diese Übersicht wurde automatisch vom Reisebüro Hülsmann"):
            pf.space_before = Pt(10)
            pf.space_after = Pt(0)
            for run in para.runs:
                _font(run, color=MUTED)
        elif value.startswith(("Reisezeitraum:", "Reisedatum:", "Reisepreis:",
                               "Anzahl Reisende:", "Alter der ältesten Person:",
                               "Geltungsbereich:")):
            pf.space_after = Pt(5)
            pf.keep_with_next = True
            lead, rest = para.text.split(":", 1)
            para.text = lead + ":\t" + rest.lstrip()
            pf.left_indent = Cm(4.9)
            pf.first_line_indent = Cm(-4.9)
            pf.tab_stops.clear_all()
            pf.tab_stops.add_tab_stop(Cm(4.9))
            for run in para.runs:
                _font(run)
        elif value:
            for run in para.runs:
                _font(run)
        if value == "Hinweise zu Ihrem Reiseschutz":
            notes_started = True
        if notes_started:
            _style_notes(para, headings)

    for table in doc.tables:
        _style_table(table)

    # Beide Fußzeilen haben unterschiedliche Aufgaben: Rechtshinweis und Kontakt.
    for section in doc.sections:
        # Das alte transparente Bildlogo überlagert seinen eigenen Schriftzug
        # in Word. Ein schlichter Textkopf bleibt auch im Schwarzweißdruck klar.
        for is_first, part in ((False, section.header), (True, section.first_page_header)):
            if is_first:
                para = part.paragraphs[0]
                para.clear()
                para.alignment = WD_ALIGN_PARAGRAPH.CENTER
                para.paragraph_format.space_after = Pt(0)
                _font(para.add_run("Reisebüro Hülsmann"), bold=True)
                _font(para.add_run("\nWir beraten Sie gerne!"))
                for extra in list(part.paragraphs[1:]):
                    extra._element.getparent().remove(extra._element)
            else:
                for para in part.paragraphs:
                    for run in para.runs:
                        _font(run)
        for part in (section.footer, section.first_page_footer):
            for para in part.paragraphs:
                legal = para.text.startswith("Diese Übersicht wurde automatisch")
                para.alignment = (WD_ALIGN_PARAGRAPH.LEFT if legal
                                  else WD_ALIGN_PARAGRAPH.CENTER)
                para.paragraph_format.line_spacing = 1.05
                para.paragraph_format.space_after = Pt(0)
                for run in para.runs:
                    _font(run, color=MUTED, size=8.5 if legal else 9.5)
                    run.font.italic = legal

    doc.save(output_path)
    return output_path
