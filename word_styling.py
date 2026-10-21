"""Werte in die Word-Vorlage einsetzen und den Abschlussblock formatieren.

Die Vorlage bestimmt die Gestaltung. Die Abstaende des Abschlussblocks
werden zuletzt im Export festgelegt: 32, 8 und 16 Punkt vor den Absaetzen.
Die Schnittstelle export_doc(template_path, output_path, daten) bleibt erhalten.
"""

from docx import Document
from docx.oxml.ns import qn
from docx.shared import Pt
from docx.text.paragraph import Paragraph


def _formatiere_abschluss(doc):
    abstaende = {
        "Für den Abschluss benötigen wir Ihre IBAN.": 32,
        "Für Rückfragen stehen wir Ihnen gerne zur Verfügung.": 8,
        "Ihr Reisebüro Hülsmann": 16,
    }
    for absatz in list(doc.paragraphs):
        text = absatz.text.strip()
        if text not in abstaende:
            continue

        # Echte Absatzabstaende statt zusaetzlicher Leerabsaetze verwenden.
        davor = absatz._p.getprevious()
        while davor is not None and davor.tag == qn("w:p"):
            if Paragraph(davor, absatz._parent).text.strip():
                break
            if davor.xpath(".//w:drawing | .//w:pict | .//w:br | .//w:sectPr | .//w:pageBreakBefore"):
                break
            vorher = davor.getprevious()
            davor.getparent().remove(davor)
            davor = vorher

        # Eine alte Abschluss-Ueberschrift gehoert nicht zum vereinbarten Block.
        if text == "Für den Abschluss benötigen wir Ihre IBAN.":
            if davor is not None and davor.tag == qn("w:p"):
                if Paragraph(davor, absatz._parent).text.strip() == "Abschluss":
                    davor.getparent().remove(davor)

        formatierung = absatz.paragraph_format
        formatierung.space_before = Pt(abstaende[text])
        formatierung.space_after = Pt(0)
        formatierung.line_spacing = 1.1
        formatierung.keep_with_next = text != "Ihr Reisebüro Hülsmann"
        formatierung.keep_together = True


def export_doc(template_path, output_path, daten):
    doc = Document(template_path)

    def ersetze_absatz(absatz, preis=False):
        for run in absatz.runs:
            if "{" not in run.text:
                continue
            run.text = run.text.format(**daten)
            if preis:
                run.font.size = Pt(10)
                props = run._element.get_or_add_rPr()
                for width in list(props.findall(qn("w:w"))):
                    props.remove(width)

    for absatz in doc.paragraphs:
        ersetze_absatz(absatz)

    for tabelle in doc.tables:
        for zeile in tabelle.rows:
            for zelle in zeile.cells:
                for absatz in zelle.paragraphs:
                    ersetze_absatz(absatz, preis=True)

    _formatiere_abschluss(doc)
    doc.save(output_path)
    return output_path
