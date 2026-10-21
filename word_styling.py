"""Werte in die gestaltete Word-Vorlage einsetzen.

Die Gestaltung stammt aus angebot.docx. Dieser Export fügt keine
Überschriften hinzu und überschreibt keine Fußzeilen oder Absatzabstände.

Die Schnittstelle entspricht dem urspruenglich funktionierenden Export:
export_doc("angebot.docx", ausgabepfad, daten).
"""

from docx import Document
from docx.oxml.ns import qn
from docx.shared import Pt


def export_doc(template_path, output_path, daten):
    doc = Document(template_path)

    def ersetze_absatz(absatz, preis=False):
        for run in absatz.runs:
            if "{" not in run.text:
                continue
            run.text = run.text.format(**daten)
            if preis:
                run.font.size = Pt(10)
                # Nur die langen Platzhalternamen waren in der Vorlage gestaucht.
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

    doc.save(output_path)
    return output_path
