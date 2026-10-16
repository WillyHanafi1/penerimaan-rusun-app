"""Retain cached fixed-tariff arithmetic when openpyxl saves a daily master."""

from io import BytesIO
import re
from xml.etree import ElementTree as ET
from zipfile import ZipFile

NS = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'


def save_master(workbook, path):
    # Fixed prices such as =280000+90000 do not depend on daily payment edits.
    # Preserve their existing Excel results so tomorrow's data_only reader can
    # still read the tariff without requiring an intervening Excel recalculation.
    cached = {}
    with ZipFile(path) as original:
        for name in original.namelist():
            if not re.fullmatch(r'xl/worksheets/sheet\d+\.xml', name):
                continue
            root = ET.fromstring(original.read(name))
            shared = {formula.get('si'): formula.text for formula in root.iter(NS + 'f')
                      if formula.get('t') == 'shared' and formula.text}
            for cell in root.iter(NS + 'c'):
                formula, value = cell.find(NS + 'f'), cell.find(NS + 'v')
                expression = (formula.text or shared.get(formula.get('si'))) if formula is not None else None
                if (expression and value is not None and value.text
                        and re.fullmatch(r'[\d\s.+*/^%(),-]+', expression)):
                    cached[(name, cell.get('r'))] = (expression, value.text)
    saved = BytesIO()
    workbook.save(saved)
    saved.seek(0)
    output = BytesIO()
    with ZipFile(saved) as archive, ZipFile(output, 'w') as result:
        for entry in archive.infolist():
            content = archive.read(entry.filename)
            if entry.filename.startswith('xl/worksheets/') and entry.filename.endswith('.xml'):
                root = ET.fromstring(content)
                changed = False
                for cell in root.iter(NS + 'c'):
                    prior = cached.get((entry.filename, cell.get('r')))
                    formula = cell.find(NS + 'f')
                    if prior and formula is not None and formula.text == prior[0]:
                        value = cell.find(NS + 'v')
                        if value is None:
                            value = ET.SubElement(cell, NS + 'v')
                        value.text = prior[1]
                        changed = True
                if changed:
                    content = ET.tostring(root, encoding='utf-8')
            result.writestr(entry, content)
    with open(path, 'wb') as target:
        target.write(output.getvalue())
