"""Fictional deterministic workbooks. No random generation is used."""

from datetime import datetime
from io import BytesIO
from pathlib import Path
from zipfile import ZipFile, ZipInfo, ZIP_DEFLATED

from openpyxl import Workbook
from openpyxl.workbook.defined_name import DefinedName

FIXED_TIME = datetime(2020, 1, 1)


def save_deterministic(workbook, path):
    """Normalize OOXML properties and ZIP timestamps for byte repeatability."""
    workbook.properties.created = FIXED_TIME
    workbook.properties.modified = FIXED_TIME
    workbook.properties.creator = "Spreadsheet Reverse Engineer synthetic fixtures"
    buffer = BytesIO()
    workbook.save(buffer)
    with ZipFile(buffer) as source, ZipFile(path, "w", ZIP_DEFLATED) as target:
        for filename in sorted(source.namelist()):
            data = source.read(filename)
            if filename == "docProps/core.xml":
                import re
                data = re.sub(rb"(<dcterms:modified[^>]*>).*?(</dcterms:modified>)", rb"\g<1>2020-01-01T00:00:00Z\g<2>", data)
            info = ZipInfo(filename, (2020, 1, 1, 0, 0, 0))
            info.compress_type = ZIP_DEFLATED
            info.external_attr = 0o600 << 16
            target.writestr(info, data)


def base():
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "README"
    sheet["A1"] = "SYNTHETIC FICTIONAL FIXTURE. Not real finances."
    sheet["A2"] = "Descriptive analysis only; no correctness certification. Migration plan is a proposal."
    return workbook


def generate(out):
    directory = Path(out)
    directory.mkdir(parents=True, exist_ok=True)
    clean = base()
    inputs = clean.create_sheet("Inputs")
    inputs["A1"], inputs["B1"] = "Fictional units", "Fictional unit price"
    inputs["A2"], inputs["B2"] = 4, 12
    calc = clean.create_sheet("Calc")
    calc["A1"] = "Fictional revenue"
    calc["A2"] = "=Inputs!A2*Inputs!B2"
    clean.defined_names.add(DefinedName("OutputRevenue", attr_text="'Calc'!$A$2"))
    save_deterministic(clean, directory / "clean.xlsx")
    pricing = base()
    sheet = pricing.create_sheet("Pricing")
    for key, value in {"A1": "Fictional pricing model", "A2": "Basic", "A3": "Plus", "B2": 10,
                       "B3": 20, "C2": "=B2*1.2", "C3": "=B3*1.2", "D2": '=IF(B2>15,C2,B2)',
                       "E2": '=VLOOKUP(A2,Rates!A2:B3,2,FALSE)', "F2": "=SUM(C2:C3)",
                       "G2": "=TODAY()", "H2": "=TaxRate*B2", "J2": 999,
                       "C8": "=C9+1", "C9": "=C8+1", "C10": "=C8*B2",
                       "K2": "=Rates!B2+Tax!A2+Shipping!A2"}.items():
        sheet[key] = value
    rates = pricing.create_sheet("Rates")
    for key, value in {"A1": "Fictional tier", "B1": "Fictional rate", "A2": "Basic", "B2": 5,
                       "A3": "Plus", "B3": 8}.items():
        rates[key] = value
    pricing.create_sheet("Tax")["A2"] = 0.2
    pricing.create_sheet("Shipping")["A2"] = 3
    pricing.defined_names.add(DefinedName("TaxRate", attr_text="'Tax'!$A$2"))
    for name, ref in {"OutputTotal": "'Pricing'!$F$2", "OutputDecision": "'Pricing'!$D$2",
                      "OutputLookup": "'Pricing'!$E$2", "OutputTax": "'Pricing'!$H$2",
                      "OutputSpaghetti": "'Pricing'!$K$2", "OutputClock": "'Pricing'!$G$2"}.items():
        pricing.defined_names.add(DefinedName(name, attr_text=ref))
    save_deterministic(pricing, directory / "pricing.xlsx")
    return [str(directory / "clean.xlsx"), str(directory / "pricing.xlsx")]
