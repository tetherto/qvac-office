import struct
import zlib
from pathlib import Path

from pptx import Presentation
from pptx.util import Inches, Pt

here = Path(__file__).parent


def png(width, height, rgb):
    def chunk(kind, data):
        body = kind + data
        return struct.pack(">I", len(data)) + body + struct.pack(">I", zlib.crc32(body))

    row = b"\x00" + bytes(rgb) * width
    raw = row * height
    return (
        b"\x89PNG\r\n\x1a\n"
        + chunk(b"IHDR", struct.pack(">IIBBBBB", width, height, 8, 2, 0, 0, 0))
        + chunk(b"IDAT", zlib.compress(raw))
        + chunk(b"IEND", b"")
    )


(here / "red.png").write_bytes(png(40, 20, (255, 0, 0)))

prs = Presentation()

cover = prs.slides.add_slide(prs.slide_layouts[0])
cover.shapes.title.text = "Why the Sky Is Blue"
cover.placeholders[1].text = "A short tour of Rayleigh scattering"

content = prs.slides.add_slide(prs.slide_layouts[1])
content.shapes.title.text = "Scattering"
body = content.placeholders[1].text_frame
body.text = "Sunlight holds every colour"
run = body.paragraphs[0].runs[0]
run.font.bold = True
run.font.size = Pt(24)
second = body.add_paragraph()
second.text = "Short wavelengths scatter more"
second.level = 1
content.notes_slide.notes_text_frame.text = "Mention the sunset question."

visual = prs.slides.add_slide(prs.slide_layouts[5])
visual.shapes.title.text = "Measurements"
visual.shapes.add_picture(str(here / "red.png"), Inches(1), Inches(2), width=Inches(2))
table = visual.shapes.add_table(2, 2, Inches(4), Inches(2), Inches(4), Inches(1)).table
table.cell(0, 0).text = "Colour"
table.cell(0, 1).text = "Wavelength"
table.cell(1, 0).text = "Blue"
table.cell(1, 1).text = "450 nm"
group = visual.shapes.add_group_shape()
box = group.shapes.add_textbox(Inches(1), Inches(5), Inches(3), Inches(1))
box.text_frame.text = "Grouped caption"

prs.save(here / "basic.pptx")

from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE

charts = Presentation()
bars = charts.slides.add_slide(charts.slide_layouts[5])
bars.shapes.title.text = "Revenue"
data = CategoryChartData()
data.categories = ["Q1", "Q2", "Q3"]
data.add_series("2025", (10.0, 12.5, 14.0))
data.add_series("2026", (11.0, 13.0, 15.5))
frame = bars.shapes.add_chart(
    XL_CHART_TYPE.COLUMN_CLUSTERED, Inches(1), Inches(1.5), Inches(8), Inches(5), data
)
frame.chart.has_title = True
frame.chart.chart_title.text_frame.text = "Revenue by quarter"
frame.chart.chart_title.text_frame.paragraphs[0].runs[0].font.bold = True

pie = charts.slides.add_slide(charts.slide_layouts[6])
share = CategoryChartData()
share.categories = ["North", "South"]
share.add_series("Share", (0.6, 0.4))
pie.shapes.add_chart(XL_CHART_TYPE.PIE, Inches(1), Inches(1), Inches(6), Inches(5), share)

charts.save(here / "charts.pptx")

from docx import Document as WordDocument
from docx.shared import Inches as WordInches

report = WordDocument()
report.add_heading("Quarterly Report", level=0)
report.add_paragraph("Prepared by the finance team.")
report.add_heading("Revenue", level=1)
intro = report.add_paragraph("Revenue grew ")
intro.add_run("twelve percent").bold = True
intro.add_run(" year over year.")
report.add_paragraph("North led the growth", style="List Bullet")
report.add_paragraph("South held flat", style="List Bullet")
report.add_heading("Regions", level=2)
report.add_paragraph("Every region reported.")
report.add_heading("Costs", level=1)
report.add_paragraph("Costs fell as contracts ended.")
grid = report.add_table(rows=2, cols=3)
grid.style = "Table Grid"
for i, v in enumerate(["Metric", "2025", "2026"]):
    grid.cell(0, i).text = v
for i, v in enumerate(["Margin", "25%", "36%"]):
    grid.cell(1, i).text = v
report.add_heading("Outlook", level=1)
report.add_paragraph("Steady.")
report.add_picture(str(here / "red.png"), width=WordInches(1))
report.sections[0].header.paragraphs[0].text = "Confidential"
report.save(here / "report.docx")

import datetime

from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.workbook.defined_name import DefinedName

book = Workbook()
sales = book.active
sales.title = "Sales"
sales["A1"] = "Quarterly sales"
sales.merge_cells("A1:D1")
sales.append(["Quarter", "Units", "Price", "Revenue"])
for cell in sales[2]:
    cell.font = Font(bold=True)
rows = [("Q1", 120, 9.5), ("Q2", 150, 9.75), ("Q3", 90, 10.0)]
for i, (quarter, units, price) in enumerate(rows, start=3):
    sales.append([quarter, units, price, f"=B{i}*C{i}"])
sales["A6"] = "Total"
sales["B6"] = "=SUM(B3:B5)"
sales["D6"] = "=SUM(D3:D5)"
sales["F2"] = "Margin"
sales["F3"] = 0.25
sales["F3"].number_format = "0.0%"
sales["G2"] = "Updated"
sales["G3"] = datetime.date(2026, 3, 31)
sales["G3"].number_format = "yyyy-mm-dd"
sales["H2"] = "Final"
sales["H3"] = True
sales.column_dimensions["A"].width = 18

notes = book.create_sheet("Notes")
notes["A1"] = "Double revenue"
notes["B1"] = "=Sales!D6*2"
notes["A2"] = "Plain text"

book.defined_names["TotalUnits"] = DefinedName("TotalUnits", attr_text="Sales!$B$6")
book.save(here / "sales.xlsx")

import xlsxwriter

shared = xlsxwriter.Workbook(str(here / "shared.xlsx"))
sheet = shared.add_worksheet("Budget")
bold = shared.add_format({"bold": True})
money = shared.add_format({"num_format": "#,##0.00"})
sheet.write_row(0, 0, ["Item", "Cost"], bold)
sheet.write_row(1, 0, ["Rent", 1200], None)
sheet.write_row(2, 0, ["Food", 450.5], None)
sheet.write(1, 1, 1200, money)
sheet.write(2, 1, 450.5, money)
sheet.write_formula(3, 1, "=SUM(B2:B3)", money, 1650.5)
sheet.write(3, 0, "Total", bold)
shared.close()
