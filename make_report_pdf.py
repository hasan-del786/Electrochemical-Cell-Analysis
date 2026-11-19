"""Builds Electrochemical_Cell_Analysis_Report.pdf from electrochemical_cell_results.csv and the PNG charts.
Run electrochemical_cell_analysis.py first.   Install: pip install reportlab pandas"""
import pandas as pd
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import cm
from reportlab.platypus import (Image, KeepTogether, PageBreak, Paragraph, SimpleDocTemplate,
                                Spacer, Table, TableStyle)

df = pd.read_csv("electrochemical_cell_results.csv")
sd_r = df.SD_V.round(3)
hi, lo = df.loc[df.Mean_V.idxmax()], df.loc[df.Mean_V.idxmin()]
low_sd = ", ".join(c.replace("Cell ", "") for c in df.Cell[sd_r == sd_r.min()])
high_sd = ", ".join(c.replace("Cell ", "") for c in df.Cell[sd_r == sd_r.max()])
cvmax, cvmin = df.loc[df.CV_percent.idxmax()], df.loc[df.CV_percent.idxmin()]

ss = getSampleStyleSheet()
body = ParagraphStyle("b", parent=ss["Normal"], fontName="Helvetica", fontSize=10, leading=14.5, spaceAfter=6)
h1 = ParagraphStyle("h1", parent=ss["Heading1"], fontName="Helvetica-Bold", fontSize=14, textColor=colors.HexColor("#1F4E79"), spaceBefore=10, spaceAfter=6)
title = ParagraphStyle("t", parent=ss["Title"], fontName="Helvetica-Bold", fontSize=20, leading=25, textColor=colors.HexColor("#1F4E79"))
sub = ParagraphStyle("s", parent=body, alignment=TA_CENTER, textColor=colors.HexColor("#444444"))
small = ParagraphStyle("sm", parent=body, fontSize=8.5, leading=11, textColor=colors.HexColor("#444444"))
bullet = ParagraphStyle("bl", parent=body, leftIndent=14, bulletIndent=2, spaceAfter=3)
cap = ParagraphStyle("cap", parent=small, alignment=TA_CENTER, spaceAfter=10)

def table(data, widths, hdr=colors.HexColor("#1F4E79")):
    t = Table(data, colWidths=widths, repeatRows=1)
    t.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), hdr), ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"), ("FONTNAME", (0, 1), (-1, -1), "Helvetica"),
        ("FONTSIZE", (0, 0), (-1, -1), 9), ("ALIGN", (0, 0), (-1, -1), "CENTER"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"), ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#999999")),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#EEF3F8")]),
        ("TOPPADDING", (0, 0), (-1, -1), 4), ("BOTTOMPADDING", (0, 0), (-1, -1), 4)]))
    return t

def footer(canvas, doc):
    canvas.saveState(); canvas.setFont("Helvetica", 8); canvas.setFillColor(colors.HexColor("#666666"))
    canvas.drawString(2 * cm, 1.2 * cm, "Electrochemical Cell Analysis - simulated data for learning and analysis")
    canvas.drawRightString(A4[0] - 2 * cm, 1.2 * cm, f"Page {doc.page}"); canvas.restoreState()

S = []
S += [Spacer(1, 0.4 * cm), Paragraph("Electrochemical Cell Analysis Using Python and Excel", title),
      Paragraph("Mean voltage, sample standard deviation and variation across 10 illustrative cells", sub),
      Paragraph("Chemistry data-analysis project", sub), Spacer(1, 0.3 * cm)]

S += [Paragraph("1. Aim", h1),
      Paragraph("To calculate the mean voltage, sample standard deviation (SD) and relative variation for three repeated "
                "voltage readings on each of 10 illustrative electrochemical cells, and to present the results in "
                "Excel and Python so they can be reused in a chemistry portfolio.", body),
      Paragraph("<b>Data note.</b> All voltages are <b>simulated values for learning and analysis. They are not experimental "
                "results</b> and no laboratory measurements were made for this project.", body)]

S += [Paragraph("2. Data", h1),
      Paragraph("Three repeated readings per cell. Mean and SD are calculated from the three readings.", body)]
rows = [["Cell", "Reading 1 (V)", "Reading 2 (V)", "Reading 3 (V)", "Mean (V)", "SD (V)", "CV (%)"]]
for _, r in df.iterrows():
    rows.append([r.Cell, f"{r.Reading_1_V:.2f}", f"{r.Reading_2_V:.2f}", f"{r.Reading_3_V:.2f}",
                 f"{r.Mean_V:.2f}", f"{r.SD_V:.2f}", f"{r.CV_percent:.2f}"])
S += [table(rows, [2.2*cm, 2.7*cm, 2.7*cm, 2.7*cm, 2.2*cm, 2.0*cm, 2.0*cm])]

S += [Paragraph("3. Methods", h1),
      Paragraph("<b>Mean voltage</b> = (x1 + x2 + x3) / 3. In Excel: <font face='Courier'>=AVERAGE(B2:D2)</font> in E2.", body),
      Paragraph("<b>Sample standard deviation</b>: s = sqrt[ sum (xi - mean)^2 / (n - 1) ], with n = 3. "
                "In Excel: <font face='Courier'>=STDEV.S(B2:D2)</font> in F2; in Python: "
                "<font face='Courier'>DataFrame.std(axis=1, ddof=1)</font>. A smaller SD means the repeated readings agree more closely.", body),
      Paragraph("<b>Coefficient of variation (CV)</b> = SD / mean x 100. It expresses the spread relative to the voltage, "
                "so cells with different voltages can be compared fairly.", body),
      Paragraph("Formulas are entered in row 2 and filled down to row 11. The Python script reproduces the Excel values exactly.", body)]

S += [PageBreak(), Paragraph("4. Results", h1)]
res = [["Metric", "Value", "Cell(s)"],
       ["Highest mean voltage", f"{hi.Mean_V:.2f} V", hi.Cell],
       ["Lowest mean voltage", f"{lo.Mean_V:.2f} V", lo.Cell],
       ["Mean of the 10 cell means", f"{df.Mean_V.mean():.3f} V", "-"],
       ["Lowest SD", f"{sd_r.min():.2f} V", f"Cells {low_sd}"],
       ["Highest SD", f"{sd_r.max():.2f} V", f"Cells {high_sd}"],
       ["Highest CV", f"{cvmax.CV_percent:.1f} %", cvmax.Cell],
       ["Lowest CV", f"{cvmin.CV_percent:.2f} %", cvmin.Cell]]
S += [table(res, [6.5*cm, 3.5*cm, 5.5*cm]), Spacer(1, 0.3 * cm)]

S += [Image("chart_mean_voltage.png", width=14.5*cm, height=14.5*cm*4.2/7.5),
      Paragraph("Figure 1. Mean voltage of each cell; error bars show the sample SD.", cap),
      Image("chart_standard_deviation.png", width=14.5*cm, height=14.5*cm*4.2/7.5),
      Paragraph("Figure 2. Sample SD of each cell. Green = lowest SD, red = highest SD.", cap)]

S += [PageBreak(), Paragraph("5. Interpretation", h1)]
pts = [
    f"<b>Voltage range.</b> Mean voltages span {lo.Mean_V:.2f} V ({lo.Cell}) to {hi.Mean_V:.2f} V ({hi.Cell}); "
    f"the mean of the 10 cell means is {df.Mean_V.mean():.3f} V.",
    f"<b>Consistency.</b> Cells {low_sd} have the lowest SD ({sd_r.min():.2f} V), so their readings are the most consistent. "
    f"Cells {high_sd} have the highest SD ({sd_r.max():.2f} V).",
    f"<b>Relative variation.</b> {cvmax.Cell} has the largest CV ({cvmax.CV_percent:.1f} %) because a 0.05 V SD is large compared with a "
    f"0.50 V mean. {cvmin.Cell} has the smallest CV ({cvmin.CV_percent:.2f} %): the same 0.02 V SD is small compared with 1.50 V. "
    "Absolute SD and CV can therefore rank cells differently.",
    "<b>Voltage and spread are not linked here.</b> The highest-voltage cell is among the most consistent, but the SD groups "
    "do not follow the voltage order, so no trend between voltage level and spread can be claimed from these 10 cells.",
    "<b>Why simulated spreads look neat.</b> In each cell the three readings are evenly spaced, so the SD equals the step "
    "between readings (0.02, 0.03 or 0.05 V). Real measurements would not be this regular.",
]
S += [Paragraph(p, bullet, bulletText="\u2022") for p in pts]
S += [Image("chart_repeated_readings.png", width=12.5*cm, height=12.5*cm*4.2/7.5),
      Paragraph("Figure 3. Individual readings (markers) and cell means (black dashes).", cap)]

S += [Paragraph("6. Limitations", h1)]
lims = ["The data are simulated, so the results describe the method, not real cell behaviour.",
        "With only n = 3 readings per cell, each SD is a rough estimate; it is itself uncertain.",
        "No chemistry (electrode materials, electrolyte, temperature, concentration) is modelled; "
        "differences between cells are not explained, only quantified.",
        "No significance tests were run. A next step would be to measure real cells and compare them, for example with a t-test or ANOVA."]
S += [Paragraph(p, bullet, bulletText="\u2022") for p in lims]

S += [Paragraph("7. Files and reuse", h1),
      Paragraph("The accompanying zip contains the data CSV, the Python analysis script, the Excel workbook "
                "(sheets: Cell Data, Summary, Charts; all results use live formulas), the three charts, and this report. "
                "Change the blue input readings in Excel or edit the CSV and re-run the script to analyse new data.", body),
      Paragraph("<b>Suggested portfolio / CV wording (accurate):</b> \"Electrochemical cell data analysis (simulated dataset): "
                "calculated mean voltage, sample standard deviation and coefficient of variation for 10 cells using Python "
                "(pandas, matplotlib) and Excel formulas (AVERAGE, STDEV.S); produced charts and a summary report.\"", body)]

SimpleDocTemplate("Electrochemical_Cell_Analysis_Report.pdf", pagesize=A4, leftMargin=2*cm, rightMargin=2*cm,
                  topMargin=1.8*cm, bottomMargin=2*cm, title="Electrochemical Cell Analysis Using Python and Excel",
                  author="Hasan Ali").build(S, onFirstPage=footer, onLaterPages=footer)
print("Saved Electrochemical_Cell_Analysis_Report.pdf")
