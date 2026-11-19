"""
Electrochemical Cell Analysis Using Python and Excel
----------------------------------------------------
Reads three repeated voltage readings for each of 10 illustrative electrochemical cells
(electrochemical_cell_data.csv) and:
  1. calculates mean voltage, sample standard deviation (n - 1) and coefficient of variation,
  2. flags the most and least consistent cells,
  3. exports a CSV of results, three PNG charts and an Excel workbook that uses live
     Excel formulas (AVERAGE, STDEV.S) plus native Excel column charts.

DATA NOTE: the voltages are SIMULATED for learning and analysis. They are NOT experimental results.

Install:  pip install pandas numpy matplotlib openpyxl
Run:      python electrochemical_cell_analysis.py
"""
import numpy as np
import pandas as pd
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

READ = ["Reading_1_V", "Reading_2_V", "Reading_3_V"]

# ---------- 1. Load data and calculate statistics ----------
df = pd.read_csv("electrochemical_cell_data.csv")
df["Mean_V"] = df[READ].mean(axis=1)
df["SD_V"] = df[READ].std(axis=1, ddof=1)          # sample SD, same as Excel STDEV.S
df["CV_percent"] = df["SD_V"] / df["Mean_V"] * 100  # relative variation
df.round(4).to_csv("electrochemical_cell_results.csv", index=False)

sd_r = df["SD_V"].round(3)
print(df.round(3).to_string(index=False))
print(f"\nHighest mean : {df.loc[df.Mean_V.idxmax(), 'Cell']} ({df.Mean_V.max():.2f} V)")
print(f"Lowest mean  : {df.loc[df.Mean_V.idxmin(), 'Cell']} ({df.Mean_V.min():.2f} V)")
print(f"Lowest SD    : {', '.join(df.Cell[sd_r == sd_r.min()])} ({sd_r.min():.2f} V)")
print(f"Highest SD   : {', '.join(df.Cell[sd_r == sd_r.max()])} ({sd_r.max():.2f} V)")
print(f"Mean of the 10 cell means: {df.Mean_V.mean():.3f} V")

# ---------- 2. Charts (matplotlib) ----------
labels = [c.replace("Cell ", "") for c in df.Cell]
plt.rcParams.update({"font.family": "DejaVu Sans", "font.size": 10})

fig, ax = plt.subplots(figsize=(7.5, 4.2))
ax.bar(labels, df.Mean_V, yerr=df.SD_V, capsize=4, color="#2E75B6", edgecolor="black", linewidth=0.6)
for i, v in enumerate(df.Mean_V):
    ax.text(i, v + df.SD_V[i] + 0.03, f"{v:.2f}", ha="center", fontsize=8)
ax.set(xlabel="Cell", ylabel="Mean voltage (V)", title="Mean voltage of 10 simulated cells (error bars = sample SD)")
ax.set_ylim(0, 1.8); ax.grid(axis="y", alpha=0.3)
fig.tight_layout(); fig.savefig("chart_mean_voltage.png", dpi=200); plt.close(fig)

fig, ax = plt.subplots(figsize=(7.5, 4.2))
colors = ["#70AD47" if s == sd_r.min() else "#C00000" if s == sd_r.max() else "#7F7F7F" for s in sd_r]
ax.bar(labels, df.SD_V, color=colors, edgecolor="black", linewidth=0.6)
ax.set(xlabel="Cell", ylabel="Sample standard deviation (V)", title="Measurement consistency (green = lowest SD, red = highest SD)")
ax.grid(axis="y", alpha=0.3)
fig.tight_layout(); fig.savefig("chart_standard_deviation.png", dpi=200); plt.close(fig)

fig, ax = plt.subplots(figsize=(7.5, 4.2))
x = np.arange(len(df))
for col, mk in zip(READ, ["o", "s", "^"]):
    ax.scatter(x, df[col], marker=mk, s=28, label=col.replace("_V", "").replace("_", " "), zorder=3)
ax.scatter(x, df.Mean_V, marker="_", s=300, color="black", label="Mean", zorder=4)
ax.set_xticks(x); ax.set_xticklabels(labels)
ax.set(xlabel="Cell", ylabel="Voltage (V)", title="Individual repeated readings and their means")
ax.legend(ncol=4, fontsize=8, loc="upper right"); ax.set_ylim(0.3, 1.8); ax.grid(alpha=0.3)
fig.tight_layout(); fig.savefig("chart_repeated_readings.png", dpi=200); plt.close(fig)

# ---------- 3. Excel workbook (live formulas) ----------
F = "Arial"
hdr_font, hdr_fill = Font(name=F, bold=True, color="FFFFFF"), PatternFill("solid", fgColor="1F4E79")
thin = Side(style="thin", color="999999"); box = Border(left=thin, right=thin, top=thin, bottom=thin)
inp = Font(name=F, color="0000FF"); calc = Font(name=F, color="000000")

wb = Workbook()
ws = wb.active; ws.title = "Cell Data"
heads = ["Cell", "Reading 1 (V)", "Reading 2 (V)", "Reading 3 (V)", "Mean Voltage (V)",
         "Sample SD (V)", "CV (%)", "Consistency"]
for j, h in enumerate(heads, 1):
    c = ws.cell(row=1, column=j, value=h)
    c.font, c.fill, c.border = hdr_font, hdr_fill, box
    c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

for i, row in df.iterrows():
    r = i + 2
    ws.cell(r, 1, row.Cell).font = calc
    for j, col in enumerate(READ, 2):
        c = ws.cell(r, j, float(row[col])); c.font = inp; c.number_format = "0.00"
    ws.cell(r, 5, f"=AVERAGE(B{r}:D{r})").number_format = "0.00"
    ws.cell(r, 6, f"=_xlfn.STDEV.S(B{r}:D{r})").number_format = "0.000"
    ws.cell(r, 7, f'=IF(E{r}=0,"",F{r}/E{r}*100)').number_format = "0.00"
    ws.cell(r, 8, f'=IF(ROUND(F{r},3)=ROUND(MIN($F$2:$F$11),3),"Most consistent",'
                  f'IF(ROUND(F{r},3)=ROUND(MAX($F$2:$F$11),3),"Least consistent",""))')
    for j in range(1, 9):
        c = ws.cell(r, j); c.border = box
        c.alignment = Alignment(horizontal="center")
        if j >= 5: c.font = calc

notes = [
    "Legend: blue = input readings (edit these); black = formulas that recalculate automatically.",
    "Mean = AVERAGE(B:D). Sample SD = STDEV.S(B:D), which divides by n - 1 (n = 3 readings). CV (%) = SD / Mean x 100.",
    "Consistency compares SD rounded to 3 decimals, so ties (equal SDs) are flagged together.",
    "DATA NOTE: all voltages are simulated, illustrative values supplied for learning and analysis. They are not experimental results.",
]
for k, t in enumerate(notes):
    ws.cell(14 + k, 1, t).font = Font(name=F, italic=True, size=9)
for col, w in zip("ABCDEFGH", [10, 14, 14, 14, 17, 14, 10, 18]):
    ws.column_dimensions[col].width = w
ws.row_dimensions[1].height = 32; ws.freeze_panes = "A2"

# Summary sheet
sm = wb.create_sheet("Summary")
for j, h in enumerate(["Metric", "Value", "Cell (first match)"], 1):
    c = sm.cell(1, j, h); c.font, c.fill, c.border = hdr_font, hdr_fill, box
    c.alignment = Alignment(horizontal="center")
D = "'Cell Data'!"
rows = [
    ("Highest mean voltage (V)", f"=MAX({D}E2:E11)", f"=INDEX({D}A2:A11,MATCH(B2,{D}E2:E11,0))", "0.00"),
    ("Lowest mean voltage (V)", f"=MIN({D}E2:E11)", f"=INDEX({D}A2:A11,MATCH(B3,{D}E2:E11,0))", "0.00"),
    ("Mean of the 10 cell means (V)", f"=AVERAGE({D}E2:E11)", None, "0.000"),
    ("Lowest sample SD (V)", f"=MIN({D}F2:F11)", f'=INDEX({D}A2:A11,MATCH("Most consistent",{D}H2:H11,0))', "0.000"),
    ("Number of cells with lowest SD", f'=COUNTIF({D}H2:H11,"Most consistent")', None, "0"),
    ("Highest sample SD (V)", f"=MAX({D}F2:F11)", f'=INDEX({D}A2:A11,MATCH("Least consistent",{D}H2:H11,0))', "0.000"),
    ("Number of cells with highest SD", f'=COUNTIF({D}H2:H11,"Least consistent")', None, "0"),
    ("Mean of the 10 SDs (V)", f"=AVERAGE({D}F2:F11)", None, "0.000"),
    ("Highest CV (%)", f"=MAX({D}G2:G11)", f"=INDEX({D}A2:A11,MATCH(B10,{D}G2:G11,0))", "0.00"),
    ("Lowest CV (%)", f"=MIN({D}G2:G11)", f"=INDEX({D}A2:A11,MATCH(B11,{D}G2:G11,0))", "0.00"),
]
for i, (m, f1, f2, nf) in enumerate(rows, 2):
    sm.cell(i, 1, m); sm.cell(i, 2, f1).number_format = nf
    if f2: sm.cell(i, 3, f2)
    for j in range(1, 4):
        c = sm.cell(i, j); c.font = calc; c.border = box
        if j > 1: c.alignment = Alignment(horizontal="center")
sm.cell(13, 1, "Every value on this sheet is a formula that reads the 'Cell Data' sheet. Data are simulated.").font = Font(name=F, italic=True, size=9)
sm.column_dimensions["A"].width = 36; sm.column_dimensions["B"].width = 14; sm.column_dimensions["C"].width = 20

# Native Excel charts
cs = wb.create_sheet("Charts")
def column_chart(title, ycol, ytitle, anchor):
    ch = BarChart(); ch.type = "col"; ch.title = title; ch.legend = None
    ch.add_data(Reference(ws, min_col=ycol, min_row=1, max_row=11), titles_from_data=True)
    ch.set_categories(Reference(ws, min_col=1, min_row=2, max_row=11))
    ch.y_axis.title = ytitle; ch.x_axis.title = "Cell"
    ch.x_axis.delete = False; ch.y_axis.delete = False
    ch.width, ch.height = 18, 9
    cs.add_chart(ch, anchor)
column_chart("Mean Voltage of 10 Simulated Cells", 5, "Mean voltage (V)", "A1")
column_chart("Sample Standard Deviation of Repeated Readings", 6, "Sample SD (V)", "A21")

for s in (ws, sm, cs):
    s.sheet_view.showGridLines = True
for s in (ws, sm):
    for rr in s.iter_rows():
        for c in rr:
            if c.font.name != F: c.font = Font(name=F, size=c.font.size, italic=c.font.italic, bold=c.font.bold, color=c.font.color)
wb.save("electrochemical_cell_analysis.xlsx")
print("\nSaved: results CSV, 3 PNG charts, electrochemical_cell_analysis.xlsx")
