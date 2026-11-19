# Electrochemical Cell Analysis Using Python and Excel

Calculates the mean voltage, sample standard deviation (n - 1) and coefficient of variation for three repeated
readings on each of 10 illustrative electrochemical cells, then exports an Excel workbook (live formulas +
native charts), PNG charts and a PDF report.

**DATA NOTE: all voltages are simulated for learning and analysis. They are not experimental results.**

## Files
| File | Purpose |
|---|---|
| electrochemical_cell_data.csv | Input: 3 readings x 10 cells (edit to analyse new data) |
| electrochemical_cell_analysis.py | Main script: statistics, CSV, charts, Excel workbook |
| make_report_pdf.py | Builds the PDF report from the results and charts |
| electrochemical_cell_results.csv | Output: readings, mean, SD, CV |
| electrochemical_cell_analysis.xlsx | Sheets: Cell Data (AVERAGE, STDEV.S), Summary, Charts |
| chart_mean_voltage.png / chart_standard_deviation.png / chart_repeated_readings.png | Figures |
| Electrochemical_Cell_Analysis_Report.pdf | 3-page report |

## Run
    pip install -r requirements.txt
    python electrochemical_cell_analysis.py
    python make_report_pdf.py

## Excel formulas (row 2, fill down to row 11)
    E2 (mean)       =AVERAGE(B2:D2)
    F2 (sample SD)  =STDEV.S(B2:D2)
    G2 (CV %)       =F2/E2*100

## Key results
Highest mean: Cell 3 (1.50 V). Lowest mean: Cell 9 (0.50 V). Lowest SD (0.02 V): cells 1, 2, 3, 5.
Highest SD (0.05 V): cells 7, 8, 9, 10. Mean of the 10 cell means: 0.995 V.

## Limits
Simulated data; n = 3 per cell gives rough SD estimates; no electrochemistry (electrodes, electrolyte,
temperature) is modelled; no significance tests are run.
