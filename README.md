# Electrochemical-Cell-Analysis
A reproducible educational data-analysis project that calculates mean voltage, sample standard deviation (SD), and coefficient of variation (CV) for repeated readings from 10 illustrative electrochemical cells. It exports a processed CSV, publication-style PNG figures, an Excel workbook with live formulas and charts, and a PDF report.
Data integrity statement: every voltage value bundled with this repository is simulated for learning and analysis. The dataset is not experimental and must not be presented as laboratory evidence or measured electrochemical performance.

Project overview
The input dataset contains three voltage readings for each of 10 named cells. The Python workflow validates the input, calculates descriptive statistics, creates visualizations, and builds a formatted workbook. A second script produces a PDF report from the generated results and figures.
Problem statement
Repeated readings can vary. This project demonstrates how to summarize the central voltage level and the spread of repeated readings, then present those statistics in reusable analysis outputs. It is a computational exercise, not a model of electrode or electrolyte chemistry.
Objectives
- Load and validate a structured CSV dataset.
- Calculate mean voltage, sample SD (using n - 1), and CV percentage.
- Identify cells with the lowest and highest rounded sample SD.
- Generate clear comparison charts.
- Export results to CSV, Excel, and PDF formats.
- Demonstrate a reproducible, documented Python workflow.
Key features
- Input checks for required columns, missing/non-numeric/non-finite readings, duplicate cell names, and non-positive voltages.
- pandas-based descriptive statistics (ddof=1 for sample SD).
- Three Matplotlib figures: mean voltage with SD error bars, SD comparison, and repeated readings with means.
- Excel workbook with editable input readings, AVERAGE, STDEV.S, CV formulas, summary metrics, and native charts.
- PDF report with dataset, method, results, interpretation, figures, and limitations.
- Unit tests for the bundled data and common invalid-input cases.
Technologies and tools
Technology	Purpose
Python	Analysis workflow and report generation
pandas	CSV loading and tabular calculations
NumPy	Numeric validation and chart coordinates
Matplotlib	PNG charts and project banner
openpyxl	Excel workbook, formulas, formatting, and charts
ReportLab	PDF report generation
unittest	Automated validation tests


Project architecture and workflow
See [the workflow diagram](docs/PROJECT_WORKFLOW.md).
1. Read data/raw/electrochemical_cell_data.csv.
2. Validate cell identifiers and voltage readings.
3. Calculate mean, sample SD, and CV.
4. Write the processed CSV to data/processed/.
5. Generate figures in figures/ and the Excel workbook in reports/.
6. Generate the PDF report in reports/.
Folder structure
Electrochemical_Cell_Analysis_Project_Professional/
├── data/
│   ├── raw/                 # Original input dataset
│   └── processed/           # Calculated results CSV
├── docs/                    # Workflow documentation
├── figures/                 # Charts and project banner
├── reports/                 # Excel workbook and PDF report
├── src/
│   ├── electrochemical_cell_analysis.py
│   └── make_report_pdf.py
├── tests/
│   └── test_analysis.py
├── .gitignore
├── LICENSE_NOTICE.md
├── README.md
└── requirements.txt
Installation requirements
- Python 3.10 or newer is recommended.
- pip and a terminal/command prompt.
- Windows, macOS, or Linux.
Setup
From the project root, create and activate a virtual environment.

Windows (PowerShell)
py -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install --upgrade pip
pip install -r requirements.txt
macOS / Linux
python3 -m venv .venv
source .venv/bin/activate
python -m pip install --upgrade pip
pip install -r requirements.txt
How to run
Run commands from the project root.
python src/electrochemical_cell_analysis.py
python src/make_report_pdf.py
The analysis script also accepts custom paths:
python src/electrochemical_cell_analysis.py --input path/to/input.csv --output path/to/output_project
The PDF generator expects the processed CSV and three figures to exist, so run the analysis first. Paths are resolved relative to the script/project, not the shell's current directory.
Usage instructions
1. Make a backup of the sample CSV.
2. Edit data/raw/electrochemical_cell_data.csv while retaining the required headers: Cell, Reading_1_V, Reading_2_V, and Reading_3_V.
3. Use unique cell names and positive numeric readings.
4. Run the analysis script, then the PDF script.
5. Open reports/electrochemical_cell_analysis.xlsx. Blue cells on the Cell Data sheet are input readings; calculated fields use Excel formulas.
The script validates the bundled schema and reports invalid input rather than silently continuing.
Results and findings
These values are recalculated from the bundled simulated dataset:
Finding	Result
Number of cells	10
Readings per cell	3
Highest mean voltage	Cell 3 — 1.50 V
Lowest mean voltage	Cell 9 — 0.50 V
Mean of the 10 cell means	0.995 V
Lowest sample SD	0.02 V — Cells 1, 2, 3, and 5
Highest sample SD	0.05 V — Cells 7, 8, 9, and 10
Highest CV	Cell 9 — 10.00%
Lowest CV	Cell 3 — 1.33%


The results describe only the supplied simulated numbers. They do not establish the performance or quality of real electrochemical cells.
Output previews
Mean voltage

Sample standard deviation

Repeated readings

Statistical methods
For three readings x1, x2, and x3:
- Mean: (x1 + x2 + x3) / 3
- Sample standard deviation: sqrt(sum((xi - mean)^2) / (n - 1)), with n = 3.
- Coefficient of variation: (sample SD / mean) * 100.
Python uses DataFrame.std(axis=1, ddof=1). Excel uses STDEV.S.
Limitations
- All supplied readings are simulated, not measured.
- Only three readings are supplied per cell, so SD estimates are based on a small sample.
- Electrode materials, electrolyte composition, temperature, concentration, and other electrochemical conditions are not modelled.
- No inferential tests (such as t-tests or ANOVA) are performed.
- Excel formulas recalculate when opened in a spreadsheet application that supports these functions; openpyxl writes formulas but does not calculate their displayed cached values itself.
Future improvements
- Add real, documented laboratory datasets with measurement conditions and provenance.
- Support a configurable number of repeated readings rather than exactly three named input columns.
- Add confidence intervals and appropriate inferential tests when justified by the experimental design.
- Add continuous integration and code-coverage reporting.
- Add workbook formula-result verification in Microsoft Excel or LibreOffice in an automated environment.
Conclusion
This project demonstrates an end-to-end workflow for validating repeated voltage readings, calculating descriptive statistics, and presenting results in CSV, Excel, chart, and PDF formats. Its value is as a reproducible programming and data-presentation exercise; experimental claims require genuine measurements and documented methods.
Author
Hasan Ali — author attribution carried over from the original report metadata. Update this section if a different display name or affiliation is required.
License
No license was present in the supplied project. See [LICENSE_NOTICE.md](LICENSE_NOTICE.md) before publishing. Add a recognized license only after deciding the terms under which others may use the work.
GitHub publication checklist
- Suggested repository name: electrochemical-cell-analysis
- Suggested description: Educational Python and Excel workflow for analyzing repeated electrochemical-cell voltage readings using mean, sample standard deviation, and coefficient of variation.
- Suggested topics: python, data-analysis, electrochemistry, pandas, numpy, matplotlib, excel, openpyxl, reportlab, statistics
- Review the dataset and remove any private or sensitive information before publishing.
- Choose and add a license if you want to permit reuse.
