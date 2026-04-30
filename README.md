# Excel Sheet Comparison Tool

A command-line Python tool that compares two Excel files by statistically profiling every column and producing a structured comparison report.

## What It Does

Takes two `.xlsx` files and outputs a `comparison_output.xlsx` containing:

**Summary sheet** — high-level side-by-side comparison:
- Row and column counts
- Total null counts and percentages
- Columns unique to each file
- Numeric vs text column breakdown

**Column Detail sheet** — per-column statistics for both files:
- Data type, row count, null/blank counts and percentages
- Unique value count
- Min, max, mean, median, standard deviation (numeric columns)
- Q1, Q3, and outlier count using the IQR method (numeric columns)
- Most common value and its frequency (all columns)

## Requirements

- Python 3.10+
- pandas
- openpyxl

## Setup

```bash
pip install pandas openpyxl
```

## Usage

```bash
python BrendonSheetComparison.py file_a.xlsx file_b.xlsx
```

Output is written to `comparison_output.xlsx` in the current directory.


## Roadmap

### Error Handling & Validation
- [x] Graceful handling of missing or invalid file paths
- [x] Detect and report unreadable or corrupt Excel files
- [ ] Handle password-protected workbooks
- [ ] Validate that input files contain data (not empty sheets)
- [ ] Try/except blocks around column profiling to handle unexpected data types
- [ ] Meaningful error messages with guidance on how to fix the issue

### Multi-File Comparison
- [ ] Accept any number of input files (not just two)
- [ ] Dynamic summary sheet that scales columns to match the number of files
- [ ] Column Detail sheet with one row per file per column
- [ ] Support comparing specific sheets within multi-sheet workbooks

### Output Polish
- [ ] Written highlight summary on the Summary sheet (auto-generated text calling out key differences)
- [ ] Value formatting — percentages to 2dp, large numbers with commas, clean null display
- [ ] Column width auto-sizing based on content length
- [ ] Row height and text wrapping for readability
- [ ] Conditional formatting — highlight cells where stats diverge significantly between files
- [ ] Freeze header rows and panes for easier navigation

### Deeper Analysis
- [ ] Side-by-side column comparison (File A and File B stats on the same row per column)
- [ ] Data type mismatch detection (flag columns where the type differs between files)
- [ ] Row-level diff using a key column to identify added, removed, and changed rows
- [ ] Value distribution comparison — flag when a column's spread shifts significantly

### Usability
- [ ] Support for CSV input files alongside Excel
- [ ] Optional config file to set thresholds (e.g. what % null is "too many", outlier sensitivity)
- [ ] Command line flags for output path and sheet selection (using argparse)
- [ ] Logging instead of print statements

### Visualisation
- [ ] Embedded bar/histogram charts in the output for numeric distributions
- [ ] Heatmap-style summary grid showing where the biggest differences are
