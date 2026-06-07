# Langmuir-Probe-Automated-Batch-Data-Processing

Batch-processing tools for Langmuir probe current-voltage sweep analysis. The code estimates key plasma parameters including floating potential, plasma potential, electron temperature, ion saturation current, and plasma density.
This repository was created from an existing research/prototype script and reorganized so it can be imported as a Python package or run from the command line.
## What this does

- Reads a folder of Langmuir probe Excel files
- Identifies bad data files and optionally quarantines them
- Calculates zero-crossing and floating potential
- Estimates plasma potential using a second-derivative method with fallback logic
- Calculates electron current, electron temperature, density, and ion saturation current
- Writes a summary spreadsheet to `results/lp_data_post_processing_summary.xlsx`

## Repository Structure

```
├── src/langmuir_probe_batch/
│   ├── __init__.py
│   ├── cli.py          # command-line workflow
│   ├── constants.py    # physical constants and default probe geometry
│   └── core.py         # original analysis functions, no hard-coded execution
├── docs/
│   └── original_script.py
├── tests/
├── pyproject.toml
├── requirements.txt
├── .gitignore
└── README.md
```
## Install

```bash
python -m venv .venv
.venv\Scripts\activate  # Windows
# source .venv/bin/activate  # macOS/Linux

pip install -e .
```
## Run
```bash
lp-batch "path/to/langmuir_probe_excel_files" --output-dir results
```
Common options:
```bash
lp-batch "data/52925/500th Recycle" \
  --pattern "*.xlsx" \
  --shape cylindrical \
  --probe-diameter-m 0.00127 \
  --output-dir results
```
By default, suspicious files are copied to `results/bad_data` but not deleted. To also delete them from the input folder after copying:
```bash
lp-batch "data" --remove-bad-data
```
## Input data expectations
Each input workbook should contain voltage and current columns. The original workflow assumes:
- Column 0: probe voltage / bias voltage
- Column 1: probe current
- A column named `Current` is used for bad-data screening

## Important notes
This is still research/prototype code. The physics methods and fitting windows should be validated against known data before using results for publication, design decisions, or customer deliverables.
The original script had local Windows paths and ran immediately when opened. Those paths were removed from the CLI workflow. A copy of the original script is preserved in `docs/original_script.py` for traceability.

## Development
Run an import smoke test:
```bash
python -c "import langmuir_probe_batch; print(langmuir_probe_batch.__version__)"
```
Future cleanup ideas:
- Add sample test data
- Add unit tests for each calculation function
- Replace broad `except` blocks with specific exceptions
- Move plotting output paths into CLI configuration
- Convert legacy CamelCase function names to snake_case wrappers
- Add configuration files for probe geometry and gas species
