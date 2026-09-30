# Weekly Return Generator

Python tool for automating the generation of weekly personnel and plant returns from project planning Excel files.

## Goal

The goal of this project is to replace a manual Excel/Power Query workflow with a Python-based process that can:

- Read project planning Excel files
- Detect relevant planning sections and their structure
- Process personnel and plant/equipment data
- Apply return business rules
- Generate a formatted weekly return from an Excel template
- Gradually become robust to different planning layouts

## Current Features

The current version can:

- Read planning data from Excel
- Automatically detect the date range in the planning
- Detect relevant planning sections without relying on fixed row numbers
- Extract and clean personnel and equipment sections
- Apply different business rules to personnel and equipment
- Generate weekly personnel and equipment data
- Validate the requested reporting week
- Check that the complete seven-day period is available in the planning
- Generate the final return using an Excel template
- Dynamically handle different numbers of personnel and equipment entries

## Project Structure

```
weekly-return-generator/
│
├── src/
│   ├── main.py
│   ├── config.py
│   ├── planning.py
│   ├── processing.py
│   ├── rules.py
│   └── output.py
│
├── templates/
│   └── template_return.xlsx
│
├── data/
│   └── Staff Planning 2027
│
├── output/
│   └── ...
│
├── requirements.txt
├── .gitignore
└── README.md
```

### Modules

- `main.py` — application flow and user input
- `config.py` — file paths, section names and business rules
- `planning.py` — detection of dates and planning sections
- `processing.py` — data extraction, cleaning and transformation
- `rules.py` — personnel and equipment business rules
- `output.py` — generation and formatting of the Excel return

## Requirements

- Python 3.12+
- pandas
- openpyxl

Install the required packages with:

```bash
pip install -r requirements.txt
```

## Usage

The program asks the user to provide the starting date of the reporting week.

The date must be a Monday. Both `YYYY-MM-DD` and `DD-MM-YYYY` formats are accepted.

The program also validates that the complete seven-day reporting period is available in the planning file before processing begins.

## Status

🚧 **Under active development**

The core workflow is functional. Future development will focus on improving robustness across different project planning layouts, expanding configuration options, and further separating project-specific rules from the processing logic.

## Disclaimer

This project is developed as a personal Python learning and automation project.

The repository may contain example planning structures, business rules and templates used to demonstrate the workflow. Real project planning files and confidential company data are not included.