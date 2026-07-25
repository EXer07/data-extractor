# data-extractor

[![ci](https://github.com/EXer07/data-extractor/actions/workflows/ci.yml/badge.svg)](https://github.com/EXer07/data-extractor/actions/workflows/ci.yml)
[![python](https://img.shields.io/badge/Python-3.11-3776AB.svg?style=flat&logo=python&logoColor=white)](https://www.python.org)
[![Poetry](https://img.shields.io/endpoint?url=https://python-poetry.org/badge/v0.json)](https://python-poetry.org/)
[![Code style: black](https://img.shields.io/badge/code%20style-black-000000.svg)](https://github.com/psf/black)
[![Checked with pyright](https://microsoft.github.io/pyright/img/pyright_badge.svg)](https://microsoft.github.io/pyright/)

Python tools for extracting public Power BI semantic tables, validating API
datasets, and producing structured Excel, Parquet, DuckDB, and JSONL outputs.
The original Selenium console and GUI scraper remains available.

<img src="./docs/gui_screenshot.png" alt="GUI Screenshot" width="400"/>

## Prerequisites

-   [Python 3.11](https://www.python.org/downloads/release/python-311/)
-   [Poetry](https://python-poetry.org/docs/) (optional)

## Installation

1. Clone the repository:

```bash
git clone https://github.com/EXer07/data-extractor.git
cd data-extractor
```

2. Install dependencies using Poetry:

```bash
poetry install
```

For non-poetry users, a `requirements.txt` file is also provided.

## Configuration

To set up configuration:

1. Rename `config.example.yml` to `config.yml`.

2. Update the `config.yml` file with your specific settings.

-   To switch between GUI and Console mode, change the `mode` value to either `gui` or `console`.
-   Depending on the mode, the `gui` or `console` section of the config file will be used. The other section will be ignored, but you can keep it in the file if you still want to have the possibility to switch between modes.

```yml
# EXAMPLE CONFIG FILE

mode: gui # REQUIRED: Options: gui or console

should_uncheck_filter: true # OPTIONAL (default=false): Find checkbox filter and uncheck all checkboxes before scraping
max_rows: null # OPTIONAL (default=None): Set a maximum number of rows to scrape (e.g. for reducing scraping time during testing)

console:
    url: https://app.powerbi.com/XXXXX # REQUIRED: URL to the Power BI report that should be scraped
    is_headless: true # OPTIONAL (default=true): 'true' hides the the browser window during scraping
    output_format: excel # OPTIONAL (default=excel): Options: excel, csv
    output_path: ./table.xlsx # OPTIONAL (default="./table.xlsx"): File extension should match the output_format (i.e. .xlsx for excel and .csv for csv)

gui:
    language: en # OPTIONAL (defaul=en): Options: en, da
    program_name: Power BI Table Scraper # OPTIONAL (defaul=Power BI Table Scraper) The program name that should be displayed in the GUI

    # Default values in the GUI. Can be changed by the user.
    default_values:
        url: https://app.powerbi.com/XXXXX # OPTIONAL (default=None): URL to the Power BI report that should be scraped
        is_headless: true # OPTIONAL (default=true): 'true' hides the the browser window during scraping
        output_format: excel # OPTIONAL (default=excel): excel or csv
        output_path: null # OPTIONAL(default=None): User is always required to browse for a valid path before being able to run the scraper unless a default path is specified here. File extension should match the output_format (i.e. .xlsx for excel and .csv for csv)
```

i.e. the minimum required configuration for the console mode is:

```yml
mode: console
console:
    url: https://app.powerbi.com/XXXXX
```

and for the GUI mode:

```yml
mode: gui
```

## Usage

After setting up `./config.yml`:

```bash
python main.py
```

For the GUI mode, follow the on-screen instructions. For the Console mode, scraping will start automatically based on the settings defined in `config.yml`.

## Dataset workflows

- `extract_all.py`: extract all exposed raw tables from a public Power BI report.
- `create_three_excel_files.py`: split the first rice report into reader-facing workbooks.
- `rename_powerbi_2_excel.py`: rename and format the satellite rice-supply workbook.
- `build_rag_dataset.py`: create Parquet, DuckDB, JSONL, metadata, and quality reports for hybrid RAG.
- `export_clean_excel.py`: export a clean Excel copy of the first rice dataset.

See [`SKILLS.md`](SKILLS.md) for source URLs, validation requirements, known
Power BI limits, API notes, and cross-computer handoff instructions.

Generated datasets, local configuration, API keys, and virtual environments are
excluded by `.gitignore`.

## Creating a Standalone Executable with PyInstaller

To create a standalone executable of the tool, run the following command:

```bash
pyinstaller --onefile --noconsole --name="Power BI Table Scraper" ./src/main.py
```

The executable will be created in the `./dist` folder - remember to include the `config.yml` file in the same folder as the executable. Now the tool can be run without having to install Python or any dependencies.
