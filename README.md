# Sales Data Cleaner

A practical Python data-cleaning and reporting project that transforms raw sales records into validated analytics outputs and Excel reports.

## Overview

This project demonstrates an end-to-end data workflow:

1. Data loading and inspection
2. Cleaning duplicates and missing values
3. Validation and normalization
4. Metric engineering (revenue/discount fields)
5. Grouped sales analysis
6. Business summary generation
7. Exporting final reports

## Tech Stack

- Python
- pandas
- numpy
- openpyxl

## Project Files

- `sales_cleaner.py`: Main pipeline script
- `generate_data.py`: Utility script to generate test/messy input data
- `raw_sales.xlsx`: Input dataset (local)
- `clean_sales_report.xlsx`: Cleaned transaction output
- `high_value_orders.xlsx`: Orders with high final amount
- `city_summary.xlsx`: City-wise aggregates
- `product_summary.xlsx`: Product-wise aggregates

## Installation

```bash
pip install -r requirements.txt
```

## Usage

Run the cleaner pipeline:

```bash
python sales_cleaner.py
```

Generate synthetic raw data first (optional):

```bash
python generate_data.py
```

## Output Reports

After execution, the following files are generated in the project directory:

- `clean_sales_report.xlsx`
- `high_value_orders.xlsx`
- `city_summary.xlsx`
- `product_summary.xlsx`

## Professional Project Hygiene

- Dependency pinning via `requirements.txt`
- Repository cleanup via `.gitignore`
- Licensing via `LICENSE`
- Contribution workflow via `CONTRIBUTING.md`

## License

This project is licensed under the MIT License.