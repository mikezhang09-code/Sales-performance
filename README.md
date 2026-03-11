# Sales-performance

This repository contains tools for analyzing telecom sales performance data and generating scoring and ranking models.

## Scripts & Tools

- `generate_excel.py`: Reads the raw sales data (`Mike-prepared.xlsx`), computes multiple KPIs, dynamically calculates the formulas for performance scoring across different regions, and exports the results as a formatted `.xlsx` file along with a `.zip` archive.
- `analysis/compute_ranking.py`: Takes the raw excel data and computes an overall score and rank based on weighted KPIs, directly parsing the underlying XML for performance.

## Setup & Usage

Ensure you have a Python environment set up with the required dependencies (e.g., `openpyxl`, `pandas`).

```bash
python3 -m venv venv
source venv/bin/activate
pip install openpyxl pandas
python generate_excel.py
```