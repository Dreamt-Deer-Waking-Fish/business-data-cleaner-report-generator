# Business Data Cleaner & Report Generator

![Python](https://img.shields.io/badge/Python-3.9%2B-blue)
![Output](https://img.shields.io/badge/Output-Excel%20%2B%20PNG-green)
![Use Case](https://img.shields.io/badge/Use%20Case-Sales%20Reporting-orange)

Client-style automation for small businesses that receive messy sales exports from POS, ecommerce, inventory, or accounting systems.

The workflow imports multiple CSV or Excel files, standardizes inconsistent columns, cleans common data issues, removes duplicate orders, calculates key sales metrics, and exports a polished Excel report plus chart images.

## Client Problem

Many small businesses lose time every month manually cleaning sales exports before they can answer basic questions:

- How much revenue did we generate?
- Which products performed best?
- Which months or regions are trending up or down?
- Which rows are duplicated, incomplete, or inconsistent?

This project turns that manual reporting process into a repeatable command-line workflow.

## Delivered Solution

- Multi-file CSV/XLSX import from `sample_data/`
- Column normalization for inconsistent export headers
- Missing value cleanup for customer, product, quantity, price, and revenue fields
- Duplicate order removal by `order_id`
- Revenue, order count, average order value, top product, and monthly revenue calculations
- Excel workbook with multiple stakeholder-friendly sheets
- PNG chart exports for quick sharing in reports or proposals
- Run log for client handoff and troubleshooting

## Project Structure

```text
business-data-cleaner-report-generator/
  main.py
  requirements.txt
  README.md
  portfolio_description.md
  docs/
  sample_data/
    january_orders.csv
    february_orders.csv
    march_orders.csv
  output/
    summary_report.xlsx
    top_products_revenue.png
    revenue_by_month.png
  screenshots/
```

## Setup

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
```

If you are running this from the repository root, install the shared dependency file instead:

```powershell
pip install -r requirements.txt
```

## Usage

Run the sample workflow:

```powershell
python main.py
```

Run against custom folders:

```powershell
python main.py --input-dir "sample_data" --output-dir "output"
```

## Outputs

```text
output/summary_report.xlsx
output/top_products_revenue.png
output/revenue_by_month.png
output/run.log
```

The Excel workbook includes:

| Sheet | Purpose |
| --- | --- |
| `Summary` | Total revenue, total orders, and average order value |
| `Top Products` | Ranked product revenue and unit volume |
| `Revenue By Month` | Monthly revenue and order count |
| `Cleaned Orders` | Standardized order-level dataset |

## Validation

```powershell
python -m py_compile main.py
python main.py
```

Successful run criteria:

- Excel report is generated
- Both chart PNG files are generated
- Invalid date rows are logged and skipped
- Duplicate orders are removed
- Final log reports the number of processed clean orders

## Example Client Applications

- Monthly sales reporting for retail shops, salons, cafes, and ecommerce sellers
- POS export cleanup before management review
- Combining sales files from multiple locations
- Preparing clean order data for Power BI, Looker Studio, or accounting imports
- Replacing repetitive spreadsheet cleanup tasks with a repeatable workflow

## Security and Data Handling

This demo uses sample data only. For real client work, raw exports should be handled in a private repository or secure transfer location. No credentials are required for this project.

## Known Limits

- The included mapping covers common sales export fields, not every possible POS schema.
- The report is batch-oriented and does not run on a schedule by default.
- Refund, tax, shipping, and discount logic can be added for client-specific workflows.

## Upgrade Ideas

- Client-specific column mapping config
- Scheduled monthly report generation
- Email or Slack delivery
- Dashboard export
- Data validation report for rejected rows
