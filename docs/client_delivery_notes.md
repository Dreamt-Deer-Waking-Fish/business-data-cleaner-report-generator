# Client Delivery Notes

## What the Client Receives

- A reusable Python script for cleaning sales exports
- Sample input files that demonstrate the expected format
- Generated Excel and PNG outputs
- Documentation for setup, usage, and troubleshooting
- A screenshot checklist for portfolio or handoff material

## Acceptance Checklist

- `python -m py_compile main.py` passes
- `python main.py` runs without crashing
- `output/summary_report.xlsx` is created
- `output/top_products_revenue.png` is created
- `output/revenue_by_month.png` is created
- Runtime logs are generated locally and excluded from Git

## Client Customization Options

- Add custom column aliases
- Add tax, discount, refund, or shipping calculations
- Add scheduled execution
- Add email delivery
- Add dashboard export
