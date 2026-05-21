# Architecture Diagram

![Architecture diagram](../assets/architecture_diagram.png)

## Components

| Layer | Responsibility |
| --- | --- |
| Client exports | Raw CSV/XLSX files from POS, ecommerce, inventory, or accounting tools |
| Cleaning engine | Loads files, standardizes columns, fills common missing values, and removes duplicate orders |
| Metrics layer | Calculates revenue, order count, average order value, top products, and monthly revenue |
| Output layer | Produces an Excel workbook, PNG charts, and a local run log |

## Data Flow

```mermaid
flowchart LR
    A[CSV/XLSX sales exports] --> B[Load files with pandas]
    B --> C[Normalize columns and clean rows]
    C --> D[Calculate business metrics]
    D --> E[Excel workbook]
    D --> F[PNG charts]
    C --> G[Cleaned Orders sheet]
```

## Client Notes

The architecture is intentionally batch-oriented so a client can drop new exports into `sample_data/` and rerun the same command for a fresh report.
