from __future__ import annotations

import argparse
import logging
from pathlib import Path
from typing import Dict, Iterable, List

import matplotlib.pyplot as plt
import pandas as pd


COLUMN_ALIASES: Dict[str, str] = {
    "order id": "order_id",
    "order_id": "order_id",
    "id": "order_id",
    "date": "order_date",
    "order date": "order_date",
    "order_date": "order_date",
    "product": "product",
    "product name": "product",
    "item": "product",
    "customer": "customer",
    "customer name": "customer",
    "qty": "quantity",
    "quantity": "quantity",
    "units": "quantity",
    "unit price": "unit_price",
    "unit_price": "unit_price",
    "price": "unit_price",
    "sales": "revenue",
    "revenue": "revenue",
    "total": "revenue",
    "region": "region",
    "sales region": "region",
}


def configure_logging(output_dir: Path) -> None:
    output_dir.mkdir(parents=True, exist_ok=True)
    logging.getLogger("matplotlib").setLevel(logging.WARNING)
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s | %(levelname)s | %(message)s",
        handlers=[
            logging.StreamHandler(),
            logging.FileHandler(output_dir / "run.log", mode="w", encoding="utf-8"),
        ],
    )


def normalize_column_name(column: str) -> str:
    normalized = str(column).strip().lower().replace("-", " ").replace("_", " ")
    normalized = " ".join(normalized.split())
    return COLUMN_ALIASES.get(normalized, normalized.replace(" ", "_"))


def coalesce_duplicate_columns(frame: pd.DataFrame) -> pd.DataFrame:
    cleaned = pd.DataFrame(index=frame.index)
    for column in dict.fromkeys(frame.columns):
        matches = frame.loc[:, frame.columns == column]
        if matches.shape[1] == 1:
            cleaned[column] = matches.iloc[:, 0]
        else:
            cleaned[column] = matches.bfill(axis=1).iloc[:, 0]
            logging.info("Coalesced %s duplicate '%s' columns", matches.shape[1], column)
    return cleaned


def load_input_files(input_dir: Path) -> pd.DataFrame:
    files = sorted(
        list(input_dir.glob("*.csv"))
        + list(input_dir.glob("*.xlsx"))
        + list(input_dir.glob("*.xls"))
    )
    if not files:
        raise FileNotFoundError(f"No CSV or Excel files found in {input_dir}")

    frames: List[pd.DataFrame] = []
    for file_path in files:
        try:
            if file_path.suffix.lower() == ".csv":
                frame = pd.read_csv(file_path)
            else:
                frame = pd.read_excel(file_path)
            frame["source_file"] = file_path.name
            frames.append(frame)
            logging.info("Loaded %s rows from %s", len(frame), file_path.name)
        except Exception as exc:
            logging.error("Failed to load %s: %s", file_path.name, exc)
            raise

    return pd.concat(frames, ignore_index=True)


def clean_orders(raw: pd.DataFrame) -> pd.DataFrame:
    orders = raw.copy()
    orders.columns = [normalize_column_name(column) for column in orders.columns]
    orders = coalesce_duplicate_columns(orders)

    required_columns = ["order_id", "order_date", "product", "customer", "quantity", "unit_price"]
    for column in required_columns:
        if column not in orders.columns:
            orders[column] = pd.NA

    orders["order_id"] = orders["order_id"].astype("string").str.strip()
    missing_order_ids = orders["order_id"].isna() | (orders["order_id"] == "")
    orders.loc[missing_order_ids, "order_id"] = "MISSING-" + orders.index[missing_order_ids].astype(str)
    orders["order_date"] = pd.to_datetime(orders["order_date"], errors="coerce")
    orders["product"] = orders["product"].fillna("Unknown Product").astype(str).str.strip()
    orders["customer"] = orders["customer"].fillna("Unknown Customer").astype(str).str.strip()
    orders["region"] = orders.get("region", "Unassigned")
    orders["region"] = orders["region"].fillna("Unassigned").astype(str).str.strip()
    orders["quantity"] = pd.to_numeric(orders["quantity"], errors="coerce").fillna(1)
    orders["unit_price"] = pd.to_numeric(orders["unit_price"], errors="coerce").fillna(0)

    if "revenue" in orders.columns:
        orders["revenue"] = pd.to_numeric(orders["revenue"], errors="coerce")
    else:
        orders["revenue"] = pd.NA
    orders["revenue"] = orders["revenue"].fillna(orders["quantity"] * orders["unit_price"])

    missing_dates = orders["order_date"].isna().sum()
    if missing_dates:
        logging.warning("Dropped %s rows with invalid order dates", missing_dates)
        orders = orders.dropna(subset=["order_date"])

    before = len(orders)
    orders = orders.drop_duplicates(subset=["order_id"], keep="first")
    logging.info("Removed %s duplicate orders", before - len(orders))

    orders["month"] = orders["order_date"].dt.to_period("M").astype(str)
    orders = orders.sort_values("order_date").reset_index(drop=True)
    return orders[
        [
            "order_id",
            "order_date",
            "month",
            "customer",
            "product",
            "region",
            "quantity",
            "unit_price",
            "revenue",
            "source_file",
        ]
    ]


def calculate_metrics(orders: pd.DataFrame) -> Dict[str, pd.DataFrame]:
    total_revenue = float(orders["revenue"].sum())
    total_orders = int(orders["order_id"].nunique())
    average_order_value = total_revenue / total_orders if total_orders else 0

    summary = pd.DataFrame(
        [
            {"metric": "Total Revenue", "value": round(total_revenue, 2)},
            {"metric": "Total Orders", "value": total_orders},
            {"metric": "Average Order Value", "value": round(average_order_value, 2)},
        ]
    )
    top_products = (
        orders.groupby("product", as_index=False)
        .agg(total_revenue=("revenue", "sum"), total_quantity=("quantity", "sum"), orders=("order_id", "nunique"))
        .sort_values("total_revenue", ascending=False)
    )
    revenue_by_month = (
        orders.groupby("month", as_index=False)
        .agg(total_revenue=("revenue", "sum"), orders=("order_id", "nunique"))
        .sort_values("month")
    )
    return {
        "summary": summary,
        "top_products": top_products,
        "revenue_by_month": revenue_by_month,
    }


def export_excel_report(orders: pd.DataFrame, metrics: Dict[str, pd.DataFrame], output_dir: Path) -> Path:
    report_path = output_dir / "summary_report.xlsx"
    with pd.ExcelWriter(report_path, engine="openpyxl") as writer:
        metrics["summary"].to_excel(writer, sheet_name="Summary", index=False)
        metrics["top_products"].to_excel(writer, sheet_name="Top Products", index=False)
        metrics["revenue_by_month"].to_excel(writer, sheet_name="Revenue By Month", index=False)
        orders.to_excel(writer, sheet_name="Cleaned Orders", index=False)
    logging.info("Excel report saved to %s", report_path)
    return report_path


def save_bar_chart(data: pd.DataFrame, x: str, y: str, title: str, output_path: Path) -> None:
    plt.figure(figsize=(10, 6))
    plt.bar(data[x].astype(str).tolist(), data[y].tolist())
    plt.title(title)
    plt.xlabel(x.replace("_", " ").title())
    plt.ylabel(y.replace("_", " ").title())
    plt.xticks(rotation=35, ha="right")
    plt.tight_layout()
    plt.savefig(output_path, dpi=150)
    plt.close()
    logging.info("Chart saved to %s", output_path)


def generate_charts(metrics: Dict[str, pd.DataFrame], output_dir: Path) -> None:
    top_products = metrics["top_products"].head(8)
    revenue_by_month = metrics["revenue_by_month"]
    save_bar_chart(top_products, "product", "total_revenue", "Top Products by Revenue", output_dir / "top_products_revenue.png")
    save_bar_chart(revenue_by_month, "month", "total_revenue", "Revenue by Month", output_dir / "revenue_by_month.png")


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Clean messy sales exports and generate an Excel business report.")
    parser.add_argument("--input-dir", default="sample_data", help="Folder containing CSV/XLSX files.")
    parser.add_argument("--output-dir", default="output", help="Folder where reports and charts are saved.")
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    input_dir = Path(args.input_dir)
    output_dir = Path(args.output_dir)
    configure_logging(output_dir)

    try:
        raw_orders = load_input_files(input_dir)
        clean_data = clean_orders(raw_orders)
        metrics = calculate_metrics(clean_data)
        export_excel_report(clean_data, metrics, output_dir)
        generate_charts(metrics, output_dir)
        logging.info("Done. Processed %s clean orders.", len(clean_data))
    except Exception as exc:
        logging.exception("Report generation failed: %s", exc)
        raise SystemExit(1)


if __name__ == "__main__":
    main()
