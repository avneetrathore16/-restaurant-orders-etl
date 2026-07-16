"""ETL pipeline for restaurant orders.

Run this module directly to create a curated CSV, SQLite database, and
validation report from the supplied Orders.csv and Restaurants.csv files.
"""

from __future__ import annotations

import argparse
import json
import re
import sqlite3
from dataclasses import asdict, dataclass
from pathlib import Path

import pandas as pd


PROJECT_DIR = Path(__file__).resolve().parent
DEFAULT_ORDERS = Path.home() / "Downloads" / "Orders.csv"
DEFAULT_RESTAURANTS = Path.home() / "Downloads" / "Restaurants.csv"


@dataclass
class PipelineReport:
    source_orders: int
    source_restaurants: int
    curated_orders: int
    rejected_orders: int
    unmatched_restaurants: int
    duplicate_order_ids: int
    missing_values_after_cleaning: int


def snake_case(value: str) -> str:
    """Convert source headers to predictable, database-friendly names."""
    value = re.sub(r"[^a-zA-Z0-9]+", "_", value.strip())
    return re.sub(r"_+", "_", value).strip("_").lower()


def resolve_source(explicit: str | None, filename: str) -> Path:
    """Prefer a project-local raw file, then use the supplied Downloads file."""
    candidates = [
        Path(explicit).expanduser() if explicit else None,
        PROJECT_DIR / "data" / "raw" / filename,
        Path.home() / "Downloads" / filename,
    ]
    for path in candidates:
        if path and path.exists():
            return path
    raise FileNotFoundError(
        f"Could not find {filename}. Put it in data/raw/ or pass its path explicitly."
    )


def clean_orders(raw: pd.DataFrame) -> tuple[pd.DataFrame, int, int]:
    orders = raw.copy()
    orders.columns = [snake_case(column) for column in orders.columns]
    text_columns = orders.select_dtypes(include="object").columns
    for column in text_columns:
        orders[column] = orders[column].astype("string").str.strip()

    orders["order_date"] = pd.to_datetime(orders["order_date"], errors="coerce")
    for column in [
        "restaurant_id", "quantity_of_items", "order_amount",
        "delivery_time_taken_mins", "customer_rating_food",
        "customer_rating_delivery",
    ]:
        orders[column] = pd.to_numeric(orders[column], errors="coerce")

    orders["payment_mode"] = (
        orders["payment_mode"].str.lower().str.replace(r"\s+", " ", regex=True)
        .replace({
            "upi": "UPI", "cash on delivery": "Cash on Delivery",
            "credit card": "Credit Card", "debit card": "Debit Card",
            "cash": "Cash",
        })
    )
    orders["payment_mode"] = orders["payment_mode"].str.title().replace({"Upi": "UPI"})

    duplicate_ids = int(orders["order_id"].duplicated(keep="first").sum())
    orders = orders.drop_duplicates(subset="order_id", keep="first")
    orders = orders.drop_duplicates()

    valid = (
        orders["order_id"].notna()
        & orders["order_date"].notna()
        & orders["restaurant_id"].notna()
        & orders["quantity_of_items"].gt(0)
        & orders["order_amount"].gt(0)
        & orders["delivery_time_taken_mins"].gt(0)
        & orders["customer_rating_food"].between(1, 5)
        & orders["customer_rating_delivery"].between(1, 5)
    )
    rejected = int((~valid).sum())
    return orders.loc[valid].copy(), duplicate_ids, rejected


def clean_restaurants(raw: pd.DataFrame) -> pd.DataFrame:
    restaurants = raw.copy()
    restaurants.columns = [snake_case(column) for column in restaurants.columns]
    restaurants = restaurants.rename(columns={
        "restaurantid": "restaurant_id",
        "restaurantname": "restaurant_name",
    })
    for column in restaurants.select_dtypes(include="object").columns:
        restaurants[column] = restaurants[column].astype("string").str.strip()
    restaurants["restaurant_id"] = pd.to_numeric(restaurants["restaurant_id"], errors="coerce")
    restaurants["restaurant_name"] = restaurants["restaurant_name"].str.replace(r"\s+", " ", regex=True).str.title()
    for column in ["cuisine", "zone", "category"]:
        restaurants[column] = restaurants[column].str.replace(r"\s+", " ", regex=True).str.title()
    return restaurants.dropna(subset=["restaurant_id"]).drop_duplicates(subset="restaurant_id", keep="first")


def add_features(curated: pd.DataFrame) -> pd.DataFrame:
    curated["order_month"] = curated["order_date"].dt.month_name()
    curated["day_of_week"] = curated["order_date"].dt.day_name()
    curated["is_weekend_order"] = curated["order_date"].dt.dayofweek.ge(5).map({True: "Yes", False: "No"})
    curated["delivery_speed"] = pd.cut(
        curated["delivery_time_taken_mins"],
        bins=[0, 20, 40, float("inf")], labels=["Fast", "Normal", "Slow"], include_lowest=True,
    ).astype("string")
    curated["average_rating"] = curated[["customer_rating_food", "customer_rating_delivery"]].mean(axis=1).round(2)
    return curated


def run_pipeline(orders_path: str | None = None, restaurants_path: str | None = None, output_dir: str | Path | None = None) -> tuple[pd.DataFrame, PipelineReport]:
    """Extract, clean, validate, join, enrich, and load the restaurant data."""
    orders_file = resolve_source(orders_path, "Orders.csv")
    restaurants_file = resolve_source(restaurants_path, "Restaurants.csv")
    output = Path(output_dir) if output_dir else PROJECT_DIR / "data" / "processed"
    output.mkdir(parents=True, exist_ok=True)

    raw_orders = pd.read_csv(orders_file)
    raw_restaurants = pd.read_csv(restaurants_file)
    orders, duplicate_ids, rejected = clean_orders(raw_orders)
    restaurants = clean_restaurants(raw_restaurants)
    curated = orders.merge(restaurants, on="restaurant_id", how="left", validate="many_to_one")
    unmatched = int(curated["restaurant_name"].isna().sum())
    curated = add_features(curated)
    curated = curated.sort_values("order_date").reset_index(drop=True)

    curated.to_csv(output / "processed_orders.csv", index=False, date_format="%Y-%m-%d %H:%M:%S")
    with sqlite3.connect(output / "restaurant_orders.db") as connection:
        curated.to_sql("curated_orders", connection, if_exists="replace", index=False)
        restaurants.to_sql("restaurants", connection, if_exists="replace", index=False)

    report = PipelineReport(
        source_orders=len(raw_orders), source_restaurants=len(raw_restaurants),
        curated_orders=len(curated), rejected_orders=rejected,
        unmatched_restaurants=unmatched, duplicate_order_ids=duplicate_ids,
        missing_values_after_cleaning=int(curated.isna().sum().sum()),
    )
    (output / "validation_report.json").write_text(json.dumps(asdict(report), indent=2), encoding="utf-8")
    return curated, report


def main() -> None:
    parser = argparse.ArgumentParser(description="Run the restaurant orders ETL pipeline.")
    parser.add_argument("--orders", help="Path to Orders.csv")
    parser.add_argument("--restaurants", help="Path to Restaurants.csv")
    parser.add_argument("--output-dir", help="Directory for curated outputs")
    args = parser.parse_args()
    _, report = run_pipeline(args.orders, args.restaurants, args.output_dir)
    print(json.dumps(asdict(report), indent=2))


if __name__ == "__main__":
    main()
