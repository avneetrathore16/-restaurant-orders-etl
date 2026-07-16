# Restaurant Orders Data Curation

An end-to-end Python ETL project and Streamlit dashboard for restaurant order analytics.

## What it does

- Profiles and cleans `Orders.csv` and `Restaurants.csv`
- Validates dates, quantities, amounts, delivery times, ratings, and duplicate IDs
- Joins orders to restaurant metadata and creates analytics features
- Loads a curated CSV, SQLite database, and JSON validation report
- Serves an interactive dashboard with filters and downloadable results

## Quick start

```bash
cd restaurant-orders-etl
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
python etl_pipeline.py --orders ~/Downloads/Orders.csv --restaurants ~/Downloads/Restaurants.csv
streamlit run dashboard.py
```

Raw data can alternatively be placed in `data/raw/Orders.csv` and `data/raw/Restaurants.csv`; the pipeline discovers those files automatically.

## Outputs

The pipeline writes these reproducible artifacts to `data/processed/`:

- `processed_orders.csv` — clean joined and enriched data
- `restaurant_orders.db` — SQLite tables: `curated_orders`, `restaurants`
- `validation_report.json` — row counts and validation outcomes

## Feature engineering

`order_month`, `day_of_week`, `is_weekend_order`, `delivery_speed`, and `average_rating`.
