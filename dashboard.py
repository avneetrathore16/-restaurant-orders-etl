"""Interactive Streamlit dashboard for curated restaurant orders."""

from pathlib import Path

import pandas as pd
import plotly.express as px
import streamlit as st

from etl_pipeline import PROJECT_DIR, run_pipeline

st.set_page_config(page_title="Restaurant Orders", page_icon="🍽️", layout="wide")


@st.cache_data(show_spinner="Preparing curated data…")
def get_data() -> pd.DataFrame:
    output_file = PROJECT_DIR / "data" / "processed" / "processed_orders.csv"
    if not output_file.exists():
        run_pipeline()
    data = pd.read_csv(output_file, parse_dates=["order_date"])
    return data


data = get_data()
st.title("🍽️ Restaurant Orders Intelligence")
st.caption("A curated analytics view built from Orders and Restaurants data.")

with st.sidebar:
    st.header("Filters")
    zones = st.multiselect("Zone", sorted(data["zone"].dropna().unique()))
    cuisines = st.multiselect("Cuisine", sorted(data["cuisine"].dropna().unique()))
    payment_modes = st.multiselect("Payment mode", sorted(data["payment_mode"].dropna().unique()))
    dates = st.date_input("Order date range", value=(data["order_date"].min().date(), data["order_date"].max().date()))

filtered = data.copy()
if zones:
    filtered = filtered[filtered["zone"].isin(zones)]
if cuisines:
    filtered = filtered[filtered["cuisine"].isin(cuisines)]
if payment_modes:
    filtered = filtered[filtered["payment_mode"].isin(payment_modes)]
if isinstance(dates, tuple) and len(dates) == 2:
    filtered = filtered[filtered["order_date"].dt.date.between(*dates)]

if filtered.empty:
    st.warning("No orders match these filters.")
    st.stop()

revenue = filtered["order_amount"].sum()
avg_delivery = filtered["delivery_time_taken_mins"].mean()
avg_rating = filtered["average_rating"].mean()
metrics = st.columns(4)
metrics[0].metric("Orders", f"{len(filtered):,}")
metrics[1].metric("Revenue", f"₹{revenue:,.0f}")
metrics[2].metric("Avg. delivery", f"{avg_delivery:.1f} min")
metrics[3].metric("Avg. rating", f"{avg_rating:.2f} / 5")

left, right = st.columns(2)
restaurant_orders = filtered.groupby("restaurant_name", as_index=False).agg(orders=("order_id", "count"), revenue=("order_amount", "sum")).sort_values("orders", ascending=False).head(12)
left.plotly_chart(px.bar(restaurant_orders, x="orders", y="restaurant_name", orientation="h", title="Top restaurants by orders", color="revenue", color_continuous_scale="Oranges"), width="stretch")
payment = filtered.groupby("payment_mode", as_index=False).size().rename(columns={"size": "orders"})
right.plotly_chart(px.pie(payment, names="payment_mode", values="orders", title="Payment modes", hole=.45), width="stretch")

left, right = st.columns(2)
left.plotly_chart(px.histogram(filtered, x="order_amount", nbins=25, title="Order amount distribution", color_discrete_sequence=["#ef8354"]), width="stretch")
right.plotly_chart(px.box(filtered, x="delivery_speed", y="delivery_time_taken_mins", category_orders={"delivery_speed": ["Fast", "Normal", "Slow"]}, title="Delivery time by speed tier", color="delivery_speed"), width="stretch")

left, right = st.columns(2)
cuisine = filtered.groupby("cuisine", as_index=False).size().rename(columns={"size": "orders"}).sort_values("orders", ascending=False)
left.plotly_chart(px.bar(cuisine, x="cuisine", y="orders", title="Cuisine distribution", color="orders", color_continuous_scale="Teal"), width="stretch")
zone = filtered.groupby("zone", as_index=False).agg(revenue=("order_amount", "sum")).sort_values("revenue", ascending=False)
right.plotly_chart(px.bar(zone, x="zone", y="revenue", title="Revenue by zone", color="revenue", color_continuous_scale="Blues"), width="stretch")

st.subheader("Curated orders")
st.dataframe(filtered.sort_values("order_date", ascending=False), width="stretch", hide_index=True)
st.download_button("Download filtered CSV", filtered.to_csv(index=False).encode("utf-8"), "filtered_restaurant_orders.csv", "text/csv")
