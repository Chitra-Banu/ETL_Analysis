import pandas as pd
import streamlit as st
import plotly.express as px
from statsmodels.tsa.holtwinters import ExponentialSmoothing
import numpy as np

# Load the dataset
@st.cache_data
def load_data(file_path):
    data = pd.read_excel(file_path)
    return data

data = load_data("ELB-Sales-Data.xlsx")  

st.title("Elbrit Dashboard")

# Total Sales Metrics
st.header("Total Sales Metrics")

total_primary_sales = data["Primary Sales"].sum()
total_sales_return = data["Sales Return"].sum()
total_breakage = data["Breakage"].sum()
total_expiry_return = data["Against Expiry"].sum()

st.metric("Total Primary Sales", f"{total_primary_sales:,.2f}")
st.metric("Total Sales Return", f"{total_sales_return:,.2f}")
st.metric("Total Breakage", f"{total_breakage:,.2f}")
st.metric("Total Expiry Return", f"{total_expiry_return:,.2f}")

# Total Offers
st.subheader("Total Offers")

# Safely handle missing columns
total_claims = data["Claims"].sum() if "Claims" in data.columns else 0
free_quantity = data["Free Item"].sum() if "Free Item" in data.columns else 0
rate_difference = data["Rate Difference"].sum() if "Rate Difference" in data.columns else 0

offer_data = {
    "Claims": total_claims,
    "Free Quantity": free_quantity,
    "Rate Difference": rate_difference,
}

if sum(offer_data.values()) > 0:
    offer_df = pd.DataFrame(offer_data.items(), columns=["Offer Type", "Value"])
    fig_pie = px.pie(offer_df, names="Offer Type", values="Value", title="Offers Distribution")
    st.plotly_chart(fig_pie)
else:
    st.warning("No data available for Offers Distribution.")

# Sales Teams Primary Sales
st.subheader("September - Sales Teams Primary Sales")
if "Sales Team" in data.columns and "Primary Sales" in data.columns:
    team_sales = data.groupby("Sales Team")["Primary Sales"].sum().reset_index()
    fig_bar = px.bar(team_sales, x="Sales Team", y="Primary Sales", title="Sales Teams Primary Sales")
    st.plotly_chart(fig_bar)
else:
    st.warning("No data available for Sales Teams Primary Sales.")

# Forecasting Section
st.header("Forecasted Values")
time_series_data = data.groupby("Date")["Primary Sales"].sum().sort_index()


if len(time_series_data) >= 24:
    model = ExponentialSmoothing(
        time_series_data, trend="add", seasonal="add", seasonal_periods=12
    ).fit()

    forecast = model.forecast(2) 
    forecast_df = pd.DataFrame({
        "Month": ["October", "November"],
        "Forecasted Primary Sales": forecast.values,
    })

    st.write("Forecasted Primary Sales:")
    st.table(forecast_df)

  
    forecasted_metrics = {
        "Sales Return": data["Sales Return"].mean(),
        "Breakage": data["Breakage"].mean(),
        "Expiry Return": data["Against Expiry"].mean(),
    }
    st.write("Forecasted Metrics for October and November:")
    for metric, value in forecasted_metrics.items():
        st.metric(metric, f"{value:,.2f}")
else:
    st.warning("Not enough data points to perform forecasting. At least 24 data points are required.")


