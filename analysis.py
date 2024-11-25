
import pandas as pd
import streamlit as st
import matplotlib.pyplot as plt
from statsmodels.tsa.holtwinters import ExponentialSmoothing

# Load the dataset
file_path = "ELB-Sales-Data.xlsx"  
data = pd.read_excel(file_path)

# Preprocess data
data.columns = data.columns.str.strip()  
data['Date'] = pd.to_datetime(data['Date'])  # Convert to datetime
data['Month'] = data['Date'].dt.strftime('%B') 

# Sidebar filters
st.sidebar.title("Filters")
hq = st.sidebar.selectbox("Select HQ", data['HQ'].unique())
month = st.sidebar.selectbox("Select Month", ['April', 'May', 'June', 'July', 'August', 'September'])

# Filter data
filtered_data = data[data['HQ'] == hq]
month_data = filtered_data[filtered_data['Month'] == month]

# Dashboard layout
st.title("Elbrit Life Sciences Sales Analysis Dashboard")

# Visualization Functions
def plot_bar(data, x, y, title):
    fig, ax = plt.subplots()
    data.groupby(x)[y].sum().plot(kind='bar', ax=ax)
    ax.set_title(title)
    st.pyplot(fig)

# Question 1: Highest-selling product in the selected month
if not month_data.empty:
    top_product = month_data.groupby('Item Name')['Primary Sales'].sum().idxmax()
    st.metric(f"Highest-selling product in {month}", top_product)
    plot_bar(month_data, 'Item Name', 'Primary Sales', f"Top Products in {month}")

# Question 2: Highest sales for CND Chennai in May
if hq == 'CND Chennai' and month == 'May':
    top_cnd_product = month_data.groupby('Item Name')['Primary Sales'].sum().idxmax()
    st.metric("Top Product for CND Chennai in May", top_cnd_product)

# Question 3: Max stock returns for Bangalore HQ in October
if hq == 'Bangalore' and month == 'October':
    top_return_customer = month_data.groupby('Customer')['Sales Return'].sum().idxmax()
    st.metric("Customer with max stock returns in October", top_return_customer)

# Question 4: Sales team with max % primary sales returned due to expiry
if not filtered_data.empty:
    expiry_returns = filtered_data.groupby('Sales Team')['Against Expiry'].sum()
    total_sales = filtered_data.groupby('Sales Team')['Primary Sales'].sum()
    return_rate = (expiry_returns / total_sales * 100).idxmax()
    st.metric("Team with highest % expiry returns", return_rate)

# Question 5: Percentage of sales affected by breakage
if 'Breakage' in data.columns:
    breakage_pct = (data['Breakage'].sum() / data['Primary Sales'].sum()) * 100
    st.metric("Overall % sales affected by breakage", f"{breakage_pct:.2f}%")

# Question 6: Primary sales for Delhi HQ in September
if hq == 'Delhi' and month == 'September':
    delhi_sales = month_data['Primary Sales'].sum()
    st.metric("Primary Sales for Delhi in September", f"{delhi_sales:.2f}")

# Question 7: Sales of Britorva 20 for PALEPU PHARMA in September
if month == 'September':
    britorva_sales = month_data[
        (month_data['Item Name'] == 'Britorva 20') &
        (month_data['Customer'] == 'PALEPU PHARMA DIST PVT LTD')
    ]['Primary Sales'].sum()
    st.metric("Britorva 20 Sales for PALEPU PHARMA in September", f"{britorva_sales:.2f}")

# Bonus: Forecasting
st.header("Forecasting Primary Sales")
time_series_data = data.groupby('Date')['Primary Sales'].sum().resample('M').sum()
if len(time_series_data) > 0:
    model = ExponentialSmoothing(time_series_data, trend='add', seasonal='add', seasonal_periods=12)
    model_fit = model.fit()
    forecast = model_fit.forecast(steps=2)
    st.line_chart(pd.DataFrame({'Actual': time_series_data, 'Forecast': forecast}))
    st.write("Forecasted values for next two months:")
    st.write(forecast)


