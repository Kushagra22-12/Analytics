import pandas as pd
import streamlit as st
import plotly.express as px
from sklearn.linear_model import LinearRegression
import numpy as np

st.title("📊 Advanced Rolling Plan Analytics Dashboard")

# File Upload
uploaded_file = st.file_uploader("Upload Excel File", type=["xlsx"])

if uploaded_file:
    # Read Excel
    df = pd.read_excel(uploaded_file, sheet_name=0, header=2, engine='openpyxl')
    df.dropna(how='all', inplace=True)

    # Identify columns
    production_cols = [col for col in df.columns if 'Production' in str(col)]
    sales_cols = [col for col in df.columns if 'Sales' in str(col)]
    customer_col = 'Customer'
    product_col = 'Product'
    deficit_col = 'Dificit Qty.'

    # Melt data
    prod_data = df.melt(id_vars=[customer_col, product_col], value_vars=production_cols, var_name='Month', value_name='Production')
    sales_data = df.melt(id_vars=[customer_col, product_col], value_vars=sales_cols, var_name='Month', value_name='Sales')
    merged = pd.merge(prod_data, sales_data, on=[customer_col, product_col, 'Month'], how='outer')

    # Add deficit
    if deficit_col in df.columns:
        deficit_data = df[[customer_col, product_col, deficit_col]].dropna()
        merged = pd.merge(merged, deficit_data, on=[customer_col, product_col], how='left')

    # Extract month-year and convert to datetime
    merged['MonthText'] = merged['Month'].str.extract(r"(Oct|Nov|Dec|Jan|Feb)")
    merged['YearText'] = merged['Month'].str.extract(r"(\\d{2})")
    month_map = {'Oct': 10, 'Nov': 11, 'Dec': 12, 'Jan': 1, 'Feb': 2}
    merged['MonthNum'] = merged['MonthText'].map(month_map)
    merged['Year'] = pd.to_numeric(merged['YearText'], errors='coerce').fillna(25).astype(int) + 2000
    merged.loc[merged['MonthNum'].isin([1, 2]), 'Year'] += 1
    merged['Date'] = pd.to_datetime(dict(year=merged['Year'], month=merged['MonthNum'], day=1))

    # Convert numeric columns
    merged['Production'] = pd.to_numeric(merged['Production'], errors='coerce')
    merged['Sales'] = pd.to_numeric(merged['Sales'], errors='coerce')
    merged['Dificit Qty.'] = pd.to_numeric(merged['Dificit Qty.'], errors='coerce')

    # Sidebar Filters
    st.sidebar.header("🔍 Filters")
    selected_date = st.sidebar.selectbox("Select Month", sorted(merged['Date'].dropna().unique()))
    selected_customer = st.sidebar.multiselect("Select Customer", merged[customer_col].dropna().unique())
    selected_product = st.sidebar.multiselect("Select Product", merged[product_col].dropna().unique())

    # Apply filters
    filtered = merged[merged['Date'] == selected_date]
    if selected_customer:
        filtered = filtered[filtered[customer_col].isin(selected_customer)]
    if selected_product:
        filtered = filtered[filtered[product_col].isin(selected_product)]

    # KPI Cards
    st.subheader(f"📅 Summary for {selected_date.strftime('%B %Y')}")
    c1, c2, c3 = st.columns(3)
    c1.metric("Total Production", f"{filtered['Production'].sum():.2f}")
    c2.metric("Total Sales", f"{filtered['Sales'].sum():.2f}")
    c3.metric("Total Deficit", f"{filtered['Dificit Qty.'].sum():.2f}")

    # Monthly Trend Chart
    st.subheader("📈 Monthly Production vs Sales Trend")
    monthly_summary = merged.groupby('Date')[['Production', 'Sales']].sum().reset_index()
    fig_trend = px.line(monthly_summary, x='Date', y=['Production', 'Sales'], markers=True)
    st.plotly_chart(fig_trend)

    # Customer-wise Summary
    st.subheader("👥 Customer-wise Production and Sales")
    cust_summary = filtered.groupby(customer_col)[['Production', 'Sales']].sum().reset_index()
    fig_cust = px.bar(cust_summary, x=customer_col, y=['Production', 'Sales'], barmode='group')
    st.plotly_chart(fig_cust)

    # Comparison with Previous and Next Month
    date_list = sorted(merged['Date'].dropna().unique())
    date_index = date_list.index(selected_date)
    prev_date = date_list[date_index - 1] if date_index > 0 else None
    next_date = date_list[date_index + 1] if date_index < len(date_list) - 1 else None

    if prev_date:
        st.subheader(f"⬅️ Comparison with Previous Month: {prev_date.strftime('%B %Y')}")
        prev_data = merged[merged['Date'] == prev_date]
        comp_prev = pd.merge(filtered, prev_data, on=[customer_col, product_col], suffixes=('_current', '_prev'))
        comp_prev['Production_Diff'] = comp_prev['Production_current'] - comp_prev['Production_prev']
        comp_prev['Sales_Diff'] = comp_prev['Sales_current'] - comp_prev['Sales_prev']
        st.dataframe(comp_prev[[customer_col, product_col, 'Production_Diff', 'Sales_Diff']])

    if next_date:
        st.subheader(f"➡️ Comparison with Next Month: {next_date.strftime('%B %Y')}")
        next_data = merged[merged['Date'] == next_date]
        comp_next = pd.merge(filtered, next_data, on=[customer_col, product_col], suffixes=('_current', '_next'))
        comp_next['Production_Diff'] = comp_next['Production_next'] - comp_next['Production_current']
        comp_next['Sales_Diff'] = comp_next['Sales_next'] - comp_next['Sales_current']
        st.dataframe(comp_next[[customer_col, product_col, 'Production_Diff', 'Sales_Diff']])

    # Forecasting Next Month's Production and Sales
    st.subheader("🔮 Forecasting Next Month's Production and Sales")
    forecast_data = monthly_summary.dropna()
    forecast_data['MonthIndex'] = np.arange(len(forecast_data))
    model_prod = LinearRegression().fit(forecast_data[['MonthIndex']], forecast_data['Production'])
    model_sales = LinearRegression().fit(forecast_data[['MonthIndex']], forecast_data['Sales'])
    next_index = len(forecast_data)
    st.write(f"🔮 Forecasted Production: {model_prod.predict([[next_index]])[0]:.2f}")
    st.write(f"🔮 Forecasted Sales: {model_sales.predict([[next_index]])[0]:.2f}")

    # Anomaly Detection
    st.subheader("🚨 Anomaly Detection")
    prod_mean, prod_std = monthly_summary['Production'].mean(), monthly_summary['Production'].std()
    sales_mean, sales_std = monthly_summary['Sales'].mean(), monthly_summary['Sales'].std()
    anomalies = monthly_summary[
        (monthly_summary['Production'] > prod_mean + 2 * prod_std) |
        (monthly_summary['Production'] < prod_mean - 2 * prod_std) |
        (monthly_summary['Sales'] > sales_mean + 2 * sales_std) |
        (monthly_summary['Sales'] < sales_mean - 2 * sales_std)
    ]
    st.dataframe(anomalies)

    # Download Option
    st.subheader("📥 Download Filtered Data")
    st.download_button("Download CSV", data=filtered.to_csv(index=False).encode('utf-8'),
                       file_name="filtered_data.csv", mime="text/csv")
else:
    st.info("Please upload an Excel file to begin analysis.")
