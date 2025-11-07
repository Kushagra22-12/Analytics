import pandas as pd
import streamlit as st
import plotly.express as px

st.title("📊 Rolling Plan Analytics Dashboard")

# ✅ File Upload
uploaded_file = st.file_uploader("Upload Excel File", type=["xlsx"])

if uploaded_file:
    # Read Excel
    df = pd.read_excel(uploaded_file, sheet_name=0, header=2, engine='openpyxl')
    df.dropna(how='all', inplace=True)

    # Extract relevant columns
    date_columns = [col for col in df.columns if isinstance(col, str) and any(month in col for month in ['Oct', 'Nov', 'Dec', 'Jan', 'Feb'])]
    production_columns = [col for col in date_columns if 'Production' in col]
    sales_columns = [col for col in date_columns if 'Sales' in col]
    customer_column = 'Customer'
    product_column = 'Product'
    deficit_column = 'Dificit Qty.'

    # Melt production and sales data
    production_data = df.melt(id_vars=[customer_column, product_column], value_vars=production_columns, var_name='Date', value_name='Production')
    sales_data = df.melt(id_vars=[customer_column, product_column], value_vars=sales_columns, var_name='Date', value_name='Sales')

    # Merge production and sales
    merged_data = pd.merge(production_data, sales_data, on=[customer_column, product_column, 'Date'], how='outer')

    # Add deficit column
    if deficit_column in df.columns:
        deficit_data = df[[customer_column, product_column, deficit_column]].dropna()
        merged_data = pd.merge(merged_data, deficit_data, on=[customer_column, product_column], how='left')

    # ✅ Extract month-year and convert to datetime
    merged_data['MonthText'] = merged_data['Date'].str.extract(r"(Oct|Nov|Dec|Jan|Feb)")
    merged_data['YearText'] = merged_data['Date'].str.extract(r"(\\d{2})")

    # Map month abbreviations to numbers
    month_map = {'Oct': 10, 'Nov': 11, 'Dec': 12, 'Jan': 1, 'Feb': 2}
    merged_data['Month'] = merged_data['MonthText'].map(month_map)

    # ✅ Handle NaN and convert year
    merged_data['Year'] = pd.to_numeric(merged_data['YearText'], errors='coerce').fillna(25).astype(int) + 2000

    # Adjust year for Jan & Feb (next year)
    merged_data.loc[merged_data['Month'].isin([1, 2]), 'Year'] += 1

    # Create proper datetime
    merged_data['Date'] = pd.to_datetime(merged_data[['Year', 'Month']].assign(DAY=1))

    # Convert numeric columns
    merged_data['Production'] = pd.to_numeric(merged_data['Production'], errors='coerce')
    merged_data['Sales'] = pd.to_numeric(merged_data['Sales'], errors='coerce')
    merged_data['Dificit Qty.'] = pd.to_numeric(merged_data['Dificit Qty.'], errors='coerce')

    # Sidebar Filters
    st.sidebar.header("🔍 Filters")
    selected_date = st.sidebar.selectbox("Select Month", sorted(merged_data['Date'].unique()))
    selected_customer = st.sidebar.multiselect("Select Customer", merged_data[customer_column].dropna().unique())
    selected_product = st.sidebar.multiselect("Select Product", merged_data[product_column].dropna().unique())

    # Apply filters
    filtered_data = merged_data[merged_data['Date'] == selected_date]
    if selected_customer:
        filtered_data = filtered_data[filtered_data[customer_column].isin(selected_customer)]
    if selected_product:
        filtered_data = filtered_data[filtered_data[product_column].isin(selected_product)]

    # KPI Cards
    total_production = filtered_data['Production'].sum()
    total_sales = filtered_data['Sales'].sum()
    total_deficit = filtered_data['Dificit Qty.'].sum()

    st.subheader(f"📅 Summary for {selected_date.strftime('%B %Y')}")
    col1, col2, col3 = st.columns(3)
    col1.metric("Total Production", f"{total_production:.2f}")
    col2.metric("Total Sales", f"{total_sales:.2f}")
    col3.metric("Total Deficit", f"{total_deficit:.2f}")

    # Charts
    st.subheader("📈 Monthly Production vs Sales Trend")
    monthly_summary = merged_data.groupby('Date')[['Production', 'Sales']].sum().reset_index()
    fig_monthly = px.line(monthly_summary, x='Date', y=['Production', 'Sales'], markers=True)
    st.plotly_chart(fig_monthly)

    st.subheader("👥 Customer-wise Production and Sales")
    customer_summary = filtered_data.groupby(customer_column)[['Production', 'Sales']].sum().reset_index()
    fig_customer = px.bar(customer_summary, x=customer_column, y=['Production', 'Sales'], barmode='group')
    st.plotly_chart(fig_customer)

    # Comparison with Previous and Next Month
    date_list = sorted(merged_data['Date'].unique())
    date_index = date_list.index(selected_date)
    prev_date = date_list[date_index - 1] if date_index > 0 else None
    next_date = date_list[date_index + 1] if date_index < len(date_list) - 1 else None

    if prev_date:
        st.subheader(f"⬅️ Comparison with Previous Month: {prev_date.strftime('%B %Y')}")
        prev_data = merged_data[merged_data['Date'] == prev_date]
        comparison_prev = pd.merge(filtered_data, prev_data, on=[customer_column, product_column], suffixes=('_current', '_prev'))
        comparison_prev['Production_Diff'] = comparison_prev['Production_current'] - comparison_prev['Production_prev']
        comparison_prev['Sales_Diff'] = comparison_prev['Sales_current'] - comparison_prev['Sales_prev']
        st.dataframe(comparison_prev[[customer_column, product_column, 'Production_Diff', 'Sales_Diff']])

    if next_date:
        st.subheader(f"➡️ Comparison with Next Month: {next_date.strftime('%B %Y')}")
        next_data = merged_data[merged_data['Date'] == next_date]
        comparison_next = pd.merge(filtered_data, next_data, on=[customer_column, product_column], suffixes=('_current', '_next'))
        comparison_next['Production_Diff'] = comparison_next['Production_next'] - comparison_next['Production_current']
        comparison_next['Sales_Diff'] = comparison_next['Sales_next'] - comparison_next['Sales_current']
        st.dataframe(comparison_next[[customer_column, product_column, 'Production_Diff', 'Sales_Diff']])

    # Download CSV
    st.subheader("📥 Download Summary as CSV")
    csv = filtered_data.to_csv(index=False).encode('utf-8')
    st.download_button("Download Filtered Data", data=csv, file_name="filtered_summary.csv", mime="text/csv")
else:
    st.info("Please upload an Excel file to start analysis.")
