import streamlit as st
import pandas as pd
from datetime import datetime
import os

@st.cache_data
def load_excel_data(file_path, sheet_name):
    return pd.read_excel(file_path, sheet_name=sheet_name)
    
# Path to your Excel file
excel_file_path = "Simulator Rules_v02.xlsx"

# Load Excel data
file_path = 'data.xlsx'  # Path to your Excel file
sheet_name = 'data'  # Replace with your sheet name
df = load_excel_data(file_path, sheet_name=sheet_name)

with open(excel_file_path, "rb") as file:
    st.download_button(
        label="📂 Click Here to Download Reference File with Rules for this Simulator",
        data=file,
        file_name="Simulator Rules_v02.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

# Ensure 'Year Month' and 'Product Group' are integers for calculations
df['Year Month'] = df['Year Month'].fillna(0).astype(int)
df['Product Group'] = df['Product Group'].fillna(0).astype(int)

# CSS for styling
st.markdown("""
    <style>
    .header {
        color: red;
        font-weight: bold;
        font-size: 24px;
        text-align: center;
    }
    .full-width-table {
        width: 100% !important;
        overflow-x: auto;
        font-size: 18px; /* Increase font size for table values */
    }
    .full-width-table table {
        width: 100%;
        border-collapse: collapse;
        table-layout: fixed;
    }
    th, td {
        padding: 8px !important;
        text-align: left;
        font-weight: bold;
    }
    th {
        background-color: #f0f0f0;
    }
    .note {
        color: red;
        font-weight: bold;
        font-size: 18px; /* Increase font size for the note */
    }
    </style>
""", unsafe_allow_html=True)

# Header for the application title
st.markdown("<h1 class='header'>PAMA - Cost and Budget Simulator for Budget Preparations</h1>", unsafe_allow_html=True)

# Displaying all values in USD note
st.markdown("<p class='note'>Note : All the values are in USD</p>", unsafe_allow_html=True)

# Header for filters section
st.header("Filters Box")

# Dictionary to store selected filters
filters = {
    "LPGU": [],
    "Supplier Type": [],
    "Supplier GUID": [],
    "Location": [],
    "Business Line": [],
    "Supplier Name": [],
    "MDF": []
}

# Function to dynamically filter options based on selected values
def get_filtered_options(df, filters, selected_column):
    filtered_df = df.copy()
    for column, selected_values in filters.items():
        if selected_values and column != selected_column:
            filtered_df = filtered_df[filtered_df[column].isin(selected_values)]
    return filtered_df[selected_column].dropna().unique().tolist()

# Arrange top 6 filters in a 2x3 grid
cols = st.columns(4)
filter_columns = list(filters.keys())

for i, column in enumerate(filter_columns):
    with cols[i % 4]:  # Display 4 filters per row
        options = get_filtered_options(df, filters, column)
        selected_values = st.multiselect(f"Select {column}", options)
        filters[column] = selected_values if selected_values else options

# Divider for input parameters
st.header("Input Parameters")

# Input fields for spend and price change
input_cols = st.columns(4)

# Spend change
with input_cols[0]:
    spend_change = st.number_input("Spend change (%)", value=0.0)

# Price change
with input_cols[1]:
    price_change = st.number_input("Price change (%)", value=0.0) / 100  # Convert to decimal

# Year and month selection for spend and price change start dates
current_year = datetime.now().year
years = list(range(current_year, current_year + 6))  # Current year plus next 5 years
months = [f"{month:02d}" for month in range(1, 13)]

# Start date for spend change (year and month only)
with input_cols[2]:
    selected_year_spend = st.selectbox("Start date spend change - Year", years)
    selected_month_spend = st.selectbox("Start date spend change - Month", months, index=datetime.now().month - 1)

# Start date for price change (year and month only)
with input_cols[3]:
    selected_year_price = st.selectbox("Start date price change - Year", years)
    selected_month_price = st.selectbox("Start date price change - Month", months, index=datetime.now().month - 1)

# Combine selected year and month into YYYYMM format for any calculations or display purposes
start_date_spend_change = f"{selected_year_spend}{selected_month_spend}"
start_date_price_change = f"{selected_year_price}{selected_month_price}"

# Global max "Year Month" in the entire dataset (not affected by filters)
global_max_year_month = df['Year Month'].max()
global_max_month = int(str(global_max_year_month)[4:])  # Extract month from YYYYMM format
global_max_month_name = datetime.strptime(str(global_max_month), "%m").strftime("%B")  # Convert to month name

# Clear and Apply buttons
if st.button("Clear Filters"):
    for key in filters.keys():
        filters[key] = []  # Clear all selected filters

if st.button("Apply Filters"):
    filtered_df = df.copy()
    for column, selected_values in filters.items():
        if selected_values:
            filtered_df = filtered_df[filtered_df[column].isin(selected_values)]
    
    # YTD Rep Spend and monthly spend calculation using global max month
    ytd_rep_spend = round(filtered_df['Rep Spend'].sum())  # Round to integer for display
    monthly_spend = round(ytd_rep_spend / global_max_month) if global_max_month != 0 else 0
    
    # Generate columns for remaining months after global max month
    remaining_months = [datetime.strptime(str(m), "%m").strftime("%B") for m in range(global_max_month + 1, 13)]
    remaining_spends = {month: monthly_spend for month in remaining_months}
    
    # Calculate the total spend including YTD and projected monthly spends
    total_projected_spend = ytd_rep_spend + sum(remaining_spends.values())
    
    # Calculate new columns based on global max month
    filtered_df['Spend YE'] = (filtered_df['Spend'] / global_max_month) * 12
    filtered_df['Spend Rep. YE'] = (filtered_df['Rep Spend'] / global_max_month) * 12
    filtered_df['Spend Non Rep YE'] = (filtered_df['Non rep Spend'] / global_max_month) * 12
    filtered_df['NMI YE'] = (filtered_df['Total NMI Act'] / global_max_month) * 12
    filtered_df['NMI Rep YE'] = (filtered_df['NMI Rep Act'] / global_max_month) * 12
    filtered_df['NMI Non Rep YE'] = (filtered_df['NMI Non-Rep Act'] / global_max_month) * 12

    # Display Available Data
    st.markdown("### Available Data (Live Available Data in SCMIS)")
    available_data_df = filtered_df.drop(columns=["Spend YE", "Spend Rep. YE", "Spend Non Rep YE",
                                                  "NMI YE", "NMI Rep YE", "NMI Non Rep YE"]).copy()
    
    available_data_df["Year Month"] = available_data_df["Year Month"].astype(str)
    available_data_df["Product Group"] = available_data_df["Product Group"].astype(str)

    financial_columns = [
        "Invoice Quantity", "Invoice QTY Rep", "Invoice QTY Non Rep", 
        "Spend", "Rep Spend", "Non rep Spend", "Total NMI Act", 
        "NMI Rep Act", "NMI Non-Rep Act"
    ]
    
    for col in financial_columns:
        available_data_df[col] = available_data_df[col].apply(lambda x: f"{int(x):,}" if pd.notnull(x) and isinstance(x, (int, float)) else x)


    st.markdown("<div class='full-width-table'>", unsafe_allow_html=True)
    st.table(available_data_df.reset_index(drop=True))

    # Display Calculations
    total_invoice_qty_rep = filtered_df['Invoice QTY Rep'].sum()
    average_price_ytd = ytd_rep_spend / total_invoice_qty_rep if total_invoice_qty_rep != 0 else 0
    average_new_price = average_price_ytd * (1 - price_change)
    monthly_volumes_impacted = total_invoice_qty_rep / global_max_month if global_max_month != 0 else 0
    delta_nmi_impact_monthly = (average_price_ytd - average_new_price) * monthly_volumes_impacted

    st.markdown("### Calculations (Required for NMI Calculations)")
    calculations_df = pd.DataFrame({
        "Average Price YTD": [average_price_ytd],
        "Average New Price": [average_new_price],
        "Monthly Volumes Impacted": [monthly_volumes_impacted],
        "Delta NMI Impact | Monthly": [delta_nmi_impact_monthly]
    })
    st.markdown("<div class='full-width-table'>", unsafe_allow_html=True)
    st.table(calculations_df)

    # Display Grand Total
    grand_total_values = filtered_df[financial_columns].sum().astype(int)
    grand_total_df = pd.DataFrame(grand_total_values).T
    grand_total_df.index = ["YTD Values"]
    for col in financial_columns:
        grand_total_df[col] = grand_total_df[col].apply(lambda x: f"{x:,}")

    # Rep Spend Extrapolation Table
    st.markdown("### Rep Spend Extrapolation (Linearly Extrapolated Values)")
    ytd_spend_data = {
        f"YTD Spend with {global_max_month_name}": [ytd_rep_spend],
        **{month: monthly_spend for month in remaining_months},
        "Total": [total_projected_spend]
    }
    ytd_spend_df = pd.DataFrame(ytd_spend_data)
    ytd_spend_df = ytd_spend_df.applymap(lambda x: f"{int(x):,}")
    st.markdown("<div class='full-width-table'>", unsafe_allow_html=True)
    st.table(ytd_spend_df)

    # Rep Spend Variation Table
    st.markdown("### Rep. Spend Variation (with applied (%) change)")
    adjusted_remaining_spends = {}
    spend_start_yyyymm = int(f"{selected_year_spend}{selected_month_spend}")
    price_start_yyyymm = int(f"{selected_year_price}{selected_month_price}")
    
    for month_name in remaining_months:
        month_num = datetime.strptime(month_name, "%B").month
        yyyymm = int(f"{current_year}{month_num:02d}")
        adjusted_value = monthly_spend
        if yyyymm >= spend_start_yyyymm:
            adjusted_value *= (1 - spend_change / 100)
        if yyyymm >= price_start_yyyymm:
            adjusted_value *= (1 - price_change)
        adjusted_remaining_spends[month_name] = round(adjusted_value)
    
    adjusted_total_projected_spend = ytd_rep_spend + sum(adjusted_remaining_spends.values())
    adjusted_ytd_spend_data = {
        f"YTD Spend with {global_max_month_name}": [ytd_rep_spend],
        **adjusted_remaining_spends,
        "Total": [adjusted_total_projected_spend]
    }
    adjusted_ytd_spend_df = pd.DataFrame(adjusted_ytd_spend_data)
    adjusted_ytd_spend_df = adjusted_ytd_spend_df.applymap(lambda x: f"{int(x):,}")
    st.markdown("<div class='full-width-table'>", unsafe_allow_html=True)
    st.table(adjusted_ytd_spend_df)

    # NMI Rep Variation Table
    st.markdown("### NMI Rep Variation (with applied (%) change)")
    nmi_rep_act = filtered_df['NMI Rep Act'].sum()
    monthly_nmi_rep_value = round(nmi_rep_act / global_max_month) if global_max_month != 0 else 0
    adjusted_nmi_rep_values = {}
    for month_name in remaining_months:
        month_num = datetime.strptime(month_name, "%B").month
        yyyymm = int(f"{current_year}{month_num:02d}")
        if yyyymm >= price_start_yyyymm:
            adjusted_nmi_rep_value = monthly_nmi_rep_value + delta_nmi_impact_monthly
        else:
            adjusted_nmi_rep_value = monthly_nmi_rep_value
        adjusted_nmi_rep_values[month_name] = round(adjusted_nmi_rep_value)

    total_nmi_rep_variation = nmi_rep_act + sum(adjusted_nmi_rep_values.values())
    nmi_rep_variation_data = {
        f"NMI Rep with {global_max_month_name}": [nmi_rep_act],
        **adjusted_nmi_rep_values,
        "Total": [total_nmi_rep_variation]
    }
    nmi_rep_variation_df = pd.DataFrame(nmi_rep_variation_data)
    nmi_rep_variation_df = nmi_rep_variation_df.applymap(lambda x: f"{int(x):,}")
    st.markdown("<div class='full-width-table'>", unsafe_allow_html=True)
    st.table(nmi_rep_variation_df)

    # Additional Calculations for Simulated, YE Extrapolated, and YTD Values
    spend_non_rep_sim = filtered_df['Spend Non Rep YE'].sum()
    nmi_non_rep_sim = filtered_df['NMI Non Rep YE'].sum()
    spend_rep_sim = adjusted_total_projected_spend  # Total from Rep Spend Variation
    nmi_rep_sim = total_nmi_rep_variation  # Total from NMI Rep Variation
    spend_sim = spend_rep_sim + spend_non_rep_sim
    nmi_sim = nmi_rep_sim + nmi_non_rep_sim

    # DataFrames for Simulated, YE Extrapolated, and YTD Values tables
    simulated_values_df = pd.DataFrame([{
        "Spend Sim": spend_sim,
        "Spend Rep. Sim": spend_rep_sim,
        "Spend Non Rep Sim": spend_non_rep_sim,
        "NMI Sim": nmi_sim,
        "NMI Rep Sim": nmi_rep_sim,
        "NMI Non Rep Sim": nmi_non_rep_sim
    }])

    ye_extrapolated_df = filtered_df[["Spend YE", "Spend Rep. YE", "Spend Non Rep YE", 
                                      "NMI YE", "NMI Rep YE", "NMI Non Rep YE"]].sum().to_frame().T

    ytd_values_df = filtered_df[["Invoice Quantity", "Invoice QTY Rep", "Invoice QTY Non Rep", 
                                 "Spend", "Rep Spend", "Non rep Spend", "Total NMI Act", 
                                 "NMI Rep Act", "NMI Non-Rep Act"]].sum().to_frame().T

    # Format DataFrames
    for col in simulated_values_df.columns:
        simulated_values_df[col] = simulated_values_df[col].apply(lambda x: f"{int(x):,}")
    for col in ye_extrapolated_df.columns:
        ye_extrapolated_df[col] = ye_extrapolated_df[col].apply(lambda x: f"{int(x):,}")
    for col in ytd_values_df.columns:
        ytd_values_df[col] = ytd_values_df[col].apply(lambda x: f"{int(x):,}")

    # Display tables with headers
    st.markdown("### Simulated Values (Final values with Percentage Change)")
    st.markdown("<div class='full-width-table'>", unsafe_allow_html=True)
    st.table(simulated_values_df)

    st.markdown("### Year End Linearly Extrapolated Values")
    st.markdown("<div class='full-width-table'>", unsafe_allow_html=True)
    st.table(ye_extrapolated_df)

    st.markdown("### YTD Available Values")
    st.markdown("<div class='full-width-table'>", unsafe_allow_html=True)
    st.table(ytd_values_df)

else:
    st.write("Use the filters above and click 'Apply Filters' to see the summarized results.")
