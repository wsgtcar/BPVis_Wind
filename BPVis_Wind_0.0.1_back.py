import streamlit as st
import pandas as pd
import numpy as np
import os
import re
import plotly.express as px

# -----------------------------------------------------------
# App setup
# -----------------------------------------------------------
st.set_page_config(page_title="WSGT_BPVis_Wind 0.0.1", page_icon="Pamo_Icon_White.png", layout="wide")



col1, col2 = st.columns(2)
with col2:
    st.image("WS_Logo.jpg", width=900)


st.sidebar.image("Pamo_Icon_Black.png", width=80)
st.sidebar.write("## BPVis Wind")
st.sidebar.write("Version 0.0.1")

st.title("Wind Power Potential Analysis")
st.markdown("""
This tool calculates the **monthly and annual electricity generation** (kWh)
from wind frequency data and turbine parameters.
""")
TEMPLATE_PATH = "wind_frequency_template.xlsx"

# -----------------------------------------------------------
# Section 1: Download Excel Template
# -----------------------------------------------------------
st.header("📘 Step 1: Download Wind Frequency Template")

st.write("""
Download the Excel template below and fill in your **monthly wind frequency data**
(in hours per wind speed bin).
""")

if os.path.exists(TEMPLATE_PATH):
    with open(TEMPLATE_PATH, "rb") as f:
        st.download_button(
            label="Download Excel Template",
            data=f,
            file_name="wind_frequency_template.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
else:
    st.error("⚠️ Template file not found. Please run the 'create_wind_template.py' script first.")

# -----------------------------------------------------------
# Section 2: User Inputs
# -----------------------------------------------------------
st.header("⚙️ Step 2: Turbine and Environmental Parameters")

col1, col2 = st.columns(2)
with col1:
    rotor_area = st.number_input("Rotor Area (m²)", min_value=1.0, value=50.0, step=1.0)
    efficiency = st.number_input("Efficiency (–)", min_value=0.0, max_value=1.0, value=0.4, step=0.01)
with col2:
    air_density = st.number_input("Air Density (kg/m³)", min_value=0.5, max_value=2.0, value=1.225, step=0.01)
    cut_in = st.number_input("Cut-in Speed (m/s)", min_value=0.0, value=3.0, step=0.5)
    cut_out = st.number_input("Cut-out Speed (m/s)", min_value=3.0, value=25.0, step=0.5)

# -----------------------------------------------------------
# Section 3: Upload User Data
# -----------------------------------------------------------
st.header("📤 Step 3: Upload Your Wind Frequency Data")

uploaded_file = st.file_uploader("Upload your filled Excel file", type=["xlsx"])

if uploaded_file:
    try:
        df = pd.read_excel(uploaded_file)
        st.success("✅ File uploaded successfully!")
        st.dataframe(df.head())

        # -----------------------------------------------------------
        # Parse wind speed bins
        # -----------------------------------------------------------
        def parse_speed_bin(bin_str):
            numbers = re.findall(r"[\d.]+", str(bin_str))
            if len(numbers) == 2:
                return (float(numbers[0]) + float(numbers[1])) / 2
            elif len(numbers) == 1:
                return float(numbers[0])
            else:
                return np.nan

        df["v_mid"] = df["Wind Speed Bin (m/s)"].apply(parse_speed_bin)

        # -----------------------------------------------------------
        # Filter by cut-in and cut-out speeds
        # -----------------------------------------------------------
        df = df[(df["v_mid"] >= cut_in) & (df["v_mid"] <= cut_out)]

        # -----------------------------------------------------------
        # Power calculation per bin (W)
        # -----------------------------------------------------------
        df["Power_W"] = 0.5 * air_density * rotor_area * (df["v_mid"] ** 3) * efficiency

        # -----------------------------------------------------------
        # Calculate energy per month (kWh)
        # -----------------------------------------------------------
        months = [m for m in df.columns if m not in ["Wind Speed Bin (m/s)", "v_mid", "Power_W"]]
        monthly_energy = {m: (df[m] * df["Power_W"]).sum() / 1000 for m in months}  # Wh → kWh
        total_annual_energy = sum(monthly_energy.values())

        # -----------------------------------------------------------
        # Results Display
        # -----------------------------------------------------------
        st.header("📊 Step 4: Results")

        results_df = pd.DataFrame({
            "Month": months,
            "Energy (kWh)": [monthly_energy[m] for m in months]
        })

        st.subheader("Monthly Energy Generation")
        fig = px.bar(results_df, x="Month", y="Energy (kWh)", text_auto=".0f",
                     title="Monthly Generated Electricity (kWh)",
                     labels={"Energy (kWh)": "Energy [kWh]"})
        fig.update_traces(textposition="outside")
        fig.update_layout(showlegend=False)
        st.plotly_chart(fig, use_container_width=True)

        st.subheader("Summary Table")
        st.dataframe(results_df.style.format({"Energy (kWh)": "{:,.0f}"}))

        st.success(f"💡 **Total Annual Electricity Generation:** {total_annual_energy:,.0f} kWh")

    except Exception as e:
        st.error(f"Error reading or processing file: {e}")
else:
    st.info("⬆️ Please upload your filled Excel file to proceed with the calculations.")
