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


if os.path.exists(TEMPLATE_PATH):
    with open(TEMPLATE_PATH, "rb") as f:
        with st.sidebar:
            st.markdown("---")
            st.markdown("### Download Template")
            st.download_button(
                label="Download Excel Template",
                data=f,
                file_name="wind_frequency_template.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
else:
    st.error("⚠️ Template file not found. Please run the 'create_wind_template.py' script first.")


with st.sidebar:
    st.markdown("---")
    st.markdown("### Design Parameters")
    rotor_area = st.number_input("Rotor Area (m²)", min_value=1.0, value=50.0, step=1.0)
    efficiency = st.number_input("Efficiency (–)", min_value=0.0, max_value=1.0, value=0.4, step=0.01)
    air_density = st.number_input("Air Density (kg/m³)", min_value=0.5, max_value=2.0, value=1.23, step=0.01)
    cut_in = st.number_input("Cut-in Speed (m/s)", min_value=0.0, value=2.0, step=0.5)
    cut_out = st.number_input("Cut-out Speed (m/s)", min_value=3.0, value=20.0, step=0.5)


uploaded_file = st.file_uploader("Upload your filled Excel file", type=["xlsx"])

if uploaded_file:
    try:
        df = pd.read_excel(uploaded_file)
        st.success("✅ File uploaded successfully!")
        with st.expander("Uploaded Data", expanded=False):
            st.dataframe(df.head())

        # -----------------------------------------------------------
        # Parse wind speed bins
        # -----------------------------------------------------------
        def parse_speed_bin(bin_str):
            """Parse the wind speed bin label into a numeric midpoint (handles all dash types and multi-digit numbers)."""
            if pd.isna(bin_str):
                return np.nan

            # Normalize all possible dash characters
            clean = str(bin_str).strip().replace("–", "-").replace("—", "-").replace("−", "-")

            # Split into parts only once (in case of weird values like "10-12 m/s")
            parts = clean.split("-", maxsplit=1)

            try:
                if len(parts) == 2:
                    v1 = float(re.findall(r"\d+\.?\d*", parts[0])[0])
                    v2 = float(re.findall(r"\d+\.?\d*", parts[1])[0])
                    return (v1 + v2) / 2
                elif len(parts) == 1 and re.search(r"\d", parts[0]):
                    return float(re.findall(r"\d+\.?\d*", parts[0])[0])
                else:
                    return np.nan
            except Exception:
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
        st.header("Results")


        results_df = pd.DataFrame({
            "Month": months,
            "Energy (kWh)": [monthly_energy[m] for m in months]
        })
        a, b = st.columns([3,1])
        with a:
            st.subheader("Monthly Energy Generation")
            fig = px.bar(results_df, x="Month", y="Energy (kWh)", text_auto=".0f",
                         title="Monthly Generated Electricity (kWh)",
                         labels={"Energy (kWh)": "Energy [kWh]"},
                         color_discrete_sequence=["#1f77b4"],  # 👈 custom color (hex or CSS)
                         )
            fig.update_traces(textposition="inside")
            fig.update_layout(showlegend=False)
            st.plotly_chart(fig, use_container_width=True)
        with b:
            st.metric("Total Annual Energy",f"{total_annual_energy/1000:.1f} MWh")

        with st.expander("Summary Table", expanded=False):
            st.subheader("Summary Table")
            st.dataframe(results_df.style.format({"Energy (kWh)": "{:,.0f}"}))



    except Exception as e:
        st.error(f"Error reading or processing file: {e}")

    st.subheader("Annual Wind Speed Distribution")

    # Compute total annual hours per wind speed bin
    df["Annual Hours"] = df[months].sum(axis=1)

    # Sort bins by wind speed midpoint
    df_sorted = df.sort_values("v_mid", ascending=True)

    # Create horizontal bar chart
    fig_dist = px.bar(
        df_sorted,
        y="Wind Speed Bin (m/s)",
        x="Annual Hours",
        orientation="h",
        text="Annual Hours",
        color="v_mid",
        color_continuous_scale=px.colors.sequential.Blues,
        title="Annual Wind Speed Frequency (hours per bin)",
        labels={"Annual Hours": "Hours", "Wind Speed Bin (m/s)": "Wind Speed Bin"},
    )
    fig_dist.update_traces(textposition="outside")
    fig_dist.update_layout(
        yaxis=dict(title="", categoryorder="array", categoryarray=df_sorted["Wind Speed Bin (m/s)"]),
        coloraxis_colorbar=dict(title="Wind Speed [m/s]"),
        showlegend=False,
        height=500,
    )
    st.plotly_chart(fig_dist, use_container_width=True)
else:
    st.info("⬆️ Please upload your filled Excel file to proceed with the calculations.")



with st.sidebar:
    st.markdown("---")
    st.caption("*A product of*")
    st.image("WS_Logo.png", width=300)
    st.caption("Werner Sobek Green Technologies GmbH")
    st.caption("Fachgruppe Simulation")
    st.markdown("---")
    st.caption("*Coded by*")
    st.caption("Rodrigo Carvalho")
    st.caption("*Need help? Contact me under:*")
    st.caption("*email:* rodrigo.carvalho@wernersobek.com")
    st.caption("*Tel* +49.40.6963863-14")
    st.caption("*Mob* +49.171.964.7850")