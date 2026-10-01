import streamlit as st
import os
import time

# Import your script function (assuming you name your file reit_engine.py)
from Reit_to_excel_withTNX import get_reit_data
output_filename = "REIT_Analysis_Dashboard.xlsx"

st.set_page_config(page_title="REIT Analysis Generator", page_icon="📈", layout="centered")

st.title("📊 Automated REIT Hybrid Screener & Dashboard")
st.markdown("""
This automated reporting engine evaluates global REITs by combining **fundamental metrics** (NAV, Dividend Yield, Payout Frequency), **debt risk tiers** (Gearing, Dynamic ICR), and **technical indicators** (50 & 200 EMAs) into a single executive Excel dashboard.
""")

st.info("💡 **Portfolio Demo Mode:** Click the button below to run the live pipeline, fetch live Yahoo Finance data, generate technical charts, and compile the final spreadsheet.")

if st.button("🚀 Run Report & Generate Dashboard", type="primary"):
    with st.spinner("Fetching live market data, scraping bond yields, and building Excel sheets... Please wait."):
        try:
            # Run your script function
            get_reit_data()
            
            # Check if file was successfully generated
            if os.path.exists(output_filename):
                st.success("✅ Report generated successfully!")
                
                # Provide download button for the user
                with open(output_filename, "rb") as file:
                    st.download_button(
                        label="📥 Download REIT_Analysis_Dashboard.xlsx",
                        data=file,
                        file_name=output_filename,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
            else:
                st.error("❌ Error: The spreadsheet file was not created.")
        except Exception as e:
            st.error(f"An error occurred during execution: {e}")

st.markdown("---")
st.markdown("### 🔍 What this script automates:")
st.markdown("""
* **Dynamic Benchmarking:** Automatically checks US 10-Year Treasury yields (`^TNX`) or scrapes Singapore Government Bond yields as a baseline risk-free rate.
* **Resilient Financial Parsing:** Computes dynamic Interest Coverage Ratios (ICR) from quarterly income statements when standard fields are missing.
* **Technical Charting:** Dynamically plots 50-day and 200-day Exponential Moving Averages (EMAs) using `matplotlib` and embeds the charts directly into individual Excel tabs.
* **Automated Styling:** Applies conditional color formatting based on custom signal logic (Accumulate vs. Decumulate) via `openpyxl`.
""")
