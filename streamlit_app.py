"""
RENEWABLE ENERGY OPTIMIZATION TOOL - STREAMLIT VERSION
=======================================================
Cloud-based web application for hybrid renewable energy system optimization
PV + Wind + Hydro + BESS with Firm Capacity Analysis
"""

import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
import plotly.express as px
from datetime import datetime
import io
import sys
import os
from io import BytesIO

# Try to import optimization code
try:
    from optimize_gridsearch_hydro_static_HOMERNPCFIXED_COMB import (
        read_inputs,
        grid_search_optimize_hydro,
        find_optimal_solution
    )
    OPTIMIZATION_AVAILABLE = True
except ImportError:
    OPTIMIZATION_AVAILABLE = False
def build_input_excel_from_streamlit(
    # Configuration
    simulation_hours, 
    target_unmet_percent,  # <-- CRITICAL
    discount_rate, 
    inflation_rate, 
    project_lifetime,
    # PV
    pv_min, pv_max, pv_step, pv_capex, pv_opex, pv_lifetime, pv_lcoe,
    # Wind
    wind_min, wind_max, wind_step, wind_capex, wind_opex, wind_lifetime, wind_lcoe,
    # Hydro
    hydro_min, hydro_max, hydro_step, 
    hydro_hours_per_day,  # <-- CHANGED from hydro_window_min, hydro_window_max
    hydro_capex, hydro_opex, hydro_lifetime, hydro_lcoe,
    # BESS
    bess_min, bess_max, bess_step, bess_duration, bess_min_soc, bess_max_soc,
    bess_charge_eff, bess_discharge_eff, bess_power_capex, bess_energy_capex,
    bess_opex, bess_lifetime, bess_replacement_cost,
    # Profiles
    load_profile_df, 
    pv_profile_df, 
    wind_profile_df, 
    hydro_profile_df=None
):
    """Build complete Excel input file from Streamlit parameters."""
    
    output = BytesIO()
    
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        
        # Configuration sheet
        config_data = {
            'Parameter': ['Simulation Hours', 'Target Unmet Load (%)', 'Optimization Method',
                         'Discount Rate (%)', 'Inflation Rate (%)', 'Project Lifetime (years)',
                         'Use Dynamic LCOE'],
            'Value': [simulation_hours, target_unmet_percent, 'GRID_SEARCH',
                     discount_rate, inflation_rate, project_lifetime, 'NO']
        }
        pd.DataFrame(config_data).to_excel(writer, sheet_name='Configuration', index=False)
        
        # Grid Search Config sheet
        grid_config_data = {
            'Component': ['PV', 'Wind', 'Hydro', 'BESS'],
            'Min_Capacity': [pv_min, wind_min, hydro_min, bess_min],
            'Max_Capacity': [pv_max, wind_max, hydro_max, bess_max],
            'Step_Size': [pv_step, wind_step, hydro_step, bess_step]
        }
        pd.DataFrame(grid_config_data).to_excel(writer, sheet_name='Grid_Search_Config', index=False)
        
        # Solar_PV sheet
        solar_data = {
            'Parameter': ['Initial Capacity (MW)', 'Min Capacity (MW)', 'Max Capacity (MW)',
                         'Step Size (MW)', 'CapEx ($/kW)', 'OpEx ($/kW/year)',
                         'Lifetime (years)', 'LCOE ($/MWh)', 'Derating Factor (%)'],
            'Value': [pv_min, pv_min, pv_max, pv_step, pv_capex, pv_opex,
                     pv_lifetime, pv_lcoe, 100]
        }
        pd.DataFrame(solar_data).to_excel(writer, sheet_name='Solar_PV', index=False)
        
        # Wind sheet
        wind_data = {
            'Parameter': ['Initial Capacity (MW)', 'Min Capacity (MW)', 'Max Capacity (MW)',
                         'Step Size (MW)', 'CAPEX ($/kW)', 'OPEX ($/kW/year)',
                         'Lifetime (years)', 'LCOE ($/MWh)', 'Hub Height (m)',
                         'Derating Factor (%)'],
            'Value': [wind_min, wind_min, wind_max, wind_step, wind_capex, wind_opex,
                     wind_lifetime, wind_lcoe, 80, 100]
        }
        pd.DataFrame(wind_data).to_excel(writer, sheet_name='Wind', index=False)
        
       # Hydro sheet
        hydro_data = {
            'Parameter': ['Initial Capacity (MW)', 'Min Capacity (MW)', 'Max Capacity (MW)',
                         'Step Size (MW)', 'CapEx ($/kW)', 'OpEx ($/kW/year)',
                         'Lifetime (years)', 'LCOE ($/MWh)', 'Operating Hours'],
            'Value': [hydro_min, hydro_min, hydro_max, hydro_step, 
                     hydro_capex, hydro_opex, hydro_lifetime, hydro_lcoe,
                     hydro_hours_per_day]
        }
        pd.DataFrame(hydro_data).to_excel(writer, sheet_name='Hydro', index=False)
        
        # BESS sheet
        bess_data = {
            'Parameter': ['Initial Power (MW)', 'Min Power (MW)', 'Max Power (MW)',
                         'Step Size (MW)', 'Duration (hours)', 'Min SOC (%)', 'Max SOC (%)',
                         'Charging Efficiency (%)', 'Discharging Efficiency (%)',
                         'Power CAPEX ($/kW)', 'Energy CAPEX ($/kWh)', 'OPEX ($/kWh/year)',
                         'Lifetime (years)', 'Replacement Cost (%)'],
            'Value': [bess_min, bess_min, bess_max, bess_step, bess_duration,
                     bess_min_soc, bess_max_soc, bess_charge_eff, bess_discharge_eff,
                     bess_power_capex, bess_energy_capex, bess_opex,
                     bess_lifetime, bess_replacement_cost]
        }
        pd.DataFrame(bess_data).to_excel(writer, sheet_name='BESS', index=False)
        
        # Profile sheets
        load_profile_df.to_excel(writer, sheet_name='Load_Profile', index=False)
        pv_profile_df.to_excel(writer, sheet_name='PVsyst_Profile', index=False)
        wind_profile_df.to_excel(writer, sheet_name='Wind_Profile', index=False)
        
        # Hydro profile
        if hydro_profile_df is not None:
            hydro_profile_df.to_excel(writer, sheet_name='Hydro_Profile', index=False)
        else:
            default_hydro = pd.DataFrame({
                'Hour': range(8760),
                'Output_kW': [1.0] * 8760
            })
            default_hydro.to_excel(writer, sheet_name='Hydro_Profile', index=False)
    
    output.seek(0)
    return output
# Page configuration
st.set_page_config(
    page_title="RE Optimization Tool",
    page_icon="🌞",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS for better styling
st.markdown("""
<style>
    .main-header {
        font-size: 2.5rem;
        font-weight: bold;
        color: #1f77b4;
        text-align: center;
        margin-bottom: 2rem;
    }
    .metric-card {
        background-color: #f0f2f6;
        padding: 1rem;
        border-radius: 0.5rem;
        margin: 0.5rem 0;
    }
    .stButton>button {
        width: 100%;
        background-color: #1f77b4;
        color: white;
        font-weight: bold;
    }
</style>
""", unsafe_allow_html=True)

# Title
st.markdown('<p class="main-header">🌞 Renewable Energy Optimization Tool</p>', unsafe_allow_html=True)
st.markdown("**Hybrid System Design: PV + Wind + Hydro + Battery Storage**")
st.markdown("---")

# Initialize session state
if 'optimization_complete' not in st.session_state:
    st.session_state.optimization_complete = False
if 'results' not in st.session_state:
    st.session_state.results = None

# Sidebar for inputs
with st.sidebar:
    st.header("⚙️ System Configuration")
    
    # ========================================================================
    # SOLAR PV CONFIGURATION
    # ========================================================================
    with st.expander("☀️ SOLAR PV", expanded=True):
        st.subheader("Capacity Range")
        col1, col2 = st.columns(2)
        with col1:
            pv_min = st.number_input("Min (MW)", value=4.0, min_value=0.0, step=0.5, key="pv_min")
        with col2:
            pv_max = st.number_input("Max (MW)", value=10.0, min_value=0.0, step=0.5, key="pv_max")
        pv_step = st.number_input("Step (MW)", value=0.5, min_value=0.1, step=0.1, key="pv_step")
        
        st.subheader("Financial Parameters")
        col1, col2 = st.columns(2)
        with col1:
            pv_capex = st.number_input("CAPEX ($/kW)", value=1000, step=10, key="pv_capex")
            pv_opex = st.number_input("OPEX ($/kW/yr)", value=10, step=1, key="pv_opex")
        with col2:
            pv_lifetime = st.number_input("Lifetime (years)", value=25, step=1, key="pv_life")
            pv_lcoe = st.number_input("LCOE ($/MWh)", value=35.0, step=1.0, key="pv_lcoe")
    
    # ========================================================================
    # WIND CONFIGURATION
    # ========================================================================
    with st.expander("💨 WIND"):
        st.subheader("Capacity Range")
        col1, col2 = st.columns(2)
        with col1:
            wind_min = st.number_input("Min (MW)", value=0.0, min_value=0.0, step=0.5, key="wind_min")
        with col2:
            wind_max = st.number_input("Max (MW)", value=3.0, min_value=0.0, step=0.5, key="wind_max")
        wind_step = st.number_input("Step (MW)", value=0.5, min_value=0.1, step=0.1, key="wind_step")
        
        st.subheader("Financial Parameters")
        col1, col2 = st.columns(2)
        with col1:
            wind_capex = st.number_input("CAPEX ($/kW)", value=1200, step=10, key="wind_capex")
            wind_opex = st.number_input("OPEX ($/kW/yr)", value=15, step=1, key="wind_opex")
        with col2:
            wind_lifetime = st.number_input("Lifetime (years)", value=20, step=1, key="wind_life")
            wind_lcoe = st.number_input("LCOE ($/MWh)", value=45.0, step=1.0, key="wind_lcoe")
    
   # ========================================================================
# HYDRO CONFIGURATION
# ========================================================================
with st.expander("💧 HYDRO"):
    st.subheader("Capacity Range")
    col1, col2 = st.columns(2)
    with col1:
        hydro_min = st.number_input("Min (MW)", value=0.0, min_value=0.0, step=0.5, key="hydro_min")
    with col2:
        hydro_max = st.number_input("Max (MW)", value=2.0, min_value=0.0, step=0.5, key="hydro_max")
    hydro_step = st.number_input("Step (MW)", value=0.5, min_value=0.1, step=0.1, key="hydro_step")
    
    # ADD THIS SECTION:
    st.subheader("Operating Configuration")
    hydro_hours_per_day = st.number_input(
        "Operating Hours (hours/day)", 
        value=6, 
        min_value=1, 
        max_value=24, 
        step=1,
        key="hydro_hours",
        help="How many consecutive hours per day should hydro operate? System will auto-optimize the time window."
    )
    
    st.info(f"💡 System will test all possible {hydro_hours_per_day}-hour windows and find the optimal operating time")
    
    # Example windows
    if hydro_hours_per_day < 24:
        st.caption(f"Example windows to test: 00:00-{hydro_hours_per_day:02d}:00, 01:00-{hydro_hours_per_day+1:02d}:00, etc.")
    
    # REMOVE these old parameters (if they exist):
    # hydro_window_min = ...  # DELETE
    # hydro_window_max = ...  # DELETE
    
    st.subheader("Financial Parameters")
    col1, col2 = st.columns(2)
    with col1:
        hydro_capex = st.number_input("CapEx ($/kW)", value=2000, step=10, key="hydro_capex")
        hydro_opex = st.number_input("OpEx ($/kW/yr)", value=20, step=1, key="hydro_opex")
    with col2:
        hydro_lifetime = st.number_input("Lifetime (years)", value=50, step=1, key="hydro_life")
        hydro_lcoe = st.number_input("LCOE ($/MWh)", value=40.0, step=1.0, key="hydro_lcoe")
    
    # ========================================================================
    # BESS CONFIGURATION
    # ========================================================================
    # ========================================================================
    # BESS CONFIGURATION - UPDATED FOR YOUR OPTIMIZATION CODE
    # ========================================================================
    with st.expander("🔋 BATTERY STORAGE"):
        st.subheader("Power Range")
        col1, col2 = st.columns(2)
        with col1:
            bess_min = st.number_input("Min Power (MW)", value=5.0, min_value=0.0, step=1.0, key="bess_min")
        with col2:
            bess_max = st.number_input("Max Power (MW)", value=20.0, min_value=0.0, step=1.0, key="bess_max")
        bess_step = st.number_input("Step (MW)", value=1.0, min_value=0.1, step=0.1, key="bess_step")
        
        st.subheader("Storage Parameters")
        col1, col2 = st.columns(2)
        with col1:
            bess_duration = st.number_input("Duration (hours)", value=4.0, min_value=0.5, step=0.5, key="bess_dur", 
                                           help="Energy capacity = Power × Duration")
            bess_min_soc = st.number_input("Min SOC (%)", value=20, min_value=0, max_value=100, key="bess_min_soc",
                                          help="Minimum State of Charge - Usually 20%")
            bess_charge_eff = st.number_input("Charging Efficiency (%)", value=95, min_value=50, max_value=100, key="bess_charge_eff",
                                             help="Energy efficiency when charging")
        with col2:
            bess_lifetime = st.number_input("Lifetime (years)", value=15, min_value=1, max_value=30, step=1, key="bess_life")
            bess_max_soc = st.number_input("Max SOC (%)", value=100, min_value=0, max_value=100, key="bess_max_soc",
                                          help="Maximum State of Charge - Usually 100%")
            bess_discharge_eff = st.number_input("Discharging Efficiency (%)", value=95, min_value=50, max_value=100, key="bess_discharge_eff",
                                                help="Energy efficiency when discharging")
        
        st.subheader("Financial Parameters")
        col1, col2 = st.columns(2)
        with col1:
            bess_power_capex = st.number_input("Power CAPEX ($/kW)", value=300, step=10, key="bess_power_capex",
                                              help="Capital cost for power capacity (inverter, etc)")
            bess_energy_capex = st.number_input("Energy CAPEX ($/kWh)", value=200, step=10, key="bess_energy_capex",
                                               help="Capital cost for energy storage (battery cells)")
        with col2:
            bess_opex = st.number_input("OPEX ($/kWh/yr)", value=2, step=1, key="bess_opex",
                                       help="Annual operating & maintenance cost")
            bess_replacement_cost = st.number_input("Replacement Cost (%)", value=80, min_value=0, max_value=100, key="bess_replacement",
                                                   help="Replacement cost as % of original capital cost")
    
    # ========================================================================
    # FINANCIAL CONFIGURATION
    # ========================================================================
    with st.expander("💰 FINANCIAL PARAMETERS"):
        discount_rate = st.number_input("Discount Rate (%)", value=8.0, min_value=0.0, max_value=20.0, step=0.5)
        inflation_rate = st.number_input("Inflation Rate (%)", value=2.0, min_value=0.0, max_value=10.0, step=0.5)
        project_lifetime = st.number_input("Project Lifetime (years)", value=25, min_value=1, max_value=50, step=1)
    # ========================================================================
    # OPTIMIZATION CONFIGURATION
    # ========================================================================
    with st.expander("🎯 OPTIMIZATION SETTINGS"):
    st.subheader("Reliability Target")
    target_unmet_percent = st.number_input(
        "Target Unmet Load (%)", 
        value=0.0, 
        min_value=0.0, 
        max_value=5.0, 
        step=0.1,
        key="target_unmet",
        help="Maximum acceptable unmet load. 0% = 100% reliable system"
    )
    
    if target_unmet_percent == 0.0:
        st.info("🎯 Target: 100% reliable system (no unmet load)")
    else:
        st.info(f"🎯 Target: ≥{100-target_unmet_percent:.1f}% reliability")
    # ========================================================================
    # FILE UPLOADS
    # ========================================================================
    st.header("📁 Upload Profiles")
    
    load_file = st.file_uploader("Load Profile (8760 hours)", type=['csv', 'xlsx'], key="load_file")
    pv_file = st.file_uploader("Solar PV Profile (1 kW)", type=['csv', 'xlsx'], key="pv_file")
    wind_file = st.file_uploader("Wind Profile (1 kW)", type=['csv', 'xlsx'], key="wind_file")
    
    st.info("💡 Tip: Upload Excel or CSV files with 'Hour' and 'Load_kW' (or Output_kW) columns")

# Main area - Tabs
tab1, tab2, tab3, tab4 = st.tabs(["🏠 Home", "⚙️ Optimize", "📊 Results", "📈 Analysis"])

# ============================================================================
# TAB 1: HOME
# ============================================================================
with tab1:
    st.header("Welcome to the Renewable Energy Optimization Tool")
    
    st.markdown("""
    ### 🎯 Purpose
    This tool optimizes hybrid renewable energy systems to minimize costs while meeting reliability targets.
    
    ### 🔧 Features
    - **Multi-source optimization:** Solar PV + Wind + Hydro + Battery Storage
    - **Grid search algorithm:** Tests all combinations to find global optimum
    - **Firm capacity analysis:** Identifies systems with 100% reliability
    - **NPC calculation:** Industry-standard financial analysis
    - **Hourly dispatch:** Complete 8760-hour simulation
    - **Interactive visualization:** Charts and graphs for easy understanding
    
    ### 📋 How to Use
    1. **Configure components** in the sidebar (left)
    2. **Upload profiles** (load, PV, wind) or use defaults
    3. **Run optimization** in the "Optimize" tab
    4. **View results** in the "Results" tab
    5. **Analyze performance** in the "Analysis" tab
    6. **Export results** to Excel or PDF
    
    ### 🚀 Quick Start
    1. Adjust capacity ranges in sidebar
    2. Go to "Optimize" tab
    3. Click "▶️ RUN OPTIMIZATION"
    4. View optimal system configuration
    """)
    
    # Calculate search space
    pv_options = int((pv_max - pv_min) / pv_step) + 1 if pv_step > 0 else 1
    wind_options = int((wind_max - wind_min) / wind_step) + 1 if wind_step > 0 else 1
    hydro_options = int((hydro_max - hydro_min) / hydro_step) + 1 if hydro_step > 0 else 1
    bess_options = int((bess_max - bess_min) / bess_step) + 1 if bess_step > 0 else 1
    total_combinations = pv_options * wind_options * hydro_options * bess_options
    
    st.subheader("📊 Current Configuration Summary")
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        st.metric("PV Options", f"{pv_options}", f"{pv_min}-{pv_max} MW")
    with col2:
        st.metric("Wind Options", f"{wind_options}", f"{wind_min}-{wind_max} MW")
    with col3:
        st.metric("Hydro Options", f"{hydro_options}", f"{hydro_min}-{hydro_max} MW")
    with col4:
        st.metric("BESS Options", f"{bess_options}", f"{bess_min}-{bess_max} MW")
    
    st.info(f"**Total Search Space:** {total_combinations:,} combinations")
    
    if total_combinations > 10000:
        st.warning("⚠️ Large search space! Consider increasing step sizes for faster optimization.")

# ============================================================================
# TAB 2: OPTIMIZE
# ============================================================================
with tab2:
    st.header("⚙️ Run Optimization")
    
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        st.metric("Search Space", f"{total_combinations:,}", "combinations")
    with col2:
        est_time = max(1, total_combinations * 0.05 / 60)  # Rough estimate
        st.metric("Est. Time", f"{est_time:.1f} min", "approximate")
    with col3:
        st.metric("Components", "4", "PV+Wind+Hydro+BESS")
    with col4:
        st.metric("Target Unmet", "0%", "100% reliable")
    
    st.markdown("---")
    
    # Input validation
    validation_passed = True
    validation_messages = []
    
    if pv_max <= pv_min:
        validation_passed = False
        validation_messages.append("❌ PV Max must be greater than PV Min")
    if wind_max < wind_min:
        validation_passed = False
        validation_messages.append("❌ Wind Max must be greater than or equal to Wind Min")
    if hydro_max < hydro_min:
        validation_passed = False
        validation_messages.append("❌ Hydro Max must be greater than or equal to Hydro Min")
    if bess_max <= bess_min:
        validation_passed = False
        validation_messages.append("❌ BESS Max must be greater than BESS Min")
    if not load_file:
        validation_passed = False
        validation_messages.append("❌ Please upload Load Profile")
    if not pv_file:
        validation_passed = False
        validation_messages.append("❌ Please upload Solar PV Profile")
    if not wind_file:
        validation_passed = False
        validation_messages.append("❌ Please upload Wind Profile")
    
    if validation_messages:
        for msg in validation_messages:
            st.error(msg)
    else:
        st.success("✅ All inputs validated. Ready to optimize!")
    
    # Optimization button
    if st.button("▶️ RUN OPTIMIZATION", type="primary", disabled=not validation_passed, use_container_width=True):
        
        if not OPTIMIZATION_AVAILABLE:
            st.error("❌ Optimization code not available. Please check GitHub repository.")
        else:
            try:
                st.subheader("🔄 Optimization in Progress...")
                progress_bar = st.progress(0)
                status_text = st.empty()
                
                # Read uploaded profiles
                status_text.text("📊 Reading uploaded profiles...")
                progress_bar.progress(10)
                
                # Load profile
                if load_file.name.endswith('.csv'):
                    load_df = pd.read_csv(load_file)
                else:
                    load_df = pd.read_excel(load_file)
                
                # PV profile
                if pv_file.name.endswith('.csv'):
                    pv_df = pd.read_csv(pv_file)
                else:
                    pv_df = pd.read_excel(pv_file)
                
                # Wind profile
                if wind_file.name.endswith('.csv'):
                    wind_df = pd.read_csv(wind_file)
                else:
                    wind_df = pd.read_excel(wind_file)
                
                # Hydro profile (optional)
                hydro_df = None
                if 'hydro_file' in locals() and hydro_file is not None:
                    if hydro_file.name.endswith('.csv'):
                        hydro_df = pd.read_csv(hydro_file)
                    else:
                        hydro_df = pd.read_excel(hydro_file)
                
                # Build Excel from parameters
                status_text.text("🔨 Building input file from your parameters...")
                progress_bar.progress(15)
                
                excel_bytes = build_input_excel_from_streamlit(
                    8760, 
                    target_unmet_percent,  # <-- Use sidebar value, not hardcoded 0
                    discount_rate, 
                    inflation_rate, 
                    project_lifetime,
                    pv_min, pv_max, pv_step, pv_capex, pv_opex, pv_lifetime, pv_lcoe,
                    wind_min, wind_max, wind_step, wind_capex, wind_opex, wind_lifetime, wind_lcoe,
                    hydro_min, hydro_max, hydro_step, 
                    hydro_hours_per_day,  # <-- Pass this instead of window_min/max
                    hydro_capex, hydro_opex, hydro_lifetime, hydro_lcoe,
                    bess_min, bess_max, bess_step, bess_duration, bess_min_soc, bess_max_soc,
                    bess_charge_eff, bess_discharge_eff, bess_power_capex, bess_energy_capex,
                    bess_opex, bess_lifetime, bess_replacement_cost,
                    load_df, pv_df, wind_df, hydro_df

                    )

                
                # Save temp file
                temp_file = "temp_input_generated.xlsx"
                with open(temp_file, "wb") as f:
                    f.write(excel_bytes.getvalue())
                
                # Run optimization
                status_text.text("⚙️ Loading optimization engine...")
                progress_bar.progress(25)
                
                import optimize_gridsearch_hydro_static_HOMERNPCFIXED_COMB as opt_module
                opt_module.INPUT_FILE = temp_file
                
                status_text.text("📖 Reading configuration...")
                progress_bar.progress(30)
                
                result = opt_module.read_inputs()
                if len(result) == 9:
                    config, grid_config, solar, wind, hydro, bess, load_profile, pvsyst_profile, wind_profile = result
                    hydro_profile_opt = None
                else:
                    config, grid_config, solar, wind, hydro, bess, load_profile, pvsyst_profile, wind_profile, hydro_profile_opt = result
                
                status_text.text("🔍 Running grid search optimization... (this may take several minutes)")
                progress_bar.progress(35)
                
                results_df = opt_module.grid_search_optimize_hydro(
                    config, grid_config, solar, wind, hydro, bess,
                    load_profile, pvsyst_profile, wind_profile, hydro_profile_opt,
                    lcoe_tables=None
                )
                
                progress_bar.progress(85)
                status_text.text("🎯 Finding optimal solution...")
                
                optimal = opt_module.find_optimal_solution(results_df)
                
                progress_bar.progress(100)
                status_text.text("✅ Complete!")
                
                # Clean up
                if os.path.exists(temp_file):
                    os.remove(temp_file)
                
                if optimal is not None:
                    st.session_state.results = {
                        'pv_capacity': optimal['PV_kW'] / 1000,
                        'wind_capacity': optimal['Wind_kW'] / 1000,
                        'hydro_capacity': optimal['Hydro_kW'] / 1000,
                        'hydro_window_start': optimal['Hydro_Window_Start'],
                        'hydro_window_end': optimal['Hydro_Window_End'],
                        'bess_power': optimal['BESS_Power_kW'] / 1000,
                        'bess_energy': optimal['BESS_Capacity_kWh'] / 1000,
                        'npc': optimal['NPC_$'],
                        'lcoe': optimal['LCOE_$/MWh'],
                        'unmet_pct': optimal['Unmet_%'],
                        'firm_capacity': optimal.get('Firm_Capacity_MW', 0),
                        're_penetration': optimal.get('RE_Penetration_%', 100),
                        'pv_npc': optimal.get('PV_NPC_$', 0),
                        'wind_npc': optimal.get('Wind_NPC_$', 0),
                        'hydro_npc': optimal.get('Hydro_NPC_$', 0),
                        'bess_npc': optimal.get('BESS_NPC_$', 0),
                        'results_df': results_df
                    }
                    st.session_state.optimization_complete = True
                    
                    st.success("✅ Optimization Complete!")
                    st.balloons()
                    st.info("👉 Go to **Results** tab to view optimal configuration")
                else:
                    st.error("❌ No optimal solution found. Check your parameters.")
                    
            except Exception as e:
                st.error(f"❌ Error during optimization: {str(e)}")
                st.exception(e)
                
                # Clean up
                if os.path.exists("temp_input_generated.xlsx"):
                    os.remove("temp_input_generated.xlsx")

# ============================================================================
# TAB 3: RESULTS
# ============================================================================
with tab3:
    if not st.session_state.optimization_complete:
        st.info("ℹ️ No results yet. Run optimization first in the 'Optimize' tab.")
    else:
        st.header("📊 Optimization Results")
        
        results = st.session_state.results
        
        # Key metrics
        st.subheader("🎯 Key Performance Indicators")
        col1, col2, col3, col4 = st.columns(4)
        
        with col1:
            st.metric(
                "Total NPC",
                f"${results['npc']/1e6:.2f}M",
                help="Net Present Cost over project lifetime"
            )
        
        with col2:
            st.metric(
                "System LCOE",
                f"${results['lcoe']:.2f}/MWh",
                help="Levelized Cost of Energy"
            )
        
        with col3:
            st.metric(
                "Firm Capacity",
                f"{results['firm_capacity']:.3f} MW",
                delta="100% Reliable" if results['unmet_pct'] == 0 else f"{results['unmet_pct']:.2f}% Unmet",
                delta_color="normal" if results['unmet_pct'] == 0 else "inverse",
                help="Maximum load system can serve with 100% reliability"
            )
        
        with col4:
            st.metric(
                "RE Penetration",
                f"{results['re_penetration']:.1f}%",
                help="Renewable energy as % of total generation"
            )
        
        st.markdown("---")
        
        # Optimal configuration
        st.subheader("⚡ Optimal System Configuration")
        
        config_col1, config_col2 = st.columns(2)
        
        with config_col1:
            st.markdown("#### Generation & Storage")
            config_data = {
                'Component': ['☀️ Solar PV', '💨 Wind', '💧 Hydro', '🔋 BESS Power', '🔋 BESS Energy'],
                'Capacity': [
                    results['pv_capacity'],
                    results['wind_capacity'],
                    results['hydro_capacity'],
                    results['bess_power'],
                    results['bess_energy']
                ],
                'Unit': ['MW', 'MW', 'MW', 'MW', 'MWh']
            }
            st.dataframe(
                pd.DataFrame(config_data),
                use_container_width=True,
                hide_index=True
            )
            
            st.info(f"💧 Hydro Operating Window: {results['hydro_window_start']:02d}:00 - {results['hydro_window_end']:02d}:00")
        
        with config_col2:
            st.markdown("#### Cost Breakdown")
            cost_data = {
                'Component': ['☀️ Solar PV', '💨 Wind', '💧 Hydro', '🔋 BESS'],
                'NPC ($M)': [
                    results['pv_npc']/1e6,
                    results['wind_npc']/1e6,
                    results['hydro_npc']/1e6,
                    results['bess_npc']/1e6
                ],
                '% of Total': [
                    results['pv_npc']/results['npc']*100,
                    results['wind_npc']/results['npc']*100,
                    results['hydro_npc']/results['npc']*100,
                    results['bess_npc']/results['npc']*100
                ]
            }
            st.dataframe(
                pd.DataFrame(cost_data).style.format({
                    'NPC ($M)': '{:.2f}',
                    '% of Total': '{:.1f}%'
                }),
                use_container_width=True,
                hide_index=True
            )
        
        st.markdown("---")
        
        # Charts
        st.subheader("📊 Visual Analysis")
        
        chart_col1, chart_col2 = st.columns(2)
        
        with chart_col1:
            # Cost breakdown pie chart
            fig_cost = go.Figure(data=[go.Pie(
                labels=['Solar PV', 'Wind', 'Hydro', 'BESS'],
                values=[
                    results['pv_npc'],
                    results['wind_npc'],
                    results['hydro_npc'],
                    results['bess_npc']
                ],
                marker=dict(colors=['#FDB462', '#80B1D3', '#8DD3C7', '#FB8072']),
                textinfo='label+percent',
                hovertemplate='%{label}<br>$%{value:,.0f}<br>%{percent}<extra></extra>'
            )])
            fig_cost.update_layout(
                title="Cost Distribution by Component",
                showlegend=True,
                height=400
            )
            st.plotly_chart(fig_cost, use_container_width=True)
        
        with chart_col2:
            # Capacity comparison bar chart
            fig_capacity = go.Figure(data=[
                go.Bar(
                    x=['Solar PV', 'Wind', 'Hydro', 'BESS Power'],
                    y=[
                        results['pv_capacity'],
                        results['wind_capacity'],
                        results['hydro_capacity'],
                        results['bess_power']
                    ],
                    marker=dict(color=['#FDB462', '#80B1D3', '#8DD3C7', '#FB8072']),
                    text=[
                        f"{results['pv_capacity']:.1f} MW",
                        f"{results['wind_capacity']:.1f} MW",
                        f"{results['hydro_capacity']:.1f} MW",
                        f"{results['bess_power']:.1f} MW"
                    ],
                    textposition='outside'
                )
            ])
            fig_capacity.update_layout(
                title="Installed Capacity by Component",
                yaxis_title="Capacity (MW)",
                showlegend=False,
                height=400
            )
            st.plotly_chart(fig_capacity, use_container_width=True)
        
        st.markdown("---")
        
        # Export buttons
        st.subheader("📥 Export Results")
        col1, col2, col3 = st.columns(3)
        
        with col1:
            # Create Excel export
            excel_buffer = io.BytesIO()
            with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
                pd.DataFrame([results]).to_excel(writer, sheet_name='Summary', index=False)
            excel_buffer.seek(0)
            
            st.download_button(
                label="📊 Download Excel Report",
                data=excel_buffer,
                file_name=f"optimization_results_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
        
        with col2:
            st.download_button(
                label="📄 Download PDF Report",
                data="PDF report generation coming soon",
                file_name=f"report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf",
                mime="application/pdf",
                use_container_width=True,
                disabled=True
            )
        
        with col3:
            st.download_button(
                label="📈 Download Charts",
                data="Chart export coming soon",
                file_name=f"charts_{datetime.now().strftime('%Y%m%d_%H%M%S')}.zip",
                mime="application/zip",
                use_container_width=True,
                disabled=True
            )

# ============================================================================
# TAB 4: ANALYSIS
# ============================================================================
with tab4:
    if not st.session_state.optimization_complete:
        st.info("ℹ️ No results yet. Run optimization first in the 'Optimize' tab.")
    else:
        st.header("📈 Detailed Analysis")
        
        st.subheader("🔋 Hourly Dispatch Simulation")
        st.info("📊 8760-hour detailed dispatch visualization coming soon")
        
        st.subheader("💰 Financial Deep Dive")
        st.info("📊 Component-wise cost breakdown analysis coming soon")
        
        st.subheader("⚡ Performance Metrics")
        st.info("📊 Capacity factors and utilization analysis coming soon")

# Footer
st.markdown("---")
st.markdown("""
<div style='text-align: center; color: #666;'>
    <p>Renewable Energy Optimization Tool v1.0</p>
    <p>Developed for Hybrid System Design | PV + Wind + Hydro + BESS</p>
</div>

""", unsafe_allow_html=True)








