"""
RENEWABLE ENERGY OPTIMIZATION TOOL - FINAL COMPLETE VERSION
============================================================
Streamlit web application for hybrid renewable energy system optimization
PV + Wind + Hydro + BESS with HOMER-style NPC calculation

This version uses your exact Anaconda optimization code for guaranteed accuracy.
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

# Import optimization code - using your exact file
try:
    import optimize_gridsearch_hydro_static_STREAMLITCHECK as opt_module
    OPTIMIZATION_AVAILABLE = True
except ImportError:
    OPTIMIZATION_AVAILABLE = False
    st.error("❌ Optimization module not found. Please ensure optimize_gridsearch_hydro_static_STREAMLITCHECK.py is in the repository.")


# ==============================================================================
# PAGE CONFIGURATION
# ==============================================================================

st.set_page_config(
    page_title="RE Optimization Tool",
    page_icon="🌞",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS
st.markdown("""
<style>
    .main-header {
        font-size: 2.5rem;
        font-weight: bold;
        color: #1f77b4;
        text-align: center;
        margin-bottom: 2rem;
    }
</style>
""", unsafe_allow_html=True)

# Title
st.markdown('<p class="main-header">🌞 Renewable Energy Optimization Tool</p>', unsafe_allow_html=True)
st.markdown("**Hybrid System Designer: PV + Wind + Hydro + Battery Storage**")
st.markdown("---")

# Initialize session state
if 'optimization_complete' not in st.session_state:
    st.session_state.optimization_complete = False
if 'results' not in st.session_state:
    st.session_state.results = None


# ==============================================================================
# SIDEBAR - INPUT PARAMETERS
# ==============================================================================

with st.sidebar:
    st.header("⚙️ System Configuration")
    
    # ========================================================================
    # SOLAR PV
    # ========================================================================
    with st.expander("☀️ SOLAR PV", expanded=True):
        st.subheader("Capacity Range")
        col1, col2 = st.columns(2)
        with col1:
            pv_min = st.number_input("Min (MW)", value=1.0, min_value=0.0, step=0.5, key="pv_min")
        with col2:
            pv_max = st.number_input("Max (MW)", value=5.0, min_value=0.0, step=0.5, key="pv_max")
        pv_step = st.number_input("Step (MW)", value=0.5, min_value=0.1, step=0.1, key="pv_step")
        
        st.subheader("Financial Parameters")
        col1, col2 = st.columns(2)
        with col1:
            pv_capex = st.number_input("CapEx ($/kW)", value=1000, step=10, key="pv_capex")
            pv_opex = st.number_input("OpEx ($/kW/yr)", value=10, step=1, key="pv_opex")
        with col2:
            pv_lifetime = st.number_input("Lifetime (years)", value=25, step=1, key="pv_life")
            pv_lcoe = st.number_input("LCOE ($/MWh)", value=35.0, step=1.0, key="pv_lcoe")
    
    # ========================================================================
    # WIND
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
            wind_capex = st.number_input("CapEx ($/kW)", value=1200, step=10, key="wind_capex")
            wind_opex = st.number_input("OpEx ($/kW/yr)", value=15, step=1, key="wind_opex")
        with col2:
            wind_lifetime = st.number_input("Lifetime (years)", value=20, step=1, key="wind_life")
            wind_lcoe = st.number_input("LCOE ($/MWh)", value=45.0, step=1.0, key="wind_lcoe")
    
    # ========================================================================
    # HYDRO
    # ========================================================================
    with st.expander("💧 HYDRO"):
        st.subheader("Capacity Range")
        col1, col2 = st.columns(2)
        with col1:
            hydro_min = st.number_input("Min (MW)", value=0.0, min_value=0.0, step=0.5, key="hydro_min")
        with col2:
            hydro_max = st.number_input("Max (MW)", value=2.0, min_value=0.0, step=0.5, key="hydro_max")
        hydro_step = st.number_input("Step (MW)", value=0.5, min_value=0.1, step=0.1, key="hydro_step")
        
        st.subheader("Operating Configuration")
        hydro_hours_per_day = st.number_input(
            "Operating Hours (hours/day)", 
            value=6, 
            min_value=1, 
            max_value=24, 
            step=1,
            key="hydro_hours",
            help="System will auto-optimize the time window"
        )
        st.info(f"💡 System will test all possible {hydro_hours_per_day}-hour windows")
        
        st.subheader("Financial Parameters")
        col1, col2 = st.columns(2)
        with col1:
            hydro_capex = st.number_input("CapEx ($/kW)", value=2000, step=10, key="hydro_capex")
            hydro_opex = st.number_input("OpEx ($/kW/yr)", value=20, step=1, key="hydro_opex")
        with col2:
            hydro_lifetime = st.number_input("Lifetime (years)", value=50, step=1, key="hydro_life")
            hydro_lcoe = st.number_input("LCOE ($/MWh)", value=40.0, step=1.0, key="hydro_lcoe")
    
    # ========================================================================
    # BESS
    # ========================================================================
    with st.expander("🔋 BATTERY STORAGE"):
        st.subheader("Power Range")
        col1, col2 = st.columns(2)
        with col1:
            bess_min = st.number_input("Min (MW)", value=5.0, min_value=0.0, step=1.0, key="bess_min")
        with col2:
            bess_max = st.number_input("Max (MW)", value=20.0, min_value=0.0, step=1.0, key="bess_max")
        bess_step = st.number_input("Step (MW)", value=1.0, min_value=0.1, step=0.1, key="bess_step")
        
        st.subheader("Storage Parameters")
        col1, col2 = st.columns(2)
        with col1:
            bess_duration = st.number_input("Duration (hours)", value=4.0, min_value=0.5, step=0.5, key="bess_dur")
            bess_min_soc = st.number_input("Min SOC (%)", value=20, min_value=0, max_value=100, key="bess_min_soc")
            bess_charge_eff = st.number_input("Charging Eff (%)", value=95, min_value=50, max_value=100, key="bess_charge_eff")
        with col2:
            bess_lifetime = st.number_input("Lifetime (years)", value=15, step=1, key="bess_life")
            bess_max_soc = st.number_input("Max SOC (%)", value=100, min_value=0, max_value=100, key="bess_max_soc")
            bess_discharge_eff = st.number_input("Discharging Eff (%)", value=95, min_value=50, max_value=100, key="bess_discharge_eff")
        
        st.subheader("Financial Parameters")
        col1, col2 = st.columns(2)
        with col1:
            bess_power_capex = st.number_input("Power CapEx ($/kW)", value=300, step=10, key="bess_power_capex")
            bess_energy_capex = st.number_input("Energy CapEx ($/kWh)", value=200, step=10, key="bess_energy_capex")
        with col2:
            bess_opex = st.number_input("OpEx ($/kWh/yr)", value=2, step=1, key="bess_opex")
            bess_replacement_cost = st.number_input("Replacement (%)", value=80, min_value=0, max_value=100, key="bess_replacement")
    
    # ========================================================================
    # FINANCIAL PARAMETERS
    # ========================================================================
    with st.expander("💰 FINANCIAL PARAMETERS"):
        discount_rate = st.number_input("Discount Rate (%)", value=8.0, min_value=0.0, max_value=20.0, step=0.5)
        inflation_rate = st.number_input("Inflation Rate (%)", value=2.0, min_value=0.0, max_value=10.0, step=0.5)
        project_lifetime = st.number_input("Project Lifetime (years)", value=25, min_value=1, max_value=50, step=1)
    
    # ========================================================================
    # OPTIMIZATION SETTINGS
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
            help="0% = 100% reliable system"
        )
        if target_unmet_percent == 0.0:
            st.info("🎯 Target: 100% reliable system")
        else:
            st.info(f"🎯 Target: ≥{100-target_unmet_percent:.1f}% reliability")
    
    # ========================================================================
    # FILE UPLOADS
    # ========================================================================
    st.header("📁 Upload Profiles")
    st.info("💡 Upload 8760-hour profile data files")
    
    load_file = st.file_uploader("Load Profile", type=['csv', 'xlsx'], key="load_file")
    if load_file:
        st.success(f"✅ {load_file.name}")
    
    pv_file = st.file_uploader("PV Profile (1 kW)", type=['csv', 'xlsx'], key="pv_file")
    if pv_file:
        st.success(f"✅ {pv_file.name}")
    
    wind_file = st.file_uploader("Wind Profile (1 kW)", type=['csv', 'xlsx'], key="wind_file")
    if wind_file:
        st.success(f"✅ {wind_file.name}")
    
    hydro_file = st.file_uploader("Hydro Profile (Optional)", type=['csv', 'xlsx'], key="hydro_file")
    if hydro_file:
        st.success(f"✅ {hydro_file.name}")


# ==============================================================================
# MAIN AREA - TABS
# ==============================================================================

tab1, tab2, tab3 = st.tabs(["🏠 Home", "⚙️ Optimize", "📊 Results"])

# Calculate search space
pv_options = int((pv_max - pv_min) / pv_step) + 1 if pv_step > 0 else 1
wind_options = int((wind_max - wind_min) / wind_step) + 1 if wind_step > 0 else 1
hydro_options = int((hydro_max - hydro_min) / hydro_step) + 1 if hydro_step > 0 else 1
bess_options = int((bess_max - bess_min) / bess_step) + 1 if bess_step > 0 else 1
total_combinations = pv_options * wind_options * hydro_options * bess_options


# ==============================================================================
# TAB 1: HOME
# ==============================================================================
with tab1:
    st.header("Welcome to the Renewable Energy Optimization Tool")
    
    st.markdown("""
    ### 🎯 Purpose
    Optimize hybrid renewable energy systems to minimize costs while meeting reliability targets.
    
    ### 🔧 Features
    - Multi-source optimization: PV + Wind + Hydro + BESS
    - Grid search algorithm with HOMER-style NPC calculation
    - Firm capacity analysis and hourly dispatch
    - Interactive visualization and Excel export
    """)
    
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
        st.warning("⚠️ Large search space! Consider increasing step sizes.")


# ==============================================================================
# TAB 2: OPTIMIZE
# ==============================================================================
with tab2:
    st.header("⚙️ Run Optimization")
    
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        st.metric("Search Space", f"{total_combinations:,}")
    with col2:
        est_time = max(1, total_combinations * 0.05 / 60)
        st.metric("Est. Time", f"{est_time:.1f} min")
    with col3:
        st.metric("Components", "4")
    with col4:
        st.metric("Target Unmet", f"{target_unmet_percent}%")
    
    st.markdown("---")
    
    # Validation
    validation_passed = True
    validation_messages = []
    
    if not load_file:
        validation_passed = False
        validation_messages.append("❌ Please upload Load Profile")
    if not pv_file:
        validation_passed = False
        validation_messages.append("❌ Please upload PV Profile")
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
            st.error("❌ Optimization code not available")
        else:
            try:
                st.subheader("🔄 Optimization in Progress...")
                progress_bar = st.progress(0)
                status_text = st.empty()
                
                # Read profiles
                status_text.text("📊 Reading uploaded profiles...")
                progress_bar.progress(10)
                
                if load_file.name.endswith('.csv'):
                    load_df = pd.read_csv(load_file)
                else:
                    load_df = pd.read_excel(load_file)
                
                if pv_file.name.endswith('.csv'):
                    pv_df = pd.read_csv(pv_file)
                else:
                    pv_df = pd.read_excel(pv_file)
                
                if wind_file.name.endswith('.csv'):
                    wind_df = pd.read_csv(wind_file)
                else:
                    wind_df = pd.read_excel(wind_file)
                
                hydro_df = None
                if hydro_file is not None:
                    if hydro_file.name.endswith('.csv'):
                        hydro_df = pd.read_csv(hydro_file)
                    else:
                        hydro_df = pd.read_excel(hydro_file)
                
                # Build Excel (keeping this for compatibility with your optimization code)
                status_text.text("🔨 Building input file...")
                progress_bar.progress(15)
                
                # Convert MW to kW
                pv_min_kw = pv_min * 1000
                pv_max_kw = pv_max * 1000
                pv_step_kw = pv_step * 1000
                wind_min_kw = wind_min * 1000
                wind_max_kw = wind_max * 1000
                wind_step_kw = wind_step * 1000
                hydro_min_kw = hydro_min * 1000
                hydro_max_kw = hydro_max * 1000
                hydro_step_kw = hydro_step * 1000
                bess_min_kw = bess_min * 1000
                bess_max_kw = bess_max * 1000
                bess_step_kw = bess_step * 1000
                
                # Build Excel input matching Anaconda format
                output = BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    # Configuration
                    pd.DataFrame({
                        'Parameter': ['Simulation Hours', 'Target Unmet Load (%)', 'Optimization Method',
                                     'Discount Rate (%)', 'Inflation Rate (%)', 'Project Lifetime (years)', 'Use Dynamic LCOE'],
                        'Value': [8760, target_unmet_percent, 'GRID_SEARCH', discount_rate, inflation_rate, project_lifetime, 'NO']
                    }).to_excel(writer, sheet_name='Configuration', index=False)
                    
                    # Grid_Search_Config
                    pd.DataFrame({
                        'Parameter': ['Enable Grid Search', 'PV Search Start', 'PV Search End', 'PV Search Step',
                                     'Wind Search Start', 'Wind Search End', 'Wind Search Step',
                                     'Hydro Search Start', 'Hydro Search End', 'Hydro Search Step',
                                     'BESS Search Start', 'BESS Search End', 'BESS Search Step',
                                     'Max Combinations', 'Optimization Objective', 'Show Top N Solutions'],
                        'Value': ['YES', pv_min_kw, pv_max_kw, pv_step_kw, wind_min_kw, wind_max_kw, wind_step_kw,
                                 hydro_min_kw, hydro_max_kw, hydro_step_kw, bess_min_kw, bess_max_kw, bess_step_kw,
                                 100000000, 'NPC', 5]
                    }).to_excel(writer, sheet_name='Grid_Search_Config', index=False)
                    
                    # Solar_PV
                    pd.DataFrame({
                        'Parameter': ['LCOE', 'PVsyst Baseline', 'Capex', 'O&M Cost', 'Lifetime'],
                        'Value': [pv_lcoe, 1.0, pv_capex, pv_opex, pv_lifetime]
                    }).to_excel(writer, sheet_name='Solar_PV', index=False)
                    
                    # Wind
                    pd.DataFrame({
                        'Parameter': ['Include Wind?', 'LCOE', 'Capex', 'O&M Cost', 'Lifetime'],
                        'Value': ['YES' if wind_max > 0 else 'NO', wind_lcoe, wind_capex, wind_opex, wind_lifetime]
                    }).to_excel(writer, sheet_name='Wind', index=False)
                    
                    # Hydro
                    pd.DataFrame({
                        'Parameter': ['Include Hydro?', 'LCOE', 'Capex', 'O&M Cost', 'Lifetime', 'Operating Hours'],
                        'Value': ['YES' if hydro_max > 0 else 'NO', hydro_lcoe, hydro_capex, hydro_opex, hydro_lifetime, hydro_hours_per_day]
                    }).to_excel(writer, sheet_name='Hydro', index=False)
                    
                    # BESS
                    pd.DataFrame({
                        'Parameter': ['Duration', 'LCOS', 'Charge Efficiency', 'Discharge Efficiency', 'Min SOC', 'Max SOC',
                                     'Power Capex', 'Energy Capex', 'O&M Cost', 'Lifetime'],
                        'Value': [bess_duration, 0, bess_charge_eff, bess_discharge_eff, bess_min_soc, bess_max_soc,
                                 bess_power_capex, bess_energy_capex, bess_opex, bess_lifetime]
                    }).to_excel(writer, sheet_name='BESS', index=False)
                    
                    # Profiles
                    load_df.to_excel(writer, sheet_name='Load_Profile', index=False)
                    pv_df.to_excel(writer, sheet_name='PVsyst_Profile', index=False)
                    wind_df.to_excel(writer, sheet_name='Wind_Profile', index=False)
                    
                    if hydro_df is not None:
                        hydro_df.to_excel(writer, sheet_name='Hydro_Profile', index=False)
                    else:
                        pd.DataFrame({'Hour': range(8760), 'Output_kW': [1.0] * 8760}).to_excel(writer, sheet_name='Hydro_Profile', index=False)
                
                output.seek(0)
                
                # Save temp file
                status_text.text("💾 Saving temporary file...")
                progress_bar.progress(20)
                
                import tempfile
                temp_dir = tempfile.gettempdir()
                temp_file = os.path.join(temp_dir, "temp_input_generated.xlsx")
                
                with open(temp_file, "wb") as f:
                    f.write(output.getvalue())
                
                if not os.path.exists(temp_file):
                    raise FileNotFoundError(f"Failed to create {temp_file}")
                
                # Run optimization using your exact code
                status_text.text("⚙️ Loading optimization engine...")
                progress_bar.progress(25)
                
                opt_module.INPUT_FILE = temp_file
                
                status_text.text("📖 Reading configuration...")
                progress_bar.progress(30)
                
                result = opt_module.read_inputs()
                if len(result) == 9:
                    config, grid_config, solar, wind, hydro, bess, load_profile, pvsyst_profile, wind_profile = result
                    hydro_profile_opt = None
                else:
                    config, grid_config, solar, wind, hydro, bess, load_profile, pvsyst_profile, wind_profile, hydro_profile_opt = result
                
                status_text.text("🔍 Running optimization... (may take several minutes)")
                progress_bar.progress(35)
                
                results_df = opt_module.grid_search_optimize_hydro(
                    config, grid_config, solar, wind, hydro, bess,
                    load_profile, pvsyst_profile, wind_profile, hydro_profile_opt
                )
                
                progress_bar.progress(85)
                status_text.text("🎯 Finding optimal solution...")
                
                optimal = opt_module.find_optimal_solution(results_df)
                
                progress_bar.progress(100)
                status_text.text("✅ Complete!")
                # After optimal = opt_module.find_optimal_solution(results_df)
                st.write("**🔍 DIAGNOSTIC - First Row from Results:**")
                first_row = results_df.iloc[0]
                st.write(f"PV_OM_$ from results_df: ${first_row.get('PV_OM_$', 0):,.0f}")
                st.write(f"Expected: ~$775,651 (Present Value)")
                if first_row.get('PV_OM_$', 0) < 100000:
                    st.error("❌ WRONG! Getting annual cost instead of Present Value!")
                    st.write("This means the optimization code is using wrong calculations.")
                # Clean up
                if os.path.exists(temp_file):
                    os.remove(temp_file)
                
                if optimal is not None:
                    # Store results
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
                        'pv_npc': optimal.get('PV_NPC_$', 0),
                        'wind_npc': optimal.get('Wind_NPC_$', 0),
                        'hydro_npc': optimal.get('Hydro_NPC_$', 0),
                        'bess_npc': optimal.get('BESS_NPC_$', 0),
                        'results_df': results_df,
                        'optimal_row': optimal.to_dict()
                    }
                    
                    st.session_state.optimization_complete = True
                    
                    st.success("✅ Optimization Complete!")
                    st.balloons()
                    st.info("👉 Go to **Results** tab")
                else:
                    st.error("❌ No optimal solution found")
                    
            except Exception as e:
                st.error(f"❌ Error: {str(e)}")
                st.exception(e)
                
                try:
                    if 'temp_file' in locals() and os.path.exists(temp_file):
                        os.remove(temp_file)
                except:
                    pass


# ==============================================================================
# TAB 3: RESULTS
# ==============================================================================
with tab3:
    if not st.session_state.optimization_complete:
        st.info("ℹ️ No results yet. Run optimization first.")
    else:
        st.header("📊 Optimization Results")
        
        results = st.session_state.results
        
        # Key metrics
        col1, col2, col3, col4 = st.columns(4)
        with col1:
            st.metric("Total NPC", f"${results['npc']/1e6:.2f}M")
        with col2:
            st.metric("System LCOE", f"${results['lcoe']:.2f}/MWh")
        with col3:
            st.metric("PV Capacity", f"{results['pv_capacity']:.1f} MW")
        with col4:
            st.metric("Unmet Load", f"{results['unmet_pct']:.2f}%")
        
        st.markdown("---")
        
        # Configuration
        st.subheader("⚡ Optimal Configuration")
        config_data = {
            'Component': ['Solar PV', 'Wind', 'Hydro', 'BESS Power', 'BESS Energy'],
            'Capacity': [
                f"{results['pv_capacity']:.1f} MW",
                f"{results['wind_capacity']:.1f} MW",
                f"{results['hydro_capacity']:.1f} MW",
                f"{results['bess_power']:.1f} MW",
                f"{results['bess_energy']:.1f} MWh"
            ],
            'NPC ($M)': [
                f"${results['pv_npc']/1e6:.2f}",
                f"${results['wind_npc']/1e6:.2f}",
                f"${results['hydro_npc']/1e6:.2f}",
                f"${results['bess_npc']/1e6:.2f}",
                "(included in BESS Power)"
            ]
        }
        st.dataframe(pd.DataFrame(config_data), use_container_width=True, hide_index=True)
        
        st.info(f"💧 Hydro Window: {int(results['hydro_window_start']):02d}:00 - {int(results['hydro_window_end']):02d}:00")
        
        # Download Results
        st.markdown("---")
        st.subheader("📥 Download Results")
        
        # Export to Excel with all results
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            # Summary
            pd.DataFrame({
                'Parameter': ['Total NPC ($)', 'System LCOE ($/MWh)', 'Unmet Load (%)',
                             'PV (MW)', 'Wind (MW)', 'Hydro (MW)', 'BESS Power (MW)', 'BESS Energy (MWh)'],
                'Value': [results['npc'], results['lcoe'], results['unmet_pct'],
                         results['pv_capacity'], results['wind_capacity'], results['hydro_capacity'],
                         results['bess_power'], results['bess_energy']]
            }).to_excel(writer, sheet_name='Summary', index=False)
            
            # Cost Breakdown
            optimal_row = results.get('optimal_row', {})
            pd.DataFrame({
                'Component': ['PV', 'Wind', 'Hydro', 'BESS', 'System'],
                'Capital ($)': [
                    optimal_row.get('PV_Capital_$', 0),
                    optimal_row.get('Wind_Capital_$', 0),
                    optimal_row.get('Hydro_Capital_$', 0),
                    optimal_row.get('BESS_Capital_$', 0),
                    optimal_row.get('Capital_$', 0)
                ],
                'O&M ($)': [
                    optimal_row.get('PV_OM_$', 0),
                    optimal_row.get('Wind_OM_$', 0),
                    optimal_row.get('Hydro_OM_$', 0),
                    optimal_row.get('BESS_OM_$', 0),
                    optimal_row.get('OM_$', 0)
                ],
                'Total NPC ($)': [
                    results['pv_npc'],
                    results['wind_npc'],
                    results['hydro_npc'],
                    results['bess_npc'],
                    results['npc']
                ]
            }).to_excel(writer, sheet_name='Cost_Breakdown', index=False)
            
            # All Results
            results['results_df'].to_excel(writer, sheet_name='All_Results', index=False)
            
            # Feasible Solutions
            feasible = results['results_df'][results['results_df']['Feasible'] == True].sort_values('NPC_$')
            if len(feasible) > 0:
                feasible.to_excel(writer, sheet_name='Feasible_Solutions', index=False)
        
        output.seek(0)
        
        st.download_button(
            label="📥 Download Complete Results (Excel)",
            data=output.getvalue(),
            file_name=f"optimization_results_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        
        # Charts
        st.markdown("---")
        col1, col2 = st.columns(2)
        
        with col1:
            fig_cost = go.Figure(data=[go.Pie(
                labels=['PV', 'Wind', 'Hydro', 'BESS'],
                values=[results['pv_npc'], results['wind_npc'], results['hydro_npc'], results['bess_npc']],
                marker=dict(colors=['#FDB462', '#80B1D3', '#8DD3C7', '#FB8072'])
            )])
            fig_cost.update_layout(title="Cost Distribution", height=400)
            st.plotly_chart(fig_cost, use_container_width=True)
        
        with col2:
            fig_cap = go.Figure(data=[go.Bar(
                x=['PV', 'Wind', 'Hydro', 'BESS'],
                y=[results['pv_capacity'], results['wind_capacity'], results['hydro_capacity'], results['bess_power']],
                marker=dict(color=['#FDB462', '#80B1D3', '#8DD3C7', '#FB8072'])
            )])
            fig_cap.update_layout(title="Installed Capacity", yaxis_title="MW", height=400)
            st.plotly_chart(fig_cap, use_container_width=True)


# Footer
st.markdown("---")
st.markdown("""
<div style='text-align: center; color: #666;'>
    <p>Renewable Energy Optimization Tool v1.0</p>
    <p>HOMER-style NPC Calculation</p>
</div>
""", unsafe_allow_html=True)


