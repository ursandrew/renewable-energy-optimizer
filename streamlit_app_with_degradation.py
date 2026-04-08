"""
RENEWABLE ENERGY OPTIMIZATION TOOL - VERSION 5.1
==================================================
UPGRADES FROM v5.0:
- Wind degradation: Simple (annual rate) + Advanced (upload curve) — mirrors PV structure
- Hydro degradation: Grace period + annual rate UI (user controls stable years & rate)
  with live editable year-by-year preview table
- BESS degradation: Simple mode added (annual capacity rate + user-set efficiencies)
  Charge/discharge efficiency inputs re-enabled in simple mode
- LCOS call sites fixed: pass npc_data['crf'] instead of project_lifetime

Author: SJ | March 2026
"""

import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
from plotly.subplots import make_subplots
from datetime import datetime
from io import BytesIO

try:
    import optimize_gridsearch_hydro_WITH_DEGRADATION as opt_module
    OPTIMIZATION_AVAILABLE = True
except ImportError:
    OPTIMIZATION_AVAILABLE = False
    st.error("❌ Optimization module not found")


# ==============================================================================
# SUNGROW BESS DEPLOYMENT
# ==============================================================================

def calculate_bess_deployment_sungrow(bess_power_mw, bess_capacity_mwh):
    import math
    container_capacity_mwh = 10.0
    container_power_mw = 5.0
    container_length_m = 6.058
    container_width_m = 2.438
    back_to_back_spacing_m = 0.150
    mvs_spacing_m = 3.500
    mvs_width_m = 2.000
    adjacent_spacing_m = 1.500
    perimeter_clearance_m = 5.000

    num_containers_energy = math.ceil(bess_capacity_mwh / container_capacity_mwh)
    num_containers_power = math.ceil(bess_power_mw / container_power_mw)
    num_containers = max(num_containers_energy, num_containers_power)
    actual_capacity_mwh = num_containers * container_capacity_mwh
    actual_power_mw = num_containers * container_power_mw
    num_mvs_units = math.ceil(num_containers / 2)

    if num_containers <= 2:
        total_length_m = container_length_m + (2 * perimeter_clearance_m)
        container_section_width = container_width_m + back_to_back_spacing_m + container_width_m
        mvs_section_width = mvs_spacing_m + mvs_width_m
        total_width_m = container_section_width + mvs_section_width + (2 * perimeter_clearance_m)
        layout_desc = "1 section (2 containers back-to-back + 1 MVS unit)"
    else:
        num_sections = math.ceil(num_containers / 2)
        section_length = container_length_m + adjacent_spacing_m
        total_length_m = num_sections * section_length - adjacent_spacing_m + 2 * perimeter_clearance_m
        container_section_width = container_width_m + back_to_back_spacing_m + container_width_m
        mvs_section_width = mvs_spacing_m + mvs_width_m
        total_width_m = container_section_width + mvs_section_width + (2 * perimeter_clearance_m)
        layout_desc = f"{num_sections} sections side-by-side ({num_containers} containers + {num_mvs_units} MVS units)"

    total_area_m2 = total_length_m * total_width_m
    total_area_hectares = total_area_m2 / 10000
    total_area_acres = total_area_hectares * 2.471
    power_density = actual_power_mw / total_area_hectares if total_area_hectares > 0 else 0
    energy_density = actual_capacity_mwh / total_area_hectares if total_area_hectares > 0 else 0

    return {
        'num_containers': num_containers, 'container_model': 'PowerTitan 2.0',
        'container_capacity_mwh': container_capacity_mwh, 'container_power_mw': container_power_mw,
        'actual_capacity_mwh': actual_capacity_mwh, 'actual_power_mw': actual_power_mw,
        'num_mvs_units': num_mvs_units, 'total_length_m': total_length_m,
        'total_width_m': total_width_m, 'total_area_m2': total_area_m2,
        'total_area_hectares': total_area_hectares, 'total_area_acres': total_area_acres,
        'power_density_mw_per_ha': power_density, 'energy_density_mwh_per_ha': energy_density,
        'layout_description': layout_desc
    }


# ==============================================================================
# DEGRADATION FILE PARSERS
# ==============================================================================

def parse_pv_degradation_file(uploaded_file):
    try:
        df = pd.read_csv(uploaded_file) if uploaded_file.name.endswith('.csv') else pd.read_excel(uploaded_file)
        if 'Year' not in df.columns or 'PV_Degradation_%' not in df.columns:
            st.error("❌ Required columns: Year, PV_Degradation_%")
            return None
        deg_curve = dict(zip(df['Year'].astype(int), df['PV_Degradation_%']))
        st.success(f"✓ Loaded PV degradation curve: {len(deg_curve)} years")
        return deg_curve
    except Exception as e:
        st.error(f"❌ Error: {str(e)}")
        return None


def parse_wind_degradation_file(uploaded_file):
    try:
        df = pd.read_csv(uploaded_file) if uploaded_file.name.endswith('.csv') else pd.read_excel(uploaded_file)
        if 'Year' not in df.columns or 'Wind_Degradation_%' not in df.columns:
            st.error("❌ Required columns: Year, Wind_Degradation_%")
            return None
        deg_curve = dict(zip(df['Year'].astype(int), df['Wind_Degradation_%']))
        st.success(f"✓ Loaded Wind degradation curve: {len(deg_curve)} years")
        return deg_curve
    except Exception as e:
        st.error(f"❌ Error: {str(e)}")
        return None


def parse_bess_degradation_file(uploaded_file):
    try:
        df = pd.read_csv(uploaded_file) if uploaded_file.name.endswith('.csv') else pd.read_excel(uploaded_file)
        required_cols = ['Year', 'Capacity_Retention_%', 'Charging_Efficiency_%', 'Discharging_Efficiency_%']
        if not all(col in df.columns for col in required_cols):
            st.error(f"❌ Required columns: {', '.join(required_cols)}")
            return None
        deg_data = {}
        for _, row in df.iterrows():
            year = int(row['Year'])
            deg_data[year] = {
                'capacity': row['Capacity_Retention_%'],
                'charge_eff': row['Charging_Efficiency_%'],
                'discharge_eff': row['Discharging_Efficiency_%']
            }
        st.success(f"✓ Loaded BESS degradation curve: {len(deg_data)} years")
        return deg_data
    except Exception as e:
        st.error(f"❌ Error: {str(e)}")
        return None


def create_wind_degradation_template():
    years = list(range(1, 26))
    degradation = [0.3 * (year - 1) for year in years]
    df = pd.DataFrame({'Year': years, 'Wind_Degradation_%': degradation})
    return df.to_csv(index=False).encode('utf-8')


def create_pv_degradation_template():
    years = list(range(1, 26))
    degradation = [0.4 * (year - 1) for year in years]
    df = pd.DataFrame({'Year': years, 'PV_Degradation_%': degradation})
    return df.to_csv(index=False).encode('utf-8')


def create_bess_degradation_template():
    years = list(range(1, 26))
    capacity = [100 - (0.5 * (year - 1)) for year in years]
    df = pd.DataFrame({
        'Year': years, 'Capacity_Retention_%': capacity,
        'Charging_Efficiency_%': [90.0] * 25, 'Discharging_Efficiency_%': [98.5] * 25
    })
    return df.to_csv(index=False).encode('utf-8')


# ==============================================================================
# VISUALIZATION
# ==============================================================================

def create_single_day_dispatch_profile(results):
    if 'optimal_dispatch' not in results or results['optimal_dispatch'] is None:
        return None
    dispatch_df = results['optimal_dispatch'].copy()
    if 'Hour_of_Day' not in dispatch_df.columns:
        dispatch_df['Absolute_Hour'] = dispatch_df.index
        dispatch_df['Hour_of_Day'] = dispatch_df['Hour'] if 'Hour' in dispatch_df.columns else dispatch_df.index % 24
    else:
        dispatch_df['Absolute_Hour'] = dispatch_df.get('Hour', dispatch_df.index)
    dispatch_df['Day'] = dispatch_df['Absolute_Hour'] // 24
    pv_col = 'PV_Available_kW' if 'PV_Available_kW' in dispatch_df.columns else 'PV_Output_kW'
    daily_pv = dispatch_df.groupby('Day')[pv_col].sum()
    median_pv_day = daily_pv.sort_values().index[len(daily_pv) // 2]
    start_hour = median_pv_day * 24
    day_profile = dispatch_df[
        (dispatch_df['Absolute_Hour'] >= start_hour) &
        (dispatch_df['Absolute_Hour'] < start_hour + 24)
    ].copy()

    fig = make_subplots(specs=[[{"secondary_y": True}]])
    fig.add_trace(go.Scatter(x=day_profile['Hour_of_Day'], y=day_profile['Hydro_Output_kW'] / 1000,
        name='Hydro', fill='tozeroy', fillcolor='rgba(141,211,199,0.6)',
        line=dict(width=0.5, color='rgba(141,211,199,1)'),
        hovertemplate='Hour %{x}<br>Hydro: %{y:.2f} MW<extra></extra>'), secondary_y=False)
    fig.add_trace(go.Scatter(x=day_profile['Hour_of_Day'], y=day_profile[pv_col] / 1000,
        name='PV', fill='tozeroy', fillcolor='rgba(255,219,92,0.6)',
        line=dict(width=0.5, color='rgba(255,219,92,1)'),
        hovertemplate='Hour %{x}<br>PV: %{y:.2f} MW<extra></extra>'), secondary_y=False)
    if 'Wind_Output_kW' in day_profile.columns:
        fig.add_trace(go.Scatter(x=day_profile['Hour_of_Day'], y=day_profile['Wind_Output_kW'] / 1000,
            name='Wind', fill='tozeroy', fillcolor='rgba(179,226,205,0.6)',
            line=dict(width=0.5, color='rgba(179,226,205,1)'),
            hovertemplate='Hour %{x}<br>Wind: %{y:.2f} MW<extra></extra>'), secondary_y=False)
    fig.add_trace(go.Scatter(x=day_profile['Hour_of_Day'], y=day_profile['Load_kW'] / 1000,
        name='Load', mode='lines', line=dict(color='red', width=2),
        hovertemplate='Hour %{x}<br>Load: %{y:.2f} MW<extra></extra>'), secondary_y=False)
    if 'BESS_SOC_pct' in day_profile.columns:
        fig.add_trace(go.Scatter(x=day_profile['Hour_of_Day'], y=day_profile['BESS_SOC_pct'],
            name='BESS SOC', mode='lines', line=dict(color='purple', width=2, dash='dash'),
            hovertemplate='Hour %{x}<br>SOC: %{y:.1f}%<extra></extra>'), secondary_y=True)
    fig.update_xaxes(title_text="Hour of Day", range=[0, 23])
    fig.update_yaxes(title_text="Power (MW)", secondary_y=False)
    fig.update_yaxes(title_text="BESS SOC (%)", secondary_y=True, range=[0, 100])
    fig.update_layout(title=f'Typical Day Dispatch Profile (Day {median_pv_day + 1} - Median PV)',
        hovermode='x unified', height=500, showlegend=True,
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1))
    return fig


def create_energy_mix_pie(results):
    optimal = results['optimal_solution']
    values, labels, colors = [], [], []
    if optimal['PV_Energy_kWh'] > 0:
        values.append(optimal['PV_Energy_kWh'] / 1000); labels.append('Solar PV'); colors.append('#FFDB5C')
    if optimal['Wind_Energy_kWh'] > 0:
        values.append(optimal['Wind_Energy_kWh'] / 1000); labels.append('Wind'); colors.append('#B3E2CD')
    if optimal['Hydro_Energy_kWh'] > 0:
        values.append(optimal['Hydro_Energy_kWh'] / 1000); labels.append('Hydro'); colors.append('#8DD3C7')
    if not values:
        return None
    fig = go.Figure(data=[go.Pie(labels=labels, values=values, hole=0.4,
        marker=dict(colors=colors))])
    fig.update_layout(title='Annual Energy Mix', height=400,
        annotations=[dict(text='MWh', x=0.5, y=0.5, font_size=20, showarrow=False)])
    return fig


def build_excel_export(results, optimal, opt_module):
    output = BytesIO()
    crf = results['npc_data']['crf']
    bess_annual_discharge = 0
    lcos = 0
    if results.get('optimal_dispatch') is not None:
        bess_annual_discharge = results['optimal_dispatch']['BESS_Discharge_wieff_kW'].sum()
    bess_npc_val = optimal.get('BESS_NPC', 0)
    if bess_annual_discharge > 0 and hasattr(opt_module, 'calculate_bess_lcos_from_npc'):
        lcos = opt_module.calculate_bess_lcos_from_npc(bess_npc_val, bess_annual_discharge, crf)
    re_penetration = (
        optimal['PV_Energy_kWh'] + optimal['Wind_Energy_kWh'] + optimal['Hydro_Energy_kWh']
    ) / optimal['Total_Load_kWh'] * 100 if optimal['Total_Load_kWh'] > 0 else 0

    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        summary_data = {
            'Parameter': ['PV Capacity (MW)', 'Wind Capacity (MW)', 'Hydro Capacity (MW)',
                'BESS Power (MW)', 'BESS Energy (MWh)', 'Net Present Cost ($M)',
                'LCOE ($/MWh)', 'LCOS ($/MWh)', 'Unmet Load (%)', 'RE Penetration (%)',
                'Annual PV Energy (MWh)', 'Annual Wind Energy (MWh)',
                'Annual Hydro Energy (MWh)', 'Annual Load (MWh)',
                'Total Capital ($M)', 'Annual O&M ($k)', 'Annual BESS Discharge (MWh)'],
            'Value': [
                round(optimal['PV_kW'] / 1000, 3), round(optimal['Wind_kW'] / 1000, 3),
                round(optimal['Hydro_kW'] / 1000, 3), round(optimal['BESS_Power_kW'] / 1000, 3),
                round(optimal['BESS_Capacity_kWh'] / 1000, 3),
                round(optimal['NPC_Total'] / 1_000_000, 4),
                round(optimal['LCOE_per_kWh'] * 1000, 4), round(lcos * 1000, 4),
                round(optimal['Unmet_Load_Percent'], 4), round(re_penetration, 2),
                round(optimal['PV_Energy_kWh'] / 1000, 1), round(optimal['Wind_Energy_kWh'] / 1000, 1),
                round(optimal['Hydro_Energy_kWh'] / 1000, 1), round(optimal['Total_Load_kWh'] / 1000, 1),
                round(optimal['CapEx_Total'] / 1_000_000, 4),
                round(optimal['OpEx_Annual'] / 1000, 2), round(bess_annual_discharge / 1000, 1)
            ]
        }
        pd.DataFrame(summary_data).to_excel(writer, sheet_name='Optimal Configuration', index=False)
        if 'all_results' in results:
            results['all_results'].to_excel(writer, sheet_name='All Combinations', index=False)
        if results.get('optimal_dispatch') is not None:
            results['optimal_dispatch'].to_excel(writer, sheet_name='Dispatch Year 1', index=False)
        if results.get('degradation_enabled') and 'degradation_analysis' in results:
            deg_data = results['degradation_analysis']
            deg_data['yearly_metrics'].to_excel(writer, sheet_name='Degradation Summary', index=False)
            for year_key, dispatch_df in deg_data['selected_year_dispatch'].items():
                yr = year_key.split('_')[1]
                dispatch_df.to_excel(writer, sheet_name=f'Dispatch Year {yr}', index=False)
    output.seek(0)
    return output


# ==============================================================================
# PAGE CONFIG
# ==============================================================================

st.set_page_config(page_title="Energy Modeling Optimizer v5.1", page_icon="⚡",
    layout="wide", initial_sidebar_state="expanded")

st.markdown("""
<div style="display:flex;align-items:center;justify-content:center;gap:14px;margin-bottom:4px">
    <svg width="52" height="52" viewBox="0 0 52 52" xmlns="http://www.w3.org/2000/svg">
        <rect width="52" height="52" rx="0" fill="#0047AB"/>
        <text x="18" y="38" font-family="Arial,sans-serif" font-size="32"
              font-weight="bold" fill="white" text-anchor="middle">S</text>
        <text x="36" y="37" font-family="Arial,sans-serif" font-size="18"
              font-weight="bold" fill="white" text-anchor="middle">J</text>
        <circle cx="38" cy="18" r="4" fill="#E63946"/>
    </svg>
    <p style="font-size:2.5rem;font-weight:bold;color:#1f77b4;margin:0">
        Energy Modeling Optimizer
    </p>
</div>
""", unsafe_allow_html=True)
st.markdown("**Hybrid System Designer: PV + Wind + Hydro + Battery Storage**")
st.markdown("---")

if 'optimization_complete' not in st.session_state:
    st.session_state.optimization_complete = False
if 'results' not in st.session_state:
    st.session_state.results = None


# ==============================================================================
# SIDEBAR
# ==============================================================================

with st.sidebar:
    st.header("⚙️ System Configuration")
    st.subheader("🔌 Component Selection")

    col1, col2 = st.columns(2)
    with col1:
        enable_pv    = st.checkbox("☀️ Solar PV", value=True, key="enable_pv")
        enable_wind  = st.checkbox("💨 Wind",     value=True, key="enable_wind")
        enable_hydro = st.checkbox("💧 Hydro",    value=True, key="enable_hydro")
    with col2:
        enable_bess  = st.checkbox("🔋 BESS",     value=True, key="enable_bess")

    if not any([enable_pv, enable_wind, enable_hydro, enable_bess]):
        st.error("⚠️ At least one component must be enabled!")

    st.markdown("---")

    # ── SOLAR PV ──────────────────────────────────────────────────────────────
    with st.expander("☀️ SOLAR PV", expanded=enable_pv):
        if not enable_pv:
            st.warning("⚠️ Solar PV is DISABLED")
            pv_min = pv_max = pv_step = 0.0; pv_capex = 1000; pv_opex = 10; pv_lifetime = 25
            apply_pv_degradation = False; pv_deg_method = None
            pv_annual_deg_rate = None; pv_deg_file = None
        else:
            col1, col2 = st.columns(2)
            with col1: pv_min = st.number_input("Min (MW)", value=1.0, min_value=0.0, step=0.5, key="pv_min")
            with col2: pv_max = st.number_input("Max (MW)", value=5.0, min_value=0.0, step=0.5, key="pv_max")
            pv_step = st.number_input("Step (MW)", value=1.0, min_value=0.1, step=0.1, key="pv_step")

            col1, col2 = st.columns(2)
            with col1:
                pv_capex = st.number_input("CapEx ($/kW)", value=1000, step=10, key="pv_capex")
                pv_opex  = st.number_input("OpEx ($/kW/yr)", value=10, step=1, key="pv_opex")
            with col2:
                pv_lifetime = st.number_input("Lifetime (years)", value=25, step=1, key="pv_life")

            st.markdown("---")
            apply_pv_degradation = st.checkbox("Apply PV Degradation Analysis", value=False, key="apply_pv_deg")
            if apply_pv_degradation:
                pv_deg_method = st.radio("Method:", ["Simple (Annual Rate)", "Advanced (Upload Curve)"],
                    horizontal=True, key="pv_deg_method")
                if pv_deg_method == "Simple (Annual Rate)":
                    pv_annual_deg_rate = st.number_input("Annual Degradation Rate (%/yr)", value=0.40,
                        min_value=0.0, max_value=2.0, step=0.05, key="pv_annual_deg")
                    yr25 = (1 - (1 - pv_annual_deg_rate / 100) ** 24) * 100
                    st.info(f"📊 Cumulative degradation at Year 25: ~{yr25:.2f}%")
                    pv_deg_file = None
                    st.download_button("📥 Download Template CSV", create_pv_degradation_template(),
                        "pv_degradation_template.csv", "text/csv", key="pv_tpl")
                else:
                    pv_deg_file = st.file_uploader("Upload PV Degradation Curve",
                        type=['csv', 'xlsx'], key="pv_deg_file")
                    pv_annual_deg_rate = None
                    st.download_button("📥 Download Template CSV", create_pv_degradation_template(),
                        "pv_degradation_template.csv", "text/csv", key="pv_tpl_adv")
            else:
                pv_deg_method = None; pv_annual_deg_rate = None; pv_deg_file = None

    # ── WIND ──────────────────────────────────────────────────────────────────
    with st.expander("💨 WIND"):
        if not enable_wind:
            st.warning("⚠️ Wind is DISABLED")
            wind_min = wind_max = wind_step = 0.0; wind_capex = 1200; wind_opex = 15; wind_lifetime = 25
            apply_wind_degradation = False; wind_deg_method = None
            wind_annual_deg_rate = None; wind_deg_file = None
        else:
            col1, col2 = st.columns(2)
            with col1: wind_min = st.number_input("Min (MW)", value=1.0, min_value=0.0, step=0.5, key="wind_min")
            with col2: wind_max = st.number_input("Max (MW)", value=5.0, min_value=0.0, step=0.5, key="wind_max")
            wind_step = st.number_input("Step (MW)", value=1.0, min_value=0.1, step=0.1, key="wind_step")

            col1, col2 = st.columns(2)
            with col1:
                wind_capex = st.number_input("CapEx ($/kW)", value=1200, step=10, key="wind_capex")
                wind_opex  = st.number_input("OpEx ($/kW/yr)", value=15, step=1, key="wind_opex")
            with col2:
                wind_lifetime = st.number_input("Lifetime (years)", value=25, step=1, key="wind_life")

            st.markdown("---")
            apply_wind_degradation = st.checkbox("Apply Wind Degradation Analysis", value=False,
                key="apply_wind_deg",
                help="Model turbine output decline from blade erosion, bearing wear, etc. Typical: 0.1–0.5%/yr")
            if apply_wind_degradation:
                wind_deg_method = st.radio("Method:",
                    ["Simple (Annual Rate)", "Advanced (Upload Curve)"],
                    horizontal=True, key="wind_deg_method")
                if wind_deg_method == "Simple (Annual Rate)":
                    wind_annual_deg_rate = st.number_input(
                        "Annual Degradation Rate (%/yr)", value=0.30,
                        min_value=0.0, max_value=2.0, step=0.05, key="wind_annual_deg",
                        help="Typical wind turbine degradation: 0.1–0.5%/year")
                    yr25_wind = (1 - (1 - wind_annual_deg_rate / 100) ** 24) * 100
                    st.info(f"📊 Cumulative degradation at Year 25: ~{yr25_wind:.2f}%")
                    wind_deg_file = None
                    st.download_button("📥 Download Template CSV", create_wind_degradation_template(),
                        "wind_degradation_template.csv", "text/csv", key="wind_tpl")
                else:
                    wind_deg_file = st.file_uploader("Upload Wind Degradation Curve",
                        type=['csv', 'xlsx'],
                        help="Required columns: Year, Wind_Degradation_%  (cumulative %)",
                        key="wind_deg_file")
                    wind_annual_deg_rate = None
                    if wind_deg_file:
                        st.success(f"✓ Uploaded: {wind_deg_file.name}")
                    st.download_button("📥 Download Template CSV", create_wind_degradation_template(),
                        "wind_degradation_template.csv", "text/csv", key="wind_tpl_adv")
            else:
                wind_deg_method = None; wind_annual_deg_rate = None; wind_deg_file = None

    # ── HYDRO ─────────────────────────────────────────────────────────────────
    with st.expander("💧 HYDRO"):
        if not enable_hydro:
            st.warning("⚠️ Hydro is DISABLED")
            hydro_min = hydro_max = hydro_step = 0.0; hydro_hours_per_day = 6
            hydro_capex = 1500; hydro_opex = 20; hydro_lifetime = 50
            apply_hydro_degradation = False
            hydro_stable_years = 15; hydro_deg_rate_after = 0.5
        else:
            col1, col2 = st.columns(2)
            with col1: hydro_min = st.number_input("Min (MW)", value=1.0, min_value=0.0, step=0.5, key="hydro_min")
            with col2: hydro_max = st.number_input("Max (MW)", value=5.0, min_value=0.0, step=0.5, key="hydro_max")
            hydro_step = st.number_input("Step (MW)", value=1.0, min_value=0.1, step=0.1, key="hydro_step")
            hydro_hours_per_day = st.number_input("Operating Hours/Day", value=6,
                min_value=1, max_value=24, step=1, key="hydro_hours")

            col1, col2 = st.columns(2)
            with col1:
                hydro_capex = st.number_input("CapEx ($/kW)", value=1500, step=10, key="hydro_capex")
                hydro_opex  = st.number_input("OpEx ($/kW/yr)", value=20, step=1, key="hydro_opex")
            with col2:
                hydro_lifetime = st.number_input("Lifetime (years)", value=50, step=1, key="hydro_life")

            st.markdown("---")
            apply_hydro_degradation = st.checkbox("Apply Hydro Degradation Analysis", value=False,
                key="apply_hydro_deg",
                help="Hydro plants typically maintain full output for 15–20 years before gradual decline from turbine wear and sediment.")

            if apply_hydro_degradation:
                st.info("ℹ️ Hydro plants have a stable grace period before output declines.")
                col1, col2 = st.columns(2)
                with col1:
                    hydro_stable_years = st.number_input(
                        "Stable Period (years)", value=15, min_value=1,
                        max_value=int(project_lifetime) if 'project_lifetime' in dir() else 25,
                        step=1, key="hydro_stable_years",
                        help="Years with no output degradation (typically 15–20 years)")
                with col2:
                    hydro_deg_rate_after = st.number_input(
                        "Annual Rate After Stable Period (%/yr)", value=0.50,
                        min_value=0.0, max_value=5.0, step=0.05, key="hydro_deg_rate",
                        help="Annual output degradation after the stable period ends")

                # Build preview table
                preview_lt = 25
                hydro_preview_table = {}
                for yr in range(1, preview_lt + 1):
                    if yr <= hydro_stable_years:
                        hydro_preview_table[yr] = 100.0
                    else:
                        years_after = yr - hydro_stable_years
                        hydro_preview_table[yr] = round(
                            (1 - hydro_deg_rate_after / 100) ** years_after * 100, 2)

                # Show preview table (collapsible)
                with st.expander("📋 Preview Degradation Table (all years)", expanded=False):
                    preview_df = pd.DataFrame({
                        'Year': list(hydro_preview_table.keys()),
                        'Output Factor (%)': list(hydro_preview_table.values()),
                        'Degradation (%)': [round(100 - v, 2) for v in hydro_preview_table.values()]
                    })
                    # Highlight stable vs degrading rows
                    def color_rows(row):
                        if row['Year'] <= hydro_stable_years:
                            return ['background-color: #e8f5e9'] * 3
                        else:
                            return ['background-color: #fff3e0'] * 3
                    st.dataframe(preview_df.style.apply(color_rows, axis=1),
                        use_container_width=True, hide_index=True, height=300)
                    st.caption("🟢 Green = stable period  |  🟠 Orange = degrading period")

                # Key milestones
                yr_end = preview_lt
                end_factor = hydro_preview_table.get(yr_end, 100)
                st.info(
                    f"📊 Year {hydro_stable_years}: 100.0% output (end of stable period)\n\n"
                    f"📊 Year {yr_end}: {end_factor:.2f}% output "
                    f"({100-end_factor:.2f}% cumulative loss)"
                )
            else:
                hydro_stable_years = 15
                hydro_deg_rate_after = 0.5

    # ── BESS ──────────────────────────────────────────────────────────────────
    with st.expander("🔋 BATTERY STORAGE"):
        if not enable_bess:
            st.warning("⚠️ BESS is DISABLED")
            bess_min = bess_max = bess_step = 0.0; bess_duration = 4.0
            bess_min_soc = 10.0; bess_max_soc = 90.0
            bess_charge_eff = 90.0; bess_discharge_eff = 95.0
            bess_power_capex = 300; bess_energy_capex = 300; bess_opex = 10; bess_lifetime = 15
            apply_bess_degradation = False; bess_deg_mode = None
            bess_annual_cap_deg = None; bess_deg_file = None
            bess_annual_charge_eff_deg = 0.0; bess_annual_discharge_eff_deg = 0.0
        else:
            col1, col2 = st.columns(2)
            with col1: bess_min = st.number_input("Min Power (MW)", value=1.0, min_value=0.0, step=0.5, key="bess_min")
            with col2: bess_max = st.number_input("Max Power (MW)", value=5.0, min_value=0.0, step=0.5, key="bess_max")
            bess_step = st.number_input("Step (MW)", value=1.0, min_value=0.1, step=0.1, key="bess_step")

            col1, col2 = st.columns(2)
            with col1:
                bess_duration = st.number_input("Duration (hours)", value=4.0, min_value=0.5, step=0.5, key="bess_duration")
                bess_min_soc  = st.number_input("Min SOC (%)", value=10.0, min_value=0.0, max_value=50.0, step=5.0, key="bess_min_soc")
            with col2:
                bess_max_soc  = st.number_input("Max SOC (%)", value=90.0, min_value=50.0, max_value=100.0, step=5.0, key="bess_max_soc")

            st.markdown("---")
            apply_bess_degradation = st.checkbox("Apply BESS Degradation Analysis", value=False,
                key="apply_bess_deg")

            if apply_bess_degradation:
                # ── Degradation mode selection ──
                bess_deg_mode = st.radio(
                    "Degradation Method:",
                    ["Simple (Annual Rate)", "Advanced (Upload CSV Curve)"],
                    horizontal=True, key="bess_deg_mode",
                    help="Simple: enter one annual capacity degradation rate. Advanced: upload a year-by-year CSV curve."
                )

                if bess_deg_mode == "Simple (Annual Rate)":
                    st.markdown("**Capacity Degradation**")
                    bess_annual_cap_deg = st.number_input(
                        "Annual Capacity Degradation Rate (%/yr)", value=2.0,
                        min_value=0.0, max_value=10.0, step=0.1, key="bess_annual_cap_deg",
                        help="Battery capacity retention degrades each year. Typical: 1.5–3%/yr")
                    yr25_cap = (1 - bess_annual_cap_deg / 100) ** 24 * 100
                    st.info(f"📊 Capacity retention at Year 25: ~{yr25_cap:.1f}%  |  Loss: ~{100-yr25_cap:.1f}%")

                    st.markdown("**Efficiency Parameters**")
                    st.caption("Set Year-1 baseline values and annual degradation rate for each.")
                    col1, col2 = st.columns(2)
                    with col1:
                        bess_charge_eff = st.number_input(
                            "Charge Efficiency Year-1 (%)", value=90.0,
                            min_value=50.0, max_value=100.0, step=1.0, key="bess_charge_eff_simple")
                        bess_annual_charge_eff_deg = st.number_input(
                            "Charge Eff. Degradation (%/yr)", value=0.5,
                            min_value=0.0, max_value=5.0, step=0.1, key="bess_chg_eff_deg",
                            help="Annual decline in charging efficiency. Set 0 to keep constant.")
                    with col2:
                        bess_discharge_eff = st.number_input(
                            "Discharge Efficiency Year-1 (%)", value=95.0,
                            min_value=50.0, max_value=100.0, step=1.0, key="bess_discharge_eff_simple")
                        bess_annual_discharge_eff_deg = st.number_input(
                            "Discharge Eff. Degradation (%/yr)", value=0.2,
                            min_value=0.0, max_value=5.0, step=0.1, key="bess_dis_eff_deg",
                            help="Annual decline in discharging efficiency. Set 0 to keep constant.")

                    # Year-25 preview for efficiencies
                    yr25_chg = bess_charge_eff    * (1 - bess_annual_charge_eff_deg    / 100) ** 24
                    yr25_dis = bess_discharge_eff * (1 - bess_annual_discharge_eff_deg / 100) ** 24
                    st.info(
                        f"📊 Year 25 → "
                        f"Charge eff: {yr25_chg:.2f}%  |  "
                        f"Discharge eff: {yr25_dis:.2f}%"
                    )
                    bess_deg_file = None

                else:  # Advanced CSV mode
                    bess_annual_cap_deg = None
                    bess_annual_charge_eff_deg = 0.0
                    bess_annual_discharge_eff_deg = 0.0
                    st.info("ℹ️ Efficiency values are controlled by the degradation CSV (inputs below are ignored).")
                    bess_deg_file = st.file_uploader(
                        "Select BESS Degradation CSV File", type=['csv', 'xlsx'],
                        help="Required columns: Year, Capacity_Retention_%, Charging_Efficiency_%, Discharging_Efficiency_%",
                        key="bess_deg_file")
                    if bess_deg_file:
                        st.success(f"✓ Uploaded: {bess_deg_file.name}")
                        try:
                            preview = parse_bess_degradation_file(bess_deg_file)
                            if preview:
                                st.info(
                                    f"Year 1 → Capacity: {preview[1]['capacity']:.1f}% | "
                                    f"Chg: {preview[1]['charge_eff']:.1f}% | "
                                    f"Dis: {preview[1]['discharge_eff']:.1f}%\n\n"
                                    f"Year 25 → Capacity: {preview[25]['capacity']:.1f}% | "
                                    f"Chg: {preview[25]['charge_eff']:.1f}% | "
                                    f"Dis: {preview[25]['discharge_eff']:.1f}%"
                                )
                                bess_deg_file.seek(0)
                        except Exception:
                            pass

                    # Preset downloads
                    st.markdown("**Download a preset curve:**")
                    col1, col2, col3 = st.columns(3)
                    with col1:
                        try: nmc_data = open('bess_degradation_lithium_nmc.csv', 'rb').read()
                        except: nmc_data = create_bess_degradation_template()
                        st.download_button("📥 Lithium NMC", nmc_data,
                            "bess_degradation_lithium_nmc.csv", "text/csv", key="bess_nmc")
                    with col2:
                        try: lfp_data = open('bess_degradation_lithium_lfp.csv', 'rb').read()
                        except: lfp_data = create_bess_degradation_template()
                        st.download_button("📥 Lithium LFP", lfp_data,
                            "bess_degradation_lithium_lfp.csv", "text/csv", key="bess_lfp")
                    with col3:
                        try: sodi_data = open('bess_degradation_sodium_ion.csv', 'rb').read()
                        except: sodi_data = create_bess_degradation_template()
                        st.download_button("📥 Sodium-Ion", sodi_data,
                            "bess_degradation_sodium_ion.csv", "text/csv", key="bess_sodi")

                    # Efficiency inputs disabled in CSV mode
                    col1, col2 = st.columns(2)
                    with col1:
                        bess_charge_eff = st.number_input("Charge Efficiency (%)", value=90.0,
                            min_value=50.0, max_value=100.0, step=1.0,
                            key="bess_charge_eff_csv", disabled=True)
                    with col2:
                        bess_discharge_eff = st.number_input("Discharge Efficiency (%)", value=95.0,
                            min_value=50.0, max_value=100.0, step=1.0,
                            key="bess_discharge_eff_csv", disabled=True)

            else:
                # No degradation — both efficiency inputs enabled
                bess_deg_mode = None; bess_annual_cap_deg = None; bess_deg_file = None
                bess_annual_charge_eff_deg = 0.0; bess_annual_discharge_eff_deg = 0.0
                col1, col2 = st.columns(2)
                with col1:
                    bess_charge_eff = st.number_input("Charge Efficiency (%)", value=90.0,
                        min_value=50.0, max_value=100.0, step=1.0, key="bess_charge_eff")
                with col2:
                    bess_discharge_eff = st.number_input("Discharge Efficiency (%)", value=95.0,
                        min_value=50.0, max_value=100.0, step=1.0, key="bess_discharge_eff")

            st.subheader("Financial Parameters")
            col1, col2 = st.columns(2)
            with col1:
                bess_power_capex  = st.number_input("Power CapEx ($/kW)", value=300, step=10, key="bess_power_capex")
                bess_energy_capex = st.number_input("Energy CapEx ($/kWh)", value=300, step=10, key="bess_energy_capex")
            with col2:
                bess_opex     = st.number_input("OpEx ($/kW/yr)", value=10, step=1, key="bess_opex")
                bess_lifetime = st.number_input("Lifetime (years)", value=15, step=1, key="bess_life")

    st.markdown("---")

    # ── UPLOAD PROFILES ───────────────────────────────────────────────────────
    with st.expander("📁 UPLOAD PROFILES", expanded=True):
        st.markdown("**Upload your energy profiles (CSV/Excel):**")
        load_file = st.file_uploader("📊 Load Profile (Required)", type=['csv', 'xlsx'],
            key="load_file", help="8760-hour load profile in kW")
        if load_file: st.success(f"✓ {load_file.name}")

        pv_file = st.file_uploader("☀️ PV Profile (Required if PV enabled)", type=['csv', 'xlsx'],
            key="pv_file", help="8760-hour PV generation profile (1 kW baseline)")
        if pv_file: st.success(f"✓ {pv_file.name}")

        wind_file = st.file_uploader("💨 Wind Profile (Required if Wind enabled)", type=['csv', 'xlsx'],
            key="wind_file", help="8760-hour wind generation profile")
        if wind_file: st.success(f"✓ {wind_file.name}")

        hydro_file = st.file_uploader("💧 Hydro Profile (Optional)", type=['csv', 'xlsx'],
            key="hydro_file", help="8760-hour hydro availability profile. If omitted, constant 24/7 availability assumed.")
        if hydro_file: st.success(f"✓ {hydro_file.name}")

    st.markdown("---")

    # ── PROJECT PARAMETERS ────────────────────────────────────────────────────
    with st.expander("💰 PROJECT PARAMETERS"):
        col1, col2 = st.columns(2)
        with col1:
            discount_rate  = st.number_input("Nominal Discount Rate (%)", value=8.0,
                min_value=0.0, max_value=30.0, step=0.5, key="discount_rate")
            inflation_rate = st.number_input("Inflation Rate (%)", value=2.0,
                min_value=0.0, max_value=10.0, step=0.5, key="inflation_rate")
        with col2:
            project_lifetime = st.number_input("Project Lifetime (years)", value=25,
                min_value=10, max_value=50, step=5, key="project_lifetime")
            target_unmet_percent = st.number_input("Target Max Unmet Load (%)", value=5.0,
                min_value=0.0, max_value=20.0, step=0.5, key="target_unmet")

        st.markdown("---")
        st.markdown("**📊 Cash Flow LCOE — Cost Overrides**")
        st.caption(
            "The cash flow LCOE method uses nominal costs per unit capacity. "
            "These default to the values entered above but can be overridden here "
            "for the economic analysis tab."
        )

        # PV
        with st.expander("☀️ PV — Cash Flow Costs", expanded=False):
            col1, col2, col3 = st.columns(3)
            with col1:
                cf_pv_capex_per_mwp = st.number_input(
                    "PV CAPEX (USD/MWp)", value=float(pv_capex if enable_pv else 1000) * 1000,
                    min_value=0.0, step=10000.0, key="cf_pv_capex",
                    help="Capital cost per MWp of DC installed capacity")
            with col2:
                cf_pv_om_per_mwp = st.number_input(
                    "PV Fixed O&M (USD/MWp-yr)", value=float(pv_opex if enable_pv else 10) * 1000,
                    min_value=0.0, step=1000.0, key="cf_pv_om",
                    help="Annual O&M per MWp (escalates with inflation)")
            with col3:
                cf_pv_inv_bop_per_mwac = st.number_input(
                    "PV Inverter/BoP (USD/MWac)", value=50000.0,
                    min_value=0.0, step=1000.0, key="cf_pv_inv",
                    help="One-off inverter & Balance-of-Plant replacement cost per MWac at Year 0")
            cf_dc_ac_ratio = st.number_input(
                "DC/AC Ratio", value=1.5, min_value=1.0, max_value=2.0, step=0.05,
                key="cf_dc_ac", help="Used to derive PV AC from PV DC (MWp) capacity")

        # Wind
        with st.expander("💨 Wind — Cash Flow Costs", expanded=False):
            col1, col2 = st.columns(2)
            with col1:
                cf_wind_capex_per_kw = st.number_input(
                    "Wind CAPEX (USD/kW)", value=float(wind_capex if enable_wind else 1200),
                    min_value=0.0, step=10.0, key="cf_wind_capex")
            with col2:
                cf_wind_om_per_kw = st.number_input(
                    "Wind O&M (USD/kW-yr)", value=float(wind_opex if enable_wind else 15),
                    min_value=0.0, step=1.0, key="cf_wind_om")

        # Hydro
        with st.expander("💧 Hydro — Cash Flow Costs", expanded=False):
            col1, col2 = st.columns(2)
            with col1:
                cf_hydro_capex_per_kw = st.number_input(
                    "Hydro CAPEX (USD/kW)", value=float(hydro_capex if enable_hydro else 1500),
                    min_value=0.0, step=10.0, key="cf_hydro_capex")
            with col2:
                cf_hydro_om_per_kw = st.number_input(
                    "Hydro O&M (USD/kW-yr)", value=float(hydro_opex if enable_hydro else 20),
                    min_value=0.0, step=1.0, key="cf_hydro_om")

        # BESS
        with st.expander("🔋 BESS — Cash Flow Costs", expanded=False):
            col1, col2 = st.columns(2)
            with col1:
                cf_bess_capex_per_kwh = st.number_input(
                    "BESS CAPEX (USD/kWh)", value=float(bess_energy_capex if enable_bess else 300),
                    min_value=0.0, step=10.0, key="cf_bess_capex",
                    help="Capital cost per kWh of BESS energy capacity")
            with col2:
                cf_bess_om_per_kwh = st.number_input(
                    "BESS Fixed O&M (USD/kWh-yr)", value=20.0,
                    min_value=0.0, step=1.0, key="cf_bess_om",
                    help="Annual O&M per kWh of BESS capacity (escalates with inflation)")


# ==============================================================================
# MAIN — RUN OPTIMIZATION
# ==============================================================================

st.header("🚀 Run Optimization")

validation_errors = []
if not OPTIMIZATION_AVAILABLE:
    validation_errors.append("❌ Optimization module not available")
if load_file is None:
    validation_errors.append("❌ Load profile is required")
if enable_pv and pv_file is None:
    validation_errors.append("❌ PV profile required when PV is enabled")
if enable_wind and wind_file is None:
    validation_errors.append("❌ Wind profile required when Wind is enabled")

if validation_errors:
    for e in validation_errors: st.error(e)
    st.stop()

if not enable_wind and wind_file is None:
    pass
if enable_hydro and hydro_file is None:
    st.info("ℹ️ Hydro enabled but no profile uploaded — constant 24/7 availability assumed.")


if st.button("▶️ RUN OPTIMIZATION", type="primary", use_container_width=True):
    with st.spinner("Running optimization..."):
        try:
            progress_bar = st.progress(0)
            status_text  = st.empty()

            # ── Load profiles ──
            status_text.text("📂 Loading input profiles...")
            progress_bar.progress(10)

            load_df = pd.read_csv(load_file) if load_file.name.endswith('.csv') else pd.read_excel(load_file)
            pv_df   = (pd.read_csv(pv_file) if pv_file.name.endswith('.csv') else pd.read_excel(pv_file)) if pv_file else pd.DataFrame({'PVsyst_kW': [0] * 8760})
            wind_df = (pd.read_csv(wind_file) if wind_file.name.endswith('.csv') else pd.read_excel(wind_file)) if wind_file else pd.DataFrame({'Wind_kW': [0] * 8760})
            hydro_df = (pd.read_csv(hydro_file) if hydro_file.name.endswith('.csv') else pd.read_excel(hydro_file)) if hydro_file else pd.DataFrame({'Hydro_Available_kW': [1.0] * 8760})

            def extract_profile(df, name):
                raw = df.iloc[:, 0].values if len(df.columns) == 1 else df.iloc[:, 1].values
                if len(raw) < 8760:
                    st.error(f"❌ {name} profile has {len(raw)} rows — needs 8760.")
                    st.stop()
                return raw[:8760].astype(float)

            load_profile   = extract_profile(load_df,  "Load")
            pvsyst_profile = extract_profile(pv_df,    "PV")
            wind_profile   = extract_profile(wind_df,  "Wind")

            progress_bar.progress(20)

            # ── Build configs ──
            status_text.text("⚙️ Configuring parameters...")

            config = {
                'simulation_hours': 8760,
                'target_unmet_percent': target_unmet_percent,
                'discount_rate': discount_rate / 100,
                'inflation_rate': inflation_rate / 100,
                'project_lifetime': project_lifetime,
                'pv_lifetime': pv_lifetime,
                'wind_lifetime': wind_lifetime,
                'hydro_lifetime': hydro_lifetime,
                'bess_lifetime': bess_lifetime
            }

            grid_config = {
                'pv_start':   pv_min * 1000,   'pv_end':   pv_max * 1000,   'pv_step':   pv_step * 1000,
                'wind_start': wind_min * 1000,  'wind_end': wind_max * 1000, 'wind_step': wind_step * 1000,
                'hydro_start':hydro_min * 1000, 'hydro_end':hydro_max * 1000,'hydro_step':hydro_step * 1000,
                'bess_start': bess_min * 1000,  'bess_end': bess_max * 1000, 'bess_step': bess_step * 1000
            }

            solar_config = {'capex_per_kw': pv_capex, 'om_per_kw_year': pv_opex,
                            'lifetime': pv_lifetime, 'baseline_kw': 1.0}
            wind_config  = {'capex_per_kw': wind_capex, 'om_per_kw_year': wind_opex,
                            'lifetime': wind_lifetime, 'enabled': enable_wind}
            hydro_config = {'capex_per_kw': hydro_capex, 'om_per_kw_year': hydro_opex,
                            'lifetime': hydro_lifetime, 'hours_per_day': hydro_hours_per_day}
            bess_config  = {
                'duration_hours': bess_duration,
                'min_soc': bess_min_soc, 'max_soc': bess_max_soc,
                'charge_eff':    bess_charge_eff / 100,
                'discharge_eff': bess_discharge_eff / 100,
                'power_capex_per_kw': bess_power_capex,
                'energy_capex_per_kwh': bess_energy_capex,
                'om_per_kw_year': bess_opex, 'lifetime': bess_lifetime
            }

            progress_bar.progress(30)

            # ── Grid search ──
            status_text.text("⚙️ Running grid search optimization...")
            results_df = opt_module.grid_search_optimize_hydro(
                config, grid_config, solar_config, wind_config,
                hydro_config, bess_config, load_profile, pvsyst_profile, wind_profile, None)
            progress_bar.progress(60)

            # ── Find optimal ──
            status_text.text("🔍 Finding optimal solution...")
            optimal = opt_module.find_optimal_solution(results_df)
            if optimal is None:
                st.error("❌ No feasible solution found! Adjust ranges or unmet load target.")
                st.stop()
            progress_bar.progress(70)

            # ── Re-run optimal dispatch ──
            status_text.text("📈 Generating optimal dispatch profile...")
            optimal_dispatch_df = opt_module.calculate_dispatch_with_hydro(
                load_profile, pvsyst_profile, wind_profile,
                optimal['PV_kW'], optimal['Wind_kW'], optimal['Hydro_kW'],
                optimal['BESS_Power_kW'], optimal['BESS_Capacity_kWh'],
                solar_config, wind_config, hydro_config, bess_config,
                int(optimal['Hydro_Window_Start']), int(optimal['Hydro_Window_End'])
            )
            progress_bar.progress(75)

            # ── NPC & electrical metrics ──
            status_text.text("📊 Calculating financial metrics...")
            npc_data = opt_module.calculate_npc_homer_style(
                optimal['PV_kW'], optimal['Wind_kW'], optimal['Hydro_kW'],
                optimal['BESS_Power_kW'], optimal['BESS_Capacity_kWh'],
                solar_config, wind_config, hydro_config, bess_config, config,
                None, False, optimal['Total_Energy_Served_kWh']
            )
            component_capacities = {
                'pv_kw': optimal['PV_kW'], 'wind_kw': optimal['Wind_kW'],
                'hydro_kw': optimal['Hydro_kW'], 'bess_kwh': optimal['BESS_Capacity_kWh']
            }
            component_configs = {
                'bess_max_soc': bess_max_soc / 100,
                'bess_min_soc': bess_min_soc / 100,
                'bess_lifetime': bess_lifetime
            }
            electrical_metrics = opt_module.calculate_electrical_metrics(
                optimal_dispatch_df, component_capacities, component_configs,
                npc_data, project_lifetime
            )

            # ── Degradation analysis ──
            use_degradation = (apply_pv_degradation or apply_wind_degradation or
                               apply_hydro_degradation or apply_bess_degradation)

            if use_degradation:
                status_text.text("🔬 Running multi-year degradation analysis...")

                # PV
                pv_deg_type = None; pv_deg_data = None
                if apply_pv_degradation:
                    if pv_deg_method == "Simple (Annual Rate)":
                        pv_deg_type = 'simple'; pv_deg_data = pv_annual_deg_rate
                    else:
                        if pv_deg_file:
                            pv_deg_type = 'curve'
                            pv_deg_data = parse_pv_degradation_file(pv_deg_file)
                            if pv_deg_data is None: st.stop()

                # Wind
                wind_deg_type = None; wind_deg_data_param = None
                if apply_wind_degradation:
                    if wind_deg_method == "Simple (Annual Rate)":
                        wind_deg_type = 'simple'; wind_deg_data_param = wind_annual_deg_rate
                    else:
                        if wind_deg_file:
                            wind_deg_type = 'curve'
                            wind_deg_data_param = parse_wind_degradation_file(wind_deg_file)
                            if wind_deg_data_param is None: st.stop()

                # Hydro — build table from grace period + rate
                hydro_deg_table = None
                if apply_hydro_degradation and enable_hydro:
                    hydro_deg_table = opt_module.build_default_hydro_deg_table(
                        project_lifetime, hydro_stable_years, hydro_deg_rate_after
                    )

                # BESS
                bess_deg_data_param = None
                if apply_bess_degradation:
                    if bess_deg_mode == "Simple (Annual Rate)":
                        bess_deg_data_param = opt_module.build_bess_simple_degradation_data(
                            project_lifetime,
                            bess_annual_cap_deg,
                            bess_charge_eff,
                            bess_discharge_eff,
                            bess_annual_charge_eff_deg,
                            bess_annual_discharge_eff_deg
                        )
                    else:
                        if bess_deg_file:
                            bess_deg_data_param = parse_bess_degradation_file(bess_deg_file)
                            if bess_deg_data_param is None: st.stop()

                degradation_results = opt_module.run_multi_year_degradation_analysis(
                    optimal.to_dict(),
                    load_profile, pvsyst_profile, wind_profile,
                    solar_config, wind_config, hydro_config, bess_config,
                    project_lifetime=project_lifetime,
                    pv_degradation_type=pv_deg_type,
                    pv_degradation_data=pv_deg_data,
                    wind_degradation_type=wind_deg_type,
                    wind_degradation_data=wind_deg_data_param,
                    hydro_degradation_table=hydro_deg_table,
                    bess_degradation_data=bess_deg_data_param
                )
                progress_bar.progress(90)

            # ── Cash Flow LCOE (Manager's Method) ──
            status_text.text("📊 Calculating Cash Flow LCOE...")

            # Derive PV AC from optimal DC capacity
            pv_dc_mwp_opt  = optimal['PV_kW'] / 1000.0
            pv_ac_mwac_opt = pv_dc_mwp_opt / cf_dc_ac_ratio

            cf_lcoe_results = opt_module.calculate_lcoe_cashflow_method(
                pv_dc_mwp          = pv_dc_mwp_opt,
                pv_ac_mwac         = pv_ac_mwac_opt,
                wind_kw            = optimal['Wind_kW'],
                hydro_kw           = optimal['Hydro_kW'],
                bess_capacity_kwh  = optimal['BESS_Capacity_kWh'],
                annual_energy_mwh  = optimal['Total_Energy_Served_kWh'] / 1000.0,
                pv_capex_per_mwp          = cf_pv_capex_per_mwp,
                pv_fixed_om_per_mwp_yr    = cf_pv_om_per_mwp,
                pv_inverter_bop_per_mwac  = cf_pv_inv_bop_per_mwac,
                wind_capex_per_kw         = cf_wind_capex_per_kw,
                wind_om_per_kw_yr         = cf_wind_om_per_kw,
                hydro_capex_per_kw        = cf_hydro_capex_per_kw,
                hydro_om_per_kw_yr        = cf_hydro_om_per_kw,
                bess_capex_per_kwh        = cf_bess_capex_per_kwh,
                bess_om_per_kwh_yr        = cf_bess_om_per_kwh,
                nominal_discount_rate_pct = discount_rate,
                inflation_rate_pct        = inflation_rate,
                project_lifetime          = project_lifetime,
            )

            # ── Store results ──
            st.session_state.results = {
                'optimal_solution': optimal,
                'all_results': results_df,
                'optimal_dispatch': optimal_dispatch_df,
                'config': config,
                'degradation_enabled': use_degradation,
                'electrical_metrics': electrical_metrics,
                'npc_data': npc_data,
                'cf_lcoe': cf_lcoe_results,
            }
            if use_degradation:
                st.session_state.results['degradation_analysis'] = degradation_results

            st.session_state.optimization_complete = True
            progress_bar.progress(100)
            status_text.text("✅ Optimization complete!")
            st.success("🎉 Optimization completed successfully!")
            st.balloons()

        except Exception as e:
            st.error(f"❌ Optimization failed: {str(e)}")
            import traceback
            st.code(traceback.format_exc())
            st.stop()


# ==============================================================================
# RESULTS DISPLAY
# ==============================================================================

if st.session_state.optimization_complete and st.session_state.results is not None:

    results = st.session_state.results
    optimal = results['optimal_solution']
    crf     = results['npc_data']['crf']

    st.markdown("---")
    st.header("📊 Optimization Results")

    tabs = ["📊 Summary", "💰 Cost & Performance", "📈 Economic Analysis"]
    if results.get('degradation_enabled', False):
        tabs.append("🔬 Degradation")
    tab_objects = st.tabs(tabs)
    tab1 = tab_objects[0]
    tab2 = tab_objects[1]
    tab_econ = tab_objects[2]
    tab3 = tab_objects[3] if results.get('degradation_enabled', False) else None

    # ── TAB 1: SUMMARY ────────────────────────────────────────────────────────
    with tab1:
        st.subheader("Optimal System Configuration")
        col1, col2, col3, col4 = st.columns(4)
        with col1: st.metric("☀️ Solar PV", f"{optimal['PV_kW']/1000:.2f} MW", delta=f"{optimal['PV_Energy_kWh']/1000:.0f} MWh/yr")
        with col2: st.metric("💨 Wind",     f"{optimal['Wind_kW']/1000:.2f} MW", delta=f"{optimal['Wind_Energy_kWh']/1000:.0f} MWh/yr")
        with col3: st.metric("💧 Hydro",    f"{optimal['Hydro_kW']/1000:.2f} MW", delta=f"{optimal['Hydro_Energy_kWh']/1000:.0f} MWh/yr")
        with col4: st.metric("🔋 BESS",     f"{optimal['BESS_Power_kW']/1000:.2f} MW", delta=f"{optimal['BESS_Capacity_kWh']/1000:.1f} MWh")

        st.markdown("---")
        st.subheader("Key Performance Indicators")
        re_penetration = (optimal['PV_Energy_kWh'] + optimal['Wind_Energy_kWh'] + optimal['Hydro_Energy_kWh']) / optimal['Total_Load_kWh'] * 100

        col1, col2, col3, col4 = st.columns(4)
        with col1: st.metric("Net Present Cost", f"${optimal['NPC_Total']/1_000_000:.2f}M")
        with col2: st.metric("LCOE", f"${optimal['LCOE_per_kWh']*1000:.2f}/MWh")
        with col3:
            delta_label = "Target" if optimal['Unmet_Load_Percent'] <= target_unmet_percent else "Over"
            delta_color = "normal" if optimal['Unmet_Load_Percent'] <= target_unmet_percent else "inverse"
            st.metric("Unmet Load", f"{optimal['Unmet_Load_Percent']:.2f}%", delta=delta_label, delta_color=delta_color)
        with col4: st.metric("RE Penetration", f"{re_penetration:.1f}%")

        st.markdown("---")
        col1, col2 = st.columns(2)
        with col1:
            bess_annual_discharge = results['optimal_dispatch']['BESS_Discharge_wieff_kW'].sum() if results.get('optimal_dispatch') is not None else 0
            bess_npc_val = optimal.get('BESS_NPC', 0)
            lcos_val = opt_module.calculate_bess_lcos_from_npc(bess_npc_val, bess_annual_discharge, crf) if bess_annual_discharge > 0 else 0
            st.metric("🔋 BESS LCOS", f"${lcos_val*1000:.2f}/MWh")
        with col2:
            st.metric("Annual BESS Discharge", f"{bess_annual_discharge/1000:.0f} MWh")

        st.markdown("---")
        if optimal['BESS_Power_kW'] > 0:
            st.subheader("🏗️ BESS Deployment Details (Sungrow PowerTitan 2.0)")
            bess_dep = calculate_bess_deployment_sungrow(optimal['BESS_Power_kW']/1000, optimal['BESS_Capacity_kWh']/1000)
            col1, col2, col3 = st.columns(3)
            with col1:
                st.metric("Containers Required", f"{bess_dep['num_containers']}")
                st.metric("Container Model", bess_dep['container_model'])
                st.metric("MVS Units", f"{bess_dep['num_mvs_units']}")
            with col2:
                st.metric("Deployed Capacity", f"{bess_dep['actual_capacity_mwh']:.1f} MWh")
                st.metric("Deployed Power", f"{bess_dep['actual_power_mw']:.1f} MW")
                st.markdown("**Layout**"); st.write(bess_dep['layout_description'])
            with col3:
                st.metric("Total Area", f"{bess_dep['total_area_hectares']:.2f} ha ({bess_dep['total_area_acres']:.2f} acres)")
                st.metric("Power Density", f"{bess_dep['power_density_mw_per_ha']:.2f} MW/ha")
                st.metric("Energy Density", f"{bess_dep['energy_density_mwh_per_ha']:.2f} MWh/ha")

    # ── TAB 2: COST & PERFORMANCE ─────────────────────────────────────────────
    with tab2:
        st.subheader("💰 Cost Summary")
        col1, col2, col3, col4 = st.columns(4)
        with col1: st.metric("Total NPC", f"${optimal['NPC_Total']/1_000_000:.2f}M")
        with col2: st.metric("Total Capital", f"${optimal['CapEx_Total']/1_000_000:.2f}M")
        with col3: st.metric("System LCOE", f"${optimal['LCOE_per_kWh']*1000:.2f}/MWh")
        with col4:
            em = results.get('electrical_metrics', {})
            bess_lcos_em = em.get('bess', {}).get('levelized_cost_per_mwh', 0)
            st.metric("BESS LCOS", f"${bess_lcos_em:.2f}/MWh" if bess_lcos_em > 0 else "N/A")

        st.markdown("---")
        st.subheader("Net Present Cost Breakdown")

        comp_map = [('Solar PV','PV'),('Wind','Wind'),('Hydro','Hydro'),('BESS','BESS')]
        comp_names, comp_npc = [], []
        for label, key in comp_map:
            cap = optimal.get(f'{key}_kW' if key != 'BESS' else 'BESS_Power_kW', 0)
            if cap > 0:
                comp_names.append(label)
                comp_npc.append(optimal[f'{key}_NPC'] / 1e6)

        col1, col2 = st.columns(2)
        with col1:
            bar_colors = {'Solar PV':'#FDB462','Wind':'#80B1D3','Hydro':'#8DD3C7','BESS':'#FB8072'}
            fig_comp = go.Figure(data=[go.Bar(
                x=comp_names, y=comp_npc,
                marker_color=[bar_colors.get(n,'#BEBADA') for n in comp_names],
                text=[f'${v:.2f}M' for v in comp_npc], textposition='outside')])
            fig_comp.update_layout(title='NPC by Component', height=380, showlegend=False,
                plot_bgcolor='white', paper_bgcolor='white', font=dict(color='#333333'))
            fig_comp.update_yaxes(gridcolor='#EEEEEE')
            st.plotly_chart(fig_comp, use_container_width=True)

        with col2:
            cost_names  = ['Capital','Replacement','O&M','Salvage']
            cost_values = [optimal['CapEx_Total']/1e6, optimal.get('Total_Replacement',0)/1e6,
                           optimal.get('Total_OM',0)/1e6, -optimal.get('Total_Salvage',0)/1e6]
            fig_type = go.Figure(data=[go.Bar(
                x=cost_names, y=cost_values,
                marker_color=['#2E7D32','#1976D2','#F57C00','#C62828'],
                text=[f'${v:.2f}M' for v in cost_values], textposition='outside')])
            fig_type.update_layout(title='NPC by Cost Type', height=380, showlegend=False,
                plot_bgcolor='white', paper_bgcolor='white', font=dict(color='#333333'))
            fig_type.update_yaxes(gridcolor='#EEEEEE')
            st.plotly_chart(fig_type, use_container_width=True)

        st.markdown("---")
        st.subheader("Detailed Component Cost Breakdown")
        detailed = []
        for label, key in comp_map:
            cap = optimal.get('BESS_Power_kW' if key == 'BESS' else f'{key}_kW', 0)
            if cap > 0:
                cap_str = (f"{cap/1000:.2f} MW / {optimal['BESS_Capacity_kWh']/1000:.1f} MWh"
                           if key == 'BESS' else f"{cap/1000:.2f} MW")
                detailed.append({'Component': label, 'Capacity': cap_str,
                    'CapEx ($M)':       f"${optimal[f'{key}_CapEx']/1e6:.3f}",
                    'Replacement ($M)': f"${optimal.get(f'{key}_Replacement',0)/1e6:.3f}",
                    'O&M PV ($M)':      f"${optimal.get(f'{key}_OM',0)/1e6:.3f}",
                    'Salvage ($M)':     f"${optimal.get(f'{key}_Salvage',0)/1e6:.3f}",
                    'NPC ($M)':         f"${optimal[f'{key}_NPC']/1e6:.3f}"})
        if detailed:
            st.dataframe(pd.DataFrame(detailed), use_container_width=True, hide_index=True)

        st.markdown("---")
        st.subheader("💵 Nominal Cash Flow Analysis")
        proj_lt = results['config']['project_lifetime']
        years = list(range(0, proj_lt + 1))
        capital_flow = [0.0] * len(years); operating_flow = [0.0] * len(years)
        replacement_flow = [0.0] * len(years); salvage_flow = [0.0] * len(years)
        capital_flow[0] = -optimal['CapEx_Total'] / 1e6
        annual_om = optimal.get('Total_OM', 0) / 1e6 / proj_lt if proj_lt > 0 else 0
        for yr in range(1, proj_lt + 1): operating_flow[yr] = -annual_om
        lt_map = {'PV': results['config'].get('pv_lifetime',25), 'Wind': results['config'].get('wind_lifetime',25),
                  'Hydro': results['config'].get('hydro_lifetime',50), 'BESS': results['config'].get('bess_lifetime',15)}
        for label, key in comp_map:
            repl_val = optimal.get(f'{key}_Replacement', 0) / 1e6
            if repl_val > 0:
                comp_lt = lt_map.get(key, 20)
                yr = comp_lt
                while yr < proj_lt:
                    if yr < len(replacement_flow):
                        replacement_flow[yr] -= repl_val / max(1, proj_lt // comp_lt)
                    yr += comp_lt
        salvage_flow[-1] = optimal.get('Total_Salvage', 0) / 1e6

        fig_cf = go.Figure()
        fig_cf.add_trace(go.Bar(name='Capital', x=years, y=capital_flow, marker_color='#2E7D32'))
        fig_cf.add_trace(go.Bar(name='Operating', x=years, y=operating_flow, marker_color='#F57C00'))
        fig_cf.add_trace(go.Bar(name='Replacement', x=years, y=replacement_flow, marker_color='#1976D2'))
        fig_cf.add_trace(go.Bar(name='Salvage', x=years, y=salvage_flow, marker_color='#43A047'))
        fig_cf.update_layout(title='Nominal Cash Flow Over Project Lifetime',
            xaxis_title='Year', yaxis_title='Cash Flow ($M)', barmode='relative', height=420,
            plot_bgcolor='white', paper_bgcolor='white', font=dict(color='#333333'),
            legend=dict(orientation='h', yanchor='bottom', y=1.02, xanchor='right', x=1))
        st.plotly_chart(fig_cf, use_container_width=True)

        st.markdown("---")
        st.subheader("⚡ Electrical Performance Metrics")
        em = results.get('electrical_metrics', {})
        if em:
            col1, col2 = st.columns(2)
            def make_gen_table(comp_dict, label):
                return pd.DataFrame({'Metric': ['Rated Capacity','Mean Output','Capacity Factor',
                    'Total Production','Hours of Operation','Levelized Cost (LCOE)'],
                    'Value': [f"{comp_dict.get('rated_capacity_kw',0):,.1f} kW",
                        f"{comp_dict.get('mean_output_kw',0):,.1f} kW",
                        f"{comp_dict.get('capacity_factor_pct',0):.2f}%",
                        f"{comp_dict.get('total_production_kwh',0):,.0f} kWh/yr",
                        f"{comp_dict.get('hours_of_operation',0):,.0f} hrs/yr",
                        f"${comp_dict.get('levelized_cost_per_kwh',0):.4f}/kWh"]})
            with col1:
                if optimal['PV_kW'] > 0:
                    st.markdown("**☀️ Solar PV**")
                    st.dataframe(make_gen_table(em.get('pv',{}), 'PV'), use_container_width=True, hide_index=True)
                if optimal['Hydro_kW'] > 0:
                    st.markdown("**💧 Hydro**")
                    st.dataframe(make_gen_table(em.get('hydro',{}), 'Hydro'), use_container_width=True, hide_index=True)
            with col2:
                if optimal['Wind_kW'] > 0:
                    st.markdown("**💨 Wind**")
                    st.dataframe(make_gen_table(em.get('wind',{}), 'Wind'), use_container_width=True, hide_index=True)
                if optimal['BESS_Power_kW'] > 0:
                    st.markdown("**🔋 Battery Storage**")
                    bess = em.get('bess', {})
                    bess_table = pd.DataFrame({'Metric': ['Nominal Capacity','Usable Capacity','Autonomy',
                        'Energy In','Energy Out','Losses','Annual Throughput','Expected Life','Levelized Cost (LCOS)'],
                        'Value': [f"{bess.get('nominal_capacity_kwh',0):,.1f} kWh",
                            f"{bess.get('usable_capacity_kwh',0):,.1f} kWh",
                            f"{bess.get('autonomy_hours',0):.2f} hours",
                            f"{bess.get('energy_in_kwh',0):,.0f} kWh/yr",
                            f"{bess.get('energy_out_kwh',0):,.0f} kWh/yr",
                            f"{bess.get('losses_kwh',0):,.0f} kWh/yr",
                            f"{bess.get('annual_throughput_kwh',0):,.0f} kWh/yr",
                            f"{bess.get('expected_life_years',0):.0f} years",
                            f"${bess.get('levelized_cost_per_kwh',0):.4f}/kWh"]})
                    st.dataframe(bess_table, use_container_width=True, hide_index=True)

        st.markdown("---")
        st.subheader("📊 Annual Energy Production Mix")
        col1, col2 = st.columns(2)
        with col1:
            pie = create_energy_mix_pie(results)
            if pie: st.plotly_chart(pie, use_container_width=True)
        with col2:
            gen_data = []; total_gen = 0
            for src, key in [('Solar PV','PV'),('Wind','Wind'),('Hydro','Hydro')]:
                val = optimal.get(f'{key}_Energy_kWh', 0) / 1000
                if val > 0:
                    gen_data.append({'Source': src, 'Energy (MWh/yr)': f"{val:,.1f}"}); total_gen += val
            for row in gen_data:
                row['Share (%)'] = f"{float(row['Energy (MWh/yr)'].replace(',','')) / total_gen * 100:.1f}%"
            gen_data.append({'Source':'Total','Energy (MWh/yr)':f"{total_gen:,.1f}",'Share (%)':'100.0%'})
            st.dataframe(pd.DataFrame(gen_data), use_container_width=True, hide_index=True)
            st.markdown("**Energy Balance**")
            total_load = optimal['Total_Load_kWh'] / 1000
            unmet = optimal['Unmet_Load_kWh'] / 1000
            curtail = optimal.get('Total_Curtailment_kWh', 0) / 1000
            st.dataframe(pd.DataFrame({'Metric':['Annual Load','Energy Served','Unmet Load','Curtailment'],
                'Value (MWh)':[f"{total_load:,.1f}",f"{total_load-unmet:,.1f}",
                               f"{unmet:,.1f}",f"{curtail:,.1f}"]}),
                use_container_width=True, hide_index=True)

        st.markdown("---")
        st.subheader("📈 Typical Day Dispatch Profile")
        dispatch_fig = create_single_day_dispatch_profile(results)
        if dispatch_fig:
            st.plotly_chart(dispatch_fig, use_container_width=True)
            st.caption("Representative day based on median PV production day")

    # ── TAB ECON: ECONOMIC ANALYSIS (CASH FLOW LCOE) ─────────────────────────
    with tab_econ:
        st.header("📈 Economic Analysis — Cash Flow LCOE Method")
        st.caption(
            "LCOE calculated using year-by-year discounted cash flows. "
            "CAPEX at Year 0, O&M escalated annually with inflation, "
            "discounted at nominal discount rate. "
            "Compare with the HOMER-style LCOE in the Cost & Performance tab."
        )

        cf = results.get('cf_lcoe', {})
        if not cf:
            st.warning("Cash Flow LCOE data not available. Re-run optimization.")
        else:
            # ── KPI Row ──
            col1, col2, col3, col4 = st.columns(4)
            with col1:
                st.metric("CF LCOE", f"${cf['lcoe_per_mwh']:.2f}/MWh",
                    help="Cash flow method: NPV_Costs / NPV_Energy")
            with col2:
                st.metric("CF LCOE ($/kWh)", f"${cf['lcoe_per_kwh']:.5f}/kWh")
            with col3:
                st.metric("HOMER LCOE", f"${optimal['LCOE_per_kWh']*1000:.2f}/MWh",
                    help="HOMER method: NPC × CRF / Annual_Energy")
            with col4:
                diff = cf['lcoe_per_mwh'] - optimal['LCOE_per_kWh'] * 1000
                st.metric("Difference", f"${diff:+.2f}/MWh",
                    delta=f"{diff/max(optimal['LCOE_per_kWh']*1000, 0.001)*100:+.1f}%",
                    delta_color="off")

            st.markdown("---")

            # ── Key Financial Assumptions ──
            st.subheader("⚙️ Financial Assumptions Used")
            cfg = results['config']
            col1, col2, col3, col4 = st.columns(4)
            with col1: st.metric("Nominal Discount Rate", f"{cfg['discount_rate']*100:.1f}%")
            with col2: st.metric("Inflation Rate",        f"{cfg['inflation_rate']*100:.1f}%")
            with col3: st.metric("Project Lifetime",      f"{int(cfg['project_lifetime'])} years")
            with col4:
                st.metric("NPV Costs", f"${cf['npv_costs']/1e6:.2f}M")

            st.markdown("---")

            # ── CAPEX Breakdown ──
            st.subheader("🏗️ Capital Expenditure Breakdown (Year 0)")
            capex_bk = cf['capex_breakdown']
            capex_items = {k: v for k, v in capex_bk.items() if k != 'Total' and v > 0}
            col1, col2 = st.columns(2)
            with col1:
                capex_df = pd.DataFrame({
                    'Component': list(capex_items.keys()),
                    'CAPEX ($)': [f"${v:,.0f}" for v in capex_items.values()],
                    'Share (%)': [f"{v/capex_bk['Total']*100:.1f}%" for v in capex_items.values()]
                })
                capex_df.loc[len(capex_df)] = ['TOTAL', f"${capex_bk['Total']:,.0f}", '100.0%']
                st.dataframe(capex_df, use_container_width=True, hide_index=True)
            with col2:
                fig_capex = go.Figure(data=[go.Pie(
                    labels=list(capex_items.keys()),
                    values=list(capex_items.values()),
                    hole=0.4,
                    marker=dict(colors=['#FFDB5C','#FDB462','#80B1D3','#8DD3C7','#FB8072'])
                )])
                fig_capex.update_layout(title='CAPEX by Component', height=350,
                    annotations=[dict(text='CAPEX', x=0.5, y=0.5, font_size=14, showarrow=False)])
                st.plotly_chart(fig_capex, use_container_width=True)

            st.markdown("---")

            # ── Base O&M Breakdown ──
            st.subheader("🔧 Base O&M Breakdown (Year 1, before inflation escalation)")
            om_bk = cf['base_om_breakdown']
            om_items = {k: v for k, v in om_bk.items() if v > 0}
            col1, col2 = st.columns(2)
            with col1:
                om_total = cf['base_om_annual']
                om_df = pd.DataFrame({
                    'Component':   list(om_items.keys()),
                    'Base O&M ($)': [f"${v:,.0f}" for v in om_items.values()],
                    'Share (%)':   [f"{v/om_total*100:.1f}%" for v in om_items.values()]
                })
                om_df.loc[len(om_df)] = ['TOTAL', f"${om_total:,.0f}", '100.0%']
                st.dataframe(om_df, use_container_width=True, hide_index=True)
            with col2:
                yr_last = int(cfg['project_lifetime'])
                infl = cfg['inflation_rate']
                om_yr_last = om_total * (1 + infl) ** yr_last
                st.metric(f"Year 1 O&M (base)", f"${om_total*1e-6:.2f}M/yr")
                st.metric(f"Year {yr_last} O&M (inflated)", f"${om_yr_last*1e-6:.2f}M/yr",
                    delta=f"+{(om_yr_last/om_total - 1)*100:.1f}% from Year 1",
                    delta_color="off")
                inflation_multiplier = (1 + infl) ** yr_last
                st.metric("Inflation Multiplier", f"{inflation_multiplier:.3f}×",
                    help=f"(1 + {infl*100:.1f}%)^{yr_last} = {inflation_multiplier:.3f}")

            st.markdown("---")

            # ── Year-by-Year Cash Flow Chart ──
            st.subheader("📊 Year-by-Year Discounted Cash Flows")
            cf_df = cf['cashflow_df']

            fig_cf = go.Figure()
            # Discounted CAPEX bar at year 0
            fig_cf.add_trace(go.Bar(
                x=[0], y=[-cf_df.loc[0, 'Discounted_Cost'] / 1e6],
                name='Discounted CAPEX (Year 0)',
                marker_color='#C62828',
                hovertemplate='Year 0<br>CAPEX: $%{customdata:.2f}M<extra></extra>',
                customdata=[cf_df.loc[0, 'Discounted_Cost'] / 1e6]
            ))
            # Discounted O&M bars for years 1+
            om_years = cf_df[cf_df['Year'] > 0]
            fig_cf.add_trace(go.Bar(
                x=om_years['Year'],
                y=-om_years['Discounted_Cost'] / 1e6,
                name='Discounted O&M',
                marker_color='#F57C00',
                hovertemplate='Year %{x}<br>Disc. O&M: $%{customdata:.2f}M<extra></extra>',
                customdata=om_years['Discounted_Cost'] / 1e6
            ))
            # Discounted energy line on secondary axis
            fig_cf.add_trace(go.Scatter(
                x=cf_df['Year'], y=cf_df['Discounted_Energy_MWh'] / 1000,
                name='Discounted Energy (GWh)', mode='lines+markers',
                line=dict(color='#1976D2', width=2),
                yaxis='y2',
                hovertemplate='Year %{x}<br>Disc. Energy: %{y:.1f} GWh<extra></extra>'
            ))
            fig_cf.update_layout(
                title='Discounted Cash Flows and Energy by Year',
                xaxis_title='Year',
                yaxis_title='Cost ($M, negative = outflow)',
                yaxis2=dict(title='Discounted Energy (GWh)', overlaying='y',
                            side='right', showgrid=False),
                barmode='relative', height=450,
                plot_bgcolor='white', paper_bgcolor='white', font=dict(color='#333333'),
                legend=dict(orientation='h', yanchor='bottom', y=1.02, xanchor='right', x=1)
            )
            st.plotly_chart(fig_cf, use_container_width=True)

            st.markdown("---")

            # ── LCOE Comparison ──
            st.subheader("🔍 LCOE Method Comparison")
            homer_lcoe = optimal['LCOE_per_kWh'] * 1000
            cf_lcoe_val = cf['lcoe_per_mwh']
            col1, col2 = st.columns(2)
            with col1:
                comp_df = pd.DataFrame({
                    'Metric': [
                        'LCOE ($/MWh)',
                        'LCOE ($/kWh)',
                        'Discount Approach',
                        'O&M Treatment',
                        'Replacement Costs',
                        'Salvage Value',
                        'Cash Flow Basis',
                    ],
                    'HOMER Method': [
                        f"${homer_lcoe:.2f}",
                        f"${homer_lcoe/1000:.5f}",
                        'Real discount rate',
                        'Constant (real terms)',
                        'Modeled at lifetime intervals',
                        'Included (linear depreciation)',
                        'NPC × CRF / Annual Energy',
                    ],
                    'Cash Flow Method': [
                        f"${cf_lcoe_val:.2f}",
                        f"${cf_lcoe_val/1000:.5f}",
                        'Nominal discount rate',
                        'Escalates with inflation',
                        'Not modeled (CAPEX at Y0 only)',
                        'Not modeled',
                        'NPV(Costs) / NPV(Energy)',
                    ],
                })
                st.dataframe(comp_df, use_container_width=True, hide_index=True)
            with col2:
                fig_comp = go.Figure(data=[go.Bar(
                    x=['HOMER Method', 'Cash Flow Method'],
                    y=[homer_lcoe, cf_lcoe_val],
                    marker_color=['#1976D2', '#2E7D32'],
                    text=[f'${homer_lcoe:.2f}', f'${cf_lcoe_val:.2f}'],
                    textposition='outside',
                )])
                fig_comp.update_layout(
                    title='LCOE Comparison ($/MWh)',
                    yaxis_title='LCOE ($/MWh)', height=380,
                    showlegend=False,
                    plot_bgcolor='white', paper_bgcolor='white', font=dict(color='#333333')
                )
                fig_comp.update_yaxes(gridcolor='#EEEEEE')
                st.plotly_chart(fig_comp, use_container_width=True)

            st.markdown("---")

            # ── Full Cash Flow Table ──
            st.subheader("📋 Full Year-by-Year Cash Flow Table")
            display_cf = cf_df.copy()
            display_cf['Total_Cost'] = display_cf['Total_Cost'].map('${:,.0f}'.format)
            display_cf['Discounted_Cost'] = display_cf['Discounted_Cost'].map('${:,.0f}'.format)
            display_cf['Annual_Energy_MWh'] = display_cf['Annual_Energy_MWh'].map('{:,.1f}'.format)
            display_cf['Discounted_Energy_MWh'] = display_cf['Discounted_Energy_MWh'].map('{:,.1f}'.format)
            display_cf['OM_PV']    = display_cf['OM_PV'].map('${:,.0f}'.format)
            display_cf['OM_Wind']  = display_cf['OM_Wind'].map('${:,.0f}'.format)
            display_cf['OM_Hydro'] = display_cf['OM_Hydro'].map('${:,.0f}'.format)
            display_cf['OM_BESS']  = display_cf['OM_BESS'].map('${:,.0f}'.format)
            st.dataframe(display_cf, use_container_width=True, hide_index=True)

            # Summary row
            col1, col2, col3 = st.columns(3)
            with col1: st.metric("NPV of Costs",  f"${cf['npv_costs']:,.0f}")
            with col2: st.metric("NPV of Energy", f"{cf['npv_energy_mwh']:,.1f} MWh")
            with col3: st.metric("LCOE",          f"${cf['lcoe_per_mwh']:.4f}/MWh")

    # ── TAB 3: DEGRADATION ────────────────────────────────────────────────────
    if tab3 is not None:
        with tab3:
            st.header("🔬 Multi-Year Degradation Analysis")
            deg_data    = results['degradation_analysis']
            yearly_df   = deg_data['yearly_metrics']
            deg_summary = deg_data['degradation_summary']
            proj_lt_deg = results['config']['project_lifetime']

            # ── Summary KPIs ──
            col1, col2, col3, col4 = st.columns(4)
            with col1:
                pv_yr_last = deg_summary.get('pv_degradation_year_last', deg_summary.get('pv_degradation_year_25', 0))
                st.metric(f"PV Degradation @ Year {proj_lt_deg}", f"{pv_yr_last:.2f}%")
            with col2:
                wind_yr_last = deg_summary.get('wind_degradation_year_last', deg_summary.get('wind_degradation_year_25', 0))
                st.metric(f"Wind Degradation @ Year {proj_lt_deg}", f"{wind_yr_last:.2f}%")
            with col3:
                hydro_yr_last = deg_summary.get('hydro_degradation_year_last', deg_summary.get('hydro_degradation_year_25', 0))
                st.metric(f"Hydro Degradation @ Year {proj_lt_deg}", f"{hydro_yr_last:.2f}%")
            with col4:
                bess_ret = deg_summary.get('bess_retention_year_last', deg_summary.get('bess_retention_year_25', 100))
                st.metric(f"BESS Retention @ Year {proj_lt_deg}", f"{bess_ret:.1f}%")

            col1, col2, col3 = st.columns(3)
            with col1: st.metric("Avg Unmet Load", f"{deg_summary['avg_unmet_pct']:.2f}%")
            with col2: st.metric("Max Unmet Load", f"{deg_summary['max_unmet_pct']:.2f}%")
            with col3: st.metric("Total Energy Served", f"{deg_summary['total_energy_served_25yr_GWh']:.1f} GWh")

            st.markdown("---")

            # ── Component Degradation Chart ──
            st.subheader("📈 Component Degradation Over Project Life")
            fig_deg = go.Figure()

            if 'PV_Degradation_%' in yearly_df.columns and yearly_df['PV_Degradation_%'].max() > 0:
                fig_deg.add_trace(go.Scatter(x=yearly_df['Year'], y=yearly_df['PV_Degradation_%'],
                    name='PV Degradation (%)', mode='lines+markers',
                    line=dict(color='#FFDB5C', width=2), marker=dict(size=5)))

            if 'Wind_Degradation_%' in yearly_df.columns and yearly_df['Wind_Degradation_%'].max() > 0:
                fig_deg.add_trace(go.Scatter(x=yearly_df['Year'], y=yearly_df['Wind_Degradation_%'],
                    name='Wind Degradation (%)', mode='lines+markers',
                    line=dict(color='#80B1D3', width=2), marker=dict(size=5)))

            if 'Hydro_Degradation_%' in yearly_df.columns and yearly_df['Hydro_Degradation_%'].max() > 0:
                fig_deg.add_trace(go.Scatter(x=yearly_df['Year'], y=yearly_df['Hydro_Degradation_%'],
                    name='Hydro Degradation (%)', mode='lines+markers',
                    line=dict(color='#8DD3C7', width=2), marker=dict(size=5)))

            if 'BESS_Retention_%' in yearly_df.columns:
                bess_loss = 100 - yearly_df['BESS_Retention_%']
                if bess_loss.max() > 0:
                    fig_deg.add_trace(go.Scatter(x=yearly_df['Year'], y=bess_loss,
                        name='BESS Capacity Loss (%)', mode='lines+markers',
                        line=dict(color='#FB8072', width=2, dash='dash'), marker=dict(size=5)))

            fig_deg.update_layout(title='Cumulative Degradation by Component',
                xaxis_title='Project Year', yaxis_title='Cumulative Degradation (%)',
                hovermode='x unified', height=420,
                legend=dict(orientation='h', yanchor='bottom', y=1.02, xanchor='right', x=1))
            st.plotly_chart(fig_deg, use_container_width=True)

            st.markdown("---")

            # ── Unmet Load Over Time ──
            st.subheader("⚠️ Unmet Load Profile Over Project Life")
            fig_unmet = go.Figure()
            fig_unmet.add_trace(go.Scatter(x=yearly_df['Year'], y=yearly_df['Unmet_%'],
                name='Unmet Load %', mode='lines+markers',
                line=dict(color='#E63946', width=2), marker=dict(size=6),
                fill='tozeroy', fillcolor='rgba(230,57,70,0.1)'))
            fig_unmet.add_hline(y=target_unmet_percent, line_dash='dash',
                line_color='orange', annotation_text=f'Target ({target_unmet_percent}%)')
            fig_unmet.update_layout(xaxis_title='Project Year', yaxis_title='Unmet Load (%)',
                hovermode='x unified', height=380)
            st.plotly_chart(fig_unmet, use_container_width=True)

            st.markdown("---")

            # ── Annual Energy Generation ──
            st.subheader("⚡ Annual Energy Generation by Source")
            fig_gen = go.Figure()
            if 'PV_Energy_MWh' in yearly_df.columns:
                fig_gen.add_trace(go.Scatter(x=yearly_df['Year'], y=yearly_df['PV_Energy_MWh'],
                    name='PV', mode='lines', line=dict(color='#FFDB5C', width=2)))
            if 'Wind_Energy_MWh' in yearly_df.columns:
                fig_gen.add_trace(go.Scatter(x=yearly_df['Year'], y=yearly_df['Wind_Energy_MWh'],
                    name='Wind', mode='lines', line=dict(color='#80B1D3', width=2)))
            if 'Hydro_Energy_MWh' in yearly_df.columns:
                fig_gen.add_trace(go.Scatter(x=yearly_df['Year'], y=yearly_df['Hydro_Energy_MWh'],
                    name='Hydro', mode='lines', line=dict(color='#8DD3C7', width=2)))
            fig_gen.update_layout(xaxis_title='Project Year', yaxis_title='Energy (MWh)',
                hovermode='x unified', height=400,
                legend=dict(orientation='h', yanchor='bottom', y=1.02, xanchor='right', x=1))
            st.plotly_chart(fig_gen, use_container_width=True)

            st.markdown("---")

            # ── Full Metrics Table ──
            st.subheader("📊 Annual Performance Metrics Table")
            display_cols = ['Year']
            if 'PV_Degradation_%'   in yearly_df.columns: display_cols.append('PV_Degradation_%')
            if 'Wind_Degradation_%' in yearly_df.columns: display_cols.append('Wind_Degradation_%')
            if 'Hydro_Degradation_%'in yearly_df.columns: display_cols.append('Hydro_Degradation_%')
            if 'BESS_Retention_%'   in yearly_df.columns: display_cols.append('BESS_Retention_%')
            display_cols += ['PV_Energy_MWh','Wind_Energy_MWh','Hydro_Energy_MWh',
                             'Load_MWh','Served_MWh','Unmet_MWh','Unmet_%',
                             'BESS_Throughput_MWh','Curtailment_MWh']
            display_cols = [c for c in display_cols if c in yearly_df.columns]
            st.dataframe(yearly_df[display_cols].round(2), use_container_width=True, hide_index=True)

    # ── EXPORT ────────────────────────────────────────────────────────────────
    st.markdown("---")
    st.header("📥 Export Results")
    try:
        excel_buffer = build_excel_export(results, optimal, opt_module)
        st.download_button(
            label="📥 Download Full Results (Excel — All Sheets)",
            data=excel_buffer.getvalue(),
            file_name=f"optimization_results_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True, type="primary")

        sheets = ["✅ Optimal Configuration & KPIs", "✅ All Combinations"]
        if results.get('optimal_dispatch') is not None:
            sheets.append("✅ Dispatch Year 1 (Optimal)")
        if results.get('degradation_enabled') and 'degradation_analysis' in results:
            sheets.append("✅ Degradation Summary")
            for yk in results['degradation_analysis']['selected_year_dispatch'].keys():
                sheets.append(f"✅ Dispatch Year {yk.split('_')[1]}")
        st.markdown("**Excel file contains:**")
        for s in sheets: st.markdown(f"- {s}")
    except Exception as e:
        st.error(f"❌ Excel export failed: {str(e)}")
        import traceback
        st.code(traceback.format_exc())

else:
    st.info("👆 Configure system parameters in the sidebar and click 'RUN OPTIMIZATION' to begin.")


# ==============================================================================
# FOOTER
# ==============================================================================
st.markdown("---")
st.markdown("""
<div style="text-align:center;color:#666;">
    <p><strong>Energy Modeling Optimizer v5.1</strong> | Professional Hybrid System Design Tool</p>
    <p>Developed by SJ | March 2026</p>
</div>
""", unsafe_allow_html=True)
