"""
RENEWABLE ENERGY OPTIMIZATION TOOL - VERSION 5.0 WITH DEGRADATION
==================================================================
Complete version with user-configurable degradation analysis

NEW IN V5.0:
- User-configurable PV degradation (simple annual rate OR custom curve upload)
- User-configurable BESS degradation (3 presets OR custom curve upload)
- 25-year multi-year simulation with degradation
- Multi-year hourly dispatch exports (Years 1,2,5,10,15,20,25)
- Professional SJ branding
- Direct Python architecture (no Excel intermediary)
- Single Excel export with multiple sheets

FIXES IN THIS VERSION:
- BESS SOC unit bug fixed (min_soc/max_soc now correctly divided by 100)
- Export Results moved inside optimization_complete block (fixes IndexError)
- Broken try/except/else structure fixed
- Layout description truncation fixed (st.write instead of st.metric)
- LCOS calculation and display added
- Single Excel download replaces multiple CSV buttons
- Optimal dispatch re-run after optimization for dispatch profile chart

Author: SJ
Date: March 2026
"""

import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
from plotly.subplots import make_subplots
from datetime import datetime
from io import BytesIO

# Import optimization module WITH DEGRADATION
try:
    import optimize_gridsearch_hydro_WITH_DEGRADATION as opt_module
    OPTIMIZATION_AVAILABLE = True
except ImportError:
    try:
        import optimize_gridsearch_hydro_CLEAN as opt_module
        OPTIMIZATION_AVAILABLE = True
        st.warning("⚠️ Using base optimization module (no degradation support)")
    except ImportError:
        OPTIMIZATION_AVAILABLE = False
        st.error("❌ Optimization module not found")


# ==============================================================================
# SUNGROW BESS DEPLOYMENT CALCULATION
# ==============================================================================

def calculate_bess_deployment_sungrow(bess_power_mw, bess_capacity_mwh):
    """Calculate BESS deployment using REAL Sungrow PowerTitan 2.0 specifications."""
    import math

    container_capacity_mwh = 10.0
    container_power_mw = 5.0
    container_length_m = 6.058
    container_width_m = 2.438

    back_to_back_spacing_m = 0.150
    adjacent_spacing_m = 1.500
    mvs_spacing_m = 3.500
    mvs_width_m = 2.000
    perimeter_clearance_m = 5.000

    num_containers_energy = math.ceil(bess_capacity_mwh / container_capacity_mwh)
    num_containers_power = math.ceil(bess_power_mw / container_power_mw)
    num_containers = max(num_containers_energy, num_containers_power)

    actual_capacity_mwh = num_containers * container_capacity_mwh
    actual_power_mw = num_containers * container_power_mw
    num_mvs_units = math.ceil(num_containers / 2)

    if num_containers <= 2:
        total_length_m = container_length_m + (2 * perimeter_clearance_m)
        container_section_width = (container_width_m + back_to_back_spacing_m + container_width_m)
        mvs_section_width = mvs_spacing_m + mvs_width_m
        total_width_m = container_section_width + mvs_section_width + (2 * perimeter_clearance_m)
        layout_desc = "1 section (2 containers back-to-back + 1 MVS unit)"
    else:
        num_sections = math.ceil(num_containers / 2)
        section_length = container_length_m + adjacent_spacing_m
        total_length_m = (num_sections * section_length - adjacent_spacing_m + 2 * perimeter_clearance_m)
        container_section_width = (container_width_m + back_to_back_spacing_m + container_width_m)
        mvs_section_width = mvs_spacing_m + mvs_width_m
        total_width_m = container_section_width + mvs_section_width + (2 * perimeter_clearance_m)
        layout_desc = f"{num_sections} sections side-by-side ({num_containers} containers + {num_mvs_units} MVS units)"

    total_area_m2 = total_length_m * total_width_m
    total_area_hectares = total_area_m2 / 10000
    total_area_acres = total_area_hectares * 2.471
    container_footprint_m2 = num_containers * (container_length_m * container_width_m)
    power_density_mw_per_ha = actual_power_mw / total_area_hectares if total_area_hectares > 0 else 0
    energy_density_mwh_per_ha = actual_capacity_mwh / total_area_hectares if total_area_hectares > 0 else 0

    return {
        'num_containers': num_containers,
        'container_model': 'PowerTitan 2.0',
        'container_capacity_mwh': container_capacity_mwh,
        'container_power_mw': container_power_mw,
        'actual_capacity_mwh': actual_capacity_mwh,
        'actual_power_mw': actual_power_mw,
        'num_mvs_units': num_mvs_units,
        'total_length_m': total_length_m,
        'total_width_m': total_width_m,
        'total_area_m2': total_area_m2,
        'total_area_hectares': total_area_hectares,
        'total_area_acres': total_area_acres,
        'container_footprint_m2': container_footprint_m2,
        'power_density_mw_per_ha': power_density_mw_per_ha,
        'energy_density_mwh_per_ha': energy_density_mwh_per_ha,
        'layout_description': layout_desc
    }


# ==============================================================================
# DEGRADATION FILE PARSING HELPERS
# ==============================================================================

def parse_pv_degradation_file(uploaded_file):
    """Parse uploaded PV degradation curve file."""
    try:
        if uploaded_file.name.endswith('.csv'):
            df = pd.read_csv(uploaded_file)
        else:
            df = pd.read_excel(uploaded_file)

        if 'Year' not in df.columns or 'PV_Degradation_%' not in df.columns:
            st.error("❌ Invalid file format. Required columns: Year, PV_Degradation_%")
            return None

        deg_curve = dict(zip(df['Year'].astype(int), df['PV_Degradation_%']))

        if len(deg_curve) < 25:
            st.warning(f"⚠️ File contains only {len(deg_curve)} years. Should have 25 years.")

        st.success(f"✓ Loaded PV degradation curve: {len(deg_curve)} years")
        return deg_curve

    except Exception as e:
        st.error(f"❌ Error parsing PV degradation file: {str(e)}")
        return None


def parse_bess_degradation_file(uploaded_file):
    """Parse uploaded BESS degradation curve file."""
    try:
        if uploaded_file.name.endswith('.csv'):
            df = pd.read_csv(uploaded_file)
        else:
            df = pd.read_excel(uploaded_file)

        required_cols = ['Year', 'Capacity_Retention_%', 'Charging_Efficiency_%', 'Discharging_Efficiency_%']
        if not all(col in df.columns for col in required_cols):
            st.error(f"❌ Invalid file format. Required columns: {', '.join(required_cols)}")
            return None

        deg_data = {}
        for _, row in df.iterrows():
            year = int(row['Year'])
            deg_data[year] = {
                'capacity': row['Capacity_Retention_%'],
                'charge_eff': row['Charging_Efficiency_%'],
                'discharge_eff': row['Discharging_Efficiency_%']
            }

        if len(deg_data) < 25:
            st.warning(f"⚠️ File contains only {len(deg_data)} years. Should have 25 years.")

        st.success(f"✓ Loaded BESS degradation curve: {len(deg_data)} years")
        return deg_data

    except Exception as e:
        st.error(f"❌ Error parsing BESS degradation file: {str(e)}")
        return None


def create_pv_degradation_template():
    """Create PV degradation template CSV."""
    years = list(range(1, 26))
    degradation = [0.4 * (year - 1) for year in years]
    df = pd.DataFrame({'Year': years, 'PV_Degradation_%': degradation})
    return df.to_csv(index=False).encode('utf-8')


def create_bess_degradation_template():
    """Create BESS degradation template CSV."""
    years = list(range(1, 26))
    capacity = [100 - (0.5 * (year - 1)) for year in years]
    charge_eff = [90.0] * 25
    discharge_eff = [98.5] * 25
    df = pd.DataFrame({
        'Year': years,
        'Capacity_Retention_%': capacity,
        'Charging_Efficiency_%': charge_eff,
        'Discharging_Efficiency_%': discharge_eff
    })
    return df.to_csv(index=False).encode('utf-8')


# ==============================================================================
# VISUALIZATION FUNCTIONS
# ==============================================================================

def create_single_day_dispatch_profile(results):
    """Create single median-PV-day dispatch profile chart."""
    if 'optimal_dispatch' not in results or results['optimal_dispatch'] is None:
        return None

    dispatch_df = results['optimal_dispatch'].copy()

    # Handle both old and new hour formats
    if 'Hour_of_Day' not in dispatch_df.columns:
        if dispatch_df['Hour'].max() <= 23:
            dispatch_df['Absolute_Hour'] = dispatch_df.index
            dispatch_df['Hour_of_Day'] = dispatch_df['Hour']
        else:
            dispatch_df['Absolute_Hour'] = dispatch_df['Hour']
            dispatch_df['Hour_of_Day'] = dispatch_df['Hour'] % 24
    else:
        if 'Hour' in dispatch_df.columns and dispatch_df['Hour'].max() > 24:
            dispatch_df['Absolute_Hour'] = dispatch_df['Hour']
        else:
            dispatch_df['Absolute_Hour'] = dispatch_df.index

    dispatch_df['Day'] = dispatch_df['Absolute_Hour'] // 24

    # Find median PV day
    pv_col = 'PV_Available_kW' if 'PV_Available_kW' in dispatch_df.columns else 'PV_Output_kW'
    daily_pv = dispatch_df.groupby('Day')[pv_col].sum()
    median_pv_day = daily_pv.sort_values().index[len(daily_pv) // 2]

    # Extract that day
    start_hour = median_pv_day * 24
    end_hour = start_hour + 24
    day_profile = dispatch_df[
        (dispatch_df['Absolute_Hour'] >= start_hour) &
        (dispatch_df['Absolute_Hour'] < end_hour)
    ].copy()

    fig = make_subplots(specs=[[{"secondary_y": True}]])

    # Generation components (overlapping from zero)
    fig.add_trace(go.Scatter(
        x=day_profile['Hour_of_Day'],
        y=day_profile['Hydro_Output_kW'] / 1000,
        name='Hydro',
        fill='tozeroy',
        fillcolor='rgba(141, 211, 199, 0.6)',
        line=dict(width=0.5, color='rgba(141, 211, 199, 1)'),
        hovertemplate='Hour %{x}<br>Hydro: %{y:.2f} MW<extra></extra>'
    ), secondary_y=False)

    fig.add_trace(go.Scatter(
        x=day_profile['Hour_of_Day'],
        y=day_profile[pv_col] / 1000,
        name='PV',
        fill='tozeroy',
        fillcolor='rgba(255, 219, 92, 0.6)',
        line=dict(width=0.5, color='rgba(255, 219, 92, 1)'),
        hovertemplate='Hour %{x}<br>PV: %{y:.2f} MW<extra></extra>'
    ), secondary_y=False)

    if 'Wind_Output_kW' in day_profile.columns:
        fig.add_trace(go.Scatter(
            x=day_profile['Hour_of_Day'],
            y=day_profile['Wind_Output_kW'] / 1000,
            name='Wind',
            fill='tozeroy',
            fillcolor='rgba(179, 226, 205, 0.6)',
            line=dict(width=0.5, color='rgba(179, 226, 205, 1)'),
            hovertemplate='Hour %{x}<br>Wind: %{y:.2f} MW<extra></extra>'
        ), secondary_y=False)

    # Load line
    fig.add_trace(go.Scatter(
        x=day_profile['Hour_of_Day'],
        y=day_profile['Load_kW'] / 1000,
        name='Load',
        mode='lines',
        line=dict(color='red', width=2),
        hovertemplate='Hour %{x}<br>Load: %{y:.2f} MW<extra></extra>'
    ), secondary_y=False)

    # BESS SOC on secondary axis
    if 'BESS_SOC_pct' in day_profile.columns:
        fig.add_trace(go.Scatter(
            x=day_profile['Hour_of_Day'],
            y=day_profile['BESS_SOC_pct'],
            name='BESS SOC',
            mode='lines',
            line=dict(color='purple', width=2, dash='dash'),
            hovertemplate='Hour %{x}<br>SOC: %{y:.1f}%<extra></extra>'
        ), secondary_y=True)

    fig.update_xaxes(title_text="Hour of Day", range=[0, 23])
    fig.update_yaxes(title_text="Power (MW)", secondary_y=False)
    fig.update_yaxes(title_text="BESS SOC (%)", secondary_y=True, range=[0, 100])

    fig.update_layout(
        title=f'Typical Day Dispatch Profile (Day {median_pv_day + 1} - Median PV)',
        hovermode='x unified',
        height=500,
        showlegend=True,
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1)
    )

    return fig


def create_energy_mix_pie(results):
    """Create energy mix pie chart."""
    optimal = results['optimal_solution']

    values, labels, colors = [], [], []

    if optimal['PV_Energy_kWh'] > 0:
        values.append(optimal['PV_Energy_kWh'] / 1000)
        labels.append('Solar PV')
        colors.append('#FFDB5C')

    if optimal['Wind_Energy_kWh'] > 0:
        values.append(optimal['Wind_Energy_kWh'] / 1000)
        labels.append('Wind')
        colors.append('#B3E2CD')

    if optimal['Hydro_Energy_kWh'] > 0:
        values.append(optimal['Hydro_Energy_kWh'] / 1000)
        labels.append('Hydro')
        colors.append('#8DD3C7')

    if not values:
        return None

    fig = go.Figure(data=[go.Pie(
        labels=labels,
        values=values,
        hole=0.4,
        marker=dict(colors=colors)
    )])

    fig.update_layout(
        title='Annual Energy Mix',
        height=400,
        showlegend=True,
        annotations=[dict(text='MWh', x=0.5, y=0.5, font_size=20, showarrow=False)]
    )

    return fig


def build_excel_export(results, optimal, opt_module):
    """Build multi-sheet Excel export and return BytesIO buffer."""
    output = BytesIO()

    # Calculate LCOS
    bess_annual_discharge = 0
    lcos = 0
    if results.get('optimal_dispatch') is not None:
        bess_annual_discharge = results['optimal_dispatch']['BESS_Discharge_wieff_kW'].sum()
    bess_npc_val = optimal.get('BESS_NPC', 0)
    if bess_annual_discharge > 0 and hasattr(opt_module, 'calculate_bess_lcos_from_npc'):
        lcos = opt_module.calculate_bess_lcos_from_npc(
            bess_npc_val,
            bess_annual_discharge,
            results['config']['project_lifetime']
        )

    re_penetration = (
        optimal['PV_Energy_kWh'] + optimal['Wind_Energy_kWh'] + optimal['Hydro_Energy_kWh']
    ) / optimal['Total_Load_kWh'] * 100 if optimal['Total_Load_kWh'] > 0 else 0

    with pd.ExcelWriter(output, engine='openpyxl') as writer:

        # ── Sheet 1: Optimal Configuration & KPIs ──
        summary_data = {
            'Parameter': [
                'PV Capacity (MW)', 'Wind Capacity (MW)', 'Hydro Capacity (MW)',
                'BESS Power (MW)', 'BESS Energy (MWh)',
                'Net Present Cost ($M)', 'LCOE ($/MWh)', 'LCOS ($/MWh)',
                'Unmet Load (%)', 'RE Penetration (%)',
                'Annual PV Energy (MWh)', 'Annual Wind Energy (MWh)',
                'Annual Hydro Energy (MWh)', 'Annual Load (MWh)',
                'Total Capital ($M)', 'Annual O&M ($k)',
                'Annual BESS Discharge (MWh)'
            ],
            'Value': [
                round(optimal['PV_kW'] / 1000, 3),
                round(optimal['Wind_kW'] / 1000, 3),
                round(optimal['Hydro_kW'] / 1000, 3),
                round(optimal['BESS_Power_kW'] / 1000, 3),
                round(optimal['BESS_Capacity_kWh'] / 1000, 3),
                round(optimal['NPC_Total'] / 1_000_000, 4),
                round(optimal['LCOE_per_kWh'] * 1000, 4),
                round(lcos * 1000, 4),
                round(optimal['Unmet_Load_Percent'], 4),
                round(re_penetration, 2),
                round(optimal['PV_Energy_kWh'] / 1000, 1),
                round(optimal['Wind_Energy_kWh'] / 1000, 1),
                round(optimal['Hydro_Energy_kWh'] / 1000, 1),
                round(optimal['Total_Load_kWh'] / 1000, 1),
                round(optimal['CapEx_Total'] / 1_000_000, 4),
                round(optimal['OpEx_Annual'] / 1000, 2),
                round(bess_annual_discharge / 1000, 1)
            ]
        }
        pd.DataFrame(summary_data).to_excel(writer, sheet_name='Optimal Configuration', index=False)

        # ── Sheet 2: All Combinations ──
        if 'all_results' in results:
            results['all_results'].to_excel(writer, sheet_name='All Combinations', index=False)

        # ── Sheet 3: Hourly Dispatch (Optimal Year 1) ──
        if results.get('optimal_dispatch') is not None:
            results['optimal_dispatch'].to_excel(writer, sheet_name='Dispatch Year 1', index=False)

        # ── Sheets 4+: Degradation year dispatches ──
        if results.get('degradation_enabled') and 'degradation_analysis' in results:
            deg_data = results['degradation_analysis']
            deg_data['yearly_metrics'].to_excel(writer, sheet_name='Degradation Summary', index=False)

            for year_key, dispatch_df in deg_data['selected_year_dispatch'].items():
                year_num = year_key.split('_')[1]
                sheet_name = f'Dispatch Year {year_num}'
                dispatch_df.to_excel(writer, sheet_name=sheet_name, index=False)

    output.seek(0)
    return output


# ==============================================================================
# PAGE CONFIGURATION
# ==============================================================================

st.set_page_config(
    page_title="Energy Modeling Optimizer v5.0",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="expanded"
)

# SJ logo + title
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

# Initialize session state
if 'optimization_complete' not in st.session_state:
    st.session_state.optimization_complete = False
if 'results' not in st.session_state:
    st.session_state.results = None


# ==============================================================================
# SIDEBAR - SYSTEM CONFIGURATION
# ==============================================================================

with st.sidebar:
    st.header("⚙️ System Configuration")

    st.subheader("🔌 Component Selection")
    st.markdown("**Select components to include:**")

    col1, col2 = st.columns(2)
    with col1:
        enable_pv = st.checkbox("☀️ Solar PV", value=True, key="enable_pv")
        enable_wind = st.checkbox("💨 Wind", value=True, key="enable_wind")
        enable_hydro = st.checkbox("💧 Hydro", value=True, key="enable_hydro")
    with col2:
        enable_bess = st.checkbox("🔋 BESS", value=True, key="enable_bess")

    if not any([enable_pv, enable_wind, enable_hydro, enable_bess]):
        st.error("⚠️ At least one component must be enabled!")

    st.markdown("---")

    # ── SOLAR PV ──
    with st.expander("☀️ SOLAR PV", expanded=enable_pv):
        if not enable_pv:
            st.warning("⚠️ Solar PV is DISABLED")
            pv_min = 0.0; pv_max = 0.0; pv_step = 1.0
            pv_capex = 1000; pv_opex = 10; pv_lifetime = 25
            apply_pv_degradation = False
            pv_deg_method = None; pv_annual_deg_rate = None; pv_deg_file = None
        else:
            st.subheader("Capacity Range")
            col1, col2 = st.columns(2)
            with col1:
                pv_min = st.number_input("Min (MW)", value=1.0, min_value=0.0, step=0.5, key="pv_min")
            with col2:
                pv_max = st.number_input("Max (MW)", value=5.0, min_value=0.0, step=0.5, key="pv_max")
            pv_step = st.number_input("Step (MW)", value=1.0, min_value=0.1, step=0.1, key="pv_step")

            st.subheader("Financial Parameters")
            col1, col2 = st.columns(2)
            with col1:
                pv_capex = st.number_input("CapEx ($/kW)", value=1000, step=10, key="pv_capex")
                pv_opex = st.number_input("OpEx ($/kW/yr)", value=10, step=1, key="pv_opex")
            with col2:
                pv_lifetime = st.number_input("Lifetime (years)", value=25, step=1, key="pv_life")

            st.markdown("---")
            st.subheader("🔬 Advanced Analysis")

            apply_pv_degradation = st.checkbox(
                "Apply PV Degradation Analysis", value=False,
                help="Enable 25-year PV degradation modeling", key="apply_pv_degradation"
            )

            if apply_pv_degradation:
                pv_deg_method = st.radio(
                    "Degradation Method:",
                    ["Simple (Annual Rate)", "Advanced (Upload Curve)"],
                    horizontal=True, key="pv_deg_method"
                )

                if pv_deg_method == "Simple (Annual Rate)":
                    pv_annual_deg_rate = st.number_input(
                        "Annual Degradation Rate (%/year)", value=0.40,
                        min_value=0.0, max_value=2.0, step=0.05,
                        help="Typical range: 0.3-0.8%/year for crystalline silicon PV",
                        key="pv_annual_deg"
                    )
                    year_25_preview = (1 - (1 - pv_annual_deg_rate / 100) ** 25) * 100
                    st.info(f"ℹ️ Typical: 0.3-0.8%/year\n\n📊 Preview: ~{year_25_preview:.2f}% degradation at Year 25")
                    pv_deg_file = None
                    st.download_button(
                        label="📥 Download Template CSV",
                        data=create_pv_degradation_template(),
                        file_name="pv_degradation_template.csv",
                        mime="text/csv", key="pv_template_download"
                    )
                else:
                    pv_deg_file = st.file_uploader(
                        "Upload PV Degradation Curve", type=['csv', 'xlsx'],
                        help="Required columns: Year, PV_Degradation_%", key="pv_deg_file"
                    )
                    pv_annual_deg_rate = None
                    if pv_deg_file:
                        st.success(f"✓ Uploaded: {pv_deg_file.name}")
                    st.download_button(
                        label="📥 Download Template CSV",
                        data=create_pv_degradation_template(),
                        file_name="pv_degradation_template.csv",
                        mime="text/csv", key="pv_template_download_adv"
                    )
            else:
                pv_deg_method = None; pv_annual_deg_rate = None; pv_deg_file = None

    # ── WIND ──
    with st.expander("💨 WIND"):
        if not enable_wind:
            st.warning("⚠️ Wind is DISABLED")
            wind_min = 0.0; wind_max = 0.0; wind_step = 1.0
            wind_capex = 1200; wind_opex = 15; wind_lifetime = 25
        else:
            st.subheader("Capacity Range")
            col1, col2 = st.columns(2)
            with col1:
                wind_min = st.number_input("Min (MW)", value=1.0, min_value=0.0, step=0.5, key="wind_min")
            with col2:
                wind_max = st.number_input("Max (MW)", value=5.0, min_value=0.0, step=0.5, key="wind_max")
            wind_step = st.number_input("Step (MW)", value=1.0, min_value=0.1, step=0.1, key="wind_step")

            st.subheader("Financial Parameters")
            col1, col2 = st.columns(2)
            with col1:
                wind_capex = st.number_input("CapEx ($/kW)", value=1200, step=10, key="wind_capex")
                wind_opex = st.number_input("OpEx ($/kW/yr)", value=15, step=1, key="wind_opex")
            with col2:
                wind_lifetime = st.number_input("Lifetime (years)", value=25, step=1, key="wind_life")

    # ── HYDRO ──
    with st.expander("💧 HYDRO"):
        if not enable_hydro:
            st.warning("⚠️ Hydro is DISABLED")
            hydro_min = 0.0; hydro_max = 0.0; hydro_step = 1.0
            hydro_hours_per_day = 6
            hydro_capex = 1500; hydro_opex = 20; hydro_lifetime = 50
        else:
            st.subheader("Capacity Range")
            col1, col2 = st.columns(2)
            with col1:
                hydro_min = st.number_input("Min (MW)", value=1.0, min_value=0.0, step=0.5, key="hydro_min")
            with col2:
                hydro_max = st.number_input("Max (MW)", value=5.0, min_value=0.0, step=0.5, key="hydro_max")
            hydro_step = st.number_input("Step (MW)", value=1.0, min_value=0.1, step=0.1, key="hydro_step")

            st.subheader("Operating Configuration")
            hydro_hours_per_day = st.number_input(
                "Operating Hours/Day", value=6, min_value=1, max_value=24, step=1,
                key="hydro_hours",
                help="Target hours per day for hydro operation (optimizer will find optimal window)"
            )

            st.subheader("Financial Parameters")
            col1, col2 = st.columns(2)
            with col1:
                hydro_capex = st.number_input("CapEx ($/kW)", value=1500, step=10, key="hydro_capex")
                hydro_opex = st.number_input("OpEx ($/kW/yr)", value=20, step=1, key="hydro_opex")
            with col2:
                hydro_lifetime = st.number_input("Lifetime (years)", value=50, step=1, key="hydro_life")

    # ── BESS ──
    with st.expander("🔋 BATTERY STORAGE"):
        if not enable_bess:
            st.warning("⚠️ BESS is DISABLED")
            bess_min = 0.0; bess_max = 0.0; bess_step = 1.0
            bess_duration = 4.0; bess_min_soc = 10.0; bess_max_soc = 90.0
            bess_charge_eff = 90.0; bess_discharge_eff = 95.0
            bess_power_capex = 300; bess_energy_capex = 300
            bess_opex = 10; bess_lifetime = 15
            apply_bess_degradation = False
            bess_chemistry = None; bess_deg_file = None
        else:
            st.subheader("Power Range")
            col1, col2 = st.columns(2)
            with col1:
                bess_min = st.number_input("Min Power (MW)", value=1.0, min_value=0.0, step=0.5, key="bess_min")
            with col2:
                bess_max = st.number_input("Max Power (MW)", value=5.0, min_value=0.0, step=0.5, key="bess_max")
            bess_step = st.number_input("Step (MW)", value=1.0, min_value=0.1, step=0.1, key="bess_step")

            st.subheader("Storage Parameters")
            col1, col2 = st.columns(2)
            with col1:
                bess_duration = st.number_input("Duration (hours)", value=4.0, min_value=0.5, step=0.5, key="bess_duration")
                bess_min_soc = st.number_input("Min SOC (%)", value=10.0, min_value=0.0, max_value=50.0, step=5.0, key="bess_min_soc")
            with col2:
                bess_max_soc = st.number_input("Max SOC (%)", value=90.0, min_value=50.0, max_value=100.0, step=5.0, key="bess_max_soc")

            st.markdown("---")
            st.subheader("🔬 Advanced Analysis")

            apply_bess_degradation = st.checkbox(
                "Apply BESS Degradation Analysis", value=False,
                help="Enable battery capacity retention modeling over project life",
                key="apply_bess_degradation"
            )

            st.markdown("---")
            if apply_bess_degradation:
                st.info("ℹ️ Efficiency values will be controlled by the degradation curve (inputs disabled)")

            col1, col2 = st.columns(2)
            with col1:
                bess_charge_eff = st.number_input(
                    "Charge Efficiency (%)", value=90.0, min_value=50.0, max_value=100.0,
                    step=1.0, key="bess_charge_eff", disabled=apply_bess_degradation
                )
            with col2:
                bess_discharge_eff = st.number_input(
                    "Discharge Efficiency (%)", value=95.0, min_value=50.0, max_value=100.0,
                    step=1.0, key="bess_discharge_eff", disabled=apply_bess_degradation
                )

            st.subheader("Financial Parameters")
            col1, col2 = st.columns(2)
            with col1:
                bess_power_capex = st.number_input("Power CapEx ($/kW)", value=300, step=10, key="bess_power_capex")
                bess_energy_capex = st.number_input("Energy CapEx ($/kWh)", value=300, step=10, key="bess_energy_capex")
            with col2:
                bess_opex = st.number_input("OpEx ($/kW/yr)", value=10, step=1, key="bess_opex")
                bess_lifetime = st.number_input("Lifetime (years)", value=15, step=1, key="bess_life")

            if apply_bess_degradation:
                st.markdown("**Upload BESS Degradation Curve:**")
                bess_deg_file = st.file_uploader(
                    "Select BESS Degradation CSV File", type=['csv', 'xlsx'],
                    help="Upload CSV with columns: Year, Capacity_Retention_%, Charging_Efficiency_%, Discharging_Efficiency_%",
                    key="bess_deg_file"
                )

                if bess_deg_file:
                    st.success(f"✓ Uploaded: {bess_deg_file.name}")
                    try:
                        deg_data_preview = parse_bess_degradation_file(bess_deg_file)
                        if deg_data_preview:
                            st.info(
                                f"✓ Year 1:  {deg_data_preview[1]['capacity']:.1f}% capacity | "
                                f"Charge: {deg_data_preview[1]['charge_eff']:.2f}% | "
                                f"Discharge: {deg_data_preview[1]['discharge_eff']:.2f}%\n\n"
                                f"✓ Year 10: {deg_data_preview[10]['capacity']:.1f}% capacity | "
                                f"Charge: {deg_data_preview[10]['charge_eff']:.2f}% | "
                                f"Discharge: {deg_data_preview[10]['discharge_eff']:.2f}%\n\n"
                                f"✓ Year 25: {deg_data_preview[25]['capacity']:.1f}% capacity | "
                                f"Charge: {deg_data_preview[25]['charge_eff']:.2f}% | "
                                f"Discharge: {deg_data_preview[25]['discharge_eff']:.2f}%"
                            )
                            bess_deg_file.seek(0)
                    except:
                        pass
                else:
                    st.warning("⚠️ Please upload a BESS degradation CSV file to proceed")

                st.markdown("**Need a template? Download a preset:**")
                col1, col2, col3 = st.columns(3)
                with col1:
                    try:
                        nmc_data = open('bess_degradation_lithium_nmc.csv', 'rb').read()
                    except:
                        nmc_data = create_bess_degradation_template()
                    st.download_button(label="📥 Lithium NMC", data=nmc_data,
                        file_name="bess_degradation_lithium_nmc.csv", mime="text/csv",
                        key="bess_nmc_download", help="Standard Lithium NMC degradation curve")
                with col2:
                    try:
                        lfp_data = open('bess_degradation_lithium_lfp.csv', 'rb').read()
                    except:
                        lfp_data = create_bess_degradation_template()
                    st.download_button(label="📥 Lithium LFP", data=lfp_data,
                        file_name="bess_degradation_lithium_lfp.csv", mime="text/csv",
                        key="bess_lfp_download", help="Long-life Lithium LFP degradation curve")
                with col3:
                    try:
                        sodi_data = open('bess_degradation_sodium_ion.csv', 'rb').read()
                    except:
                        sodi_data = create_bess_degradation_template()
                    st.download_button(label="📥 Sodium-Ion", data=sodi_data,
                        file_name="bess_degradation_sodium_ion.csv", mime="text/csv",
                        key="bess_sodi_download", help="Emerging Sodium-Ion degradation curve")

                bess_chemistry = "Custom (From CSV)"
            else:
                bess_chemistry = None
                bess_deg_file = None

    st.markdown("---")

    # ── UPLOAD PROFILES ──
    with st.expander("📁 UPLOAD PROFILES", expanded=True):
        st.markdown("**Upload your energy profiles (CSV/Excel):**")

        load_file = st.file_uploader(
            "📊 Load Profile (Required)", type=['csv', 'xlsx'],
            key="load_file", help="8760-hour load profile in kW"
        )
        if load_file:
            st.success(f"✓ {load_file.name}")

        pv_file = st.file_uploader(
            "☀️ PV Profile (Required if PV enabled)", type=['csv', 'xlsx'],
            key="pv_file", help="8760-hour PV generation profile (1 kW baseline)"
        )
        if pv_file:
            st.success(f"✓ {pv_file.name}")

        wind_file = st.file_uploader(
            "💨 Wind Profile (Required if Wind enabled)", type=['csv', 'xlsx'],
            key="wind_file", help="8760-hour wind generation profile"
        )
        if wind_file:
            st.success(f"✓ {wind_file.name}")

        hydro_file = st.file_uploader(
            "💧 Hydro Profile (Optional - for variable hydro)", type=['csv', 'xlsx'],
            key="hydro_file",
            help="8760-hour hydro availability profile. Optional - if not provided, hydro assumed constant 24/7."
        )
        if hydro_file:
            st.success(f"✓ {hydro_file.name}")

    st.markdown("---")

    # ── PROJECT PARAMETERS ──
    with st.expander("💰 PROJECT PARAMETERS"):
        st.subheader("Economic Parameters")
        col1, col2 = st.columns(2)
        with col1:
            discount_rate = st.number_input(
                "Nominal Discount Rate (%)", value=8.0, min_value=0.0, max_value=30.0,
                step=0.5, key="discount_rate"
            )
            inflation_rate = st.number_input(
                "Inflation Rate (%)", value=2.0, min_value=0.0, max_value=10.0,
                step=0.5, key="inflation_rate"
            )
        with col2:
            project_lifetime = st.number_input(
                "Project Lifetime (years)", value=25, min_value=10, max_value=50,
                step=5, key="project_lifetime"
            )
            target_unmet_percent = st.number_input(
                "Target Max Unmet Load (%)", value=5.0, min_value=0.0, max_value=20.0,
                step=0.5, key="target_unmet"
            )


# ==============================================================================
# MAIN CONTENT - OPTIMIZATION EXECUTION
# ==============================================================================

st.header("🚀 Run Optimization")

# Validation
validation_errors = []

if not OPTIMIZATION_AVAILABLE:
    validation_errors.append("❌ Optimization module not available")
if load_file is None:
    validation_errors.append("❌ Load profile is required")
if enable_pv and pv_file is None:
    validation_errors.append("❌ PV profile required when PV is enabled")
if enable_wind and wind_file is None:
    validation_errors.append("❌ Wind profile required when Wind is enabled")
if enable_hydro and hydro_file is None:
    st.info("ℹ️ Hydro is enabled but no profile uploaded. Will use constant availability.")

if validation_errors:
    for error in validation_errors:
        st.error(error)
    st.stop()

# RUN OPTIMIZATION BUTTON
if st.button("▶️ RUN OPTIMIZATION", type="primary", use_container_width=True):

    with st.spinner("Running optimization..."):
        try:
            progress_bar = st.progress(0)
            status_text = st.empty()

            # Step 1: Load profiles
            status_text.text("📂 Loading input profiles...")
            progress_bar.progress(10)

            load_df = pd.read_csv(load_file) if load_file.name.endswith('.csv') else pd.read_excel(load_file)

            if pv_file:
                pv_df = pd.read_csv(pv_file) if pv_file.name.endswith('.csv') else pd.read_excel(pv_file)
            else:
                pv_df = pd.DataFrame({'PVsyst_kW': [0] * 8760})

            if wind_file:
                wind_df = pd.read_csv(wind_file) if wind_file.name.endswith('.csv') else pd.read_excel(wind_file)
            else:
                wind_df = pd.DataFrame({'Wind_kW': [0] * 8760})

            if hydro_file:
                hydro_df = pd.read_csv(hydro_file) if hydro_file.name.endswith('.csv') else pd.read_excel(hydro_file)
            else:
                hydro_df = pd.DataFrame({'Hydro_Available_kW': [1.0] * 8760})

            load_profile = load_df.iloc[:, 0].values if len(load_df.columns) == 1 else load_df.iloc[:, 1].values
            pvsyst_profile = pv_df.iloc[:, 0].values if len(pv_df.columns) == 1 else pv_df.iloc[:, 1].values
            wind_profile = wind_df.iloc[:, 0].values if len(wind_df.columns) == 1 else wind_df.iloc[:, 1].values

            progress_bar.progress(20)

            # Step 2: Build configs
            status_text.text("⚙️ Configuring optimization parameters...")

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
                'pv_start': pv_min * 1000, 'pv_end': pv_max * 1000, 'pv_step': pv_step * 1000,
                'wind_start': wind_min * 1000, 'wind_end': wind_max * 1000, 'wind_step': wind_step * 1000,
                'hydro_start': hydro_min * 1000, 'hydro_end': hydro_max * 1000, 'hydro_step': hydro_step * 1000,
                'bess_start': bess_min * 1000, 'bess_end': bess_max * 1000, 'bess_step': bess_step * 1000
            }

            solar_config = {
                'capex_per_kw': pv_capex,
                'om_per_kw_year': pv_opex,
                'lifetime': pv_lifetime,
                'baseline_kw': 1.0
            }

            wind_config = {
                'capex_per_kw': wind_capex,
                'om_per_kw_year': wind_opex,
                'lifetime': wind_lifetime
            }

            hydro_config = {
                'capex_per_kw': hydro_capex,
                'om_per_kw_year': hydro_opex,
                'lifetime': hydro_lifetime,
                'hours_per_day': hydro_hours_per_day
            }

            bess_config = {
                'duration_hours': bess_duration,
                'min_soc': bess_min_soc,
                'max_soc': bess_max_soc,
                'charge_eff': 0.90 if apply_bess_degradation else (bess_charge_eff / 100),
                'discharge_eff': 0.95 if apply_bess_degradation else (bess_discharge_eff / 100),
                'power_capex_per_kw': bess_power_capex,
                'energy_capex_per_kwh': bess_energy_capex,
                'om_per_kw_year': bess_opex,
                'lifetime': bess_lifetime
            }

            progress_bar.progress(30)

            # Step 3: Grid search
            status_text.text("⚙️ Running grid search optimization...")

            results_df = opt_module.grid_search_optimize_hydro(
                config, grid_config, solar_config, wind_config,
                hydro_config, bess_config,
                load_profile, pvsyst_profile, wind_profile, None
            )

            progress_bar.progress(60)

            # Step 4: Find optimal
            status_text.text("🔍 Finding optimal solution...")
            optimal = opt_module.find_optimal_solution(results_df)

            if optimal is None:
                st.error("❌ No feasible solution found! Try adjusting search ranges or unmet load target.")
                st.stop()

            progress_bar.progress(70)

            # Step 5: Re-run dispatch for optimal to get hourly profile
            status_text.text("📈 Generating optimal dispatch profile...")

            optimal_dispatch_df = opt_module.calculate_dispatch_with_hydro(
                load_profile,
                pvsyst_profile,
                wind_profile,
                optimal['PV_kW'],
                optimal['Wind_kW'],
                optimal['Hydro_kW'],
                optimal['BESS_Power_kW'],
                optimal['BESS_Capacity_kWh'],
                solar_config,
                wind_config,
                hydro_config,
                bess_config,
                int(optimal['Hydro_Window_Start']),
                int(optimal['Hydro_Window_End'])
            )

            progress_bar.progress(75)

            # Step 5b: Calculate NPC data + electrical metrics for enhanced results display
            status_text.text("📊 Calculating electrical metrics...")

            npc_data = opt_module.calculate_npc_homer_style(
                optimal['PV_kW'], optimal['Wind_kW'], optimal['Hydro_kW'],
                optimal['BESS_Power_kW'], optimal['BESS_Capacity_kWh'],
                solar_config, wind_config, hydro_config, bess_config, config,
                None, False, optimal['Total_Energy_Served_kWh']
            )

            component_capacities = {
                'pv_kw':   optimal['PV_kW'],
                'wind_kw': optimal['Wind_kW'],
                'hydro_kw': optimal['Hydro_kW'],
                'bess_kwh': optimal['BESS_Capacity_kWh']
            }

            component_configs = {
                'bess_max_soc':  bess_max_soc / 100,
                'bess_min_soc':  bess_min_soc / 100,
                'bess_lifetime': bess_lifetime
            }

            electrical_metrics = opt_module.calculate_electrical_metrics(
                optimal_dispatch_df, component_capacities, component_configs,
                npc_data, project_lifetime
            )

            # Step 6: Degradation analysis (optional)
            use_degradation = apply_pv_degradation or apply_bess_degradation

            if use_degradation:
                status_text.text("🔬 Running 25-year degradation analysis...")

                pv_deg_type = None
                pv_deg_data = None

                if apply_pv_degradation:
                    if pv_deg_method == "Simple (Annual Rate)":
                        pv_deg_type = 'simple'
                        pv_deg_data = pv_annual_deg_rate
                    else:
                        if pv_deg_file:
                            pv_deg_type = 'curve'
                            pv_deg_data = parse_pv_degradation_file(pv_deg_file)
                            if pv_deg_data is None:
                                st.error("❌ Failed to parse PV degradation file")
                                st.stop()

                bess_deg_data = None

                if apply_bess_degradation:
                    if bess_chemistry == "Custom (From CSV)":
                        if bess_deg_file:
                            bess_deg_data = parse_bess_degradation_file(bess_deg_file)
                            if bess_deg_data is None:
                                st.error("❌ Failed to parse BESS degradation file")
                                st.stop()
                    else:
                        if hasattr(opt_module, 'BESS_DEGRADATION_PRESETS'):
                            bess_deg_data = opt_module.BESS_DEGRADATION_PRESETS[bess_chemistry]
                        else:
                            st.error("❌ BESS degradation presets not available in optimization module")
                            st.stop()

                degradation_results = opt_module.run_multi_year_degradation_analysis(
                    optimal.to_dict(),
                    load_profile, pvsyst_profile, wind_profile,
                    solar_config, wind_config, hydro_config, bess_config,
                    project_lifetime=project_lifetime,
                    pv_degradation_type=pv_deg_type,
                    pv_degradation_data=pv_deg_data,
                    bess_degradation_data=bess_deg_data
                )

                progress_bar.progress(90)

            # Step 7: Store results
            status_text.text("💾 Saving results...")

            st.session_state.results = {
                'optimal_solution': optimal,
                'all_results': results_df,
                'optimal_dispatch': optimal_dispatch_df,
                'config': config,
                'degradation_enabled': use_degradation,
                'electrical_metrics': electrical_metrics,
                'npc_data': npc_data
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

    st.markdown("---")
    st.header("📊 Optimization Results")

    # Create tabs
    if results.get('degradation_enabled', False):
        tab1, tab2, tab3 = st.tabs(["📊 Summary", "💰 Cost & Performance", "🔬 Degradation"])
    else:
        tab1, tab2 = st.tabs(["📊 Summary", "💰 Cost & Performance"])

    # ── TAB 1: SUMMARY ──
    with tab1:
        st.subheader("Optimal System Configuration")

        col1, col2, col3, col4 = st.columns(4)
        with col1:
            st.metric("☀️ Solar PV", f"{optimal['PV_kW']/1000:.2f} MW",
                      delta=f"{optimal['PV_Energy_kWh']/1000:.0f} MWh/yr")
        with col2:
            st.metric("💨 Wind", f"{optimal['Wind_kW']/1000:.2f} MW",
                      delta=f"{optimal['Wind_Energy_kWh']/1000:.0f} MWh/yr")
        with col3:
            st.metric("💧 Hydro", f"{optimal['Hydro_kW']/1000:.2f} MW",
                      delta=f"{optimal['Hydro_Energy_kWh']/1000:.0f} MWh/yr")
        with col4:
            st.metric("🔋 BESS", f"{optimal['BESS_Power_kW']/1000:.2f} MW",
                      delta=f"{optimal['BESS_Capacity_kWh']/1000:.1f} MWh")

        st.markdown("---")

        # KPIs
        st.subheader("Key Performance Indicators")

        col1, col2, col3, col4 = st.columns(4)
        with col1:
            st.metric("Net Present Cost", f"${optimal['NPC_Total']/1_000_000:.2f}M")
        with col2:
            st.metric("LCOE", f"${optimal['LCOE_per_kWh']*1000:.2f}/MWh")
        with col3:
            st.metric(
                "Unmet Load", f"{optimal['Unmet_Load_Percent']:.2f}%",
                delta="Target" if optimal['Unmet_Load_Percent'] <= target_unmet_percent else "Over",
                delta_color="normal" if optimal['Unmet_Load_Percent'] <= target_unmet_percent else "inverse"
            )
        with col4:
            re_penetration = (
                optimal['PV_Energy_kWh'] + optimal['Wind_Energy_kWh'] + optimal['Hydro_Energy_kWh']
            ) / optimal['Total_Load_kWh'] * 100
            st.metric("RE Penetration", f"{re_penetration:.1f}%")

        # LCOS row
        st.markdown("---")
        col1, col2 = st.columns(2)
        with col1:
            bess_annual_discharge = (
                results['optimal_dispatch']['BESS_Discharge_wieff_kW'].sum()
                if results.get('optimal_dispatch') is not None else 0
            )
            bess_npc_val = optimal.get('BESS_NPC', 0)
            lcos_val = 0
            if bess_annual_discharge > 0 and hasattr(opt_module, 'calculate_bess_lcos_from_npc'):
                lcos_val = opt_module.calculate_bess_lcos_from_npc(
                    bess_npc_val, bess_annual_discharge,
                    results['config']['project_lifetime']
                )
            st.metric("🔋 BESS LCOS", f"${lcos_val*1000:.2f}/MWh")
        with col2:
            st.metric("Annual BESS Discharge", f"{bess_annual_discharge/1000:.0f} MWh")

        st.markdown("---")

        # BESS Deployment Details
        if optimal['BESS_Power_kW'] > 0:
            st.subheader("🏗️ BESS Deployment Details (Sungrow PowerTitan 2.0)")

            bess_deployment = calculate_bess_deployment_sungrow(
                optimal['BESS_Power_kW'] / 1000,
                optimal['BESS_Capacity_kWh'] / 1000
            )

            col1, col2, col3 = st.columns(3)
            with col1:
                st.metric("Containers Required", f"{bess_deployment['num_containers']}")
                st.metric("Container Model", bess_deployment['container_model'])
                st.metric("MVS Units", f"{bess_deployment['num_mvs_units']}")
            with col2:
                st.metric("Deployed Capacity", f"{bess_deployment['actual_capacity_mwh']:.1f} MWh")
                st.metric("Deployed Power", f"{bess_deployment['actual_power_mw']:.1f} MW")
                # Use st.write to avoid truncation of layout description
                st.markdown("**Layout**")
                st.write(bess_deployment['layout_description'])
            with col3:
                st.metric("Total Area", f"{bess_deployment['total_area_hectares']:.2f} ha ({bess_deployment['total_area_acres']:.2f} acres)")
                st.metric("Power Density", f"{bess_deployment['power_density_mw_per_ha']:.2f} MW/ha")
                st.metric("Energy Density", f"{bess_deployment['energy_density_mwh_per_ha']:.2f} MWh/ha")

    # ── TAB 2: COST & PERFORMANCE ──
    with tab2:

        # ── Section 1: KPI Row ──
        st.subheader("💰 Cost Summary")
        col1, col2, col3, col4 = st.columns(4)
        with col1:
            st.metric("Total NPC", f"${optimal['NPC_Total']/1_000_000:.2f}M")
        with col2:
            st.metric("Total Capital", f"${optimal['CapEx_Total']/1_000_000:.2f}M")
        with col3:
            st.metric("System LCOE", f"${optimal['LCOE_per_kWh']*1000:.2f}/MWh")
        with col4:
            # LCOS from electrical_metrics if available
            em = results.get('electrical_metrics', {})
            bess_lcos = em.get('bess', {}).get('levelized_cost_per_mwh', 0)
            st.metric("BESS LCOS", f"${bess_lcos:.2f}/MWh" if bess_lcos > 0 else "N/A")

        st.markdown("---")

        # ── Section 2: NPC by Component + NPC by Cost Type ──
        st.subheader("Net Present Cost Breakdown")

        # Build component data
        comp_names, comp_npc, comp_capex, comp_repl, comp_om, comp_salvage = [], [], [], [], [], []
        comp_map = [
            ('Solar PV',  optimal['PV_kW'],         'PV'),
            ('Wind',      optimal['Wind_kW'],        'Wind'),
            ('Hydro',     optimal['Hydro_kW'],       'Hydro'),
            ('BESS',      optimal['BESS_Power_kW'],  'BESS'),
        ]
        for label, capacity, key in comp_map:
            if capacity > 0:
                comp_names.append(label)
                comp_npc.append(optimal[f'{key}_NPC'] / 1e6)
                comp_capex.append(optimal[f'{key}_CapEx'] / 1e6)
                comp_repl.append(optimal.get(f'{key}_Replacement', 0) / 1e6)
                comp_om.append(optimal.get(f'{key}_OM', 0) / 1e6)
                comp_salvage.append(optimal.get(f'{key}_Salvage', 0) / 1e6)

        col1, col2 = st.columns(2)

        with col1:
            # Chart: NPC by Component
            bar_colors = {'Solar PV': '#FDB462', 'Wind': '#80B1D3', 'Hydro': '#8DD3C7', 'BESS': '#FB8072'}
            colors = [bar_colors.get(n, '#BEBADA') for n in comp_names]
            fig_npc_comp = go.Figure(data=[go.Bar(
                x=comp_names, y=comp_npc,
                marker_color=colors,
                text=[f'${v:.2f}M' for v in comp_npc],
                textposition='outside'
            )])
            fig_npc_comp.update_layout(
                title='Net Present Cost by Component',
                xaxis_title='Component', yaxis_title='NPC ($M)',
                height=380, showlegend=False,
                plot_bgcolor='white', paper_bgcolor='white',
                font=dict(color='#333333')
            )
            fig_npc_comp.update_yaxes(gridcolor='#EEEEEE')
            st.plotly_chart(fig_npc_comp, use_container_width=True)

            # Table: NPC by Component
            comp_table = pd.DataFrame({
                'Component': comp_names,
                'Total NPC ($)': [f"${v*1e6:,.0f}" for v in comp_npc]
            })
            st.dataframe(comp_table, use_container_width=True, hide_index=True)

        with col2:
            # Chart: NPC by Cost Type
            total_cap  = optimal['CapEx_Total'] / 1e6
            total_repl = optimal.get('Total_Replacement', 0) / 1e6
            total_om   = optimal.get('Total_OM', 0) / 1e6
            total_salv = optimal.get('Total_Salvage', 0) / 1e6

            cost_type_names   = ['Capital', 'Replacement', 'O&M', 'Salvage']
            cost_type_values  = [total_cap, total_repl, total_om, -total_salv]
            cost_type_colors  = ['#2E7D32', '#1976D2', '#F57C00', '#C62828']

            fig_npc_type = go.Figure(data=[go.Bar(
                x=cost_type_names, y=cost_type_values,
                marker_color=cost_type_colors,
                text=[f'${v:.2f}M' for v in cost_type_values],
                textposition='outside'
            )])
            fig_npc_type.update_layout(
                title='Net Present Cost by Cost Type',
                xaxis_title='Cost Type', yaxis_title='Cost ($M)',
                height=380, showlegend=False,
                plot_bgcolor='white', paper_bgcolor='white',
                font=dict(color='#333333')
            )
            fig_npc_type.update_yaxes(gridcolor='#EEEEEE')
            st.plotly_chart(fig_npc_type, use_container_width=True)

            # Table: Cost Type totals
            cost_type_table = pd.DataFrame({
                'Cost Type': cost_type_names,
                'System Total ($)': [
                    f"${optimal['CapEx_Total']:,.0f}",
                    f"${optimal.get('Total_Replacement', 0):,.0f}",
                    f"${optimal.get('Total_OM', 0):,.0f}",
                    f"${optimal.get('Total_Salvage', 0):,.0f}"
                ]
            })
            st.dataframe(cost_type_table, use_container_width=True, hide_index=True)

        st.markdown("---")

        # ── Section 3: Detailed Component Cost Table ──
        st.subheader("Detailed Component Cost Breakdown")
        detailed_cost_data = []
        for label, capacity, key in comp_map:
            if capacity > 0:
                cap_str = (f"{capacity/1000:.2f} MW" if key != 'BESS'
                           else f"{optimal['BESS_Power_kW']/1000:.2f} MW / {optimal['BESS_Capacity_kWh']/1000:.1f} MWh")
                detailed_cost_data.append({
                    'Component':        label,
                    'Capacity':         cap_str,
                    'CapEx ($M)':       f"${optimal[f'{key}_CapEx']/1e6:.3f}",
                    'Replacement ($M)': f"${optimal.get(f'{key}_Replacement', 0)/1e6:.3f}",
                    'O&M PV ($M)':      f"${optimal.get(f'{key}_OM', 0)/1e6:.3f}",
                    'Salvage ($M)':     f"${optimal.get(f'{key}_Salvage', 0)/1e6:.3f}",
                    'NPC ($M)':         f"${optimal[f'{key}_NPC']/1e6:.3f}"
                })
        if detailed_cost_data:
            st.dataframe(pd.DataFrame(detailed_cost_data), use_container_width=True, hide_index=True)

        st.markdown("---")

        # ── Section 4: Cash Flow Chart ──
        st.subheader("💵 Nominal Cash Flow Analysis")
        project_lt = results['config']['project_lifetime']
        years = list(range(0, project_lt + 1))

        capital_flow    = [0.0] * len(years)
        operating_flow  = [0.0] * len(years)
        replacement_flow = [0.0] * len(years)
        salvage_flow    = [0.0] * len(years)

        # Year 0: full capital outlay
        capital_flow[0] = -optimal['CapEx_Total'] / 1e6

        # Annual O&M — use present value / project lifetime as nominal annual proxy
        total_om_pv = optimal.get('Total_OM', 0) / 1e6
        annual_om = total_om_pv / project_lt if project_lt > 0 else 0
        for yr in range(1, project_lt + 1):
            operating_flow[yr] = -annual_om

        # Replacement: spread at component lifetime intervals
        for label, capacity, key in comp_map:
            if capacity > 0:
                comp_repl_val = optimal.get(f'{key}_Replacement', 0) / 1e6
                if comp_repl_val > 0:
                    # Approximate replacement year from config
                    lt_map = {'PV': results['config'].get('pv_lifetime', 30),
                              'Wind': results['config'].get('wind_lifetime', 20),
                              'Hydro': results['config'].get('hydro_lifetime', 50),
                              'BESS': results['config'].get('bess_lifetime', 15)}
                    comp_lt = lt_map.get(key, 20)
                    yr = comp_lt
                    while yr < project_lt:
                        if yr < len(replacement_flow):
                            replacement_flow[yr] -= comp_repl_val / max(1, project_lt // comp_lt)
                        yr += comp_lt

        # Final year: salvage (positive)
        salvage_flow[-1] = optimal.get('Total_Salvage', 0) / 1e6

        fig_cf = go.Figure()
        fig_cf.add_trace(go.Bar(name='Capital',     x=years, y=capital_flow,     marker_color='#2E7D32'))
        fig_cf.add_trace(go.Bar(name='Operating',   x=years, y=operating_flow,   marker_color='#F57C00'))
        fig_cf.add_trace(go.Bar(name='Replacement', x=years, y=replacement_flow, marker_color='#1976D2'))
        fig_cf.add_trace(go.Bar(name='Salvage',     x=years, y=salvage_flow,     marker_color='#43A047'))
        fig_cf.update_layout(
            title='Nominal Cash Flow Over Project Lifetime',
            xaxis_title='Year', yaxis_title='Cash Flow ($M)',
            barmode='relative', height=420,
            showlegend=True,
            plot_bgcolor='white', paper_bgcolor='white',
            font=dict(color='#333333'),
            legend=dict(orientation='h', yanchor='bottom', y=1.02, xanchor='right', x=1)
        )
        fig_cf.update_xaxes(gridcolor='#EEEEEE')
        fig_cf.update_yaxes(gridcolor='#EEEEEE')
        st.plotly_chart(fig_cf, use_container_width=True)

        st.markdown("---")

        # ── Section 5: Electrical Performance Metrics ──
        st.subheader("⚡ Electrical Performance Metrics")

        em = results.get('electrical_metrics', {})
        if em:
            col1, col2 = st.columns(2)

            with col1:
                # Solar PV table
                if optimal['PV_kW'] > 0:
                    st.markdown("**☀️ Solar PV**")
                    pv = em.get('pv', {})
                    pv_table = pd.DataFrame({
                        'Metric': ['Rated Capacity', 'Mean Output', 'Capacity Factor',
                                   'Total Production', 'Hours of Operation', 'Levelized Cost (LCOE)'],
                        'Value': [
                            f"{pv.get('rated_capacity_kw', 0):,.1f} kW",
                            f"{pv.get('mean_output_kw', 0):,.1f} kW",
                            f"{pv.get('capacity_factor_pct', 0):.2f}%",
                            f"{pv.get('total_production_kwh', 0):,.0f} kWh/yr",
                            f"{pv.get('hours_of_operation', 0):,.0f} hrs/yr",
                            f"${pv.get('levelized_cost_per_kwh', 0):.4f}/kWh"
                        ]
                    })
                    st.dataframe(pv_table, use_container_width=True, hide_index=True)

                # Hydro table
                if optimal['Hydro_kW'] > 0:
                    st.markdown("**💧 Hydro**")
                    hydro = em.get('hydro', {})
                    hydro_table = pd.DataFrame({
                        'Metric': ['Rated Capacity', 'Mean Output', 'Capacity Factor',
                                   'Total Production', 'Hours of Operation', 'Levelized Cost (LCOE)'],
                        'Value': [
                            f"{hydro.get('rated_capacity_kw', 0):,.1f} kW",
                            f"{hydro.get('mean_output_kw', 0):,.1f} kW",
                            f"{hydro.get('capacity_factor_pct', 0):.2f}%",
                            f"{hydro.get('total_production_kwh', 0):,.0f} kWh/yr",
                            f"{hydro.get('hours_of_operation', 0):,.0f} hrs/yr",
                            f"${hydro.get('levelized_cost_per_kwh', 0):.4f}/kWh"
                        ]
                    })
                    st.dataframe(hydro_table, use_container_width=True, hide_index=True)

            with col2:
                # Wind table
                if optimal['Wind_kW'] > 0:
                    st.markdown("**💨 Wind**")
                    wind = em.get('wind', {})
                    wind_table = pd.DataFrame({
                        'Metric': ['Rated Capacity', 'Mean Output', 'Capacity Factor',
                                   'Total Production', 'Hours of Operation', 'Levelized Cost (LCOE)'],
                        'Value': [
                            f"{wind.get('rated_capacity_kw', 0):,.1f} kW",
                            f"{wind.get('mean_output_kw', 0):,.1f} kW",
                            f"{wind.get('capacity_factor_pct', 0):.2f}%",
                            f"{wind.get('total_production_kwh', 0):,.0f} kWh/yr",
                            f"{wind.get('hours_of_operation', 0):,.0f} hrs/yr",
                            f"${wind.get('levelized_cost_per_kwh', 0):.4f}/kWh"
                        ]
                    })
                    st.dataframe(wind_table, use_container_width=True, hide_index=True)

                # BESS table
                if optimal['BESS_Power_kW'] > 0:
                    st.markdown("**🔋 Battery Storage**")
                    bess = em.get('bess', {})
                    bess_table = pd.DataFrame({
                        'Metric': ['Nominal Capacity', 'Usable Capacity', 'Autonomy',
                                   'Energy In', 'Energy Out', 'Losses',
                                   'Annual Throughput', 'Expected Life', 'Levelized Cost (LCOS)'],
                        'Value': [
                            f"{bess.get('nominal_capacity_kwh', 0):,.1f} kWh",
                            f"{bess.get('usable_capacity_kwh', 0):,.1f} kWh",
                            f"{bess.get('autonomy_hours', 0):.2f} hours",
                            f"{bess.get('energy_in_kwh', 0):,.0f} kWh/yr",
                            f"{bess.get('energy_out_kwh', 0):,.0f} kWh/yr",
                            f"{bess.get('losses_kwh', 0):,.0f} kWh/yr",
                            f"{bess.get('annual_throughput_kwh', 0):,.0f} kWh/yr",
                            f"{bess.get('expected_life_years', 0):.0f} years",
                            f"${bess.get('levelized_cost_per_kwh', 0):.4f}/kWh"
                        ]
                    })
                    st.dataframe(bess_table, use_container_width=True, hide_index=True)
        else:
            st.info("Electrical metrics not available. Re-run optimization to generate them.")

        st.markdown("---")

        # ── Section 6: Energy Production Mix ──
        st.subheader("📊 Annual Energy Production Mix")

        col1, col2 = st.columns(2)
        with col1:
            energy_mix_fig = create_energy_mix_pie(results)
            if energy_mix_fig:
                st.plotly_chart(energy_mix_fig, use_container_width=True)
        with col2:
            gen_data = []
            total_gen_mwh = 0
            for source, key in [('Solar PV', 'PV'), ('Wind', 'Wind'), ('Hydro', 'Hydro')]:
                val = optimal.get(f'{key}_Energy_kWh', 0) / 1000
                if val > 0:
                    gen_data.append({'Source': source, 'Energy (MWh/yr)': f"{val:,.1f}"})
                    total_gen_mwh += val
            if gen_data:
                # Add percentages
                for row in gen_data:
                    val = float(row['Energy (MWh/yr)'].replace(',', ''))
                    row['Share (%)'] = f"{val / total_gen_mwh * 100:.1f}%"
                gen_data.append({'Source': 'Total', 'Energy (MWh/yr)': f"{total_gen_mwh:,.1f}", 'Share (%)': '100.0%'})
                st.dataframe(pd.DataFrame(gen_data), use_container_width=True, hide_index=True)

            # Energy balance summary
            st.markdown("**Energy Balance**")
            total_load = optimal['Total_Load_kWh'] / 1000
            unmet = optimal['Unmet_Load_kWh'] / 1000
            curtailment = optimal.get('Total_Curtailment_kWh', 0) / 1000
            balance_data = pd.DataFrame({
                'Metric': ['Annual Load', 'Energy Served', 'Unmet Load', 'Curtailment'],
                'Value (MWh)': [
                    f"{total_load:,.1f}",
                    f"{total_load - unmet:,.1f}",
                    f"{unmet:,.1f}",
                    f"{curtailment:,.1f}"
                ]
            })
            st.dataframe(balance_data, use_container_width=True, hide_index=True)

        st.markdown("---")

        # ── Section 7: Single Day Dispatch Profile ──
        st.subheader("📈 Typical Day Dispatch Profile")
        dispatch_fig = create_single_day_dispatch_profile(results)
        if dispatch_fig:
            st.plotly_chart(dispatch_fig, use_container_width=True)
            st.caption("Representative day based on median PV production day")
        else:
            st.info("Dispatch profile not available")

    # ── TAB 3: DEGRADATION ──
    if results.get('degradation_enabled', False):
        with tab3:
            st.header("🔬 25-Year Degradation Analysis")

            deg_data = results['degradation_analysis']
            deg_summary = deg_data['degradation_summary']
            yearly_df = deg_data['yearly_metrics']

            col1, col2, col3 = st.columns(3)
            with col1:
                st.metric("PV Degradation @ Year 25", f"{deg_summary['pv_degradation_year_25']:.2f}%")
            with col2:
                st.metric("BESS Retention @ Year 25", f"{deg_summary['bess_retention_year_25']:.1f}%")
            with col3:
                st.metric("Avg Unmet Load (25 years)", f"{deg_summary['avg_unmet_pct']:.2f}%")

            st.markdown("---")

            st.subheader("📈 System Performance Over Project Life")

            fig = make_subplots(specs=[[{"secondary_y": True}]])
            fig.add_trace(go.Scatter(
                x=yearly_df['Year'], y=yearly_df['Unmet_%'],
                name='Unmet Load %', mode='lines+markers',
                line=dict(color='#E63946', width=2), marker=dict(size=6)
            ), secondary_y=False)
            fig.add_trace(go.Scatter(
                x=yearly_df['Year'], y=yearly_df['PV_Degradation_%'],
                name='PV Degradation %', mode='lines',
                line=dict(color='#FDB462', width=2, dash='dash')
            ), secondary_y=True)
            fig.add_trace(go.Scatter(
                x=yearly_df['Year'], y=yearly_df['BESS_Retention_%'],
                name='BESS Retention %', mode='lines',
                line=dict(color='#8DD3C7', width=2, dash='dot')
            ), secondary_y=True)

            fig.update_xaxes(title_text="Project Year")
            fig.update_yaxes(title_text="Unmet Load (%)", secondary_y=False)
            fig.update_yaxes(title_text="Degradation/Retention (%)", secondary_y=True, range=[0, 105])
            fig.update_layout(
                title='System Performance with Degradation Over 25 Years',
                hovermode='x unified', height=500,
                legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1)
            )
            st.plotly_chart(fig, use_container_width=True)

            st.markdown("---")

            st.subheader("⚡ Annual Energy Generation")
            fig2 = go.Figure()
            fig2.add_trace(go.Scatter(
                x=yearly_df['Year'], y=yearly_df['PV_Energy_MWh'],
                name='PV', mode='lines', line=dict(color='#FFDB5C', width=2)
            ))
            fig2.add_trace(go.Scatter(
                x=yearly_df['Year'], y=yearly_df['Wind_Energy_MWh'],
                name='Wind', mode='lines', line=dict(color='#B3E2CD', width=2)
            ))
            fig2.add_trace(go.Scatter(
                x=yearly_df['Year'], y=yearly_df['Hydro_Energy_MWh'],
                name='Hydro', mode='lines', line=dict(color='#8DD3C7', width=2)
            ))
            fig2.update_layout(
                title='Annual Generation by Source',
                xaxis_title='Project Year', yaxis_title='Energy (MWh)',
                hovermode='x unified', height=400
            )
            st.plotly_chart(fig2, use_container_width=True)

            st.markdown("---")

            st.subheader("📊 Annual Performance Metrics")
            st.dataframe(yearly_df.copy().round(2), use_container_width=True, hide_index=True)

            st.markdown("---")

            st.subheader("📈 25-Year Summary")
            col1, col2, col3 = st.columns(3)
            with col1:
                st.metric("Total Energy Served",
                          f"{deg_summary['total_energy_served_25yr_GWh']:.1f} GWh",
                          help="Total energy served over 25-year project life")
            with col2:
                st.metric("Max Unmet Load", f"{deg_summary['max_unmet_pct']:.2f}%",
                          help="Highest unmet load percentage across all years")
            with col3:
                st.metric("Total Curtailment", f"{yearly_df['Curtailment_MWh'].sum():.0f} MWh",
                          help="Total energy curtailed over 25 years")

    # ──────────────────────────────────────────────────────────────────────────
    # EXPORT RESULTS — inside optimization_complete block
    # ──────────────────────────────────────────────────────────────────────────
    st.markdown("---")
    st.header("📥 Export Results")

    try:
        excel_buffer = build_excel_export(results, optimal, opt_module)

        st.download_button(
            label="📥 Download Full Results (Excel — All Sheets)",
            data=excel_buffer.getvalue(),
            file_name=f"optimization_results_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
            type="primary"
        )

        # Sheet contents summary for the user
        sheets = ["✅ Optimal Configuration & KPIs", "✅ All Combinations"]
        if results.get('optimal_dispatch') is not None:
            sheets.append("✅ Dispatch Year 1 (Optimal)")
        if results.get('degradation_enabled') and 'degradation_analysis' in results:
            sheets.append("✅ Degradation Summary (25-year metrics)")
            for yk in results['degradation_analysis']['selected_year_dispatch'].keys():
                yr = yk.split('_')[1]
                sheets.append(f"✅ Dispatch Year {yr}")

        st.markdown("**Excel file contains:**")
        for s in sheets:
            st.markdown(f"- {s}")

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
<div style="text-align: center; color: #666;">
    <p><strong>Energy Modeling Optimizer v5.0</strong> | Professional Hybrid System Design Tool</p>
    <p>Developed by SJ | March 2026</p>
</div>
""", unsafe_allow_html=True)
