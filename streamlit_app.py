"""
RENEWABLE ENERGY OPTIMIZATION TOOL - COMPLETE VERSION WITH CHECKBOXES
======================================================================
Features:
- Component enable/disable toggles
- Phase 1 + Phase 2 enhancements
- Sungrow BESS deployment metrics
- Fixed LCOE calculation from NPC
- Simplified cost analysis
- Typical daily dispatch profile
"""

import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
import plotly.express as px
from plotly.subplots import make_subplots
from datetime import datetime
import os
from io import BytesIO

# Import optimization code
try:
    import optimize_gridsearch_hydro_static_STREAMLITCHECK as opt_module
    OPTIMIZATION_AVAILABLE = True
except ImportError:
    OPTIMIZATION_AVAILABLE = False
    st.error("❌ Optimization module not found")


# ==============================================================================
# SUNGROW BESS DEPLOYMENT CALCULATION
# ==============================================================================

def calculate_bess_deployment_sungrow(bess_power_mw, bess_capacity_mwh):
    """
    Calculate BESS deployment using REAL Sungrow PowerTitan 2.0 specifications.
    Based on actual layout drawing with corrected spacing interpretation.
    """
    import math
    
    # Sungrow PowerTitan 2.0 specifications
    container_capacity_mwh = 10.0
    container_power_mw = 5.0
    container_length_m = 6.058
    container_width_m = 2.438
    container_height_m = 2.896
    
    # Spacing requirements from Sungrow layout drawing
    back_to_back_spacing_m = 0.150
    adjacent_spacing_m = 1.500
    mvs_spacing_m = 3.500
    mvs_width_m = 2.000
    perimeter_clearance_m = 5.000
    
    # Calculate number of containers
    num_containers_energy = math.ceil(bess_capacity_mwh / container_capacity_mwh)
    num_containers_power = math.ceil(bess_power_mw / container_power_mw)
    num_containers = max(num_containers_energy, num_containers_power)
    
    # Actual installed capacity
    actual_capacity_mwh = num_containers * container_capacity_mwh
    actual_power_mw = num_containers * container_power_mw
    
    # Number of MVS units (1 per 2 containers)
    num_mvs_units = math.ceil(num_containers / 2)
    
    # Calculate layout
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
        'container_dimensions': f"{container_length_m:.2f} × {container_width_m:.2f} × {container_height_m:.2f} m",
        'actual_capacity_mwh': actual_capacity_mwh,
        'actual_power_mw': actual_power_mw,
        'num_mvs_units': num_mvs_units,
        'layout_description': layout_desc,
        'container_footprint_m2': container_footprint_m2,
        'total_length_m': total_length_m,
        'total_width_m': total_width_m,
        'site_dimensions': f"{total_length_m:.1f} × {total_width_m:.1f} m",
        'total_area_m2': total_area_m2,
        'total_area_hectares': total_area_hectares,
        'total_area_acres': total_area_acres,
        'power_density_mw_per_ha': power_density_mw_per_ha,
        'energy_density_mwh_per_ha': energy_density_mwh_per_ha,
        'spacing_back_to_back_mm': 150,
        'spacing_adjacent_mm': 1500,
        'spacing_mvs_mm': 3500
    }


# ==============================================================================
# EXCEL EXPORT
# ==============================================================================

def export_results_industry_format(results_dict, results_df, optimal_row, config_params):
    """Export results in industry standard Excel format."""
    output = BytesIO()
    
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        # Sheet 1: Summary
        summary_data = []
        summary_data.extend([
            ['Parameter', 'Value'],
            ['Optimization Method', 'GRID_SEARCH'],
            ['NPC Calculation Method', 'Present Value Analysis'],
            ['Target Unmet Load (%)', config_params.get('target_unmet_percent', 0.1)],
            ['', ''],
            ['Nominal Discount Rate (%)', config_params.get('discount_rate', 8.0)],
            ['Inflation Rate (%)', config_params.get('inflation_rate', 2.0)],
            ['Project Lifetime (years)', config_params.get('project_lifetime', 25)],
            ['', ''],
            ['PV Capacity (kW)', results_dict['pv_capacity'] * 1000],
            ['Wind Capacity (kW)', results_dict['wind_capacity'] * 1000],
            ['Hydro Capacity (kW)', results_dict['hydro_capacity'] * 1000],
            ['BESS Power (kW)', results_dict['bess_power'] * 1000],
            ['BESS Capacity (kWh)', results_dict['bess_energy'] * 1000],
            ['', ''],
            ['Total NPC ($)', results_dict['npc']],
            ['System LCOE ($/MWh)', results_dict['lcoe']],
            ['Unmet Load (%)', results_dict['unmet_pct']],
        ])
        pd.DataFrame(summary_data).to_excel(writer, sheet_name='Summary', index=False, header=False)
        
        # Sheet 2: Cost Breakdown
        cost_breakdown = pd.DataFrame({
            'Component': ['PV', 'Wind', 'Hydro', 'BESS', 'System'],
            'Capital ($)': [
                optimal_row.get('PV_Capital_$', 0),
                optimal_row.get('Wind_Capital_$', 0),
                optimal_row.get('Hydro_Capital_$', 0),
                optimal_row.get('BESS_Capital_$', 0),
                optimal_row.get('Capital_$', 0)
            ],
            'NPC ($)': [
                optimal_row.get('PV_NPC_$', 0),
                optimal_row.get('Wind_NPC_$', 0),
                optimal_row.get('Hydro_NPC_$', 0),
                optimal_row.get('BESS_NPC_$', 0),
                optimal_row.get('NPC_$', 0)
            ]
        })
        cost_breakdown.to_excel(writer, sheet_name='Cost_Breakdown', index=False)
        
        # Sheet 3: All Results
        results_df.to_excel(writer, sheet_name='All_Results', index=False)
    
    output.seek(0)
    return output


# ==============================================================================
# COST ANALYSIS CHARTS
# ==============================================================================

def create_cost_analysis_charts_with_tables(results, optimal_row):
    """Create cost analysis charts with tables."""
    charts_and_tables = {}
    
    # Chart 1: Net Present Cost by Component
    components = ['PV', 'Wind', 'Hydro', 'BESS']
    npc_values = [
        optimal_row.get('PV_NPC_$', 0) / 1e6,
        optimal_row.get('Wind_NPC_$', 0) / 1e6,
        optimal_row.get('Hydro_NPC_$', 0) / 1e6,
        optimal_row.get('BESS_NPC_$', 0) / 1e6
    ]
    
    fig1 = go.Figure(data=[
        go.Bar(x=components, y=npc_values, 
               marker_color=['#FDB462', '#80B1D3', '#8DD3C7', '#FB8072'],
               text=[f'${v:.2f}M' for v in npc_values],
               textposition='outside')
    ])
    fig1.update_layout(
        title='Net Present Cost by Component',
        xaxis_title='Component',
        yaxis_title='NPC ($M)',
        height=400,
        showlegend=False
    )
    
    table1 = pd.DataFrame({
        'Component': components,
        'Total NPC ($)': [f"${optimal_row.get('PV_NPC_$', 0):,.0f}",
                         f"${optimal_row.get('Wind_NPC_$', 0):,.0f}",
                         f"${optimal_row.get('Hydro_NPC_$', 0):,.0f}",
                         f"${optimal_row.get('BESS_NPC_$', 0):,.0f}"]
    })
    
    charts_and_tables['npc_by_component'] = {'chart': fig1, 'table': table1}
    
    # Chart 2: Net Present Cost by Cost Type
    cost_types = ['Capital', 'Replacement', 'O&M', 'Salvage']
    cost_values = [
        optimal_row.get('Capital_$', 0) / 1e6,
        optimal_row.get('Replacement_$', 0) / 1e6,
        optimal_row.get('OM_$', 0) / 1e6,
        -optimal_row.get('Salvage_$', 0) / 1e6
    ]
    
    fig2 = go.Figure(data=[
        go.Bar(x=cost_types, y=cost_values,
               marker_color=['#2E7D32', '#1976D2', '#F57C00', '#C62828'],
               text=[f'${v:.2f}M' for v in cost_values],
               textposition='outside')
    ])
    fig2.update_layout(
        title='Net Present Cost by Cost Type',
        xaxis_title='Cost Type',
        yaxis_title='Cost ($M)',
        height=400,
        showlegend=False
    )
    
    table2 = pd.DataFrame({
        'Cost Type': cost_types,
        'System Total ($)': [f"${optimal_row.get('Capital_$', 0):,.0f}",
                            f"${optimal_row.get('Replacement_$', 0):,.0f}",
                            f"${optimal_row.get('OM_$', 0):,.0f}",
                            f"${optimal_row.get('Salvage_$', 0):,.0f}"]
    })
    
    charts_and_tables['npc_by_cost_type'] = {'chart': fig2, 'table': table2}
    
    return charts_and_tables


# ==============================================================================
# CASH FLOW CHART
# ==============================================================================

def create_fixed_cash_flow_chart(results, optimal_row):
    """Create corrected cash flow chart."""
    project_lifetime = results.get('config_params', {}).get('project_lifetime', 25)
    years = list(range(0, project_lifetime + 1))
    
    capital_flow = [0] * len(years)
    operating_flow = [0] * len(years)
    salvage_flow = [0] * len(years)
    
    capital_flow[0] = -optimal_row.get('Capital_$', 0) / 1e6
    annual_om = optimal_row.get('OM_$', 0) / project_lifetime / 1e6
    for year in range(1, project_lifetime + 1):
        operating_flow[year] = -annual_om
    salvage_flow[-1] = optimal_row.get('Salvage_$', 0) / 1e6
    
    fig = go.Figure()
    fig.add_trace(go.Bar(name='Capital', x=years, y=capital_flow, marker_color='#2E7D32'))
    fig.add_trace(go.Bar(name='Operating', x=years, y=operating_flow, marker_color='#F57C00'))
    fig.add_trace(go.Bar(name='Salvage', x=years, y=salvage_flow, marker_color='#388E3C'))
    
    fig.update_layout(
        title='Nominal Cash Flow',
        xaxis_title='Year',
        yaxis_title='Cash Flow ($M)',
        barmode='relative',
        height=450
    )
    
    return fig


# ==============================================================================
# ELECTRICAL METRICS TABLES
# ==============================================================================

def create_electrical_metrics_tables(electrical_metrics, bess_power_mw=0, bess_capacity_mwh=0):
    """Create formatted tables for electrical metrics."""
    tables = {}
    
    if electrical_metrics:
        # PV Table
        pv_data = pd.DataFrame({
            'Metric': ['Rated Capacity', 'Mean Output', 'Capacity Factor', 
                      'Total Production', 'Hours of Operation', 'Levelized Cost (LCOE)'],
            'Value': [
                f"{electrical_metrics['pv']['rated_capacity_kw']:,.1f} kW",
                f"{electrical_metrics['pv']['mean_output_kw']:,.1f} kW",
                f"{electrical_metrics['pv']['capacity_factor_pct']:.2f}%",
                f"{electrical_metrics['pv']['total_production_kwh']:,.0f} kWh/yr",
                f"{electrical_metrics['pv']['hours_of_operation']:,.0f} hrs/yr",
                f"${electrical_metrics['pv']['levelized_cost_per_kwh']:.4f}/kWh"
            ]
        })
        tables['pv'] = pv_data
        
        # Wind Table
        wind_data = pd.DataFrame({
            'Metric': ['Rated Capacity', 'Mean Output', 'Capacity Factor',
                      'Total Production', 'Hours of Operation', 'Levelized Cost (LCOE)'],
            'Value': [
                f"{electrical_metrics['wind']['rated_capacity_kw']:,.1f} kW",
                f"{electrical_metrics['wind']['mean_output_kw']:,.1f} kW",
                f"{electrical_metrics['wind']['capacity_factor_pct']:.2f}%",
                f"{electrical_metrics['wind']['total_production_kwh']:,.0f} kWh/yr",
                f"{electrical_metrics['wind']['hours_of_operation']:,.0f} hrs/yr",
                f"${electrical_metrics['wind']['levelized_cost_per_kwh']:.4f}/kWh"
            ]
        })
        tables['wind'] = wind_data
        
        # Hydro Table
        hydro_data = pd.DataFrame({
            'Metric': ['Rated Capacity', 'Mean Output', 'Capacity Factor',
                      'Total Production', 'Hours of Operation', 'Levelized Cost (LCOE)'],
            'Value': [
                f"{electrical_metrics['hydro']['rated_capacity_kw']:,.1f} kW",
                f"{electrical_metrics['hydro']['mean_output_kw']:,.1f} kW",
                f"{electrical_metrics['hydro']['capacity_factor_pct']:.2f}%",
                f"{electrical_metrics['hydro']['total_production_kwh']:,.0f} kWh/yr",
                f"{electrical_metrics['hydro']['hours_of_operation']:,.0f} hrs/yr",
                f"${electrical_metrics['hydro']['levelized_cost_per_kwh']:.4f}/kWh"
            ]
        })
        tables['hydro'] = hydro_data
        
        # BESS Performance Table
        bess_performance = pd.DataFrame({
            'Metric': ['Nominal Capacity', 'Usable Capacity', 'Autonomy', 'Energy In',
                      'Energy Out', 'Losses', 'Annual Throughput', 'Expected Life', 'Levelized Cost (LCOS)'],
            'Value': [
                f"{electrical_metrics['bess']['nominal_capacity_kwh']:,.1f} kWh",
                f"{electrical_metrics['bess']['usable_capacity_kwh']:,.1f} kWh",
                f"{electrical_metrics['bess']['autonomy_hours']:.2f} hours",
                f"{electrical_metrics['bess']['energy_in_kwh']:,.0f} kWh/yr",
                f"{electrical_metrics['bess']['energy_out_kwh']:,.0f} kWh/yr",
                f"{electrical_metrics['bess']['losses_kwh']:,.0f} kWh/yr",
                f"{electrical_metrics['bess']['annual_throughput_kwh']:,.0f} kWh/yr",
                f"{electrical_metrics['bess']['expected_life_years']:.0f} years",
                f"${electrical_metrics['bess']['levelized_cost_per_kwh']:.4f}/kWh"
            ]
        })
        tables['bess_performance'] = bess_performance
        
        # BESS Deployment Table
        if bess_capacity_mwh > 0:
            deployment = calculate_bess_deployment_sungrow(bess_power_mw, bess_capacity_mwh)
            bess_deployment = pd.DataFrame({
                'Metric': [
                    '--- DEPLOYMENT SPECIFICATIONS ---',
                    'OEM / Model',
                    'Number of Containers',
                    'Layout Configuration',
                    'Site Dimensions (L×W)',
                    'Total Site Area',
                    'Area (hectares)',
                    'Area (acres)',
                    'Power Density',
                    'Energy Density'
                ],
                'Value': [
                    '',
                    f"Sungrow {deployment['container_model']}",
                    f"{deployment['num_containers']} units (10 MWh / 5 MW each)",
                    deployment['layout_description'],
                    deployment['site_dimensions'],
                    f"{deployment['total_area_m2']:,.0f} m²",
                    f"{deployment['total_area_hectares']:.4f} ha",
                    f"{deployment['total_area_acres']:.3f} acres",
                    f"{deployment['power_density_mw_per_ha']:.1f} MW/ha",
                    f"{deployment['energy_density_mwh_per_ha']:.1f} MWh/ha"
                ]
            })
            tables['bess_deployment'] = bess_deployment
        else:
            tables['bess_deployment'] = pd.DataFrame({
                'Metric': ['Note'],
                'Value': ['No BESS in optimal configuration']
            })
    
    return tables


# ==============================================================================
# TYPICAL DAILY DISPATCH PROFILE
# ==============================================================================

def create_typical_daily_dispatch_profile(results):
    """Create typical 24-hour dispatch profile chart."""
    if 'optimal_dispatch' in results:
        dispatch_df = results['optimal_dispatch']
        dispatch_df['Hour_of_Day'] = dispatch_df['Hour'] % 24
        
        avg_profile = dispatch_df.groupby('Hour_of_Day').agg({
            'Load_kW': 'mean',
            'PV_Output_kW': 'mean',
            'Wind_Output_kW': 'mean',
            'Hydro_Output_kW': 'mean',
            'BESS_SOC_kWh': 'mean',
            'BESS_Discharge_kW': 'mean',
            'BESS_Charge_kW': 'mean'
        }).reset_index()
        
        avg_profile['Load_MW'] = avg_profile['Load_kW'] / 1000
        avg_profile['PV_MW'] = avg_profile['PV_Output_kW'] / 1000
        avg_profile['Wind_MW'] = avg_profile['Wind_Output_kW'] / 1000
        avg_profile['Hydro_MW'] = avg_profile['Hydro_Output_kW'] / 1000
        
        bess_capacity_kwh = results.get('bess_energy', 1) * 1000
        avg_profile['BESS_SOC_%'] = (avg_profile['BESS_SOC_kWh'] / bess_capacity_kwh) * 100
    else:
        hours = list(range(24))
        avg_profile = pd.DataFrame({
            'Hour_of_Day': hours,
            'Load_MW': [5] * 24,
            'PV_MW': [0] * 24,
            'Wind_MW': [0] * 24,
            'Hydro_MW': [0] * 24,
            'BESS_SOC_%': [50] * 24
        })
    
    fig = make_subplots(specs=[[{"secondary_y": True}]])
    
    fig.add_trace(go.Scatter(x=avg_profile['Hour_of_Day'], y=avg_profile['Hydro_MW'],
                            name='Hydro', mode='lines', line=dict(width=0),
                            stackgroup='one', fillcolor='rgba(141, 211, 199, 0.7)'),
                 secondary_y=False)
    
    fig.add_trace(go.Scatter(x=avg_profile['Hour_of_Day'], y=avg_profile['PV_MW'],
                            name='PV', mode='lines', line=dict(width=0),
                            stackgroup='one', fillcolor='rgba(253, 180, 98, 0.7)'),
                 secondary_y=False)
    
    fig.add_trace(go.Scatter(x=avg_profile['Hour_of_Day'], y=avg_profile['Wind_MW'],
                            name='Wind', mode='lines', line=dict(width=0),
                            stackgroup='one', fillcolor='rgba(128, 177, 211, 0.7)'),
                 secondary_y=False)
    
    fig.add_trace(go.Scatter(x=avg_profile['Hour_of_Day'], y=avg_profile['Load_MW'],
                            name='Load', mode='lines',
                            line=dict(width=3, color='#FFD700')),
                 secondary_y=False)
    
    fig.add_trace(go.Scatter(x=avg_profile['Hour_of_Day'], y=avg_profile['BESS_SOC_%'],
                            name='BESS SOC', mode='lines',
                            line=dict(width=3, color='#FF00FF')),
                 secondary_y=True)
    
    fig.update_xaxes(title_text="Hours", tickmode='linear', tick0=0, dtick=2, range=[0, 24])
    fig.update_yaxes(title_text="Power (MW)", secondary_y=False)
    fig.update_yaxes(title_text="BESS SOC (%)", secondary_y=True, range=[0, 120])
    
    fig.update_layout(
        title='Typical Daily Dispatch Profile',
        hovermode='x unified',
        height=500,
        showlegend=True
    )
    
    return fig


# ==============================================================================
# STREAMLIT APP
# ==============================================================================

st.set_page_config(
    page_title="RE Optimization Tool",
    page_icon="🌞",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.markdown('<p style="font-size:2.5rem;font-weight:bold;color:#1f77b4;text-align:center">🌞 Renewable Energy Optimization Tool</p>', unsafe_allow_html=True)
st.markdown("**Hybrid System Designer: PV + Wind + Hydro + Battery Storage**")
st.markdown("---")

# Initialize session state
if 'optimization_complete' not in st.session_state:
    st.session_state.optimization_complete = False
if 'results' not in st.session_state:
    st.session_state.results = None

# ========== SIDEBAR ==========
with st.sidebar:
    st.header("⚙️ System Configuration")
    
    # Component Selection
    st.subheader("🔌 Component Selection")
    st.markdown("**Select components to include:**")
    
    col1, col2 = st.columns(2)
    with col1:
        enable_pv = st.checkbox("☀️ Solar PV", value=True, key="enable_pv")
        enable_wind = st.checkbox("💨 Wind", value=True, key="enable_wind")
    with col2:
        enable_hydro = st.checkbox("💧 Hydro", value=True, key="enable_hydro")
        enable_bess = st.checkbox("🔋 BESS", value=True, key="enable_bess")
    
    if not any([enable_pv, enable_wind, enable_hydro, enable_bess]):
        st.error("⚠️ At least one component must be enabled!")
    
    st.markdown("---")
    
    # Solar PV
    with st.expander("☀️ SOLAR PV", expanded=enable_pv):
        if not enable_pv:
            st.warning("⚠️ Solar PV is DISABLED")
            pv_min = 0.0
            pv_max = 0.0
            pv_step = 1.0
            pv_capex = 1000
            pv_opex = 10
            pv_lifetime = 25
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
    
    # Wind
    with st.expander("💨 WIND"):
        if not enable_wind:
            st.warning("⚠️ Wind is DISABLED")
            wind_min = 0.0
            wind_max = 0.0
            wind_step = 1.0
            wind_capex = 1200
            wind_opex = 15
            wind_lifetime = 20
        else:
            st.subheader("Capacity Range")
            col1, col2 = st.columns(2)
            with col1:
                wind_min = st.number_input("Min (MW)", value=0.0, min_value=0.0, step=0.5, key="wind_min")
            with col2:
                wind_max = st.number_input("Max (MW)", value=3.0, min_value=0.0, step=0.5, key="wind_max")
            wind_step = st.number_input("Step (MW)", value=1.0, min_value=0.1, step=0.1, key="wind_step")
            
            st.subheader("Financial Parameters")
            col1, col2 = st.columns(2)
            with col1:
                wind_capex = st.number_input("CapEx ($/kW)", value=1200, step=10, key="wind_capex")
                wind_opex = st.number_input("OpEx ($/kW/yr)", value=15, step=1, key="wind_opex")
            with col2:
                wind_lifetime = st.number_input("Lifetime (years)", value=20, step=1, key="wind_life")
    
    # Hydro
    with st.expander("💧 HYDRO"):
        if not enable_hydro:
            st.warning("⚠️ Hydro is DISABLED")
            hydro_min = 0.0
            hydro_max = 0.0
            hydro_step = 1.0
            hydro_hours_per_day = 8
            hydro_capex = 2000
            hydro_opex = 20
            hydro_lifetime = 50
        else:
            st.subheader("Capacity Range")
            col1, col2 = st.columns(2)
            with col1:
                hydro_min = st.number_input("Min (MW)", value=0.0, min_value=0.0, step=0.5, key="hydro_min")
            with col2:
                hydro_max = st.number_input("Max (MW)", value=2.0, min_value=0.0, step=0.5, key="hydro_max")
            hydro_step = st.number_input("Step (MW)", value=1.0, min_value=0.1, step=0.1, key="hydro_step")
            
            st.subheader("Operating Configuration")
            hydro_hours_per_day = st.number_input("Operating Hours (hours/day)", value=8, min_value=1, max_value=24, step=1, key="hydro_hours")
            
            st.subheader("Financial Parameters")
            col1, col2 = st.columns(2)
            with col1:
                hydro_capex = st.number_input("CapEx ($/kW)", value=2000, step=10, key="hydro_capex")
                hydro_opex = st.number_input("OpEx ($/kW/yr)", value=20, step=1, key="hydro_opex")
            with col2:
                hydro_lifetime = st.number_input("Lifetime (years)", value=50, step=1, key="hydro_life")
    
    # BESS
    with st.expander("🔋 BATTERY STORAGE"):
        if not enable_bess:
            st.warning("⚠️ BESS is DISABLED")
            bess_min = 0.0
            bess_max = 0.0
            bess_step = 1.0
            bess_duration = 4.0
            bess_min_soc = 20
            bess_max_soc = 100
            bess_charge_eff = 95
            bess_discharge_eff = 95
            bess_lifetime = 15
            bess_power_capex = 300
            bess_energy_capex = 200
            bess_opex = 2
        else:
            st.subheader("Power Range")
            col1, col2 = st.columns(2)
            with col1:
                bess_min = st.number_input("Min (MW)", value=5.0, min_value=0.0, step=1.0, key="bess_min")
            with col2:
                bess_max = st.number_input("Max (MW)", value=20.0, min_value=0.0, step=1.0, key="bess_max")
            bess_step = st.number_input("Step (MW)", value=5.0, min_value=0.1, step=0.1, key="bess_step")
            
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
                bess_opex = st.number_input("OpEx ($/kW/yr)", value=2, step=1, key="bess_opex")
    
    # Financial Parameters
    with st.expander("💰 FINANCIAL PARAMETERS"):
        discount_rate = st.number_input("Discount Rate (%)", value=8.0, min_value=0.0, max_value=20.0, step=0.5)
        inflation_rate = st.number_input("Inflation Rate (%)", value=2.0, min_value=0.0, max_value=10.0, step=0.5)
        project_lifetime = st.number_input("Project Lifetime (years)", value=25, min_value=1, max_value=50, step=1)
    
    # Optimization Settings
    with st.expander("🎯 OPTIMIZATION SETTINGS"):
        target_unmet_percent = st.number_input("Target Unmet Load (%)", value=0.1, min_value=0.0, max_value=5.0, step=0.1, key="target_unmet")
    
    # File Uploads
    st.header("📁 Upload Profiles")
    load_file = st.file_uploader("Load Profile", type=['csv', 'xlsx'], key="load_file")
    
    if enable_pv:
        pv_file = st.file_uploader("PV Profile (1 kW)", type=['csv', 'xlsx'], key="pv_file")
    else:
        pv_file = None
        st.caption("⚠️ PV profile not required (PV disabled)")
    
    if enable_wind:
        wind_file = st.file_uploader("Wind Profile (1 kW)", type=['csv', 'xlsx'], key="wind_file")
    else:
        wind_file = None
        st.caption("⚠️ Wind profile not required (Wind disabled)")
    
    if enable_hydro:
        hydro_file = st.file_uploader("Hydro Profile (Optional)", type=['csv', 'xlsx'], key="hydro_file")
    else:
        hydro_file = None
        st.caption("⚠️ Hydro profile not required (Hydro disabled)")

# ========== MAIN TABS ==========
tab1, tab2, tab3 = st.tabs(["🏠 Home", "⚙️ Optimize", "📊 Results"])

# Calculate search space
pv_options = int((pv_max - pv_min) / pv_step) + 1 if pv_step > 0 and enable_pv else 1
wind_options = int((wind_max - wind_min) / wind_step) + 1 if wind_step > 0 and enable_wind else 1
hydro_options = int((hydro_max - hydro_min) / hydro_step) + 1 if hydro_step > 0 and enable_hydro else 1
bess_options = int((bess_max - bess_min) / bess_step) + 1 if bess_step > 0 and enable_bess else 1
total_combinations = pv_options * wind_options * hydro_options * bess_options

# TAB 1: HOME
with tab1:
    st.header("Welcome to the Renewable Energy Optimization Tool")
    st.markdown("""
    ### 🎯 Purpose
    Optimize hybrid renewable energy systems to minimize Net Present Cost while meeting reliability targets.
    
    ### 🔌 Flexible System Design
    Enable/disable any combination of components:
    - **PV + BESS** (Solar + Storage)
    - **PV + Wind + BESS** (Hybrid Solar-Wind)
    - **PV + Wind + Hydro + BESS** (Full Hybrid)
    - And many more!
    """)
    
    # Active Configuration
    active_components = []
    if enable_pv:
        active_components.append("☀️ Solar PV")
    if enable_wind:
        active_components.append("💨 Wind")
    if enable_hydro:
        active_components.append("💧 Hydro")
    if enable_bess:
        active_components.append("🔋 BESS")
    
    if len(active_components) > 0:
        st.success(f"**Enabled Components:** {' + '.join(active_components)}")
    
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        if enable_pv:
            st.metric("PV Options", f"{pv_options}", f"{pv_min}-{pv_max} MW")
        else:
            st.metric("PV", "Disabled", "❌")
    with col2:
        if enable_wind:
            st.metric("Wind Options", f"{wind_options}", f"{wind_min}-{wind_max} MW")
        else:
            st.metric("Wind", "Disabled", "❌")
    with col3:
        if enable_hydro:
            st.metric("Hydro Options", f"{hydro_options}", f"{hydro_min}-{hydro_max} MW")
        else:
            st.metric("Hydro", "Disabled", "❌")
    with col4:
        if enable_bess:
            st.metric("BESS Options", f"{bess_options}", f"{bess_min}-{bess_max} MW")
        else:
            st.metric("BESS", "Disabled", "❌")
    
    st.info(f"**Total Search Space:** {total_combinations:,} combinations")

# TAB 2: OPTIMIZE
with tab2:
    st.header("⚙️ Run Optimization")
    
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        st.metric("Search Space", f"{total_combinations:,}")
    with col2:
        est_time = max(1, total_combinations * 0.05 / 60)
        st.metric("Est. Time", f"{est_time:.1f} min")
    with col3:
        num_enabled = sum([enable_pv, enable_wind, enable_hydro, enable_bess])
        st.metric("Components", f"{num_enabled}")
    with col4:
        st.metric("Target Unmet", f"{target_unmet_percent}%")
    
    st.markdown("---")
    
    # Validation
    validation_passed = True
    validation_messages = []
    
    if not any([enable_pv, enable_wind, enable_hydro, enable_bess]):
        validation_passed = False
        validation_messages.append("❌ At least one component must be enabled")
    
    if not load_file:
        validation_passed = False
        validation_messages.append("❌ Load Profile required")
    
    if enable_pv and not pv_file:
        validation_passed = False
        validation_messages.append("❌ PV Profile required (PV is enabled)")
    
    if enable_wind and not wind_file:
        validation_passed = False
        validation_messages.append("❌ Wind Profile required (Wind is enabled)")
    
    if not validation_passed:
        for msg in validation_messages:
            st.error(msg)
    else:
        st.success("✅ All inputs validated. Ready to optimize!")
    
    if st.button("▶️ RUN OPTIMIZATION", type="primary", disabled=not validation_passed, use_container_width=True):
        
        if not OPTIMIZATION_AVAILABLE:
            st.error("❌ Optimization code not available")
        else:
            try:
                st.subheader("🔄 Optimization in Progress...")
                progress_bar = st.progress(0)
                status_text = st.empty()
                
                # Read profiles
                status_text.text("📊 Reading profiles...")
                progress_bar.progress(10)
                
                load_df = pd.read_csv(load_file) if load_file.name.endswith('.csv') else pd.read_excel(load_file)
                
                if enable_pv and pv_file:
                    pv_df = pd.read_csv(pv_file) if pv_file.name.endswith('.csv') else pd.read_excel(pv_file)
                else:
                    pv_df = pd.DataFrame({'Hour': range(8760), 'Output_kW': [0.0] * 8760})
                
                if enable_wind and wind_file:
                    wind_df = pd.read_csv(wind_file) if wind_file.name.endswith('.csv') else pd.read_excel(wind_file)
                else:
                    wind_df = pd.DataFrame({'Hour': range(8760), 'Output_kW': [0.0] * 8760})
                
                if enable_hydro and hydro_file:
                    hydro_df = pd.read_csv(hydro_file) if hydro_file.name.endswith('.csv') else pd.read_excel(hydro_file)
                else:
                    hydro_df = pd.DataFrame({'Hour': range(8760), 'Output_kW': [1.0] * 8760})
                
                # Build Excel input
                status_text.text("🔨 Building input file...")
                progress_bar.progress(15)
                
                output = BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    # Configuration
                    pd.DataFrame({
                        'Parameter': ['Simulation Hours', 'Target Unmet Load (%)', 'Optimization Method',
                                     'Discount Rate (%)', 'Inflation Rate (%)', 'Project Lifetime (years)', 'Use Dynamic LCOE'],
                        'Value': [8760, target_unmet_percent, 'GRID_SEARCH', discount_rate, inflation_rate, project_lifetime, 'NO']
                    }).to_excel(writer, sheet_name='Configuration', index=False)
                    
                    # Grid Search Config
                    pd.DataFrame({
                        'Parameter': ['Enable Grid Search', 'PV Search Start', 'PV Search End', 'PV Search Step',
                                     'Wind Search Start', 'Wind Search End', 'Wind Search Step',
                                     'Hydro Search Start', 'Hydro Search End', 'Hydro Search Step',
                                     'BESS Search Start', 'BESS Search End', 'BESS Search Step',
                                     'Max Combinations', 'Optimization Objective', 'Show Top N Solutions'],
                        'Value': ['YES', pv_min*1000, pv_max*1000, pv_step*1000, 
                                 wind_min*1000, wind_max*1000, wind_step*1000,
                                 hydro_min*1000, hydro_max*1000, hydro_step*1000, 
                                 bess_min*1000, bess_max*1000, bess_step*1000,
                                 100000000, 'NPC', 5]
                    }).to_excel(writer, sheet_name='Grid_Search_Config', index=False)
                    
                    # Component sheets (use dummy LCOE value 0, will be calculated from NPC)
                    pd.DataFrame({
                        'Parameter': ['LCOE', 'PVsyst Baseline', 'Capex', 'O&M Cost', 'Lifetime'],
                        'Value': [0, 1.0, pv_capex, pv_opex, pv_lifetime]
                    }).to_excel(writer, sheet_name='Solar_PV', index=False)
                    
                    pd.DataFrame({
                        'Parameter': ['Include Wind?', 'LCOE', 'Capex', 'O&M Cost', 'Lifetime'],
                        'Value': ['YES' if enable_wind else 'NO', 0, wind_capex, wind_opex, wind_lifetime]
                    }).to_excel(writer, sheet_name='Wind', index=False)
                    
                    pd.DataFrame({
                        'Parameter': ['Include Hydro?', 'LCOE', 'Capex', 'O&M Cost', 'Lifetime', 'Operating Hours'],
                        'Value': ['YES' if enable_hydro else 'NO', 0, hydro_capex, hydro_opex, hydro_lifetime, hydro_hours_per_day]
                    }).to_excel(writer, sheet_name='Hydro', index=False)
                    
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
                    hydro_df.to_excel(writer, sheet_name='Hydro_Profile', index=False)
                
                output.seek(0)
                
                # Save temp file
                import tempfile
                temp_file = os.path.join(tempfile.gettempdir(), "temp_input_generated.xlsx")
                with open(temp_file, "wb") as f:
                    f.write(output.getvalue())
                
                # Run optimization
                status_text.text("⚙️ Running optimization...")
                progress_bar.progress(30)
                
                opt_module.INPUT_FILE = temp_file
                result = opt_module.read_inputs()
                
                if len(result) == 9:
                    config, grid_config, solar, wind, hydro, bess, load_profile, pvsyst_profile, wind_profile = result
                else:
                    config, grid_config, solar, wind, hydro, bess, load_profile, pvsyst_profile, wind_profile = result[:9]
                
                results_df = opt_module.grid_search_optimize_hydro(
                    config, grid_config, solar, wind, hydro, bess,
                    load_profile, pvsyst_profile, wind_profile, None
                )
                
                progress_bar.progress(85)
                optimal = opt_module.find_optimal_solution(results_df)
                progress_bar.progress(90)
                
                if optimal is not None:
                    # Calculate electrical metrics
                    optimal_dispatch = opt_module.calculate_dispatch_with_hydro(
                        load_profile, pvsyst_profile, wind_profile,
                        optimal['PV_kW'], optimal['Wind_kW'], optimal['Hydro_kW'],
                        optimal['BESS_Power_kW'], optimal['BESS_Capacity_kWh'],
                        solar, wind, hydro, bess,
                        int(optimal['Hydro_Window_Start']), int(optimal['Hydro_Window_End'])
                    )
                    
                    component_capacities = {
                        'pv_kw': optimal['PV_kW'],
                        'wind_kw': optimal['Wind_kW'],
                        'hydro_kw': optimal['Hydro_kW'],
                        'bess_kwh': optimal['BESS_Capacity_kWh']
                    }
                    
                    component_configs = {
                        'pv_lcoe': 0,
                        'wind_lcoe': 0,
                        'hydro_lcoe': 0,
                        'bess_max_soc': bess_max_soc / 100,
                        'bess_min_soc': bess_min_soc / 100,
                        'bess_lifetime': bess_lifetime
                    }
                    
                    # Get NPC data for LCOE calculation
                    npc_breakdown = {
                        'pv': {'npc': optimal.get('PV_NPC_$', 0)},
                        'wind': {'npc': optimal.get('Wind_NPC_$', 0)},
                        'hydro': {'npc': optimal.get('Hydro_NPC_$', 0)},
                        'bess': {'npc': optimal.get('BESS_NPC_$', 0)}
                    }
                    
                    electrical_metrics = opt_module.calculate_electrical_metrics(
                        optimal_dispatch, component_capacities, component_configs,
                        npc_breakdown, project_lifetime
                    )
                    
                    progress_bar.progress(100)
                    
                    # Clean up
                    if os.path.exists(temp_file):
                        os.remove(temp_file)
                    
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
                        'results_df': results_df,
                        'optimal_row': optimal.to_dict(),
                        'electrical_metrics': electrical_metrics,
                        'optimal_dispatch': optimal_dispatch,
                        'config_params': {
                            'discount_rate': discount_rate,
                            'inflation_rate': inflation_rate,
                            'project_lifetime': project_lifetime,
                            'target_unmet_percent': target_unmet_percent
                        }
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

# TAB 3: RESULTS
with tab3:
    if not st.session_state.optimization_complete:
        st.info("ℹ️ No results yet. Run optimization first.")
    else:
        st.header("📊 Optimization Results")
        
        results = st.session_state.results
        optimal_row = results.get('optimal_row', {})
        
        # Key Metrics
        col1, col2, col3 = st.columns(3)
        with col1:
            st.metric("Total NPC", f"${results['npc']/1e6:.2f}M")
        with col2:
            st.metric("System LCOE", f"${results['lcoe']:.2f}/MWh")
        with col3:
            st.metric("Unmet Load", f"{results['unmet_pct']:.3f}%")
        
        # Component Configuration
        col1, col2, col3, col4 = st.columns(4)
        with col1:
            st.markdown("**☀️ Solar PV**")
            st.metric("Capacity", f"{results['pv_capacity']:.2f} MW")
            if 'electrical_metrics' in results and results['electrical_metrics']:
                pv_lcoe = results['electrical_metrics']['pv']['levelized_cost_per_kwh'] * 1000
                st.metric("LCOE", f"${pv_lcoe:.2f}/MWh")
        
        with col2:
            st.markdown("**💨 Wind**")
            st.metric("Capacity", f"{results['wind_capacity']:.2f} MW")
            if 'electrical_metrics' in results and results['electrical_metrics']:
                wind_lcoe = results['electrical_metrics']['wind']['levelized_cost_per_kwh'] * 1000
                st.metric("LCOE", f"${wind_lcoe:.2f}/MWh")
        
        with col3:
            st.markdown("**💧 Hydro**")
            st.metric("Capacity", f"{results['hydro_capacity']:.2f} MW")
            if 'electrical_metrics' in results and results['electrical_metrics']:
                hydro_lcoe = results['electrical_metrics']['hydro']['levelized_cost_per_kwh'] * 1000
                st.metric("LCOE", f"${hydro_lcoe:.2f}/MWh")
        
        with col4:
            st.markdown("**🔋 Battery**")
            st.metric("Power", f"{results['bess_power']:.2f} MW")
            st.metric("Energy", f"{results['bess_energy']:.2f} MWh")
        
        st.markdown("---")
        
        # Cost Analysis
        st.subheader("💰 Cost Analysis")
        cost_charts = create_cost_analysis_charts_with_tables(results, optimal_row)
        
        col1, col2 = st.columns(2)
        with col1:
            st.markdown("**By Component**")
            st.plotly_chart(cost_charts['npc_by_component']['chart'], use_container_width=True)
            st.dataframe(cost_charts['npc_by_component']['table'], use_container_width=True, hide_index=True)
        
        with col2:
            st.markdown("**By Cost Type**")
            st.plotly_chart(cost_charts['npc_by_cost_type']['chart'], use_container_width=True)
            st.dataframe(cost_charts['npc_by_cost_type']['table'], use_container_width=True, hide_index=True)
        
        st.markdown("---")
        
        # Cash Flow
        st.subheader("💵 Cash Flow Analysis")
        cash_flow_fig = create_fixed_cash_flow_chart(results, optimal_row)
        st.plotly_chart(cash_flow_fig, use_container_width=True)
        
        st.markdown("---")
        
        # Electrical Metrics
        st.subheader("⚡ Electrical Performance Metrics")
        
        if 'electrical_metrics' in results and results['electrical_metrics']:
            elec_tables = create_electrical_metrics_tables(
                results['electrical_metrics'],
                bess_power_mw=results.get('bess_power', 0),
                bess_capacity_mwh=results.get('bess_energy', 0)
            )
            
            col1, col2 = st.columns(2)
            with col1:
                st.markdown("**☀️ Solar PV**")
                st.dataframe(elec_tables['pv'], use_container_width=True, hide_index=True)
            
            with col2:
                st.markdown("**💨 Wind**")
                st.dataframe(elec_tables['wind'], use_container_width=True, hide_index=True)
            
            col1, col2 = st.columns(2)
            with col1:
                st.markdown("**💧 Hydro**")
                st.dataframe(elec_tables['hydro'], use_container_width=True, hide_index=True)
            
            with col2:
                st.markdown("**🔋 Battery Storage - Performance**")
                st.dataframe(elec_tables['bess_performance'], use_container_width=True, hide_index=True)
            
            st.markdown("---")
            st.markdown("**🏗️ Battery Storage - Deployment (Sungrow PowerTitan 2.0)**")
            st.dataframe(elec_tables['bess_deployment'], use_container_width=True, hide_index=True)
        
        st.markdown("---")
        
        # Typical Daily Profile
        st.subheader("📈 Typical Daily Dispatch Profile")
        daily_profile_fig = create_typical_daily_dispatch_profile(results)
        st.plotly_chart(daily_profile_fig, use_container_width=True)
        
        st.markdown("---")
        
        # Download
        st.subheader("📥 Download Results")
        excel_output = export_results_industry_format(
            results, results['results_df'], results['optimal_row'], 
            results['config_params']
        )
        
        col1, col2, col3 = st.columns([2, 1, 2])
        with col2:
            st.download_button(
                label="📥 Download Excel",
                data=excel_output,
                file_name=f"results_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
                type="primary"
            )

# Footer
st.markdown("---")
st.markdown('<div style="text-align:center;color:#666"><p>RE Optimization Tool v3.0 | Professional NPC Analysis</p></div>', unsafe_allow_html=True)
