"""
RENEWABLE ENERGY OPTIMIZATION TOOL - COMPLETE VERSION WITH DEGRADATION
=======================================================================
Version: 4.0 - Integrated Degradation Analysis
Features:
- Full degradation analysis with 25-year hourly simulation
- Independent PV/BESS degradation selection
- Multi-year hourly export (Years 1, 2, 5, 10, 15, 20, 25)
- Fixed BESS SOC initialization with Year 1 degradation
- Enhanced Excel export with degradation sheets

Author: SJ
Date: February 2026
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

# Import degradation analysis - FIXED VERSION
try:
    import optimize_with_degradation_FIXED as deg_module
    DEGRADATION_AVAILABLE = True
except ImportError:
    try:
        import optimize_with_degradation as deg_module
        DEGRADATION_AVAILABLE = True
    except ImportError:
        DEGRADATION_AVAILABLE = False
        print("⚠️ Warning: Degradation module not found")

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
    container_height_m = 2.896
    
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
# EXCEL EXPORT WITH DEGRADATION SUPPORT
# ==============================================================================

def export_results_industry_format(results_dict, results_df, optimal_row, 
                                   config_params, degradation_results=None):
    """
    Export results in industry standard Excel format.
    NOW INCLUDES DEGRADATION ANALYSIS SHEETS.
    """
    output = BytesIO()
    
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        # Sheet 1: Summary
        summary_data = []
        
        # Check if degradation was applied
        if degradation_results and 'degradation_applied' in degradation_results:
            deg_settings = degradation_results['degradation_applied']
            deg_status = []
            if deg_settings['pv']:
                deg_status.append("PV")
            if deg_settings['bess']:
                deg_status.append("BESS")
            deg_text = " + ".join(deg_status) if deg_status else "None"
        else:
            deg_text = "None"
        
        summary_data.extend([
            ['Parameter', 'Value'],
            ['Optimization Method', 'GRID_SEARCH'],
            ['NPC Calculation Method', 'Present Value Analysis'],
            ['Degradation Analysis', deg_text],
            ['', ''],
            ['Target Unmet Load (%)', config_params.get('target_unmet_percent', 0.1)],
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
        ])
        
        # Add degradation summary if available
        if degradation_results:
            summary_data.extend([
                ['--- DEGRADATION ANALYSIS ---', ''],
                ['Year 1 NPC ($)', degradation_results.get('npc_year1', 0)],
                ['25-Year NPC ($)', degradation_results.get('npc_25year', 0)],
                ['BESS Replacement Cost PV ($)', degradation_results.get('replacement_cost_pv', 0)],
                ['Year 1 LCOE ($/MWh)', degradation_results.get('lcoe_year1', 0)],
                ['25-Year LCOE ($/MWh)', degradation_results.get('lcoe_25year', 0)],
                ['PV Total Degradation (%)', degradation_results.get('pv_deg_total', 0)],
                ['BESS Capacity Loss 20Y (%)', degradation_results.get('bess_loss_20y', 0)],
                ['', ''],
            ])
        else:
            summary_data.extend([
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
            'Replacement ($)': [
                optimal_row.get('PV_Replacement_$', 0),
                optimal_row.get('Wind_Replacement_$', 0),
                optimal_row.get('Hydro_Replacement_$', 0),
                optimal_row.get('BESS_Replacement_$', 0),
                optimal_row.get('Replacement_$', 0)
            ],
            'OM ($)': [
                optimal_row.get('PV_OM_$', 0),
                optimal_row.get('Wind_OM_$', 0),
                optimal_row.get('Hydro_OM_$', 0),
                optimal_row.get('BESS_OM_$', 0),
                optimal_row.get('OM_$', 0)
            ],
            'Salvage ($)': [
                optimal_row.get('PV_Salvage_$', 0),
                optimal_row.get('Wind_Salvage_$', 0),
                optimal_row.get('Hydro_Salvage_$', 0),
                optimal_row.get('BESS_Salvage_$', 0),
                optimal_row.get('Salvage_$', 0)
            ],
            'NPC ($)': [
                optimal_row.get('PV_NPC_$', 0),
                optimal_row.get('Wind_NPC_$', 0),
                optimal_row.get('Hydro_NPC_$', 0),
                optimal_row.get('BESS_NPC_$', 0),
                optimal_row.get('NPC_$', 0)
            ],
            'Annualized ($/yr)': [
                optimal_row.get('PV_Annualized_$/yr', 0),
                optimal_row.get('Wind_Annualized_$/yr', 0),
                optimal_row.get('Hydro_Annualized_$/yr', 0),
                optimal_row.get('BESS_Annualized_$/yr', 0),
                optimal_row.get('Annualized_$/yr', 0)
            ]
        })
        cost_breakdown.to_excel(writer, sheet_name='Cost_Breakdown', index=False)
        
        # Sheet 3: All Results
        results_df.to_excel(writer, sheet_name='All_Results', index=False)
        
        # Sheet 4+: Degradation Analysis Sheets
        if degradation_results and 'hourly_dispatch' in degradation_results:
            # Export yearly summary
            if 'yearly_summary' in degradation_results:
                degradation_results['yearly_summary'].to_excel(
                    writer, sheet_name='Degradation_25Years', index=False
                )
            
            # Export hourly dispatch for selected years
            hourly_dispatch = degradation_results['hourly_dispatch']
            for year_key in sorted(hourly_dispatch.keys()):
                year_num = year_key.split('_')[1]
                sheet_name = f'Year_{year_num}_Hourly'
                hourly_dispatch[year_key].to_excel(writer, sheet_name=sheet_name, index=False)
        
        # Fallback: Standard Year 1 dispatch if no degradation
        elif 'optimal_dispatch' in results_dict:
            results_dict['optimal_dispatch'].to_excel(writer, sheet_name='Year_1_Hourly', index=False)
    
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
    """Create simplified cash flow chart."""
    project_lifetime = results.get('config_params', {}).get('project_lifetime', 25)
    years = list(range(0, project_lifetime + 1))
    
    capital_flow = [0] * len(years)
    replacement_flow = [0] * len(years)
    salvage_flow = [0] * len(years)
    
    capital_flow[0] = -optimal_row.get('Capital_$', 0) / 1e6
    
    total_replacement = optimal_row.get('Replacement_$', 0) / 1e6
    bess_lifetime = 15
    
    for year in range(bess_lifetime, project_lifetime, bess_lifetime):
        replacement_flow[year] = -total_replacement
    
    salvage_flow[-1] = optimal_row.get('Salvage_$', 0) / 1e6
    
    fig = go.Figure()
    fig.add_trace(go.Bar(name='Capital', x=years, y=capital_flow, marker_color='#2E7D32'))
    fig.add_trace(go.Bar(name='Replacement', x=years, y=replacement_flow, marker_color='#1976D2'))
    fig.add_trace(go.Bar(name='Salvage', x=years, y=salvage_flow, marker_color='#388E3C'))
    
    fig.update_layout(
        title='Nominal Cash Flow (Major Costs Only)',
        xaxis_title='Year',
        yaxis_title='Cash Flow ($M)',
        barmode='relative',
        height=450,
        showlegend=True,
        annotations=[
            dict(
                text="Note: Annual O&M costs not shown (too small to visualize effectively)",
                xref="paper", yref="paper",
                x=0.5, y=-0.15,
                showarrow=False,
                font=dict(size=10, color="gray")
            )
        ]
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
# SINGLE DAY DISPATCH PROFILE
# ==============================================================================

def create_single_day_dispatch_profile(results):
    """Create dispatch profile for a SINGLE REPRESENTATIVE DAY."""
    if 'optimal_dispatch' in results:
        dispatch_df = results['optimal_dispatch'].copy()

        # ─────────────────────────────────────────────────────────────────────
        # COLUMN DETECTION
        # We need THREE things:
        #   1. pv_available_col  → TOTAL PV generated (for the area fill)
        #   2. wind_col / hydro_col → other generation
        #   3. day-selection col  → use AVAILABLE PV to find median day
        # ─────────────────────────────────────────────────────────────────────
        if 'PV_Available_kW' in dispatch_df.columns:
            # Degradation/Anaconda format  ← your current output
            pv_available_col = 'PV_Available_kW'
            wind_col  = 'Wind_Output_kW'  if 'Wind_Output_kW'  in dispatch_df.columns else None
            hydro_col = 'Hydro_Output_kW' if 'Hydro_Output_kW' in dispatch_df.columns else None

        elif 'PV_Output_kW' in dispatch_df.columns:
            # Standard base-optimisation format
            pv_available_col = 'PV_Output_kW'
            wind_col  = 'Wind_Output_kW'  if 'Wind_Output_kW'  in dispatch_df.columns else None
            hydro_col = 'Hydro_Output_kW' if 'Hydro_Output_kW' in dispatch_df.columns else None

        elif 'PV_to_Load_kW' in dispatch_df.columns:
            # Fallback: only PV_to_Load available (less accurate but won't crash)
            pv_available_col = 'PV_to_Load_kW'
            wind_col  = None
            hydro_col = None
        else:
            pv_available_col = None
            wind_col  = None
            hydro_col = None

        # Fill missing wind/hydro with zeros
        if wind_col is None:
            dispatch_df['_wind_zero'] = 0
            wind_col = '_wind_zero'
        if hydro_col is None:
            dispatch_df['_hydro_zero'] = 0
            hydro_col = '_hydro_zero'
        if pv_available_col is None:
            dispatch_df['_pv_zero'] = 0
            pv_available_col = '_pv_zero'

        # ─────────────────────────────────────────────────────────────────────
        # DAY DETECTION
        # Hour column is 0-23 repeated (new format) or continuous 0-8759
        # ─────────────────────────────────────────────────────────────────────
        if dispatch_df['Hour'].max() <= 23:
            dispatch_df['Continuous_Hour'] = dispatch_df.index
            dispatch_df['Day'] = dispatch_df['Continuous_Hour'] // 24
        else:
            dispatch_df['Day'] = dispatch_df['Hour'] // 24

        # Pick the MEDIAN PV day using AVAILABLE PV (bell-curve shape)
        daily_pv = dispatch_df.groupby('Day')[pv_available_col].sum()
        median_pv_day = daily_pv.sort_values().index[len(daily_pv) // 2]

        # Extract that day
        day_profile = dispatch_df[dispatch_df['Day'] == median_pv_day].copy()

        # Hour of day axis
        if day_profile['Hour'].max() <= 23:
            day_profile['Hour_of_Day'] = day_profile['Hour']
        else:
            day_profile['Hour_of_Day'] = day_profile['Hour'] % 24

        # Convert to MW
        day_profile['Load_MW']  = day_profile['Load_kW']          / 1000
        day_profile['PV_MW']    = day_profile[pv_available_col]   / 1000   # ← AVAILABLE, not to-load
        day_profile['Wind_MW']  = day_profile[wind_col]           / 1000
        day_profile['Hydro_MW'] = day_profile[hydro_col]          / 1000

        # BESS SOC
        bess_capacity_kwh = results.get('bess_energy', 1) * 1000
        if bess_capacity_kwh > 0 and 'BESS_SOC_kWh' in day_profile.columns:
            day_profile['BESS_SOC_%'] = (day_profile['BESS_SOC_kWh'] / bess_capacity_kwh * 100).clip(0, 100)
        else:
            day_profile['BESS_SOC_%'] = 50

    else:
        hours = list(range(24))
        day_profile = pd.DataFrame({
            'Hour_of_Day': hours,
            'Load_MW':    [0.5] * 24,
            'PV_MW':      [0]   * 24,
            'Wind_MW':    [0]   * 24,
            'Hydro_MW':   [0]   * 24,
            'BESS_SOC_%': [50]  * 24,
        })
    
    fig = make_subplots(specs=[[{"secondary_y": True}]])
    
    fig.add_trace(go.Scatter(
        x=day_profile['Hour_of_Day'],
        y=day_profile['Hydro_MW'],
        name='Hydro',
        mode='lines',
        line=dict(width=0),
        stackgroup='one',
        fillcolor='rgba(141, 211, 199, 0.7)',
        hovertemplate='Hour %{x}<br>Hydro: %{y:.2f} MW<extra></extra>'
    ), secondary_y=False)
    
    fig.add_trace(go.Scatter(
        x=day_profile['Hour_of_Day'],
        y=day_profile['PV_MW'],
        name='PV',
        mode='lines',
        line=dict(width=0),
        stackgroup='one',
        fillcolor='rgba(253, 180, 98, 0.7)',
        hovertemplate='Hour %{x}<br>PV: %{y:.2f} MW<extra></extra>'
    ), secondary_y=False)
    
    fig.add_trace(go.Scatter(
        x=day_profile['Hour_of_Day'],
        y=day_profile['Wind_MW'],
        name='Wind',
        mode='lines',
        line=dict(width=0),
        stackgroup='one',
        fillcolor='rgba(128, 177, 211, 0.7)',
        hovertemplate='Hour %{x}<br>Wind: %{y:.2f} MW<extra></extra>'
    ), secondary_y=False)
    
    fig.add_trace(go.Scatter(
        x=day_profile['Hour_of_Day'],
        y=day_profile['Load_MW'],
        name='Load',
        mode='lines',
        line=dict(width=3, color='#FFD700'),
        hovertemplate='Hour %{x}<br>Load: %{y:.2f} MW<extra></extra>'
    ), secondary_y=False)
    
    fig.add_trace(go.Scatter(
        x=day_profile['Hour_of_Day'],
        y=day_profile['BESS_SOC_%'],
        name='BESS SOC',
        mode='lines',
        line=dict(width=3, color='#FF00FF'),
        hovertemplate='Hour %{x}<br>SOC: %{y:.1f}%<extra></extra>'
    ), secondary_y=True)
    
    fig.update_xaxes(
        title_text="Hours",
        tickmode='linear',
        tick0=0,
        dtick=2,
        range=[0, 24]
    )
    
    fig.update_yaxes(
        title_text="Power (MW)",
        secondary_y=False
    )
    
    fig.update_yaxes(
        title_text="BESS SOC (%)",
        secondary_y=True,
        range=[0, 120]
    )
    
    fig.update_layout(
        title='Single Day Dispatch Profile (Representative Day)',
        hovermode='x unified',
        height=500,
        showlegend=True,
        legend=dict(
            orientation="h",
            yanchor="bottom",
            y=1.02,
            xanchor="right",
            x=1
        )
    )
    
    return fig


# ==============================================================================
# ENERGY MIX PIE CHART
# ==============================================================================

def create_energy_mix_pie_chart(optimal_row):
    """Create energy production mix pie chart."""
    pv_energy = optimal_row.get('PV_Energy_kWh', 0) / 1000
    wind_energy = optimal_row.get('Wind_Energy_kWh', 0) / 1000
    hydro_energy = optimal_row.get('Hydro_Energy_kWh', 0) / 1000
    
    total_energy = pv_energy + wind_energy + hydro_energy
    
    if total_energy > 0:
        pv_pct = (pv_energy / total_energy) * 100
        wind_pct = (wind_energy / total_energy) * 100
        hydro_pct = (hydro_energy / total_energy) * 100
    else:
        pv_pct = wind_pct = hydro_pct = 0
    
    fig = go.Figure(data=[go.Pie(
        labels=['PV', 'Wind', 'Hydro'],
        values=[pv_energy, wind_energy, hydro_energy],
        marker=dict(colors=['#FDB462', '#80B1D3', '#8DD3C7']),
        textinfo='label+percent',
        hovertemplate='%{label}<br>%{value:.1f} MWh/yr<br>%{percent}<extra></extra>',
        hole=0.3
    )])
    
    fig.update_layout(
        title='Annual Energy Production Mix',
        height=400
    )
    
    energy_table = pd.DataFrame({
        'Component': ['PV', 'Wind', 'Hydro', 'Total'],
        'Energy (MWh/yr)': [
            f"{pv_energy:,.1f}",
            f"{wind_energy:,.1f}",
            f"{hydro_energy:,.1f}",
            f"{total_energy:,.1f}"
        ],
        'Percentage (%)': [
            f"{pv_pct:.1f}%",
            f"{wind_pct:.1f}%",
            f"{hydro_pct:.1f}%",
            "100.0%"
        ]
    })
    
    return fig, energy_table


# ==============================================================================
# EMISSIONS TABLE
# ==============================================================================

def create_emissions_table(optimal_row, results):
    """Create emissions summary table (placeholder for future implementation)."""
    # Placeholder - emissions calculations will be implemented later
    emissions_data = pd.DataFrame({
        'Pollutant': ['Carbon Dioxide', 'Carbon Monoxide', 'Unburned Hydrocarbons', 
                     'Particulate Matter', 'Sulfur Dioxide', 'Nitrogen Oxides'],
        'Quantity': [0, 0, 0, 0, 0, 0],
        'Units': ['kg/yr', 'kg/yr', 'kg/yr', 'kg/yr', 'kg/yr', 'kg/yr']
    })
    
    return emissions_data


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
st.markdown("**Hybrid System Designer: PV + Wind + Hydro + Battery Storage + CCGT + Hydrogen**")
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
        enable_hydro = st.checkbox("💧 Hydro", value=True, key="enable_hydro")
    with col2:
        enable_bess = st.checkbox("🔋 BESS", value=True, key="enable_bess")
        enable_ccgt = st.checkbox("🔥 CCGT", value=False, disabled=True, key="enable_ccgt", 
                                  help="Coming soon - Combined Cycle Gas Turbine")
        enable_hydrogen = st.checkbox("⚗️ Hydrogen", value=False, disabled=True, key="enable_hydrogen",
                                     help="Coming soon - Hydrogen production/storage")
    
    if not any([enable_pv, enable_wind, enable_hydro, enable_bess]):
        st.error("⚠️ At least one component must be enabled!")
    
    st.markdown("---")
    
    # ==============================================================================
    # SOLAR PV - WITH DEGRADATION CHECKBOX AND PROFILE UPLOAD
    # ==============================================================================
    with st.expander("☀️ SOLAR PV", expanded=enable_pv):
        if not enable_pv:
            st.warning("⚠️ Solar PV is DISABLED")
            pv_min = 0.0
            pv_max = 0.0
            pv_step = 1.0
            pv_capex = 1000
            pv_opex = 10
            pv_lifetime = 25
            pv_file = None
            apply_pv_degradation = False
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
            
            # DEGRADATION CHECKBOX (IN COMPONENT TAB)
            st.markdown("---")
            apply_pv_degradation = st.checkbox(
                "🔬 Apply PV Degradation Analysis",
                value=False,
                help="Enable 25-year PV degradation modeling (~9.8% total degradation)",
                key="apply_pv_degradation"
            )
            if apply_pv_degradation:
                st.info("✓ PV Degradation: ~9.8% over 25 years")
            
            # PROFILE UPLOAD (IN COMPONENT TAB)
            st.markdown("---")
            st.subheader("📊 PV Profile Upload")
            pv_file = st.file_uploader(
                "Upload PV Generation Profile (1 kW baseline)",
                type=['csv', 'xlsx'],
                key="pv_file",
                help="Hourly PV generation profile (8760 hours)"
            )
            if pv_file:
                st.success(f"✓ Uploaded: {pv_file.name}")
    
    # ==============================================================================
    # WIND - WITH PROFILE UPLOAD
    # ==============================================================================
    with st.expander("💨 WIND"):
        if not enable_wind:
            st.warning("⚠️ Wind is DISABLED")
            wind_min = 0.0
            wind_max = 0.0
            wind_step = 1.0
            wind_capex = 1200
            wind_opex = 15
            wind_lifetime = 20
            wind_file = None
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
            
            # PROFILE UPLOAD (IN COMPONENT TAB)
            st.markdown("---")
            st.subheader("📊 Wind Profile Upload")
            wind_file = st.file_uploader(
                "Upload Wind Generation Profile (1 kW baseline)",
                type=['csv', 'xlsx'],
                key="wind_file",
                help="Hourly wind generation profile (8760 hours)"
            )
            if wind_file:
                st.success(f"✓ Uploaded: {wind_file.name}")
    
    # ==============================================================================
    # HYDRO - WITH PROFILE UPLOAD
    # ==============================================================================
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
            hydro_file = None
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
            
            # PROFILE UPLOAD (IN COMPONENT TAB) - OPTIONAL
            st.markdown("---")
            st.subheader("📊 Hydro Profile Upload (Optional)")
            hydro_file = st.file_uploader(
                "Upload Hydro Availability Profile (optional)",
                type=['csv', 'xlsx'],
                key="hydro_file",
                help="Hourly hydro availability profile (8760 hours). If not provided, uses uniform availability."
            )
            if hydro_file:
                st.success(f"✓ Uploaded: {hydro_file.name}")
    
    # ==============================================================================
    # BESS - WITH DEGRADATION CHECKBOX
    # ==============================================================================
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
            apply_bess_degradation = False
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
                bess_min_soc = st.number_input("Min SOC (%)", value=20.0, min_value=0.0, max_value=100.0, step=0.1, key="bess_min_soc")
                bess_charge_eff = st.number_input("Charging Eff (%)", value=92.94, min_value=50.0, max_value=100.0, step=0.01, key="bess_charge_eff")
            with col2:
                bess_lifetime = st.number_input("Lifetime (years)", value=15, step=1, key="bess_life")
                bess_max_soc = st.number_input("Max SOC (%)", value=100.0, min_value=0.0, max_value=100.0, step=0.1, key="bess_max_soc")
                bess_discharge_eff = st.number_input("Discharging Eff (%)", value=91.78, min_value=50.0, max_value=100.0, step=0.01, key="bess_discharge_eff")
            
            st.subheader("Financial Parameters")
            col1, col2 = st.columns(2)
            with col1:
                bess_power_capex = st.number_input("Power CapEx ($/kW)", value=300, step=10, key="bess_power_capex")
                bess_energy_capex = st.number_input("Energy CapEx ($/kWh)", value=200, step=10, key="bess_energy_capex")
            with col2:
                bess_opex = st.number_input("OpEx ($/kW/yr)", value=2, step=1, key="bess_opex")
            
            # DEGRADATION CHECKBOX (IN COMPONENT TAB)
            st.markdown("---")
            apply_bess_degradation = st.checkbox(
                "🔬 Apply BESS Degradation Analysis",
                value=False,
                help="Enable 25-year BESS degradation modeling with replacement at year 21",
                key="apply_bess_degradation"
            )
            if apply_bess_degradation:
                st.info("✓ BESS Degradation: ~31% loss by year 20, replacement at year 21")
    
    # ==============================================================================
    # CCGT (Placeholder)
    # ==============================================================================
    with st.expander("🔥 CCGT (Coming Soon)", expanded=False):
        st.info("⚠️ Combined Cycle Gas Turbine module will be available in a future release")
        ccgt_min = 0.0
        ccgt_max = 0.0
        ccgt_step = 1.0
    
    # ==============================================================================
    # HYDROGEN (Placeholder)
    # ==============================================================================
    with st.expander("⚗️ HYDROGEN (Coming Soon)", expanded=False):
        st.info("⚠️ Hydrogen production and storage module will be available in a future release")
        hydrogen_min = 0.0
        hydrogen_max = 0.0
        hydrogen_step = 1.0
    
    # ==============================================================================
    # LOAD PROFILE SETUP
    # ==============================================================================
    st.markdown("---")
    st.header("📁 Load Profile Setup")
    
    # Load Type Dropdown
    load_type = st.selectbox(
        "Load Type",
        options=["Select Load Type", "Residential", "Commercial", "Industrial", "Community"],
        index=0,
        help="Select the type of load profile you want to upload"
    )
    
    # Show file uploader only after load type is selected
    if load_type != "Select Load Type":
        load_file = st.file_uploader(
            f"📊 Upload {load_type} Load Profile",
            type=['csv', 'xlsx'],
            key="load_file",
            help=f"Upload hourly load profile for {load_type} application (8760 hours)"
        )
        if load_file:
            st.success(f"✓ Uploaded: {load_file.name}")
    else:
        load_file = None
        st.caption("⚠️ Please select a load type first")
    
    # ==============================================================================
    # FINANCIAL PARAMETERS
    # ==============================================================================
    st.markdown("---")
    with st.expander("💰 FINANCIAL PARAMETERS"):
        discount_rate = st.number_input("Discount Rate (%)", value=8.0, min_value=0.0, max_value=20.0, step=0.5)
        inflation_rate = st.number_input("Inflation Rate (%)", value=2.0, min_value=0.0, max_value=10.0, step=0.5)
        project_lifetime = st.number_input("Project Lifetime (years)", value=25, min_value=1, max_value=50, step=1)
    
    # ==============================================================================
    # EMISSIONS SETTINGS
    # ==============================================================================
    with st.expander("🌍 EMISSIONS SETTINGS"):
        calculate_emissions = st.checkbox(
            "Calculate Emissions",
            value=False,
            help="Enable to calculate and display emissions (CO2, NOx, etc.) in results"
        )
        if calculate_emissions:
            st.info("📊 Emissions table will be displayed in Results tab")
    
    # ==============================================================================
    # OPTIMIZATION SETTINGS
    # ==============================================================================
    with st.expander("🎯 OPTIMIZATION SETTINGS"):
        target_unmet_percent = st.number_input("Target Unmet Load (%)", value=0.1, min_value=0.0, max_value=50.0, step=0.1, key="target_unmet")

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
    - **Future: CCGT and Hydrogen integration**
    
    ### 🔬 Advanced Degradation Analysis
    - **PV Degradation**: ~9.8% capacity loss over 25 years
    - **BESS Degradation**: ~31% capacity loss by year 20 with automatic replacement at year 21
    - **Full 25-Year Hourly Simulation**: Generates detailed hourly dispatch for selected years
    - **Enhanced Excel Export**: Multi-sheet output with yearly summary and hourly data
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
    
    # Show degradation status
    degradation_status = []
    if apply_pv_degradation:
        degradation_status.append("☀️ PV Degradation")
    if apply_bess_degradation:
        degradation_status.append("🔋 BESS Degradation")
    
    if len(degradation_status) > 0:
        st.info(f"**Degradation Analysis:** {' + '.join(degradation_status)}")
    
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
    
    if load_type != "Select Load Type":
        st.success(f"**Load Type Selected:** {load_type}")

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
    
    if load_type == "Select Load Type":
        validation_passed = False
        validation_messages.append("❌ Please select a load type")
    
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
                
                # Check if degradation analysis should be used
                use_degradation = apply_pv_degradation or apply_bess_degradation
                
                if use_degradation and not DEGRADATION_AVAILABLE:
                    st.error("❌ Degradation analysis module not available")
                    st.stop()
                
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
                    
                    # Component sheets
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
                
                # Choose which optimization module to use
                if use_degradation:
                    st.info(f"🔬 Using degradation analysis (PV: {apply_pv_degradation}, BESS: {apply_bess_degradation})")
                    
                    # Use degradation module for optimization
                    deg_module.INPUT_FILE = temp_file
                    result = deg_module.read_inputs()
                    
                    if len(result) == 9:
                        config, grid_config, solar, wind, hydro, bess, load_profile, pvsyst_profile, wind_profile = result
                    else:
                        config, grid_config, solar, wind, hydro, bess, load_profile, pvsyst_profile, wind_profile = result[:9]
                    
                    # Run grid search (WITHOUT degradation - finds optimal config)
                    results_df = deg_module.grid_search_optimize_hydro(
                        config, grid_config, solar, wind, hydro, bess,
                        load_profile, pvsyst_profile, wind_profile, None
                    )
                    
                    progress_bar.progress(70)
                    optimal = deg_module.find_optimal_solution(results_df)
                    
                    if optimal is not None:
                        status_text.text("🔬 Running 25-year degradation analysis...")
                        progress_bar.progress(75)
                        
                        # Prepare profiles for degradation analysis
                        profiles = {
                            'load': load_profile,
                            'pv': pvsyst_profile,
                            'wind': wind_profile,
                            'hydro': np.ones(len(load_profile))  # Default uniform hydro availability
                        }
                        
                        # Prepare config params
                        config_for_deg = {
                            'discount_rate': discount_rate,
                            'inflation_rate': inflation_rate,
                            'project_lifetime': project_lifetime,
                            'target_unmet_percent': target_unmet_percent,
                            'bess_charge_eff': bess_charge_eff,
                            'bess_discharge_eff': bess_discharge_eff,
                            'bess_max_soc': bess_max_soc,
                            'bess_min_soc': bess_min_soc
                        }
                        
                        # Run COMPLETE degradation analysis with hourly simulation
                        # Set export_all_years=True to export ALL 25 years for verification
                        degradation_results = deg_module.run_degradation_analysis_complete(
                            optimal.to_dict(),
                            config_for_deg,
                            profiles,
                            apply_pv=apply_pv_degradation,
                            apply_bess=apply_bess_degradation,
                            years_to_export=[1, 2, 5, 10, 15, 20, 25],  # Selected years
                            export_all_years=False  # Change to True to export all 25 years
                        )
                        
                        progress_bar.progress(90)
                        
                        # Use Year 1 hourly dispatch from degradation analysis
                        if 'year_1' in degradation_results['hourly_dispatch']:
                            # Get Year 1 with Anaconda column names (for display/Excel)
                            year_1_anaconda = degradation_results['hourly_dispatch']['year_1']
                            
                            # Create version with expected column names for electrical metrics
                            optimal_dispatch = year_1_anaconda.copy()
                            if 'PV_to_Load_kW' in optimal_dispatch.columns:
                                optimal_dispatch['PV_Output_kW'] = optimal_dispatch['PV_to_Load_kW']
                            if 'Wind_Output_kW' not in optimal_dispatch.columns:
                                optimal_dispatch['Wind_Output_kW'] = 0
                            if 'Hydro_Output_kW' not in optimal_dispatch.columns:
                                optimal_dispatch['Hydro_Output_kW'] = 0
                            if 'BESS_Charge_kW' not in optimal_dispatch.columns:
                                optimal_dispatch['BESS_Charge_kW'] = optimal_dispatch.get('BESS_Charge_woeff_kW', 0)
                            if 'BESS_Discharge_kW' not in optimal_dispatch.columns:
                                optimal_dispatch['BESS_Discharge_kW'] = optimal_dispatch.get('BESS_Discharge_wieff_kW', 0)
                            if 'Excess_kW' not in optimal_dispatch.columns:
                                optimal_dispatch['Excess_kW'] = optimal_dispatch.get('Curtailment_kW', 0)
                        else:
                            optimal_dispatch = None
                    else:
                        degradation_results = None
                        optimal_dispatch = None
                    
                else:
                    # Standard optimization (no degradation)
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
                    degradation_results = None
                    
                    if optimal is not None:
                        optimal_dispatch = opt_module.calculate_dispatch_with_hydro(
                            load_profile, pvsyst_profile, wind_profile,
                            optimal['PV_kW'], optimal['Wind_kW'], optimal['Hydro_kW'],
                            optimal['BESS_Power_kW'], optimal['BESS_Capacity_kWh'],
                            solar, wind, hydro, bess,
                            int(optimal['Hydro_Window_Start']), int(optimal['Hydro_Window_End'])
                        )
                    else:
                        optimal_dispatch = None
                
                progress_bar.progress(95)
                
                if optimal is not None:
                    # Calculate electrical metrics
                    module = deg_module if use_degradation else opt_module
                    
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
                    
                    electrical_metrics = module.calculate_electrical_metrics(
                        optimal_dispatch, component_capacities, component_configs,
                        npc_breakdown, project_lifetime
                    )
                    
                    progress_bar.progress(100)
                    
                    # Clean up
                    if os.path.exists(temp_file):
                        os.remove(temp_file)
                    
                    # Store results with degradation data
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
                        },
                        'load_type': load_type,
                        'calculate_emissions': calculate_emissions,
                        'degradation_settings': {
                            'pv': apply_pv_degradation,
                            'bess': apply_bess_degradation
                        },
                        'degradation_results': degradation_results  # NEW: Store degradation results
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
        
        # Show degradation status if applied
        if 'degradation_settings' in results:
            deg_settings = results['degradation_settings']
            if deg_settings['pv'] or deg_settings['bess']:
                deg_status = []
                if deg_settings['pv']:
                    deg_status.append("☀️ PV")
                if deg_settings['bess']:
                    deg_status.append("🔋 BESS")
                st.info(f"🔬 **Degradation Analysis Applied:** {' + '.join(deg_status)}")
        
        # Key Metrics
        col1, col2, col3 = st.columns(3)
        with col1:
            st.metric("Total NPC", f"${results['npc']/1e6:.2f}M")
        with col2:
            st.metric("System LCOE", f"${results['lcoe']:.2f}/MWh")
        with col3:
            st.metric("Unmet Load", f"{results['unmet_pct']:.3f}%")
        
        # Display Load Type
        if 'load_type' in results:
            st.info(f"**Load Type:** {results['load_type']}")
        
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
        
        # Degradation Analysis Summary (if available)
        if 'degradation_results' in results and results['degradation_results']:
            degradation_results = results['degradation_results']
            
            st.subheader("🔬 Degradation Analysis Summary")
            
            col1, col2, col3, col4 = st.columns(4)
            with col1:
                st.metric("Year 1 NPC", f"${degradation_results['npc_year1']/1e6:.2f}M")
            with col2:
                st.metric("25-Year NPC", f"${degradation_results['npc_25year']/1e6:.2f}M")
            with col3:
                st.metric("Replacement Cost", f"${degradation_results['replacement_cost_pv']/1e6:.2f}M")
            with col4:
                npc_increase = ((degradation_results['npc_25year'] / degradation_results['npc_year1']) - 1) * 100
                st.metric("NPC Increase", f"{npc_increase:.1f}%")
            
            # Show yearly summary table
            st.markdown("**25-Year Performance Summary**")
            st.dataframe(degradation_results['yearly_summary'], use_container_width=True, hide_index=True)
            
            # List exported years
            years_list = ", ".join([f"Year {y}" for y in degradation_results.get('years_exported', [])])
            st.info(f"📊 **Hourly Dispatch Exported for:** {years_list}")
            
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
        
        # Emissions Table (if enabled)
        if results.get('calculate_emissions', False):
            st.subheader("🌍 Emissions Summary")
            emissions_table = create_emissions_table(optimal_row, results)
            st.dataframe(emissions_table, use_container_width=True, hide_index=True)
            st.caption("Note: Emissions calculations based on system configuration and dispatch profile")
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
        
        # Single Day Dispatch Profile
        st.subheader("📈 Single Day Dispatch Profile")
        daily_profile_fig = create_single_day_dispatch_profile(results)
        st.plotly_chart(daily_profile_fig, use_container_width=True)
        st.caption("Shows dispatch for one representative day (median PV production)")
        
        st.markdown("---")
        
        # Energy Mix Pie Chart
        st.subheader("📊 Energy Production Mix")
        
        col1, col2 = st.columns(2)
        with col1:
            pie_fig, energy_table = create_energy_mix_pie_chart(optimal_row)
            st.plotly_chart(pie_fig, use_container_width=True)
        
        with col2:
            st.markdown("**Energy Production Breakdown**")
            st.dataframe(energy_table, use_container_width=True, hide_index=True)
        
        st.markdown("---")
        
        # Download
        st.subheader("📥 Download Results")
        
        # Get degradation results if available
        degradation_results = results.get('degradation_results', None)
        
        # Use the export function with degradation support
        excel_output = export_results_industry_format(
            results, results['results_df'], results['optimal_row'], 
            results['config_params'], degradation_results
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
st.markdown('<div style="text-align:center;color:#666"><p>RE Optimization Tool v4.0 | Complete Degradation Analysis Integration</p></div>', unsafe_allow_html=True)
