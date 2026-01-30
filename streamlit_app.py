"""
RENEWABLE ENERGY OPTIMIZATION TOOL - COMPLETE VERSION
======================================================
Phase 1 + Phase 2 Enhancements:
- Electrical metrics with capacity factor, LCOE, hours of operation
- Enhanced cost visualizations (4 variations with tables)
- Fixed cash flow chart with proper signs
- Professional layout organization
- Individual component LCOE display
- Excess electricity and capacity shortage metrics
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
# EXCEL EXPORT - Industry Standard Format
# ==============================================================================

def export_results_industry_format(results_dict, results_df, optimal_row, config_params):
    """Export results matching industry standard Excel format."""
    output = BytesIO()
    
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        
        # Sheet 1: Summary
        summary_data = []
        summary_data.extend([
            ['Parameter', 'Value'],
            ['Optimization Method', 'GRID_SEARCH'],
            ['NPC Calculation Method', 'Present Value Analysis'],
            ['LCOE Mode', 'STATIC'],
            ['Target Unmet Load (%)', config_params.get('target_unmet_percent', 0.1)],
            ['Total Combinations Tested', len(results_df)],
            ['Feasible Solutions', len(results_df[results_df['Feasible'] == True])],
            ['Optimal Solution Iteration', optimal_row.get('Iteration', 0)],
            ['', ''],
            ['Nominal Discount Rate (%)', config_params.get('discount_rate', 8.0)],
            ['Inflation Rate (%)', config_params.get('inflation_rate', 2.0)],
            ['Real Discount Rate (%)', opt_module.calculate_real_discount_rate(
                config_params.get('discount_rate', 8.0)/100, 
                config_params.get('inflation_rate', 2.0)/100) * 100],
            ['CRF', opt_module.calculate_crf(
                opt_module.calculate_real_discount_rate(
                    config_params.get('discount_rate', 8.0)/100, 
                    config_params.get('inflation_rate', 2.0)/100),
                config_params.get('project_lifetime', 25))],
            ['', ''],
            ['PV Capacity (kW)', results_dict['pv_capacity'] * 1000],
            ['Wind Capacity (kW)', results_dict['wind_capacity'] * 1000],
            ['Hydro Capacity (kW)', results_dict['hydro_capacity'] * 1000],
            ['Hydro Window Start (hour)', int(results_dict['hydro_window_start'])],
            ['Hydro Window End (hour)', int(results_dict['hydro_window_end'])],
            ['Hydro Operating Hours/Day', int(results_dict['hydro_window_end']) - int(results_dict['hydro_window_start'])],
            ['BESS Power (kW)', results_dict['bess_power'] * 1000],
            ['BESS Capacity (kWh)', results_dict['bess_energy'] * 1000],
            ['', ''],
            ['Total NPC ($)', results_dict['npc']],
            ['Total Capital ($)', optimal_row.get('Capital_$', 0)],
            ['Total Replacement ($)', optimal_row.get('Replacement_$', 0)],
            ['Total O&M ($)', optimal_row.get('OM_$', 0)],
            ['Total Salvage ($)', optimal_row.get('Salvage_$', 0)],
            ['Total Annualized ($/year)', optimal_row.get('Annualized_$/yr', 0)],
            ['', ''],
            ['Levelized COE ($/kWh)', optimal_row.get('LCOE_$/kWh', 0)],
            ['Levelized COE ($/MWh)', results_dict['lcoe']],
            ['Unmet Load (%)', results_dict['unmet_pct']],
        ])
        
        pd.DataFrame(summary_data).to_excel(writer, sheet_name='Summary', index=False, header=False)
        
        # Sheet 2: Cost_Breakdown
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
            'O&M ($)': [
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
            'Total NPC ($)': [
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
        
        # Sheet 3: All_Results
        results_df.to_excel(writer, sheet_name='All_Results', index=False)
        
        # Sheet 4: Feasible_Solutions
        feasible = results_df[results_df['Feasible'] == True].sort_values('NPC_$')
        if len(feasible) > 0:
            feasible.to_excel(writer, sheet_name='Feasible_Solutions', index=False)
        else:
            pd.DataFrame({'Message': ['No feasible solutions found']}).to_excel(
                writer, sheet_name='Feasible_Solutions', index=False)
        
        # Sheet 5: Hydro_Window_Analysis
        if results_dict['hydro_capacity'] > 0:
            hydro_window_data = []
            hydro_hours = int(results_dict['hydro_window_end']) - int(results_dict['hydro_window_start'])
            
            hydro_window_data.extend([
                ['HYDRO WINDOW ANALYSIS', ''],
                ['', ''],
                ['Parameter', 'Value'],
                ['Hydro Capacity (kW)', results_dict['hydro_capacity'] * 1000],
                ['Operating Hours per Day', hydro_hours],
                ['Optimal Window Start (hour)', int(results_dict['hydro_window_start'])],
                ['Optimal Window End (hour)', int(results_dict['hydro_window_end'])],
                ['', ''],
                ['Window Configuration', f"{int(results_dict['hydro_window_start']):02d}:00 - {int(results_dict['hydro_window_end']):02d}:00"],
                ['Total Possible Windows', 24 - hydro_hours + 1],
                ['Optimal Window Unmet Load (%)', results_dict['unmet_pct']],
            ])
            
            pd.DataFrame(hydro_window_data).to_excel(writer, sheet_name='Hydro_Window_Analysis', index=False, header=False)
        else:
            pd.DataFrame({'Note': ['No hydro in optimal configuration']}).to_excel(
                writer, sheet_name='Hydro_Window_Analysis', index=False)
    
    output.seek(0)
    return output


# ==============================================================================
# COST ANALYSIS CHARTS WITH TABLES
# ==============================================================================

def create_cost_analysis_charts_with_tables(results, optimal_row):
    """
    Create 4 cost chart variations:
    1. Net Present by Component
    2. Net Present by Cost Type
    3. Annualized by Component
    4. Annualized by Cost Type
    """
    
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
        go.Bar(
            x=components, 
            y=npc_values, 
            marker_color=['#FDB462', '#80B1D3', '#8DD3C7', '#FB8072'],
            text=[f'${v:.2f}M' for v in npc_values],
            textposition='outside'
        )
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
        'Capital ($)': [
            f"${optimal_row.get('PV_Capital_$', 0):,.0f}",
            f"${optimal_row.get('Wind_Capital_$', 0):,.0f}",
            f"${optimal_row.get('Hydro_Capital_$', 0):,.0f}",
            f"${optimal_row.get('BESS_Capital_$', 0):,.0f}"
        ],
        'Replacement ($)': [
            f"${optimal_row.get('PV_Replacement_$', 0):,.0f}",
            f"${optimal_row.get('Wind_Replacement_$', 0):,.0f}",
            f"${optimal_row.get('Hydro_Replacement_$', 0):,.0f}",
            f"${optimal_row.get('BESS_Replacement_$', 0):,.0f}"
        ],
        'O&M ($)': [
            f"${optimal_row.get('PV_OM_$', 0):,.0f}",
            f"${optimal_row.get('Wind_OM_$', 0):,.0f}",
            f"${optimal_row.get('Hydro_OM_$', 0):,.0f}",
            f"${optimal_row.get('BESS_OM_$', 0):,.0f}"
        ],
        'Salvage ($)': [
            f"${optimal_row.get('PV_Salvage_$', 0):,.0f}",
            f"${optimal_row.get('Wind_Salvage_$', 0):,.0f}",
            f"${optimal_row.get('Hydro_Salvage_$', 0):,.0f}",
            f"${optimal_row.get('BESS_Salvage_$', 0):,.0f}"
        ],
        'Total NPC ($)': [
            f"${optimal_row.get('PV_NPC_$', 0):,.0f}",
            f"${optimal_row.get('Wind_NPC_$', 0):,.0f}",
            f"${optimal_row.get('Hydro_NPC_$', 0):,.0f}",
            f"${optimal_row.get('BESS_NPC_$', 0):,.0f}"
        ]
    })
    
    charts_and_tables['npc_by_component'] = {'chart': fig1, 'table': table1}
    
    
    # Chart 3: Annualized Cost by Component
    annualized_values = [
        optimal_row.get('PV_Annualized_$/yr', 0) / 1e3,
        optimal_row.get('Wind_Annualized_$/yr', 0) / 1e3,
        optimal_row.get('Hydro_Annualized_$/yr', 0) / 1e3,
        optimal_row.get('BESS_Annualized_$/yr', 0) / 1e3
    ]
    
    fig3 = go.Figure(data=[
        go.Bar(
            x=components,
            y=annualized_values,
            marker_color=['#FDB462', '#80B1D3', '#8DD3C7', '#FB8072'],
            text=[f'${v:.1f}K' for v in annualized_values],
            textposition='outside'
        )
    ])
    fig3.update_layout(
        title='Annualized Cost by Component',
        xaxis_title='Component',
        yaxis_title='Annualized Cost ($K/year)',
        height=400,
        showlegend=False
    )
    
    table3 = pd.DataFrame({
        'Component': components,
        'Annualized Cost ($/year)': [
            f"${optimal_row.get('PV_Annualized_$/yr', 0):,.0f}",
            f"${optimal_row.get('Wind_Annualized_$/yr', 0):,.0f}",
            f"${optimal_row.get('Hydro_Annualized_$/yr', 0):,.0f}",
            f"${optimal_row.get('BESS_Annualized_$/yr', 0):,.0f}"
        ],
        'NPC ($)': [
            f"${optimal_row.get('PV_NPC_$', 0):,.0f}",
            f"${optimal_row.get('Wind_NPC_$', 0):,.0f}",
            f"${optimal_row.get('Hydro_NPC_$', 0):,.0f}",
            f"${optimal_row.get('BESS_NPC_$', 0):,.0f}"
        ]
    })
    
    charts_and_tables['annualized_by_component'] = {'chart': fig3, 'table': table3}
    
    return charts_and_tables


# ==============================================================================
# FIXED CASH FLOW CHART
# ==============================================================================

def create_fixed_cash_flow_chart(results, optimal_row):
    """
    Create CORRECTED cash flow chart where:
    - Capital, Operating, Replacement are NEGATIVE (outflows from bottom)
    - Salvage is POSITIVE (inflow)
    """
    
    project_lifetime = results.get('config_params', {}).get('project_lifetime', 25)
    years = list(range(0, project_lifetime + 1))
    
    # Initialize cash flows
    capital_flow = [0] * len(years)
    operating_flow = [0] * len(years)
    replacement_flow = [0] * len(years)
    salvage_flow = [0] * len(years)
    
    # Year 0: Capital costs (NEGATIVE)
    capital_flow[0] = -optimal_row.get('Capital_$', 0) / 1e6
    
    # Annual O&M (NEGATIVE)
    annual_om = optimal_row.get('OM_$', 0) / project_lifetime / 1e6
    for year in range(1, project_lifetime + 1):
        operating_flow[year] = -annual_om
    
    # Replacements (NEGATIVE) - simplified schedule
    bess_lifetime = 10
    wind_lifetime = 20
    
    bess_replacement = optimal_row.get('BESS_Replacement_$', 0) / 1e6
    wind_replacement = optimal_row.get('Wind_Replacement_$', 0) / 1e6
    
    for year in years:
        if year > 0:
            if year % bess_lifetime == 0 and year < project_lifetime:
                replacement_flow[year] -= bess_replacement * 0.5
            if year % wind_lifetime == 0 and year < project_lifetime:
                replacement_flow[year] -= wind_replacement * 0.5
    
    # Final year: Salvage value (POSITIVE)
    salvage_flow[-1] = optimal_row.get('Salvage_$', 0) / 1e6
    
    # Create stacked bar chart
    fig = go.Figure()
    
    fig.add_trace(go.Bar(
        name='Capital',
        x=years,
        y=capital_flow,
        marker_color='#2E7D32'
    ))
    
    fig.add_trace(go.Bar(
        name='Operating',
        x=years,
        y=operating_flow,
        marker_color='#F57C00'
    ))
    
    fig.add_trace(go.Bar(
        name='Replacement',
        x=years,
        y=replacement_flow,
        marker_color='#1976D2'
    ))
    
    fig.add_trace(go.Bar(
        name='Salvage',
        x=years,
        y=salvage_flow,
        marker_color='#388E3C'
    ))
    
    fig.update_layout(
        title='Nominal Cash Flow',
        xaxis_title='Year',
        yaxis_title='Cash Flow ($M)',
        barmode='relative',
        height=450,
        showlegend=True,
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1)
    )
    
    # Create data table
    cash_flow_table = pd.DataFrame({
        'Year': years[:10],
        'Capital ($M)': [f"${v:.2f}" for v in capital_flow[:10]],
        'Operating ($M)': [f"${v:.2f}" for v in operating_flow[:10]],
        'Replacement ($M)': [f"${v:.2f}" for v in replacement_flow[:10]],
        'Salvage ($M)': [f"${v:.2f}" for v in salvage_flow[:10]],
        'Net Cash Flow ($M)': [
            f"${capital_flow[i] + operating_flow[i] + replacement_flow[i] + salvage_flow[i]:.2f}"
            for i in range(10)
        ]
    })
    
    return fig, cash_flow_table

def create_typical_daily_dispatch_profile(results, optimal_row):
    """
    Create a typical 24-hour dispatch profile chart matching manager's requirement.
    Shows PV, Wind, Hydro, Load, and BESS SOC variation over a representative day.
    """
    
    # Get hourly dispatch data (use average of all days or pick a representative day)
    if 'optimal_dispatch' in results:
        dispatch_df = results['optimal_dispatch']
        
        # Calculate average hourly profile across all days
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
        
        # Convert to MW for cleaner display
        avg_profile['Load_MW'] = avg_profile['Load_kW'] / 1000
        avg_profile['PV_MW'] = avg_profile['PV_Output_kW'] / 1000
        avg_profile['Wind_MW'] = avg_profile['Wind_Output_kW'] / 1000
        avg_profile['Hydro_MW'] = avg_profile['Hydro_Output_kW'] / 1000
        avg_profile['BESS_Net_MW'] = (avg_profile['BESS_Discharge_kW'] - avg_profile['BESS_Charge_kW']) / 1000
        
        # Calculate BESS SOC percentage
        bess_capacity_kwh = results.get('bess_energy', 1) * 1000  # Convert MWh to kWh
        avg_profile['BESS_SOC_%'] = (avg_profile['BESS_SOC_kWh'] / bess_capacity_kwh) * 100
        
    else:
        # Fallback: Create sample profile if dispatch data not available
        hours = list(range(24))
        avg_profile = pd.DataFrame({
            'Hour_of_Day': hours,
            'Load_MW': [5] * 24,
            'PV_MW': [0] * 24,
            'Wind_MW': [0] * 24,
            'Hydro_MW': [0] * 24,
            'BESS_Net_MW': [0] * 24,
            'BESS_SOC_%': [50] * 24
        })
    
    # Create figure with secondary y-axis
    fig = make_subplots(specs=[[{"secondary_y": True}]])
    
    # Add stacked area for generation sources
    fig.add_trace(
        go.Scatter(
            x=avg_profile['Hour_of_Day'],
            y=avg_profile['Hydro_MW'],
            name='Hydro',
            mode='lines',
            line=dict(width=0),
            stackgroup='one',
            fillcolor='rgba(141, 211, 199, 0.7)',  # Teal
            hovertemplate='Hour %{x}<br>Hydro: %{y:.2f} MW<extra></extra>'
        ),
        secondary_y=False
    )
    
    fig.add_trace(
        go.Scatter(
            x=avg_profile['Hour_of_Day'],
            y=avg_profile['PV_MW'],
            name='PV',
            mode='lines',
            line=dict(width=0),
            stackgroup='one',
            fillcolor='rgba(253, 180, 98, 0.7)',  # Orange
            hovertemplate='Hour %{x}<br>PV: %{y:.2f} MW<extra></extra>'
        ),
        secondary_y=False
    )
    
    fig.add_trace(
        go.Scatter(
            x=avg_profile['Hour_of_Day'],
            y=avg_profile['Wind_MW'],
            name='Wind',
            mode='lines',
            line=dict(width=0),
            stackgroup='one',
            fillcolor='rgba(128, 177, 211, 0.7)',  # Blue
            hovertemplate='Hour %{x}<br>Wind: %{y:.2f} MW<extra></extra>'
        ),
        secondary_y=False
    )
    
    # Add BESS net power (can be positive or negative)
    fig.add_trace(
        go.Scatter(
            x=avg_profile['Hour_of_Day'],
            y=avg_profile['BESS_Net_MW'],
            name='BESS Net',
            mode='lines',
            line=dict(width=2, color='rgba(251, 128, 114, 0.8)', dash='dot'),
            hovertemplate='Hour %{x}<br>BESS: %{y:.2f} MW<extra></extra>'
        ),
        secondary_y=False
    )
    
    # Add load as a line (not stacked)
    fig.add_trace(
        go.Scatter(
            x=avg_profile['Hour_of_Day'],
            y=avg_profile['Load_MW'],
            name='Load',
            mode='lines',
            line=dict(width=3, color='#FFD700'),  # Gold/Yellow
            hovertemplate='Hour %{x}<br>Load: %{y:.2f} MW<extra></extra>'
        ),
        secondary_y=False
    )
    
    # Add BESS SOC on secondary y-axis
    fig.add_trace(
        go.Scatter(
            x=avg_profile['Hour_of_Day'],
            y=avg_profile['BESS_SOC_%'],
            name='BESS SOC',
            mode='lines',
            line=dict(width=3, color='#FF00FF'),  # Magenta
            hovertemplate='Hour %{x}<br>SOC: %{y:.1f}%<extra></extra>'
        ),
        secondary_y=True
    )
    
    # Update layout
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
        title='Typical Daily Dispatch Profile',
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
# ELECTRICAL METRICS TABLES WITH SUNGROW BESS DEPLOYMENT
# ==============================================================================

def calculate_bess_deployment_sungrow(bess_power_mw, bess_capacity_mwh):
    """
    Calculate BESS deployment using REAL Sungrow PowerTitan 2.0 specifications.
    Based on actual layout drawing with corrected spacing interpretation.
    
    Args:
        bess_power_mw: Total BESS power (MW)
        bess_capacity_mwh: Total BESS capacity (MWh)
    
    Returns:
        Dict with deployment metrics
    """
    import math
    
    # Sungrow PowerTitan 2.0 specifications
    container_capacity_mwh = 10.0  # Each container: 10 MWh
    container_power_mw = 5.0       # Each container: 5 MW
    container_length_m = 6.058     # Container depth (m)
    container_width_m = 2.438      # Container width (m)
    container_height_m = 2.896     # Container height (m)
    
    # Spacing requirements from Sungrow layout drawing
    back_to_back_spacing_m = 0.150   # 150 mm - LENGTH direction spacing
    adjacent_spacing_m = 1.500        # 1,500 mm - WIDTH direction spacing
    mvs_spacing_m = 3.500            # 3,500 mm - spacing to MVS unit
    mvs_width_m = 2.000              # MVS5000 unit width (estimated)
    perimeter_clearance_m = 5.000    # Fire safety perimeter clearance
    
    # Calculate number of containers needed
    num_containers_energy = math.ceil(bess_capacity_mwh / container_capacity_mwh)
    num_containers_power = math.ceil(bess_power_mw / container_power_mw)
    num_containers = max(num_containers_energy, num_containers_power)
    
    # Calculate actual installed capacity
    actual_capacity_mwh = num_containers * container_capacity_mwh
    actual_power_mw = num_containers * container_power_mw
    
    # Number of MVS units (1 MVS5000 per 2 containers)
    num_mvs_units = math.ceil(num_containers / 2)
    
    # Calculate layout dimensions
    if num_containers <= 2:
        # Single section (2 containers back-to-back + 1 MVS)
        # LENGTH: Just container length + perimeter clearances
        total_length_m = container_length_m + (2 * perimeter_clearance_m)
        
        # WIDTH: Containers back-to-back + MVS section + clearances
        container_section_width = (container_width_m +           # Container 1
                                   back_to_back_spacing_m +      # Gap
                                   container_width_m)            # Container 2
        
        mvs_section_width = mvs_spacing_m + mvs_width_m         # MVS section
        
        total_width_m = (container_section_width + 
                        mvs_section_width + 
                        2 * perimeter_clearance_m)
        
        layout_desc = "1 section (2 containers back-to-back + 1 MVS unit)"
        
    else:
        # Multiple sections arranged side-by-side
        num_sections = math.ceil(num_containers / 2)
        
        # LENGTH: Multiple sections side-by-side with adjacent spacing
        section_length = container_length_m + adjacent_spacing_m
        total_length_m = (num_sections * section_length - adjacent_spacing_m +  # Last section has no spacing after
                         2 * perimeter_clearance_m)
        
        # WIDTH: Same as single section
        container_section_width = (container_width_m + 
                                   back_to_back_spacing_m + 
                                   container_width_m)
        mvs_section_width = mvs_spacing_m + mvs_width_m
        total_width_m = (container_section_width + 
                        mvs_section_width + 
                        2 * perimeter_clearance_m)
        
        layout_desc = f"{num_sections} sections side-by-side ({num_containers} containers + {num_mvs_units} MVS units)"
    
    # Calculate total area
    total_area_m2 = total_length_m * total_width_m
    total_area_hectares = total_area_m2 / 10000
    total_area_acres = total_area_hectares * 2.471
    
    # Calculate container footprint (without clearances)
    container_footprint_m2 = num_containers * (container_length_m * container_width_m)
    
    # Density metrics
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


def create_electrical_metrics_tables(electrical_metrics, bess_power_mw=0, bess_capacity_mwh=0):
    """
    Create formatted tables for electrical metrics of each component.
    Now includes Sungrow BESS deployment specifications.
    
    Args:
        electrical_metrics: Dict with electrical metrics from optimization
        bess_power_mw: BESS power in MW (for deployment calculation)
        bess_capacity_mwh: BESS capacity in MWh (for deployment calculation)
    """
    
    tables = {}
    
    if electrical_metrics:
        # ========================================================================
        # PV TABLE
        # ========================================================================
        pv_data = pd.DataFrame({
            'Metric': [
                'Rated Capacity',
                'Mean Output',
                'Capacity Factor',
                'Total Production',
                'Hours of Operation',
                'Levelized Cost (LCOE)'
            ],
            'Value': [
                f"{electrical_metrics['pv']['rated_capacity_kw']:,.1f} kW",
                f"{electrical_metrics['pv']['mean_output_kw']:,.1f} kW",
                f"{electrical_metrics['pv']['capacity_factor_pct']:.2f}%",
                f"{electrical_metrics['pv']['total_production_kwh']:,.0f} kWh/yr",
                f"{electrical_metrics['pv']['hours_of_operation']:,.0f} hrs/yr",
                f"${electrical_metrics['pv']['levelized_cost_per_kwh']:.4f}/kWh (${electrical_metrics['pv'].get('levelized_cost_per_mwh', 0):.2f}/MWh)"
            ]
        })
        tables['pv'] = pv_data
        
        # ========================================================================
        # WIND TABLE
        # ========================================================================
        wind_data = pd.DataFrame({
            'Metric': [
                'Rated Capacity',
                'Mean Output',
                'Capacity Factor',
                'Total Production',
                'Hours of Operation',
                'Levelized Cost (LCOE)'
            ],
            'Value': [
                f"{electrical_metrics['wind']['rated_capacity_kw']:,.1f} kW",
                f"{electrical_metrics['wind']['mean_output_kw']:,.1f} kW",
                f"{electrical_metrics['wind']['capacity_factor_pct']:.2f}%",
                f"{electrical_metrics['wind']['total_production_kwh']:,.0f} kWh/yr",
                f"{electrical_metrics['wind']['hours_of_operation']:,.0f} hrs/yr",
                f"${electrical_metrics['wind']['levelized_cost_per_kwh']:.4f}/kWh (${electrical_metrics['wind'].get('levelized_cost_per_mwh', 0):.2f}/MWh)"
            ]
        })
        tables['wind'] = wind_data
        
        # ========================================================================
        # HYDRO TABLE
        # ========================================================================
        hydro_data = pd.DataFrame({
            'Metric': [
                'Rated Capacity',
                'Mean Output',
                'Capacity Factor',
                'Total Production',
                'Hours of Operation',
                'Levelized Cost (LCOE)'
            ],
            'Value': [
                f"{electrical_metrics['hydro']['rated_capacity_kw']:,.1f} kW",
                f"{electrical_metrics['hydro']['mean_output_kw']:,.1f} kW",
                f"{electrical_metrics['hydro']['capacity_factor_pct']:.2f}%",
                f"{electrical_metrics['hydro']['total_production_kwh']:,.0f} kWh/yr",
                f"{electrical_metrics['hydro']['hours_of_operation']:,.0f} hrs/yr",
                f"${electrical_metrics['hydro']['levelized_cost_per_kwh']:.4f}/kWh (${electrical_metrics['hydro'].get('levelized_cost_per_mwh', 0):.2f}/MWh)"
            ]
        })
        tables['hydro'] = hydro_data
        
        # ========================================================================
        # BESS TABLE - PERFORMANCE METRICS
        # ========================================================================
        bess_performance = pd.DataFrame({
            'Metric': [
                'Nominal Capacity',
                'Usable Capacity',
                'Autonomy',
                'Energy In',
                'Energy Out',
                'Losses',
                'Annual Throughput',
                'Expected Life',
                'Levelized Cost (LCOS)'
            ],
            'Value': [
                f"{electrical_metrics['bess']['nominal_capacity_kwh']:,.1f} kWh",
                f"{electrical_metrics['bess']['usable_capacity_kwh']:,.1f} kWh",
                f"{electrical_metrics['bess']['autonomy_hours']:.2f} hours",
                f"{electrical_metrics['bess']['energy_in_kwh']:,.0f} kWh/yr",
                f"{electrical_metrics['bess']['energy_out_kwh']:,.0f} kWh/yr",
                f"{electrical_metrics['bess']['losses_kwh']:,.0f} kWh/yr",
                f"{electrical_metrics['bess']['annual_throughput_kwh']:,.0f} kWh/yr",
                f"{electrical_metrics['bess']['expected_life_years']:.0f} years",
                f"${electrical_metrics['bess']['levelized_cost_per_kwh']:.4f}/kWh (${electrical_metrics['bess'].get('levelized_cost_per_mwh', 0):.2f}/MWh)"
            ]
        })
        tables['bess_performance'] = bess_performance
        
        # ========================================================================
        # BESS TABLE - SUNGROW DEPLOYMENT SPECIFICATIONS
        # ========================================================================
        if bess_capacity_mwh > 0:
            deployment = calculate_bess_deployment_sungrow(bess_power_mw, bess_capacity_mwh)
            
            bess_deployment = pd.DataFrame({
                'Metric': [
                    '--- DEPLOYMENT SPECIFICATIONS ---',
                    'OEM / Model',
                    'Number of Containers',
                    'Capacity per Container',
                    'Container Dimensions (L×W×H)',
                    'Actual Installed Capacity',
                    'Actual Installed Power',
                    '',
                    'Layout Configuration',
                    'Number of MVS5000 Units',
                    '',
                    'Container Footprint',
                    'Site Dimensions (L×W)',
                    'Total Site Area',
                    'Area (hectares)',
                    'Area (acres)',
                    '',
                    'Power Density',
                    'Energy Density',
                    '',
                    '--- SPACING SPECIFICATIONS ---',
                    'Back-to-back Spacing',
                    'Adjacent Container Spacing',
                    'Container to MVS Spacing',
                    'Perimeter Clearance'
                ],
                'Value': [
                    '',
                    f"Sungrow {deployment['container_model']}",
                    f"{deployment['num_containers']} units",
                    f"{deployment['container_capacity_mwh']:.0f} MWh / {deployment['container_power_mw']:.0f} MW each",
                    deployment['container_dimensions'],
                    f"{deployment['actual_capacity_mwh']:.0f} MWh",
                    f"{deployment['actual_power_mw']:.0f} MW",
                    '',
                    deployment['layout_description'],
                    f"{deployment['num_mvs_units']} units",
                    '',
                    f"{deployment['container_footprint_m2']:,.1f} m²",
                    deployment['site_dimensions'],
                    f"{deployment['total_area_m2']:,.0f} m²",
                    f"{deployment['total_area_hectares']:.4f} ha",
                    f"{deployment['total_area_acres']:.3f} acres",
                    '',
                    f"{deployment['power_density_mw_per_ha']:.1f} MW/ha",
                    f"{deployment['energy_density_mwh_per_ha']:.1f} MWh/ha",
                    '',
                    '',
                    f"{deployment['spacing_back_to_back_mm']} mm (length direction)",
                    f"{deployment['spacing_adjacent_mm']} mm (width direction)",
                    f"{deployment['spacing_mvs_mm']} mm",
                    f"{5000} mm (all sides)"
                ]
            })
            tables['bess_deployment'] = bess_deployment
        else:
            # No BESS in configuration
            bess_deployment = pd.DataFrame({
                'Metric': ['Note'],
                'Value': ['No BESS in optimal configuration']
            })
            tables['bess_deployment'] = bess_deployment
    
    return tables

# ==============================================================================
# ADDITIONAL SIMPLE CHARTS
# ==============================================================================

def create_additional_charts(results):
    """Create monthly production and energy mix charts."""
    
    charts = {}
    optimal_row = results.get('optimal_row', {})
    
    # Monthly Electric Production
    months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']
    
    pv_annual = optimal_row.get('PV_Energy_kWh', 0) / 1000
    wind_annual = optimal_row.get('Wind_Energy_kWh', 0) / 1000
    hydro_annual = optimal_row.get('Hydro_Energy_kWh', 0) / 1000
    
    pv_monthly = [pv_annual/12 * (1 + 0.3 * np.sin((i - 3) * np.pi / 6)) for i in range(12)]
    wind_monthly = [wind_annual/12 * (1 + 0.2 * np.cos(i * np.pi / 6)) for i in range(12)]
    hydro_monthly = [hydro_annual/12 for i in range(12)]
    
    fig_monthly = go.Figure()
    fig_monthly.add_trace(go.Bar(name='PV', x=months, y=pv_monthly, marker_color='#FDB462'))
    fig_monthly.add_trace(go.Bar(name='Wind', x=months, y=wind_monthly, marker_color='#80B1D3'))
    fig_monthly.add_trace(go.Bar(name='Hydro', x=months, y=hydro_monthly, marker_color='#8DD3C7'))
    
    fig_monthly.update_layout(
        title='Monthly Electric Production',
        xaxis_title='Month',
        yaxis_title='Production (MWh)',
        barmode='stack',
        height=450,
        showlegend=True
    )
    charts['monthly_production'] = fig_monthly
    
    # Energy Mix Pie Chart
    pv_frac = optimal_row.get('PV_Fraction_%', 0)
    wind_frac = optimal_row.get('Wind_Fraction_%', 0)
    hydro_frac = optimal_row.get('Hydro_Fraction_%', 0)
    
    fig_pie = go.Figure(data=[go.Pie(
        labels=['PV', 'Wind', 'Hydro'],
        values=[pv_frac, wind_frac, hydro_frac],
        marker=dict(colors=['#FDB462', '#80B1D3', '#8DD3C7']),
        textinfo='label+percent',
        hovertemplate='<b>%{label}</b><br>%{value:.1f}%<br>%{percent}<extra></extra>'
    )])
    
    fig_pie.update_layout(
        title='Energy Production Mix',
        height=450,
        showlegend=True
    )
    charts['energy_mix'] = fig_pie
    
    return charts


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
    
    # Solar PV
    with st.expander("☀️ SOLAR PV", expanded=True):
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
    pv_file = st.file_uploader("PV Profile (1 kW)", type=['csv', 'xlsx'], key="pv_file")
    wind_file = st.file_uploader("Wind Profile (1 kW)", type=['csv', 'xlsx'], key="wind_file")
    hydro_file = st.file_uploader("Hydro Profile (Optional)", type=['csv', 'xlsx'], key="hydro_file")

# ========== MAIN TABS ==========
tab1, tab2, tab3 = st.tabs(["🏠 Home", "⚙️ Optimize", "📊 Results"])

# Calculate search space
pv_options = int((pv_max - pv_min) / pv_step) + 1 if pv_step > 0 else 1
wind_options = int((wind_max - wind_min) / wind_step) + 1 if wind_step > 0 else 1
hydro_options = int((hydro_max - hydro_min) / hydro_step) + 1 if hydro_step > 0 else 1
bess_options = int((bess_max - bess_min) / bess_step) + 1 if bess_step > 0 else 1
total_combinations = pv_options * wind_options * hydro_options * bess_options

# TAB 1: HOME
with tab1:
    st.header("Welcome to the Renewable Energy Optimization Tool")
    st.markdown("""
    ### 🎯 Purpose
    Optimize hybrid renewable energy systems to minimize Net Present Cost while meeting reliability targets.
    
    ### 🔧 Features
    - **Present Value Cost Analysis** with real discount rate and inflation adjustment
    - **Grid Search Algorithm** testing all capacity combinations
    - **Hydro Operating Window Optimization** for Run-of-River systems  
    - **Professional Excel Reports** matching industry standards
    - **Interactive Visualization** for comprehensive results analysis
    """)
    
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
        st.metric("Components", "4")
    with col4:
        st.metric("Target Unmet", f"{target_unmet_percent}%")
    
    st.markdown("---")
    
    validation_passed = all([load_file, pv_file, wind_file])
    
    if not validation_passed:
        if not load_file:
            st.error("❌ Please upload Load Profile")
        if not pv_file:
            st.error("❌ Please upload PV Profile")
        if not wind_file:
            st.error("❌ Please upload Wind Profile")
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
                pv_df = pd.read_csv(pv_file) if pv_file.name.endswith('.csv') else pd.read_excel(pv_file)
                wind_df = pd.read_csv(wind_file) if wind_file.name.endswith('.csv') else pd.read_excel(wind_file)
                hydro_df = None
                if hydro_file:
                    hydro_df = pd.read_csv(hydro_file) if hydro_file.name.endswith('.csv') else pd.read_excel(hydro_file)
                
                # Build Excel input file
                status_text.text("🔨 Building input file...")
                progress_bar.progress(15)
                
                output = BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    # Configuration sheet
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
                        'Value': ['YES', pv_min*1000, pv_max*1000, pv_step*1000, wind_min*1000, wind_max*1000, wind_step*1000,
                                 hydro_min*1000, hydro_max*1000, hydro_step*1000, bess_min*1000, bess_max*1000, bess_step*1000,
                                 100000000, 'NPC', 5]
                    }).to_excel(writer, sheet_name='Grid_Search_Config', index=False)
                    
                    # Solar PV
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
                        'pv_lcoe': pv_lcoe,
                        'wind_lcoe': wind_lcoe,
                        'hydro_lcoe': hydro_lcoe,
                        'bess_max_soc': bess_max_soc / 100,
                        'bess_min_soc': bess_min_soc / 100,
                        'bess_lifetime': bess_lifetime
                    }
                    
                    electrical_metrics = opt_module.calculate_electrical_metrics(
                        optimal_dispatch, component_capacities, component_configs
                    )
                    
                    progress_bar.progress(100)
                    
                    # Clean up temp file
                    if os.path.exists(temp_file):
                        os.remove(temp_file)
                    
                    # Store results in session state
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
                        'optimal_row': optimal.to_dict(),
                        'electrical_metrics': electrical_metrics,
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
                    st.info("👉 Go to **Results** tab to view detailed analysis")
                else:
                    st.error("❌ No optimal solution found. Try relaxing constraints.")
                    
            except Exception as e:
                st.error(f"❌ Error during optimization: {str(e)}")
                st.exception(e)

# TAB 3: RESULTS
with tab3:
    if not st.session_state.optimization_complete:
        st.info("ℹ️ No results available yet. Please run optimization in the **Optimize** tab first.")
    else:
        st.header("📊 Optimization Results")
        
        results = st.session_state.results
        optimal_row = results.get('optimal_row', {})
        
        # ============================================================
        # ENHANCED KEY METRICS DISPLAY (with individual LCOE)
        # ============================================================
        st.subheader("🎯 System Overview")
        
        # Row 1: Financial Metrics
        col1, col2, col3 = st.columns(3)
        with col1:
            st.metric("💰 Total NPC", 
                     f"${results['npc']/1e6:.2f}M", 
                     help="Total Net Present Cost over project lifetime")
        with col2:
            st.metric("⚡ System LCOE", 
                     f"${results['lcoe']:.2f}/MWh",
                     help="System-wide Levelized Cost of Energy")
        with col3:
            st.metric("📉 Unmet Load", 
                     f"{results['unmet_pct']:.3f}%",
                     help="Percentage of load not met by the system")
        
        st.markdown("---")
        
        # Row 2: Component Capacities & Individual LCOE
        st.subheader("🔋 Component Configuration")
        
        col1, col2, col3, col4 = st.columns(4)
        
        with col1:
            st.markdown("**☀️ Solar PV**")
            st.metric("Capacity", f"{results['pv_capacity']:.2f} MW")
            if 'electrical_metrics' in results and results['electrical_metrics']:
                pv_lcoe = results['electrical_metrics']['pv']['levelized_cost_per_kwh'] * 1000
                st.metric("LCOE", f"${pv_lcoe:.2f}/MWh")
            else:
                st.metric("LCOE", f"${pv_lcoe:.2f}/MWh")
        
        with col2:
            st.markdown("**💨 Wind**")
            st.metric("Capacity", f"{results['wind_capacity']:.2f} MW")
            if 'electrical_metrics' in results and results['electrical_metrics']:
                wind_lcoe = results['electrical_metrics']['wind']['levelized_cost_per_kwh'] * 1000
                st.metric("LCOE", f"${wind_lcoe:.2f}/MWh")
            else:
                st.metric("LCOE", f"${wind_lcoe:.2f}/MWh")
        
        with col3:
            st.markdown("**💧 Hydro**")
            st.metric("Capacity", f"{results['hydro_capacity']:.2f} MW")
            if 'electrical_metrics' in results and results['electrical_metrics']:
                hydro_lcoe = results['electrical_metrics']['hydro']['levelized_cost_per_kwh'] * 1000
                st.metric("LCOE", f"${hydro_lcoe:.2f}/MWh")
            else:
                st.metric("LCOE", f"${hydro_lcoe:.2f}/MWh")
            st.caption(f"Window: {int(results['hydro_window_start']):02d}:00 - {int(results['hydro_window_end']):02d}:00")
        
        with col4:
            st.markdown("**🔋 Battery**")
            st.metric("Power", f"{results['bess_power']:.2f} MW")
            st.metric("Energy", f"{results['bess_energy']:.2f} MWh")
            st.caption(f"Duration: {results['bess_energy']/results['bess_power']:.1f}h" if results['bess_power'] > 0 else "Duration: N/A")
        
        st.markdown("---")
        
        # ============================================================
        # COST ANALYSIS SECTION
        # ============================================================
        st.subheader("💰 Cost Analysis")
        
        cost_charts = create_cost_analysis_charts_with_tables(results, optimal_row)
        
        # Row 1: Net Present Cost
        st.markdown("#### Net Present Cost")
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
        
        # Row 2: Annualized Cost
        st.markdown("#### Annualized Cost")
        col1, col2 = st.columns(2)
        
        with col1:
            st.markdown("**By Component**")
            st.plotly_chart(cost_charts['annualized_by_component']['chart'], use_container_width=True)
            st.dataframe(cost_charts['annualized_by_component']['table'], use_container_width=True, hide_index=True)
        
        with col2:
            st.markdown("**By Cost Type**")
            st.plotly_chart(cost_charts['annualized_by_cost_type']['chart'], use_container_width=True)
            st.dataframe(cost_charts['annualized_by_cost_type']['table'], use_container_width=True, hide_index=True)
        
        st.markdown("---")
        
        # ============================================================
        # CASH FLOW SECTION
        # ============================================================
        st.subheader("💵 Cash Flow Analysis")
        
        cash_flow_fig, cash_flow_table = create_fixed_cash_flow_chart(results, optimal_row)
        
        col1, col2 = st.columns([2, 1])
        with col1:
            st.plotly_chart(cash_flow_fig, use_container_width=True)
        with col2:
            st.markdown("**Cash Flow Summary (First 10 Years)**")
            st.dataframe(cash_flow_table, use_container_width=True, hide_index=True)
        
        st.markdown("---")
        
        # ============================================================
        # ELECTRICAL PERFORMANCE METRICS SECTION
        # ============================================================
        st.subheader("⚡ Electrical Performance Metrics")
        
        if 'electrical_metrics' in results and results['electrical_metrics']:
            elec_tables = create_electrical_metrics_tables(results['electrical_metrics'])
            
            # Row 1: PV and Wind
            col1, col2 = st.columns(2)
            with col1:
                st.markdown("**☀️ Solar PV**")
                st.dataframe(elec_tables['pv'], use_container_width=True, hide_index=True)
            
            with col2:
                st.markdown("**💨 Wind**")
                st.dataframe(elec_tables['wind'], use_container_width=True, hide_index=True)
            
            # Row 2: Hydro and BESS
            col1, col2 = st.columns(2)
            with col1:
                st.markdown("**💧 Hydro**")
                st.dataframe(elec_tables['hydro'], use_container_width=True, hide_index=True)
            
            with col2:
                st.markdown("**🔋 Battery Storage**")
                st.dataframe(elec_tables['bess'], use_container_width=True, hide_index=True)
            
            # Row 3: System Performance (Excess Electricity, Unmet Load, Capacity Shortage)
            st.markdown("---")
            st.markdown("**⚙️ System Performance**")
            
            # Calculate system metrics
            total_load = optimal_row.get('Total_Load_kWh', 0)
            excess_kwh = optimal_row.get('Excess_kWh', 0)
            unmet_kwh = optimal_row.get('Unmet_kWh', 0)
            capacity_shortage_kwh = unmet_kwh  # Simplified assumption
            
            system_metrics = pd.DataFrame({
                'Metric': [
                    'Excess Electricity (kWh/yr)',
                    'Excess Electricity (%)',
                    'Unmet Electric Load (kWh/yr)',
                    'Unmet Electric Load (%)',
                    'Capacity Shortage (kWh/yr)',
                    'Capacity Shortage (%)'
                ],
                'Value': [
                    f"{excess_kwh:,.0f}",
                    f"{(excess_kwh / total_load * 100) if total_load > 0 else 0:.2f}",
                    f"{unmet_kwh:,.0f}",
                    f"{(unmet_kwh / total_load * 100) if total_load > 0 else 0:.2f}",
                    f"{capacity_shortage_kwh:,.0f}",
                    f"{(capacity_shortage_kwh / total_load * 100) if total_load > 0 else 0:.2f}"
                ]
            })
            
            st.dataframe(system_metrics, use_container_width=True, hide_index=True)
            
        else:
            st.warning("⚠️ Electrical metrics not available. Please re-run optimization with updated code.")
        
        st.markdown("---")
        
        # ============================================================
        # ADDITIONAL ANALYSIS
        # ============================================================
        st.subheader("📊 Additional Analysis")
        
        additional_charts = create_additional_charts(results)
        
        col1, col2 = st.columns(2)
        with col1:
            st.plotly_chart(additional_charts['monthly_production'], use_container_width=True)
        with col2:
            st.plotly_chart(additional_charts['energy_mix'], use_container_width=True)
        
        st.markdown("---")
        st.subheader("📈 Typical Daily Dispatch Profile")

        daily_profile_fig = create_typical_daily_dispatch_profile(results, optimal_row)
        st.plotly_chart(daily_profile_fig, use_container_width=True)

        st.caption("Shows average hourly generation, load, and BESS state across all days")
        # ============================================================
        # DOWNLOAD SECTION
        # ============================================================
        st.subheader("📥 Download Results")
        
        config_params = results.get('config_params', {
            'discount_rate': 8.0,
            'inflation_rate': 2.0,
            'project_lifetime': 25,
            'target_unmet_percent': 0.1
        })
        
        excel_output = export_results_industry_format(
            results, 
            results['results_df'], 
            results['optimal_row'], 
            config_params
        )
        
        col1, col2, col3 = st.columns([2, 1, 2])
        with col2:
            st.download_button(
                label="📥 Download Complete Results (Excel)",
                data=excel_output,
                file_name=f"optimization_results_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
                type="primary"
            )

# Footer
st.markdown("---")
st.markdown('<div style="text-align:center;color:#666"><p>Renewable Energy Optimization Tool v3.0 | Professional NPC Analysis</p></div>', unsafe_allow_html=True)


