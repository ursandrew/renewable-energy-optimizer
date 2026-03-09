"""
DEGRADATION ANALYSIS FUNCTIONS
To be added to optimize_gridsearch_hydro_CLEAN_FIXED.py after line 779
"""

import pandas as pd
import numpy as np


# ==============================================================================
# DEGRADATION ANALYSIS FUNCTIONS
# ==============================================================================

def apply_pv_degradation_simple(pv_profile, year, annual_rate_pct):
    """
    Apply PV degradation using simple annual rate.
    
    Args:
        pv_profile: 8760-hour PV generation array (kW)
        year: Project year (1-25)
        annual_rate_pct: Annual degradation rate (e.g., 0.40 for 0.4%/year)
    
    Returns:
        Degraded PV profile (kW)
    """
    annual_rate = annual_rate_pct / 100  # Convert to decimal
    degradation_factor = (1 - annual_rate) ** (year - 1)
    return pv_profile * degradation_factor


def apply_pv_degradation_curve(pv_profile, year, degradation_curve):
    """
    Apply PV degradation using custom curve.
    
    Args:
        pv_profile: 8760-hour PV generation array (kW)
        year: Project year (1-25)
        degradation_curve: Dict {year: cumulative_degradation_%}
    
    Returns:
        Degraded PV profile (kW)
    """
    if year not in degradation_curve:
        return pv_profile  # No degradation data for this year
    
    cumulative_deg_pct = degradation_curve[year]
    degradation_factor = 1 - (cumulative_deg_pct / 100)
    return pv_profile * degradation_factor


def apply_bess_degradation(bess_capacity_kwh, bess_power_kw, year, degradation_data):
    """
    Apply BESS degradation to capacity and efficiencies.
    
    Args:
        bess_capacity_kwh: Initial BESS capacity (kWh)
        bess_power_kw: BESS power rating (kW)
        year: Project year (1-25)
        degradation_data: Dict with year -> {'capacity': %, 'charge_eff': %, 'discharge_eff': %}
    
    Returns:
        Tuple: (degraded_capacity_kwh, charge_eff, discharge_eff)
    """
    if year not in degradation_data:
        # No degradation data - return original
        return bess_capacity_kwh, None, None
    
    year_data = degradation_data[year]
    
    # Apply capacity retention
    capacity_retention_pct = year_data['capacity']
    degraded_capacity = bess_capacity_kwh * (capacity_retention_pct / 100)
    
    # Get efficiencies (as decimals)
    charge_eff = year_data['charge_eff'] / 100
    discharge_eff = year_data['discharge_eff'] / 100
    
    return degraded_capacity, charge_eff, discharge_eff


def run_multi_year_degradation_analysis(
    optimal_config, load_profile, pvsyst_profile, wind_profile,
    solar_config, wind_config, hydro_config, bess_config,
    project_lifetime=25,
    pv_degradation_type=None,  # 'simple' or 'curve' or None
    pv_degradation_data=None,  # annual_rate OR curve dict
    bess_degradation_data=None  # BESS degradation dict or None
):
    """
    Run degradation analysis for all project years.
    
    Args:
        optimal_config: Dict with optimal system configuration
        load_profile: 8760-hour load profile
        pvsyst_profile: 8760-hour PV profile (1 kW baseline)
        wind_profile: 8760-hour wind profile
        solar_config, wind_config, hydro_config, bess_config: Component configs
        project_lifetime: Project duration (years)
        pv_degradation_type: 'simple', 'curve', or None
        pv_degradation_data: If simple: annual_rate_pct, if curve: {year: deg%}
        bess_degradation_data: {year: {'capacity': %, 'charge_eff': %, 'discharge_eff': %}}
    
    Returns:
        Dictionary with:
        - yearly_metrics: DataFrame with annual performance
        - selected_year_dispatch: Dict with dispatch for years [1,2,5,10,15,20,25]
        - degradation_summary: Summary statistics
    """
    print("\n" + "="*70)
    print("RUNNING 25-YEAR DEGRADATION ANALYSIS")
    print("="*70)
    
    yearly_results = []
    selected_years = [1, 2, 5, 10, 15, 20, 25]
    yearly_dispatch = {}
    
    # Baseline (Year 1) performance
    baseline_pv_energy = optimal_config.get('PV_Energy_kWh', 0)
    baseline_bess_capacity = optimal_config.get('BESS_Capacity_kWh', 0)
    
    for year in range(1, project_lifetime + 1):
        if year % 5 == 0 or year == 1:
            print(f"  Processing Year {year}...")
        
        # Apply PV degradation
        if pv_degradation_type == 'simple' and pv_degradation_data is not None:
            pv_gen_degraded = apply_pv_degradation_simple(
                pvsyst_profile, year, pv_degradation_data
            )
            pv_deg_pct = (1 - (1 - pv_degradation_data/100) ** (year-1)) * 100
        elif pv_degradation_type == 'curve' and pv_degradation_data is not None:
            pv_gen_degraded = apply_pv_degradation_curve(
                pvsyst_profile, year, pv_degradation_data
            )
            pv_deg_pct = pv_degradation_data.get(year, 0)
        else:
            pv_gen_degraded = pvsyst_profile
            pv_deg_pct = 0
        
        # Apply BESS degradation
        if bess_degradation_data is not None:
            bess_capacity_degraded, charge_eff_deg, discharge_eff_deg = apply_bess_degradation(
                optimal_config['BESS_Capacity_kWh'],
                optimal_config['BESS_Power_kW'],
                year,
                bess_degradation_data
            )
            
            # Update BESS config with degraded efficiencies
            bess_config_degraded = bess_config.copy()
            if charge_eff_deg is not None:
                bess_config_degraded['charge_eff'] = charge_eff_deg
                bess_config_degraded['discharge_eff'] = discharge_eff_deg
            
            bess_retention_pct = (bess_capacity_degraded / baseline_bess_capacity * 100) if baseline_bess_capacity > 0 else 100
        else:
            bess_capacity_degraded = optimal_config['BESS_Capacity_kWh']
            bess_config_degraded = bess_config
            bess_retention_pct = 100
        
        # Run dispatch simulation for this year
        dispatch_df = calculate_dispatch_with_hydro(
            load_profile,
            pv_gen_degraded,
            wind_profile,
            optimal_config['PV_kW'],
            optimal_config['Wind_kW'],
            optimal_config['Hydro_kW'],
            optimal_config['BESS_Power_kW'],
            bess_capacity_degraded,
            solar_config,
            wind_config,
            hydro_config,
            bess_config_degraded,
            int(optimal_config.get('Hydro_Window_Start', 0)),
            int(optimal_config.get('Hydro_Window_End', 24))
        )
        
        # Calculate annual metrics
        total_load = dispatch_df['Load_kW'].sum()
        total_unmet = dispatch_df['Unmet_Load_kW'].sum()
        total_served = total_load - total_unmet
        
        annual_metrics = {
            'Year': year,
            'PV_Degradation_%': pv_deg_pct,
            'BESS_Retention_%': bess_retention_pct,
            'PV_Energy_MWh': dispatch_df['PV_Available_kW'].sum() / 1000,
            'Wind_Energy_MWh': dispatch_df['Wind_Output_kW'].sum() / 1000,
            'Hydro_Energy_MWh': dispatch_df['Hydro_Output_kW'].sum() / 1000,
            'Load_MWh': total_load / 1000,
            'Served_MWh': total_served / 1000,
            'Unmet_MWh': total_unmet / 1000,
            'Unmet_%': (total_unmet / total_load * 100) if total_load > 0 else 0,
            'BESS_Throughput_MWh': dispatch_df['BESS_Discharge_wieff_kW'].sum() / 1000,
            'Curtailment_MWh': dispatch_df['Curtailment_kW'].sum() / 1000,
        }
        
        yearly_results.append(annual_metrics)
        
        # Store dispatch for selected years
        if year in selected_years:
            dispatch_df['Year'] = year
            yearly_dispatch[f'Year_{year}'] = dispatch_df.copy()
    
    # Create summary
    yearly_metrics_df = pd.DataFrame(yearly_results)
    
    degradation_summary = {
        'pv_degradation_year_1': yearly_metrics_df.loc[0, 'PV_Degradation_%'],
        'pv_degradation_year_25': yearly_metrics_df.loc[24, 'PV_Degradation_%'],
        'bess_retention_year_1': yearly_metrics_df.loc[0, 'BESS_Retention_%'],
        'bess_retention_year_25': yearly_metrics_df.loc[24, 'BESS_Retention_%'],
        'avg_unmet_pct': yearly_metrics_df['Unmet_%'].mean(),
        'max_unmet_pct': yearly_metrics_df['Unmet_%'].max(),
        'total_energy_served_25yr_GWh': yearly_metrics_df['Served_MWh'].sum() / 1000,
    }
    
    print(f"\n✓ Degradation Analysis Complete")
    print(f"  PV Degradation: {pv_deg_pct:.2f}% at Year 25")
    print(f"  BESS Retention: {bess_retention_pct:.1f}% at Year 25")
    print(f"  Average Unmet Load: {degradation_summary['avg_unmet_pct']:.2f}%")
    print("="*70)
    
    return {
        'yearly_metrics': yearly_metrics_df,
        'selected_year_dispatch': yearly_dispatch,
        'degradation_summary': degradation_summary,
        'optimal_config': optimal_config
    }


# ==============================================================================
# PRESET DEGRADATION DATA
# ==============================================================================

# BESS degradation presets based on battery chemistry (from user's data)
BESS_DEGRADATION_PRESETS = {
    "Lithium NMC (Standard)": {
        1: {'capacity': 97.32, 'charge_eff': 88.84, 'discharge_eff': 98.50},
        2: {'capacity': 95.06, 'charge_eff': 88.75, 'discharge_eff': 98.50},
        3: {'capacity': 93.01, 'charge_eff': 88.55, 'discharge_eff': 98.50},
        4: {'capacity': 91.10, 'charge_eff': 88.47, 'discharge_eff': 98.50},
        5: {'capacity': 89.27, 'charge_eff': 88.27, 'discharge_eff': 98.50},
        6: {'capacity': 87.50, 'charge_eff': 88.09, 'discharge_eff': 98.50},
        7: {'capacity': 85.76, 'charge_eff': 87.99, 'discharge_eff': 98.50},
        8: {'capacity': 84.05, 'charge_eff': 87.81, 'discharge_eff': 98.50},
        9: {'capacity': 82.37, 'charge_eff': 87.61, 'discharge_eff': 98.50},
        10: {'capacity': 80.73, 'charge_eff': 87.52, 'discharge_eff': 98.50},
        11: {'capacity': 79.12, 'charge_eff': 87.43, 'discharge_eff': 98.50},
        12: {'capacity': 77.54, 'charge_eff': 87.33, 'discharge_eff': 98.50},
        13: {'capacity': 75.99, 'charge_eff': 87.24, 'discharge_eff': 98.50},
        14: {'capacity': 74.48, 'charge_eff': 87.14, 'discharge_eff': 98.50},
        15: {'capacity': 72.99, 'charge_eff': 87.14, 'discharge_eff': 98.50},
        16: {'capacity': 71.53, 'charge_eff': 87.05, 'discharge_eff': 98.50},
        17: {'capacity': 70.10, 'charge_eff': 86.96, 'discharge_eff': 98.50},
        18: {'capacity': 68.70, 'charge_eff': 86.84, 'discharge_eff': 98.50},
        19: {'capacity': 67.32, 'charge_eff': 86.72, 'discharge_eff': 98.50},
        20: {'capacity': 65.98, 'charge_eff': 86.58, 'discharge_eff': 98.50},
        21: {'capacity': 97.32, 'charge_eff': 88.84, 'discharge_eff': 98.50},  # Replacement
        22: {'capacity': 95.06, 'charge_eff': 88.75, 'discharge_eff': 98.50},
        23: {'capacity': 93.01, 'charge_eff': 88.55, 'discharge_eff': 98.50},
        24: {'capacity': 91.10, 'charge_eff': 88.47, 'discharge_eff': 98.50},
        25: {'capacity': 89.27, 'charge_eff': 88.27, 'discharge_eff': 98.50},
    },
    
    "Lithium LFP (Long Life)": {
        1: {'capacity': 98.50, 'charge_eff': 90.00, 'discharge_eff': 98.50},
        2: {'capacity': 97.80, 'charge_eff': 89.90, 'discharge_eff': 98.50},
        3: {'capacity': 97.20, 'charge_eff': 89.80, 'discharge_eff': 98.50},
        4: {'capacity': 96.60, 'charge_eff': 89.70, 'discharge_eff': 98.50},
        5: {'capacity': 96.00, 'charge_eff': 89.60, 'discharge_eff': 98.50},
        6: {'capacity': 95.50, 'charge_eff': 89.50, 'discharge_eff': 98.50},
        7: {'capacity': 95.00, 'charge_eff': 89.40, 'discharge_eff': 98.50},
        8: {'capacity': 94.50, 'charge_eff': 89.30, 'discharge_eff': 98.50},
        9: {'capacity': 94.00, 'charge_eff': 89.20, 'discharge_eff': 98.50},
        10: {'capacity': 93.50, 'charge_eff': 89.10, 'discharge_eff': 98.50},
        11: {'capacity': 93.00, 'charge_eff': 89.00, 'discharge_eff': 98.50},
        12: {'capacity': 92.60, 'charge_eff': 88.95, 'discharge_eff': 98.50},
        13: {'capacity': 92.20, 'charge_eff': 88.90, 'discharge_eff': 98.50},
        14: {'capacity': 91.80, 'charge_eff': 88.85, 'discharge_eff': 98.50},
        15: {'capacity': 91.40, 'charge_eff': 88.80, 'discharge_eff': 98.50},
        16: {'capacity': 91.00, 'charge_eff': 88.75, 'discharge_eff': 98.50},
        17: {'capacity': 90.60, 'charge_eff': 88.70, 'discharge_eff': 98.50},
        18: {'capacity': 90.20, 'charge_eff': 88.65, 'discharge_eff': 98.50},
        19: {'capacity': 89.80, 'charge_eff': 88.60, 'discharge_eff': 98.50},
        20: {'capacity': 89.50, 'charge_eff': 88.55, 'discharge_eff': 98.50},
        21: {'capacity': 89.20, 'charge_eff': 88.50, 'discharge_eff': 98.50},
        22: {'capacity': 88.90, 'charge_eff': 88.45, 'discharge_eff': 98.50},
        23: {'capacity': 88.60, 'charge_eff': 88.40, 'discharge_eff': 98.50},
        24: {'capacity': 88.30, 'charge_eff': 88.35, 'discharge_eff': 98.50},
        25: {'capacity': 88.00, 'charge_eff': 88.30, 'discharge_eff': 98.50},
    },
    
    "Sodium-Ion (Emerging)": {
        1: {'capacity': 95.00, 'charge_eff': 85.00, 'discharge_eff': 95.00},
        2: {'capacity': 94.20, 'charge_eff': 84.90, 'discharge_eff': 95.00},
        3: {'capacity': 93.40, 'charge_eff': 84.80, 'discharge_eff': 95.00},
        4: {'capacity': 92.60, 'charge_eff': 84.70, 'discharge_eff': 95.00},
        5: {'capacity': 91.80, 'charge_eff': 84.60, 'discharge_eff': 95.00},
        6: {'capacity': 91.00, 'charge_eff': 84.50, 'discharge_eff': 95.00},
        7: {'capacity': 90.20, 'charge_eff': 84.40, 'discharge_eff': 95.00},
        8: {'capacity': 89.50, 'charge_eff': 84.30, 'discharge_eff': 95.00},
        9: {'capacity': 88.80, 'charge_eff': 84.20, 'discharge_eff': 95.00},
        10: {'capacity': 88.10, 'charge_eff': 84.10, 'discharge_eff': 95.00},
        11: {'capacity': 87.50, 'charge_eff': 84.00, 'discharge_eff': 95.00},
        12: {'capacity': 86.90, 'charge_eff': 83.95, 'discharge_eff': 95.00},
        13: {'capacity': 86.30, 'charge_eff': 83.90, 'discharge_eff': 95.00},
        14: {'capacity': 85.80, 'charge_eff': 83.85, 'discharge_eff': 95.00},
        15: {'capacity': 85.30, 'charge_eff': 83.80, 'discharge_eff': 95.00},
        16: {'capacity': 84.80, 'charge_eff': 83.75, 'discharge_eff': 95.00},
        17: {'capacity': 84.30, 'charge_eff': 83.70, 'discharge_eff': 95.00},
        18: {'capacity': 83.80, 'charge_eff': 83.65, 'discharge_eff': 95.00},
        19: {'capacity': 83.40, 'charge_eff': 83.60, 'discharge_eff': 95.00},
        20: {'capacity': 83.00, 'charge_eff': 83.55, 'discharge_eff': 95.00},
        21: {'capacity': 82.60, 'charge_eff': 83.50, 'discharge_eff': 95.00},
        22: {'capacity': 82.20, 'charge_eff': 83.45, 'discharge_eff': 95.00},
        23: {'capacity': 81.80, 'charge_eff': 83.40, 'discharge_eff': 95.00},
        24: {'capacity': 81.50, 'charge_eff': 83.35, 'discharge_eff': 95.00},
        25: {'capacity': 81.20, 'charge_eff': 83.30, 'discharge_eff': 95.00},
    }
}
