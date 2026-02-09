"""
DEGRADATION ANALYSIS MODULE FOR STREAMLIT - FIXED VERSION
==========================================================
Full degradation engine for PV + BESS systems with COMPLETE hourly simulation

FIXES IN THIS VERSION:
1. Hour numbering: 0-23 for each year (not continuous 0-8759)
2. BESS SOC carryover: Maintains SOC from year to year (not reset to 50%)
3. Column order: Matches Anaconda output exactly
4. Proper degradation application from Year 1

IMPORTANT NOTES:
- Capacity Factor: This is based on actual energy produced vs nameplate capacity
  Low CF (~6%) is normal when BESS stores excess PV for later use
  High unmet load means PV+BESS can't meet demand, not that PV isn't producing
- BESS SOC: Starts at 50% for Year 1, then carries over year-to-year
- Degradation: Applied from Year 1 (94.46% retention = 5.54% degradation)

Author: SJ
Version: 3.1 - Fixed Hour Numbering, SOC Carryover, Column Order
"""

import pandas as pd
import numpy as np

# ==============================================================================
# OEM DEGRADATION DATA
# ==============================================================================

# PV Degradation - Cumulative percentage loss
PV_DEG = {
    1: 0, 2: 0.41, 3: 0.82, 4: 1.22, 5: 1.63, 6: 2.04, 7: 2.45, 8: 2.86, 9: 3.27, 10: 3.67,
    11: 4.08, 12: 4.49, 13: 4.90, 14: 5.31, 15: 5.71, 16: 6.12, 17: 6.53, 18: 6.94, 19: 7.35, 20: 7.76,
    21: 8.16, 22: 8.57, 23: 8.98, 24: 9.39, 25: 9.80
}

# BESS Capacity Retention - Percentage of original capacity
BESS_CAP_RET = {
    1: 94.46, 2: 92.14, 3: 90.33, 4: 88.71, 5: 87.19, 6: 85.75, 7: 84.37, 8: 83.00, 9: 81.70, 10: 80.45,
    11: 79.20, 12: 78.00, 13: 76.76, 14: 75.57, 15: 74.40, 16: 73.23, 17: 72.10, 18: 70.96, 19: 69.84, 20: 68.73
}

# BESS Charging Efficiency Retention - Fraction of Year 1 value
BESS_CHARGE_EFF_RETENTION = {
    1: 1.0000, 2: 0.9989, 3: 0.9981, 4: 0.9972, 5: 0.9963,
    6: 0.9955, 7: 0.9947, 8: 0.9939, 9: 0.9930, 10: 0.9922,
    11: 0.9914, 12: 0.9906, 13: 0.9897, 14: 0.9888, 15: 0.9881,
    16: 0.9873, 17: 0.9864, 18: 0.9856, 19: 0.9848, 20: 0.9840
}

# BESS Discharging Efficiency Retention - Fraction of Year 1 value
BESS_DISCHARGE_EFF_RETENTION = {
    1: 1.0000, 2: 0.9989, 3: 0.9981, 4: 0.9973, 5: 0.9964,
    6: 0.9955, 7: 0.9948, 8: 0.9939, 9: 0.9930, 10: 0.9923,
    11: 0.9915, 12: 0.9907, 13: 0.9898, 14: 0.9889, 15: 0.9882,
    16: 0.9874, 17: 0.9865, 18: 0.9858, 19: 0.9849, 20: 0.9841
}

# Global variable for input file path
INPUT_FILE = None


# ==============================================================================
# IMPORT BASE OPTIMIZATION FUNCTIONS
# ==============================================================================

def read_inputs():
    """
    Wrapper for base module read_inputs.
    
    CRITICAL: This sets the INPUT_FILE in the base module before calling read_inputs.
    """
    import optimize_gridsearch_hydro_static_STREAMLITCHECK as base_module
    
    # Set the INPUT_FILE in the base module to match this module's INPUT_FILE
    if INPUT_FILE is not None:
        base_module.INPUT_FILE = INPUT_FILE
    
    return base_module.read_inputs()


def calculate_npc_homer_style(*args, **kwargs):
    """Wrapper for base module calculate_npc_homer_style."""
    import optimize_gridsearch_hydro_static_STREAMLITCHECK as base_module
    return base_module.calculate_npc_homer_style(*args, **kwargs)


def calculate_electrical_metrics(dispatch_df, component_capacities, component_configs,
                                 npc_breakdown, project_lifetime):
    """Wrapper for base module calculate_electrical_metrics."""
    import optimize_gridsearch_hydro_static_STREAMLITCHECK as base_module
    return base_module.calculate_electrical_metrics(
        dispatch_df, component_capacities, component_configs,
        npc_breakdown, project_lifetime
    )


def find_optimal_solution(results_df):
    """Wrapper for base module find_optimal_solution."""
    import optimize_gridsearch_hydro_static_STREAMLITCHECK as base_module
    return base_module.find_optimal_solution(results_df)


# ==============================================================================
# GRID SEARCH OPTIMIZATION (NO DEGRADATION DURING OPTIMIZATION)
# ==============================================================================

def grid_search_optimize_hydro(config, grid_config, solar, wind, hydro, bess, 
                               load_profile, pvsyst_profile, wind_profile, hydro_profile):
    """
    Grid search optimization WITHOUT degradation.
    
    This finds the optimal configuration using Year 1 performance.
    Degradation is applied separately via run_degradation_analysis().
    """
    
    # Import and call base optimization
    import optimize_gridsearch_hydro_static_STREAMLITCHECK as base_module
    
    # Call base optimization (no degradation during optimization)
    results_df = base_module.grid_search_optimize_hydro(
        config, grid_config, solar, wind, hydro, bess,
        load_profile, pvsyst_profile, wind_profile, hydro_profile
    )
    
    return results_df


# ==============================================================================
# COMPLETE 25-YEAR DEGRADATION ANALYSIS WITH HOURLY SIMULATION
# ==============================================================================

def run_degradation_analysis_complete(optimal_row, config_params, profiles, 
                                      apply_pv=True, apply_bess=True,
                                      years_to_export=[1, 2, 5, 10, 15, 20, 25],
                                      export_all_years=False):
    """
    Run COMPLETE 25-year degradation analysis with HOURLY dispatch simulation.
    
    Parameters:
    -----------
    optimal_row : dict
        Optimal solution from grid search
    config_params : dict
        Configuration parameters (discount_rate, project_lifetime, etc.)
    profiles : dict
        Dictionary with keys 'load', 'pv', 'wind', 'hydro' containing hourly profiles
    apply_pv : bool
        Apply PV degradation
    apply_bess : bool
        Apply BESS degradation
    years_to_export : list
        Which years to export hourly data (default: [1, 2, 5, 10, 15, 20, 25])
    export_all_years : bool
        If True, export all 25 years (overrides years_to_export)
    
    Returns:
    --------
    dict : Complete degradation analysis results with hourly dispatch
    
    IMPORTANT FIXES:
    - Hour numbering: 0-23 for each year (not continuous)
    - BESS SOC: Carries over from year to year (not reset to 50%)
    - Column order: Matches Anaconda output exactly
    """
    
    print("\n" + "="*80)
    print("RUNNING COMPLETE 25-YEAR DEGRADATION ANALYSIS")
    print("="*80)
    
    # Extract capacities
    pv_kw = optimal_row.get('PV_kW', 0)
    wind_kw = optimal_row.get('Wind_kW', 0)
    hydro_kw = optimal_row.get('Hydro_kW', 0)
    bess_power_kw = optimal_row.get('BESS_Power_kW', 0)
    bess_capacity_kwh = optimal_row.get('BESS_Capacity_kWh', 0)
    
    hydro_start = int(optimal_row.get('Hydro_Window_Start', 8))
    hydro_end = int(optimal_row.get('Hydro_Window_End', 16))
    
    # Extract costs
    npc_y1 = optimal_row.get('NPC_$', 0)
    lcoe_y1 = optimal_row.get('LCOE_$/MWh', 0)
    bess_npc = optimal_row.get('BESS_NPC_$', 0)
    
    # Get parameters
    discount_rate = config_params.get('discount_rate', 8.0)
    if discount_rate > 1:
        discount_rate = discount_rate / 100
    
    project_lifetime = config_params.get('project_lifetime', 25)
    
    # BESS parameters - CRITICAL: Use sidebar values correctly
    base_charge_eff = config_params.get('bess_charge_eff', 92.94)
    if base_charge_eff > 1:
        base_charge_eff = base_charge_eff / 100
    
    base_discharge_eff = config_params.get('bess_discharge_eff', 91.78)
    if base_discharge_eff > 1:
        base_discharge_eff = base_discharge_eff / 100
    
    max_soc = config_params.get('bess_max_soc', 100.0)
    if max_soc > 1:
        max_soc = max_soc / 100
    
    min_soc = config_params.get('bess_min_soc', 20.0)
    if min_soc > 1:
        min_soc = min_soc / 100
    
    bess_duration_h = bess_capacity_kwh / bess_power_kw if bess_power_kw > 0 else 2.0
    
    # Extract profiles
    load_profile_kw = profiles['load']
    pv_profile_pu = profiles['pv']
    wind_profile_pu = profiles['wind']
    hydro_profile_pu = profiles.get('hydro', np.ones(len(load_profile_kw)))
    
    hours_per_year = len(load_profile_kw)
    
    print(f"\nSystem Configuration:")
    print(f"  PV: {pv_kw:,.0f} kW")
    print(f"  Wind: {wind_kw:,.0f} kW")
    print(f"  Hydro: {hydro_kw:,.0f} kW (Hours {hydro_start}-{hydro_end})")
    print(f"  BESS: {bess_power_kw:,.0f} kW / {bess_capacity_kwh:,.0f} kWh")
    print(f"\nBESS Parameters:")
    print(f"  Base Charge Eff: {base_charge_eff*100:.2f}%")
    print(f"  Base Discharge Eff: {base_discharge_eff*100:.2f}%")
    print(f"  SOC Range: {min_soc*100:.1f}% - {max_soc*100:.1f}%")
    print(f"\nDegradation Settings:")
    print(f"  PV Degradation: {'ENABLED' if apply_pv else 'DISABLED'}")
    print(f"  BESS Degradation: {'ENABLED' if apply_bess else 'DISABLED'}")
    print(f"\nSimulating {project_lifetime} years with {hours_per_year} hours per year...")
    
    # Determine which years to export
    if export_all_years:
        years_to_export_list = list(range(1, min(project_lifetime + 1, 26)))
        print(f"Exporting ALL {len(years_to_export_list)} years")
    else:
        years_to_export_list = years_to_export
        print(f"Exporting selected years: {years_to_export_list}")
    
    # Storage for results
    yearly_results = []
    hourly_results_by_year = {}
    
    # Simulate each year
    for year in range(1, min(project_lifetime + 1, 26)):
        print(f"\n  Year {year}: ", end='', flush=True)
        
        # ======================
        # APPLY PV DEGRADATION
        # ======================
        if apply_pv:
            pv_deg_pct = PV_DEG.get(year, 9.8)
            pv_degradation_factor = 1.0 - (pv_deg_pct / 100)
            degraded_pv_capacity = pv_kw * pv_degradation_factor
        else:
            pv_deg_pct = 0
            pv_degradation_factor = 1.0
            degraded_pv_capacity = pv_kw
        
        # ======================
        # APPLY BESS DEGRADATION
        # ======================
        if apply_bess:
            # Handle replacement at year 21
            if year <= 20:
                bess_age = year
                replaced = False
            else:
                bess_age = year - 20  # Reset age after replacement
                replaced = (year == 21)
            
            # Capacity retention
            capacity_retention_pct = BESS_CAP_RET.get(bess_age, 70)
            capacity_retention = capacity_retention_pct / 100
            
            # Efficiency retention factors
            charge_eff_retention = BESS_CHARGE_EFF_RETENTION.get(bess_age, 0.98)
            discharge_eff_retention = BESS_DISCHARGE_EFF_RETENTION.get(bess_age, 0.98)
            
            # Apply degradation
            degraded_bess_capacity = bess_capacity_kwh * capacity_retention
            degraded_bess_power = bess_power_kw * capacity_retention
            charge_eff = base_charge_eff * charge_eff_retention
            discharge_eff = base_discharge_eff * discharge_eff_retention
        else:
            capacity_retention_pct = 100
            capacity_retention = 1.0
            charge_eff_retention = 1.0
            discharge_eff_retention = 1.0
            degraded_bess_capacity = bess_capacity_kwh
            degraded_bess_power = bess_power_kw
            charge_eff = base_charge_eff
            discharge_eff = base_discharge_eff
            replaced = False
        
        print(f"PV={pv_degradation_factor*100:.2f}%, BESS Cap={capacity_retention*100:.2f}%, ", end='', flush=True)
        
        # ================================
        # INITIALIZE BESS FOR THIS YEAR
        # ================================
        # CRITICAL: For Year 1, start at 50% of degraded capacity
        # For subsequent years, carry over final SOC from previous year
        if year == 1:
            soc_kwh = degraded_bess_capacity * 0.5  # Start at 50% for Year 1
        else:
            # Carry over SOC from previous year, but adjust for capacity change
            # If capacity decreased (degradation), maintain same percentage
            # If capacity increased (replacement), maintain same percentage
            if 'previous_soc_pct' in locals():
                soc_kwh = degraded_bess_capacity * previous_soc_pct
            else:
                soc_kwh = degraded_bess_capacity * 0.5  # Fallback
        
        # Yearly accumulators
        year_pv_gen = 0
        year_wind_gen = 0
        year_hydro_gen = 0
        year_pv_to_load = 0
        year_wind_to_load = 0
        year_hydro_to_load = 0
        year_pv_excess = 0
        year_bess_charge = 0
        year_bess_discharge = 0
        year_load = 0
        year_unmet = 0
        year_curtailment = 0
        
        # Store hourly data for selected years
        if year in years_to_export_list:
            hourly_data = []
            store_hourly = True
        else:
            store_hourly = False
        
        # ================================
        # HOURLY SIMULATION FOR THIS YEAR
        # ================================
        for h in range(hours_per_year):
            load_h = load_profile_kw[h]
            
            # Generation with degradation applied
            pv_available = degraded_pv_capacity * pv_profile_pu[h]
            wind_available = wind_kw * wind_profile_pu[h]  # No wind degradation
            
            # Hydro availability (time window + profile)
            hour_of_day = h % 24
            if hydro_start <= hour_of_day < hydro_end:
                hydro_available = hydro_kw * hydro_profile_pu[h]
            else:
                hydro_available = 0
            
            year_pv_gen += pv_available
            year_wind_gen += wind_available
            year_hydro_gen += hydro_available
            year_load += load_h
            
            # ================================
            # DISPATCH LOGIC (Merit Order)
            # ================================
            remaining_load = load_h
            
            # 1. HYDRO (highest priority)
            hydro_to_load = min(hydro_available, remaining_load)
            remaining_load -= hydro_to_load
            year_hydro_to_load += hydro_to_load
            
            # 2. PV
            pv_to_load = min(pv_available, remaining_load)
            remaining_load -= pv_to_load
            pv_excess = pv_available - pv_to_load
            year_pv_to_load += pv_to_load
            
            # 3. WIND
            wind_to_load = min(wind_available, remaining_load)
            remaining_load -= wind_to_load
            wind_excess = wind_available - wind_to_load
            year_wind_to_load += wind_to_load
            
            # 4. BESS CHARGING (if excess renewable energy)
            total_excess = pv_excess + wind_excess
            bess_charge_woeff = 0
            bess_charge_wieff = 0
            curtailment = 0
            
            if total_excess > 0 and degraded_bess_capacity > 0:
                # Maximum charge limited by power and available capacity
                max_charge_power = min(total_excess, degraded_bess_power)
                available_capacity = degraded_bess_capacity * max_soc - soc_kwh
                max_charge_energy = min(max_charge_power, available_capacity / charge_eff)
                
                bess_charge_woeff = max_charge_energy  # Energy from renewables
                bess_charge_wieff = bess_charge_woeff * charge_eff  # Stored in BESS
                soc_kwh += bess_charge_wieff
                
                year_bess_charge += bess_charge_woeff
                
                # Curtailment
                curtailment = total_excess - bess_charge_woeff
                year_curtailment += curtailment
            else:
                curtailment = total_excess
                year_curtailment += curtailment
            
            # 5. BESS DISCHARGING (if remaining load)
            bess_discharge_woeff = 0
            bess_discharge_wieff = 0
            unmet = 0
            
            if remaining_load > 0 and degraded_bess_capacity > 0:
                available_energy = soc_kwh - degraded_bess_capacity * min_soc
                max_discharge_power = min(remaining_load, degraded_bess_power)
                max_discharge_energy_wieff = min(max_discharge_power, available_energy * discharge_eff)
                
                bess_discharge_wieff = max_discharge_energy_wieff  # Energy to load
                bess_discharge_woeff = bess_discharge_wieff / discharge_eff  # From BESS storage
                soc_kwh -= bess_discharge_woeff
                
                year_bess_discharge += bess_discharge_wieff
                
                # Unmet load
                unmet = remaining_load - bess_discharge_wieff
                year_unmet += unmet
            else:
                unmet = remaining_load
                year_unmet += unmet
            
            # Store hourly data if this year is selected for export
            if store_hourly:
                # Match Anaconda column order exactly
                hourly_data.append({
                    'Hour': h % 24,  # Reset to 0-23 for each day
                    'Load_kW': load_h,
                    'PV_Available_kW': pv_available,
                    'PV_to_Load_kW': pv_to_load,
                    'PV_Excess_kW': pv_excess,
                    'BESS_Charge_woeff_kW': bess_charge_woeff,
                    'BESS_Charge_wieff_kW': bess_charge_wieff,
                    'BESS_Discharge_woeff_kW': bess_discharge_woeff,
                    'BESS_Discharge_wieff_kW': bess_discharge_wieff,
                    'BESS_SOC_kWh': soc_kwh,
                    'BESS_SOC_pct': (soc_kwh / degraded_bess_capacity * 100) if degraded_bess_capacity > 0 else 0,
                    'Curtailment_kW': curtailment,
                    'Unmet_Load_kW': unmet
                })
        
        # Store hourly results for this year
        if store_hourly:
            hourly_results_by_year[f'year_{year}'] = pd.DataFrame(hourly_data)
            print(f"Hourly data stored", end='', flush=True)
        
        # Store final SOC percentage for carryover to next year
        if degraded_bess_capacity > 0:
            previous_soc_pct = soc_kwh / degraded_bess_capacity
        else:
            previous_soc_pct = 0.5
        
        # Calculate yearly metrics
        unmet_pct = (year_unmet / year_load * 100) if year_load > 0 else 0
        
        yearly_results.append({
            'Year': year,
            'PV_Capacity_kW': degraded_pv_capacity,
            'PV_Degradation_%': pv_deg_pct,
            'BESS_Capacity_kWh': degraded_bess_capacity,
            'BESS_Power_kW': degraded_bess_power,
            'BESS_Retention_%': capacity_retention_pct,
            'Charge_Efficiency_%': charge_eff * 100,
            'Discharge_Efficiency_%': discharge_eff * 100,
            'PV_Generation_kWh': year_pv_gen,
            'Wind_Generation_kWh': year_wind_gen,
            'Hydro_Generation_kWh': year_hydro_gen,
            'Total_Generation_kWh': year_pv_gen + year_wind_gen + year_hydro_gen,
            'BESS_Charge_kWh': year_bess_charge,
            'BESS_Discharge_kWh': year_bess_discharge,
            'Load_kWh': year_load,
            'Unmet_Load_kWh': year_unmet,
            'Unmet_%': unmet_pct,
            'Curtailment_kWh': year_curtailment,
            'Replaced': '🔋 REPLACED' if replaced else ''
        })
        
        print(" ✓")
    
    # Calculate BESS replacement cost
    if apply_bess and bess_capacity_kwh > 0:
        # Replacement cost = 80% of original (no installation cost)
        replacement_cost_nominal = bess_npc * 0.8
        
        # Discount to present value (year 0)
        replacement_cost_pv = replacement_cost_nominal / ((1 + discount_rate) ** 20)
    else:
        replacement_cost_pv = 0
    
    # Total NPC with degradation
    npc_25y = npc_y1 + replacement_cost_pv
    
    # LCOE adjustment (rough estimate)
    if apply_pv or apply_bess:
        # Energy production decreases due to degradation
        # Rough estimate: 5-7% increase in LCOE
        lcoe_25y = lcoe_y1 * 1.06
    else:
        lcoe_25y = lcoe_y1
    
    print("\n" + "="*80)
    print("DEGRADATION ANALYSIS COMPLETE")
    print("="*80)
    print(f"\nNPC Summary:")
    print(f"  Year 1 NPC: ${npc_y1:,.2f}")
    print(f"  BESS Replacement Cost (PV): ${replacement_cost_pv:,.2f}")
    print(f"  25-Year NPC: ${npc_25y:,.2f}")
    print(f"\nLCOE Summary:")
    print(f"  Year 1 LCOE: ${lcoe_y1:.2f}/MWh")
    print(f"  25-Year LCOE: ${lcoe_25y:.2f}/MWh")
    print(f"\nHourly Data Exported for Years: {years_to_export_list}")
    print("="*80 + "\n")
    
    return {
        'yearly_summary': pd.DataFrame(yearly_results),
        'hourly_dispatch': hourly_results_by_year,
        'npc_year1': npc_y1,
        'npc_25year': npc_25y,
        'replacement_cost_pv': replacement_cost_pv,
        'lcoe_year1': lcoe_y1,
        'lcoe_25year': lcoe_25y,
        'pv_deg_total': PV_DEG[25] if apply_pv else 0,
        'bess_loss_20y': (100 - BESS_CAP_RET[20]) if apply_bess else 0,
        'degradation_applied': {
            'pv': apply_pv,
            'bess': apply_bess
        },
        'years_exported': years_to_export_list
    }


# ==============================================================================
# BACKWARD COMPATIBILITY - Old function name
# ==============================================================================

def run_degradation_analysis(*args, **kwargs):
    """
    Backward compatibility wrapper.
    Redirects to the complete analysis function.
    """
    return run_degradation_analysis_complete(*args, **kwargs)
