"""
GRID SEARCH OPTIMIZER - PV + WIND + HYDRO + BESS (CLEAN ARCHITECTURE)
=====================================================================

REFACTORED VERSION V4.0:
- Standardized hourly output format (matches degradation format)
- Energy balance validation columns included
- Direct Python data structures (no Excel intermediary needed)
- HOMER Pro NPC calculation methodology
- Ready for Streamlit direct integration

Features:
- PV + Wind + Hydro + BESS hybrid optimization
- HOMER-style NPC calculation with real discount rate
- 4D grid search (PV × Wind × Hydro × BESS)
- Automatic hydro operating window optimization
- Energy balance tracking for validation
"""

import pandas as pd
import numpy as np
from datetime import datetime
import time

# ==============================================================================
# HOMER NPC CALCULATION FUNCTIONS
# ==============================================================================

def calculate_real_discount_rate(nominal_rate, inflation_rate):
    """
    Calculate real discount rate adjusted for inflation (HOMER method).
    
    Formula: i_real = (i_nom - i_inf) / (1 + i_inf)
    """
    return (nominal_rate - inflation_rate) / (1 + inflation_rate)


def calculate_crf(discount_rate, lifetime_years):
    """
    Calculate Capital Recovery Factor (CRF) - HOMER method.
    
    Formula: CRF = [i(1+i)^n] / [(1+i)^n - 1]
    """
    if discount_rate == 0:
        return 1.0 / lifetime_years
    
    numerator = discount_rate * (1 + discount_rate) ** lifetime_years
    denominator = (1 + discount_rate) ** lifetime_years - 1
    
    return numerator / denominator


def calculate_present_value_factor(discount_rate, lifetime_years):
    """Calculate Present Value Factor (inverse of CRF)."""
    if discount_rate == 0:
        return lifetime_years
    
    return ((1 + discount_rate) ** lifetime_years - 1) / \
           (discount_rate * (1 + discount_rate) ** lifetime_years)


def calculate_salvage_value(component_capital, component_lifetime, project_lifetime, age_at_start=0):
    """
    Calculate salvage value of component at end of project (HOMER method).
    
    HOMER assumes linear depreciation:
    Salvage = Capital × (Remaining_Life / Component_Lifetime)
    """
    age_at_end = (age_at_start + project_lifetime) % component_lifetime
    
    if age_at_end == 0:
        return 0.0
    
    remaining_life = component_lifetime - age_at_end
    salvage_fraction = remaining_life / component_lifetime
    
    return component_capital * salvage_fraction


def calculate_replacement_cost_pv(component_capital, component_lifetime, project_lifetime, 
                                  discount_rate, replacement_cost_multiplier=0.8):
    """
    Calculate present value of all replacement costs over project lifetime (HOMER method).
    
    HOMER assumes:
    - Replacements occur at end of each component lifetime
    - Replacement cost = 80% of original capital (typical assumption)
    - Each replacement is discounted to present value
    """
    if component_lifetime >= project_lifetime:
        return 0.0
    
    replacement_cost = component_capital * replacement_cost_multiplier
    total_replacement_pv = 0.0
    
    replacement_year = component_lifetime
    while replacement_year < project_lifetime:
        pv_of_replacement = replacement_cost / ((1 + discount_rate) ** replacement_year)
        total_replacement_pv += pv_of_replacement
        replacement_year += component_lifetime
    
    return total_replacement_pv


def calculate_component_npc_homer(capital_cost, annual_om_cost, component_lifetime, 
                                 project_lifetime, discount_rate, replacement_cost_multiplier=0.8):
    """
    Calculate NPC for a single component using HOMER methodology.
    
    Formula:
    Annualized = CRF × (Capital + Replacement_PV - Salvage_PV) + Annual_O&M
    NPC = Annualized / CRF
    """
    crf = calculate_crf(discount_rate, project_lifetime)
    
    # Replacement costs
    replacement_pv = calculate_replacement_cost_pv(
        capital_cost, component_lifetime, project_lifetime, 
        discount_rate, replacement_cost_multiplier
    )
    
    # Salvage value
    salvage_value = calculate_salvage_value(
        capital_cost, component_lifetime, project_lifetime
    )
    salvage_pv = salvage_value / ((1 + discount_rate) ** project_lifetime)
    
    # O&M present value
    pv_factor = calculate_present_value_factor(discount_rate, project_lifetime)
    om_pv = annual_om_cost * pv_factor
    
    # Total NPC
    npc = capital_cost + replacement_pv + om_pv - salvage_pv
    
    # Annualized cost
    annualized_cost = npc * crf
    
    return {
        'capital': capital_cost,
        'replacement_pv': replacement_pv,
        'om_pv': om_pv,
        'salvage_pv': salvage_pv,
        'npc': npc,
        'annualized': annualized_cost,
        'crf': crf
    }


# ==============================================================================
# DISPATCH SIMULATION WITH ENERGY BALANCE TRACKING
# ==============================================================================

def calculate_dispatch_with_hydro(load_profile, pvsyst_profile, wind_profile,
                                  pv_capacity, wind_capacity, hydro_capacity,
                                  bess_power, bess_capacity,
                                  solar_config, wind_config, hydro_config, bess_config,
                                  hydro_window_start, hydro_window_end):
    """
    Calculate hourly dispatch for PV + Wind + Hydro + BESS system.
    
    REFACTORED VERSION WITH ENERGY BALANCE TRACKING:
    - Standardized output format (matches degradation format)
    - Tracks efficiency losses explicitly
    - Provides energy balance validation columns
    
    Merit Order Dispatch:
    1. PV (non-dispatchable)
    2. Wind (non-dispatchable)
    3. Hydro (dispatchable, only during operating window)
    4. BESS Discharge
    5. Unmet Load
    """
    
    hours = len(load_profile)
    pv_baseline = solar_config['baseline_kw']
    wind_baseline = 1.0
    charge_eff = bess_config['charge_eff']
    discharge_eff = bess_config['discharge_eff']
    min_soc = bess_config['min_soc']
    max_soc = bess_config['max_soc']
    
    # Calculate SOC limits based on capacity
    min_soc_kwh = min_soc * bess_capacity
    max_soc_kwh = max_soc * bess_capacity
    
    # Initialize arrays for STANDARDIZED FORMAT
    pv_available = np.zeros(hours)
    pv_to_load = np.zeros(hours)
    pv_excess = np.zeros(hours)
    wind_output = np.zeros(hours)
    hydro_output = np.zeros(hours)
    hydro_active_flag = np.zeros(hours)
    
    # BESS tracking with efficiency breakdown
    bess_charge_woeff = np.zeros(hours)  # Before efficiency (energy input)
    bess_charge_wieff = np.zeros(hours)  # After efficiency (actually stored)
    bess_discharge_woeff = np.zeros(hours)  # Before efficiency (energy from battery)
    bess_discharge_wieff = np.zeros(hours)  # After efficiency (usable output)
    bess_soc_kwh = np.zeros(hours)
    bess_soc_pct = np.zeros(hours)
    
    # Load satisfaction
    curtailment = np.zeros(hours)
    unmet_load = np.zeros(hours)
    
    # Start with 50% SOC
    current_soc = 0.5 * bess_capacity
    
    for h in range(hours):
        hour_of_day = h % 24
        
        # Check if hydro is in operating window
        hydro_active = (hour_of_day >= hydro_window_start) and (hour_of_day < hydro_window_end)
        hydro_active_flag[h] = 1 if hydro_active else 0
        
        load = load_profile[h]
        pv_avail = pvsyst_profile[h] * (pv_capacity / pv_baseline)
        wind_avail = wind_profile[h] * (wind_capacity / wind_baseline) if wind_config.get('enabled', False) else 0
        
        # Store available renewable generation
        pv_available[h] = pv_avail
        wind_output[h] = wind_avail
        
        # Calculate how much renewable goes to load
        total_renewable = pv_avail + wind_avail
        renewable_to_load = min(total_renewable, load)
        
        # Allocate proportionally
        if total_renewable > 0:
            pv_to_load[h] = renewable_to_load * (pv_avail / total_renewable)
        else:
            pv_to_load[h] = 0
        
        pv_excess[h] = pv_avail - pv_to_load[h]
        
        # Net load after renewable
        net_load = load - total_renewable
        
        # Dispatch hydro if active and needed
        if net_load > 0 and hydro_active and hydro_capacity > 0:
            hydro_output[h] = min(net_load, hydro_capacity)
            net_load -= hydro_output[h]
        
        if net_load > 0:
            # Need to discharge battery
            available_discharge = max(0, current_soc - min_soc_kwh)
            
            # Energy available after discharge efficiency
            actual_discharge_output = available_discharge * discharge_eff
            
            # How much we actually discharge (limited by power and need)
            discharge_output = min(bess_power, actual_discharge_output, net_load)
            
            if discharge_output > 0:
                # Energy drawn from battery (before efficiency loss)
                energy_from_battery = discharge_output / discharge_eff
                
                bess_discharge_woeff[h] = energy_from_battery
                bess_discharge_wieff[h] = discharge_output
                
                current_soc -= energy_from_battery
            
            unmet_load[h] = max(0, net_load - discharge_output)
            
        else:
            # Excess renewable - charge battery
            excess = abs(net_load)
            available_space = max(0, max_soc_kwh - current_soc)
            
            # How much can we charge (limited by power, space, and excess)
            # Energy that would be stored (after efficiency)
            max_charge_stored = min(bess_power, available_space, excess * charge_eff)
            
            # Energy input needed (before efficiency)
            charge_input = max_charge_stored / charge_eff
            
            bess_charge_woeff[h] = charge_input
            bess_charge_wieff[h] = max_charge_stored
            
            current_soc += max_charge_stored
            
            # Remaining excess becomes curtailment
            curtailment[h] = max(0, excess - charge_input)
        
        # Enforce SOC limits (safety)
        current_soc = max(min_soc_kwh, min(current_soc, max_soc_kwh))
        
        # Store SOC
        bess_soc_kwh[h] = current_soc
        bess_soc_pct[h] = (current_soc / bess_capacity * 100) if bess_capacity > 0 else 0
    
    # Create STANDARDIZED DataFrame (matches degradation format)
    results = pd.DataFrame({
        'Hour': list(range(hours)),
        'Hour_of_Day': [h % 24 for h in range(hours)],
        'Load_kW': load_profile,
        
        # PV breakdown
        'PV_Available_kW': pv_available,
        'PV_to_Load_kW': pv_to_load,
        'PV_Excess_kW': pv_excess,
        
        # Wind, Hydro
        'Wind_Output_kW': wind_output,
        'Hydro_Output_kW': hydro_output,
        'Hydro_Active': hydro_active_flag,
        
        # BESS detailed breakdown (STANDARDIZED - ENERGY BALANCE READY)
        'BESS_Charge_woeff_kW': bess_charge_woeff,
        'BESS_Charge_wieff_kW': bess_charge_wieff,
        'BESS_Discharge_woeff_kW': bess_discharge_woeff,
        'BESS_Discharge_wieff_kW': bess_discharge_wieff,
        
        # BESS state
        'BESS_SOC_kWh': bess_soc_kwh,
        'BESS_SOC_pct': bess_soc_pct,
        
        # Load satisfaction
        'Curtailment_kW': curtailment,
        'Unmet_Load_kW': unmet_load,
    })
    
    # ADD ENERGY BALANCE VALIDATION COLUMNS
    results['Total_Generation_kW'] = (
        results['PV_Available_kW'] + 
        results['Wind_Output_kW'] + 
        results['Hydro_Output_kW']
    )
    
    results['BESS_Charge_Loss_kW'] = (
        results['BESS_Charge_woeff_kW'] - 
        results['BESS_Charge_wieff_kW']
    )
    
    results['BESS_Discharge_Loss_kW'] = (
        results['BESS_Discharge_woeff_kW'] - 
        results['BESS_Discharge_wieff_kW']
    )
    
    results['Total_Losses_kW'] = (
        results['BESS_Charge_Loss_kW'] + 
        results['BESS_Discharge_Loss_kW']
    )
    
    # Energy balance check
    # LHS: Total generation
    results['Energy_Balance_LHS_kW'] = (
        results['PV_Available_kW'] + 
        results['Wind_Output_kW'] + 
        results['Hydro_Output_kW'] +
        results['BESS_Discharge_wieff_kW']  
    )
    
    # RHS: Where energy went
    results['Energy_Balance_RHS_kW'] = (
        (results['Load_kW'] - results['Unmet_Load_kW']) +  # ✅ Served load only
        results['BESS_Charge_woeff_kW'] +
        results['Curtailment_kW']   
    )
    
    # Error (should be ~0)
    results['Energy_Balance_Error_kW'] = (
        results['Energy_Balance_LHS_kW'] - 
        results['Energy_Balance_RHS_kW']
    )
    
    return results


# ==============================================================================
# HYDRO WINDOW OPTIMIZATION
# ==============================================================================

def find_optimal_hydro_window(load_profile, pvsyst_profile, wind_profile,
                              pv_capacity, wind_capacity, hydro_capacity,
                              bess_power, bess_capacity,
                              solar_config, wind_config, hydro_config, bess_config,
                              return_all_windows=False):
    """
    Find optimal hydro operating window by testing all possible N-hour consecutive windows.
    
    Args:
        return_all_windows: If True, returns all window results for analysis
    
    Returns:
        If return_all_windows=False: (best_start, best_end, best_unmet_percent)
        If return_all_windows=True: (best_start, best_end, best_unmet_percent, all_windows_df)
    """
    
    hours_per_day = hydro_config['hours_per_day']
    
    # Test all possible windows
    window_results = []
    
    for start_hour in range(24 - hours_per_day + 1):
        end_hour = start_hour + hours_per_day
        
        # Run dispatch with this window
        dispatch = calculate_dispatch_with_hydro(
            load_profile, pvsyst_profile, wind_profile,
            pv_capacity, wind_capacity, hydro_capacity,
            bess_power, bess_capacity,
            solar_config, wind_config, hydro_config, bess_config,
            start_hour, end_hour
        )
        
        # Calculate unmet load percentage
        total_load = dispatch['Load_kW'].sum()
        total_unmet = dispatch['Unmet_Load_kW'].sum()
        unmet_percent = (total_unmet / total_load * 100) if total_load > 0 else 0
        
        window_results.append({
            'start_hour': start_hour,
            'end_hour': end_hour,
            'unmet_percent': unmet_percent
        })
    
    # Find window with lowest unmet load
    best_window = min(window_results, key=lambda x: x['unmet_percent'])
    
    if return_all_windows:
        # Create DataFrame for analysis
        windows_df = pd.DataFrame(window_results)
        windows_df['rank'] = windows_df['unmet_percent'].rank()
        windows_df['window_range'] = windows_df.apply(
            lambda x: f"{int(x['start_hour']):02d}:00-{int(x['end_hour']):02d}:00", axis=1
        )
        windows_df = windows_df.sort_values('unmet_percent')
        
        return (best_window['start_hour'], best_window['end_hour'], 
                best_window['unmet_percent'], windows_df)
    else:
        return (best_window['start_hour'], best_window['end_hour'], 
                best_window['unmet_percent'])


# ==============================================================================
# HOMER-STYLE NPC CALCULATION
# ==============================================================================

def calculate_npc_homer_style(pv_capacity, wind_capacity, hydro_capacity, bess_power, bess_capacity,
                              solar_config, wind_config, hydro_config, bess_config, project_config,
                              lcoe_tables, use_dynamic, total_energy_served_annual):
    """
    Calculate system NPC using HOMER Pro methodology including Hydro.
    
    Args:
        total_energy_served_annual (float): Annual energy served to load (kWh/year)
    
    Returns detailed cost breakdown matching HOMER format.
    """
    
    nominal_discount_rate = project_config.get('discount_rate', 0.08)
    inflation_rate = project_config.get('inflation_rate', 0.02)
    project_lifetime = project_config.get('project_lifetime', 25)
    
    # Calculate real discount rate (HOMER style)
    real_discount_rate = calculate_real_discount_rate(nominal_discount_rate, inflation_rate)
    
    # Component lifetimes
    pv_lifetime = solar_config.get('lifetime', 25)
    wind_lifetime = wind_config.get('lifetime', 25)
    hydro_lifetime = hydro_config.get('lifetime', 25)
    bess_lifetime = bess_config.get('lifetime', 10)
    
    # ==== PV Component ====
    if pv_capacity > 0:
        pv_capital = pv_capacity * solar_config['capex_per_kw']
        pv_om_annual = pv_capacity * solar_config['om_per_kw_year']
        pv_npc_data = calculate_component_npc_homer(
            pv_capital, pv_om_annual, pv_lifetime, 
            project_lifetime, real_discount_rate
        )
    else:
        pv_npc_data = {
            'capital': 0, 'replacement_pv': 0, 'om_pv': 0, 
            'salvage_pv': 0, 'npc': 0, 'annualized': 0, 'crf': 0
        }
    
    # ==== Wind Component ====
    if wind_capacity > 0:
        wind_capital = wind_capacity * wind_config['capex_per_kw']
        wind_om_annual = wind_capacity * wind_config['om_per_kw_year']
        wind_npc_data = calculate_component_npc_homer(
            wind_capital, wind_om_annual, wind_lifetime, 
            project_lifetime, real_discount_rate
        )
    else:
        wind_npc_data = {
            'capital': 0, 'replacement_pv': 0, 'om_pv': 0, 
            'salvage_pv': 0, 'npc': 0, 'annualized': 0, 'crf': 0
        }
    
    # ==== Hydro Component ====
    if hydro_capacity > 0:
        hydro_capital = hydro_capacity * hydro_config['capex_per_kw']
        hydro_om_annual = hydro_capacity * hydro_config['om_per_kw_year']
        hydro_npc_data = calculate_component_npc_homer(
            hydro_capital, hydro_om_annual, hydro_lifetime, 
            project_lifetime, real_discount_rate
        )
    else:
        hydro_npc_data = {
            'capital': 0, 'replacement_pv': 0, 'om_pv': 0, 
            'salvage_pv': 0, 'npc': 0, 'annualized': 0, 'crf': 0
        }
    
    # ==== BESS Component ====
    if bess_power > 0:
        bess_capital = (bess_power * bess_config['power_capex_per_kw'] + 
                       bess_capacity * bess_config['energy_capex_per_kwh'])
        bess_om_annual = bess_power * bess_config['om_per_kw_year']
        bess_npc_data = calculate_component_npc_homer(
            bess_capital, bess_om_annual, bess_lifetime, 
            project_lifetime, real_discount_rate,
            replacement_cost_multiplier=0.8
        )
    else:
        bess_npc_data = {
            'capital': 0, 'replacement_pv': 0, 'om_pv': 0, 
            'salvage_pv': 0, 'npc': 0, 'annualized': 0, 'crf': 0
        }
    
    # ==== System Totals ====
    total_capital = (pv_npc_data['capital'] + wind_npc_data['capital'] + 
                    hydro_npc_data['capital'] + bess_npc_data['capital'])
    total_replacement_pv = (pv_npc_data['replacement_pv'] + wind_npc_data['replacement_pv'] + 
                           hydro_npc_data['replacement_pv'] + bess_npc_data['replacement_pv'])
    total_om_pv = (pv_npc_data['om_pv'] + wind_npc_data['om_pv'] + 
                  hydro_npc_data['om_pv'] + bess_npc_data['om_pv'])
    total_salvage_pv = (pv_npc_data['salvage_pv'] + wind_npc_data['salvage_pv'] + 
                       hydro_npc_data['salvage_pv'] + bess_npc_data['salvage_pv'])
    total_npc = (pv_npc_data['npc'] + wind_npc_data['npc'] + 
                hydro_npc_data['npc'] + bess_npc_data['npc'])
    total_annualized = (pv_npc_data['annualized'] + wind_npc_data['annualized'] + 
                       hydro_npc_data['annualized'] + bess_npc_data['annualized'])
    
    # Calculate Levelized Cost of Energy (LCOE) - INDUSTRY STANDARD METHOD
    # LCOE = Annualized Cost / Annual Energy Delivered
    
    if total_energy_served_annual > 0:
        lcoe = total_annualized / total_energy_served_annual  # $/kWh per year
    else:
        lcoe = 0
    
    system_crf = calculate_crf(real_discount_rate, project_lifetime)
    
    return {
        # Project parameters
        'nominal_discount_rate': nominal_discount_rate,
        'inflation_rate': inflation_rate,
        'real_discount_rate': real_discount_rate,
        'project_lifetime': project_lifetime,
        'crf': system_crf,
        
        # Component breakdowns
        'pv': pv_npc_data,
        'wind': wind_npc_data,
        'hydro': hydro_npc_data,
        'bess': bess_npc_data,
        
        # System totals
        'total_capital': total_capital,
        'total_replacement_pv': total_replacement_pv,
        'total_om_pv': total_om_pv,
        'total_salvage_pv': total_salvage_pv,
        'total_npc': total_npc,
        'total_annualized': total_annualized,
        
        # Performance metrics
        'total_energy_served_annual': total_energy_served_annual,
        'lcoe': lcoe,
        'lcoe_per_kwh': lcoe,
        'lcoe_per_mwh': lcoe * 1000
    }


# ==============================================================================
# GRID SEARCH OPTIMIZATION
# ==============================================================================

def grid_search_optimize_hydro(config, grid_config, solar, wind, hydro, bess,
                               load_profile, pvsyst_profile, wind_profile, lcoe_tables):
    """4D Grid search optimization with HOMER NPC calculation and hydro window optimization."""
    
    print("\n" + "="*70)
    print("GRID SEARCH OPTIMIZATION (CLEAN ARCHITECTURE)")
    print("="*70)
    
    use_dynamic = config.get('use_dynamic_lcoe', False)
    target_unmet = config['target_unmet_percent']
    
    # Calculate real discount rate
    nominal_rate = config['discount_rate']
    inflation_rate = config.get('inflation_rate', 0.02)
    real_rate = calculate_real_discount_rate(nominal_rate, inflation_rate)
    
    print(f"\nDiscount Rates:")
    print(f"  Nominal: {nominal_rate*100:.2f}%")
    print(f"  Inflation: {inflation_rate*100:.2f}%")
    print(f"  Real (HOMER): {real_rate*100:.4f}%")
    
    # Generate capacity ranges (FIXED: Always include endpoint)
    def generate_range_with_endpoint(start, end, step):
        """Generate range from start to end with given step, always including end."""
        values = []
        current = start
        while current <= end:
            values.append(current)
            current += step
        
        if values[-1] != end:
            values.append(end)
        
        return np.array(values)
    
    pv_range = generate_range_with_endpoint(grid_config['pv_start'], 
                                            grid_config['pv_end'], 
                                            grid_config['pv_step'])
    wind_range = generate_range_with_endpoint(grid_config['wind_start'], 
                                              grid_config['wind_end'], 
                                              grid_config['wind_step'])
    hydro_range = generate_range_with_endpoint(grid_config['hydro_start'], 
                                               grid_config['hydro_end'], 
                                               grid_config['hydro_step'])
    bess_range = generate_range_with_endpoint(grid_config['bess_start'], 
                                              grid_config['bess_end'], 
                                              grid_config['bess_step'])
    
    total_combinations = len(pv_range) * len(wind_range) * len(hydro_range) * len(bess_range)
    
    print(f"\nSearch Space:")
    print(f"  PV: {len(pv_range)} options")
    print(f"  Wind: {len(wind_range)} options")
    print(f"  Hydro: {len(hydro_range)} options")
    print(f"  BESS: {len(bess_range)} options")
    print(f"  Total: {total_combinations} combinations")
    print("="*70)
    
    start_time = time.time()
    results = []
    count = 0
    
    for pv_cap in pv_range:
        for wind_cap in wind_range:
            for hydro_cap in hydro_range:
                for bess_power in bess_range:
                    count += 1
                    
                    if count % 100 == 0:
                        elapsed = time.time() - start_time
                        rate = count / elapsed
                        remaining = (total_combinations - count) / rate
                        print(f"  Progress: {count}/{total_combinations} ({count/total_combinations*100:.1f}%) - ETA: {remaining:.0f}s")
                    
                    bess_capacity = bess_power * bess['duration_hours']
                    
                    # Find optimal hydro window
                    hydro_start, hydro_end, _ = find_optimal_hydro_window(
                        load_profile, pvsyst_profile, wind_profile,
                        pv_cap, wind_cap, hydro_cap,
                        bess_power, bess_capacity,
                        solar, wind, hydro, bess
                    )
                    
                    # Dispatch simulation with optimal window
                    dispatch = calculate_dispatch_with_hydro(
                        load_profile, pvsyst_profile, wind_profile,
                        pv_cap, wind_cap, hydro_cap,
                        bess_power, bess_capacity,
                        solar, wind, hydro, bess,
                        hydro_start, hydro_end
                    )
                    
                    # Calculate metrics
                    total_load = dispatch['Load_kW'].sum()
                    total_unmet = dispatch['Unmet_Load_kW'].sum()
                    unmet_percent = (total_unmet / total_load * 100) if total_load > 0 else 0
                    
                    feasible = unmet_percent <= target_unmet
                    
                    # Energy served (for LCOE calculation)
                    total_energy_served = total_load - total_unmet
                    
                    # HOMER-style NPC calculation
                    npc_data = calculate_npc_homer_style(
                        pv_cap, wind_cap, hydro_cap, bess_power, bess_capacity,
                        solar, wind, hydro, bess, config,
                        lcoe_tables, use_dynamic, total_energy_served
                    )
                    
                    # Energy statistics
                    pv_energy = dispatch['PV_Available_kW'].sum()
                    wind_energy = dispatch['Wind_Output_kW'].sum()
                    hydro_energy = dispatch['Hydro_Output_kW'].sum()
                    bess_discharge_annual = dispatch['BESS_Discharge_wieff_kW'].sum()
                    excess = dispatch['Curtailment_kW'].sum()
                    
                    # Calculate cycles
                    cycles_per_year = bess_discharge_annual / bess_capacity if bess_capacity > 0 else 0
                    
                    # Store result
                    results.append({
                        'Iteration': count,
                        'PV_kW': pv_cap,
                        'Wind_kW': wind_cap,
                        'Hydro_kW': hydro_cap,
                        'Hydro_Window_Start': hydro_start,
                        'Hydro_Window_End': hydro_end,
                        'BESS_Power_kW': bess_power,
                        'BESS_Capacity_kWh': bess_capacity,
                        'BESS_Annual_Discharge_kWh': bess_discharge_annual,
                        'BESS_Cycles_Per_Year': cycles_per_year,
                        'Unmet_%': unmet_percent,
                        'Feasible': feasible,
                        
                        # HOMER NPC breakdown
                        'NPC_$': npc_data['total_npc'],
                        'Capital_$': npc_data['total_capital'],
                        'Replacement_$': npc_data['total_replacement_pv'],
                        'OM_$': npc_data['total_om_pv'],
                        'Salvage_$': npc_data['total_salvage_pv'],
                        'Annualized_$/yr': npc_data['total_annualized'],
                        
                        # Component NPC breakdown
                        'PV_Capital_$': npc_data['pv']['capital'],
                        'PV_Replacement_$': npc_data['pv']['replacement_pv'],
                        'PV_OM_$': npc_data['pv']['om_pv'],
                        'PV_Salvage_$': npc_data['pv']['salvage_pv'],
                        'PV_NPC_$': npc_data['pv']['npc'],
                        'PV_Annualized_$/yr': npc_data['pv']['annualized'],
                        
                        'Wind_Capital_$': npc_data['wind']['capital'],
                        'Wind_Replacement_$': npc_data['wind']['replacement_pv'],
                        'Wind_OM_$': npc_data['wind']['om_pv'],
                        'Wind_Salvage_$': npc_data['wind']['salvage_pv'],
                        'Wind_NPC_$': npc_data['wind']['npc'],
                        'Wind_Annualized_$/yr': npc_data['wind']['annualized'],
                        
                        'Hydro_Capital_$': npc_data['hydro']['capital'],
                        'Hydro_Replacement_$': npc_data['hydro']['replacement_pv'],
                        'Hydro_OM_$': npc_data['hydro']['om_pv'],
                        'Hydro_Salvage_$': npc_data['hydro']['salvage_pv'],
                        'Hydro_NPC_$': npc_data['hydro']['npc'],
                        'Hydro_Annualized_$/yr': npc_data['hydro']['annualized'],
                        
                        'BESS_Capital_$': npc_data['bess']['capital'],
                        'BESS_Replacement_$': npc_data['bess']['replacement_pv'],
                        'BESS_OM_$': npc_data['bess']['om_pv'],
                        'BESS_Salvage_$': npc_data['bess']['salvage_pv'],
                        'BESS_NPC_$': npc_data['bess']['npc'],
                        'BESS_Annualized_$/yr': npc_data['bess']['annualized'],
                        
                        # Performance metrics
                        'LCOE_$/kWh': npc_data['lcoe'],
                        'LCOE_$/MWh': npc_data['lcoe'] * 1000,
                        'Real_Discount_Rate_%': npc_data['real_discount_rate'] * 100,
                        'CRF': npc_data['crf'],
                        
                        'Total_Load_kWh': total_load,
                        'Total_Energy_Served_kWh': total_energy_served,
                        'PV_Energy_kWh': pv_energy,
                        'Wind_Energy_kWh': wind_energy,
                        'Hydro_Energy_kWh': hydro_energy,
                        'Unmet_kWh': total_unmet,
                        'Excess_kWh': excess,
                        
                        # Renewable energy metrics
                        'PV_Fraction_%': (pv_energy / (pv_energy + wind_energy + hydro_energy + 0.001) * 100),
                        'Wind_Fraction_%': (wind_energy / (pv_energy + wind_energy + hydro_energy + 0.001) * 100),
                        'Hydro_Fraction_%': (hydro_energy / (pv_energy + wind_energy + hydro_energy + 0.001) * 100),
                        'RE_Penetration_%': (total_energy_served / total_load * 100) if total_load > 0 else 0,
                        'BESS_Contribution_%': (bess_discharge_annual / total_load * 100) if total_load > 0 else 0,
                        'Excess_Fraction_%': (excess / (pv_energy + wind_energy + hydro_energy + 0.001) * 100) if (pv_energy + wind_energy + hydro_energy) > 0 else 0
                    })
    
    elapsed = time.time() - start_time
    print(f"\n✓ Completed in {elapsed:.1f}s ({elapsed/total_combinations:.3f}s per combination)")
    
    return pd.DataFrame(results)


# ==============================================================================
# ELECTRICAL METRICS CALCULATION
# ==============================================================================

def calculate_electrical_metrics(dispatch_df, component_capacities, component_configs, 
                                component_npc_data, project_lifetime):
    """
    Calculate electrical performance metrics including actual LCOE from NPC.
    
    Args:
        dispatch_df: DataFrame with hourly dispatch results (STANDARDIZED FORMAT)
        component_capacities: dict with 'pv_kw', 'wind_kw', 'hydro_kw', 'bess_kwh'
        component_configs: dict with component parameters
        component_npc_data: dict with NPC breakdown from calculate_npc_homer_style
        project_lifetime: Project lifetime in years
    
    Returns:
        Dictionary with metrics for PV, Wind, Hydro, and BESS
    """
    
    metrics = {}
    
    # PV Metrics (using standardized column names)
    if component_capacities['pv_kw'] > 0:
        pv_output = dispatch_df['PV_Available_kW'].values
        pv_total_production = pv_output.sum()
        pv_hours_operation = (pv_output > 0).sum()
        pv_mean_output = pv_output[pv_output > 0].mean() if pv_hours_operation > 0 else 0
        pv_capacity_factor = (pv_total_production / (component_capacities['pv_kw'] * 8760)) * 100
        
        # Calculate actual LCOE from NPC
        pv_lcoe = calculate_component_lcoe_from_npc(
            component_npc_data['pv']['npc'],
            pv_total_production,
            project_lifetime
        )
        
        metrics['pv'] = {
            'rated_capacity_kw': component_capacities['pv_kw'],
            'mean_output_kw': pv_mean_output,
            'capacity_factor_pct': pv_capacity_factor,
            'total_production_kwh': pv_total_production,
            'hours_of_operation': pv_hours_operation,
            'levelized_cost_per_kwh': pv_lcoe,
            'levelized_cost_per_mwh': pv_lcoe * 1000
        }
    else:
        metrics['pv'] = {
            'rated_capacity_kw': 0, 'mean_output_kw': 0, 'capacity_factor_pct': 0,
            'total_production_kwh': 0, 'hours_of_operation': 0, 
            'levelized_cost_per_kwh': 0, 'levelized_cost_per_mwh': 0
        }
    
    # Wind Metrics
    if component_capacities['wind_kw'] > 0:
        wind_output = dispatch_df['Wind_Output_kW'].values
        wind_total_production = wind_output.sum()
        wind_hours_operation = (wind_output > 0).sum()
        wind_mean_output = wind_output[wind_output > 0].mean() if wind_hours_operation > 0 else 0
        wind_capacity_factor = (wind_total_production / (component_capacities['wind_kw'] * 8760)) * 100
        
        # Calculate actual LCOE from NPC
        wind_lcoe = calculate_component_lcoe_from_npc(
            component_npc_data['wind']['npc'],
            wind_total_production,
            project_lifetime
        )
        
        metrics['wind'] = {
            'rated_capacity_kw': component_capacities['wind_kw'],
            'mean_output_kw': wind_mean_output,
            'capacity_factor_pct': wind_capacity_factor,
            'total_production_kwh': wind_total_production,
            'hours_of_operation': wind_hours_operation,
            'levelized_cost_per_kwh': wind_lcoe,
            'levelized_cost_per_mwh': wind_lcoe * 1000
        }
    else:
        metrics['wind'] = {
            'rated_capacity_kw': 0, 'mean_output_kw': 0, 'capacity_factor_pct': 0,
            'total_production_kwh': 0, 'hours_of_operation': 0,
            'levelized_cost_per_kwh': 0, 'levelized_cost_per_mwh': 0
        }
    
    # Hydro Metrics
    if component_capacities['hydro_kw'] > 0:
        hydro_output = dispatch_df['Hydro_Output_kW'].values
        hydro_total_production = hydro_output.sum()
        hydro_hours_operation = (hydro_output > 0).sum()
        hydro_mean_output = hydro_output[hydro_output > 0].mean() if hydro_hours_operation > 0 else 0
        hydro_capacity_factor = (hydro_total_production / (component_capacities['hydro_kw'] * 8760)) * 100
        
        # Calculate actual LCOE from NPC
        hydro_lcoe = calculate_component_lcoe_from_npc(
            component_npc_data['hydro']['npc'],
            hydro_total_production,
            project_lifetime
        )
        
        metrics['hydro'] = {
            'rated_capacity_kw': component_capacities['hydro_kw'],
            'mean_output_kw': hydro_mean_output,
            'capacity_factor_pct': hydro_capacity_factor,
            'total_production_kwh': hydro_total_production,
            'hours_of_operation': hydro_hours_operation,
            'levelized_cost_per_kwh': hydro_lcoe,
            'levelized_cost_per_mwh': hydro_lcoe * 1000
        }
    else:
        metrics['hydro'] = {
            'rated_capacity_kw': 0, 'mean_output_kw': 0, 'capacity_factor_pct': 0,
            'total_production_kwh': 0, 'hours_of_operation': 0,
            'levelized_cost_per_kwh': 0, 'levelized_cost_per_mwh': 0
        }
    
    # BESS Metrics (using standardized column names)
    if component_capacities['bess_kwh'] > 0:
        bess_charge_woeff = dispatch_df['BESS_Charge_woeff_kW'].values
        bess_discharge_wieff = dispatch_df['BESS_Discharge_wieff_kW'].values
        
        energy_in = bess_charge_woeff.sum()
        energy_out = bess_discharge_wieff.sum()
        
        # Calculate losses from detailed columns
        charge_loss = (dispatch_df['BESS_Charge_woeff_kW'] - dispatch_df['BESS_Charge_wieff_kW']).sum()
        discharge_loss = (dispatch_df['BESS_Discharge_woeff_kW'] - dispatch_df['BESS_Discharge_wieff_kW']).sum()
        losses = charge_loss + discharge_loss
        
        annual_throughput = energy_out
        
        mean_load = dispatch_df['Load_kW'].mean()
        usable_capacity = component_capacities['bess_kwh'] * (
            component_configs['bess_max_soc'] - component_configs['bess_min_soc']
        )
        autonomy_hours = usable_capacity / mean_load if mean_load > 0 else 0
        
        # Calculate actual LCOS from NPC
        bess_lcos = calculate_bess_lcos_from_npc(
            component_npc_data['bess']['npc'],
            annual_throughput,
            project_lifetime
        )
        
        metrics['bess'] = {
            'nominal_capacity_kwh': component_capacities['bess_kwh'],
            'usable_capacity_kwh': usable_capacity,
            'autonomy_hours': autonomy_hours,
            'energy_in_kwh': energy_in,
            'energy_out_kwh': energy_out,
            'losses_kwh': losses,
            'annual_throughput_kwh': annual_throughput,
            'expected_life_years': component_configs['bess_lifetime'],
            'levelized_cost_per_kwh': bess_lcos,
            'levelized_cost_per_mwh': bess_lcos * 1000
        }
    else:
        metrics['bess'] = {
            'nominal_capacity_kwh': 0, 'usable_capacity_kwh': 0, 'autonomy_hours': 0,
            'energy_in_kwh': 0, 'energy_out_kwh': 0, 'losses_kwh': 0,
            'annual_throughput_kwh': 0, 'expected_life_years': 0,
            'levelized_cost_per_kwh': 0, 'levelized_cost_per_mwh': 0
        }
    
    return metrics


def calculate_component_lcoe_from_npc(component_npc, annual_energy_kwh, project_lifetime):
    """
    Calculate component-specific LCOE from NPC and energy production.
    
    Args:
        component_npc: Component Net Present Cost ($)
        annual_energy_kwh: Annual energy production (kWh/year)
        project_lifetime: Project lifetime (years)
    
    Returns:
        LCOE in $/kWh
    """
    if annual_energy_kwh <= 0:
        return 0.0
    
    total_lifetime_energy = annual_energy_kwh * project_lifetime
    lcoe = component_npc / total_lifetime_energy
    
    return lcoe


def calculate_bess_lcos_from_npc(bess_npc, annual_throughput_kwh, project_lifetime):
    """
    Calculate BESS Levelized Cost of Storage (LCOS) from NPC and throughput.
    
    Args:
        bess_npc: BESS Net Present Cost ($)
        annual_throughput_kwh: Annual energy throughput (kWh/year discharged)
        project_lifetime: Project lifetime (years)
    
    Returns:
        LCOS in $/kWh
    """
    if annual_throughput_kwh <= 0:
        return 0.0
    
    total_lifetime_throughput = annual_throughput_kwh * project_lifetime
    lcos = bess_npc / total_lifetime_throughput
    
    return lcos


# ==============================================================================
# FIND OPTIMAL SOLUTION
# ==============================================================================

def find_optimal_solution(results_df):
    """Find optimal solution from results."""
    
    feasible = results_df[results_df['Feasible'] == True]
    
    if len(feasible) == 0:
        print("\n⚠️  WARNING: No feasible solutions found!")
        return None
    
    optimal = feasible.loc[feasible['NPC_$'].idxmin()]
    
    print("\n" + "="*70)
    print("✓ OPTIMAL SOLUTION FOUND (CLEAN ARCHITECTURE)!")
    print("="*70)
    print(f"  Iteration:  #{int(optimal['Iteration'])} of {len(results_df)}")
    print(f"  PV:         {optimal['PV_kW']:.0f} kW")
    print(f"  Wind:       {optimal['Wind_kW']:.0f} kW")
    print(f"  Hydro:      {optimal['Hydro_kW']:.0f} kW (Window: {int(optimal['Hydro_Window_Start']):02d}:00-{int(optimal['Hydro_Window_End']):02d}:00)")
    print(f"  BESS Power: {optimal['BESS_Power_kW']:.0f} kW")
    print(f"  BESS Capacity: {optimal['BESS_Capacity_kWh']:.0f} kWh")
    print(f"\n  === HOMER-STYLE COSTS ===")
    print(f"  Total NPC:     ${optimal['NPC_$']:,.2f}")
    print(f"  Capital:       ${optimal['Capital_$']:,.2f}")
    print(f"  Replacement:   ${optimal['Replacement_$']:,.2f}")
    print(f"  O&M:           ${optimal['OM_$']:,.2f}")
    print(f"  Salvage:       ${optimal['Salvage_$']:,.2f}")
    print(f"  Annualized:    ${optimal['Annualized_$/yr']:,.2f}/year")
    print(f"\n  === PERFORMANCE ===")
    print(f"  LCOE:       ${optimal['LCOE_$/kWh']:.4f}/kWh (${optimal['LCOE_$/MWh']:.2f}/MWh)")
    print(f"  Unmet Load: {optimal['Unmet_%']:.2f}%")
    print(f"  Real Rate:  {optimal['Real_Discount_Rate_%']:.4f}%")
    print(f"  CRF:        {optimal['CRF']:.6f}")
    print("="*70)
    
    return optimal


# ==============================================================================
# MODULE ENTRY POINT (FOR TESTING)
# ==============================================================================

if __name__ == "__main__":
    print("\n" + "="*70)
    print("OPTIMIZATION MODULE - CLEAN ARCHITECTURE v4.0")
    print("Ready for direct Python integration with Streamlit")
    print("="*70)
    print("\nThis module is designed to be imported by Streamlit.")
    print("It does NOT require Excel files as intermediary.")
    print("\nKey features:")
    print("  ✅ Direct Python data structures")
    print("  ✅ Standardized hourly output format")
    print("  ✅ Energy balance validation columns")
    print("  ✅ HOMER Pro NPC methodology")
    print("="*70 + "\n")
