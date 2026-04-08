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

FIXES:
- Wind output bug fixed: wind_capacity > 0 check replaces missing 'enabled' flag
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
    """
    crf = calculate_crf(discount_rate, project_lifetime)
    replacement_pv = calculate_replacement_cost_pv(
        capital_cost, component_lifetime, project_lifetime,
        discount_rate, replacement_cost_multiplier
    )
    salvage_value = calculate_salvage_value(capital_cost, component_lifetime, project_lifetime)
    salvage_pv = salvage_value / ((1 + discount_rate) ** project_lifetime)
    pv_factor = calculate_present_value_factor(discount_rate, project_lifetime)
    om_pv = annual_om_cost * pv_factor
    npc = capital_cost + replacement_pv + om_pv - salvage_pv
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
                                  hydro_window_start, hydro_window_end,
                                  initial_soc_kwh=None):
    """
    Calculate hourly dispatch for PV + Wind + Hydro + BESS system.

    REFACTORED VERSION WITH ENERGY BALANCE TRACKING:
    - Standardized output format (matches degradation format)
    - Tracks efficiency losses explicitly
    - Provides energy balance validation columns

    FIX: Wind output now uses `wind_capacity > 0` check instead of
         `wind_config.get('enabled', False)` which was never set and caused
         wind_output to be zero even when wind was selected.

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
    min_soc_kwh = (min_soc / 100) * bess_capacity
    max_soc_kwh = (max_soc / 100) * bess_capacity

    # Initialize arrays for STANDARDIZED FORMAT
    pv_available = np.zeros(hours)
    pv_to_load = np.zeros(hours)
    pv_excess = np.zeros(hours)
    wind_output = np.zeros(hours)
    hydro_output = np.zeros(hours)
    hydro_active_flag = np.zeros(hours)

    # BESS tracking with efficiency breakdown
    bess_charge_woeff = np.zeros(hours)
    bess_charge_wieff = np.zeros(hours)
    bess_discharge_woeff = np.zeros(hours)
    bess_discharge_wieff = np.zeros(hours)
    bess_soc_kwh = np.zeros(hours)
    bess_soc_pct = np.zeros(hours)

    # Load satisfaction
    curtailment = np.zeros(hours)
    unmet_load = np.zeros(hours)

    # Start with provided SOC or default to 50%
    if initial_soc_kwh is not None:
        current_soc = float(initial_soc_kwh)
    else:
        current_soc = 0.5 * bess_capacity

    for h in range(hours):
        hour_of_day = h % 24

        # Check if hydro is in operating window
        hydro_active = (hour_of_day >= hydro_window_start) and (hour_of_day < hydro_window_end)
        hydro_active_flag[h] = 1 if hydro_active else 0

        load = load_profile[h]
        pv_avail = pvsyst_profile[h] * (pv_capacity / pv_baseline) if pv_capacity > 0 else 0

        # ── FIX: Use wind_capacity > 0 instead of wind_config.get('enabled', False) ──
        wind_avail = wind_profile[h] * (wind_capacity / wind_baseline) if wind_capacity > 0 else 0

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
            actual_discharge_output = available_discharge * discharge_eff
            discharge_output = min(bess_power, actual_discharge_output, net_load)

            if discharge_output > 0:
                energy_from_battery = discharge_output / discharge_eff
                bess_discharge_woeff[h] = energy_from_battery
                bess_discharge_wieff[h] = discharge_output
                current_soc -= energy_from_battery

            unmet_load[h] = max(0, net_load - discharge_output)

        else:
            # Excess renewable - charge battery
            excess = abs(net_load)
            available_space = max(0, max_soc_kwh - current_soc)
            max_charge_stored = min(bess_power, available_space, excess * charge_eff)
            charge_input = max_charge_stored / charge_eff if max_charge_stored > 0 else 0

            bess_charge_woeff[h] = charge_input
            bess_charge_wieff[h] = max_charge_stored

            current_soc += max_charge_stored
            curtailment[h] = max(0, excess - charge_input)

        # Enforce SOC limits (safety)
        current_soc = max(min_soc_kwh, min(current_soc, max_soc_kwh))

        # Store SOC
        bess_soc_kwh[h] = current_soc
        bess_soc_pct[h] = (current_soc / bess_capacity * 100) if bess_capacity > 0 else 0

    # Create STANDARDIZED DataFrame
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

        # BESS detailed breakdown
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

    results['Energy_Balance_LHS_kW'] = (
        results['PV_Available_kW'] +
        results['Wind_Output_kW'] +
        results['Hydro_Output_kW'] +
        results['BESS_Discharge_wieff_kW']
    )

    results['Energy_Balance_RHS_kW'] = (
        (results['Load_kW'] - results['Unmet_Load_kW']) +
        results['BESS_Charge_woeff_kW'] +
        results['Curtailment_kW']
    )

    results['Energy_Balance_Error_kW'] = (
        results['Energy_Balance_LHS_kW'] -
        results['Energy_Balance_RHS_kW']
    )

    results._final_soc_kwh = current_soc

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
    """
    hours_per_day = hydro_config['hours_per_day']
    window_results = []

    for start_hour in range(24 - hours_per_day + 1):
        end_hour = start_hour + hours_per_day

        dispatch = calculate_dispatch_with_hydro(
            load_profile, pvsyst_profile, wind_profile,
            pv_capacity, wind_capacity, hydro_capacity,
            bess_power, bess_capacity,
            solar_config, wind_config, hydro_config, bess_config,
            start_hour, end_hour
        )

        total_load = dispatch['Load_kW'].sum()
        total_unmet = dispatch['Unmet_Load_kW'].sum()
        unmet_percent = (total_unmet / total_load * 100) if total_load > 0 else 0

        window_results.append({
            'start_hour': start_hour,
            'end_hour': end_hour,
            'unmet_percent': unmet_percent
        })

    best_window = min(window_results, key=lambda x: x['unmet_percent'])

    if return_all_windows:
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
    """
    nominal_discount_rate = project_config.get('discount_rate', 0.08)
    inflation_rate = project_config.get('inflation_rate', 0.02)
    project_lifetime = project_config.get('project_lifetime', 25)

    real_discount_rate = calculate_real_discount_rate(nominal_discount_rate, inflation_rate)

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
        pv_npc_data = {'capital': 0, 'replacement_pv': 0, 'om_pv': 0,
                       'salvage_pv': 0, 'npc': 0, 'annualized': 0, 'crf': 0}

    # ==== Wind Component ====
    if wind_capacity > 0:
        wind_capital = wind_capacity * wind_config['capex_per_kw']
        wind_om_annual = wind_capacity * wind_config['om_per_kw_year']
        wind_npc_data = calculate_component_npc_homer(
            wind_capital, wind_om_annual, wind_lifetime,
            project_lifetime, real_discount_rate
        )
    else:
        wind_npc_data = {'capital': 0, 'replacement_pv': 0, 'om_pv': 0,
                         'salvage_pv': 0, 'npc': 0, 'annualized': 0, 'crf': 0}

    # ==== Hydro Component ====
    if hydro_capacity > 0:
        hydro_capital = hydro_capacity * hydro_config['capex_per_kw']
        hydro_om_annual = hydro_capacity * hydro_config['om_per_kw_year']
        hydro_npc_data = calculate_component_npc_homer(
            hydro_capital, hydro_om_annual, hydro_lifetime,
            project_lifetime, real_discount_rate
        )
    else:
        hydro_npc_data = {'capital': 0, 'replacement_pv': 0, 'om_pv': 0,
                          'salvage_pv': 0, 'npc': 0, 'annualized': 0, 'crf': 0}

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
        bess_npc_data = {'capital': 0, 'replacement_pv': 0, 'om_pv': 0,
                         'salvage_pv': 0, 'npc': 0, 'annualized': 0, 'crf': 0}

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

    # LCOE = Annualized Cost / Annual Energy Delivered
    lcoe = total_annualized / total_energy_served_annual if total_energy_served_annual > 0 else 0
    system_crf = calculate_crf(real_discount_rate, project_lifetime)

    return {
        'nominal_discount_rate': nominal_discount_rate,
        'inflation_rate': inflation_rate,
        'real_discount_rate': real_discount_rate,
        'project_lifetime': project_lifetime,
        'crf': system_crf,
        'pv': pv_npc_data,
        'wind': wind_npc_data,
        'hydro': hydro_npc_data,
        'bess': bess_npc_data,
        'total_capital': total_capital,
        'total_replacement_pv': total_replacement_pv,
        'total_om_pv': total_om_pv,
        'total_salvage_pv': total_salvage_pv,
        'total_npc': total_npc,
        'total_annualized': total_annualized,
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

    nominal_rate = config['discount_rate']
    inflation_rate = config.get('inflation_rate', 0.02)
    real_rate = calculate_real_discount_rate(nominal_rate, inflation_rate)

    print(f"\nDiscount Rates:")
    print(f"  Nominal: {nominal_rate*100:.2f}%")
    print(f"  Inflation: {inflation_rate*100:.2f}%")
    print(f"  Real (HOMER): {real_rate*100:.4f}%")

    def generate_range_with_endpoint(start, end, step):
        values = []
        current = start
        while current <= end:
            values.append(current)
            current += step
        if values[-1] != end:
            values.append(end)
        return np.array(values)

    pv_range = generate_range_with_endpoint(grid_config['pv_start'], grid_config['pv_end'], grid_config['pv_step'])
    wind_range = generate_range_with_endpoint(grid_config['wind_start'], grid_config['wind_end'], grid_config['wind_step'])
    hydro_range = generate_range_with_endpoint(grid_config['hydro_start'], grid_config['hydro_end'], grid_config['hydro_step'])
    bess_range = generate_range_with_endpoint(grid_config['bess_start'], grid_config['bess_end'], grid_config['bess_step'])

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

                    hydro_start, hydro_end, _ = find_optimal_hydro_window(
                        load_profile, pvsyst_profile, wind_profile,
                        pv_cap, wind_cap, hydro_cap,
                        bess_power, bess_capacity,
                        solar, wind, hydro, bess
                    )

                    dispatch = calculate_dispatch_with_hydro(
                        load_profile, pvsyst_profile, wind_profile,
                        pv_cap, wind_cap, hydro_cap,
                        bess_power, bess_capacity,
                        solar, wind, hydro, bess,
                        hydro_start, hydro_end
                    )

                    total_load = dispatch['Load_kW'].sum()
                    total_unmet = dispatch['Unmet_Load_kW'].sum()
                    unmet_percent = (total_unmet / total_load * 100) if total_load > 0 else 0
                    feasible = unmet_percent <= target_unmet
                    total_energy_served = total_load - total_unmet

                    npc_data = calculate_npc_homer_style(
                        pv_cap, wind_cap, hydro_cap, bess_power, bess_capacity,
                        solar, wind, hydro, bess, config,
                        lcoe_tables, use_dynamic, total_energy_served
                    )

                    pv_energy = dispatch['PV_Available_kW'].sum()
                    wind_energy = dispatch['Wind_Output_kW'].sum()
                    hydro_energy = dispatch['Hydro_Output_kW'].sum()
                    bess_discharge_annual = dispatch['BESS_Discharge_wieff_kW'].sum()
                    excess = dispatch['Curtailment_kW'].sum()
                    cycles_per_year = bess_discharge_annual / bess_capacity if bess_capacity > 0 else 0

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
                        'NPC_$': npc_data['total_npc'],
                        'Capital_$': npc_data['total_capital'],
                        'Replacement_$': npc_data['total_replacement_pv'],
                        'OM_$': npc_data['total_om_pv'],
                        'Salvage_$': npc_data['total_salvage_pv'],
                        'Annualized_$/yr': npc_data['total_annualized'],
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

    LCOE per component = (Component_NPC × CRF) / Annual_Energy
    LCOS for BESS      = (BESS_NPC × CRF) / Annual_Throughput

    CRF is extracted from component_npc_data['crf'] which is calculated using
    the real discount rate (HOMER methodology).
    """
    metrics = {}

    # Extract CRF from NPC data (calculated with real discount rate)
    crf = component_npc_data.get('crf', None)
    if crf is None or crf <= 0:
        # Fallback: recalculate from project lifetime (undiscounted approximation)
        crf = 1.0 / project_lifetime

    # PV Metrics
    if component_capacities['pv_kw'] > 0:
        pv_output = dispatch_df['PV_Available_kW'].values
        pv_total_production = pv_output.sum()
        pv_hours_operation = (pv_output > 0).sum()
        pv_mean_output = pv_output[pv_output > 0].mean() if pv_hours_operation > 0 else 0
        pv_capacity_factor = (pv_total_production / (component_capacities['pv_kw'] * 8760)) * 100
        pv_lcoe = calculate_component_lcoe_from_npc(
            component_npc_data['pv']['npc'], pv_total_production, crf
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
        metrics['pv'] = {'rated_capacity_kw': 0, 'mean_output_kw': 0, 'capacity_factor_pct': 0,
                         'total_production_kwh': 0, 'hours_of_operation': 0,
                         'levelized_cost_per_kwh': 0, 'levelized_cost_per_mwh': 0}

    # Wind Metrics
    if component_capacities['wind_kw'] > 0:
        wind_output = dispatch_df['Wind_Output_kW'].values
        wind_total_production = wind_output.sum()
        wind_hours_operation = (wind_output > 0).sum()
        wind_mean_output = wind_output[wind_output > 0].mean() if wind_hours_operation > 0 else 0
        wind_capacity_factor = (wind_total_production / (component_capacities['wind_kw'] * 8760)) * 100
        wind_lcoe = calculate_component_lcoe_from_npc(
            component_npc_data['wind']['npc'], wind_total_production, crf
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
        metrics['wind'] = {'rated_capacity_kw': 0, 'mean_output_kw': 0, 'capacity_factor_pct': 0,
                           'total_production_kwh': 0, 'hours_of_operation': 0,
                           'levelized_cost_per_kwh': 0, 'levelized_cost_per_mwh': 0}

    # Hydro Metrics
    if component_capacities['hydro_kw'] > 0:
        hydro_output = dispatch_df['Hydro_Output_kW'].values
        hydro_total_production = hydro_output.sum()
        hydro_hours_operation = (hydro_output > 0).sum()
        hydro_mean_output = hydro_output[hydro_output > 0].mean() if hydro_hours_operation > 0 else 0
        hydro_capacity_factor = (hydro_total_production / (component_capacities['hydro_kw'] * 8760)) * 100
        hydro_lcoe = calculate_component_lcoe_from_npc(
            component_npc_data['hydro']['npc'], hydro_total_production, crf
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
        metrics['hydro'] = {'rated_capacity_kw': 0, 'mean_output_kw': 0, 'capacity_factor_pct': 0,
                            'total_production_kwh': 0, 'hours_of_operation': 0,
                            'levelized_cost_per_kwh': 0, 'levelized_cost_per_mwh': 0}

    # BESS Metrics
    if component_capacities['bess_kwh'] > 0:
        bess_charge_woeff = dispatch_df['BESS_Charge_woeff_kW'].values
        bess_discharge_wieff = dispatch_df['BESS_Discharge_wieff_kW'].values
        energy_in = bess_charge_woeff.sum()
        energy_out = bess_discharge_wieff.sum()
        charge_loss = (dispatch_df['BESS_Charge_woeff_kW'] - dispatch_df['BESS_Charge_wieff_kW']).sum()
        discharge_loss = (dispatch_df['BESS_Discharge_woeff_kW'] - dispatch_df['BESS_Discharge_wieff_kW']).sum()
        losses = charge_loss + discharge_loss
        annual_throughput = energy_out
        mean_load = dispatch_df['Load_kW'].mean()
        usable_capacity = component_capacities['bess_kwh'] * (
            component_configs['bess_max_soc'] - component_configs['bess_min_soc']
        )
        autonomy_hours = usable_capacity / mean_load if mean_load > 0 else 0
        bess_lcos = calculate_bess_lcos_from_npc(
            component_npc_data['bess']['npc'], annual_throughput, crf
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
        metrics['bess'] = {'nominal_capacity_kwh': 0, 'usable_capacity_kwh': 0, 'autonomy_hours': 0,
                           'energy_in_kwh': 0, 'energy_out_kwh': 0, 'losses_kwh': 0,
                           'annual_throughput_kwh': 0, 'expected_life_years': 0,
                           'levelized_cost_per_kwh': 0, 'levelized_cost_per_mwh': 0}

    return metrics


def calculate_component_lcoe_from_npc(component_npc, annual_energy_kwh, crf):
    """
    Calculate component-specific LCOE from NPC and annual energy production.

    Formula: LCOE = (Component_NPC × CRF) / Annual_Energy_kWh

    This is the correct HOMER-equivalent method. The CRF (Capital Recovery Factor)
    converts the NPC to an equivalent uniform annual cost, which is then divided
    by annual energy output to get $/kWh.

    Note: The old formula NPC / (E × n) is INCORRECT — it ignores the time value
    of money (equivalent to assuming a 0% discount rate). At 8% nominal / 2%
    inflation, CRF ≈ 0.0774 vs 1/25 = 0.040, a ~48% understatement.

    Args:
        component_npc:     Component Net Present Cost ($)
        annual_energy_kwh: Annual energy production (kWh/year)
        crf:               Capital Recovery Factor (real discount rate based)

    Returns:
        LCOE in $/kWh
    """
    if annual_energy_kwh <= 0 or crf <= 0:
        return 0.0
    return (component_npc * crf) / annual_energy_kwh


def calculate_bess_lcos_from_npc(bess_npc, annual_throughput_kwh, crf):
    """
    Calculate BESS Levelized Cost of Storage (LCOS) from NPC and annual throughput.

    Formula: LCOS = (BESS_NPC × CRF) / Annual_Throughput_kWh

    This is the correct method consistent with HOMER methodology. CRF converts
    the NPC into an equivalent annual cost, divided by annual energy discharged.

    Note: The old formula NPC / (throughput × n) is INCORRECT — it ignores the
    time value of money. At 8% nominal / 2% inflation, this understates LCOS
    by ~48%.

    Args:
        bess_npc:              BESS Net Present Cost ($)
        annual_throughput_kwh: Annual energy discharged (kWh/year)
        crf:                   Capital Recovery Factor (real discount rate based)

    Returns:
        LCOS in $/kWh
    """
    if annual_throughput_kwh <= 0 or crf <= 0:
        return 0.0
    return (bess_npc * crf) / annual_throughput_kwh


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

    translated = optimal.copy()
    translated['NPC_Total']            = optimal['NPC_$']
    translated['CapEx_Total']          = optimal['Capital_$']
    translated['OpEx_Annual']          = optimal['Annualized_$/yr']
    translated['LCOE_per_kWh']         = optimal['LCOE_$/kWh']
    translated['Unmet_Load_Percent']   = optimal['Unmet_%']
    translated['Unmet_Load_kWh']       = optimal['Unmet_kWh']
    translated['Total_Curtailment_kWh']= optimal['Excess_kWh']
    translated['Total_Load_kWh']       = optimal['Total_Load_kWh']
    translated['PV_NPC']               = optimal['PV_NPC_$']
    translated['Wind_NPC']             = optimal['Wind_NPC_$']
    translated['Hydro_NPC']            = optimal['Hydro_NPC_$']
    translated['BESS_NPC']             = optimal['BESS_NPC_$']
    translated['PV_CapEx']             = optimal['PV_Capital_$']
    translated['Wind_CapEx']           = optimal['Wind_Capital_$']
    translated['Hydro_CapEx']          = optimal['Hydro_Capital_$']
    translated['BESS_CapEx']           = optimal['BESS_Capital_$']
    translated['PV_OpEx_Annual']       = optimal['PV_Annualized_$/yr']
    translated['Wind_OpEx_Annual']     = optimal['Wind_Annualized_$/yr']
    translated['Hydro_OpEx_Annual']    = optimal['Hydro_Annualized_$/yr']
    translated['BESS_OpEx_Annual']     = optimal['BESS_Annualized_$/yr']
    translated['PV_Replacement']       = optimal['PV_Replacement_$']
    translated['Wind_Replacement']     = optimal['Wind_Replacement_$']
    translated['Hydro_Replacement']    = optimal['Hydro_Replacement_$']
    translated['BESS_Replacement']     = optimal['BESS_Replacement_$']
    translated['PV_OM']                = optimal['PV_OM_$']
    translated['Wind_OM']              = optimal['Wind_OM_$']
    translated['Hydro_OM']             = optimal['Hydro_OM_$']
    translated['BESS_OM']              = optimal['BESS_OM_$']
    translated['PV_Salvage']           = optimal['PV_Salvage_$']
    translated['Wind_Salvage']         = optimal['Wind_Salvage_$']
    translated['Hydro_Salvage']        = optimal['Hydro_Salvage_$']
    translated['BESS_Salvage']         = optimal['BESS_Salvage_$']
    translated['Total_Replacement']    = optimal['Replacement_$']
    translated['Total_OM']             = optimal['OM_$']
    translated['Total_Salvage']        = optimal['Salvage_$']
    translated['Real_Discount_Rate']   = optimal['Real_Discount_Rate_%'] / 100
    translated['CRF']                  = optimal['CRF']
    translated['PV_Energy_kWh']        = optimal['PV_Energy_kWh']
    translated['Wind_Energy_kWh']      = optimal['Wind_Energy_kWh']
    translated['Hydro_Energy_kWh']     = optimal['Hydro_Energy_kWh']
    translated['Total_Energy_Served_kWh'] = optimal['Total_Energy_Served_kWh']

    print("\n" + "="*70)
    print("✓ OPTIMAL SOLUTION FOUND")
    print("="*70)
    print(f"  PV:         {optimal['PV_kW']:.0f} kW")
    print(f"  Wind:       {optimal['Wind_kW']:.0f} kW")
    print(f"  Hydro:      {optimal['Hydro_kW']:.0f} kW")
    print(f"  BESS Power: {optimal['BESS_Power_kW']:.0f} kW")
    print(f"  NPC:        ${optimal['NPC_$']:,.2f}")
    print(f"  LCOE:       ${optimal['LCOE_$/kWh']:.4f}/kWh")
    print(f"  Unmet:      {optimal['Unmet_%']:.2f}%")
    print(f"  Wind Energy:{optimal['Wind_Energy_kWh']:,.0f} kWh/yr")
    print("="*70)

    return translated


# ==============================================================================
# DEGRADATION ANALYSIS FUNCTIONS
# ==============================================================================

def apply_pv_degradation_simple(pv_profile, year, annual_rate_pct):
    """Apply PV degradation using simple annual rate.
    degradation_factor = (1 - rate)^(year-1)  so Year 1 = no degradation.
    """
    annual_rate = annual_rate_pct / 100
    degradation_factor = (1 - annual_rate) ** (year - 1)
    return pv_profile * degradation_factor


def apply_pv_degradation_curve(pv_profile, year, degradation_curve):
    """Apply PV degradation using custom cumulative curve {year: cumulative_%}."""
    if year not in degradation_curve:
        return pv_profile
    cumulative_deg_pct = degradation_curve[year]
    degradation_factor = 1 - (cumulative_deg_pct / 100)
    return pv_profile * degradation_factor


# ==============================================================================
# WIND DEGRADATION
# ==============================================================================

def apply_wind_degradation_simple(wind_profile, year, annual_rate_pct):
    """Apply wind degradation using simple annual rate.
    Identical in structure to PV simple degradation.
    Typical wind turbine degradation: 0.1–0.5%/year (blade erosion, bearing wear).
    Year 1 = no degradation (factor = 1.0).
    """
    annual_rate = annual_rate_pct / 100
    degradation_factor = (1 - annual_rate) ** (year - 1)
    return wind_profile * degradation_factor


def apply_wind_degradation_curve(wind_profile, year, degradation_curve):
    """Apply wind degradation using custom cumulative curve {year: cumulative_%}."""
    if year not in degradation_curve:
        return wind_profile
    cumulative_deg_pct = degradation_curve[year]
    degradation_factor = 1 - (cumulative_deg_pct / 100)
    return wind_profile * degradation_factor


# ==============================================================================
# HYDRO DEGRADATION
# ==============================================================================

def apply_hydro_degradation_table(hydro_capacity_kw, year, hydro_deg_table):
    """Apply hydro degradation using a user-editable year-by-year table.

    Hydro plants are known for very long low-degradation periods (15-20 years),
    followed by gradual efficiency loss from turbine wear, sediment, cavitation.

    Args:
        hydro_capacity_kw: Installed hydro capacity (kW)
        year: Project year (1 to project_lifetime)
        hydro_deg_table: dict {year: output_factor_%}
                         e.g. {1: 100, 2: 100, ..., 15: 100, 16: 99.5, ...}
                         Value = % of original rated capacity available that year.

    Returns:
        Effective hydro capacity for that year (kW)
    """
    if year not in hydro_deg_table:
        return hydro_capacity_kw  # no data → assume no degradation
    output_factor_pct = hydro_deg_table[year]
    return hydro_capacity_kw * (output_factor_pct / 100)


def build_default_hydro_deg_table(project_lifetime, stable_years=15, annual_deg_after_pct=0.5):
    """Build the default hydro degradation table.

    Default behaviour:
      Years 1 → stable_years  : 100% output (no degradation)
      Years stable_years+1 → end: compound degradation at annual_deg_after_pct %/year

    Args:
        project_lifetime: Total project years
        stable_years: Years with no degradation (default 15)
        annual_deg_after_pct: Annual degradation rate after stable period (default 0.5%)

    Returns:
        dict {year: output_factor_%}
    """
    table = {}
    for yr in range(1, project_lifetime + 1):
        if yr <= stable_years:
            table[yr] = 100.0
        else:
            years_after_stable = yr - stable_years
            factor = (1 - annual_deg_after_pct / 100) ** years_after_stable
            table[yr] = round(factor * 100, 4)
    return table


# ==============================================================================
# BESS DEGRADATION (CSV-based — existing) + SIMPLE MODE (new)
# ==============================================================================

def apply_bess_degradation(bess_capacity_kwh, bess_power_kw, year, degradation_data):
    """Apply BESS degradation from CSV curve {year: {capacity, charge_eff, discharge_eff}}."""
    if year not in degradation_data:
        return bess_capacity_kwh, None, None
    year_data = degradation_data[year]
    capacity_retention_pct = year_data['capacity']
    degraded_capacity = bess_capacity_kwh * (capacity_retention_pct / 100)
    charge_eff = year_data['charge_eff'] / 100
    discharge_eff = year_data['discharge_eff'] / 100
    return degraded_capacity, charge_eff, discharge_eff


def build_bess_simple_degradation_data(project_lifetime, annual_capacity_deg_pct,
                                        charge_eff_pct, discharge_eff_pct,
                                        annual_charge_eff_deg_pct=0.0,
                                        annual_discharge_eff_deg_pct=0.0):
    """Build a BESS degradation dict using simple annual degradation rates.

    All three parameters (capacity, charge efficiency, discharge efficiency) degrade
    independently using compound annual rates from their Year-1 baseline values.

    Degradation formula (same for all three):
        value_year_n = baseline * (1 - annual_rate)^(year - 1)
        Year 1 = baseline (no degradation applied yet)

    Args:
        project_lifetime:             Total project years
        annual_capacity_deg_pct:      Annual capacity degradation rate (e.g. 2.0 = 2%/yr)
        charge_eff_pct:               Year-1 charging efficiency % (e.g. 90.0)
        discharge_eff_pct:            Year-1 discharging efficiency % (e.g. 95.0)
        annual_charge_eff_deg_pct:    Annual charging efficiency degradation rate (e.g. 0.5 = 0.5%/yr)
                                      Set to 0.0 to hold charge efficiency constant.
        annual_discharge_eff_deg_pct: Annual discharging efficiency degradation rate (e.g. 0.2 = 0.2%/yr)
                                      Set to 0.0 to hold discharge efficiency constant.

    Returns:
        dict {year: {'capacity': %, 'charge_eff': %, 'discharge_eff': %}}
        compatible with apply_bess_degradation()
    """
    data = {}
    cap_rate = annual_capacity_deg_pct / 100
    chg_rate = annual_charge_eff_deg_pct / 100
    dis_rate = annual_discharge_eff_deg_pct / 100

    for yr in range(1, project_lifetime + 1):
        exponent = yr - 1  # Year 1 -> no degradation
        capacity      = round(100.0          * (1 - cap_rate) ** exponent, 4)
        charge_eff    = round(charge_eff_pct    * (1 - chg_rate) ** exponent, 4)
        discharge_eff = round(discharge_eff_pct * (1 - dis_rate) ** exponent, 4)
        data[yr] = {
            'capacity':      max(0.0,  capacity),
            'charge_eff':    max(50.0, charge_eff),    # floor at 50%
            'discharge_eff': max(50.0, discharge_eff)  # floor at 50%
        }
    return data


def run_multi_year_degradation_analysis(
    optimal_config, load_profile, pvsyst_profile, wind_profile,
    solar_config, wind_config, hydro_config, bess_config,
    project_lifetime=25,
    pv_degradation_type=None,
    pv_degradation_data=None,
    wind_degradation_type=None,
    wind_degradation_data=None,
    hydro_degradation_table=None,
    bess_degradation_data=None,
    initial_soc_percent=50.0
):
    """Run multi-year degradation analysis for all project years.

    Supports independent degradation for all four components:
      PV   : 'simple' (annual rate) or 'curve' (custom CSV)
      Wind : 'simple' (annual rate) or 'curve' (custom CSV)
      Hydro: year-by-year output factor table (dict {year: output_%})
             Default table has 0% degradation for first 15 years, then gradual decline.
      BESS : CSV curve dict OR simple annual rate dict
             (built via build_bess_simple_degradation_data() before calling this fn)
    """
    print("\n" + "="*70)
    print(f"RUNNING {project_lifetime}-YEAR DEGRADATION ANALYSIS")
    print("="*70)

    yearly_results = []
    selected_years = list(range(1, project_lifetime + 1))
    yearly_dispatch = {}

    baseline_bess_capacity = optimal_config.get('BESS_Capacity_kWh', 0)
    baseline_hydro_capacity = optimal_config.get('Hydro_kW', 0)
    carry_soc_kwh = (initial_soc_percent / 100) * baseline_bess_capacity

    # Build default hydro table if not provided but hydro is enabled
    if hydro_degradation_table is None and baseline_hydro_capacity > 0:
        hydro_degradation_table = build_default_hydro_deg_table(project_lifetime)

    print(f"\n  Starting SOC: {initial_soc_percent:.1f}% = {carry_soc_kwh:.1f} kWh")
    print(f"  PV degradation:    {pv_degradation_type or 'none'}")
    print(f"  Wind degradation:  {wind_degradation_type or 'none'}")
    print(f"  Hydro degradation: {'table' if hydro_degradation_table else 'none'}")
    print(f"  BESS degradation:  {'curve/simple' if bess_degradation_data else 'none'}")

    for year in range(1, project_lifetime + 1):
        if year % 5 == 0 or year == 1:
            print(f"  Processing Year {year}...")

        # ── PV degradation ──
        if pv_degradation_type == 'simple' and pv_degradation_data is not None:
            pv_gen_degraded = apply_pv_degradation_simple(pvsyst_profile, year, pv_degradation_data)
            pv_deg_pct = (1 - (1 - pv_degradation_data / 100) ** (year - 1)) * 100
        elif pv_degradation_type == 'curve' and pv_degradation_data is not None:
            pv_gen_degraded = apply_pv_degradation_curve(pvsyst_profile, year, pv_degradation_data)
            pv_deg_pct = pv_degradation_data.get(year, 0)
        else:
            pv_gen_degraded = pvsyst_profile
            pv_deg_pct = 0

        # ── Wind degradation ──
        if wind_degradation_type == 'simple' and wind_degradation_data is not None:
            wind_gen_degraded = apply_wind_degradation_simple(wind_profile, year, wind_degradation_data)
            wind_deg_pct = (1 - (1 - wind_degradation_data / 100) ** (year - 1)) * 100
        elif wind_degradation_type == 'curve' and wind_degradation_data is not None:
            wind_gen_degraded = apply_wind_degradation_curve(wind_profile, year, wind_degradation_data)
            wind_deg_pct = wind_degradation_data.get(year, 0)
        else:
            wind_gen_degraded = wind_profile
            wind_deg_pct = 0

        # ── Hydro degradation ──
        if hydro_degradation_table is not None and baseline_hydro_capacity > 0:
            hydro_cap_degraded = apply_hydro_degradation_table(
                baseline_hydro_capacity, year, hydro_degradation_table
            )
            hydro_output_factor = hydro_degradation_table.get(year, 100.0)
            hydro_deg_pct = 100.0 - hydro_output_factor
        else:
            hydro_cap_degraded = baseline_hydro_capacity
            hydro_deg_pct = 0

        # ── BESS degradation ──
        if bess_degradation_data is not None:
            bess_capacity_degraded, charge_eff_deg, discharge_eff_deg = apply_bess_degradation(
                optimal_config['BESS_Capacity_kWh'], optimal_config['BESS_Power_kW'],
                year, bess_degradation_data
            )
            bess_config_degraded = bess_config.copy()
            if charge_eff_deg is not None:
                bess_config_degraded['charge_eff'] = charge_eff_deg
                bess_config_degraded['discharge_eff'] = discharge_eff_deg
            bess_retention_pct = (bess_capacity_degraded / baseline_bess_capacity * 100) if baseline_bess_capacity > 0 else 100
        else:
            bess_capacity_degraded = optimal_config['BESS_Capacity_kWh']
            bess_config_degraded = bess_config
            bess_retention_pct = 100

        # ── Dispatch simulation for this year ──
        dispatch_df = calculate_dispatch_with_hydro(
            load_profile, pv_gen_degraded, wind_gen_degraded,
            optimal_config['PV_kW'], optimal_config['Wind_kW'], hydro_cap_degraded,
            optimal_config['BESS_Power_kW'], bess_capacity_degraded,
            solar_config, wind_config, hydro_config, bess_config_degraded,
            int(optimal_config.get('Hydro_Window_Start', 0)),
            int(optimal_config.get('Hydro_Window_End', 24)),
            initial_soc_kwh=carry_soc_kwh
        )

        carry_soc_kwh = getattr(dispatch_df, '_final_soc_kwh', carry_soc_kwh)

        if year % 5 == 0 or year <= 2:
            ending_soc_pct = (carry_soc_kwh / bess_capacity_degraded * 100) if bess_capacity_degraded > 0 else 0
            print(f"    Year {year} ending SOC: {carry_soc_kwh:.1f} kWh ({ending_soc_pct:.1f}%)")

        total_load = dispatch_df['Load_kW'].sum()
        total_unmet = dispatch_df['Unmet_Load_kW'].sum()
        total_served = total_load - total_unmet

        annual_metrics = {
            'Year': year,
            'PV_Degradation_%':    pv_deg_pct,
            'Wind_Degradation_%':  wind_deg_pct,
            'Hydro_Degradation_%': hydro_deg_pct,
            'BESS_Retention_%':    bess_retention_pct,
            'PV_Energy_MWh':       dispatch_df['PV_Available_kW'].sum() / 1000,
            'Wind_Energy_MWh':     dispatch_df['Wind_Output_kW'].sum() / 1000,
            'Hydro_Energy_MWh':    dispatch_df['Hydro_Output_kW'].sum() / 1000,
            'Load_MWh':            total_load / 1000,
            'Served_MWh':          total_served / 1000,
            'Unmet_MWh':           total_unmet / 1000,
            'Unmet_%':             (total_unmet / total_load * 100) if total_load > 0 else 0,
            'BESS_Throughput_MWh': dispatch_df['BESS_Discharge_wieff_kW'].sum() / 1000,
            'Curtailment_MWh':     dispatch_df['Curtailment_kW'].sum() / 1000,
        }

        yearly_results.append(annual_metrics)

        if year in selected_years:
            dispatch_df['Year'] = year
            yearly_dispatch[f'Year_{year}'] = dispatch_df.copy()

    yearly_metrics_df = pd.DataFrame(yearly_results)

    last_idx = len(yearly_metrics_df) - 1
    degradation_summary = {
        'pv_degradation_year_1':    yearly_metrics_df.loc[0,        'PV_Degradation_%'],
        'pv_degradation_year_last': yearly_metrics_df.loc[last_idx, 'PV_Degradation_%'],
        'wind_degradation_year_1':  yearly_metrics_df.loc[0,        'Wind_Degradation_%'],
        'wind_degradation_year_last': yearly_metrics_df.loc[last_idx,'Wind_Degradation_%'],
        'hydro_degradation_year_1': yearly_metrics_df.loc[0,        'Hydro_Degradation_%'],
        'hydro_degradation_year_last': yearly_metrics_df.loc[last_idx,'Hydro_Degradation_%'],
        'bess_retention_year_1':    yearly_metrics_df.loc[0,        'BESS_Retention_%'],
        'bess_retention_year_last': yearly_metrics_df.loc[last_idx, 'BESS_Retention_%'],
        'avg_unmet_pct':            yearly_metrics_df['Unmet_%'].mean(),
        'max_unmet_pct':            yearly_metrics_df['Unmet_%'].max(),
        'total_energy_served_GWh':  yearly_metrics_df['Served_MWh'].sum() / 1000,
        # keep old key for backward compat with display code
        'pv_degradation_year_25':   yearly_metrics_df.loc[last_idx, 'PV_Degradation_%'],
        'bess_retention_year_25':   yearly_metrics_df.loc[last_idx, 'BESS_Retention_%'],
        'total_energy_served_25yr_GWh': yearly_metrics_df['Served_MWh'].sum() / 1000,
    }

    print(f"\n✓ Degradation Analysis Complete")
    print(f"  Average Unmet Load: {degradation_summary['avg_unmet_pct']:.2f}%")
    print("="*70)

    return {
        'yearly_metrics':         yearly_metrics_df,
        'selected_year_dispatch': yearly_dispatch,
        'degradation_summary':    degradation_summary,
        'optimal_config':         optimal_config,
        'hydro_deg_table':        hydro_degradation_table,
    }


def load_bess_degradation_from_csv(csv_path):
    """Load BESS degradation data from CSV file."""
    try:
        df = pd.read_csv(csv_path)
        required_cols = ['Year', 'Capacity_Retention_%', 'Charging_Efficiency_%', 'Discharging_Efficiency_%']
        if not all(col in df.columns for col in required_cols):
            print(f"❌ Error: CSV must have columns: {', '.join(required_cols)}")
            return None
        deg_data = {}
        for _, row in df.iterrows():
            year = int(row['Year'])
            deg_data[year] = {
                'capacity': float(row['Capacity_Retention_%']),
                'charge_eff': float(row['Charging_Efficiency_%']),
                'discharge_eff': float(row['Discharging_Efficiency_%'])
            }
        print(f"✓ Loaded BESS degradation data: {len(deg_data)} years from {csv_path}")
        return deg_data
    except Exception as e:
        print(f"❌ Error loading BESS degradation CSV: {str(e)}")
        return None


# ==============================================================================
# MODULE ENTRY POINT
# ==============================================================================

if __name__ == "__main__":
    print("\n" + "="*70)
    print("OPTIMIZATION MODULE - WITH DEGRADATION v4.1")
    print("Fix: Wind output now correctly uses wind_capacity > 0 check")
    print("="*70)
