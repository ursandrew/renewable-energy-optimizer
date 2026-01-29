"""
GRID SEARCH OPTIMIZER - PV + WIND + HYDRO + BESS (HOMER NPC METHOD)
===================================================================

ENHANCED VERSION with HOMER Pro NPC Calculation:
- Uses real discount rate (inflation-adjusted)
- Accounts for component replacements
- Accounts for salvage value
- Uses Capital Recovery Factor (CRF)
- Calculates Levelized COE matching HOMER
- Detailed cost breakdown by component

Supports:
- PV + Wind + Hydro + BESS hybrid optimization
- Static & Dynamic LCOE
- 4D grid search (PV × Wind × Hydro × BESS)
- Automatic hydro operating window optimization
- HOMER-style NPC calculation
"""

import pandas as pd
import numpy as np
from datetime import datetime
import time

INPUT_FILE = "input_gridsearch_STREAMLITCHECK.xlsx"
OUTPUT_FILE = "output_gridsearch_hybrid_hydro_Streamlitheck.xlsx"


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
# STEP 1: READ INPUTS
# ==============================================================================

def read_inputs():
    """Read all inputs from Excel file."""
    
    print("\n" + "="*70)
    print("READING INPUT FILE: " + INPUT_FILE)
    print("="*70)
    
    # Configuration
    df_config = pd.read_excel(INPUT_FILE, sheet_name='Configuration')
    config = {}
    for _, row in df_config.iterrows():
        param = str(row['Parameter']).strip().lower()
        value = row['Value']
        
        if 'simulation hours' in param:
            config['simulation_hours'] = int(value)
        elif 'target unmet' in param:
            config['target_unmet_percent'] = float(value)
        elif 'optimization method' in param:
            config['optimization_method'] = str(value).upper()
        elif 'discount rate' in param and 'inflation' not in param:
            config['discount_rate'] = float(value)
        elif 'inflation' in param:
            config['inflation_rate'] = float(value)
        elif 'project lifetime' in param or 'lifetime' in param:
            config['project_lifetime'] = int(value)
        elif 'dynamic lcoe' in param:
            config['use_dynamic_lcoe'] = str(value).upper() == 'YES'
    
    # Set default inflation if not provided
    if 'inflation_rate' not in config:
        config['inflation_rate'] = 0.02  # Default 2%
        print(f"  ⚠️  Inflation Rate not found, using default: 2%")
    
    print(f"✓ Configuration: {config['simulation_hours']} hours")
    print(f"  Target Unmet: {config['target_unmet_percent']}%")
    print(f"  Nominal Discount Rate: {config['discount_rate']*100:.1f}%")
    print(f"  Inflation Rate: {config.get('inflation_rate', 0.02)*100:.1f}%")
    print(f"  LCOE Mode: {'DYNAMIC' if config.get('use_dynamic_lcoe', False) else 'STATIC'}")
    
    # Grid Search Config
    df_grid = pd.read_excel(INPUT_FILE, sheet_name='Grid_Search_Config')
    grid_config = {}
    for _, row in df_grid.iterrows():
        param = row['Parameter'].strip()
        value = row['Value']
        if param == 'PV Search Start':
            grid_config['pv_start'] = float(value)
        elif param == 'PV Search End':
            grid_config['pv_end'] = float(value)
        elif param == 'PV Search Step':
            grid_config['pv_step'] = float(value)
        elif param == 'Wind Search Start':
            grid_config['wind_start'] = float(value)
        elif param == 'Wind Search End':
            grid_config['wind_end'] = float(value)
        elif param == 'Wind Search Step':
            grid_config['wind_step'] = float(value)
        elif param == 'Hydro Search Start':
            grid_config['hydro_start'] = float(value)
        elif param == 'Hydro Search End':
            grid_config['hydro_end'] = float(value)
        elif param == 'Hydro Search Step':
            grid_config['hydro_step'] = float(value)
        elif param == 'BESS Search Start':
            grid_config['bess_start'] = float(value)
        elif param == 'BESS Search End':
            grid_config['bess_end'] = float(value)
        elif param == 'BESS Search Step':
            grid_config['bess_step'] = float(value)
        elif param == 'Max Combinations':
            grid_config['max_combinations'] = int(value)
    
    # Calculate number of combinations (includes endpoints)
    def count_options_with_endpoint(start, end, step):
        """Count options when endpoint is always included."""
        count = int((end - start) / step) + 1  # Regular steps
        # Check if endpoint aligns with step
        if (end - start) % step != 0:
            # Endpoint doesn't align, but we include it anyway
            pass  # count is already correct (includes partial step)
        return count
    
    pv_options = count_options_with_endpoint(grid_config['pv_start'], grid_config['pv_end'], grid_config['pv_step'])
    wind_options = count_options_with_endpoint(grid_config['wind_start'], grid_config['wind_end'], grid_config['wind_step'])
    hydro_options = count_options_with_endpoint(grid_config['hydro_start'], grid_config['hydro_end'], grid_config['hydro_step'])
    bess_options = count_options_with_endpoint(grid_config['bess_start'], grid_config['bess_end'], grid_config['bess_step'])
    total_combinations = pv_options * wind_options * hydro_options * bess_options
    
    print(f"✓ Grid Search Config:")
    print(f"  PV: {grid_config['pv_start']}-{grid_config['pv_end']} kW, Step={grid_config['pv_step']} → {pv_options} options")
    print(f"  Wind: {grid_config['wind_start']}-{grid_config['wind_end']} kW, Step={grid_config['wind_step']} → {wind_options} options")
    print(f"  Hydro: {grid_config['hydro_start']}-{grid_config['hydro_end']} kW, Step={grid_config['hydro_step']} → {hydro_options} options")
    print(f"  BESS: {grid_config['bess_start']}-{grid_config['bess_end']} kW, Step={grid_config['bess_step']} → {bess_options} options")
    print(f"  Total Combinations: {total_combinations}")
    
    if total_combinations > grid_config['max_combinations']:
        print(f"  ⚠️  WARNING: Exceeds max ({grid_config['max_combinations']})")
    
    # Solar PV
    df_solar = pd.read_excel(INPUT_FILE, sheet_name='Solar_PV')
    solar = {}
    for _, row in df_solar.iterrows():
        param = row['Parameter'].strip()
        value = row['Value']
        if param == 'LCOE':
            solar['lcoe'] = float(value)
        elif param == 'PVsyst Baseline':
            solar['baseline_kw'] = float(value)
        elif param == 'Capex':
            solar['capex_per_kw'] = float(value)
        elif param == 'O&M Cost':
            solar['om_per_kw_year'] = float(value)
        elif param == 'Lifetime':
            solar['lifetime'] = int(value)
    
    print(f"✓ Solar PV: LCOE=${solar['lcoe']}/MWh, Lifetime={solar['lifetime']} years")
    
    # Wind
    try:
        df_wind = pd.read_excel(INPUT_FILE, sheet_name='Wind')
        wind = {}
        for _, row in df_wind.iterrows():
            param = row['Parameter'].strip()
            value = row['Value']
            if param == 'Include Wind?':
                wind['enabled'] = str(value).upper() == 'YES'
            elif param == 'LCOE':
                wind['lcoe'] = float(value)
            elif param == 'Capex':
                wind['capex_per_kw'] = float(value)
            elif param == 'O&M Cost':
                wind['om_per_kw_year'] = float(value)
            elif param == 'Lifetime':
                wind['lifetime'] = int(value)
        
        print(f"✓ Wind: LCOE=${wind['lcoe']}/MWh, Enabled={wind.get('enabled', False)}, Lifetime={wind['lifetime']} years")
    except:
        wind = {'enabled': False, 'lcoe': 0, 'capex_per_kw': 0, 'om_per_kw_year': 0, 'lifetime': 25}
        print(f"⚠️  Wind sheet not found - Wind disabled")
    
    # Hydro
    df_hydro = pd.read_excel(INPUT_FILE, sheet_name='Hydro')
    hydro = {}
    for _, row in df_hydro.iterrows():
        param = row['Parameter'].strip()
        value = row['Value']
        if param == 'Include Hydro?':
            hydro['enabled'] = str(value).upper() == 'YES'
        elif param == 'LCOE':
            hydro['lcoe'] = float(value)
        elif param == 'Capex':
            hydro['capex_per_kw'] = float(value)
        elif param == 'O&M Cost':
            hydro['om_per_kw_year'] = float(value)
        elif param == 'Lifetime':
            hydro['lifetime'] = int(value)
        elif param == 'Operating Hours':  # FIXED: Match actual Excel parameter name
            hydro['hours_per_day'] = int(value)
    
    print(f"✓ Hydro: LCOE=${hydro['lcoe']}/MWh, Enabled={hydro.get('enabled', False)}, Lifetime={hydro['lifetime']} years")
    print(f"  Operating Hours: {hydro.get('hours_per_day', 6)} hours/day")
    
    # BESS
    df_bess = pd.read_excel(INPUT_FILE, sheet_name='BESS')
    bess = {}
    for _, row in df_bess.iterrows():
        param = row['Parameter'].strip()
        value = row['Value']
        if param == 'Duration':
            bess['duration_hours'] = float(value)
        elif param == 'LCOS':
            bess['lcos'] = float(value)
        elif param == 'Charge Efficiency':
            bess['charge_eff'] = float(value) / 100.0
        elif param == 'Discharge Efficiency':
            bess['discharge_eff'] = float(value) / 100.0
        elif param == 'Min SOC':
            bess['min_soc'] = float(value) / 100.0
        elif param == 'Max SOC':
            bess['max_soc'] = float(value) / 100.0
        elif param == 'Power Capex':
            bess['power_capex_per_kw'] = float(value)
        elif param == 'Energy Capex':
            bess['energy_capex_per_kwh'] = float(value)
        elif param == 'O&M Cost':
            bess['om_per_kw_year'] = float(value)
        elif param == 'Lifetime':
            bess['lifetime'] = int(value)
    
    print(f"✓ BESS: LCOS=${bess['lcos']}/MWh, Duration={bess['duration_hours']}h, Lifetime={bess['lifetime']} years")
    
    # Load Profile
    df_load = pd.read_excel(INPUT_FILE, sheet_name='Load_Profile')
    load_profile = df_load['Load_kW'].values
    print(f"✓ Load Profile: {len(load_profile)} hours, Peak={load_profile.max():.2f} kW")
    
    # PVsyst Profile
    df_pvsyst = pd.read_excel(INPUT_FILE, sheet_name='PVsyst_Profile')
    pvsyst_profile = df_pvsyst['Output_kW'].values
    print(f"✓ PVsyst Profile: {len(pvsyst_profile)} hours")
    
    # Wind Profile
    if wind.get('enabled', False):
        try:
            df_wind_profile = pd.read_excel(INPUT_FILE, sheet_name='Wind_Profile')
            wind_profile = df_wind_profile['Output_kW'].values
            print(f"✓ Wind Profile: {len(wind_profile)} hours")
        except:
            wind_profile = np.zeros(config['simulation_hours'])
            print(f"⚠️  Wind_Profile sheet not found - Using zeros")
    else:
        wind_profile = np.zeros(config['simulation_hours'])
    
    print("="*70 + "\n")
    
    return config, grid_config, solar, wind, hydro, bess, load_profile, pvsyst_profile, wind_profile


def read_lcoe_tables():
    """Read LCOE tables for dynamic LCOE calculation."""
    try:
        df_lcoe = pd.read_excel(INPUT_FILE, sheet_name='LCOE_Tables')
        lcoe_tables = {'Solar': [], 'Wind': [], 'Hydro': [], 'BESS': []}
        
        for idx, row in df_lcoe.iterrows():
            cap_mw = row['Solar_Cap_MW']
            lcoe = row['Solar_LCOE']
            if pd.notna(cap_mw) and pd.notna(lcoe):
                lcoe_tables['Solar'].append((float(cap_mw), float(lcoe)))
        
        if 'Wind_Cap_MW' in df_lcoe.columns:
            for idx, row in df_lcoe.iterrows():
                cap_mw = row['Wind_Cap_MW']
                lcoe = row['Wind_LCOE']
                if pd.notna(cap_mw) and pd.notna(lcoe):
                    lcoe_tables['Wind'].append((float(cap_mw), float(lcoe)))
        
        if 'Hydro_Cap_MW' in df_lcoe.columns:
            for idx, row in df_lcoe.iterrows():
                cap_mw = row['Hydro_Cap_MW']
                lcoe = row['Hydro_LCOE']
                if pd.notna(cap_mw) and pd.notna(lcoe):
                    lcoe_tables['Hydro'].append((float(cap_mw), float(lcoe)))
        
        for idx, row in df_lcoe.iterrows():
            cap_mw = row['BESS_Cap_MW']
            lcos = row['BESS_LCOE']
            if pd.notna(cap_mw) and pd.notna(lcos):
                lcoe_tables['BESS'].append((float(cap_mw), float(lcos)))
        
        print(f"✓ LCOE Tables loaded")
        return lcoe_tables
    except Exception as e:
        print(f"⚠️  LCOE_Tables not found, using static LCOE")
        return None


def get_dynamic_lcoe(capacity_kw, tech_type, lcoe_tables, static_lcoe):
    """Get LCOE using linear interpolation from tables."""
    if lcoe_tables is None or tech_type not in lcoe_tables:
        return static_lcoe
    
    table = lcoe_tables[tech_type]
    if len(table) == 0:
        return static_lcoe
    
    capacity_mw = capacity_kw / 1000.0
    table = sorted(table, key=lambda x: x[0])
    capacities = [row[0] for row in table]
    lcoes = [row[1] for row in table]
    
    if capacity_mw <= capacities[0]:
        return lcoes[0]
    if capacity_mw >= capacities[-1]:
        return lcoes[-1]
    
    for i in range(len(capacities) - 1):
        if capacities[i] <= capacity_mw <= capacities[i + 1]:
            x1, x2 = capacities[i], capacities[i + 1]
            y1, y2 = lcoes[i], lcoes[i + 1]
            if x2 == x1:
                return y1
            lcoe = y1 + (y2 - y1) * (capacity_mw - x1) / (x2 - x1)
            return lcoe
    
    return lcoes[-1]


# ==============================================================================
# STEP 2: HYDRO WINDOW OPTIMIZATION
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
        total_unmet = dispatch['Unmet_kW'].sum()
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
# STEP 3: DISPATCH SIMULATION WITH HYDRO
# ==============================================================================

def calculate_dispatch_with_hydro(load_profile, pvsyst_profile, wind_profile,
                                  pv_capacity, wind_capacity, hydro_capacity,
                                  bess_power, bess_capacity,
                                  solar_config, wind_config, hydro_config, bess_config,
                                  hydro_window_start, hydro_window_end):
    """
    Calculate hourly dispatch for PV + Wind + Hydro + BESS system.
    
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
    
    # Initialize arrays
    pv_output = np.zeros(hours)
    wind_output = np.zeros(hours)
    hydro_output = np.zeros(hours)
    bess_soc = np.zeros(hours)
    bess_charge = np.zeros(hours)
    bess_discharge = np.zeros(hours)
    unmet_load = np.zeros(hours)
    excess_renewable = np.zeros(hours)
    hydro_active_flag = np.zeros(hours)
    
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
        
        # Variable renewable generation
        pv_output[h] = pv_avail
        wind_output[h] = wind_avail
        total_renewable = pv_output[h] + wind_output[h]
        
        net_load = load - total_renewable
        
        # Dispatch hydro if active and needed (FIXED: Removed enabled check)
        if net_load > 0 and hydro_active and hydro_capacity > 0:
            hydro_output[h] = min(net_load, hydro_capacity)
            net_load -= hydro_output[h]
        
        if net_load > 0:
            # Need to discharge battery
            available_discharge = max(0, current_soc - min_soc_kwh)
            actual_discharge = available_discharge * discharge_eff
            
            bess_discharge[h] = min(bess_power, actual_discharge, net_load)
            
            if bess_discharge[h] > 0:
                energy_removed = bess_discharge[h] / discharge_eff
                current_soc -= energy_removed
            
            unmet_load[h] = max(0, net_load - bess_discharge[h])
            
        else:
            # Excess renewable - charge battery
            excess = abs(net_load)
            available_space = max(0, max_soc_kwh - current_soc)
            
            max_charge = excess * charge_eff
            bess_charge[h] = min(bess_power, available_space, max_charge)
            
            current_soc += bess_charge[h]
            
            energy_used = bess_charge[h] / charge_eff
            excess_renewable[h] = max(0, excess - energy_used)
        
        # Enforce SOC limits
        current_soc = max(min_soc_kwh, min(current_soc, max_soc_kwh))
        bess_soc[h] = current_soc
    
    results = pd.DataFrame({
        'Hour': range(hours),
        'Load_kW': load_profile,
        'PV_Output_kW': pv_output,
        'Wind_Output_kW': wind_output,
        'Hydro_Output_kW': hydro_output,
        'Hydro_Active': hydro_active_flag,
        'BESS_Charge_kW': bess_charge,
        'BESS_Discharge_kW': bess_discharge,
        'BESS_SOC_kWh': bess_soc,
        'Unmet_kW': unmet_load,
        'Excess_kW': excess_renewable
    })
    
    return results


# ==============================================================================
# STEP 4: HOMER-STYLE NPC CALCULATION
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
    
    # Calculate Levelized Cost of Energy (LCOE) - Uses LIFETIME energy
    total_energy_lifetime = total_energy_served_annual * project_lifetime
    
    if total_energy_lifetime > 0:
        lcoe = total_npc / total_energy_lifetime  # $/kWh over project lifetime
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
        'total_energy_served_lifetime': total_energy_lifetime,
        'lcoe': lcoe,
        'lcoe_per_kwh': lcoe,
        'lcoe_per_mwh': lcoe * 1000
    }


# ==============================================================================
# STEP 5: GRID SEARCH OPTIMIZATION
# ==============================================================================

def grid_search_optimize_hydro(config, grid_config, solar, wind, hydro, bess,
                               load_profile, pvsyst_profile, wind_profile, lcoe_tables):
    """4D Grid search optimization with HOMER NPC calculation and hydro window optimization."""
    
    print("\n" + "="*70)
    print("GRID SEARCH OPTIMIZATION (HOMER NPC METHOD + HYDRO)")
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
    # This ensures we test the maximum capacities specified in config
    # Note: Last step might be smaller if endpoint doesn't align with step size
    
    def generate_range_with_endpoint(start, end, step):
        """Generate range from start to end with given step, always including end."""
        # Generate regular steps
        values = []
        current = start
        while current <= end:
            values.append(current)
            current += step
        
        # Ensure endpoint is included (if not already)
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
                    
                    if count % 100 == 0:  # FIXED: Match working code's update frequency
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
                    total_unmet = dispatch['Unmet_kW'].sum()
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
                    pv_energy = dispatch['PV_Output_kW'].sum()
                    wind_energy = dispatch['Wind_Output_kW'].sum()
                    hydro_energy = dispatch['Hydro_Output_kW'].sum()
                    bess_discharge_annual = dispatch['BESS_Discharge_kW'].sum()
                    excess = dispatch['Excess_kW'].sum()
                    
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
                        
                        # Clear renewable energy metrics
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
# STEP 6: FIND OPTIMAL SOLUTION
# ==============================================================================

def find_optimal_solution(results_df):
    """Find optimal solution from results."""
    
    feasible = results_df[results_df['Feasible'] == True]
    
    if len(feasible) == 0:
        print("\n⚠️  WARNING: No feasible solutions found!")
        return None
    
    optimal = feasible.loc[feasible['NPC_$'].idxmin()]
    
    print("\n" + "="*70)
    print("✓ OPTIMAL SOLUTION FOUND (HOMER NPC METHOD)!")
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
# STEP 7: SAVE RESULTS
# ==============================================================================

def write_results(results_df, optimal, config, grid_config, solar, wind, hydro, bess,
                 load_profile, pvsyst_profile, wind_profile):
    """Save results to Excel with HOMER-style breakdown."""
    
    print(f"\nWriting results to {OUTPUT_FILE}...")
    
    with pd.ExcelWriter(OUTPUT_FILE, engine='openpyxl') as writer:
        
        # Sheet 1: Summary
        summary_data = {
            'Parameter': [
                'Optimization Method',
                'NPC Calculation Method',
                'LCOE Mode',
                'Target Unmet Load (%)',
                'Total Combinations Tested',
                'Feasible Solutions',
                'Optimal Solution Iteration',
                '',
                '=== DISCOUNT RATES ===',
                'Nominal Discount Rate (%)',
                'Inflation Rate (%)',
                'Real Discount Rate (%)',
                'CRF',
                '',
                '=== OPTIMAL SOLUTION ===',
                'PV Capacity (kW)',
                'Wind Capacity (kW)',
                'Hydro Capacity (kW)',
                'Hydro Window Start (hour)',
                'Hydro Window End (hour)',
                'Hydro Operating Hours/Day',
                'BESS Power (kW)',
                'BESS Capacity (kWh)',
                'BESS Annual Discharge (kWh)',
                'BESS Cycles Per Year',
                '',
                '=== COSTS (HOMER METHOD) ===',
                'Total NPC ($)',
                'Total Capital ($)',
                'Total Replacement ($)',
                'Total O&M ($)',
                'Total Salvage ($)',
                'Total Annualized ($/year)',
                '',
                '=== COMPONENT NPC BREAKDOWN ===',
                'PV NPC ($)',
                'Wind NPC ($)',
                'Hydro NPC ($)',
                'BESS NPC ($)',
                '',
                '=== PERFORMANCE ===',
                'Levelized COE ($/kWh)',
                'Levelized COE ($/MWh)',
                'Unmet Load (%)',
                'Energy Served (kWh/year)',
                'Energy Served (kWh lifetime)',
                '',
                '=== RENEWABLE ENERGY MIX ===',
                'PV Fraction (%)',
                'Wind Fraction (%)',
                'Hydro Fraction (%)',
                'Total RE Penetration (%)',
                '',
                '=== ENERGY STORAGE ===',
                'BESS Contribution (%)',
                'Excess Fraction (%)'
            ],
            'Value': [
                'GRID SEARCH',
                'HOMER PRO',
                'DYNAMIC' if config.get('use_dynamic_lcoe', False) else 'STATIC',
                config['target_unmet_percent'],
                len(results_df),
                len(results_df[results_df['Feasible']==True]),
                int(optimal['Iteration']),
                '',
                '',
                f"{config['discount_rate'] * 100:.2f}",
                f"{config.get('inflation_rate', 0.02) * 100:.2f}",
                f"{optimal['Real_Discount_Rate_%']:.4f}",
                f"{optimal['CRF']:.6f}",
                '',
                '',
                optimal['PV_kW'],
                optimal['Wind_kW'],
                optimal['Hydro_kW'],
                int(optimal['Hydro_Window_Start']),
                int(optimal['Hydro_Window_End']),
                int(optimal['Hydro_Window_End']) - int(optimal['Hydro_Window_Start']),
                optimal['BESS_Power_kW'],
                f"{optimal['BESS_Capacity_kWh']:.2f}",
                f"{optimal['BESS_Annual_Discharge_kWh']:.2f}",
                f"{optimal['BESS_Cycles_Per_Year']:.2f}",
                '',
                '',
                f"${optimal['NPC_$']:,.2f}",
                f"${optimal['Capital_$']:,.2f}",
                f"${optimal['Replacement_$']:,.2f}",
                f"${optimal['OM_$']:,.2f}",
                f"${optimal['Salvage_$']:,.2f}",
                f"${optimal['Annualized_$/yr']:,.2f}",
                '',
                '',
                f"${optimal['PV_NPC_$']:,.2f}",
                f"${optimal['Wind_NPC_$']:,.2f}",
                f"${optimal['Hydro_NPC_$']:,.2f}",
                f"${optimal['BESS_NPC_$']:,.2f}",
                '',
                '',
                f"${optimal['LCOE_$/kWh']:.4f}",
                f"${optimal['LCOE_$/MWh']:.2f}",
                f"{optimal['Unmet_%']:.2f}",
                f"{optimal['Total_Energy_Served_kWh']:,.0f}",
                f"{optimal['Total_Energy_Served_kWh'] * config['project_lifetime']:,.0f}",
                '',
                '',
                f"{optimal['PV_Fraction_%']:.1f}",
                f"{optimal['Wind_Fraction_%']:.1f}",
                f"{optimal['Hydro_Fraction_%']:.1f}",
                f"{optimal['RE_Penetration_%']:.1f}",
                '',
                '',
                f"{optimal['BESS_Contribution_%']:.1f}",
                f"{optimal['Excess_Fraction_%']:.1f}"
            ]
        }
        summary_df = pd.DataFrame(summary_data)
        summary_df.to_excel(writer, sheet_name='Summary', index=False)
        
        # Sheet 2: Component Cost Breakdown
        cost_breakdown = {
            'Component': ['PV', 'Wind', 'Hydro', 'BESS', 'System'],
            'Capital ($)': [
                optimal['PV_Capital_$'],
                optimal['Wind_Capital_$'],
                optimal['Hydro_Capital_$'],
                optimal['BESS_Capital_$'],
                optimal['Capital_$']
            ],
            'Replacement ($)': [
                optimal['PV_Replacement_$'],
                optimal['Wind_Replacement_$'],
                optimal['Hydro_Replacement_$'],
                optimal['BESS_Replacement_$'],
                optimal['Replacement_$']
            ],
            'O&M ($)': [
                optimal['PV_OM_$'],
                optimal['Wind_OM_$'],
                optimal['Hydro_OM_$'],
                optimal['BESS_OM_$'],
                optimal['OM_$']
            ],
            'Salvage ($)': [
                optimal['PV_Salvage_$'],
                optimal['Wind_Salvage_$'],
                optimal['Hydro_Salvage_$'],
                optimal['BESS_Salvage_$'],
                optimal['Salvage_$']
            ],
            'Total NPC ($)': [
                optimal['PV_NPC_$'],
                optimal['Wind_NPC_$'],
                optimal['Hydro_NPC_$'],
                optimal['BESS_NPC_$'],
                optimal['NPC_$']
            ],
            'Annualized ($/yr)': [
                optimal['PV_Annualized_$/yr'],
                optimal['Wind_Annualized_$/yr'],
                optimal['Hydro_Annualized_$/yr'],
                optimal['BESS_Annualized_$/yr'],
                optimal['Annualized_$/yr']
            ]
        }
        cost_df = pd.DataFrame(cost_breakdown)
        cost_df.to_excel(writer, sheet_name='Cost_Breakdown', index=False)
        
        # Sheet 3: All Results
        results_df.to_excel(writer, sheet_name='All_Results', index=False)
        
        # Sheet 4: Feasible Solutions
        feasible = results_df[results_df['Feasible']==True].sort_values('NPC_$')
        feasible.to_excel(writer, sheet_name='Feasible_Solutions', index=False)
        
        # Sheet 5: Optimal Dispatch
        optimal_dispatch = calculate_dispatch_with_hydro(
            load_profile, pvsyst_profile, wind_profile,
            optimal['PV_kW'], optimal['Wind_kW'], optimal['Hydro_kW'],
            optimal['BESS_Power_kW'], optimal['BESS_Capacity_kWh'],
            solar, wind, hydro, bess,
            int(optimal['Hydro_Window_Start']), int(optimal['Hydro_Window_End'])
        )
        optimal_dispatch.to_excel(writer, sheet_name='Hourly_Dispatch', index=False)
        
        # Sheet 6: Hydro Window Analysis (for optimal solution)
        _, _, _, windows_df = find_optimal_hydro_window(
            load_profile, pvsyst_profile, wind_profile,
            optimal['PV_kW'], optimal['Wind_kW'], optimal['Hydro_kW'],
            optimal['BESS_Power_kW'], optimal['BESS_Capacity_kWh'],
            solar, wind, hydro, bess,
            return_all_windows=True
        )
        windows_df.to_excel(writer, sheet_name='Hydro_Window_Analysis', index=False)
    
    print(f"✓ Results saved!")
    print(f"  Sheets: Summary, Cost_Breakdown, All_Results, Feasible_Solutions, Hourly_Dispatch, Hydro_Window_Analysis")


# ==============================================================================
# MAIN EXECUTION
# ==============================================================================

def main():
    """Main execution."""
    
    print("\n" + "="*70)
    print("HYBRID RE OPTIMIZATION - GRID SEARCH")
    print("PV + Wind + Hydro + BESS")
    print("ENHANCED WITH HOMER NPC CALCULATION")
    print("="*70)
    print(f"Start: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
    
    # Read inputs
    config, grid_config, solar, wind, hydro, bess, load_profile, pvsyst_profile, wind_profile = read_inputs()
    
    # Read LCOE tables if dynamic mode
    lcoe_tables = None
    if config.get('use_dynamic_lcoe', False):
        lcoe_tables = read_lcoe_tables()
    
    # Run grid search
    results_df = grid_search_optimize_hydro(config, grid_config, solar, wind, hydro, bess,
                                           load_profile, pvsyst_profile, wind_profile, lcoe_tables)
    
    # Find optimal
    optimal = find_optimal_solution(results_df)
    
    if optimal is not None:
        # Save results
        write_results(results_df, optimal, config, grid_config, solar, wind, hydro, bess,
                     load_profile, pvsyst_profile, wind_profile)
    
    print(f"\n{'='*70}")
    print(f"End: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print(f"{'='*70}\n")


if __name__ == "__main__":
    main()