"""
DEGRADATION ANALYSIS MODULE FOR STREAMLIT
==========================================
Full degradation engine for PV + BESS systems with grid search optimization

Author: SJ
Version: 2.0 - Streamlit Compatible
"""

# ==============================================================================
# IMPORTS
# ==============================================================================

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
    st.error("Optimization module not found")

# Import degradation analysis
DEGRADATION_AVAILABLE = False
try:
    import optimize_with_degradation as deg_module
    DEGRADATION_AVAILABLE = True
except ImportError:
    pass
except Exception as e:
    print(f"Warning: Degradation module not available - {e}")

# ==============================================================================
# OEM DEGRADATION DATA
# ==============================================================================

PV_DEG = {
    1: 0, 2: 0.41, 3: 0.82, 4: 1.22, 5: 1.63, 6: 2.04, 7: 2.45, 8: 2.86, 9: 3.27, 10: 3.67,
    11: 4.08, 12: 4.49, 13: 4.90, 14: 5.31, 15: 5.71, 16: 6.12, 17: 6.53, 18: 6.94, 19: 7.35, 20: 7.76,
    21: 8.16, 22: 8.57, 23: 8.98, 24: 9.39, 25: 9.80
}

BESS_CAP_RET = {
    1: 94.46, 2: 92.14, 3: 90.33, 4: 88.71, 5: 87.19, 6: 85.75, 7: 84.37, 8: 83.00, 9: 81.70, 10: 80.45,
    11: 79.20, 12: 78.00, 13: 76.76, 14: 75.57, 15: 74.40, 16: 73.23, 17: 72.10, 18: 70.96, 19: 69.84, 20: 68.73
}

# Global variable for input file path
INPUT_FILE = None


# ==============================================================================
# GRID SEARCH OPTIMIZATION WITH DEGRADATION
# ==============================================================================

def grid_search_optimize_hydro(config, grid_config, solar, wind, hydro, bess, 
                               load_profile, pvsyst_profile, wind_profile, hydro_profile):
    """
    Grid search optimization with degradation analysis.
    
    This is a wrapper that calls the base optimization function.
    Degradation is applied in post-processing via run_degradation_analysis().
    """
    
    if not BASE_MODULE_AVAILABLE:
        raise ImportError("Base optimization module not available")
    
    # Import the grid search function
    from optimize_gridsearch_hydro_static_STREAMLITCHECK import grid_search_optimize_hydro as base_grid_search
    
    # Call base optimization (no degradation during optimization)
    results_df = base_grid_search(
        config, grid_config, solar, wind, hydro, bess,
        load_profile, pvsyst_profile, wind_profile, hydro_profile
    )
    
    return results_df


# ==============================================================================
# DEGRADATION ANALYSIS (POST-PROCESSING)
# ==============================================================================

def run_degradation_analysis(optimal_row, config_params, apply_pv=True, apply_bess=True):
    """
    Run 25-year degradation analysis on optimal configuration.
    
    Parameters:
    -----------
    optimal_row : dict
        Optimal solution from grid search
    config_params : dict
        Configuration parameters (discount_rate, project_lifetime, etc.)
    apply_pv : bool
        Apply PV degradation
    apply_bess : bool
        Apply BESS degradation
    
    Returns:
    --------
    dict : Degradation analysis results
    """
    
    # Extract capacities
    pv_kw = optimal_row.get('PV_kW', 0)
    pv_mw = pv_kw / 1000
    
    bess_power_kw = optimal_row.get('BESS_Power_kW', 0)
    bess_capacity_kwh = optimal_row.get('BESS_Capacity_kWh', 0)
    bess_mw = bess_power_kw / 1000
    bess_mwh = bess_capacity_kwh / 1000
    
    # Extract costs
    npc_y1 = optimal_row.get('NPC_$', 0)
    lcoe_y1 = optimal_row.get('LCOE_$/MWh', 0)
    bess_npc = optimal_row.get('BESS_NPC_$', 0)
    
    # Get discount rate (handle both percentage and decimal forms)
    discount_rate = config_params.get('discount_rate', 8.0)
    if discount_rate > 1:
        discount_rate = discount_rate / 100
    
    project_lifetime = config_params.get('project_lifetime', 25)
    
    # Simulate 25 years
    yearly_data = []
    
    for year in range(1, min(project_lifetime + 1, 26)):
        # PV degradation
        if apply_pv:
            pv_deg_pct = PV_DEG.get(year, 9.8)
            pv_capacity_year = pv_mw * (1 - pv_deg_pct / 100)
        else:
            pv_deg_pct = 0
            pv_capacity_year = pv_mw
        
        # BESS degradation (replacement at year 21)
        if apply_bess:
            if year <= 20:
                bess_ret_pct = BESS_CAP_RET.get(year, 70)
                bess_capacity_year = bess_mwh * (bess_ret_pct / 100)
                replaced = ''
            else:
                # After replacement at year 21, use new degradation curve
                age_after_replacement = year - 20
                bess_ret_pct = BESS_CAP_RET.get(age_after_replacement, 70)
                bess_capacity_year = bess_mwh * (bess_ret_pct / 100)
                replaced = '🔋 Replaced' if year == 21 else ''
        else:
            bess_ret_pct = 100
            bess_capacity_year = bess_mwh
            replaced = ''
        
        yearly_data.append({
            'Year': year,
            'PV_MW': pv_capacity_year,
            'PV_Degradation_%': pv_deg_pct,
            'BESS_MWh': bess_capacity_year,
            'BESS_Retention_%': bess_ret_pct,
            'Status': replaced
        })
    
    df = pd.DataFrame(yearly_data)
    
    # Calculate BESS replacement cost (at year 21)
    if apply_bess and bess_mwh > 0:
        # Estimate replacement cost as 80% of original (no installation cost)
        replacement_cost_nominal = bess_npc * 0.8
        
        # Discount to present value (year 0)
        replacement_cost_pv = replacement_cost_nominal / ((1 + discount_rate) ** 20)
    else:
        replacement_cost_pv = 0
    
    # Total NPC with degradation
    npc_25y = npc_y1 + replacement_cost_pv
    
    # LCOE with degradation (simplified estimate)
    if apply_pv or apply_bess:
        # Energy production decreases due to degradation
        # Rough estimate: 5% increase in LCOE
        lcoe_25y = lcoe_y1 * 1.05
    else:
        lcoe_25y = lcoe_y1
    
    return {
        'yearly_df': df,
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
        }
    }


# ==============================================================================
# WRAPPER FUNCTIONS (for compatibility)
# ==============================================================================

# Make sure all base functions are available as exports
__all__ = [
    'grid_search_optimize_hydro',
    'run_degradation_analysis',
    'read_inputs',
    'calculate_dispatch_with_hydro',
    'calculate_npc_homer_style',
    'calculate_electrical_metrics',
    'find_optimal_solution',
    'PV_DEG',
    'BESS_CAP_RET'
]
```

## Step 3: Verify the file structure

Your directory should have:
```
/your_project_folder/
├── streamlit_app_with_degradation.py (or whatever your main app is named)
├── optimize_gridsearch_hydro_static_STREAMLITCHECK.py
└── optimize_with_degradation.py  ← NEW FILE (renamed from the old one)

