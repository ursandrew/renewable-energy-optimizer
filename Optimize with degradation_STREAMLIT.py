"""
DEGRADATION ANALYSIS MODULE FOR STREAMLIT
==========================================
Lightweight degradation engine for PV + BESS systems

Author: SJ
Version: 1.0
"""

import pandas as pd
import numpy as np

# OEM Degradation Data
PV_DEG = {1:0, 2:0.41, 3:0.82, 4:1.22, 5:1.63, 6:2.04, 7:2.45, 8:2.86, 9:3.27, 10:3.67, 
          11:4.08, 12:4.49, 13:4.90, 14:5.31, 15:5.71, 16:6.12, 17:6.53, 18:6.94, 19:7.35, 20:7.76,
          21:8.16, 22:8.57, 23:8.98, 24:9.39, 25:9.80}

BESS_CAP_RET = {1:94.46, 2:92.14, 3:90.33, 4:88.71, 5:87.19, 6:85.75, 7:84.37, 8:83.00, 9:81.70, 10:80.45,
                11:79.20, 12:78.00, 13:76.76, 14:75.57, 15:74.40, 16:73.23, 17:72.10, 18:70.96, 19:69.84, 20:68.73}

def run_degradation_analysis(optimal_row, config_params, hourly_data=None):
    """Simplified degradation analysis compatible with existing Streamlit results."""
    
    pv_mw = optimal_row.get('PV Capacity', 0)
    bess_mw = optimal_row.get('BESS Power', 0) 
    bess_mwh = optimal_row.get('BESS Energy', 0)
    duration = bess_mwh / bess_mw if bess_mw > 0 else 4
    
    npc_y1 = optimal_row.get('NPC', 0)
    lcoe_y1 = optimal_row.get('LCOE', 0)
    
    # Get parameters
    discount_rate = config_params.get('discount_rate', 0.08) / 100 if config_params.get('discount_rate', 8) > 1 else config_params.get('discount_rate', 0.08)
    
    # Simulate 25 years (simplified)
    yearly = []
    for year in range(1, 26):
        age = year if year <= 20 else (year - 20)
        pv_deg = PV_DEG.get(year, 10)
        bess_ret = BESS_CAP_RET.get(age, 70)
        
        yearly.append({
            'Year': year,
            'PV_MW': pv_mw * (1 - pv_deg/100),
            'PV_Deg_%': pv_deg,
            'BESS_MWh': bess_mwh * (bess_ret/100),
            'BESS_Ret_%': bess_ret,
            'Replaced': '🔋' if year == 21 else ''
        })
    
    df = pd.DataFrame(yearly)
    
    # BESS replacement cost
    bess_npc = optimal_row.get('BESS NPC', 0)
    if bess_npc == 0:  # Estimate if not available
        bess_npc = (bess_mw * 300 + bess_mwh * 200) * 12  # Rough estimate
    
    repl_cost = bess_npc * 0.8 / ((1 + discount_rate) ** 20)
    npc_25y = npc_y1 + repl_cost
    lcoe_25y = lcoe_y1 * 1.05  # Simplified estimate
    
    return {
        'yearly_df': df,
        'npc_year1': npc_y1,
        'npc_25year': npc_25y,
        'replacement_cost': repl_cost,
        'lcoe_year1': lcoe_y1,
        'lcoe_25year': lcoe_25y,
        'pv_deg_total': PV_DEG[25],
        'bess_loss_20y': 100 - BESS_CAP_RET[20]
    }