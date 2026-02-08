"""
STREAMLIT APP UPDATES FOR DEGRADATION ANALYSIS
===============================================
These sections need to be updated in your main Streamlit app file.

INSTRUCTIONS:
1. Replace the degradation import (around line 20)
2. Update the optimization execution (around line 760-900)
3. Update the Excel export function (around line 150)
"""

# ==============================================================================
# SECTION 1: UPDATE IMPORT (Replace around line 20)
# ==============================================================================

# OLD CODE:
"""
try:
    import optimize_with_degradation as deg_module
    DEGRADATION_AVAILABLE = True
except ImportError:
    DEGRADATION_AVAILABLE = False
"""

# NEW CODE:
try:
    import optimize_with_degradation_FIXED as deg_module
    DEGRADATION_AVAILABLE = True
except ImportError:
    DEGRADATION_AVAILABLE = False
    print("⚠️ Warning: Degradation module not found")


# ==============================================================================
# SECTION 2: UPDATE EXCEL EXPORT FUNCTION (Replace the entire function)
# ==============================================================================

def export_results_industry_format_WITH_DEGRADATION(results_dict, results_df, optimal_row, 
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
# SECTION 3: UPDATE OPTIMIZATION EXECUTION (Replace the optimization section)
# ==============================================================================

# This goes in the "RUN OPTIMIZATION" button callback
# Find the section around line 760-900 and replace with this:

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
                    degradation_results = deg_module.run_degradation_analysis_complete(
                        optimal.to_dict(),
                        config_for_deg,
                        profiles,
                        apply_pv=apply_pv_degradation,
                        apply_bess=apply_bess_degradation,
                        years_to_export=[1, 2, 5, 10, 15, 20, 25]
                    )
                    
                    progress_bar.progress(90)
                    
                    # Use Year 1 hourly dispatch from degradation analysis for display
                    if 'year_1' in degradation_results['hourly_dispatch']:
                        optimal_dispatch = degradation_results['hourly_dispatch']['year_1']
                    else:
                        # Fallback to standard dispatch
                        optimal_dispatch = deg_module.calculate_dispatch_with_hydro(
                            load_profile, pvsyst_profile, wind_profile,
                            optimal['PV_kW'], optimal['Wind_kW'], optimal['Hydro_kW'],
                            optimal['BESS_Power_kW'], optimal['BESS_Capacity_kWh'],
                            solar, wind, hydro, bess,
                            int(optimal['Hydro_Window_Start']), int(optimal['Hydro_Window_End'])
                        )
                else:
                    degradation_results = None
                    optimal_dispatch = None
                
            else:
                # Standard optimization (no degradation)
                import optimize_gridsearch_hydro_static_STREAMLITCHECK as opt_module
                
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


# ==============================================================================
# SECTION 4: UPDATE DOWNLOAD BUTTON (in Results tab)
# ==============================================================================

# Find the download button section and replace with:

st.subheader("📥 Download Results")

# Get degradation results if available
degradation_results = results.get('degradation_results', None)

# Use the updated export function
excel_output = export_results_industry_format_WITH_DEGRADATION(
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

# Add degradation summary display if available
if degradation_results:
    st.markdown("---")
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
