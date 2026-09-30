# ⚡ Energy Modeling Optimizer v5.1

A professional web application for optimizing hybrid renewable energy systems:
**Solar PV + Wind + Hydro + Battery Energy Storage (BESS)**.

## 🎯 Features

- **Multi-source optimization:** Solar PV + Wind + Hydro + Battery Storage
- **4D grid search:** Exhaustive enumeration over discretized sizing combinations
  (preferred over metaheuristics for client-facing consulting work — transparent
  and auditable results)
- **Automatic hydro operating-window optimization**
- **HOMER Pro-style NPC/LCOE calculation** (real discount rate, CRF, replacement,
  salvage) alongside a **year-by-year discounted cash flow LCOE method**
  (nominal discount rate, inflation-escalated O&M, CAPEX at Year 0) — both
  available simultaneously for comparison
- **25-year multi-year degradation analysis** — independent degradation models
  for PV, Wind, Hydro, and BESS (simple annual-rate or uploaded CSV curve)
- **BESS deployment sizing** (Sungrow PowerTitan 2.0 container layout)
- **Interactive Plotly visualizations** and **one-click Excel export**

## 💻 Technology Stack

- **Frontend:** Streamlit
- **Backend:** Python 3.11
- **Optimization:** Custom 4D grid search (`optimize_gridsearch_hydro_WITH_DEGRADATION.py`)
- **Visualization:** Plotly
- **Data processing:** Pandas, NumPy
- **Excel I/O:** openpyxl

## 📁 Project Structure

```
.
├── streamlit_app_with_degradation.py          # Main Streamlit app (entry point)
├── optimize_gridsearch_hydro_WITH_DEGRADATION.py  # Optimization + financial engine
├── requirements.txt
├── .streamlit/
│   └── config.toml                            # Streamlit theme & server config
├── .devcontainer/
│   └── devcontainer.json                      # Codespaces / VS Code devcontainer
└── README.md
```

## 🎓 Usage

### Run locally

```bash
conda activate energy_opt        # or: python -m venv .venv && source .venv/bin/activate
pip install -r requirements.txt
streamlit run streamlit_app_with_degradation.py
```

### Run in the app

1. Configure component selection (PV / Wind / Hydro / BESS) and sizing ranges in the sidebar
2. Upload your 8760-hour Load, PV, Wind, and (optionally) Hydro profiles (CSV/Excel)
3. Set financial parameters (discount rate, inflation, project lifetime)
4. Optionally enable degradation analysis per component
5. Click **▶️ RUN OPTIMIZATION**
6. Review results across the **Summary**, **Cost & Performance**, **Economic Analysis**,
   and (if enabled) **Degradation** tabs
7. Export the full result set to Excel

## 📊 System Components

1. **☀️ Solar PV** — Photovoltaic solar panels
2. **💨 Wind** — Wind turbines
3. **💧 Hydro** — Hydroelectric power (with configurable daily operating window)
4. **🔋 BESS** — Battery Energy Storage System (Sungrow PowerTitan 2.0 deployment sizing)

## 🧮 Methodology Notes

- Component LCOE / LCOS = `(Component_NPC × CRF) / Annual_Energy` — **not**
  `NPC / (Energy × n)`, which understates costs by ignoring the time value of money.
- The Cash Flow LCOE tab implements the manager's Excel-based methodology:
  CAPEX at Year 0, O&M escalated annually with inflation, discounted at the
  **nominal** discount rate (not converted to real).
- Both LCOE methods are computed and displayed side-by-side for comparison.

## 👥 Team

Developed by SMEC (M) Sdn Bhd — part of the Surbana Jurong (SJ) group.

## 📄 License

Internal company tool — All rights reserved.
