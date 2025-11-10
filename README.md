# Solar Financial Model

Production-grade, auditable financial model for utility-scale and distributed solar projects with support for **Community Solar (CDG)** and **Power Purchase Agreement (PPA)** commercial modes.

## Features

- **Dual Commercial Modes**: Community Solar and PPA with configurable parameters
- **State Program Support**: NY-Sun/VDER, NJ CSEP, IL Adjustable Block Program
- **Comprehensive Tax Modeling**: ITC/PTC with adders, MACRS depreciation, NOL carryforward
- **Debt Financing**: DSCR-based sizing, sculpted/level amortization, reserves (DSRA, O&M)
- **Complete Metrics**: Pre-tax/post-tax equity IRR, project IRR, NPV, LCOE, DSCR, payback
- **Excel Output**: 15-tab workbook with formulas, charts, and audit trail
- **Monthly Granularity**: From Notice-to-Proceed through 30-35 year operating life

## Quick Start

### Installation

```bash
# Install dependencies
pip install -r requirements.txt

# Or install as package
pip install -e .
```

### Interactive Excel Model (Recommended)

**Best for**: Finance professionals who want to edit assumptions directly in Excel

```bash
python run_formula_model.py
```

This creates `SolarModel_Interactive.xlsx` with:
- ✅ **Editable inputs** in yellow cells on Dashboard tab
- ✅ **Real-time updates** - Key metrics recalculate instantly
- ✅ **Excel formulas** - All calculations visible and auditable
- ✅ **Dropdowns** for categorical inputs (mode, etc.)
- ✅ **No re-runs needed** - Change assumptions and see results immediately

**Usage:**
1. Open `SolarModel_Interactive.xlsx`
2. Go to "Dashboard" tab
3. Edit yellow input cells (COD date, capacity, prices, etc.)
4. Green metric cells update automatically
5. Review detailed calculations in other tabs

### Python-Calculated Model

**Best for**: Programmatic scenarios, batch runs, or when you need Python flexibility

```python
from model import run_model

# Run model with JSON inputs
cashflow, metrics = run_model('example_inputs.json', 'SolarModel.xlsx')

print(f"Equity IRR: {metrics['pre_tax_irr']*100:.2f}%")
print(f"Project IRR: {metrics['project_irr']*100:.2f}%")
print(f"Min DSCR: {metrics['min_dscr']:.2f}")
```

Or from command line:

```bash
python -c "from model import run_model; run_model('example_inputs.json')"
```

## Input Structure

The model accepts a JSON configuration file. See `example_inputs.json` for complete schema.

### Key Sections

1. **Project**: Mode (CDG/PPA), COD date, model horizon, discount rate
2. **Sizing**: DC/AC capacity, performance parameters, degradation
3. **Revenue**: Mode-specific parameters (subscriber discount, PPA price, etc.)
4. **Tax Credits**: ITC/PTC with adders and basis reduction
5. **CapEx**: Equipment costs, developer fee, contingency
6. **OpEx**: Fixed/variable O&M, land, property tax
7. **Financing**: Debt sizing, interest rate, tenor, reserves
8. **Program**: State-specific incentives (NY-Sun, NJ CSEP, IL ABP)

## Output Structure

Excel workbook with 15 tabs:

- **Assumptions**: High-level summary
- **Inputs_Data**: Complete input JSON
- **Energy_8760_or_Monthly**: Production schedule
- **Revenue**: Revenue by source
- **CapEx**: Capital expenditure schedule
- **OpEx**: Operating expense breakdown
- **Debt_Finance**: Debt service and DSCR
- **Taxes_Depreciation**: Tax calculations
- **Cashflow_Waterfall**: Integrated cashflow (EBITDA → FCFE)
- **IRR_NPV_Metrics**: Financial metrics summary
- **Sensitivities**: Sensitivity analysis
- **Scenarios**: Scenario comparison
- **Charts**: Financial charts
- **Audit_Trace**: Defaults and warnings
- **Notes**: Methodology documentation

## Key Metrics

- **Equity IRR** (pre-tax and post-tax)
- **Project IRR** (unlevered)
- **NPV** at user discount rate
- **LCOE** (nominal and real, $/MWh)
- **DSCR** (min, avg, year 1)
- **Payback** period

## Testing

```bash
# Run all tests
pytest tests/

# With coverage
pytest --cov=model tests/

# Specific test
pytest tests/test_tax.py -v
```

## State Program Support

### NY-Sun / VDER

- Value stack: Energy, Environmental, ICAP, DRV, LSRV
- Community Credit ($/kWh)
- ICSA and other upfront incentives

### NJ Community Solar Energy Pilot

- Blended bill credit rates by customer class
- LMI/non-LMI/anchor subscriber discounts
- Program adders (siting, public entity, etc.)

### IL Adjustable Block Program

- REC price and term
- Front-loaded payment schedule
- Brownfield/site adders
- Smart inverter rebates

## Methodology

### Energy
Monthly production with linear degradation, availability, and curtailment derates.

### Revenue
- **CDG**: Bill credit × (1 - discount) × subscription level, adjusted for churn and bad debt
- **PPA**: Energy × PPA price (escalating), merchant tail after contract

### Tax
- ITC at COD with basis reduction
- PTC per MWh over 10 years
- MACRS depreciation (5-year default)
- NOL carryforward

### Debt
Sized to target DSCR or % of cost. Sculpted or level amortization. DSRA and O&M reserves funded at COD.

### Cashflow
```
Revenue - OpEx = EBITDA
EBITDA - Depreciation - Interest = Taxable Income
Taxable Income - Tax + Credits = CFADS
CFADS - Debt Service - CapEx = FCFE
```

## Assumptions & Limitations

**Assumptions:**
- USD currency
- Monthly periods
- Constant tax rates
- Unlimited NOL carryforward
- State program values are user inputs (not fetched)

**Limitations:**
- No tax equity partnership modeling
- Simplified merchant pricing (uses last PPA price)
- No dynamic debt sculpting
- Subscriber churn is simplified

## Troubleshooting

**IRR returns NaN**: Ensure project has positive cashflows

**DSCR too low**: Increase target or reduce tenor

**State program revenue missing**: Check `program.active` matches enabled program

## Contributing

Enhancements welcome:
- Tax equity structures
- Dynamic debt sculpting
- Advanced subscriber modeling
- Real-time tariff fetching
- Google Sheets export
- Monte Carlo simulation

## License

MIT License

## Version

v1.0.0 - Initial release with CDG/PPA modes and state program support