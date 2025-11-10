"""
Main runner module.
Orchestrates all model components and generates outputs.
"""

import pandas as pd
import numpy as np
from typing import Dict, Any, Tuple
import json

from .inputs import load_inputs
from .energy import EnergyModel
from .revenue import RevenueModel
from .capex import CapExModel
from .opex import OpExModel
from .debt import DebtModel
from .tax import TaxModel
from .cashflow import CashflowModel
from .metrics import MetricsCalculator
from .writer_excel import ExcelWriter


class SolarFinancialModel:
    """Main solar financial model orchestrator."""

    def __init__(self, inputs_path: str):
        """
        Initialize model with inputs.

        Args:
            inputs_path: Path to inputs JSON file
        """
        self.inputs_path = inputs_path
        self.inputs, self.defaults_used, self.warnings = load_inputs(inputs_path)

        # Extract key parameters
        self.sizing = self.inputs['sizing']
        self.ac_kw = self.sizing['ac_kw']
        self.dc_kw = self.sizing['dc_kw']

        # Initialize sub-models
        self.energy_model = EnergyModel(self.inputs)
        self.revenue_model = RevenueModel(self.inputs, self.ac_kw)
        self.capex_model = CapExModel(self.inputs)
        self.opex_model = OpExModel(self.inputs, self.ac_kw)
        self.debt_model = DebtModel(self.inputs)
        self.tax_model = TaxModel(self.inputs)
        self.cashflow_model = CashflowModel(self.inputs)
        self.metrics_calc = MetricsCalculator(self.inputs)

        # Results storage
        self.energy_df = None
        self.revenue_df = None
        self.capex_df = None
        self.opex_df = None
        self.debt_df = None
        self.depreciation_df = None
        self.cashflow_df = None
        self.metrics = None

    def run(self) -> Tuple[pd.DataFrame, Dict[str, Any]]:
        """
        Run complete financial model.

        Returns:
            (cashflow_dataframe, metrics_dict)
        """
        print("Running Solar Financial Model...")

        # Step 1: Calculate energy production
        print("  1. Calculating energy production...")
        self.energy_df = self.energy_model.calculate_monthly_energy()

        # Step 2: Calculate OpEx
        print("  2. Calculating operating expenses...")
        self.energy_df = self.opex_model.calculate_monthly_opex(self.energy_df)
        self.energy_df = self.opex_model.calculate_land_costs(self.energy_df)
        self.energy_df = self.opex_model.calculate_property_tax(self.energy_df)
        self.energy_df = self.opex_model.calculate_inverter_replacements(self.energy_df)
        self.energy_df = self.opex_model.calculate_total_opex(self.energy_df)

        # Step 3: Calculate revenue
        print("  3. Calculating revenue...")
        self.revenue_df = self.revenue_model.calculate_revenue(self.energy_df)

        # Step 4: Calculate CapEx
        print("  4. Calculating capital expenditures...")
        total_months = len(self.energy_df)
        self.capex_df = self.capex_model.create_capex_schedule(total_months)

        # Get upfront grants
        upfront_grants, grants_timing = self.revenue_model.get_upfront_grants()

        # Step 5: Calculate tax credits and depreciation
        print("  5. Calculating tax credits and depreciation...")
        depreciable_basis_before_itc = self.capex_model.get_depreciable_basis()
        itc_amount, depreciable_basis_after_itc = self.tax_model.calculate_itc(
            depreciable_basis_before_itc, upfront_grants
        )

        self.depreciation_df = self.tax_model.calculate_depreciation(
            depreciable_basis_after_itc, total_months
        )

        # Calculate PTC if applicable
        self.revenue_df = self.tax_model.calculate_ptc(self.revenue_df)

        # Step 6: Calculate debt (needs initial CFADS estimate)
        print("  6. Calculating debt financing...")

        # Build preliminary cashflow for debt sizing
        prelim_cf = self.revenue_df.copy()
        prelim_cf['cfads_estimate'] = prelim_cf['total_revenue'] - prelim_cf['total_opex']

        # Size debt based on operating period CFADS
        operating_cfads = prelim_cf[prelim_cf['month_in_operation'] > 0]['cfads_estimate']

        total_capex_before_idc, _ = self.capex_model.calculate_total_capex(include_idc=False)
        debt_principal = self.debt_model.size_debt(operating_cfads, total_capex_before_idc)

        # Calculate debt service
        self.debt_df = self.debt_model.calculate_debt_service(debt_principal, total_months)

        # Calculate IDC
        idc_amount = self.debt_model.calculate_idc(self.capex_df, debt_principal)

        # Update total CapEx with IDC
        self.capex_df.loc[self.inputs['project']['construction_months'] - 1, 'total_capex'] += idc_amount

        # Calculate reserves
        avg_opex = self.revenue_df[self.revenue_df['month_in_operation'] > 0]['total_opex'].mean()
        avg_debt_service = self.debt_df[self.debt_df['total_debt_service'] > 0]['total_debt_service'].mean()
        reserves_df = self.debt_model.calculate_reserves(avg_opex, avg_debt_service, total_months)

        # Step 7: Build integrated cashflow
        print("  7. Building cashflow waterfall...")
        self.cashflow_df = self.cashflow_model.build_cashflow(
            self.revenue_df, self.revenue_df, self.capex_df,
            self.debt_df, self.depreciation_df, reserves_df,
            itc_amount, upfront_grants, grants_timing
        )

        # Step 8: Calculate taxes with NOL
        print("  8. Calculating taxes...")
        self.cashflow_df = self.tax_model.calculate_taxes(self.cashflow_df)

        # Recalculate CFADS and FCFE with actual taxes
        self.cashflow_df['cfads'] = (self.cashflow_df['ebitda'] -
                                     self.cashflow_df['total_tax'] +
                                     self.cashflow_df['itc_credit'] +
                                     self.cashflow_df['upfront_grants'] -
                                     self.cashflow_df.get('dsra_funding', 0) -
                                     self.cashflow_df.get('om_reserve_funding', 0))

        self.cashflow_df['fcfe'] = (self.cashflow_df['cfads'] -
                                    self.cashflow_df['total_debt_service'] +
                                    self.cashflow_df['debt_drawdown'] -
                                    self.cashflow_df['total_capex'])

        # Recalculate equity contributions and distributions
        for idx, row in self.cashflow_df.iterrows():
            if row['fcfe'] < 0:
                self.cashflow_df.loc[idx, 'equity_contribution'] = abs(row['fcfe'])
                self.cashflow_df.loc[idx, 'equity_distribution'] = 0.0
            else:
                self.cashflow_df.loc[idx, 'equity_contribution'] = 0.0
                self.cashflow_df.loc[idx, 'equity_distribution'] = row['fcfe']

        # Step 9: Calculate DSCR
        self.cashflow_df = self.debt_model.calculate_dscr(self.cashflow_df)

        # Step 10: Add terminal value
        self.cashflow_df = self.cashflow_model.add_terminal_value(self.cashflow_df)

        # Update equity distribution with terminal value
        self.cashflow_df['equity_distribution'] += self.cashflow_df['terminal_value']

        # Step 11: Calculate metrics
        print("  9. Calculating financial metrics...")
        self.metrics = self.metrics_calc.calculate_all_metrics(self.cashflow_df)

        print("Model run complete!")
        print(f"  Equity Pre-Tax IRR: {self.metrics['pre_tax_irr']*100:.2f}%")
        print(f"  Project IRR: {self.metrics['project_irr']*100:.2f}%")
        print(f"  NPV: ${self.metrics['equity_npv']:,.0f}")
        print(f"  Min DSCR: {self.metrics.get('min_dscr', 'N/A')}")
        print(f"  Nominal LCOE: ${self.metrics['nominal_lcoe']:.2f}/MWh")

        return self.cashflow_df, self.metrics

    def export_to_excel(self, output_path: str):
        """
        Export model to Excel workbook.

        Args:
            output_path: Path for output Excel file
        """
        print(f"Exporting to Excel: {output_path}")

        writer = ExcelWriter(
            self.inputs, self.cashflow_df, self.metrics,
            self.defaults_used, self.warnings
        )

        writer.write_workbook(output_path)

        print("Export complete!")


def run_model(inputs_path: str, output_path: str = "SolarModel.xlsx") -> Tuple[pd.DataFrame, Dict[str, Any]]:
    """
    Convenience function to run model and export to Excel.

    Args:
        inputs_path: Path to inputs JSON
        output_path: Path for output Excel (default: SolarModel.xlsx)

    Returns:
        (cashflow_dataframe, metrics_dict)
    """
    model = SolarFinancialModel(inputs_path)
    cashflow, metrics = model.run()
    model.export_to_excel(output_path)

    return cashflow, metrics
