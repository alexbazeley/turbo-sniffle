"""
Cashflow waterfall module.
Integrates all components into comprehensive cashflow model.
"""

import pandas as pd
import numpy as np
from typing import Dict, Any
from datetime import datetime


class CashflowModel:
    """Integrates all financial components into cashflow waterfall."""

    def __init__(self, inputs: Dict[str, Any]):
        self.inputs = inputs
        self.project = inputs['project']
        self.construction_months = self.project['construction_months']

    def build_cashflow(self, energy_df: pd.DataFrame, revenue_df: pd.DataFrame,
                      capex_df: pd.DataFrame, debt_df: pd.DataFrame,
                      depreciation_df: pd.DataFrame, reserves_df: pd.DataFrame,
                      itc_amount: float, upfront_grants: float,
                      upfront_grants_timing: str) -> pd.DataFrame:
        """
        Build integrated cashflow waterfall.

        Args:
            energy_df: Energy and OpEx data
            revenue_df: Revenue data (can be same as energy_df if already merged)
            capex_df: CapEx schedule
            debt_df: Debt service schedule
            depreciation_df: Depreciation schedule
            reserves_df: Reserve accounts
            itc_amount: Investment Tax Credit amount
            upfront_grants: Upfront grant amount
            upfront_grants_timing: 'NTP' or 'COD'

        Returns:
            Complete cashflow DataFrame
        """
        # Start with energy/revenue base
        cf = energy_df.copy()

        # Merge CapEx
        if 'total_capex' not in cf.columns:
            cf = cf.merge(capex_df[['period', 'total_capex']], on='period', how='left')

        # Merge debt
        for col in ['debt_balance', 'interest_payment', 'principal_payment', 'total_debt_service']:
            if col in debt_df.columns:
                cf[col] = debt_df[col]

        # Merge depreciation
        for col in ['depreciation_federal', 'depreciation_state']:
            if col in depreciation_df.columns:
                cf[col] = depreciation_df[col]

        # Merge reserves
        for col in ['dsra_funding', 'om_reserve_funding']:
            if col in reserves_df.columns:
                cf[col] = reserves_df[col]

        # Fill NaN with 0
        cf = cf.fillna(0.0)

        # Add ITC (received at COD)
        cf['itc_credit'] = 0.0
        if itc_amount > 0:
            cod_period = self.construction_months - 1
            if cod_period >= 0 and cod_period < len(cf):
                cf.loc[cod_period, 'itc_credit'] = itc_amount

        # Add upfront grants
        cf['upfront_grants'] = 0.0
        if upfront_grants > 0:
            if upfront_grants_timing == 'NTP':
                grant_period = 0
            else:  # COD
                grant_period = self.construction_months - 1

            if grant_period >= 0 and grant_period < len(cf):
                cf.loc[grant_period, 'upfront_grants'] = upfront_grants

        # Calculate EBITDA
        # Revenue - OpEx (excluding debt service, depreciation, taxes)
        cf['ebitda'] = cf['total_revenue'] - cf['total_opex']

        # Calculate EBIT (EBITDA - Depreciation)
        cf['ebit'] = cf['ebitda'] - cf['depreciation_federal']

        # Calculate Taxable Income
        # EBIT - Interest
        cf['taxable_income'] = cf['ebit'] - cf['interest_payment']

        # Taxes will be calculated in tax module and added back
        # For now, placeholder
        cf['total_tax'] = 0.0

        # Calculate CFADS (Cash Flow Available for Debt Service)
        # EBITDA - Taxes - non-debt CapEx - reserve funding + grants/credits
        cf['cfads'] = (cf['ebitda'] - cf['total_tax'] +
                      cf['itc_credit'] + cf['upfront_grants'] -
                      cf['dsra_funding'] - cf['om_reserve_funding'])

        # Calculate Free Cashflow to Equity
        # CFADS - Debt Service + Debt Drawdown - CapEx
        cf['debt_drawdown'] = 0.0

        # Set debt drawdown during construction
        if self.inputs['financing'].get('use_debt', False):
            # Simplified: all debt drawn at COD
            total_debt = debt_df['debt_balance'].max()
            if self.construction_months > 0:
                cf.loc[self.construction_months - 1, 'debt_drawdown'] = total_debt

        cf['fcfe'] = (cf['cfads'] - cf['total_debt_service'] +
                     cf['debt_drawdown'] - cf['total_capex'])

        # Equity contribution (negative FCFE during construction)
        cf['equity_contribution'] = 0.0
        cf['equity_distribution'] = 0.0

        for idx, row in cf.iterrows():
            if row['fcfe'] < 0:
                cf.loc[idx, 'equity_contribution'] = abs(row['fcfe'])
                cf.loc[idx, 'equity_distribution'] = 0.0
            else:
                cf.loc[idx, 'equity_contribution'] = 0.0
                cf.loc[idx, 'equity_distribution'] = row['fcfe']

        return cf

    def add_terminal_value(self, cf: pd.DataFrame) -> pd.DataFrame:
        """Add terminal value at exit years."""
        exit_config = self.inputs.get('exit', {})
        exit_years = exit_config.get('buyout_years', [])
        exit_multiple = exit_config.get('exit_multiple', 8.0)

        cf['terminal_value'] = 0.0

        for exit_year in exit_years:
            exit_period = self.construction_months + (exit_year * 12) - 1

            if exit_period < len(cf):
                # Calculate terminal value based on that year's EBITDA
                year_start = self.construction_months + ((exit_year - 1) * 12)
                year_end = year_start + 12

                if year_end <= len(cf):
                    annual_ebitda = cf.loc[year_start:year_end-1, 'ebitda'].sum()
                    terminal_value = annual_ebitda * exit_multiple

                    # Add to the last month of that year
                    cf.loc[exit_period, 'terminal_value'] = terminal_value

        return cf

    def format_output(self, cf: pd.DataFrame) -> pd.DataFrame:
        """Format cashflow for output with proper column ordering."""
        # Define column order
        columns_order = [
            'period', 'date', 'month_in_operation', 'year_in_operation',
            # Energy
            'ac_kwh', 'ac_mwh',
            # Revenue
            'total_revenue',
            # OpEx
            'total_opex',
            # CapEx
            'total_capex',
            # Debt
            'debt_drawdown', 'debt_balance', 'interest_payment',
            'principal_payment', 'total_debt_service',
            # Tax
            'depreciation_federal', 'taxable_income', 'total_tax',
            'itc_credit', 'upfront_grants',
            # Metrics
            'ebitda', 'ebit', 'cfads', 'fcfe',
            # Equity
            'equity_contribution', 'equity_distribution',
            # Terminal
            'terminal_value'
        ]

        # Include columns that exist
        output_cols = [col for col in columns_order if col in cf.columns]

        return cf[output_cols]
