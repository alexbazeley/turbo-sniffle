"""
Tax credits and depreciation calculation module.
Handles ITC, PTC, MACRS depreciation, basis reduction, and NOL carryforward.
"""

import pandas as pd
import numpy as np
from typing import Dict, Any, Tuple


class TaxModel:
    """Calculates tax credits, depreciation, and tax liabilities."""

    # MACRS 5-year schedule (half-year convention)
    MACRS_5YR = [0.2000, 0.3200, 0.1920, 0.1152, 0.1152, 0.0576]

    # MACRS 7-year schedule
    MACRS_7YR = [0.1429, 0.2449, 0.1749, 0.1249, 0.0893, 0.0892, 0.0893, 0.0446]

    def __init__(self, inputs: Dict[str, Any]):
        self.inputs = inputs
        self.tax_credits = inputs['tax_credits']
        self.taxes = inputs['taxes']
        self.project = inputs['project']

        self.federal_rate = self.taxes['federal_tax_rate_pct'] / 100.0
        self.state_rate = self.taxes['state_tax_rate_pct'] / 100.0
        self.combined_rate = self.federal_rate + self.state_rate * (1 - self.federal_rate)

    def calculate_itc(self, depreciable_basis: float, upfront_grants: float = 0.0) -> Tuple[float, float]:
        """
        Calculate Investment Tax Credit.

        Args:
            depreciable_basis: Total depreciable basis before ITC reduction
            upfront_grants: Upfront grants (if basis reduction applies)

        Returns:
            (itc_amount, reduced_basis_for_depreciation)
        """
        if self.tax_credits['mode'] != 'ITC':
            return 0.0, depreciable_basis

        # ITC basis (may need to exclude land, other non-eligible items)
        itc_basis = depreciable_basis

        # Reduce basis if grants received and flag is set
        if upfront_grants > 0 and self.tax_credits.get('reduce_itc_basis_for_grants', False):
            itc_basis -= upfront_grants

        # Calculate ITC
        base_itc_pct = self.tax_credits['itc_pct'] / 100.0
        adders_pct = self.tax_credits.get('adders_pct', 0.0) / 100.0
        total_itc_pct = base_itc_pct + adders_pct

        itc_amount = itc_basis * total_itc_pct

        # Basis reduction for depreciation
        basis_reduction_pct = self.tax_credits['basis_reduction_pct'] / 100.0
        reduced_basis = depreciable_basis - (itc_amount * basis_reduction_pct)

        return itc_amount, reduced_basis

    def calculate_ptc(self, energy_df: pd.DataFrame) -> pd.DataFrame:
        """
        Calculate Production Tax Credit by period.

        Args:
            energy_df: DataFrame with energy production

        Returns:
            DataFrame with ptc_amount column
        """
        if self.tax_credits['mode'] != 'PTC':
            energy_df['ptc_amount'] = 0.0
            return energy_df

        ptc_usd_per_mwh = self.tax_credits.get('ptc_usd_per_mwh', 27.5)
        ptc_term_years = self.tax_credits.get('ptc_term_years', 10)
        ptc_term_months = ptc_term_years * 12

        ptc_amounts = []

        for idx, row in energy_df.iterrows():
            month_in_op = row['month_in_operation']

            if month_in_op > 0 and month_in_op <= ptc_term_months:
                ptc = row['ac_mwh'] * ptc_usd_per_mwh
                ptc_amounts.append(ptc)
            else:
                ptc_amounts.append(0.0)

        energy_df['ptc_amount'] = ptc_amounts
        return energy_df

    def calculate_depreciation(self, depreciable_basis: float, total_months: int) -> pd.DataFrame:
        """
        Calculate MACRS depreciation schedule.

        Args:
            depreciable_basis: Basis for depreciation (after ITC reduction)
            total_months: Total periods in model

        Returns:
            DataFrame with period, depreciation_federal, depreciation_state
        """
        # Get schedule
        schedule_name = self.taxes['depr_schedule']
        if schedule_name == 'MACRS_5yr':
            schedule = self.MACRS_5YR
        elif schedule_name == 'MACRS_7yr':
            schedule = self.MACRS_7YR
        else:
            # Default to 5-year
            schedule = self.MACRS_5YR

        # Bonus depreciation
        bonus_pct = self.tax_credits.get('bonus_depreciation_pct', 0.0) / 100.0

        # Calculate annual depreciation
        annual_depr = []

        if bonus_pct > 0:
            # Bonus in year 1, then MACRS on remaining
            bonus_amount = depreciable_basis * bonus_pct
            remaining_basis = depreciable_basis - bonus_amount

            annual_depr.append(bonus_amount + remaining_basis * schedule[0])
            for i in range(1, len(schedule)):
                annual_depr.append(remaining_basis * schedule[i])
        else:
            # Standard MACRS
            for rate in schedule:
                annual_depr.append(depreciable_basis * rate)

        # Convert to monthly
        # Depreciation starts at COD (month = construction_months in the main timeline)
        construction_months = self.project['construction_months']

        periods = list(range(total_months))
        monthly_depr_fed = []
        monthly_depr_state = []

        for period in periods:
            if period < construction_months:
                # No depreciation during construction
                monthly_depr_fed.append(0.0)
                monthly_depr_state.append(0.0)
            else:
                # Operating period
                month_in_op = period - construction_months
                year_in_op = month_in_op // 12

                if year_in_op < len(annual_depr):
                    # Allocate annual depreciation evenly across 12 months
                    monthly_fed = annual_depr[year_in_op] / 12.0
                    monthly_depr_fed.append(monthly_fed)

                    # State depreciation (could be different schedule if specified)
                    monthly_depr_state.append(monthly_fed)
                else:
                    monthly_depr_fed.append(0.0)
                    monthly_depr_state.append(0.0)

        df = pd.DataFrame({
            'period': periods,
            'depreciation_federal': monthly_depr_fed,
            'depreciation_state': monthly_depr_state
        })

        return df

    def calculate_taxes(self, cashflow_df: pd.DataFrame) -> pd.DataFrame:
        """
        Calculate tax liability with NOL carryforward.

        Args:
            cashflow_df: DataFrame with taxable_income column

        Returns:
            Updated DataFrame with tax columns
        """
        nol_carryforward_years = self.taxes.get('nol_carryforward_years', 20)

        nol_balance = 0.0
        federal_tax = []
        state_tax = []
        nol_utilized = []
        nol_balance_list = []

        for idx, row in cashflow_df.iterrows():
            taxable_income = row.get('taxable_income', 0.0)

            if taxable_income < 0:
                # Loss - add to NOL
                nol_balance += abs(taxable_income)
                federal_tax.append(0.0)
                state_tax.append(0.0)
                nol_utilized.append(0.0)
            else:
                # Income - use NOL if available
                if nol_balance > 0:
                    nol_used = min(nol_balance, taxable_income)
                    nol_balance -= nol_used
                    taxable_after_nol = taxable_income - nol_used
                    nol_utilized.append(nol_used)
                else:
                    taxable_after_nol = taxable_income
                    nol_utilized.append(0.0)

                # Calculate tax
                fed_tax = taxable_after_nol * self.federal_rate
                st_tax = taxable_after_nol * self.state_rate

                federal_tax.append(fed_tax)
                state_tax.append(st_tax)

            nol_balance_list.append(nol_balance)

        cashflow_df['federal_tax'] = federal_tax
        cashflow_df['state_tax'] = state_tax
        cashflow_df['total_tax'] = cashflow_df['federal_tax'] + cashflow_df['state_tax']
        cashflow_df['nol_utilized'] = nol_utilized
        cashflow_df['nol_balance'] = nol_balance_list

        return cashflow_df
