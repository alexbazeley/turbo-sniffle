"""
Debt financing calculation module.
Handles debt sizing, sculpting, DSCR, reserves (DSRA, O&M), and IDC.
"""

import pandas as pd
import numpy as np
from typing import Dict, Any, Tuple


class DebtModel:
    """Calculates debt financing, service, and coverage ratios."""

    def __init__(self, inputs: Dict[str, Any]):
        self.inputs = inputs
        self.financing = inputs['financing']
        self.project = inputs['project']

        self.use_debt = self.financing.get('use_debt', False)
        self.construction_months = self.project['construction_months']

    def size_debt(self, cfads_series: pd.Series, total_capex: float) -> float:
        """
        Size debt based on target DSCR or percentage of cost.

        Args:
            cfads_series: Cash Flow Available for Debt Service (operating years only)
            total_capex: Total capital expenditure

        Returns:
            Sized debt amount
        """
        if not self.use_debt:
            return 0.0

        sizing_method = self.financing.get('sizing_method', 'target_dscr')

        if sizing_method == 'target_dscr':
            # Size to minimum DSCR target
            target_dscr = self.financing.get('target_min_dscr', 1.30)
            interest_rate = self.financing.get('interest_rate_pct', 7.0) / 100.0 / 12.0  # Monthly
            tenor_months = self.financing.get('tenor_years', 15) * 12

            # Iterate to find debt amount that meets DSCR target
            # Use binary search or simplified approach
            # For now, estimate using average CFADS and payment

            avg_monthly_cfads = cfads_series.mean()

            # Target monthly payment = avg_cfads / target_dscr
            target_payment = avg_monthly_cfads / target_dscr

            # Calculate PV of payments = debt principal
            # Payment = Principal * (r * (1+r)^n) / ((1+r)^n - 1)
            # Rearrange: Principal = Payment * ((1+r)^n - 1) / (r * (1+r)^n)

            if interest_rate > 0:
                factor = ((1 + interest_rate) ** tenor_months - 1) / (
                    interest_rate * (1 + interest_rate) ** tenor_months
                )
                debt_amount = target_payment * factor
            else:
                debt_amount = target_payment * tenor_months

            # Cap at some % of total capex (e.g., 80%)
            max_debt = total_capex * 0.80
            debt_amount = min(debt_amount, max_debt)

        elif sizing_method == 'percent_of_cost':
            # Simple percentage
            debt_pct = self.financing.get('debt_pct_of_cost', 60.0) / 100.0
            debt_amount = total_capex * debt_pct

        else:
            debt_amount = 0.0

        return debt_amount

    def calculate_debt_service(self, debt_principal: float, total_months: int) -> pd.DataFrame:
        """
        Calculate debt service schedule.

        Args:
            debt_principal: Total debt amount
            total_months: Total periods in model

        Returns:
            DataFrame with debt service schedule
        """
        if not self.use_debt or debt_principal <= 0:
            # No debt
            df = pd.DataFrame({
                'period': range(total_months),
                'debt_balance': [0.0] * total_months,
                'interest_payment': [0.0] * total_months,
                'principal_payment': [0.0] * total_months,
                'total_debt_service': [0.0] * total_months
            })
            return df

        interest_rate_annual = self.financing.get('interest_rate_pct', 7.0) / 100.0
        interest_rate_monthly = interest_rate_annual / 12.0
        tenor_months = self.financing.get('tenor_years', 15) * 12
        amortization = self.financing.get('amortization', 'sculpted')

        # Debt service starts after construction (at COD)
        periods = list(range(total_months))
        debt_balance = []
        interest_payment = []
        principal_payment = []

        current_balance = debt_principal

        for period in periods:
            if period < self.construction_months:
                # Construction - no debt service yet
                debt_balance.append(current_balance)
                interest_payment.append(0.0)
                principal_payment.append(0.0)

            else:
                # Operating period
                month_in_debt_service = period - self.construction_months

                if month_in_debt_service < tenor_months:
                    # Calculate interest
                    interest = current_balance * interest_rate_monthly

                    if amortization == 'level':
                        # Level payment (annuity)
                        if interest_rate_monthly > 0:
                            payment = debt_principal * (
                                interest_rate_monthly * (1 + interest_rate_monthly) ** tenor_months
                            ) / ((1 + interest_rate_monthly) ** tenor_months - 1)
                        else:
                            payment = debt_principal / tenor_months

                        principal = payment - interest

                    elif amortization == 'sculpted':
                        # Simplified sculpting: equal principal payments
                        # (In practice, would iterate based on CFADS)
                        principal = debt_principal / tenor_months

                    else:
                        # Default to level
                        principal = debt_principal / tenor_months

                    # Update balance
                    current_balance -= principal

                    debt_balance.append(current_balance)
                    interest_payment.append(interest)
                    principal_payment.append(principal)

                else:
                    # Debt fully paid
                    debt_balance.append(0.0)
                    interest_payment.append(0.0)
                    principal_payment.append(0.0)

        df = pd.DataFrame({
            'period': periods,
            'debt_balance': debt_balance,
            'interest_payment': interest_payment,
            'principal_payment': principal_payment
        })

        df['total_debt_service'] = df['interest_payment'] + df['principal_payment']

        return df

    def calculate_reserves(self, monthly_opex_avg: float, monthly_debt_service_avg: float,
                          total_months: int) -> pd.DataFrame:
        """
        Calculate reserve accounts (DSRA, O&M reserve).

        Args:
            monthly_opex_avg: Average monthly OpEx
            monthly_debt_service_avg: Average monthly debt service
            total_months: Total periods

        Returns:
            DataFrame with reserve funding and releases
        """
        dsra_months = self.financing.get('dsra_months', 6)
        om_reserve_months = self.financing.get('om_reserve_months', 6)

        dsra_target = monthly_debt_service_avg * dsra_months
        om_reserve_target = monthly_opex_avg * om_reserve_months

        periods = list(range(total_months))
        dsra_funding = []
        dsra_balance = []
        om_funding = []
        om_balance = []

        # Fund reserves at COD
        for period in periods:
            if period == self.construction_months - 1:
                # At COD
                dsra_funding.append(dsra_target if self.use_debt else 0.0)
                om_funding.append(om_reserve_target)
                dsra_balance.append(dsra_target if self.use_debt else 0.0)
                om_balance.append(om_reserve_target)
            elif period < self.construction_months:
                dsra_funding.append(0.0)
                om_funding.append(0.0)
                dsra_balance.append(0.0)
                om_balance.append(0.0)
            else:
                # Operating - maintain balances (simplified)
                dsra_funding.append(0.0)
                om_funding.append(0.0)
                dsra_balance.append(dsra_target if self.use_debt else 0.0)
                om_balance.append(om_reserve_target)

        df = pd.DataFrame({
            'period': periods,
            'dsra_funding': dsra_funding,
            'dsra_balance': dsra_balance,
            'om_reserve_funding': om_funding,
            'om_reserve_balance': om_balance
        })

        return df

    def calculate_idc(self, capex_schedule: pd.DataFrame, debt_principal: float) -> float:
        """
        Calculate Interest During Construction.

        Args:
            capex_schedule: DataFrame with monthly CapEx
            debt_principal: Total debt amount

        Returns:
            Total IDC
        """
        if not self.use_debt or debt_principal <= 0:
            return 0.0

        interest_rate_annual = self.financing.get('interest_rate_pct', 7.0) / 100.0
        interest_rate_monthly = interest_rate_annual / 12.0

        idc_pct = self.inputs['capex'].get('idc_pct_of_debt_drawn', 100.0) / 100.0

        # Assume debt is drawn proportionally to CapEx during construction
        construction_capex = capex_schedule[capex_schedule['period'] < self.construction_months]
        total_construction_capex = construction_capex['total_capex'].sum()

        if total_construction_capex == 0:
            return 0.0

        # Calculate IDC month by month
        cumulative_capex = 0.0
        total_idc = 0.0

        for idx, row in construction_capex.iterrows():
            period_capex = row['total_capex']
            cumulative_capex += period_capex

            # Debt drawn proportionally
            debt_drawn = (cumulative_capex / total_construction_capex) * debt_principal

            # Interest on drawn debt
            monthly_interest = debt_drawn * interest_rate_monthly * idc_pct
            total_idc += monthly_interest

        return total_idc

    def calculate_dscr(self, cashflow_df: pd.DataFrame) -> pd.DataFrame:
        """
        Calculate Debt Service Coverage Ratio.

        Args:
            cashflow_df: DataFrame with cfads and total_debt_service

        Returns:
            Updated DataFrame with dscr column
        """
        dscr = []

        for idx, row in cashflow_df.iterrows():
            debt_service = row.get('total_debt_service', 0.0)
            cfads = row.get('cfads', 0.0)

            if debt_service > 0:
                dscr.append(cfads / debt_service)
            else:
                dscr.append(None)

        cashflow_df['dscr'] = dscr

        return cashflow_df
