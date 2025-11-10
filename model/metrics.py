"""
Financial metrics calculation module.
Calculates IRR, NPV, LCOE, DSCR, payback, and other key metrics.
"""

import pandas as pd
import numpy as np
from typing import Dict, Any, Tuple
from scipy.optimize import newton


class MetricsCalculator:
    """Calculates financial metrics from cashflow model."""

    def __init__(self, inputs: Dict[str, Any]):
        self.inputs = inputs
        self.project = inputs['project']
        self.discount_rate = self.project['discount_rate_nominal_pct'] / 100.0
        self.inflation_rate = self.project.get('inflation_pct', 2.5) / 100.0

    def calculate_xirr(self, dates: pd.Series, cashflows: pd.Series) -> float:
        """
        Calculate XIRR (Internal Rate of Return with irregular periods).

        Args:
            dates: Series of dates
            cashflows: Series of cashflows

        Returns:
            IRR as decimal (e.g., 0.12 for 12%)
        """
        # Convert dates to years from first date
        first_date = dates.iloc[0]
        years = [(d - first_date).days / 365.25 for d in dates]

        # Define NPV function
        def npv(rate):
            return sum(cf / (1 + rate) ** yr for cf, yr in zip(cashflows, years))

        # Solve for rate where NPV = 0
        try:
            irr = newton(npv, 0.1, maxiter=100)
            return irr
        except:
            # If Newton's method fails, return NaN
            return np.nan

    def calculate_xnpv(self, dates: pd.Series, cashflows: pd.Series, rate: float) -> float:
        """
        Calculate XNPV (Net Present Value with irregular periods).

        Args:
            dates: Series of dates
            cashflows: Series of cashflows
            rate: Discount rate (decimal)

        Returns:
            NPV
        """
        first_date = dates.iloc[0]
        npv = sum(cf / (1 + rate) ** ((d - first_date).days / 365.25)
                 for cf, d in zip(cashflows, dates))
        return npv

    def calculate_equity_irr(self, cf: pd.DataFrame) -> Dict[str, float]:
        """
        Calculate equity IRRs (pre-tax and post-tax).

        Args:
            cf: Cashflow DataFrame with equity_contribution and equity_distribution

        Returns:
            Dict with pre_tax_irr and post_tax_irr
        """
        # Equity cashflows (negative for contributions, positive for distributions)
        equity_cf = -cf['equity_contribution'] + cf['equity_distribution']

        # Pre-tax IRR (before tax distributions)
        # For simplicity, using same equity_cf (in practice, would calculate without tax impact)
        pretax_irr = self.calculate_xirr(cf['date'], equity_cf)

        # Post-tax IRR (with tax impact)
        # This would normally adjust for tax distributions, but for now use same
        posttax_irr = pretax_irr

        return {
            'pre_tax_irr': pretax_irr,
            'post_tax_irr': posttax_irr
        }

    def calculate_project_irr(self, cf: pd.DataFrame) -> float:
        """
        Calculate project IRR (all-in, no financing).

        Args:
            cf: Cashflow DataFrame

        Returns:
            Project IRR
        """
        # Project cashflow = Revenue - OpEx - CapEx - Tax + Credits
        project_cf = (cf['total_revenue'] - cf['total_opex'] -
                     cf['total_capex'] - cf['total_tax'] +
                     cf['itc_credit'] + cf['upfront_grants'])

        project_irr = self.calculate_xirr(cf['date'], project_cf)

        return project_irr

    def calculate_npv(self, cf: pd.DataFrame) -> Dict[str, float]:
        """
        Calculate NPVs at various discount rates.

        Args:
            cf: Cashflow DataFrame

        Returns:
            Dict with equity_npv, project_npv
        """
        # Equity NPV
        equity_cf = -cf['equity_contribution'] + cf['equity_distribution']
        equity_npv = self.calculate_xnpv(cf['date'], equity_cf, self.discount_rate)

        # Project NPV
        project_cf = (cf['total_revenue'] - cf['total_opex'] -
                     cf['total_capex'] - cf['total_tax'] +
                     cf['itc_credit'] + cf['upfront_grants'])
        project_npv = self.calculate_xnpv(cf['date'], project_cf, self.discount_rate)

        return {
            'equity_npv': equity_npv,
            'project_npv': project_npv
        }

    def calculate_payback(self, cf: pd.DataFrame) -> float:
        """
        Calculate simple payback period (years).

        Args:
            cf: Cashflow DataFrame

        Returns:
            Payback period in years
        """
        equity_cf = -cf['equity_contribution'] + cf['equity_distribution']

        cumulative = 0.0
        for idx, cashflow in enumerate(equity_cf):
            cumulative += cashflow
            if cumulative > 0:
                # Payback achieved
                return (idx + 1) / 12.0  # Convert months to years

        return np.nan  # Never pays back

    def calculate_lcoe(self, cf: pd.DataFrame) -> Dict[str, float]:
        """
        Calculate Levelized Cost of Energy (nominal and real).

        Args:
            cf: Cashflow DataFrame

        Returns:
            Dict with nominal_lcoe and real_lcoe ($/MWh)
        """
        # LCOE = NPV of all costs / NPV of all energy

        # Costs = CapEx + OpEx + Debt Service - Credits/Grants
        costs = (cf['total_capex'] + cf['total_opex'] +
                cf['total_debt_service'] - cf['itc_credit'] -
                cf['upfront_grants'])

        # Energy
        energy_mwh = cf['ac_mwh']

        # NPV of costs and energy (nominal)
        npv_costs_nominal = self.calculate_xnpv(cf['date'], costs, self.discount_rate)
        npv_energy_nominal = self.calculate_xnpv(cf['date'], energy_mwh, self.discount_rate)

        nominal_lcoe = npv_costs_nominal / npv_energy_nominal if npv_energy_nominal > 0 else np.nan

        # Real LCOE (deflate using inflation)
        real_discount_rate = (1 + self.discount_rate) / (1 + self.inflation_rate) - 1
        npv_costs_real = self.calculate_xnpv(cf['date'], costs, real_discount_rate)
        npv_energy_real = self.calculate_xnpv(cf['date'], energy_mwh, real_discount_rate)

        real_lcoe = npv_costs_real / npv_energy_real if npv_energy_real > 0 else np.nan

        return {
            'nominal_lcoe': nominal_lcoe,
            'real_lcoe': real_lcoe
        }

    def calculate_dscr_metrics(self, cf: pd.DataFrame) -> Dict[str, float]:
        """
        Calculate DSCR metrics (min, avg, etc.).

        Args:
            cf: Cashflow DataFrame with dscr column

        Returns:
            Dict with min_dscr, avg_dscr, etc.
        """
        # Filter to operating periods with debt service
        operating = cf[cf['month_in_operation'] > 0]
        with_debt = operating[operating['total_debt_service'] > 0]

        if len(with_debt) == 0 or 'dscr' not in cf.columns:
            return {
                'min_dscr': np.nan,
                'avg_dscr': np.nan,
                'year1_dscr': np.nan
            }

        min_dscr = with_debt['dscr'].min()
        avg_dscr = with_debt['dscr'].mean()

        # Year 1 average DSCR
        year1 = with_debt[with_debt['year_in_operation'] == 1]
        year1_dscr = year1['dscr'].mean() if len(year1) > 0 else np.nan

        return {
            'min_dscr': min_dscr,
            'avg_dscr': avg_dscr,
            'year1_dscr': year1_dscr
        }

    def calculate_all_metrics(self, cf: pd.DataFrame) -> Dict[str, Any]:
        """
        Calculate all financial metrics.

        Args:
            cf: Complete cashflow DataFrame

        Returns:
            Dictionary of all metrics
        """
        metrics = {}

        # IRRs
        irr_results = self.calculate_equity_irr(cf)
        metrics.update(irr_results)

        metrics['project_irr'] = self.calculate_project_irr(cf)

        # NPVs
        npv_results = self.calculate_npv(cf)
        metrics.update(npv_results)

        # Payback
        metrics['payback_years'] = self.calculate_payback(cf)

        # LCOE
        lcoe_results = self.calculate_lcoe(cf)
        metrics.update(lcoe_results)

        # DSCR
        dscr_results = self.calculate_dscr_metrics(cf)
        metrics.update(dscr_results)

        # Total equity invested
        metrics['total_equity_invested'] = cf['equity_contribution'].sum()

        # Total CapEx
        metrics['total_capex'] = cf['total_capex'].sum()

        # Total debt
        metrics['total_debt'] = cf['debt_balance'].max()

        # Total energy
        metrics['lifetime_energy_mwh'] = cf['ac_mwh'].sum()

        # Total revenue
        metrics['lifetime_revenue'] = cf['total_revenue'].sum()

        return metrics

    def run_sensitivity(self, base_metrics: Dict[str, Any],
                       sensitivity_params: Dict[str, list]) -> pd.DataFrame:
        """
        Run sensitivity analysis on key parameters.

        Args:
            base_metrics: Base case metrics
            sensitivity_params: Dict of {param_name: [low_val, base_val, high_val]}

        Returns:
            DataFrame with sensitivity results
        """
        # This is a placeholder for sensitivity analysis
        # Would require re-running model with different inputs
        sensitivity_results = []

        for param, values in sensitivity_params.items():
            for value in values:
                # Would re-run model here
                # For now, placeholder
                sensitivity_results.append({
                    'parameter': param,
                    'value': value,
                    'equity_irr': base_metrics['pre_tax_irr'],  # Placeholder
                    'project_irr': base_metrics['project_irr'],  # Placeholder
                    'npv': base_metrics['equity_npv']  # Placeholder
                })

        return pd.DataFrame(sensitivity_results)
