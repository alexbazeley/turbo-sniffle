"""Unit tests for metrics module."""

import unittest
import sys
import os
import pandas as pd
import numpy as np
from datetime import datetime, timedelta

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))

from model.metrics import MetricsCalculator


class TestMetricsCalculator(unittest.TestCase):
    """Test financial metrics calculations."""

    def setUp(self):
        """Set up test inputs."""
        self.inputs = {
            'project': {
                'discount_rate_nominal_pct': 8.0,
                'inflation_pct': 2.5
            }
        }
        self.calc = MetricsCalculator(self.inputs)

    def test_xirr_simple_case(self):
        """Test XIRR calculation with simple cashflows."""
        # Simple case: invest $1000, get back $1200 after 1 year
        dates = pd.Series([
            datetime(2025, 1, 1),
            datetime(2026, 1, 1)
        ])
        cashflows = pd.Series([-1000, 1200])

        irr = self.calc.calculate_xirr(dates, cashflows)

        # IRR should be 20%
        self.assertAlmostEqual(irr, 0.20, places=2)

    def test_xnpv_simple_case(self):
        """Test XNPV calculation."""
        dates = pd.Series([
            datetime(2025, 1, 1),
            datetime(2026, 1, 1)
        ])
        cashflows = pd.Series([-1000, 1100])

        npv = self.calc.calculate_xnpv(dates, cashflows, 0.10)

        # NPV at 10% discount = -1000 + 1100/1.1 = $0
        expected_npv = -1000 + 1100 / 1.10
        self.assertAlmostEqual(npv, expected_npv, places=2)

    def test_lcoe_calculation(self):
        """Test LCOE calculation."""
        # Create simple cashflow
        dates = [datetime(2025, 1, 1) + timedelta(days=30*i) for i in range(24)]

        cf = pd.DataFrame({
            'date': dates,
            'total_capex': [10000] + [0] * 23,
            'total_opex': [0] + [100] * 23,
            'total_debt_service': [0] * 24,
            'itc_credit': [0] + [2000] + [0] * 22,
            'upfront_grants': [0] * 24,
            'ac_mwh': [0] + [10] * 23
        })

        lcoe = self.calc.calculate_lcoe(cf)

        # LCOE should be positive
        self.assertGreater(lcoe['nominal_lcoe'], 0)
        self.assertGreater(lcoe['real_lcoe'], 0)

        # Real LCOE should be less than nominal (inflation adjustment)
        self.assertLess(lcoe['real_lcoe'], lcoe['nominal_lcoe'])

    def test_payback_calculation(self):
        """Test payback period calculation."""
        dates = [datetime(2025, 1, 1) + timedelta(days=30*i) for i in range(60)]

        # Invest $10k, get back $500/month for 2 years
        equity_contrib = [10000] + [0] * 59
        equity_distrib = [0] * 12 + [500] * 48  # Start distributions after year 1

        cf = pd.DataFrame({
            'date': dates,
            'equity_contribution': equity_contrib,
            'equity_distribution': equity_distrib
        })

        payback = self.calc.calculate_payback(cf)

        # Payback should occur around month 32 (12 + 10000/500)
        # = 12 + 20 = 32 months = 2.67 years
        self.assertGreater(payback, 2.5)
        self.assertLess(payback, 3.0)

    def test_dscr_metrics(self):
        """Test DSCR metrics calculation."""
        cf = pd.DataFrame({
            'month_in_operation': range(0, 25),
            'total_debt_service': [0] * 1 + [1000] * 24,
            'cfads': [0] * 1 + [1400] * 24,
            'dscr': [None] + [1.4] * 24
        })

        metrics = self.calc.calculate_dscr_metrics(cf)

        self.assertAlmostEqual(metrics['min_dscr'], 1.4, places=2)
        self.assertAlmostEqual(metrics['avg_dscr'], 1.4, places=2)


if __name__ == '__main__':
    unittest.main()
