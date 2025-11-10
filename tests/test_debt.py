"""Unit tests for debt module."""

import unittest
import sys
import os
import pandas as pd

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))

from model.debt import DebtModel


class TestDebtModel(unittest.TestCase):
    """Test debt financing calculations."""

    def setUp(self):
        """Set up test inputs."""
        self.inputs = {
            'project': {
                'construction_months': 12
            },
            'financing': {
                'use_debt': True,
                'sizing_method': 'target_dscr',
                'target_min_dscr': 1.30,
                'tenor_years': 15,
                'interest_rate_pct': 7.0,
                'amortization': 'sculpted',
                'upfront_fees_pct_of_debt': 2.0,
                'dsra_months': 6,
                'om_reserve_months': 6
            },
            'capex': {
                'idc_pct_of_debt_drawn': 100.0
            }
        }

    def test_debt_sizing_by_dscr(self):
        """Test debt sizing to target DSCR."""
        debt_model = DebtModel(self.inputs)

        # Create mock CFADS (monthly, operating period only)
        # $100k/month CFADS
        cfads_series = pd.Series([100000] * 180)  # 15 years

        total_capex = 10000000  # $10M

        debt_amount = debt_model.size_debt(cfads_series, total_capex)

        # Debt should be positive and less than total capex
        self.assertGreater(debt_amount, 0)
        self.assertLess(debt_amount, total_capex)

        # At 1.30 DSCR, monthly payment ~ 100k / 1.30 = $76,923
        # This should translate to debt of roughly $9-10M range
        self.assertGreater(debt_amount, 8000000)
        self.assertLess(debt_amount, 10000000)

    def test_debt_service_calculation(self):
        """Test debt service schedule."""
        debt_model = DebtModel(self.inputs)

        debt_principal = 5000000  # $5M
        total_months = 12 + (15 * 12)  # Construction + 15 year tenor

        debt_df = debt_model.calculate_debt_service(debt_principal, total_months)

        # Should have correct number of periods
        self.assertEqual(len(debt_df), total_months)

        # During construction, balance should equal principal
        construction = debt_df[debt_df['period'] < 12]
        self.assertTrue((construction['debt_balance'] == debt_principal).all())

        # After construction, balance should decline
        operating = debt_df[debt_df['period'] >= 12]
        balances = operating['debt_balance'].values

        # Balance should be monotonically decreasing
        for i in range(1, len(balances)):
            self.assertLessEqual(balances[i], balances[i-1])

        # Final balance should be zero or near zero
        final_balance = debt_df['debt_balance'].iloc[-1]
        self.assertLess(final_balance, 100)  # Allow small rounding

        # Total principal repaid should equal original debt
        total_principal_paid = debt_df['principal_payment'].sum()
        self.assertAlmostEqual(total_principal_paid, debt_principal, delta=1000)

    def test_dscr_calculation(self):
        """Test DSCR calculation."""
        debt_model = DebtModel(self.inputs)

        # Create mock cashflow
        cf = pd.DataFrame({
            'month_in_operation': range(0, 25),
            'total_debt_service': [0] * 1 + [10000] * 24,
            'cfads': [0] * 1 + [13000] * 24
        })

        cf = debt_model.calculate_dscr(cf)

        # Check DSCR values
        operating = cf[cf['month_in_operation'] > 0]

        # DSCR should be 13000 / 10000 = 1.3
        self.assertTrue(all(operating['dscr'] == 1.3))

    def test_no_debt_mode(self):
        """Test when debt is disabled."""
        self.inputs['financing']['use_debt'] = False
        debt_model = DebtModel(self.inputs)

        cfads_series = pd.Series([100000] * 180)
        debt_amount = debt_model.size_debt(cfads_series, 10000000)

        self.assertEqual(debt_amount, 0.0)


if __name__ == '__main__':
    unittest.main()
