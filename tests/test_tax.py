"""Unit tests for tax module."""

import unittest
import sys
import os

# Add parent directory to path
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))

from model.tax import TaxModel


class TestTaxModel(unittest.TestCase):
    """Test tax calculations."""

    def setUp(self):
        """Set up test inputs."""
        self.inputs = {
            'project': {'construction_months': 12, 'model_years': 25},
            'tax_credits': {
                'mode': 'ITC',
                'itc_pct': 30.0,
                'adders_pct': 10.0,
                'basis_reduction_pct': 50.0,
                'bonus_depreciation_pct': 0.0,
                'reduce_itc_basis_for_grants': False
            },
            'taxes': {
                'federal_tax_rate_pct': 21.0,
                'state_tax_rate_pct': 6.5,
                'depr_schedule': 'MACRS_5yr',
                'nol_carryforward_years': 20
            }
        }

    def test_itc_calculation_basic(self):
        """Test basic ITC calculation."""
        tax_model = TaxModel(self.inputs)

        depreciable_basis = 10000000  # $10M
        itc_amount, reduced_basis = tax_model.calculate_itc(depreciable_basis, 0.0)

        # ITC = 10M * (30% + 10%) = $4M
        expected_itc = 10000000 * 0.40
        self.assertAlmostEqual(itc_amount, expected_itc, places=2)

        # Reduced basis = 10M - (4M * 50%) = $8M
        expected_reduced_basis = 10000000 - (expected_itc * 0.50)
        self.assertAlmostEqual(reduced_basis, expected_reduced_basis, places=2)

    def test_itc_with_grant_basis_reduction(self):
        """Test ITC with grant basis reduction."""
        self.inputs['tax_credits']['reduce_itc_basis_for_grants'] = True
        tax_model = TaxModel(self.inputs)

        depreciable_basis = 10000000
        upfront_grants = 500000

        itc_amount, reduced_basis = tax_model.calculate_itc(depreciable_basis, upfront_grants)

        # ITC basis = 10M - 500K = 9.5M
        # ITC = 9.5M * 40% = $3.8M
        expected_itc = (10000000 - 500000) * 0.40
        self.assertAlmostEqual(itc_amount, expected_itc, places=2)

    def test_depreciation_macrs_5yr(self):
        """Test MACRS 5-year depreciation schedule."""
        tax_model = TaxModel(self.inputs)

        depreciable_basis = 10000000
        total_months = 12 + (25 * 12)  # Construction + operating

        depr_df = tax_model.calculate_depreciation(depreciable_basis, total_months)

        # Check annual depreciation sums to basis
        # Group by year and sum
        depr_df['year'] = depr_df['period'] // 12

        annual_depr = depr_df[depr_df['period'] >= 12].groupby('year')['depreciation_federal'].sum()

        # Year 1 should be 20% of basis
        year_1_depr = annual_depr.iloc[0]
        expected_year_1 = depreciable_basis * 0.20
        self.assertAlmostEqual(year_1_depr, expected_year_1, delta=1000)

        # Total depreciation should equal basis
        total_depr = depr_df['depreciation_federal'].sum()
        self.assertAlmostEqual(total_depr, depreciable_basis, delta=1000)

    def test_no_itc_mode(self):
        """Test when ITC mode is off."""
        self.inputs['tax_credits']['mode'] = 'None'
        tax_model = TaxModel(self.inputs)

        depreciable_basis = 10000000
        itc_amount, reduced_basis = tax_model.calculate_itc(depreciable_basis, 0.0)

        self.assertEqual(itc_amount, 0.0)
        self.assertEqual(reduced_basis, depreciable_basis)


if __name__ == '__main__':
    unittest.main()
