"""Unit tests for energy module."""

import unittest
import sys
import os
from datetime import datetime

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))

from model.energy import EnergyModel


class TestEnergyModel(unittest.TestCase):
    """Test energy production calculations."""

    def setUp(self):
        """Set up test inputs."""
        self.inputs = {
            'project': {
                'cod_date': '2027-06-01',
                'construction_months': 12,
                'model_years': 25
            },
            'sizing': {
                'dc_kw': 10000,
                'ac_kw': 8000,
                'dc_ac_ratio': 1.25,
                'capacity_factor_pct': 20.0,
                'performance_ratio_pct': 85.0,
                'degradation_pct_per_year': 0.5,
                'availability_pct': 98.0,
                'curtailment_pct': 2.0,
                'use_8760': False,
                '8760_csv_path': None
            }
        }

    def test_synthetic_energy_calculation(self):
        """Test synthetic energy generation."""
        energy_model = EnergyModel(self.inputs)
        df = energy_model.calculate_monthly_energy()

        # Should have construction + operating months
        expected_months = 12 + (25 * 12)
        self.assertEqual(len(df), expected_months)

        # Construction months should have zero production
        construction = df[df['month_in_operation'] == 0]
        self.assertEqual(construction['ac_kwh'].sum(), 0)

        # Operating months should have production
        operating = df[df['month_in_operation'] > 0]
        self.assertGreater(operating['ac_kwh'].sum(), 0)

        # Check degradation: year 2 should be less than year 1
        year1 = df[df['year_in_operation'] == 1]['ac_kwh'].sum()
        year2 = df[df['year_in_operation'] == 2]['ac_kwh'].sum()

        # Year 2 should be ~0.5% less than year 1
        expected_ratio = (100 - 0.5) / 100
        self.assertAlmostEqual(year2 / year1, expected_ratio, places=2)

    def test_energy_with_degradation(self):
        """Test that degradation is properly applied."""
        energy_model = EnergyModel(self.inputs)
        df = energy_model.calculate_monthly_energy()

        # Get year 1 and year 25 production
        year1 = df[df['year_in_operation'] == 1]['ac_kwh'].sum()
        year25 = df[df['year_in_operation'] == 25]['ac_kwh'].sum()

        # After 24 years of 0.5%/year degradation
        # Degradation factor = 1 - (0.005 * 24) = 0.88
        expected_ratio = 1 - (0.005 * 24)

        actual_ratio = year25 / year1

        self.assertAlmostEqual(actual_ratio, expected_ratio, places=2)

    def test_dc_ac_relationship(self):
        """Test DC to AC relationship."""
        energy_model = EnergyModel(self.inputs)
        df = energy_model.calculate_monthly_energy()

        # DC should be higher than AC by DC/AC ratio
        operating = df[df['month_in_operation'] > 0]

        # Get first operating month
        first_op = operating.iloc[0]

        dc_ac_ratio = first_op['dc_kwh'] / first_op['ac_kwh']

        self.assertAlmostEqual(dc_ac_ratio, 1.25, places=1)


if __name__ == '__main__':
    unittest.main()
