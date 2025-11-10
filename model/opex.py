"""
Operating expenditure calculation module.
Handles OpEx, land costs (lease/purchase), property tax/PILOT.
"""

import pandas as pd
import numpy as np
from typing import Dict, Any


class OpExModel:
    """Calculates operating expenses, land costs, and property taxes."""

    def __init__(self, inputs: Dict[str, Any], sizing_ac_kw: float):
        self.inputs = inputs
        self.opex = inputs['opex']
        self.land = inputs['land']
        self.property_tax = inputs['property_tax_pilot']
        self.project = inputs['project']

        self.ac_kw = sizing_ac_kw
        self.construction_months = self.project['construction_months']

    def calculate_monthly_opex(self, energy_df: pd.DataFrame) -> pd.DataFrame:
        """
        Calculate monthly operating expenses.

        Args:
            energy_df: DataFrame with energy production (ac_mwh column)

        Returns:
            Updated DataFrame with OpEx columns
        """
        # Fixed O&M ($/kW-year)
        fixed_om_annual = self.opex.get('fixed_om_usd_per_kw_year', 0.0) * self.ac_kw
        fixed_om_monthly = fixed_om_annual / 12.0

        # Variable O&M ($/MWh)
        variable_om_rate = self.opex.get('variable_om_usd_per_mwh', 0.0)

        # Insurance ($/kW-year)
        insurance_annual = self.opex.get('insurance_usd_per_kw_year', 0.0) * self.ac_kw
        insurance_monthly = insurance_annual / 12.0

        # Asset management ($/kW-year)
        asset_mgmt_annual = self.opex.get('asset_mgmt_usd_per_kw_year', 0.0) * self.ac_kw
        asset_mgmt_monthly = asset_mgmt_annual / 12.0

        # Other OpEx (annual)
        other_opex_annual = self.opex.get('other_opex_usd_year', 0.0)
        other_opex_monthly = other_opex_annual / 12.0

        # Calculate per period
        fixed_om = []
        variable_om = []
        insurance = []
        asset_mgmt = []
        other_opex = []

        for idx, row in energy_df.iterrows():
            month_in_op = row['month_in_operation']

            if month_in_op > 0:
                # Operating period
                fixed_om.append(fixed_om_monthly)
                variable_om.append(row['ac_mwh'] * variable_om_rate)
                insurance.append(insurance_monthly)
                asset_mgmt.append(asset_mgmt_monthly)
                other_opex.append(other_opex_monthly)
            else:
                # Construction - no OpEx
                fixed_om.append(0.0)
                variable_om.append(0.0)
                insurance.append(0.0)
                asset_mgmt.append(0.0)
                other_opex.append(0.0)

        energy_df['fixed_om'] = fixed_om
        energy_df['variable_om'] = variable_om
        energy_df['insurance'] = insurance
        energy_df['asset_management'] = asset_mgmt
        energy_df['other_opex'] = other_opex

        return energy_df

    def calculate_land_costs(self, energy_df: pd.DataFrame) -> pd.DataFrame:
        """Calculate land lease or purchase costs."""
        land_mode = self.land.get('mode', 'lease')

        if land_mode == 'lease':
            # Annual lease payment
            base_rent_per_acre = self.land.get('lease_base_usd_per_acre_year', 0.0)
            acres = self.land.get('acres', 0.0)
            escalator = self.land.get('lease_escalator_pct', 0.0) / 100.0

            base_annual_rent = base_rent_per_acre * acres

            land_cost = []

            for idx, row in energy_df.iterrows():
                month_in_op = row['month_in_operation']
                year_in_op = row['year_in_operation']

                if month_in_op > 0:
                    # Apply escalator
                    escalation_factor = (1 + escalator) ** (year_in_op - 1)
                    annual_rent = base_annual_rent * escalation_factor
                    monthly_rent = annual_rent / 12.0
                    land_cost.append(monthly_rent)
                else:
                    land_cost.append(0.0)

            energy_df['land_lease'] = land_cost

        elif land_mode == 'purchase':
            # Land purchased upfront (shown in CapEx)
            # No recurring cost
            energy_df['land_lease'] = 0.0

        return energy_df

    def calculate_property_tax(self, energy_df: pd.DataFrame) -> pd.DataFrame:
        """Calculate property tax or PILOT payments."""
        pilot_enabled = self.property_tax.get('pilot_enabled', False)

        if pilot_enabled:
            # PILOT schedule
            pilot_schedule = self.property_tax.get('pilot_schedule_usd_per_year', [])

            property_tax = []

            for idx, row in energy_df.iterrows():
                month_in_op = row['month_in_operation']
                year_in_op = row['year_in_operation']

                if month_in_op > 0:
                    year_idx = year_in_op - 1

                    if year_idx < len(pilot_schedule):
                        annual_pilot = pilot_schedule[year_idx]
                    else:
                        # Use last value if schedule runs out
                        annual_pilot = pilot_schedule[-1] if pilot_schedule else 0.0

                    monthly_pilot = annual_pilot / 12.0
                    property_tax.append(monthly_pilot)
                else:
                    property_tax.append(0.0)

            energy_df['property_tax'] = property_tax

        else:
            # Standard property tax
            tax_info = self.property_tax.get('property_tax_if_no_pilot', {})
            assessed_value = tax_info.get('assessed_value_usd', 0.0)
            mill_rate = tax_info.get('mill_rate_pct', 0.0) / 100.0

            annual_tax = assessed_value * mill_rate
            monthly_tax = annual_tax / 12.0

            property_tax = []

            for idx, row in energy_df.iterrows():
                month_in_op = row['month_in_operation']

                if month_in_op > 0:
                    property_tax.append(monthly_tax)
                else:
                    property_tax.append(0.0)

            energy_df['property_tax'] = property_tax

        return energy_df

    def calculate_inverter_replacements(self, energy_df: pd.DataFrame) -> pd.DataFrame:
        """Calculate scheduled inverter replacements."""
        # Major repairs/replacements (specified by year)
        inverter_repair_usd = self.opex.get('inverter_repairs_major_usd_year', 0.0)

        # Typically at year 10, 20, etc. (user can customize)
        # For now, simple: no scheduled replacements (can be enhanced)

        energy_df['inverter_replacement'] = 0.0

        return energy_df

    def calculate_total_opex(self, energy_df: pd.DataFrame) -> pd.DataFrame:
        """
        Calculate total operating expenses.

        Args:
            energy_df: DataFrame with energy data

        Returns:
            Updated DataFrame with total_opex column
        """
        opex_columns = [
            'fixed_om', 'variable_om', 'insurance', 'asset_management',
            'other_opex', 'land_lease', 'property_tax', 'inverter_replacement'
        ]

        # Ensure all columns exist
        for col in opex_columns:
            if col not in energy_df.columns:
                energy_df[col] = 0.0

        energy_df['total_opex'] = energy_df[opex_columns].sum(axis=1)

        return energy_df
