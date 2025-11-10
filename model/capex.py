"""
Capital expenditure calculation module.
Handles CapEx breakdown, developer fee, IDC, and timing.
"""

import pandas as pd
import numpy as np
from typing import Dict, Any, Tuple


class CapExModel:
    """Calculates capital expenditures and developer fee."""

    def __init__(self, inputs: Dict[str, Any]):
        self.inputs = inputs
        self.capex = inputs['capex']
        self.developer = inputs['developer']
        self.project = inputs['project']
        self.financing = inputs['financing']

        self.construction_months = self.project['construction_months']

    def calculate_total_epc_cost(self) -> Tuple[float, Dict[str, float]]:
        """
        Calculate total EPC cost before contingency and fees.

        Returns:
            (total_epc, breakdown_dict)
        """
        breakdown = {
            'modules': self.capex.get('modules_usd', 0.0),
            'inverters': self.capex.get('inverters_usd', 0.0),
            'racking': self.capex.get('racking_usd', 0.0),
            'bos': self.capex.get('bos_usd', 0.0),
            'civil': self.capex.get('civil_usd', 0.0),
            'interconnection': self.capex.get('interconnection_usd', 0.0),
            'owner_costs': self.capex.get('owner_costs_usd', 0.0),
            'development_soft_costs': self.capex.get('development_soft_costs_usd', 0.0),
            'epc_indirects': self.capex.get('epc_indirects_usd', 0.0),
        }

        subtotal = sum(breakdown.values())

        # Add contingency
        contingency_pct = self.capex.get('contingency_pct', 5.0) / 100.0
        contingency = subtotal * contingency_pct
        breakdown['contingency'] = contingency

        total_epc = subtotal + contingency

        return total_epc, breakdown

    def calculate_developer_fee(self, epc_cost: float) -> float:
        """Calculate developer fee."""
        mode = self.developer['developer_fee_mode']

        if mode == 'percent_of_epc':
            fee_pct = self.developer.get('developer_fee_pct', 0.0) / 100.0
            return epc_cost * fee_pct
        elif mode == 'fixed':
            return self.developer.get('developer_fee_fixed_usd', 0.0)
        else:
            return 0.0

    def calculate_total_capex(self, include_idc: bool = True) -> Tuple[float, Dict[str, float]]:
        """
        Calculate total project CapEx including developer fee and IDC.

        Args:
            include_idc: Whether to include Interest During Construction

        Returns:
            (total_capex, full_breakdown)
        """
        epc_cost, breakdown = self.calculate_total_epc_cost()

        # Developer fee
        dev_fee = self.calculate_developer_fee(epc_cost)
        breakdown['developer_fee'] = dev_fee

        # Decommissioning reserve
        decom = self.capex.get('decommissioning_reserve_usd', 0.0)
        breakdown['decommissioning_reserve'] = decom

        # Subtotal before IDC
        subtotal = epc_cost + dev_fee + decom

        # IDC (if debt is used and we want to include it)
        if include_idc and self.financing.get('use_debt', False):
            # IDC is calculated in debt module, but we can estimate here
            # For now, placeholder (will be calculated properly in debt module)
            breakdown['idc'] = 0.0  # Placeholder
        else:
            breakdown['idc'] = 0.0

        total = subtotal + breakdown['idc']
        breakdown['total'] = total

        return total, breakdown

    def create_capex_schedule(self, total_months: int) -> pd.DataFrame:
        """
        Create monthly CapEx schedule.

        Args:
            total_months: Total periods in model

        Returns:
            DataFrame with period and capex spend by category
        """
        _, breakdown = self.calculate_total_capex(include_idc=False)

        # Create monthly schedule
        periods = list(range(total_months))

        # Allocate CapEx during construction period
        # Simple allocation: spread evenly over construction months
        construction_capex_categories = [
            'modules', 'inverters', 'racking', 'bos', 'civil',
            'interconnection', 'owner_costs', 'development_soft_costs',
            'epc_indirects', 'contingency'
        ]

        monthly_data = {
            'period': periods
        }

        # Initialize all categories to 0
        for cat in construction_capex_categories + ['developer_fee', 'decommissioning_reserve']:
            monthly_data[cat] = [0.0] * total_months

        # Allocate construction CapEx evenly during construction
        for cat in construction_capex_categories:
            monthly_amount = breakdown[cat] / self.construction_months if self.construction_months > 0 else 0
            for i in range(self.construction_months):
                monthly_data[cat][i] = monthly_amount

        # Developer fee timing
        dev_fee_timing = self.developer.get('developer_fee_timing', 'COD')
        dev_fee_total = breakdown['developer_fee']

        if dev_fee_timing == 'NTP':
            # Pay at period 0
            monthly_data['developer_fee'][0] = dev_fee_total
        elif dev_fee_timing == 'COD':
            # Pay at end of construction
            if self.construction_months > 0:
                monthly_data['developer_fee'][self.construction_months - 1] = dev_fee_total
        elif dev_fee_timing == 'over_time':
            # Spread over construction
            monthly_dev_fee = dev_fee_total / self.construction_months if self.construction_months > 0 else 0
            for i in range(self.construction_months):
                monthly_data['developer_fee'][i] = monthly_dev_fee

        # Decommissioning reserve at COD
        if self.construction_months > 0:
            monthly_data['decommissioning_reserve'][self.construction_months - 1] = breakdown['decommissioning_reserve']

        df = pd.DataFrame(monthly_data)

        # Calculate total monthly capex
        df['total_capex'] = df[construction_capex_categories + ['developer_fee', 'decommissioning_reserve']].sum(axis=1)

        return df

    def get_depreciable_basis(self) -> float:
        """
        Calculate depreciable basis (before ITC reduction).

        Excludes land and other non-depreciable items.
        """
        total_capex, breakdown = self.calculate_total_capex(include_idc=False)

        # For solar, most items are depreciable
        # Exclude: land (if purchased), and possibly some owner costs
        depreciable = total_capex

        # Subtract decommissioning reserve if it's not depreciable
        depreciable -= breakdown.get('decommissioning_reserve', 0.0)

        return depreciable
