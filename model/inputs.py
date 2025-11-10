"""
Input validation and loading module.
Handles JSON schema validation, defaults, and audit trail.
"""

import json
from datetime import datetime
from typing import Dict, Any, List, Tuple
from copy import deepcopy


class InputValidator:
    """Validates and processes input JSON with defaults and audit tracking."""

    def __init__(self):
        self.defaults_used = []
        self.validation_errors = []
        self.warnings = []

    def load_and_validate(self, json_path: str) -> Dict[str, Any]:
        """Load JSON and validate with defaults."""
        with open(json_path, 'r') as f:
            data = json.load(f)

        validated = self._apply_defaults(data)
        self._validate_inputs(validated)

        if self.validation_errors:
            raise ValueError(f"Input validation failed: {self.validation_errors}")

        return validated

    def _apply_defaults(self, data: Dict[str, Any]) -> Dict[str, Any]:
        """Apply defaults for missing values."""
        result = deepcopy(data)

        # Project defaults
        if 'project' not in result:
            result['project'] = {}
        self._set_default(result['project'], 'mode', 'community_solar', 'project.mode')
        self._set_default(result['project'], 'model_years', 35, 'project.model_years')
        self._set_default(result['project'], 'construction_months', 12, 'project.construction_months')
        self._set_default(result['project'], 'discount_rate_nominal_pct', 8.0, 'project.discount_rate_nominal_pct')
        self._set_default(result['project'], 'inflation_pct', 2.5, 'project.inflation_pct')

        # Sizing defaults
        if 'sizing' not in result:
            result['sizing'] = {}
        self._set_default(result['sizing'], 'dc_ac_ratio', 1.2, 'sizing.dc_ac_ratio')
        self._set_default(result['sizing'], 'performance_ratio_pct', 86.0, 'sizing.performance_ratio_pct')
        self._set_default(result['sizing'], 'degradation_pct_per_year', 0.5, 'sizing.degradation_pct_per_year')
        self._set_default(result['sizing'], 'availability_pct', 98.0, 'sizing.availability_pct')
        self._set_default(result['sizing'], 'curtailment_pct', 0.0, 'sizing.curtailment_pct')
        self._set_default(result['sizing'], 'use_8760', False, 'sizing.use_8760')
        self._set_default(result['sizing'], '8760_csv_path', None, 'sizing.8760_csv_path')

        # Revenue common defaults
        if 'revenue_common' not in result:
            result['revenue_common'] = {}
        self._set_default(result['revenue_common'], 'rec_price_usd_per_mwh', 0.0, 'revenue_common.rec_price_usd_per_mwh')
        self._set_default(result['revenue_common'], 'rec_escalator_pct', 0.0, 'revenue_common.rec_escalator_pct')
        self._set_default(result['revenue_common'], 'capacity_revenue_usd_per_kw_year', 0.0, 'revenue_common.capacity_revenue_usd_per_kw_year')
        self._set_default(result['revenue_common'], 'capacity_term_years', 0, 'revenue_common.capacity_term_years')

        # Community solar defaults
        if 'community_solar' not in result:
            result['community_solar'] = {}
        self._set_default(result['community_solar'], 'subscriber_discount_pct', 10.0, 'community_solar.subscriber_discount_pct')
        self._set_default(result['community_solar'], 'anchor_share_pct', 30.0, 'community_solar.anchor_share_pct')
        self._set_default(result['community_solar'], 'lmi_share_pct', 20.0, 'community_solar.lmi_share_pct')
        self._set_default(result['community_solar'], 'mgmt_fee_usd_per_acct_month', 2.5, 'community_solar.mgmt_fee_usd_per_acct_month')
        self._set_default(result['community_solar'], 'annual_churn_pct', 6.0, 'community_solar.annual_churn_pct')
        self._set_default(result['community_solar'], 'bad_debt_pct_of_billings', 1.5, 'community_solar.bad_debt_pct_of_billings')
        self._set_default(result['community_solar'], 'ramp_to_full_subscribed_months', 12, 'community_solar.ramp_to_full_subscribed_months')

        # PPA defaults
        if 'ppa' not in result:
            result['ppa'] = {}
        self._set_default(result['ppa'], 'settlement', 'pay_as_produced', 'ppa.settlement')

        # Tax credits defaults
        if 'tax_credits' not in result:
            result['tax_credits'] = {}
        self._set_default(result['tax_credits'], 'mode', 'ITC', 'tax_credits.mode')
        self._set_default(result['tax_credits'], 'itc_pct', 30.0, 'tax_credits.itc_pct')
        self._set_default(result['tax_credits'], 'adders_pct', 0.0, 'tax_credits.adders_pct')
        self._set_default(result['tax_credits'], 'basis_reduction_pct', 50.0, 'tax_credits.basis_reduction_pct')
        self._set_default(result['tax_credits'], 'bonus_depreciation_pct', 0.0, 'tax_credits.bonus_depreciation_pct')
        self._set_default(result['tax_credits'], 'elective_pay', False, 'tax_credits.elective_pay')
        self._set_default(result['tax_credits'], 'reduce_itc_basis_for_grants', False, 'tax_credits.reduce_itc_basis_for_grants')

        # Taxes defaults
        if 'taxes' not in result:
            result['taxes'] = {}
        self._set_default(result['taxes'], 'federal_tax_rate_pct', 21.0, 'taxes.federal_tax_rate_pct')
        self._set_default(result['taxes'], 'state_tax_rate_pct', 0.0, 'taxes.state_tax_rate_pct')
        self._set_default(result['taxes'], 'depr_schedule', 'MACRS_5yr', 'taxes.depr_schedule')
        self._set_default(result['taxes'], 'state_depr_schedule', None, 'taxes.state_depr_schedule')
        self._set_default(result['taxes'], 'nol_carryforward_years', 20, 'taxes.nol_carryforward_years')

        # CapEx defaults
        if 'capex' not in result:
            result['capex'] = {}
        self._set_default(result['capex'], 'contingency_pct', 5.0, 'capex.contingency_pct')
        self._set_default(result['capex'], 'idc_pct_of_debt_drawn', 100.0, 'capex.idc_pct_of_debt_drawn')

        # Developer defaults
        if 'developer' not in result:
            result['developer'] = {}
        self._set_default(result['developer'], 'developer_fee_mode', 'percent_of_epc', 'developer.developer_fee_mode')
        self._set_default(result['developer'], 'developer_fee_pct', 0.0, 'developer.developer_fee_pct')
        self._set_default(result['developer'], 'developer_fee_fixed_usd', 0.0, 'developer.developer_fee_fixed_usd')
        self._set_default(result['developer'], 'developer_fee_timing', 'COD', 'developer.developer_fee_timing')

        # Land defaults
        if 'land' not in result:
            result['land'] = {}
        self._set_default(result['land'], 'mode', 'lease', 'land.mode')

        # Property tax defaults
        if 'property_tax_pilot' not in result:
            result['property_tax_pilot'] = {}
        self._set_default(result['property_tax_pilot'], 'pilot_enabled', False, 'property_tax_pilot.pilot_enabled')

        # OpEx defaults
        if 'opex' not in result:
            result['opex'] = {}

        # Financing defaults
        if 'financing' not in result:
            result['financing'] = {}
        self._set_default(result['financing'], 'use_debt', False, 'financing.use_debt')
        self._set_default(result['financing'], 'sizing_method', 'target_dscr', 'financing.sizing_method')
        self._set_default(result['financing'], 'amortization', 'sculpted', 'financing.amortization')
        self._set_default(result['financing'], 'upfront_fees_pct_of_debt', 2.0, 'financing.upfront_fees_pct_of_debt')
        self._set_default(result['financing'], 'dsra_months', 6, 'financing.dsra_months')
        self._set_default(result['financing'], 'om_reserve_months', 6, 'financing.om_reserve_months')

        # Exit defaults
        if 'exit' not in result:
            result['exit'] = {}
        self._set_default(result['exit'], 'terminal_value_method', 'exit_multiple_on_yearN_cashflow', 'exit.terminal_value_method')
        self._set_default(result['exit'], 'exit_multiple', 8.0, 'exit.exit_multiple')

        # Program defaults
        if 'program' not in result:
            result['program'] = {}
        self._set_default(result['program'], 'active', 'none', 'program.active')

        return result

    def _set_default(self, section: Dict, key: str, default: Any, path: str):
        """Set a default value and track it."""
        if key not in section or section[key] is None:
            section[key] = default
            self.defaults_used.append(f"{path} = {default}")

    def _validate_inputs(self, data: Dict[str, Any]):
        """Validate input constraints."""
        # Mode validation
        mode = data['project']['mode']
        if mode not in ['community_solar', 'ppa']:
            self.validation_errors.append(f"Invalid mode: {mode}")

        # Sizing validation
        if data['sizing'].get('dc_kw', 0) <= 0:
            self.validation_errors.append("dc_kw must be > 0")
        if data['sizing'].get('ac_kw', 0) <= 0:
            self.validation_errors.append("ac_kw must be > 0")

        # Program mutual exclusivity
        program_active = data['program']['active']
        if program_active != 'none':
            enabled_count = 0
            for prog in ['ny_vder', 'nj_csep', 'il_abp']:
                if prog in data['program'] and data['program'][prog].get('enabled', False):
                    enabled_count += 1

            if enabled_count > 1:
                self.validation_errors.append("Only one state program can be enabled at a time")

        # Tax credit mode validation
        tax_mode = data['tax_credits']['mode']
        if tax_mode not in ['ITC', 'PTC', 'None']:
            self.validation_errors.append(f"Invalid tax_credits.mode: {tax_mode}")

        # Financing validation
        if data['financing']['use_debt']:
            if 'interest_rate_pct' not in data['financing']:
                self.validation_errors.append("interest_rate_pct required when use_debt=True")
            if 'tenor_years' not in data['financing']:
                self.validation_errors.append("tenor_years required when use_debt=True")


def load_inputs(json_path: str) -> Tuple[Dict[str, Any], List[str], List[str]]:
    """
    Load and validate inputs from JSON file.

    Returns:
        (validated_data, defaults_used, warnings)
    """
    validator = InputValidator()
    data = validator.load_and_validate(json_path)
    return data, validator.defaults_used, validator.warnings
