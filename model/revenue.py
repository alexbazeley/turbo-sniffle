"""
Revenue calculation module.
Handles Community Solar, PPA, and state program revenue (NY-Sun/VDER, NJ CSEP, IL ABP).
"""

import pandas as pd
import numpy as np
from typing import Dict, Any, Tuple


class RevenueModel:
    """Calculates revenue streams for Community Solar and PPA modes."""

    def __init__(self, inputs: Dict[str, Any], sizing_ac_kw: float):
        self.inputs = inputs
        self.project = inputs['project']
        self.mode = self.project['mode']
        self.revenue_common = inputs.get('revenue_common', {})
        self.community_solar = inputs.get('community_solar', {})
        self.ppa = inputs.get('ppa', {})
        self.program = inputs.get('program', {})

        self.ac_kw = sizing_ac_kw
        self.construction_months = self.project['construction_months']

    def calculate_revenue(self, energy_df: pd.DataFrame) -> pd.DataFrame:
        """
        Calculate total revenue by mode.

        Args:
            energy_df: DataFrame with energy production

        Returns:
            Updated DataFrame with revenue columns
        """
        # Common revenue streams
        energy_df = self._calculate_rec_revenue(energy_df)
        energy_df = self._calculate_capacity_revenue(energy_df)

        # Mode-specific revenue
        if self.mode == 'community_solar':
            energy_df = self._calculate_community_solar_revenue(energy_df)
        elif self.mode == 'ppa':
            energy_df = self._calculate_ppa_revenue(energy_df)

        # State program revenue (overrides/supplements based on program)
        program_active = self.program.get('active', 'none')
        if program_active == 'ny_vder':
            energy_df = self._calculate_ny_vder_revenue(energy_df)
        elif program_active == 'nj_csep':
            energy_df = self._calculate_nj_csep_revenue(energy_df)
        elif program_active == 'il_abp':
            energy_df = self._calculate_il_abp_revenue(energy_df)

        # Calculate total revenue
        self._calculate_total_revenue(energy_df)

        return energy_df

    def _calculate_rec_revenue(self, df: pd.DataFrame) -> pd.DataFrame:
        """Calculate REC/SREC revenue."""
        rec_price = self.revenue_common.get('rec_price_usd_per_mwh', 0.0)
        rec_escalator = self.revenue_common.get('rec_escalator_pct', 0.0) / 100.0

        rec_revenue = []

        for idx, row in df.iterrows():
            month_in_op = row['month_in_operation']
            year_in_op = row['year_in_operation']

            if month_in_op > 0:
                escalation_factor = (1 + rec_escalator) ** (year_in_op - 1)
                current_rec_price = rec_price * escalation_factor
                revenue = row['ac_mwh'] * current_rec_price
                rec_revenue.append(revenue)
            else:
                rec_revenue.append(0.0)

        df['rec_revenue'] = rec_revenue
        return df

    def _calculate_capacity_revenue(self, df: pd.DataFrame) -> pd.DataFrame:
        """Calculate capacity/ICAP revenue."""
        capacity_revenue_per_kw_year = self.revenue_common.get('capacity_revenue_usd_per_kw_year', 0.0)
        capacity_term_years = self.revenue_common.get('capacity_term_years', 0)

        capacity_revenue_monthly = (capacity_revenue_per_kw_year * self.ac_kw) / 12.0

        capacity_revenue = []

        for idx, row in df.iterrows():
            month_in_op = row['month_in_operation']
            year_in_op = row['year_in_operation']

            if month_in_op > 0 and (capacity_term_years == 0 or year_in_op <= capacity_term_years):
                capacity_revenue.append(capacity_revenue_monthly)
            else:
                capacity_revenue.append(0.0)

        df['capacity_revenue'] = capacity_revenue
        return df

    def _calculate_community_solar_revenue(self, df: pd.DataFrame) -> pd.DataFrame:
        """Calculate Community Solar subscriber revenue."""
        utility_credit_value = self.community_solar.get('utility_credit_value_usd_per_mwh', 0.0)
        subscriber_discount = self.community_solar.get('subscriber_discount_pct', 10.0) / 100.0
        mgmt_fee_per_acct = self.community_solar.get('mgmt_fee_usd_per_acct_month', 0.0)
        annual_churn = self.community_solar.get('annual_churn_pct', 6.0) / 100.0
        bad_debt_pct = self.community_solar.get('bad_debt_pct_of_billings', 1.5) / 100.0
        ramp_months = self.community_solar.get('ramp_to_full_subscribed_months', 12)
        acquisition_cost = self.community_solar.get('acquisition_cost_usd_per_subscriber', 0.0)

        # Simplified: assume 100 subscribers at full subscription
        total_subscribers = 100

        subscriber_revenue = []
        mgmt_fees = []
        acquisition_costs = []
        bad_debt = []

        for idx, row in df.iterrows():
            month_in_op = row['month_in_operation']

            if month_in_op > 0:
                # Ramp factor
                if month_in_op <= ramp_months:
                    ramp_factor = month_in_op / ramp_months
                else:
                    ramp_factor = 1.0

                # Subscription level (accounting for churn)
                # Simplified: maintain ~95% subscription after ramp
                subscription_level = ramp_factor * 0.95

                # Gross bill credit value
                gross_credit = row['ac_mwh'] * utility_credit_value * 1000  # Convert to kWh basis

                # Subscriber payment (after discount)
                subscriber_payment = gross_credit * (1 - subscriber_discount) * subscription_level

                # Bad debt
                bad_debt_amount = subscriber_payment * bad_debt_pct

                # Net subscriber revenue
                net_sub_revenue = subscriber_payment - bad_debt_amount

                # Management fees
                active_subscribers = total_subscribers * subscription_level
                mgmt_fee = mgmt_fee_per_acct * active_subscribers

                # Acquisition cost (during ramp)
                if month_in_op <= ramp_months:
                    # Simplified: spread acquisition cost over ramp period
                    acquisition = (total_subscribers * acquisition_cost) / ramp_months
                else:
                    # Ongoing churn replacement
                    monthly_churn = annual_churn / 12.0
                    subscribers_to_replace = total_subscribers * monthly_churn
                    acquisition = subscribers_to_replace * acquisition_cost

                subscriber_revenue.append(net_sub_revenue)
                mgmt_fees.append(mgmt_fee)
                acquisition_costs.append(acquisition)
                bad_debt.append(bad_debt_amount)

            else:
                subscriber_revenue.append(0.0)
                mgmt_fees.append(0.0)
                acquisition_costs.append(0.0)
                bad_debt.append(0.0)

        df['subscriber_revenue'] = subscriber_revenue
        df['cdg_mgmt_fees'] = mgmt_fees
        df['cdg_acquisition_costs'] = acquisition_costs
        df['cdg_bad_debt'] = bad_debt

        return df

    def _calculate_ppa_revenue(self, df: pd.DataFrame) -> pd.DataFrame:
        """Calculate PPA revenue."""
        ppa_price = self.ppa.get('ppa_price_usd_per_mwh', 0.0)
        ppa_escalator = self.ppa.get('ppa_escalator_pct', 0.0) / 100.0
        ppa_term_years = self.ppa.get('ppa_term_years', 20)

        ppa_revenue = []
        merchant_revenue = []

        for idx, row in df.iterrows():
            month_in_op = row['month_in_operation']
            year_in_op = row['year_in_operation']

            if month_in_op > 0:
                if year_in_op <= ppa_term_years:
                    # PPA term
                    escalation_factor = (1 + ppa_escalator) ** (year_in_op - 1)
                    current_ppa_price = ppa_price * escalation_factor
                    revenue = row['ac_mwh'] * current_ppa_price
                    ppa_revenue.append(revenue)
                    merchant_revenue.append(0.0)
                else:
                    # Merchant tail (use last PPA price as default)
                    escalation_factor = (1 + ppa_escalator) ** (ppa_term_years - 1)
                    merchant_price = ppa_price * escalation_factor
                    revenue = row['ac_mwh'] * merchant_price
                    ppa_revenue.append(0.0)
                    merchant_revenue.append(revenue)
            else:
                ppa_revenue.append(0.0)
                merchant_revenue.append(0.0)

        df['ppa_revenue'] = ppa_revenue
        df['merchant_revenue'] = merchant_revenue

        return df

    def _calculate_ny_vder_revenue(self, df: pd.DataFrame) -> pd.DataFrame:
        """Calculate NY VDER value stack and Community Credit."""
        ny_vder = self.program.get('ny_vder', {})

        if not ny_vder.get('enabled', False):
            return df

        value_stack = ny_vder.get('value_stack', {})
        community_credit = ny_vder.get('community_credit', {})
        ny_sun = ny_vder.get('ny_sun', {})

        # Value stack components
        energy_rate = value_stack.get('energy_usd_per_mwh', 0.0)
        env_rate = value_stack.get('env_usd_per_mwh', 0.0)
        icap_rate = value_stack.get('icap_usd_per_kw_year', 0.0)
        drv_schedule = value_stack.get('drv_schedule_usd_per_kw_year', [])
        lsrv_tranches = value_stack.get('lsrv_tranches', [])
        loss_factor = value_stack.get('loss_factor_pct', 3.0) / 100.0

        # Community credit
        cc_rate = community_credit.get('usd_per_kwh', 0.0)
        cc_term_years = community_credit.get('term_years', 25)
        cc_escalator = community_credit.get('escalator_pct', 0.0) / 100.0

        vder_energy = []
        vder_env = []
        vder_icap = []
        vder_drv = []
        vder_lsrv = []
        community_credit_rev = []

        for idx, row in df.iterrows():
            month_in_op = row['month_in_operation']
            year_in_op = row['year_in_operation']

            if month_in_op > 0:
                # Energy and Environmental
                energy_rev = row['ac_mwh'] * energy_rate * (1 + loss_factor)
                env_rev = row['ac_mwh'] * env_rate

                # ICAP (capacity-based)
                icap_rev = (self.ac_kw * icap_rate / 12.0)

                # DRV (10-year declining schedule)
                if year_in_op <= len(drv_schedule):
                    drv_rate = drv_schedule[year_in_op - 1]
                    drv_rev = (self.ac_kw * drv_rate / 12.0)
                else:
                    drv_rev = 0.0

                # LSRV (tranches)
                lsrv_rev = 0.0
                for tranche in lsrv_tranches:
                    tranche_rate = tranche.get('usd_per_kw_year', 0.0)
                    tranche_years = tranche.get('years', 0)
                    if year_in_op <= tranche_years:
                        lsrv_rev += (self.ac_kw * tranche_rate / 12.0)

                # Community Credit
                if year_in_op <= cc_term_years:
                    escalation_factor = (1 + cc_escalator) ** (year_in_op - 1)
                    current_cc_rate = cc_rate * escalation_factor
                    cc_rev = row['ac_kwh'] * current_cc_rate
                else:
                    cc_rev = 0.0

                vder_energy.append(energy_rev)
                vder_env.append(env_rev)
                vder_icap.append(icap_rev)
                vder_drv.append(drv_rev)
                vder_lsrv.append(lsrv_rev)
                community_credit_rev.append(cc_rev)

            else:
                vder_energy.append(0.0)
                vder_env.append(0.0)
                vder_icap.append(0.0)
                vder_drv.append(0.0)
                vder_lsrv.append(0.0)
                community_credit_rev.append(0.0)

        df['vder_energy'] = vder_energy
        df['vder_environmental'] = vder_env
        df['vder_icap'] = vder_icap
        df['vder_drv'] = vder_drv
        df['vder_lsrv'] = vder_lsrv
        df['ny_community_credit'] = community_credit_rev

        # Override subscriber revenue if in CDG mode
        if self.mode == 'community_solar':
            # Calculate using VDER value stack as credit value
            df['vder_total_value'] = (df['vder_energy'] + df['vder_environmental'] +
                                      df['vder_icap'] + df['vder_drv'] + df['vder_lsrv'])

            # Recalculate subscriber revenue using VDER + Community Credit
            subscriber_discount = self.community_solar.get('subscriber_discount_pct', 10.0) / 100.0
            df['subscriber_revenue'] = (df['vder_total_value'] + df['ny_community_credit']) * (1 - subscriber_discount) * 0.95

        return df

    def _calculate_nj_csep_revenue(self, df: pd.DataFrame) -> pd.DataFrame:
        """Calculate NJ Community Solar Energy Pilot revenue."""
        nj_csep = self.program.get('nj_csep', {})

        if not nj_csep.get('enabled', False):
            return df

        customer_mix = nj_csep.get('customer_mix_pct', {})
        bill_credit_rates = nj_csep.get('bill_credit_rates_usd_per_kwh', {})
        subscriber_discounts = nj_csep.get('subscriber_discounts_pct', {})
        program_adders = nj_csep.get('program_adders', {})
        loss_factor = nj_csep.get('loss_factor_pct', 0.0) / 100.0

        # Blended bill credit rate
        residential_share = customer_mix.get('residential', 70.0) / 100.0
        small_comm_share = customer_mix.get('small_commercial', 30.0) / 100.0

        bcr_residential = bill_credit_rates.get('residential', 0.0)
        bcr_small_comm = bill_credit_rates.get('small_commercial', 0.0)

        blended_bcr = (residential_share * bcr_residential +
                       small_comm_share * bcr_small_comm)

        # Adders
        lmi_adder = program_adders.get('lmi_usd_per_mwh', 0.0)
        siting_adder = program_adders.get('siting_usd_per_mwh', 0.0)

        nj_bill_credit = []
        nj_adders = []

        for idx, row in df.iterrows():
            month_in_op = row['month_in_operation']

            if month_in_op > 0:
                # Bill credit revenue
                credit_value = row['ac_kwh'] * blended_bcr * (1 - loss_factor)

                # Apply subscriber discount (LMI/non-LMI weighted)
                lmi_share = self.community_solar.get('lmi_share_pct', 20.0) / 100.0
                lmi_discount = subscriber_discounts.get('lmi', 15.0) / 100.0
                non_lmi_discount = subscriber_discounts.get('non_lmi', 10.0) / 100.0

                blended_discount = (lmi_share * lmi_discount +
                                   (1 - lmi_share) * non_lmi_discount)

                subscriber_payment = credit_value * (1 - blended_discount) * 0.95

                # Adders
                adder_revenue = row['ac_mwh'] * (lmi_adder + siting_adder)

                nj_bill_credit.append(subscriber_payment)
                nj_adders.append(adder_revenue)

            else:
                nj_bill_credit.append(0.0)
                nj_adders.append(0.0)

        df['nj_bill_credit_revenue'] = nj_bill_credit
        df['nj_program_adders'] = nj_adders

        # Override subscriber revenue
        if self.mode == 'community_solar':
            df['subscriber_revenue'] = df['nj_bill_credit_revenue']

        return df

    def _calculate_il_abp_revenue(self, df: pd.DataFrame) -> pd.DataFrame:
        """Calculate Illinois Adjustable Block Program REC revenue."""
        il_abp = self.program.get('il_abp', {})

        if not il_abp.get('enabled', False):
            return df

        rec_price = il_abp.get('rec_price_usd', 0.0)
        rec_term_years = il_abp.get('rec_term_years', 15)
        payment_schedule = il_abp.get('payment_schedule_pct_by_year', [])
        brownfield_adder = il_abp.get('brownfield_or_site_adder_usd_per_mwh', 0.0)
        admin_fees_pct = il_abp.get('admin_fees_pct', 0.0) / 100.0

        # Normalize payment schedule
        if payment_schedule:
            total_pct = sum(payment_schedule)
            if total_pct > 0:
                payment_schedule = [p / total_pct for p in payment_schedule]

        il_rec_revenue = []
        il_brownfield_adder = []
        il_admin_fees = []

        # Calculate total RECs over term
        total_recs_by_year = {}

        for idx, row in df.iterrows():
            year_in_op = row['year_in_operation']
            if year_in_op > 0 and year_in_op <= rec_term_years:
                if year_in_op not in total_recs_by_year:
                    total_recs_by_year[year_in_op] = 0.0
                total_recs_by_year[year_in_op] += row['ac_mwh']

        # Calculate total REC value
        total_rec_value = sum(total_recs_by_year.values()) * rec_price

        # Allocate by payment schedule
        for idx, row in df.iterrows():
            month_in_op = row['month_in_operation']
            year_in_op = row['year_in_operation']

            if month_in_op > 0 and year_in_op <= rec_term_years:
                year_idx = year_in_op - 1

                if year_idx < len(payment_schedule):
                    annual_pct = payment_schedule[year_idx]
                else:
                    annual_pct = 0.0

                # Monthly allocation (evenly over 12 months)
                monthly_rec_revenue = (total_rec_value * annual_pct) / 12.0

                # Brownfield adder
                brownfield = row['ac_mwh'] * brownfield_adder

                # Admin fees
                admin_fee = monthly_rec_revenue * admin_fees_pct

                il_rec_revenue.append(monthly_rec_revenue)
                il_brownfield_adder.append(brownfield)
                il_admin_fees.append(admin_fee)

            else:
                il_rec_revenue.append(0.0)
                il_brownfield_adder.append(0.0)
                il_admin_fees.append(0.0)

        df['il_abp_rec_revenue'] = il_rec_revenue
        df['il_abp_brownfield_adder'] = il_brownfield_adder
        df['il_abp_admin_fees'] = il_admin_fees

        return df

    def _calculate_total_revenue(self, df: pd.DataFrame):
        """Calculate total revenue from all streams."""
        # Identify all revenue columns
        revenue_columns = [col for col in df.columns if 'revenue' in col.lower() or
                          col in ['vder_energy', 'vder_environmental', 'vder_icap',
                                 'vder_drv', 'vder_lsrv', 'ny_community_credit',
                                 'nj_bill_credit_revenue', 'nj_program_adders',
                                 'il_abp_rec_revenue', 'il_abp_brownfield_adder']]

        # Exclude cost columns
        exclude = ['merchant_revenue']  # Will handle separately

        revenue_cols = [col for col in revenue_columns if col not in exclude and col != 'total_revenue']

        if revenue_cols:
            df['total_revenue'] = df[revenue_cols].sum(axis=1)
        else:
            df['total_revenue'] = 0.0

        return df

    def get_upfront_grants(self) -> Tuple[float, str]:
        """
        Calculate upfront grants from state programs.

        Returns:
            (grant_amount, timing)
        """
        program_active = self.program.get('active', 'none')

        if program_active == 'ny_vder':
            ny_sun = self.program.get('ny_vder', {}).get('ny_sun', {})
            if ny_sun.get('icsa_enabled', False):
                icsa_rate = ny_sun.get('icsa_usd_per_wdc', 0.0)
                dc_kw = self.inputs['sizing']['dc_kw']
                grant_amount = icsa_rate * dc_kw
                timing = ny_sun.get('grant_timing', 'COD')
                return grant_amount, timing

            upfront_rate = ny_sun.get('upfront_other_usd_per_wdc', 0.0)
            if upfront_rate > 0:
                dc_kw = self.inputs['sizing']['dc_kw']
                grant_amount = upfront_rate * dc_kw
                timing = ny_sun.get('grant_timing', 'COD')
                return grant_amount, timing

        elif program_active == 'nj_csep':
            nj_csep = self.program.get('nj_csep', {})
            upfront_rate = nj_csep.get('program_adders', {}).get('upfront_usd_per_wdc', 0.0)
            if upfront_rate > 0:
                dc_kw = self.inputs['sizing']['dc_kw']
                return upfront_rate * dc_kw, 'COD'

        elif program_active == 'il_abp':
            il_abp = self.program.get('il_abp', {})
            inverter_rebate = il_abp.get('smart_inverter_rebate_usd_per_kwac', 0.0)
            if inverter_rebate > 0:
                return inverter_rebate * self.ac_kw, 'COD'

        return 0.0, 'COD'
