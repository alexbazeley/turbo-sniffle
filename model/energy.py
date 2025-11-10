"""
Energy production calculation module.
Handles 8760 data or synthetic generation with degradation, availability, curtailment.
"""

import pandas as pd
import numpy as np
from typing import Dict, Any, Optional
from datetime import datetime, timedelta
from dateutil.relativedelta import relativedelta


class EnergyModel:
    """Calculates monthly energy production over project life."""

    def __init__(self, inputs: Dict[str, Any]):
        self.inputs = inputs
        self.sizing = inputs['sizing']
        self.project = inputs['project']

        # Parse COD date
        self.cod_date = datetime.strptime(inputs['project']['cod_date'], '%Y-%m-%d')

        # Total months from NTP to end of life
        self.construction_months = inputs['project']['construction_months']
        self.operating_months = inputs['project']['model_years'] * 12
        self.total_months = self.construction_months + self.operating_months

    def calculate_monthly_energy(self) -> pd.DataFrame:
        """
        Calculate monthly energy production.

        Returns:
            DataFrame with columns: period, date, month_in_operation, ac_kwh, dc_kwh
        """
        # Start date is NTP (construction months before COD)
        ntp_date = self.cod_date - relativedelta(months=self.construction_months)

        periods = []
        dates = []
        months_in_operation = []

        for i in range(self.total_months):
            period_date = ntp_date + relativedelta(months=i)
            periods.append(i)
            dates.append(period_date)

            # Month in operation (0 during construction, 1+ after COD)
            if i < self.construction_months:
                months_in_operation.append(0)
            else:
                months_in_operation.append(i - self.construction_months + 1)

        df = pd.DataFrame({
            'period': periods,
            'date': dates,
            'month_in_operation': months_in_operation,
            'year_in_operation': [m // 12 + 1 if m > 0 else 0 for m in months_in_operation]
        })

        # Calculate energy production
        if self.sizing['use_8760'] and self.sizing['8760_csv_path']:
            df = self._calculate_from_8760(df)
        else:
            df = self._calculate_synthetic(df)

        return df

    def _calculate_synthetic(self, df: pd.DataFrame) -> pd.DataFrame:
        """Calculate energy using capacity factor method."""
        dc_kw = self.sizing['dc_kw']
        ac_kw = self.sizing['ac_kw']
        capacity_factor = self.sizing.get('capacity_factor_pct', 20.0) / 100.0
        performance_ratio = self.sizing['performance_ratio_pct'] / 100.0
        availability = self.sizing['availability_pct'] / 100.0
        curtailment = self.sizing['curtailment_pct'] / 100.0
        degradation_rate = self.sizing['degradation_pct_per_year'] / 100.0

        # Base annual energy (year 1)
        hours_per_year = 8760
        base_annual_kwh = ac_kw * capacity_factor * hours_per_year * performance_ratio

        energy_ac = []
        energy_dc = []

        for idx, row in df.iterrows():
            month_in_op = row['month_in_operation']

            if month_in_op == 0:
                # Construction period - no production
                energy_ac.append(0.0)
                energy_dc.append(0.0)
            else:
                # Operating period
                year_in_op = row['year_in_operation']

                # Apply degradation (linear)
                degradation_factor = 1.0 - (degradation_rate * (year_in_op - 1))

                # Days in month
                year = row['date'].year
                month = row['date'].month
                if month == 12:
                    next_month_date = datetime(year + 1, 1, 1)
                else:
                    next_month_date = datetime(year, month + 1, 1)
                days_in_month = (next_month_date - datetime(year, month, 1)).days

                # Monthly energy (proportional to days, adjusted for degradation)
                monthly_fraction = days_in_month / 365.25
                monthly_ac_kwh = (base_annual_kwh * monthly_fraction *
                                  degradation_factor * availability * (1 - curtailment))

                # DC energy (AC / inverter efficiency, simplified as DC/AC ratio)
                monthly_dc_kwh = monthly_ac_kwh * self.sizing['dc_ac_ratio']

                energy_ac.append(monthly_ac_kwh)
                energy_dc.append(monthly_dc_kwh)

        df['ac_kwh'] = energy_ac
        df['dc_kwh'] = energy_dc
        df['ac_mwh'] = df['ac_kwh'] / 1000.0

        return df

    def _calculate_from_8760(self, df: pd.DataFrame) -> pd.DataFrame:
        """Calculate energy from 8760 CSV data."""
        # Load 8760 data
        hourly_data = pd.read_csv(self.sizing['8760_csv_path'])

        # Expected columns: timestamp, ac_kwh (or ac_kw)
        if 'ac_kwh' not in hourly_data.columns and 'ac_kw' in hourly_data.columns:
            hourly_data['ac_kwh'] = hourly_data['ac_kw']

        # Ensure we have timestamp
        if 'timestamp' in hourly_data.columns:
            hourly_data['timestamp'] = pd.to_datetime(hourly_data['timestamp'])
        else:
            # Create synthetic hourly timestamps for one year
            start = datetime(2020, 1, 1)
            hourly_data['timestamp'] = [start + timedelta(hours=i) for i in range(len(hourly_data))]

        # Add month column
        hourly_data['month'] = hourly_data['timestamp'].dt.month

        # Aggregate to monthly
        monthly_base = hourly_data.groupby('month')['ac_kwh'].sum().to_dict()

        availability = self.sizing['availability_pct'] / 100.0
        curtailment = self.sizing['curtailment_pct'] / 100.0
        degradation_rate = self.sizing['degradation_pct_per_year'] / 100.0

        energy_ac = []
        energy_dc = []

        for idx, row in df.iterrows():
            month_in_op = row['month_in_operation']

            if month_in_op == 0:
                energy_ac.append(0.0)
                energy_dc.append(0.0)
            else:
                year_in_op = row['year_in_operation']
                month_num = row['date'].month

                # Get base monthly energy from 8760
                base_monthly_kwh = monthly_base.get(month_num, 0.0)

                # Apply degradation
                degradation_factor = 1.0 - (degradation_rate * (year_in_op - 1))

                monthly_ac_kwh = base_monthly_kwh * degradation_factor * availability * (1 - curtailment)
                monthly_dc_kwh = monthly_ac_kwh * self.sizing['dc_ac_ratio']

                energy_ac.append(monthly_ac_kwh)
                energy_dc.append(monthly_dc_kwh)

        df['ac_kwh'] = energy_ac
        df['dc_kwh'] = energy_dc
        df['ac_mwh'] = df['ac_kwh'] / 1000.0

        return df
