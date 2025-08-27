import pandas as pd
import numpy as np


def weekly_stat_analysis(df_all):
    """
    Analyzes weekly statistics to find interesting periods based on demand,
    renewable energy generation, duck curve severity, and ramping requirements.

    This function identifies weeks with:
    - High Demand & High Renewable (RE) generation
    - High Demand & Low RE generation
    - Severe Duck Curves (large midday dip followed by a steep evening ramp)
    - High Ramping Needs (high volatility in net demand)

    Args:
        df_all (pd.DataFrame): A DataFrame with 15-minute interval data, including
                               'TOTAL DEMAND', 'renewable', and 'NET DEMAND' columns.

    Returns:
        tuple: A tuple containing:
            - weekly_stats (pd.DataFrame): A DataFrame with the calculated weekly statistics.
            - interesting_weeks (dict): A dictionary listing the start dates of weeks for each interesting category.
    """
    # --- 1. Calculate Daily Metrics First ---
    # The 'NET DEMAND' column is crucial for duck curve and ramping analysis.
    if 'NET DEMAND' not in df_all.columns:
        raise ValueError("Input DataFrame must contain a 'NET DEMAND' column.")

    daily_resampler = df_all.resample('D')

    # Define a function to calculate the magnitude of the duck curve for a given day.
    def get_duck_magnitude(day_df):
        if day_df.empty:
            return np.nan

        # Define time windows for the midday solar peak and the evening demand peak.
        midday_window = day_df.between_time('10:00', '15:45')
        evening_window = day_df.between_time('17:00', '21:45')

        # Proceed only if there's data in both windows.
        if midday_window.empty or evening_window.empty:
            return np.nan

        midday_min_net_demand = midday_window['NET DEMAND'].min()
        evening_peak_net_demand = evening_window['NET DEMAND'].max()

        # The "duck belly to head" height.
        return evening_peak_net_demand - midday_min_net_demand

    # Apply the functions to get daily values.
    daily_duck_magnitude = daily_resampler.apply(get_duck_magnitude)
    daily_max_ramp = daily_resampler['NET DEMAND'].apply(lambda day: day.diff().abs().max())

    # --- 2. Aggregate Daily Metrics into Weekly Stats ---
    weekly_stats = pd.DataFrame()
    weekly_stats['avg_demand'] = df_all['TOTAL DEMAND'].resample('W').mean()
    weekly_stats['avg_renewable'] = df_all['renewable'].resample('W').mean()
    weekly_stats['avg_duck_magnitude'] = daily_duck_magnitude.resample('W').mean()
    weekly_stats['avg_max_ramp'] = daily_max_ramp.resample('W').mean()
    weekly_stats['net_demand_std'] = df_all['NET DEMAND'].resample('W').std()  # Proxy for overall volatility.

    weekly_stats.dropna(inplace=True)  # Ensure all weeks have valid data.

    # --- 3. Identify Interesting Weeks using Quantiles ---
    # We define "high" as the top 10% (90th percentile) and "low" as the bottom 10%.
    high_threshold = weekly_stats.quantile(0.90)
    low_threshold = weekly_stats.quantile(0.10)

    # Filter weeks based on the thresholds.
    high_demand_weeks = weekly_stats[weekly_stats['avg_demand'] > high_threshold['avg_demand']]
    low_re_weeks = weekly_stats[weekly_stats['avg_renewable'] < low_threshold['avg_renewable']]
    high_re_weeks = weekly_stats[weekly_stats['avg_renewable'] > high_threshold['avg_renewable']]

    # Find weeks that meet the combined criteria.
    high_demand_high_re = high_demand_weeks.index.intersection(high_re_weeks.index)
    high_demand_low_re = high_demand_weeks.index.intersection(low_re_weeks.index)

    high_duck_curve_weeks = weekly_stats[
        weekly_stats['avg_duck_magnitude'] > high_threshold['avg_duck_magnitude']].index
    high_ramping_weeks = weekly_stats[weekly_stats['avg_max_ramp'] > high_threshold['avg_max_ramp']].index

    # --- 4. Prepare the Output ---
    interesting_weeks = {
        "High Demand & High RE": high_demand_high_re.strftime('%Y-%m-%d').tolist(),
        "High Demand & Low RE": high_demand_low_re.strftime('%Y-%m-%d').tolist(),
        "High Duck Curve": high_duck_curve_weeks.strftime('%Y-%m-%d').tolist(),
        "High Ramping Requirements": high_ramping_weeks.strftime('%Y-%m-%d').tolist(),
    }
    return weekly_stats, interesting_weeks


def battery_fixed_size_calculations(df_filtered, min_batt_soc, batt_efficiency, battery_configs):
    # This function remains unchanged.

    # Define battery storage parameters (MW size needs to be provided)

    # Extract relevant time-series from df_filtered
    time_series = df_filtered.index
    original_surplus = df_filtered["WITH SURPLUS"].values  # Demand (+) and surplus (-)

    # Initialize storage tracking
    battery_profiles = {}
    remaining_surplus_history = {}  # Store remaining_surplus after each battery
    remaining_surplus = original_surplus.copy()  # Make a copy to update after each battery

    for battery_name, config in battery_configs.items():
        power = config["power"]
        capacity = power * config["duration"]

        # Initialize tracking variables
        battery_energy = 0
        charge_profile = []
        discharge_profile = []
        battery_state = []

        # Iterate over time intervals (15-minute resolution)
        for i, surplus in enumerate(remaining_surplus):
            if surplus < 0:  # Charging condition
                if battery_energy < capacity:
                    charge = min(abs(surplus), power)
                    charge = min(charge, (capacity - battery_energy) * 0.25)
                else:
                    charge = 0
                battery_energy = min(battery_energy + charge * 0.25 * batt_efficiency, capacity)
                remaining_surplus[i] += charge  # Reduce surplus (since charging absorbs it)
                discharge = 0
            elif surplus > 0:  # Discharging condition
                if battery_energy > capacity * min_batt_soc:
                    discharge = min(surplus, power)
                    discharge = min(discharge, (battery_energy - capacity * min_batt_soc) / 0.25)
                else:
                    discharge = 0
                battery_energy = battery_energy - discharge * 0.25
                remaining_surplus[i] -= discharge * batt_efficiency  # Reduce demand (since discharging supplies it)
                charge = 0
            else:
                charge = 0
                discharge = 0

            # Store results
            charge_profile.append(charge)
            discharge_profile.append(discharge)
            battery_state.append(battery_energy)

        # Store profiles in DataFrame
        battery_profiles[battery_name] = pd.DataFrame({
            "Timestamp": time_series,
            "Charge (MW)": charge_profile,
            "Discharge (MW)": discharge_profile,
            "Battery State (MWh)": battery_state
        }).set_index("Timestamp")

        # Store remaining surplus after this battery processes it
        remaining_surplus_history[battery_name] = remaining_surplus.copy()

    return battery_profiles, remaining_surplus_history