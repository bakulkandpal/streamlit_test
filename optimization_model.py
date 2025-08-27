# In file: optimization_model.py
import os
import sys
from pyomo.environ import (ConcreteModel, Set, Param, Var, Constraint, Objective,
                           NonNegativeReals, SolverFactory, RangeSet, Binary,
                           minimize, value)
from my_statistics import weekly_stat_analysis, battery_fixed_size_calculations
import pandas as pd
import configparser
import logging


def setup_logging():
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s [%(levelname)s] %(message)s",
        handlers=[logging.StreamHandler(sys.stdout)]
    )


setup_logging()


def read_config(config_file='parameters.ini'):
    """Read all configuration parameters from the config file (robust version)."""
    config = configparser.ConfigParser()
    config.read(config_file)
    params = {}

    # --- Step 1: Read all parameters from all sections into a flat dictionary ---
    for section in config.sections():
        for key, val in config.items(section):
            try:
                if key in ['allow_oversized_re', 'run_thermal_&_sizing_optimization']:
                    params[key] = config.getboolean(section, key)
                elif key == 'shortage_case':
                    params[key] = val
                else:
                    # Try to convert to float, otherwise keep as string
                    params[key] = float(val)
            except (ValueError, TypeError):
                params[key] = val

    # --- Step 2: Build the nested battery dictionary from the params we just read ---
    params['battery_configs'] = {}
    battery_count = 1
    # **This is the key change**: We check for the key in the `params` dictionary
    # that we already safely loaded, not in the raw config object.
    while f"battery{battery_count}_power" in params:
        power_key = f"battery{battery_count}_power"
        duration_key = f"battery{battery_count}_duration"

        power = params[power_key]
        duration = params[duration_key]

        params['battery_configs'][f"Battery {battery_count}"] = {"power": power, "duration": duration}
        battery_count += 1

    return params


def run_optimization(config_file='parameters.ini'):
    logging.info(f"*** Reading Configuration from {config_file} ***")
    try:
        params = read_config(config_file)
        logging.info("*** Configuration Loaded Successfully ***")
    except Exception as e:
        logging.error(f"Error loading configuration: {e}")
        return

    script_dir = os.path.dirname(os.path.abspath(__file__))
    results_dir = os.path.join(script_dir, 'Results')
    if not os.path.exists(results_dir):
        os.makedirs(results_dir)

    wind_size_excel_SRI = params['wind_size_excel_sri']
    wind_size_excel_SECI = params['wind_size_excel_seci']
    wind_size_actual_SRI = params['wind_size_actual_sri']
    wind_size_actual_SECI = params['wind_size_actual_seci']
    wind_karnataka = params['wind_size_karnataka']
    wind_tamil = params['wind_size_tamil']
    PV_size_gujarat = params['pv_size_gujarat']
    PV_size_telangana = params['pv_size_telangana']
    PV_size_rajasthan = params['pv_size_rajasthan']
    PV_size_actual_goa = params['pv_size_goa']
    wind_maharashtra_goa = params['wind_size_goa_or_maharashtra']
    DRE_size = params['dre_size_goa']
    Nuclear_size = params['nuclear_size']
    Biomass_size = params['biomass_size']
    Gas_size = params['gas_size']
    RTC_size = params['rtc_size']
    annual_demand_mus = params['annual_demand_mus']
    intra_state_losses = params['intra_state_power_losses']
    inter_state_losses = params['inter_state_power_losses']
    min_batt_soc = params['min_batt_soc']
    batt_efficiency = params['batt_efficiency']
    battery_configs = params['battery_configs']
    start_date = params['timeline_start_date']
    end_date = params['timeline_end_date']
    zero_start_date = params['zero_pv_goa_start_date']
    zero_end_date = params['zero_pv_goa_end_date']
    zero_start_date2 = params['zero_pv_goa_start_date2']
    zero_end_date2 = params['zero_pv_goa_end_date2']
    solar_cost_goa = params['solar_cost_goa']
    solar_cost_guj = params['solar_cost_guj']
    solar_cost_raj = params['solar_cost_raj']
    solar_cost_tel = params['solar_cost_tel']
    wind_cost_maha = params['wind_cost_maha']
    wind_cost_tamil = params['wind_cost_tamil']
    wind_cost_karnataka = params['wind_cost_karnataka']
    battery_cost_MWh = params['battery_cost_mwh']
    penalty_thermal_unmet_demand = params['penalty_thermal_unmet_demand']
    penalty_sizing_unmet_demand = params['penalty_sizing_unmet_demand']
    # pen_charge_discharge = params['pen_charge_discharge']
    max_size_batt_mwh = params['max_size_batt_mwh']
    max_charge_discharge_power_bess = params['max_charge_discharge_power_bess']
    allow_oversized_RE = params['allow_oversized_re']
    run_thermal_sizing_optimization = params['run_thermal_&_sizing_optimization']
    shortage_case = params['shortage_case']
    gdam_price_select_year = params['gdam_price_select_year']
    max_gdam_purchase = params['max_gdam_purchase']
    max_solar_goa = params['max_solar_goa']
    min_solar_goa = params['min_solar_goa']
    min_solar_guj = params['min_solar_guj']
    min_solar_raj = params['min_solar_raj']
    min_solar_tel = params['min_solar_tel']
    min_wind_maha = params['min_wind_maha']
    min_wind_tamil = params['min_wind_tamil']
    min_wind_karnataka = params['min_wind_karnataka']

    file_path = os.path.join(script_dir, params['file_path']) if not os.path.isabs(params['file_path']) else params[
        'file_path']
    file_path_wind_SRI = os.path.join(script_dir, params['file_path_wind_sri']) if not os.path.isabs(
        params['file_path_wind_sri']) else params['file_path_wind_sri']
    file_path_wind_SECI = os.path.join(script_dir, params['file_path_wind_seci']) if not os.path.isabs(
        params['file_path_wind_seci']) else params['file_path_wind_seci']
    file_path_solar_goa = os.path.join(script_dir, params['file_path_solar_goa']) if not os.path.isabs(
        params['file_path_solar_goa']) else params['file_path_solar_goa']
    file_path_solar_gujarat = os.path.join(script_dir, params['file_path_solar_gujarat']) if not os.path.isabs(
        params['file_path_solar_gujarat']) else params['file_path_solar_gujarat']
    file_path_solar_rajasthan = os.path.join(script_dir, params['file_path_solar_rajasthan']) if not os.path.isabs(
        params['file_path_solar_rajasthan']) else params['file_path_solar_rajasthan']
    file_path_solar_given = os.path.join(script_dir, params['file_path_solar_given']) if not os.path.isabs(
        params['file_path_solar_given']) else params['file_path_solar_given']
    file_path_generators = os.path.join(script_dir, params['file_path_generators']) if not os.path.isabs(
        params['file_path_generators']) else params['file_path_generators']
    file_path_solar_telangana = os.path.join(script_dir, params['file_path_solar_telangana']) if not os.path.isabs(
        params['file_path_solar_telangana']) else params['file_path_solar_telangana']
    file_path_shortage_case1 = os.path.join(script_dir, params['file_path_shortage_case1']) if not os.path.isabs(
        params['file_path_shortage_case1']) else params['file_path_shortage_case1']
    file_path_shortage_case2 = os.path.join(script_dir, params['file_path_shortage_case2']) if not os.path.isabs(
        params['file_path_shortage_case1']) else params['file_path_shortage_case1']
    file_path_gdam = os.path.join(script_dir, params['file_path_gdam']) if not os.path.isabs(
        params['file_path_gdam']) else params['file_path_gdam']

    # Check if data files exist
    for path in [file_path, file_path_wind_SRI, file_path_wind_SECI, file_path_solar_goa,
                 file_path_solar_gujarat, file_path_solar_rajasthan, file_path_solar_given,
                 file_path_generators, file_path_solar_telangana]:
        if not os.path.exists(path):
            logging.info("*** Configuration Loaded Successfully ***")
            logging.info(f"Error: File not found: {path}")
            input("Press Enter to exit...")
            return

    demand_scaling_factor = annual_demand_mus / 7471  # 7471 is the original annual MUs considered for FY30 (based on CEA estimate)

    df_gdam_price = pd.read_excel(file_path_gdam)

    ########## SOLAR & WIND
    logging.info("*** Reading Solar and Wind Data Files *** \n")
    df_solar_wind_2022 = pd.read_csv(file_path_solar_given)  # From daily sheet which 'total demand' is also taken
    df_solar_wind_positive = df_solar_wind_2022.copy()
    ### Reducing negative generation to 0 to calculate actual net demand later.
    numeric_cols = df_solar_wind_positive.select_dtypes(include=['number']).columns
    df_solar_wind_positive[numeric_cols] = df_solar_wind_positive[numeric_cols].applymap(lambda x: max(x, 0))
    # Compute total solar power of 2022 (combining STOA, LTA, MTOA)
    df_solar_wind_positive['TOTAL SOLAR'] = df_solar_wind_2022['MTOA SOLAR'] + df_solar_wind_2022['LTA SOLAR'] + \
                                            df_solar_wind_2022['STOA SOLAR']
    df_solar_wind_positive['TOTAL RENEWABLE'] = df_solar_wind_positive['TOTAL SOLAR'] + df_solar_wind_positive[
        'NON SOLAR ( WIND / HYDRO)']

    # Load Demand Data
    logging.info("*** Reading Demand Data File *** \n")
    df_demand = pd.read_csv(file_path, parse_dates=['Timestamp'], dayfirst=True)
    df_demand['Timestamp'] = pd.to_datetime(df_demand['Timestamp'], format='%d-%m-%Y %H:%M:%S')
    df_demand['TOTAL DEMAND'] = pd.to_numeric(df_demand['TOTAL DEMAND'].astype(str).str.replace(',', '').str.strip(),
                                              errors='coerce')
    df_demand = df_demand[(df_demand['Timestamp'] >= '2022-01-01') & (df_demand['Timestamp'] < '2023-01-01')]
    df_demand.set_index('Timestamp', inplace=True)

    # Ensure the 'Timestamp' column is the index and is in datetime format
    df_demand.index = pd.to_datetime(df_demand.index)

    # Resample the data to monthly frequency and sum the total demand for each month
    monthly_energy_consumption = df_demand['TOTAL DEMAND'].resample('M').sum() / 4

    # Resample the data to daily frequency and sum the total demand for each day
    daily_energy_consumption = df_demand['TOTAL DEMAND'].resample('D').sum() / 4

    def calculate_daily_energy_consumption(target_year, daily_energy_consumption_base,
                                           monthly_energy_consumption_target):
        # Ensure the 'Timestamp' column is the index and is in datetime format
        daily_energy_consumption_base.index = pd.to_datetime(daily_energy_consumption_base.index)

        # Calculate the monthly energy consumption for the base year (dividing by 1000 because monthly-target will be in Million Units - MUs)
        monthly_energy_consumption_base = daily_energy_consumption_base.resample('M').sum() / 1000

        # Calculate the daily energy consumption for the target year
        daily_energy_consumption_target = daily_energy_consumption_base.copy()
        for month in monthly_energy_consumption_target.index:
            month_base = month.replace(str(target_year), str(daily_energy_consumption_base.index.year[0]))
            daily_energy_consumption_target.loc[
                daily_energy_consumption_base.index.month == pd.to_datetime(month).month] = (
                    daily_energy_consumption_base.loc[
                        daily_energy_consumption_base.index.month == pd.to_datetime(month_base).month] *
                    monthly_energy_consumption_target[month] /
                    monthly_energy_consumption_base[month_base]
            )

        # Adjust the timestamps to reflect the target year
        daily_energy_consumption_target.index = daily_energy_consumption_target.index.map(
            lambda x: x.replace(year=target_year))

        ratio = daily_energy_consumption_target.values / daily_energy_consumption_base.values

        return daily_energy_consumption_target, ratio

    # Example usage:
    # Assume reasonable numbers for the monthly energy consumption for 2028
    monthly_energy_consumption_2030 = pd.Series({
        '2030-01-31': 633, '2030-02-28': 598, '2030-03-31': 673, '2030-04-30': 685,
        '2030-05-31': 716, '2030-06-30': 641, '2030-07-31': 581, '2030-08-31': 486,
        '2030-09-30': 594, '2030-10-31': 615, '2030-11-30': 610, '2030-12-31': 635
    })
    monthly_energy_consumption_2030 = monthly_energy_consumption_2030 * demand_scaling_factor

    # Calculate the daily energy consumption for future year
    daily_energy_consumption_2030, ratio_2030 = calculate_daily_energy_consumption(2030, daily_energy_consumption,
                                                                                   monthly_energy_consumption_2030)

    ratio_2030 = pd.Series(ratio_2030, index=pd.date_range(start='2030-01-01', periods=365, freq='D'))
    # Create full 15-minute interval index including last day
    full_index = pd.date_range(start='2030-01-01', end='2030-12-31 23:45:00', freq='15T')

    # Reindex and forward fill
    ratio_2030_resampled = ratio_2030.reindex(full_index, method='ffill')

    # Extract the TOTAL DEMAND values
    total_demand_values = df_demand['TOTAL DEMAND'].values

    # Multiply the TOTAL DEMAND values with the resampled ratio_2028
    adjusted_demand_values = total_demand_values * ratio_2030_resampled.values

    # Create a new DataFrame with the adjusted demand values and updated timestamps
    target_year = pd.to_datetime(start_date).year
    df_demand_year = pd.DataFrame({
        'Timestamp': df_demand.index.map(lambda x: x.replace(year=target_year)),
        'TOTAL DEMAND': adjusted_demand_values
    })

    # Set the Timestamp as the index
    df_demand_year.set_index('Timestamp', inplace=True)

    # Load Wind Data (both SRI and SECI)
    df_wind_SRI = pd.read_excel(file_path_wind_SRI, sheet_name="Yearly data", engine="openpyxl")
    df_wind_SECI = pd.read_excel(file_path_wind_SECI, sheet_name="Yearly data", engine="openpyxl")

    # Rename First Column to "Date"
    df_wind_SRI.rename(columns={df_wind_SRI.columns[0]: "Date"}, inplace=True)
    # Convert "Date" to DateTime Format
    df_wind_SRI["Date"] = pd.to_datetime(df_wind_SRI["Date"], format="%d-%b-%y")
    # Convert Wide Format to Long Format
    df_wind_long = df_wind_SRI.melt(id_vars=["Date"], var_name="Time", value_name="Wind Production")
    # Extract Start Time from Time Column (Fixing Any Formatting Issues)
    df_wind_long["Time"] = df_wind_long["Time"].astype(str).str.extract(r"(\d{2}:\d{2})")
    # Create Full Timestamp
    df_wind_long["Timestamp"] = pd.to_datetime(
        df_wind_long["Date"].astype(str) + " " + df_wind_long["Time"], format="%Y-%m-%d %H:%M")
    # Keep Only Required Columns
    df_wind_long = df_wind_long[["Timestamp", "Wind Production"]].set_index("Timestamp")
    # Normalize Wind Data
    df_wind_long['Wind Production'] /= wind_size_excel_SRI  # Normalize Wind Production
    df_wind_long[
        'Wind Production'] *= wind_size_actual_SRI  # Increase Wind Production based on actual size to be considered
    df_wind_long = df_wind_long.sort_index()

    ##### Transformations for SECI wind data
    # Rename First Column to "Date"
    df_wind_SECI.rename(columns={df_wind_SECI.columns[0]: "Date"}, inplace=True)
    # Convert "Date" to DateTime Format
    df_wind_SECI["Date"] = pd.to_datetime(df_wind_SECI["Date"], format="%d-%b-%y")
    # Convert Wide Format to Long Format
    df_wind_long_SECI = df_wind_SECI.melt(id_vars=["Date"], var_name="Time", value_name="Wind Production")
    # Extract Start Time from Time Column (Fixing Any Formatting Issues)
    df_wind_long_SECI["Time"] = df_wind_long_SECI["Time"].astype(str).str.extract(r"(\d{2}:\d{2})")
    # Create Full Timestamp
    df_wind_long_SECI["Timestamp"] = pd.to_datetime(
        df_wind_long_SECI["Date"].astype(str) + " " + df_wind_long_SECI["Time"], format="%Y-%m-%d %H:%M")
    # Keep Only Required Columns
    df_wind_long_SECI = df_wind_long_SECI[["Timestamp", "Wind Production"]].set_index("Timestamp")
    # Normalize Wind Data
    df_wind_long_SECI['Wind Production'] /= wind_size_excel_SECI  # Normalize Wind Production
    df_wind_long_SECI[
        'Wind Production'] *= wind_size_actual_SECI  # Increase Wind Production based on actual size to be considered
    df_wind_long_SECI = df_wind_long_SECI.sort_index()

    # Load Solar Data
    df_solar_goa = pd.read_csv(file_path_solar_goa, parse_dates=["local_time"])
    df_solar_gujarat = pd.read_csv(file_path_solar_gujarat, parse_dates=["local_time"])
    df_solar_rajasthan = pd.read_csv(file_path_solar_rajasthan, parse_dates=["local_time"])
    df_solar_telangana = pd.read_csv(file_path_solar_telangana, parse_dates=["local_time"])

    # Ensure 'local_time' is datetime format
    df_solar_goa["local_time"] = pd.to_datetime(df_solar_goa["local_time"], format="%d-%m-%Y %H:%M", errors="coerce")
    # Create a full 15-minute timestamp range
    common_index = pd.date_range(
        start=df_solar_goa["local_time"].min().replace(minute=0),  # Start at 00:00
        end=df_solar_goa["local_time"].max().replace(hour=23, minute=45),  # End at 23:45
        freq="15min")
    # Reindex to match 15-minute intervals
    df_solar_goa = df_solar_goa.set_index("local_time").reindex(common_index)
    # Forward-fill to copy hourly values to missing 15-min slots
    df_solar_goa["electricity"] = df_solar_goa["electricity"].fillna(method="ffill").fillna(0)
    # Rename columns and set index
    df_solar_goa = df_solar_goa.rename(columns={"electricity": "Solar Production"})
    df_solar_goa.index.name = "Timestamp"
    df_solar_goa[
        'Solar Production']  # This is only Normalized Solar Production (1 MW Solar). See below for actual size PV production
    df_solar_goa = df_solar_goa.sort_index()
    ### To make PV generation zero for goa in specific dates
    solar_data_year = df_solar_goa.index[0].year
    # Convert zero_start_date and zero_end_date to datetime if they're strings
    zero_start_date_dt = pd.to_datetime(zero_start_date)
    zero_end_date_dt = pd.to_datetime(zero_end_date)

    zero_start_date_dt2 = pd.to_datetime(zero_start_date2)
    zero_end_date_dt2 = pd.to_datetime(zero_end_date2)

    # Replace the year in the dates with the year from df_solar_goa
    zero_start_date_adjusted = zero_start_date_dt.replace(year=solar_data_year)
    zero_end_date_adjusted = zero_end_date_dt.replace(year=solar_data_year)
    zero_start_date_adjusted2 = zero_start_date_dt2.replace(year=solar_data_year)
    zero_end_date_adjusted2 = zero_end_date_dt2.replace(year=solar_data_year)

    df_solar_goa.loc[zero_start_date_adjusted:zero_end_date_adjusted, 'Solar Production'] = 0
    df_solar_goa.loc[zero_start_date_adjusted2:zero_end_date_adjusted2, 'Solar Production'] = 0

    ####### Do transformations for Gujarat solar data
    # Ensure 'local_time' is datetime format
    df_solar_gujarat["local_time"] = pd.to_datetime(df_solar_gujarat["local_time"], format="%d-%m-%Y %H:%M",
                                                    errors="coerce")
    # Create a full 15-minute timestamp range
    common_index = pd.date_range(
        start=df_solar_gujarat["local_time"].min().replace(minute=0),  # Start at 00:00
        end=df_solar_gujarat["local_time"].max().replace(hour=23, minute=45),  # End at 23:45
        freq="15min")
    # Reindex to match 15-minute intervals
    df_solar_gujarat = df_solar_gujarat.set_index("local_time").reindex(common_index)
    # Forward-fill to copy hourly values to missing 15-min slots
    df_solar_gujarat["electricity"] = df_solar_gujarat["electricity"].fillna(method="ffill").fillna(0)
    # Rename columns and set index
    df_solar_gujarat = df_solar_gujarat.rename(columns={"electricity": "Solar Production"})
    df_solar_gujarat.index.name = "Timestamp"
    df_solar_gujarat[
        'Solar Production']  # This is only Normalized Solar Production (1 MW Solar). See below for actual size PV production
    df_solar_gujarat = df_solar_gujarat.sort_index()

    ####### Do transformations for Telangana solar data
    # Ensure 'local_time' is datetime format
    df_solar_telangana["local_time"] = pd.to_datetime(df_solar_telangana["local_time"], format="%d-%m-%Y %H:%M",
                                                      errors="coerce")
    # Create a full 15-minute timestamp range
    common_index = pd.date_range(
        start=df_solar_telangana["local_time"].min().replace(minute=0),  # Start at 00:00
        end=df_solar_telangana["local_time"].max().replace(hour=23, minute=45),  # End at 23:45
        freq="15min")
    # Reindex to match 15-minute intervals
    df_solar_telangana = df_solar_telangana.set_index("local_time").reindex(common_index)
    # Forward-fill to copy hourly values to missing 15-min slots
    df_solar_telangana["electricity"] = df_solar_telangana["electricity"].fillna(method="ffill").fillna(0)
    # Rename columns and set index
    df_solar_telangana = df_solar_telangana.rename(columns={"electricity": "Solar Production"})
    df_solar_telangana.index.name = "Timestamp"
    df_solar_telangana[
        'Solar Production']  # This is only Normalized Solar Production (1 MW Solar). See below for actual size PV production
    df_solar_telangana = df_solar_telangana.sort_index()

    ####### Do transformations for Rajasthan solar data
    # Ensure 'local_time' is datetime format
    df_solar_rajasthan["local_time"] = pd.to_datetime(df_solar_rajasthan["local_time"], format="%d-%m-%Y %H:%M",
                                                      errors="coerce")
    # Create a full 15-minute timestamp range
    common_index = pd.date_range(
        start=df_solar_rajasthan["local_time"].min().replace(minute=0),  # Start at 00:00
        end=df_solar_rajasthan["local_time"].max().replace(hour=23, minute=45),  # End at 23:45
        freq="15min")
    # Reindex to match 15-minute intervals
    df_solar_rajasthan = df_solar_rajasthan.set_index("local_time").reindex(common_index)
    # Forward-fill to copy hourly values to missing 15-min slots
    df_solar_rajasthan["electricity"] = df_solar_rajasthan["electricity"].fillna(method="ffill").fillna(0)
    # Rename columns and set index
    df_solar_rajasthan = df_solar_rajasthan.rename(columns={"electricity": "Solar Production"})
    df_solar_rajasthan.index.name = "Timestamp"
    df_solar_rajasthan[
        'Solar Production']  # This is only Normalized Solar Production (1 MW Solar). See below for actual size PV production
    df_solar_rajasthan = df_solar_rajasthan.sort_index()

    ###### MONTHLY CUF FOR ALL RE GENERATORS
    def calculate_monthly_cuf(df, source_type, capacity=1):
        """
        Calculate monthly Capacity Utilization Factor (CUF)

        Args:
            df: DataFrame containing generation data
            source_type: 'solar' or 'wind' to specify which production column to use
            capacity: Installed capacity in MW
        """
        if source_type == 'solar':
            production_column = 'Solar Production'
        else:  # wind
            production_column = 'Wind Production'

        monthly_energy = df[production_column].resample('M').sum() / 4
        hours_in_month = monthly_energy.index.days_in_month * 24
        monthly_cuf = (monthly_energy / (capacity * hours_in_month)) * 100

        return monthly_cuf

    cuf_df = pd.DataFrame({
        'Solar Gujarat CUF (%)': calculate_monthly_cuf(df_solar_gujarat, 'solar'),
        'Solar Rajasthan CUF (%)': calculate_monthly_cuf(df_solar_rajasthan, 'solar'),
        'Solar Telangana CUF (%)': calculate_monthly_cuf(df_solar_telangana, 'solar'),
        'Solar Goa/DRE CUF (%)': calculate_monthly_cuf(df_solar_goa, 'solar'),
        'Wind SRI CUF (%)': calculate_monthly_cuf(df_wind_long, 'wind', wind_size_actual_SRI),
        'Wind SECI CUF (%)': calculate_monthly_cuf(df_wind_long_SECI, 'wind', wind_size_actual_SECI)
    })

    ############## ADJUSTING CUF OF STATES BASED ON CEA ESTIMATES
    target_cufs = {
        'gujarat': {
            1: 15.0, 2: 18.5, 3: 20.0, 4: 22.0, 5: 21.0, 6: 17.0,
            7: 12.0, 8: 12.0, 9: 16.0, 10: 17.0, 11: 14.0, 12: 12.5
        },
        'rajasthan': {
            1: 14.0, 2: 17.5, 3: 20, 4: 21.5, 5: 22.0, 6: 20,
            7: 17.5, 8: 16.0, 9: 17.0, 10: 17.0, 11: 14.0, 12: 13.5
        },
        'maharashtra_wind': {
            1: 25.0, 2: 25.0, 3: 25, 4: 30.0, 5: 40.0, 6: 52,
            7: 61.0, 8: 52.0, 9: 32.0, 10: 24.0, 11: 28.0, 12: 25.0
        },
        'tamil_wind': {
            1: 30.0, 2: 25.0, 3: 20, 4: 28.0, 5: 48.0, 6: 62,
            7: 68.0, 8: 60.0, 9: 55.0, 10: 30.0, 11: 16.0, 12: 28.0
        },
        'karnataka_wind': {
            1: 30.0, 2: 32.0, 3: 28, 4: 22.0, 5: 50.0, 6: 62,
            7: 62.0, 8: 58.0, 9: 45.0, 10: 35.0, 11: 29.0, 12: 38.0
        },
        'telangana': {
            1: 16.0, 2: 18.0, 3: 18.0, 4: 20.0, 5: 20.0, 6: 17.5,
            7: 12.5, 8: 12.5, 9: 17.5, 10: 18.0, 11: 20.0, 12: 14.0
        }
    }

    def adjust_generation_profile(df_original, original_cuf, target_cuf_by_month):
        """
        Adjust generation profile to match target monthly CUFs
        """
        df_adjusted = df_original.copy()

        for month in range(1, 13):
            # Get the original and target CUF for this month
            original_month_cuf = original_cuf.iloc[month - 1]
            target_month_cuf = target_cuf_by_month[month]

            # Calculate adjustment ratio
            adjustment_ratio = target_month_cuf / original_month_cuf

            # Apply adjustment to that month's data
            month_mask = df_adjusted.index.month == month
            df_adjusted.loc[month_mask, 'Solar Production'] *= adjustment_ratio

        return df_adjusted

    def adjust_wind_cuf_profile(df_original, original_cuf, target_cuf_by_month):
        """
        Adjust generation profile to match target monthly CUFs
        """
        df_adjusted = df_original.copy()

        for month in range(1, 13):
            # Get the original and target CUF for this month
            original_month_cuf = original_cuf.iloc[month - 1]
            target_month_cuf = target_cuf_by_month[month]

            # Calculate adjustment ratio
            adjustment_ratio = target_month_cuf / original_month_cuf

            # Apply adjustment to that month's data
            month_mask = df_adjusted.index.month == month
            df_adjusted.loc[month_mask, 'Wind Production'] *= adjustment_ratio

        return df_adjusted

    # Adjust profiles for each state
    df_solar_gujarat_adjusted = adjust_generation_profile(
        df_solar_gujarat,
        cuf_df['Solar Gujarat CUF (%)'],
        target_cufs['gujarat']
    )

    df_solar_telangana_adjusted = adjust_generation_profile(
        df_solar_telangana,
        cuf_df['Solar Telangana CUF (%)'],
        target_cufs['telangana']
    )

    df_solar_rajasthan_adjusted = adjust_generation_profile(
        df_solar_rajasthan,
        cuf_df['Solar Rajasthan CUF (%)'],
        target_cufs['rajasthan']
    )

    df_wind_maharashtra_adjusted = adjust_wind_cuf_profile(
        df_wind_long,
        cuf_df['Wind SRI CUF (%)'],
        target_cufs['maharashtra_wind']
    )

    df_wind_tamil_adjusted = adjust_wind_cuf_profile(
        df_wind_long,
        cuf_df['Wind SRI CUF (%)'],
        target_cufs['tamil_wind']
    )

    df_wind_karnataka_adjusted = adjust_wind_cuf_profile(
        df_wind_long,
        cuf_df['Wind SRI CUF (%)'],
        target_cufs['karnataka_wind']
    )

    # Merge DataFrames
    df_all = pd.DataFrame(index=df_demand_year.index)
    df_all['TOTAL DEMAND'] = df_demand_year['TOTAL DEMAND']

    # df_all['Wind Production SRI'] = df_wind_long['Wind Production'].values * (1 - inter_state_losses)
    df_all['Wind Production Maharashtra'] = df_wind_maharashtra_adjusted['Wind Production'].values * (
            1 - intra_state_losses) * (wind_maharashtra_goa / wind_size_actual_SRI)
    df_all['Wind Production Karnataka'] = df_wind_karnataka_adjusted['Wind Production'].values * (
            1 - inter_state_losses) * (wind_karnataka / wind_size_actual_SRI)
    df_all['Wind Production Tamil Nadu'] = df_wind_tamil_adjusted['Wind Production'].values * (
            1 - inter_state_losses) * (wind_tamil / wind_size_actual_SRI)
    # df_all['Wind Production SECI'] = df_wind_long_SECI['Wind Production'].values * (1 - inter_state_losses)
    df_all["Solar Production Gujarat"] = df_solar_gujarat_adjusted['Solar Production'].values * PV_size_gujarat * (
            1 - inter_state_losses)
    df_all["Solar Production Telangana"] = df_solar_telangana_adjusted[
                                               'Solar Production'].values * PV_size_telangana * (
                                                   1 - inter_state_losses)
    df_all["Solar Production Rajasthan"] = df_solar_rajasthan_adjusted[
                                               'Solar Production'].values * PV_size_rajasthan * (
                                                   1 - inter_state_losses)
    df_all['Solar Production Goa'] = df_solar_goa['Solar Production'].values * PV_size_actual_goa * (
            1 - intra_state_losses)  # PV within GOA.
    df_all['DRE Production'] = df_solar_goa['Solar Production'].values * DRE_size * (
            1 - intra_state_losses)  # DRE considered all within GOA, all PV installments
    df_all['Biomass Production'] = [Biomass_size] * len(df_all)
    df_all['Nuclear Production'] = [Nuclear_size] * len(df_all)
    df_all['Gas Production'] = [Gas_size] * len(df_all)
    df_all['RTC Production'] = [RTC_size] * len(df_all)
    df_all['Total Solar Production'] = (df_all['Solar Production Gujarat'] + df_all['Solar Production Telangana']
                                        + df_all['Solar Production Rajasthan'] + df_all['DRE Production'] + df_all[
                                            'Solar Production Goa'])
    df_all['Total Wind Production'] = df_all['Wind Production Tamil Nadu'] + df_all['Wind Production Karnataka'] + \
                                      df_all['Wind Production Maharashtra']

    # Compute renewable generation and net demand
    df_all["renewable"] = (
            df_all['Total Wind Production'] + df_all['Total Solar Production'] + df_all["Biomass Production"] +
            df_all['Nuclear Production'] + df_all['RTC Production'])
    df_all["WITH SURPLUS"] = (df_all["TOTAL DEMAND"] - df_all["renewable"] - df_all[
        'Gas Production'])  # This is the surplus which can be used to charge the battery
    df_all["NET DEMAND"] = (df_all["TOTAL DEMAND"] - df_all["renewable"] - df_all['Gas Production']).clip(
        lower=0)  # This is only the net demand which needs to be met by the generators

    df_all = df_all.sort_index()

    logging.info("*** Successfully Created the DataFrame with all Data *** \n")

    weekly_stats, interesting_weeks_dict = weekly_stat_analysis(df_all)

    # Filter the DataFrame for the specific date or time range
    df_filtered = df_all.loc[start_date:end_date]

    input_file_path_dem = os.path.join(results_dir, 'Original_Demand_&_RE.xlsx')
    df_filtered.to_excel(input_file_path_dem, index=True)

    logging.info("*** Saved Demand Input and Original RE profiles to Excel *** \n")

    logging.info("*** Starting Non-Optimized Battery Scheduling for High RE *** \n")

    ####### BATTERY CHARGING/DISCHARGING PROFILES
    battery_profiles, remaining_surplus_history = battery_fixed_size_calculations(df_filtered, min_batt_soc,
                                                                                  batt_efficiency, battery_configs)

    battery_1_profile_df = battery_profiles['Battery 1']
    battery_2_profile_df = battery_profiles['Battery 2']
    battery_3_profile_df = battery_profiles['Battery 3']

    original_surplus = df_filtered["WITH SURPLUS"].values
    remaining_surplus_battery1 = remaining_surplus_history['Battery 1']
    remaining_surplus_battery2 = remaining_surplus_history['Battery 2']
    remaining_surplus_battery3 = remaining_surplus_history['Battery 3']

    df_remaining_surplus = pd.DataFrame({
        'Original Demand or Surplus': original_surplus,
        'After Battery1 Schedule': remaining_surplus_battery1,
        'After Battery2 Schedule': remaining_surplus_battery2,
        'After Battery3 Schedule': remaining_surplus_battery3
    }, index=battery_1_profile_df.index)

    output_path = os.path.join(results_dir, 'NonOptimized_Battery_Profiles.xlsx')

    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        battery_1_profile_df.to_excel(writer, sheet_name='Battery 1', index=True)
        battery_2_profile_df.to_excel(writer, sheet_name='Battery 2', index=True)
        battery_3_profile_df.to_excel(writer, sheet_name='Battery 3', index=True)
        df_remaining_surplus.to_excel(writer, sheet_name='Remaining Surplus or Demand', index=True)

    logging.info("*** Saved Non-Optimized Battery Profiles to Excel *** \n")

    ############# OPTIMAL COMBINATION OF SOLAR ONLY
    def calculate_deficit(df_demand, df_solar_guj, df_solar_raj, df_solar_goa, df_solar_tel,
                          size_guj, size_raj, size_goa, size_tel):
        """
        Calculate total deficit given solar sizes and demand
        """
        total_solar = (
                df_solar_guj['Solar Production'] * size_guj * (1 - inter_state_losses) +
                df_solar_raj['Solar Production'] * size_raj * (1 - inter_state_losses) +
                df_solar_goa['Solar Production'] * size_goa +  # Goa is local, no interstate losses
                df_solar_tel['Solar Production'] * size_tel * (1 - inter_state_losses)  # Added Telangana
        )

        # Adjust the time index of total_solar to the year 2030
        total_solar.index = total_solar.index.map(lambda x: x.replace(year=2030))

        # Calculate deficit (positive values only)
        deficit = abs(df_demand['TOTAL DEMAND'] - total_solar).clip(lower=0)

        return {
            'total_deficit': deficit.sum(),
            'max_deficit': deficit.max(),
            'mean_deficit': deficit.mean(),
            'total_solar_capacity': size_guj + size_raj + size_goa + size_tel  # Updated total capacity
        }

    def optimize_solar_sizes(df_demand, df_solar_guj, df_solar_raj, df_solar_goa, df_solar_tel,
                             min_size=1000, max_size=1500, step=100):
        """
        Grid search to find optimal solar sizes
        """
        best_result = None
        best_metrics = {'total_deficit': float('inf')}

        # Grid search through different combinations
        for size_guj in range(min_size, max_size + step, step):
            for size_raj in range(min_size, max_size + step, step):
                for size_goa in range(min_size, max_size + step, step):
                    for size_tel in range(min_size, max_size + step, step):  # Added Telangana loop

                        # Calculate metrics for this combination
                        metrics = calculate_deficit(
                            df_demand, df_solar_guj, df_solar_raj, df_solar_goa, df_solar_tel,
                            size_guj, size_raj, size_goa, size_tel
                        )

                        # Update best result if this combination is better
                        if metrics['total_deficit'] < best_metrics['total_deficit']:
                            best_metrics = metrics
                            best_result = {
                                'Gujarat_size': size_guj,
                                'Rajasthan_size': size_raj,
                                'Goa_size': size_goa,
                                'Telangana_size': size_tel  # Added Telangana size
                            }

        return best_result, best_metrics

    ############# OPTIMAL COMBINATION OF SOLAR & WIND COMBINED
    def calculate_deficit_with_wind(df_demand, df_solar_guj, df_solar_raj, df_solar_goa,
                                    df_wind_sri, df_wind_seci,
                                    size_guj, size_raj, size_goa, size_wind_sri, size_wind_seci):
        """
        Calculate total deficit given solar and wind sizes and demand
        """
        total_generation = (
            # Solar generation
                df_solar_guj['Solar Production'] * size_guj * (1 - inter_state_losses) +
                df_solar_raj['Solar Production'] * size_raj * (1 - inter_state_losses) +
                df_solar_goa['Solar Production'] * size_goa +  # Goa is local, no interstate losses
                # Wind generation
                df_wind_sri['Wind Production'] * size_wind_sri * (1 - inter_state_losses) +
                df_wind_seci['Wind Production'] * size_wind_seci * (1 - inter_state_losses)
        )
        # Calculate deficit (positive values only)
        deficit = (df_demand['TOTAL DEMAND'] - total_generation).clip(lower=0)

        return {
            'total_deficit': deficit.sum(),
            'max_deficit': deficit.max(),
            'mean_deficit': deficit.mean(),
            'total_solar_capacity': size_guj + size_raj + size_goa,
            'total_wind_capacity': size_wind_sri + size_wind_seci,
            'total_renewable_capacity': size_guj + size_raj + size_goa + size_wind_sri + size_wind_seci
        }

    ########## SOLAR & WIND
    def optimize_renewable_sizes(df_demand, df_solar_guj, df_solar_raj, df_solar_goa,
                                 df_wind_sri, df_wind_seci,
                                 min_solar_size=500, max_solar_size=2000, solar_step=500,
                                 min_wind_size=500, max_wind_size=2000, wind_step=500):
        """
        Grid search to find optimal solar and wind sizes
        """
        best_result = None
        best_metrics = {'total_deficit': float('inf')}

        # Grid search through different combinations
        for size_guj in range(min_solar_size, max_solar_size + solar_step, solar_step):
            for size_raj in range(min_solar_size, max_solar_size + solar_step, solar_step):
                for size_goa in range(min_solar_size, max_solar_size + solar_step, solar_step):
                    for size_wind_sri in range(min_wind_size, max_wind_size + wind_step, wind_step):
                        for size_wind_seci in range(min_wind_size, max_wind_size + wind_step, wind_step):

                            # Calculate metrics for this combination
                            metrics = calculate_deficit_with_wind(
                                df_demand, df_solar_guj, df_solar_raj, df_solar_goa,
                                df_wind_sri, df_wind_seci,
                                size_guj, size_raj, size_goa, size_wind_sri, size_wind_seci
                            )

                            # Update best result if this combination is better
                            if metrics['total_deficit'] < best_metrics['total_deficit']:
                                best_metrics = metrics
                                best_result = {
                                    'Gujarat_solar_size': size_guj,
                                    'Rajasthan_solar_size': size_raj,
                                    'Goa_solar_size': size_goa,
                                    'Wind_SRI_size': size_wind_sri,
                                    'Wind_SECI_size': size_wind_seci
                                }

        return best_result, best_metrics

    # =============================================================================
    # Assume that the following variables are already defined from your preprocessing:
    #
    # - df_all: A DataFrame with a 15-minute index and a column 'net_demand'
    # - time_list: A list of time indices (integers) matching the rows of df_all
    # - time_mapping: A dictionary mapping these integer indices to actual timestamps
    # - net_demand_dict: A dictionary mapping each time index to the net demand (MW)
    # =============================================================================

    ###### FINDING CORRELATIONS
    correlation_allrenewable = df_filtered['renewable'].corr(df_filtered['TOTAL DEMAND'])

    correlation_allsolar = df_all['TOTAL DEMAND'].corr(df_all['Total Solar Production'])

    correlation_allwind = df_all['TOTAL DEMAND'].corr(df_all['Total Wind Production'])

    time_list = list(range(len(df_filtered)))
    time_mapping = dict(enumerate(df_filtered.index))
    net_demand_dict = dict(enumerate(df_filtered['WITH SURPLUS']))

    if run_thermal_sizing_optimization:
        logging.info("*** Beginning Thermal Scheduling Optimization *** \n")
        # =============================================================================
        # Generator Data for Thermal Plants
        # =============================================================================

        # Read the Excel file
        df_generators = pd.read_excel(file_path_generators)

        # Simplified version of the gen_data dictionary creation
        gen_data = df_generators.set_index('PPA Details')[['MW', 'Variable Cost']].to_dict('index')
        gen_data = {k: {'max_capacity': v['MW'], 'var_cost': v['Variable Cost']} for k, v in gen_data.items()}

        ramp_rate = 0.15  # 1% is ramp rate per minute so for 15 minutes 0.15
        gen_list = list(gen_data.keys())
        min_gen_factor = 0.5  # Minimum generation limit as a fraction of capacity

        # Penalty cost for unmet demand (load shedding), set high to discourage its use.
        # Adjust as needed

        # =============================================================================
        # Pyomo Optimization Model with Slack Variables for Demand Balance
        # =============================================================================
        model = ConcreteModel()

        # Sets: Time periods and generators
        model.T = RangeSet(0, len(time_list) - 1)
        model.I = Set(initialize=gen_list)

        # Parameters:
        # Net demand at each time period (MW)
        model.demand = Param(model.T, initialize=net_demand_dict)

        # Generator maximum capacity (MW)
        def cap_init(model, i):
            return float(gen_data[i]['max_capacity'])

        model.cap = Param(model.I, initialize=cap_init)

        # Variable cost for each generator (per MWh)
        def cost_init(model, i):
            return float(gen_data[i]['var_cost'])

        model.var_cost = Param(model.I, initialize=cost_init)

        # Ramp rate factor (fraction of capacity per period)
        model.ramp_rate = Param(initialize=ramp_rate)

        # =============================================================================
        # Decision Variables:
        # Generation output from generator i at time t (MW)
        model.x = Var(model.I, model.T, domain=NonNegativeReals)
        # Slack variable for unmet demand at time t (MW)
        model.u = Var(model.T, domain=NonNegativeReals)

        # =============================================================================
        # Objective: Minimize total cost (generation cost + penalty for unserved demand)
        # =============================================================================
        def objective_rule(model):
            generation_cost = sum(model.var_cost[i] * model.x[i, t] for i in model.I for t in model.T)
            slack_cost = sum(penalty_thermal_unmet_demand * model.u[t] for t in model.T)
            return generation_cost + slack_cost

        model.obj = Objective(rule=objective_rule)

        # =============================================================================
        # Constraints
        # =============================================================================

        # 1. Demand Balance: Thermal generation plus slack must equal net demand
        def demand_balance_rule(model, t):
            # The idea is:
            #   (Thermal generation + Renewable generation) + battery discharge
            #     - battery charge + slack = net demand
            return (sum(model.x[i, t] for i in model.I) +
                    model.u[t]) >= model.demand[t]

        model.demand_balance = Constraint(model.T, rule=demand_balance_rule)

        # 2. Generator capacity limits:
        def capacity_limit_rule(model, i, t):
            return model.x[i, t] <= model.cap[i]

        model.capacity_limit = Constraint(model.I, model.T, rule=capacity_limit_rule)

        # 3. Ramp-up constraints:
        def ramp_up_rule(model, i, t):
            if t == 0:
                return Constraint.Skip  # No ramp-down constraint for the first time period
            return model.x[i, t] - model.x[i, t - 1] <= model.ramp_rate * model.cap[i]

        model.ramp_up = Constraint(model.I, model.T, rule=ramp_up_rule)

        # 4. Ramp-down constraints:
        def ramp_down_rule(model, i, t):
            if t == 0:
                return Constraint.Skip  # No ramp-down constraint for the first time period
            return model.x[i, t - 1] - model.x[i, t] <= model.ramp_rate * model.cap[i]

        model.ramp_down = Constraint(model.I, model.T, rule=ramp_down_rule)

        # Minimum generation limit (MW)
        def min_gen_init(model, i):
            return min_gen_factor * gen_data[i]['max_capacity']

        model.min_gen = Param(model.I, initialize=min_gen_init)

        def min_gen_limit_rule(model, i, t):
            return model.x[i, t] >= model.min_gen[i]

        model.min_gen_limit = Constraint(model.I, model.T, rule=min_gen_limit_rule)

        logging.info("About to Create Solver \n")

        solver = SolverFactory('highs')
        # solver = SolverFactory('appsi_highs')  # faster but not easily compatible with pyinstaller.

        # solver = SolverFactory('cbc', executable=r"C:\Users\i60608\OneDrive\Cbc-2.10.5\bin\cbc.exe")

        logging.info("Solver Created Successfully! \n")
        logging.info(f"Solver Available: {solver.available()}")

        # solver.options['max_iter'] = 150
        # solver.options['TimeLimit'] = 300  # Set time limit for Gurobi

        try:
            logging.info("Calling solver.solve()...")

            results = solver.solve(model, tee=True)

            logging.info(f"*** Solver Status: {results.solver.status} *** \n")
        except Exception as e:
            logging.info(f"Error during solve: {type(e).__name__}: {e}")
            import traceback
            traceback.print_exc()

        # =============================================================================
        # Postprocessing: Extract the results
        # =============================================================================
        schedule = pd.DataFrame(index=[time_mapping[t] for t in model.T],
                                columns=gen_list + ['Unserved Demand', 'With Surplus'])

        for t in model.T:
            for i in model.I:
                schedule.loc[time_mapping[t], i] = value(model.x[i, t])
            schedule.loc[time_mapping[t], 'Unserved Demand'] = value(model.u[t])
            schedule.loc[time_mapping[t], 'With Surplus'] = value(model.demand[t])

        # Define the file path for the Excel file
        output_file_path_thermal = os.path.join(results_dir, 'thermal_generation.xlsx')

        # Save the schedule DataFrame to an Excel file
        schedule.to_excel(output_file_path_thermal, index=True)

        logging.info("*** Saved Optimized Thermal Schedules to Excel *** \n")

        ##### Uncomment below if plots are needed
        # plot_demand(df_filtered)

        ############## OPTIMAL SIZING OF PV, WIND & BESS FOR UNMET DEMAND
        # After solving the first optimization model, extract the unmet demand
        unmet_demand_values = [value(model.u[t]) for t in model.T]

        # Create the timestamp index
        time_indices = [time_mapping[t] for t in model.T]

        # Create the Series directly from the values and indices
        # OLD UNMET DEMAND
        # unmet_demand_series = pd.Series(unmet_demand_values, index=time_indices)

        # Assume you have normalized solar and wind profiles (1 MW capacity)
        # If not already available, you can normalize your existing profiles:
        # solar_profile = df_filtered['Solar Production Goa'] / PV_size_actual_goa
        # wind_profile = df_filtered['Wind Production Maharashtra'] / (wind_maharashtra_goa)

        ################# NEW ADDED JUGAAD
        df_all['hour'] = df_all.index.hour
        df_all['minute'] = df_all.index.minute
        df_all['month'] = df_all.index.month

        # Create a time slot identifier (0-95 for the 96 time slots in a day)
        df_all['time_slot'] = df_all['hour'] * 4 + df_all['minute'] // 15

        # Get a list of all data columns (exclude the time components we just created)
        data_columns = df_all.columns.difference(['hour', 'minute', 'month', 'time_slot'])

        # Group by month and time slot to get average for each time slot in each month
        monthly_time_slot_avg = df_all.groupby(['month', 'time_slot'])[data_columns].mean()

        # Read the Excel file with unserved demand data
        if shortage_case == 'case1':
            df_unserved = pd.read_excel(file_path_shortage_case1, parse_dates=['Timestamp'])
        elif shortage_case == 'case2':
            df_unserved = pd.read_excel(file_path_shortage_case2, parse_dates=['Timestamp'])

        # Set the Timestamp column as the index
        df_unserved.set_index('Timestamp', inplace=True)

        # Create time identifier components
        df_unserved['hour'] = df_unserved.index.hour
        df_unserved['minute'] = df_unserved.index.minute
        df_unserved['month'] = df_unserved.index.month
        df_unserved['time_slot'] = df_unserved['hour'] * 4 + df_unserved['minute'] // 15

        monthly_unserved_avg = df_unserved.groupby(['month', 'time_slot'])[['Unserved Demand']].mean()

        solar_profile_goa = monthly_time_slot_avg['Solar Production Goa'] / PV_size_actual_goa
        solar_profile_guj = monthly_time_slot_avg['Solar Production Gujarat'] / PV_size_gujarat
        solar_profile_raj = monthly_time_slot_avg['Solar Production Rajasthan'] / PV_size_rajasthan
        solar_profile_tel = monthly_time_slot_avg['Solar Production Telangana'] / PV_size_telangana
        wind_profile_maha = monthly_time_slot_avg['Wind Production Maharashtra'] / (wind_maharashtra_goa)
        wind_profile_tamil = monthly_time_slot_avg['Wind Production Tamil Nadu'] / (wind_tamil)
        wind_profile_karnataka = monthly_time_slot_avg['Wind Production Karnataka'] / (wind_karnataka)

        flattened_unserved = monthly_unserved_avg.reset_index()

        # Sort by month and time_slot to ensure consistent ordering
        flattened_unserved = flattened_unserved.sort_values(['month', 'time_slot'])

        if gdam_price_select_year == 2023:
            gdam_price_series = pd.Series(df_gdam_price['Average of MCP 2023'].values)
        elif gdam_price_select_year == 2024:
            gdam_price_series = pd.Series(df_gdam_price['Average of MCP 2024'].values)

        # Extract just the values into a series
        unmet_demand_series = pd.Series(flattened_unserved['Unserved Demand'].values)

        ################# END OF NEW ADDED JUGAAD

        logging.info("*** Beginning RE & BESS Sizing Optimization *** \n")
        model_renewable = ConcreteModel()
        # Parameters
        time_periods = list(range(len(unmet_demand_series)))
        model_renewable.T = Set(initialize=time_periods)
        model_renewable.demand = Param(model_renewable.T,
                                       initialize={t: unmet_demand_series.iloc[t] for t in time_periods})
        model_renewable.gdam_price = Param(model_renewable.T,
                                           initialize={t: gdam_price_series.iloc[t] for t in time_periods})
        model_renewable.max_gdam = Param(initialize=max_gdam_purchase)
        model_renewable.min_total_solar = Param(initialize=params['min_total_solar'], domain=NonNegativeReals)
        model_renewable.max_total_solar = Param(initialize=params['max_total_solar'], domain=NonNegativeReals)
        model_renewable.min_total_wind = Param(initialize=params['min_total_wind'], domain=NonNegativeReals)
        model_renewable.max_total_wind = Param(initialize=params['max_total_wind'], domain=NonNegativeReals)

        # Decision variables
        model_renewable.solar_size_goa = Var(domain=NonNegativeReals)  # Total solar capacity (MW)
        model_renewable.solar_size_guj = Var(domain=NonNegativeReals)  # Total solar capacity (MW)
        model_renewable.solar_size_raj = Var(domain=NonNegativeReals)  # Total solar capacity (MW)
        model_renewable.solar_size_tel = Var(domain=NonNegativeReals)  # Total solar capacity (MW)
        model_renewable.wind_size_maha = Var(domain=NonNegativeReals)  # Total wind capacity (MW)
        model_renewable.wind_size_tamil = Var(domain=NonNegativeReals)  # Total wind capacity (MW)
        model_renewable.wind_size_karnataka = Var(domain=NonNegativeReals)  # Total wind capacity (MW)

        model_renewable.gdam_purchase = Var(model_renewable.T,
                                            domain=NonNegativeReals)  # GDAM power share (MW) for each time period

        model_renewable.battery_capacity = Var(domain=NonNegativeReals)  # Battery energy capacity (MWh)
        model_renewable.max_charge_rate = Var(domain=NonNegativeReals)  # Max charge/discharge rate (MW)

        # Battery operation variables
        model_renewable.charge = Var(model_renewable.T, domain=NonNegativeReals)
        model_renewable.discharge = Var(model_renewable.T, domain=NonNegativeReals)
        model_renewable.soc = Var(model_renewable.T, domain=NonNegativeReals)
        model_renewable.deficit = Var(model_renewable.T, domain=NonNegativeReals)  # Any remaining deficit

        # Add a binary variable to indicate charging (1) or discharging (0)
        model_renewable.is_charging = Var(model_renewable.T, domain=Binary)

        # Solar and wind production at each time period
        solar_dict_goa = {t: solar_profile_goa.iloc[t] for t in time_periods}  # Normalized production (0-1)
        solar_dict_gujarat = {t: solar_profile_guj.iloc[t] for t in time_periods}
        solar_dict_rajasthan = {t: solar_profile_raj.iloc[t] for t in time_periods}
        solar_dict_tel = {t: solar_profile_tel.iloc[t] for t in time_periods}  # Normalized production (0-1)
        wind_dict_maharashtra = {t: wind_profile_maha.iloc[t] for t in time_periods}  # Normalized production (0-1)
        wind_dict_tamil = {t: wind_profile_tamil.iloc[t] for t in time_periods}  # Normalized production (0-1)
        wind_dict_karnataka = {t: wind_profile_karnataka.iloc[t] for t in time_periods}  # Normalized production (0-1)

        pen_charge_discharge = 10

        # Objective: Minimize the cost of new capacity and any remaining deficit
        def objective_rule(model):
            # Energy purchasing costs for solar and wind in each time period
            solar_energy_cost = sum(
                (solar_cost_goa * model.solar_size_goa * solar_dict_goa[t]) +
                (solar_cost_guj * model.solar_size_guj * solar_dict_gujarat[t]) +
                (solar_cost_raj * model.solar_size_raj * solar_dict_rajasthan[t]) +
                (solar_cost_tel * model.solar_size_tel * solar_dict_tel[t])
                for t in model.T
            )

            wind_energy_cost = sum(
                (wind_cost_maha * model.wind_size_maha * wind_dict_maharashtra[t]) +
                (wind_cost_tamil * model.wind_size_tamil * wind_dict_tamil[t]) +
                (wind_cost_karnataka * model.wind_size_karnataka * wind_dict_karnataka[t])
                for t in model.T
            )

            #
            battery_cost = (battery_cost_MWh * model.battery_capacity)

            # GDAM purchase costs
            gdam_cost = sum(model.gdam_price[t] * model.gdam_purchase[t] for t in model.T)

            # Deficit penalty and battery operation control
            deficit_penalty = sum(penalty_sizing_unmet_demand * model.deficit[t] for t in model.T)
            charging_discharging_control = pen_charge_discharge * sum(
                model.charge[t] + model.discharge[t] for t in model.T)

            return solar_energy_cost + wind_energy_cost + battery_cost + gdam_cost + deficit_penalty + charging_discharging_control

        model_renewable.objective = Objective(rule=objective_rule)

        # Constraints

        # Energy balance constraint
        def energy_balance_rule(model, t):
            solar_gen_goa = solar_dict_goa[t] * model.solar_size_goa
            solar_gen_guj = solar_dict_gujarat[t] * model.solar_size_guj
            solar_gen_raj = solar_dict_rajasthan[t] * model.solar_size_raj
            solar_gen_tel = solar_dict_tel[t] * model.solar_size_tel
            wind_gen_maha = wind_dict_maharashtra[t] * model.wind_size_maha
            wind_gen_tamil = wind_dict_tamil[t] * model.wind_size_tamil
            wind_gen_karnataka = wind_dict_karnataka[t] * model.wind_size_karnataka

            if allow_oversized_RE == True:
                cons_match = (
                            solar_gen_goa + solar_gen_guj + solar_gen_raj + solar_gen_tel + wind_gen_maha + wind_gen_tamil + wind_gen_karnataka +
                            model.discharge[t] - model.charge[t] + model.gdam_purchase[t] +
                            model.deficit[t] >= (model.demand[t]))
            else:
                cons_match = (
                            solar_gen_goa + solar_gen_guj + solar_gen_raj + solar_gen_tel + wind_gen_maha + wind_gen_tamil + wind_gen_karnataka +
                            model.discharge[t] - model.charge[t] + model.gdam_purchase[t] +
                            model.deficit[t] == model.demand[t])
            return cons_match

        model_renewable.energy_balance = Constraint(model_renewable.T, rule=energy_balance_rule)

        # Battery state of charge dynamics
        def soc_rule(model, t):
            if t == 0:
                return model.soc[t] == 0.5 * model.battery_capacity + (
                        model.charge[t] - model.discharge[t]) * (15 / 60)  # 15-min intervals
            else:
                return model.soc[t] == model.soc[t - 1] + (
                        model.charge[t] - model.discharge[t]) * (15 / 60)

        model_renewable.soc_constraint = Constraint(model_renewable.T, rule=soc_rule)

        # Battery charging/discharging rate limits
        def charge_rate_limit_rule(model, t):
            return model.charge[t] <= 0.1 * model.battery_capacity

        model_renewable.charge_rate_limit = Constraint(model_renewable.T, rule=charge_rate_limit_rule)

        def discharge_rate_limit_rule(model, t):
            return model.discharge[t] <= 0.1 * model.battery_capacity

        model_renewable.discharge_rate_limit = Constraint(model_renewable.T, rule=discharge_rate_limit_rule)

        # Add a constraint to make final SOC equal to initial SOC
        def final_soc_rule(model):
            t_final = model.T.last()
            initial_soc = 0.5 * model.battery_capacity

            return model.soc[t_final] == initial_soc

        model_renewable.final_soc_constraint = Constraint(rule=final_soc_rule)

        # Add constraint for daily SOC balance (every 96 time slots)
        def daily_soc_balance_rule(model, t):
            # Only apply at the end of each day (96 time slots)
            if (t + 1) % 96 != 0:
                return Constraint.Skip

            # Find the beginning of this day
            day_start = t - 95

            # SOC at end of day should equal SOC at beginning of that same day
            return model.soc[t] == model.soc[day_start]

        model_renewable.daily_soc_balance = Constraint(model_renewable.T, rule=daily_soc_balance_rule)

        # Battery capacity constraints
        def soc_max_rule(model, t):
            return model.soc[t] <= model.battery_capacity

        model_renewable.soc_max = Constraint(model_renewable.T, rule=soc_max_rule)

        def cap_max_rule(model):
            return model.battery_capacity <= max_size_batt_mwh

        model_renewable.cap_max = Constraint(rule=cap_max_rule)

        def soc_min_rule(model, t):
            return model.soc[t] >= 0.1 * model.battery_capacity  # 10% minimum SOC

        model_renewable.soc_min = Constraint(model_renewable.T, rule=soc_min_rule)

        # Battery charge/discharge rate constraints
        def charge_rate_rule(model, t):
            return model.charge[t] <= model.max_charge_rate

        model_renewable.charge_rate = Constraint(model_renewable.T, rule=charge_rate_rule)

        def discharge_rate_rule(model, t):
            return model.discharge[t] <= model.max_charge_rate

        model_renewable.discharge_rate = Constraint(model_renewable.T, rule=discharge_rate_rule)

        # C-rate constraint (relate power and energy capacity)
        def c_rate_rule(model):
            return model.max_charge_rate <= 0.5 * model.battery_capacity  # Max C-rate of 0.5C

        model_renewable.c_rate = Constraint(rule=c_rate_rule)

        def max_rate_ch_rule(model, t):
            return model.charge[t] <= max_charge_discharge_power_bess  #

        model_renewable.max_rate_ch = Constraint(model_renewable.T, rule=max_rate_ch_rule)

        def max_rate_dish_rule(model, t):
            return model.discharge[t] <= max_charge_discharge_power_bess  #

        model_renewable.max_rate_dish = Constraint(model_renewable.T, rule=max_rate_dish_rule)

        def total_solar_min_rule(model):
            total_solar = model.solar_size_goa + model.solar_size_guj + model.solar_size_raj + model.solar_size_tel
            return total_solar >= model.min_total_solar

        model_renewable.total_solar_min_constraint = Constraint(rule=total_solar_min_rule)

        def total_solar_max_rule(model):
            total_solar = model.solar_size_goa + model.solar_size_guj + model.solar_size_raj + model.solar_size_tel
            return total_solar <= model.max_total_solar

        model_renewable.total_solar_max_constraint = Constraint(rule=total_solar_max_rule)

        # Constraint rule for total wind capacity
        def total_wind_min_rule(model):
            total_wind = model.wind_size_maha + model.wind_size_tamil + model.wind_size_karnataka
            return total_wind >= model.min_total_wind

        model_renewable.total_wind_min_constraint = Constraint(rule=total_wind_min_rule)

        def total_wind_max_rule(model):
            total_wind = model.wind_size_maha + model.wind_size_tamil + model.wind_size_karnataka
            return total_wind <= model.max_total_wind

        model_renewable.total_wind_max_constraint = Constraint(rule=total_wind_max_rule)

        def max_gdam_rule(model, t):
            return model.gdam_purchase[t] <= model.max_gdam

        model_renewable.max_gdam_constraint = Constraint(model_renewable.T, rule=max_gdam_rule)

        def goa_solar_max_rule(model):
            return model.solar_size_goa <= max_solar_goa

        model_renewable.goa_solar_max_constraint = Constraint(rule=goa_solar_max_rule)

        def goa_solar_min_rule(model):
            return model.solar_size_goa >= min_solar_goa

        model_renewable.goa_solar_min_constraint = Constraint(rule=goa_solar_min_rule)

        def guj_solar_min_rule(model):
            return model.solar_size_guj >= min_solar_guj

        model_renewable.guj_solar_min_constraint = Constraint(rule=guj_solar_min_rule)

        def raj_solar_min_rule(model):
            return model.solar_size_raj >= min_solar_raj

        model_renewable.raj_solar_min_constraint = Constraint(rule=raj_solar_min_rule)

        def tel_solar_min_rule(model):
            return model.solar_size_tel >= min_solar_tel

        model_renewable.tel_solar_min_constraint = Constraint(rule=tel_solar_min_rule)

        def maha_wind_min_rule(model):
            return model.wind_size_maha >= min_wind_maha

        model_renewable.maha_wind_min_constraint = Constraint(rule=maha_wind_min_rule)

        def tamil_wind_min_rule(model):
            return model.wind_size_tamil >= min_wind_tamil

        model_renewable.tamil_wind_min_constraint = Constraint(rule=tamil_wind_min_rule)

        def karnataka_wind_min_rule(model):
            return model.wind_size_karnataka >= min_wind_karnataka

        model_renewable.karnataka_wind_min_constraint = Constraint(rule=karnataka_wind_min_rule)

        solver = SolverFactory('highs')
        # solver = SolverFactory('appsi_highs')  # faster but not easily compatible with pyinstaller.

        # solver = SolverFactory('cbc', executable=r"C:\Users\i60608\OneDrive\Cbc-2.10.5\bin\cbc.exe")

        results = solver.solve(model_renewable, tee=True)

        result_sizing = {
            'solar_size_goa': value(model_renewable.solar_size_goa),
            'solar_size_guj': value(model_renewable.solar_size_guj),
            'solar_size_raj': value(model_renewable.solar_size_raj),
            'solar_size_tel': value(model_renewable.solar_size_tel),
            'wind_size_maha': value(model_renewable.wind_size_maha),
            'wind_size_tamil': value(model_renewable.wind_size_tamil),
            'wind_size_karnataka': value(model_renewable.wind_size_karnataka),
            'battery_capacity': value(model_renewable.battery_capacity),
            'max_charge_rate': value(model_renewable.max_charge_rate),
            'total_deficit': sum(value(model_renewable.deficit[t]) for t in model_renewable.T)
        }

        gdam_purchase_series = pd.Series([value(model_renewable.gdam_purchase[t]) for t in model_renewable.T],
                                         index=time_periods)

        # Extract additional time series data from the model solution
        battery_charge = pd.Series([value(model_renewable.charge[t]) for t in model_renewable.T],
                                   index=unmet_demand_series.index)
        battery_discharge = pd.Series([value(model_renewable.discharge[t]) for t in model_renewable.T],
                                      index=unmet_demand_series.index)
        battery_soc = pd.Series([value(model_renewable.soc[t]) for t in model_renewable.T],
                                index=unmet_demand_series.index)
        remaining_deficit = pd.Series([value(model_renewable.deficit[t]) for t in model_renewable.T],
                                      index=unmet_demand_series.index)

        # Calculate net battery flow (positive = charging, negative = discharging)
        net_battery_flow = battery_charge - battery_discharge

        # Create corrected charge and discharge series
        battery_charge_corrected = pd.Series(0, index=unmet_demand_series.index)
        battery_discharge_corrected = pd.Series(0, index=unmet_demand_series.index)

        # Apply the correction logic
        battery_charge_corrected[net_battery_flow > 0] = net_battery_flow[net_battery_flow > 0]
        battery_discharge_corrected[net_battery_flow < 0] = -net_battery_flow[net_battery_flow < 0]

        # Calculate the production from each source
        solar_production_goa = solar_profile_goa * result_sizing['solar_size_goa']
        solar_production_guj = solar_profile_guj * result_sizing['solar_size_guj']
        solar_production_raj = solar_profile_raj * result_sizing['solar_size_raj']
        solar_production_tel = solar_profile_tel * result_sizing['solar_size_tel']

        wind_production_maha = wind_profile_maha * result_sizing['wind_size_maha']
        wind_production_tamil = wind_profile_tamil * result_sizing['wind_size_tamil']
        wind_production_karnataka = wind_profile_karnataka * result_sizing['wind_size_karnataka']

        # Calculate total solar and wind production (if needed)
        total_solar_production = solar_production_goa + solar_production_guj + solar_production_raj + solar_production_tel
        total_wind_production = wind_production_maha + wind_production_tamil + wind_production_karnataka

        # First convert all your data to Series with the same index type
        # Create a common index (regular, not MultiIndex)
        common_index = pd.RangeIndex(len(unmet_demand_series))

        gdam_purchase_values = [value(model_renewable.gdam_purchase[t]) for t in model_renewable.T]
        gdam_purchase_series = pd.Series(gdam_purchase_values, index=common_index)

        # Convert each Series to use this common index
        solar_production_goa_series = pd.Series(solar_production_goa.values, index=common_index)
        solar_production_guj_series = pd.Series(solar_production_guj.values, index=common_index)
        solar_production_raj_series = pd.Series(solar_production_raj.values, index=common_index)
        solar_production_tel_series = pd.Series(solar_production_tel.values, index=common_index)

        wind_production_maha_series = pd.Series(wind_production_maha.values, index=common_index)
        wind_production_tamil_series = pd.Series(wind_production_tamil.values, index=common_index)
        wind_production_karnataka_series = pd.Series(wind_production_karnataka.values, index=common_index)

        battery_discharge_series = pd.Series(battery_discharge_corrected.values, index=common_index)
        battery_charge_series = pd.Series(-battery_charge_corrected.values, index=common_index)
        remaining_deficit_series = pd.Series(remaining_deficit.values, index=common_index)
        unmet_demand_values_series = pd.Series(unmet_demand_series.values, index=common_index)
        battery_soc_series = pd.Series(battery_soc.values, index=common_index)

        # Now create the DataFrame with all Series having the same index type
        save_data = pd.DataFrame({
            'Solar Production Goa': solar_production_goa_series,
            'Solar Production Gujarat': solar_production_guj_series,
            'Solar Production Rajasthan': solar_production_raj_series,
            'Solar Production Telangana': solar_production_tel_series,
            'Wind Production Maharashtra': wind_production_maha_series,
            'Wind Production Tamil Nadu': wind_production_tamil_series,
            'Wind Production Karnataka': wind_production_karnataka_series,
            'GDAM Purchase': gdam_purchase_series,
            'Battery Discharge': battery_discharge_series,
            'Battery Charge': battery_charge_series,
            'Remaining Deficit': remaining_deficit_series,
            'Original Unmet Demand': unmet_demand_values_series,
            'Battery SOC': battery_soc_series
        })

        # Define the file path for the Excel file
        output_file_path_sizing = os.path.join(results_dir, 'Optimal_Sizing_RE_BESS.xlsx')

        # Save the schedule DataFrame to an Excel file
        save_data.to_excel(output_file_path_sizing, index=True)

        # Create an ExcelWriter object to save multiple sheets to the same file
        with pd.ExcelWriter(output_file_path_sizing, engine='openpyxl') as writer:
            # Save the time series data to one sheet
            save_data.to_excel(writer, sheet_name='Time Series Data', index=True)

            # Convert sizing results dictionary to DataFrame and save to another sheet
            result_sizing_df = pd.DataFrame.from_dict(result_sizing, orient='index', columns=['Value'])
            result_sizing_df.index.name = 'Parameter'
            result_sizing_df.to_excel(writer, sheet_name='Sizing Results')

        logging.info("*** Saved Optimized RE & BESS Size Output to Excel *** \n")

    else:
        logging.info("*** Skipping Thermal & RE-BESS Sizing Optimization *** \n")

    logging.info("*** END OF CODE *** \n")


if __name__ == "__main__":
    # This logic checks if a filename was passed as an argument.
    # If so, it uses it. Otherwise, it defaults to 'parameters.ini'.
    if len(sys.argv) > 1:
        config_file = sys.argv[1]
    else:
        config_file = 'parameters.ini'

    # Run the optimization with the correct configuration file.
    run_optimization(config_file)