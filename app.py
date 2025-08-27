import streamlit as st
import pandas as pd
import configparser
import os
import sys
import subprocess
from datetime import datetime

# --- 1. Page Configuration and Styling ---
st.set_page_config(
    page_title="Power System Optimizer",
    page_icon="⚡",
    layout="wide"
)


def load_css(file_name):
    try:
        with open(file_name) as f:
            st.markdown(f"<style>{f.read()}</style>", unsafe_allow_html=True)
    except FileNotFoundError:
        st.warning(f"CSS file '{file_name}' not found. Using default styles.")


load_css("style.css")


# --- 2. Core Application Logic and Helper Functions ---
def check_password():
    if "password_correct" not in st.session_state:
        st.session_state["password_correct"] = False

    if not st.session_state["password_correct"]:
        st.title("🔐 Login Required")
        st.markdown("Please enter the password to access the optimizer.")
        password_input = st.text_input("Password", type="password")

        if st.button("Login"):
            if password_input == st.secrets.get("APP_PASSWORD", "your_secret_password123"):
                st.session_state["password_correct"] = True
                st.rerun()
            else:
                st.error("The password you entered is incorrect.")
        return False
    return True

# In file: app.py

def display_results():
    """
    Scans the Results folder. Displays metrics if the main sizing file exists,
    and shows download links for ALL available result files.
    """
    st.header("Results & Analysis 📊")
    results_dir = "Results"
    sizing_file_path = os.path.join(results_dir, 'Optimal_Sizing_RE_BESS.xlsx')

    # --- 1. Display Key Metrics (only if the specific file exists) ---
    st.subheader("Key Sizing Results")
    if os.path.exists(sizing_file_path):
        try:
            df_results = pd.read_excel(sizing_file_path, sheet_name='Sizing Results', index_col=0)
            total_solar = df_results.loc['solar_size_goa':'solar_size_tel'].sum().iloc[0]
            total_wind = df_results.loc['wind_size_maha':'wind_size_karnataka'].sum().iloc[0]
            battery_mwh = df_results.loc['battery_capacity', 'Value']
            total_deficit = df_results.loc['total_deficit', 'Value']

            col1, col2, col3, col4 = st.columns(4)
            col1.metric("Total Solar Capacity (MW)", f"{total_solar:,.1f}")
            col2.metric("Total Wind Capacity (MW)", f"{total_wind:,.1f}")
            col3.metric("Battery Capacity (MWh)", f"{battery_mwh:,.1f}")
            col4.metric("Total Deficit (MWh)", f"{total_deficit:,.1f}", delta_color="inverse")
        except Exception as e:
            st.error(f"Could not read or parse the sizing results file: {e}")
    else:
        st.warning("Key metrics are unavailable because the sizing optimization was infeasible or did not complete.")

    st.divider()

    # --- 2. Display Download Links (for ANY Excel file found) ---
    st.subheader("Download Output Files")
    if not os.path.exists(results_dir) or not any(f.endswith('.xlsx') for f in os.listdir(results_dir)):
        st.info("No result files found. Run an optimization to generate output files.")
        return

    result_files = [f for f in os.listdir(results_dir) if f.endswith('.xlsx')]
    for file in sorted(result_files):
        file_path = os.path.join(results_dir, file)
        with open(file_path, "rb") as fp:
            st.download_button(
                label=f"📥 Download {file}", data=fp, file_name=file,
                mime="application/vnd.ms-excel", key=f"download_{file}"
            )


# --- 3. Main App Interface ---
if check_password():
    if 'log_output' not in st.session_state:
        st.session_state.log_output = "Log output from the optimization script will appear here in real-time."

    if os.path.exists("logo.png"):
        st.sidebar.image("logo.png", width=150)
    st.sidebar.title("Configuration")
    st.sidebar.markdown("Adjust the model inputs below. The optimization will use these settings.")

    DEFAULT_PARAMS = {
        'PowerParameters': {'wind_size_karnataka': 0.1, 'wind_size_tamil': 0.1, 'wind_size_goa_or_maharashtra': 100.0,
                            'pv_size_gujarat': 0.1, 'pv_size_telangana': 119.0, 'pv_size_rajasthan': 0.1,
                            'pv_size_goa': 25.0, 'dre_size_goa': 1.0, 'nuclear_size': 60.0, 'biomass_size': 42.0,
                            'gas_size': 23.3, 'rtc_size': 325.0, 'annual_demand_mus': 6813.0,
                            'intra_state_power_losses': 0.03, 'inter_state_power_losses': 0.045},
        'BatteryConfigs_NonOptimization': {'min_batt_soc': 0.1, 'batt_efficiency': 0.9, 'battery1_power': 500.0,
                                           'battery1_duration': 4.0, 'battery2_power': 500.0, 'battery2_duration': 6.0,
                                           'battery3_power': 250.0, 'battery3_duration': 4.0},
        'TimePeriods': {'timeline_start_date': '2027-09-01', 'timeline_end_date': '2027-09-05',
                        'zero_pv_goa_start_date': '2027-06-18', 'zero_pv_goa_end_date': '2027-06-19',
                        'zero_pv_goa_start_date2': '2027-07-07', 'zero_pv_goa_end_date2': '2027-07-08'},
        'CostParameters': {'run_thermal_&_sizing_optimization': True, 'penalty_thermal_unmet_demand': 99999.0,
                           'gdam_price_select_year': 2023.0, 'solar_cost_goa': 2000.0, 'solar_cost_guj': 2700.0,
                           'solar_cost_raj': 3000.0, 'solar_cost_tel': 40000.0, 'wind_cost_maha': 4000.0,
                           'wind_cost_tamil': 2800.0, 'wind_cost_karnataka': 2500.0, 'battery_cost_mwh': 4500.0,
                           'penalty_sizing_unmet_demand': 39000.0, 'max_size_batt_mwh': 3200.0,
                           'max_charge_discharge_power_bess': 400.0, 'max_gdam_purchase': 0.1, 'max_solar_goa': 500.0,
                           'min_total_solar': 0.0, 'max_total_solar': 1300.0, 'min_total_wind': 0.0,
                           'max_total_wind': 1000.0, 'min_solar_goa': 380.0, 'min_solar_guj': 566.0,
                           'min_solar_raj': 0.0, 'min_solar_tel': 0.0, 'min_wind_maha': 0.0, 'min_wind_tamil': 50.0,
                           'min_wind_karnataka': 80.0, 'allow_oversized_re': False},
        'MiscParameters': {'shortage_case': 'case2', 'wind_size_excel_sri': 40.0, 'wind_size_excel_seci': 40.0,
                           'wind_size_actual_sri': 450.0, 'wind_size_actual_seci': 450.0},
        'FilePaths': {'file_path': 'Data/combined_demand_2022_2023.csv',
                      'file_path_wind_sri': 'Data/Wind_Analysis_Sri_Morjar_2022.xlsx',
                      'file_path_wind_seci': 'Data/Wind_Analysis_SECI_2024.xlsx',
                      'file_path_solar_goa': 'Data/solar_PV_goa.csv',
                      'file_path_solar_gujarat': 'Data/solar_PV_gujarat.csv',
                      'file_path_solar_rajasthan': 'Data/solar_PV_rajasthan.csv',
                      'file_path_solar_given': 'Data/combined_solar_wind_data_2022_2023.csv',
                      'file_path_generators': 'Data/PPA Life details.xlsx',
                      'file_path_solar_telangana': 'Data/solar_PV_telangana.csv',
                      'file_path_shortage_case1': 'Data/Shortage Case1.xlsx',
                      'file_path_shortage_case2': 'Data/Shortage Case2.xlsx',
                      'file_path_gdam': 'Data/Avg MCP GDAM 2023 and 2024.xlsx'}
    }

    user_params = {}
    for section, params in DEFAULT_PARAMS.items():
        with st.sidebar.expander(f"› {section}", expanded=False):
            # --- THIS IS THE FIX for the KeyError ---
            user_params[section] = {}
            # ----------------------------------------
            for key, value in params.items():
                if isinstance(value, bool):
                    user_params[section][key] = st.checkbox(key.replace('_', ' ').title(), value,
                                                            key=f"{section}_{key}")
                elif isinstance(value, float) or isinstance(value, int):
                    user_params[section][key] = st.number_input(key.replace('_', ' ').title(), value=value,
                                                                key=f"{section}_{key}")
                elif 'date' in key:
                    try:
                        user_params[section][key] = st.date_input(key.replace('_', ' ').title(),
                                                                  value=datetime.strptime(value, '%Y-%m-%d'),
                                                                  key=f"{section}_{key}").strftime('%Y-%m-%d')
                    except:
                        user_params[section][key] = st.text_input(key.replace('_', ' ').title(), value,
                                                                  key=f"{section}_{key}")
                else:
                    if key == 'shortage_case':
                        user_params[section][key] = st.selectbox(key.replace('_', ' ').title(), ('case1', 'case2'),
                                                                 index=('case1', 'case2').index(value),
                                                                 key=f"{section}_{key}")
                    else:
                        user_params[section][key] = st.text_input(key.replace('_', ' ').title(), value,
                                                                  key=f"{section}_{key}")

    st.title("Power System Optimizer")
    tab_run, tab_log, tab_results = st.tabs(["Setup & Run 🚀", "Live Log Output 📝", "Results & Analysis 📊"])

    with tab_run:
        st.header("Start the Optimization")
        st.markdown("Once you have confirmed the settings in the sidebar, click the button below to start the process.")
        run_button = st.button("🚀 Run Optimization", type="primary", use_container_width=True)

    with tab_log:
        st.header("Live Log Output")
        log_placeholder = st.code(st.session_state.log_output, language="log")

    with tab_results:
        display_results()

    if run_button:
        if os.path.exists("Results"):
            for f in os.listdir("Results"):
                os.remove(os.path.join("Results", f))

        st.session_state.log_output = "Configuration saved. Starting optimization process...\n\n"

        config = configparser.ConfigParser()
        for section, params in user_params.items():
            config[section] = {k: str(v) for k, v in params.items()}
        config_path = 'temp_parameters.ini'
        with open(config_path, 'w') as configfile:
            config.write(configfile)

        process = subprocess.Popen(
            [sys.executable, "optimization_model.py", config_path],
            stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True, encoding='utf-8', bufsize=1
        )

        for line in iter(process.stdout.readline, ''):
            st.session_state.log_output += line
            log_placeholder.code(st.session_state.log_output, language="log")

        process.stdout.close()
        return_code = process.wait()

        if return_code == 0:
            st.toast("✅ Optimization finished successfully!", icon="🎉")
            st.balloons()
        else:
            st.toast("❌ Optimization failed. Check logs for errors.", icon="🔥")

        st.rerun()