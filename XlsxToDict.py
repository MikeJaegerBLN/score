import pandas as pd
import math
import json


def read_model_inputs(model_file_path, model_inputs_sheet):

    #dictionary to hold model input values
    model_inputs = {}

    model_inputs_df = pd.read_excel(model_file_path, engine='openpyxl', sheet_name=model_inputs_sheet)

    model_inputs_df = model_inputs_df.iloc[:, 0:4]

    model_inputs_df = model_inputs_df[model_inputs_df['No.'].notna()]

    i = 0
    while i < model_inputs_df.shape[0]:
        
        type_param = model_inputs_df.loc[i, 'Type']
        model_input_param = model_inputs_df.loc[i, 'Parameter']
        value = model_inputs_df.loc[i, 'Value']
                
        if math.isnan(type_param):
            model_inputs[model_input_param] = value
            i=i+1
        
        #if type_param value is integer, read corresponding setpoint profiles   
        else:
            value_dict = {}
            
            while i < model_inputs_df.shape[0]:            
                
                value_dict[model_input_param] = value             
                i=i+1
                
                if i == model_inputs_df.shape[0]: break
                
                if (math.isnan(model_inputs_df.loc[i, 'Type'])) & (model_inputs_df.loc[i, 'Parameter'].lower().startswith('room')):
                    model_input_param = model_inputs_df.loc[i, 'Parameter']
                    value = model_inputs_df.loc[i, 'Value']
                    
                else:
                    break
                    
            model_inputs[int(type_param)] = value_dict

    return model_inputs


def read_weather_data(model_file_path, weather_data_sheet):

    weather_data_df = pd.read_excel(model_file_path, engine='openpyxl', sheet_name=weather_data_sheet, skiprows=1)
    
    weather_data_dict = {}
    weather_data_dict['Hour'] = weather_data_df['Hour'].to_list()
    weather_data_dict['Dry_Bulb_Temperature(°C)'] = weather_data_df['Dry_Bulb_Temperature(°C)'].to_list()
    weather_data_dict['Relative_Humidity(%)'] = weather_data_df['Relative_Humidity(%)'].to_list()
    return weather_data_dict


def read_run_schedule_profile(model_file_path, run_schedule_sheet):

    run_schedule_df = pd.read_excel(model_file_path, engine='openpyxl', sheet_name=run_schedule_sheet, skiprows=2)
    #rename first column to date
    run_schedule_df.rename(columns={'Unnamed: 0' : 'Date'}, inplace=True)
    #retain 24 hour columns
    run_schedule_df = run_schedule_df.iloc[:, 0:25]
    
    run_schedule_dict = {}
    run_schedule_dict['Hour'] = []
    run_schedule_dict['Run_Schedule_Profile'] = []
    
    for i in range(8760):
        day                     = int(i/24)
        profile_hour            = i - day*24 + 1
        run_profile_number = run_schedule_df.iloc[day,profile_hour]
        run_schedule_dict['Hour'].append(int(i+1))
        run_schedule_dict['Run_Schedule_Profile'].append(int(run_profile_number))
        
    return run_schedule_dict


def read_thermal_profile(model_file_path, thermal_profile_sheet):

    thermal_profile_df = pd.read_excel(model_file_path, engine='openpyxl', sheet_name=thermal_profile_sheet, skiprows=2)
    #rename first column to date
    thermal_profile_df.rename(columns={'Unnamed: 0' : 'Date'}, inplace=True)
    #retain 24 hour columns
    thermal_profile_df = thermal_profile_df.iloc[:, 0:25]
    
    thermal_profile_dict = {}
    thermal_profile_dict['Hour'] = []
    thermal_profile_dict['Thermal_Profile'] = []
    
    for i in range(8760):
        day                     = int(i/24)
        profile_hour            = i - day*24 + 1
        thermal_profile         = thermal_profile_df.iloc[day,profile_hour]
        thermal_profile_dict['Hour'].append(int(i+1))
        thermal_profile_dict['Thermal_Profile'].append(float(thermal_profile))
    
    return thermal_profile_dict


def read_setpoint_profile(model_file_path, setpoint_profile_sheet):

    setpoint_profile_df = pd.read_excel(model_file_path, engine='openpyxl', sheet_name=setpoint_profile_sheet, skiprows=2)
    #rename first column to date
    setpoint_profile_df.rename(columns={'Unnamed: 0' : 'Date'}, inplace=True)
    #retain 24 hour columns
    setpoint_profile_df = setpoint_profile_df.iloc[:, 0:25]
    
    setpoint_profile_dict = {}
    setpoint_profile_dict['Hour'] = []
    setpoint_profile_dict['Setpoint_Profile'] = []
    
    for i in range(8760):
        day                     = int(i/24)
        profile_hour            = i - day*24 + 1
        setpoint_profile_number = setpoint_profile_df.iloc[day,profile_hour]
        setpoint_profile_dict['Hour'].append(int(i+1))
        setpoint_profile_dict['Setpoint_Profile'].append(int(setpoint_profile_number))
    return setpoint_profile_dict


def read_inputs(model_file_path, model_inputs_sheet, weather_data_sheet, run_schedule_sheet, thermal_profile_sensible_sheet, thermal_profile_latent_sheet,  setpoint_profile_sheet):

    #read model inputs
    model_inputs_dict = read_model_inputs(model_file_path, model_inputs_sheet)

    #read weather data
    weather_data_dict = read_weather_data(model_file_path, weather_data_sheet)

    #read run schedule
    run_schedule_dict = read_run_schedule_profile(model_file_path, run_schedule_sheet)

    #read thermal profile
    thermal_profile_sensible_dict = read_thermal_profile(model_file_path, thermal_profile_sensible_sheet)

    #read thermal profile
    thermal_profile_latent_dict = read_thermal_profile(model_file_path, thermal_profile_latent_sheet)

    #read setpoint profile
    setpoint_profile_dict = read_setpoint_profile(model_file_path, setpoint_profile_sheet)

    #construct a dictionary
    inputs_dict = {
        'model_inputs': model_inputs_dict,
        'weather_data': weather_data_dict,
        'run_schedule': run_schedule_dict,
        'thermal_profile_sensible': thermal_profile_sensible_dict,
        'thermal_profile_latent': thermal_profile_latent_dict,
        'setpoint_profile': setpoint_profile_dict
    }

    return inputs_dict

if __name__ == "__main__":
    
    model_file_path = r'C:\Users\Mike.Jaeger\Desktop\GitLabStuff\hvac-simulation-iia\TestCases\Test 21.xlsx'
    
    model_inputs_sheet = 'Model Inputs'
    weather_data_sheet = 'Weather Data'
    run_schedule_sheet = 'Run Schedule Profile'
    thermal_profile_sensible_sheet = 'Thermal Profile (Sensible)'
    thermal_profile_latent_sheet = 'Thermal Profile (Latent)'
    setpoint_profile_sheet = 'Set Point Profile'
    detailed_results_sheet = 'Model Results Detailed'
    summary_results_sheet = 'Model Results Summary'
    variations_result_sheet = 'Model Results Variations'
    
    inputs_dict = read_inputs(model_file_path, model_inputs_sheet, weather_data_sheet, run_schedule_sheet, thermal_profile_sensible_sheet, thermal_profile_latent_sheet, setpoint_profile_sheet)

    with open('inputs.json', 'w') as outfile:
        json.dump(inputs_dict, outfile)
