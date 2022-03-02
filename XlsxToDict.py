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
    
    while i < 45:
        
        type_param =str(model_inputs_df.loc[i, 'Type'])
        model_input_param = model_inputs_df.loc[i, 'Parameter']
        value = model_inputs_df.loc[i, 'Value']
        if not 'Setpoint' in type_param:
            model_inputs[model_input_param] = value
            i=i+1
        
        #if type_param value is integer, read corresponding setpoint profiles   
        else:
            value_dict            = {}
            for k in range(i,i+6):
                key   = model_inputs_df.loc[k, 'Parameter']
                value = model_inputs_df.loc[k, 'Value']
                value_dict[key] = value
                i += 1
            
            modified_key = int(type_param.split(' ')[2])
            model_inputs[modified_key] = value_dict
            
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

def tweak_inputs_dict(inputs_dict):
    for key in inputs_dict['model_inputs']:
        if 'HVAC Module' in str(key):
            if str(inputs_dict['model_inputs'][key])=='nan':
                inputs_dict['model_inputs'][key] = 0
                
    inputs_dict['model_inputs']['Set Point Profile 1'] = inputs_dict['model_inputs'][1]
    inputs_dict['model_inputs']['Set Point Profile 2'] = inputs_dict['model_inputs'][2]
    inputs_dict['model_inputs']['Set Point Profile 3'] = inputs_dict['model_inputs'][3]
    del inputs_dict['model_inputs'][1]
    del inputs_dict['model_inputs'][2]
    del inputs_dict['model_inputs'][3]
    inputs_dict['model_inputs']['General'] = {}
    keys = []
    for key in inputs_dict['model_inputs'].keys():
        if not 'Set Point Profile' in key and not 'General' in key:
            inputs_dict['model_inputs']['General'][key] = inputs_dict['model_inputs'][key]
            keys.append(key)
    for key in keys:
        del inputs_dict['model_inputs'][key]
        
    for key in inputs_dict:   
        for key2 in inputs_dict[key]:
            if str(inputs_dict[key][key2])=='nan' or str(inputs_dict[key][key2])=='NaN':
                inputs_dict[key][key2] = 'nan'  
            try:
                for key3 in inputs_dict[key][key2]:
                    if str(inputs_dict[key][key2][key3])=='nan' or str(inputs_dict[key][key2][key3])=='NaN':
                        inputs_dict[key][key2][key3] = 'nan'
            except:
                pass
        
    return inputs_dict

if __name__ == "__main__":
    
    with open('run_schedules.json', 'r') as f:
        schedules = json.load(f)
    
    model_file_path = r'Bayer_Bergkamen_9.xlsx'
    
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
    inputs_dict = tweak_inputs_dict(inputs_dict)
            
    #inputs_dict['run_schedule']['Run_Schedule_Profile'] = schedules[key_to_read]
    
    #for key in inputs_dict['model_inputs']['General']:
    #    print (key)
    with open('inputs.json', 'w') as f:
        json.dump(inputs_dict, f)
    
    
        
    #for key in schedules:
    #    print (schedules[key])
    
    #run_schedules = {}
    #schedule['schedule'] = inputs_dict['run_schedule']['Run_Schedule_Profile']
    
    #with open('daytime_on_nighttime_off_sat_sun_off.json', 'r') as f:
    #    run_schedules['Daytime on, Nighttime off, Sat&Sun off'] = json.load(f)
        
    #with open('daytime_on_nighttime_halfpower_sat_sun_off.json', 'r') as f:
    #    run_schedules['Daytime on, Nighttime halfpower, Sat&Sun off'] = json.load(f)
        
    #with open('daytime_on_nighttime_halfpower_sat_sun_halfpower.json', 'r') as f:
    #    run_schedules['Daytime on, Nighttime halfpower, Sat&Sun halfpower'] = json.load(f)
        
    #with open('always_on.json', 'r') as f:
    #    run_schedules['Always on'] = json.load(f)
        
    #with open('always_on_sat_sun_off.json', 'r') as f:
    #    run_schedules['Always on, Sat&Sun off'] = json.load(f)
        
    #with open('always_on_sat_sun_halfpower.json', 'r') as f:
    #    run_schedules['Always on, Sat&Sun halfpower'] = json.load(f)
        
    #with open('run_schedules.json', 'w') as f:
    #    json.dump(run_schedules, f)
   
        
       
        
            
   
