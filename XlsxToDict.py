import pandas as pd
import numpy as np
import math
import json
import pdb

def read_model_inputs(model_file_path, model_inputs_sheet):



    model_inputs_dfs = pd.read_excel(model_file_path, engine='openpyxl', sheet_name=model_inputs_sheet)
    # pdb.set_trace()
    
    n_AHUs = model_inputs_dfs.shape[1] - 6
    
    models_inputs = {}
    
    for iAHU in range(n_AHUs):
        #dictionary to hold model input values
        model_inputs = {}
        # pdb.set_trace()
        model_inputs_df = pd.concat([model_inputs_dfs.iloc[:, 0:3], model_inputs_dfs.iloc[:, 6+iAHU]], axis=1)
        model_inputs_df.rename(columns={iAHU: 'Value'}, inplace=True)
        model_inputs_df = model_inputs_df[model_inputs_df['No.'].notna()]
        # pdb.set_trace()
        i = 0
        
        while i < 49:
            # if i == 42:
            #     pdb.set_trace()
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
                # pdb.set_trace()
                modified_key = int(type_param.split(' ')[0])
                model_inputs[modified_key] = value_dict
        
        models_inputs[iAHU] = model_inputs
                
    return models_inputs
# def read_model_inputs(model_file_path, model_inputs_sheet):

#     #dictionary to hold model input values
#     model_inputs = {}

#     model_inputs_df = pd.read_excel(model_file_path, engine='openpyxl', sheet_name=model_inputs_sheet)

#     model_inputs_df = model_inputs_df.iloc[:, 0:4]

#     model_inputs_df = model_inputs_df[model_inputs_df['No.'].notna()]

#     i = 0
    
#     while i < 45:
        
#         type_param =str(model_inputs_df.loc[i, 'Type'])
#         model_input_param = model_inputs_df.loc[i, 'Parameter']
#         value = model_inputs_df.loc[i, 'Value']
#         if not 'Setpoint' in type_param:
#             model_inputs[model_input_param] = value
#             i=i+1
        
#         #if type_param value is integer, read corresponding setpoint profiles   
#         else:
#             value_dict            = {}
#             for k in range(i,i+6):
#                 key   = model_inputs_df.loc[k, 'Parameter']
#                 value = model_inputs_df.loc[k, 'Value']
#                 value_dict[key] = value
#                 i += 1
#             # pdb.set_trace()
#             modified_key = int(type_param.split(' ')[2])
#             model_inputs[modified_key] = value_dict
            
#     return model_inputs


def read_weather_data(model_file_path, weather_data_sheet):

    weather_data_df = pd.read_excel(model_file_path, engine='openpyxl', sheet_name=weather_data_sheet, skiprows=1)
    
    weather_data_dict = {}
    weather_data_dict['Hour'] = weather_data_df['Hour'].to_list()
    weather_data_dict['Dry_Bulb_Temperature(°C)'] = weather_data_df['Dry_Bulb_Temperature(°C)'].to_list()
    weather_data_dict['Relative_Humidity(%)'] = weather_data_df['Relative_Humidity(%)'].to_list()
    return weather_data_dict


def read_run_schedule_profile(model_inputs_dict):
    with open('run_schedules.json', 'r') as f:
        schedules = json.load(f)
    n_AHUs = len(model_inputs_dict.keys())
    run_schedules_dicts = {}
    for iAHU in range(n_AHUs):
        run_schedules_dict = {}
        run_schedules_dict['Hour'] = list(range(1,8761))
        
        run_schedules_dict['Run_Schedule_Profile'] = schedules[model_inputs_dict[iAHU]['Hours of operation / Part load operation']]['schedule']
        run_schedules_dicts[iAHU] = run_schedules_dict
    return run_schedules_dicts


def read_thermal_profile(models_inputs_dicts, run_schedules_dicts, flag):
   
    if flag == 'sensible':
        dict_key = 'Sensible gains (Heat)'
    elif flag == 'latent':
        dict_key = 'Latent gains (Steam)'
    else:
        print('Wrong flag!')
        pdb.set_trace()
    n_AHUs = len(run_schedules_dicts.keys())
    thermal_profile_dicts = {}
    for iAHU in range(n_AHUs):
        gains_specific = float(models_inputs_dicts[iAHU][dict_key].split(' ')[2][1:]) #W/m²
        area = float(models_inputs_dicts[iAHU]['Room Area']) # m²
        gains = gains_specific * area / 1000 #kW
        
        thermal_profile = gains * np.array(run_schedules_dicts[iAHU]['Run_Schedule_Profile'])
        
        thermal_profile_dict = {}
        thermal_profile_dict['Hour'] = list(range(1,8761))
        thermal_profile_dict['Thermal_Profile'] = thermal_profile.tolist()
        
        thermal_profile_dicts[iAHU] = thermal_profile_dict

    return thermal_profile_dicts


def read_setpoint_profile(model_inputs_dicts):
    with open('setpoint_profile.json', 'r') as f:
        setpoint_profile = json.load(f)
    n_AHUs = len(model_inputs_dicts.keys())
    setpoint_profile_dicts = {}
    for iAHU in range(n_AHUs):
        setpoint_profile_dict = {}
        setpoint_profile_dict['Hour'] = list(range(1,8761))
        setpoint_profile_dict['Setpoint_Profile'] = setpoint_profile
        setpoint_profile_dicts[iAHU] = setpoint_profile_dict
    return setpoint_profile_dict


def read_inputs(model_file_path, model_inputs_sheet, weather_data_sheet, run_schedule_sheet, thermal_profile_sensible_sheet, thermal_profile_latent_sheet,  setpoint_profile_sheet):

    #read model inputs
    models_inputs_dicts = read_model_inputs(model_file_path, model_inputs_sheet)
    
    #read weather data
    weather_data_dict = read_weather_data(model_file_path, weather_data_sheet)

    #read run schedule
    run_schedule_dicts = read_run_schedule_profile(models_inputs_dicts)

    #read thermal profile
    thermal_profile_sensible_dicts = read_thermal_profile(models_inputs_dicts, run_schedule_dicts, 'sensible')

    #read thermal profile
    thermal_profile_latent_dicts = read_thermal_profile(models_inputs_dicts, run_schedule_dicts, 'latent')
    
    #read setpoint profile
    setpoint_profile_dict = read_setpoint_profile(models_inputs_dicts)

    #construct a dictionary of dictionaries
    n_AHUs = len(run_schedule_dicts.keys())
    inputs_dicts = {}
    for iAHU in range(n_AHUs):
            
        inputs_dict = {
            'model_inputs': models_inputs_dicts[iAHU],
            'weather_data': weather_data_dict,
            'run_schedule': run_schedule_dicts[iAHU],
            'thermal_profile_sensible': thermal_profile_sensible_dicts[iAHU],
            'thermal_profile_latent': thermal_profile_latent_dicts[iAHU],
            'setpoint_profile': setpoint_profile_dict
        }
        
        inputs_dicts[iAHU] = inputs_dict
  
    return inputs_dicts

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
    
    model_file_path = r'Template_Batch.xlsx'
    
    model_inputs_sheet = 'Model Inputs'
    weather_data_sheet = 'Weather Data'
    run_schedule_sheet = 'Run Schedule Profile'
    thermal_profile_sensible_sheet = 'Thermal Profile (Sensible)'
    thermal_profile_latent_sheet = 'Thermal Profile (Latent)'
    setpoint_profile_sheet = 'Set Point Profile'
    detailed_results_sheet = 'Model Results Detailed'
    summary_results_sheet = 'Model Results Summary'
    variations_result_sheet = 'Model Results Variations'
    
    inputs_dicts = read_inputs(model_file_path, model_inputs_sheet, weather_data_sheet, run_schedule_sheet, thermal_profile_sensible_sheet, thermal_profile_latent_sheet, setpoint_profile_sheet)
    for variant in inputs_dicts:
        inputs_dicts[variant] = tweak_inputs_dict(inputs_dicts[variant])   
    
    for key in inputs_dicts:
        print (key)
        
    with open('inputs.json', 'w') as f:
        json.dump(inputs_dicts, f)
    
    
 
        
       
        
            
   
