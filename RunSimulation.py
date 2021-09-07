import json
import HVAC_CalcEngine_DictBased as HVAC
from copy import deepcopy
from openpyxl import load_workbook


def generate_variation(value, inputs_dict, mitigation, results, variations):
    
    hvac_obj = HVAC.HVAC_Internal_Calculation(inputs_dict)
    summary_values = hvac_obj.get_summary_values()
    
    for result in results:
        variations[mitigation]['absolute'][result].append(summary_values['total'][result+'_total'])
        
        try:
            percentage = summary_values['total'][result+'_total']/variations[mitigation]['absolute'][result][0]
        except:
            percentage = 0
        variations[mitigation]['percentage'][result].append(percentage)
    
    total  = (summary_values['total']['fans_total']+summary_values['total']['preheat_total']+
              summary_values['total']['heating_total']+summary_values['total']['cooling_total']+
              summary_values['total']['dehum_total']+summary_values['total']['hum_total'])
    variations[mitigation]['absolute']['total'].append(total)
    
    percentage = total/variations[mitigation]['absolute']['total'][0]
    variations[mitigation]['percentage']['total'].append(percentage)
    
    variations[mitigation]['value'].append(value)
    
    return variations

def supply_airflow_variations(airflows, inputs_dict, variations, mitigation, results):

    for airflow in airflows:
        inputs_dict_var = deepcopy(inputs_dict)
        inputs_dict_var['model_inputs']['General']['Supply Airflow'] = inputs_dict_var['model_inputs']['General']['Supply Airflow']*airflow
        
        variations = generate_variation(airflow, inputs_dict_var, mitigation, results, variations)
        
    return variations
    
def temperature_controlband_variations(controlbands, inputs_dict, variations, mitigation, results):
    for controlband in controlbands:
        inputs_dict_var = deepcopy(inputs_dict)
        for i in range(1,4):
            #if no setpoint is given for no.2 and 3
            if str(inputs_dict_var['model_inputs']['Set Point Profile '+str(i)]['Room Temperature Lower Band']) == 'nan':
                inputs_dict_var['model_inputs']['Set Point Profile '+str(i)]['Room Temperature Lower Band'] = inputs_dict_var['model_inputs']['Set Point Profile 1']['Room Temperature Lower Band']
            if str(inputs_dict_var['model_inputs']['Set Point Profile '+str(i)]['Room Temperature Upper Band']) == 'nan':
                inputs_dict_var['model_inputs']['Set Point Profile '+str(i)]['Room Temperature Upper Band'] = inputs_dict_var['model_inputs']['Set Point Profile 1']['Room Temperature Upper Band']
            
            inputs_dict_var['model_inputs']['Set Point Profile '+str(i)]['Room Temperature Lower Band'] += -controlband
            inputs_dict_var['model_inputs']['Set Point Profile '+str(i)]['Room Temperature Upper Band'] += controlband
        
        variations = generate_variation(controlband, inputs_dict_var, mitigation, results, variations)
        
    return variations

def humidity_controlband_variations(controlbands, inputs_dict, variations, mitigation, results):
    for controlband in controlbands:
        inputs_dict_var = deepcopy(inputs_dict)
        for i in range(1,4):
            #if no setpoint is given for no.2 and 3
            if str(inputs_dict_var['model_inputs']['Set Point Profile '+str(i)]['Room Humidity Lower Band']) == 'nan':
                inputs_dict_var['model_inputs']['Set Point Profile '+str(i)]['Room Humidity Lower Band'] = inputs_dict_var['model_inputs']['Set Point Profile 1']['Room Humidity Lower Band']
            if str(inputs_dict_var['model_inputs']['Set Point Profile '+str(i)]['Room Humidity Upper Band']) == 'nan':
                inputs_dict_var['model_inputs']['Set Point Profile '+str(i)]['Room Humidity Upper Band'] = inputs_dict_var['model_inputs']['Set Point Profile 1']['Room Humidity Upper Band']
            
            inputs_dict_var['model_inputs']['Set Point Profile '+str(i)]['Room Humidity Lower Band'] += -controlband
            inputs_dict_var['model_inputs']['Set Point Profile '+str(i)]['Room Humidity Upper Band'] += controlband
        
        variations = generate_variation(controlband, inputs_dict_var, mitigation, results, variations)
        
    return variations
    
def temperature_setpoint_variations(setpoints, inputs_dict, variations, mitigation, results):
    for setpoint in setpoints:
        inputs_dict_var = deepcopy(inputs_dict)
        for i in range(1,4):
            #if no setpoint is given for no.2 and 3
            if str(inputs_dict_var['model_inputs']['Set Point Profile '+str(i)]['Room Temperature Lower Band']) == 'nan':
                inputs_dict_var['model_inputs']['Set Point Profile '+str(i)]['Room Temperature Lower Band'] = inputs_dict_var['model_inputs']['Set Point Profile 1']['Room Temperature Lower Band']
            if str(inputs_dict_var['model_inputs']['Set Point Profile '+str(i)]['Room Temperature Upper Band']) == 'nan':
                inputs_dict_var['model_inputs']['Set Point Profile '+str(i)]['Room Temperature Upper Band'] = inputs_dict_var['model_inputs']['Set Point Profile 1']['Room Temperature Upper Band']
            if str(inputs_dict_var['model_inputs']['Set Point Profile '+str(i)]['Room Temperature Setpoint']) == 'nan':
                inputs_dict_var['model_inputs']['Set Point Profile '+str(i)]['Room Temperature Setpoint'] = inputs_dict_var['model_inputs']['Set Point Profile 1']['Room Temperature Setpoint']
                    
            inputs_dict_var['model_inputs']['Set Point Profile '+str(i)]['Room Temperature Lower Band'] += setpoint
            inputs_dict_var['model_inputs']['Set Point Profile '+str(i)]['Room Temperature Upper Band'] += setpoint
            inputs_dict_var['model_inputs']['Set Point Profile '+str(i)]['Room Temperature Setpoint']   += setpoint
        
        variations = generate_variation(setpoint, inputs_dict_var, mitigation, results, variations)
           
    return variations

def create_variations(summary_values, variant):
    
    global mitigations
    mitigations = ['supply airflow', 'temperature control bands', 'humidity control bands', 'temperature setpoints']
    results     = ['fans', 'preheat', 'heating', 'cooling', 'dehum', 'hum']
    total       = (summary_values['total']['fans_total']+summary_values['total']['preheat_total']+
                   summary_values['total']['heating_total']+summary_values['total']['cooling_total']+
                   summary_values['total']['dehum_total']+summary_values['total']['hum_total'])
    variations = {}
    for mitigation in mitigations:
        variations[mitigation] = {}
        variations[mitigation]['absolute'] = {'fans':    [summary_values['total']['fans_total']],
                                              'preheat': [summary_values['total']['preheat_total']],
                                              'heating': [summary_values['total']['heating_total']],
                                              'cooling': [summary_values['total']['cooling_total']],
                                              'dehum':   [summary_values['total']['dehum_total']],
                                              'hum':     [summary_values['total']['hum_total']],
                                              'total':   [total]
                                              }
        variations[mitigation]['percentage'] = {'fans':  [1],
                                              'preheat': [1],
                                              'heating': [1],
                                              'cooling': [1],
                                              'dehum':   [1],
                                              'hum':     [1],
                                              'total':   [1]
                                              }
        if mitigation==mitigations[0]:
            variations[mitigation]['value'] = [1] 
        if mitigation==mitigations[1] or mitigation==mitigations[2] or mitigation==mitigations[3]:
            variations[mitigation]['value'] = [0] 

    airflows                 = [0.95, 0.9, 0.85, 0.8, 0.75, 0.7]
    temperature_controlbands = [0.5, 1, 1.5, 2, 10]
    humidity_controlbands    = [0.05, 0.1, 0.15, 0.2]
    temperature_setpoints    = [-1, 1]
    
    variations = supply_airflow_variations(airflows, variant, variations, mitigations[0], results)
    variations = temperature_controlband_variations(temperature_controlbands, variant, variations, mitigations[1], results)
    variations = humidity_controlband_variations(humidity_controlbands, variant, variations, mitigations[2], results)
    variations = temperature_setpoint_variations(temperature_setpoints, variant, variations, mitigations[3], results)
    
    return variations

def write_summary_values(model_file_path, summary_results_sheet, summary_values):

    wb = load_workbook(model_file_path)
    wsh = wb[summary_results_sheet]

    #writing min values
    min_keys = ['fans_min', 'preheat_min', 'dehum_min', 'hum_min', 'heating_min', 'cooling_min']
    for i in range(6):
        wsh.cell(column = i+3, row = 7, value = summary_values['min'][min_keys[i]]) 

    #writing max values
    max_keys = ['fans_max', 'preheat_max', 'dehum_max', 'hum_max', 'heating_max', 'cooling_max']
    for i in range(6):
        wsh.cell(column = i+3, row = 8, value = summary_values['max'][max_keys[i]]) 

    #writing avg values
    avg_keys = ['fans_avg', 'preheat_avg', 'dehum_avg', 'hum_avg', 'heating_avg', 'cooling_avg']
    for i in range(6):
        wsh.cell(column = i+3, row = 9, value = summary_values['avg'][avg_keys[i]]) 

    #writing total values
    total_keys = ['fans_total', 'preheat_total', 'dehum_total', 'hum_total', 'heating_total', 'cooling_total']
    for i in range(6):
        wsh.cell(column = i+3, row = 10, value = summary_values['total'][total_keys[i]]) 

    #writing kpis
    kpi_keys = ['hvac_total_energy_area', 'hvac_heating', 'hvac_cooling', 'hvac_humidification', 'hvac_dehumidification', 'hvac_fans']
    for i in range(6):
        wsh.cell(column = 10, row = i+13, value = summary_values['kpi'][kpi_keys[i]]) 

    #writing hours within
    within_keys = ['tdb_hours','rh_hours','tdb_rh_hours']
    for i in range(3):
        wsh.cell(column = i+2, row = 15, value = summary_values['within'][within_keys[i]])

    #writing hours outside
    outside_keys = ['tdb_hours','rh_hours','tdb_rh_hours']
    for i in range(3):
        wsh.cell(column = i+2, row = 18, value = summary_values['outside'][outside_keys[i]])

    # Save the file
    wb.save(model_file_path)

def write_detailed_result(model_file_path, detailed_results_sheet, detailed_result):

    #load workbook and point to Model Results Detailed sheet
    wb = load_workbook(model_file_path)
    wsh = wb[detailed_results_sheet]

    # clear the existing values
    for row in wsh['B2:AR8990']:
        for cell in row:
            cell.value = None

    #write result
    output_cols = []
    for i in range(2,100):
        if wsh.cell(1, i).value is None:
            break
        output_cols.append(wsh.cell(1, i).value)

    #output_cols = ['Room_Sensible_Load[kW]', 'Room_Latent_Load[kW]', 'Room_Ratio_Line', 'Supply_Fan_Power[kW]', 'Return_Fan_Power[kW]', 'Airflow_Fresh[kg/s]', 'Airflow_Return[kg/s]', 'Airflow_Total_Supply[kg/s]', 'Preheat_Load[kW_S]', 'Heating_Load[kW_S]', 'Reheat_Load[kW_S]', 'Cooling_Load[kW_L]', 'Cooling_Load[kW_S]', 'Dehumidification_Load[kW_L]', 'Dehumidification_Load[kW_S]', 'Humidification_Load[kW_L]', 'Humidification_Load[kW_S]', 'Humidification_Water[L]', 'Fresh_Air_Tdb', 'Fresh_Air_RH', 'Return_Air_Tdb', 'Return_Air_RH', 'Post_Preheat_Tdb', 'Post_Preheat_RH', 'Post_HR_Air_Tdb', 'Post_HR_Air_RH', 'Post_Heat_Tdb', 'Post_Heat_RH', 'Post_Cool_Tdb', 'Post_Cool_RH', 'Post_Dehum_Tdb', 'Post_Dehum_RH', 'Post_Reheat_Tdb', 'Post_Reheat_RH', 'Post_Hum_Tdb', 'Post_Hum_RH', 'Required_Supply_Air_Tdb', 'Required_Supply_Air_RH', 'Final_Supply_Air_Tdb', 'Final_Supply_Air_RH', 'Final_Room_Air_Tdb', 'Final_Room_Air_RH', 'AHU_Mode']
    for i, key in enumerate(output_cols):
        if len(detailed_result[key]) > 10: # If the result has values
                for j in range (0,8760,1):
                    wsh.cell(column = i+2, row = j+2, value = detailed_result[key][j]) 
                    #if j==4: break

    # Save the file
    wb.save(model_file_path)

def write_variations_result(model_file_path, variations_result_sheet, variation_result, basecase_result):
    
    #read the file
    wb = load_workbook(model_file_path)
    wsh = wb[variations_result_sheet]
    
    #keys correponsding to columns in the variation results sheet
    summary_keys = ['fans_total', 'preheat_total', 'dehum_total', 'hum_total', 'heating_total', 'cooling_total']
    
    #Update the variations model inputs
    #for each key
    for i in range(7, 50):
        
        #if all the variations are populated, break
        if wsh.cell(i, 2).value is None:
            break

        key = wsh.cell(i, 2).value
        
        #for each column
        for j in range(5, 11):
            
            if key == 'Base Scenario':
                #write base case results
                wsh.cell(column = j, row = i, value = basecase_result[summary_keys[j-5]]) 
            else:   
                #write variations results
                wsh.cell(column = j, row = i, value = variation_result[key][summary_keys[j-5]]) 

    # Save the file
    wb.save(model_file_path)

if __name__ == "__main__":  
    
    inputs_path = 'inputs.json'    
    with open(inputs_path, 'r') as infile:
        inputs_dicts = json.load(infile)
    
    #if only one variant
    if not '0' in inputs_dicts.keys():
        inputs_dicts['0'] = inputs_dicts
   
    #if multiple variants
    for variant in inputs_dicts:
        hvac_obj                 = HVAC.HVAC_Internal_Calculation(inputs_dicts[variant])
        detailed_results         = hvac_obj.get_detailed_results()
        summary_values           = hvac_obj.get_summary_values()
        variation_summary_values = create_variations(summary_values, inputs_dicts[variant])
   
        case = 'Template_wResults.xlsx'
        model_file_path = r'C:\Users\Mike.Jaeger\Desktop\GitLabStuff\ahes-frontend\{}'.format(case)
    
        write_summary_values(model_file_path, 'Model Results Summary', summary_values)        
        write_detailed_result(model_file_path, 'Model Results Detailed', detailed_results)
    #write_variations_result(model_file_path, 'Model Results Variations', variation_summary_values, summary_values)
    
        
        
        
        