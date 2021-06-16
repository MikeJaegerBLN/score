import json
import HVAC_CalcEngine_DictBased as HVAC
from copy import deepcopy

def supply_airflow_variations(airflows, inputs_dict, variations, mitigation):
    
    for airflow in airflows:
        inputs_dict_var = deepcopy(inputs_dict)
        inputs_dict_var['model_inputs']['General']['Supply Airflow'] = inputs_dict_var['model_inputs']['General']['Supply Airflow']*airflow
        hvac_obj = HVAC.HVAC_Internal_Calculation(inputs_dict_var)
        summary_values = hvac_obj.get_summary_values()
        
        for result in results:
            variations[mitigation][result].append(summary_values['total'][result+'_total'])
        
        total  = (summary_values['total']['fans_total']+summary_values['total']['preheat_total']+
                  summary_values['total']['heating_total']+summary_values['total']['cooling_total']+
                  summary_values['total']['dehum_total']+summary_values['total']['hum_total'])
        variations[mitigation]['total'].append(total)
    
    return variations
    
def temperature_controlband_variations(controlbands, inputs_dict, variations, mitigation):
    for controlband in controlbands:
        inputs_dict_var = deepcopy(inputs_dict)
        inputs_dict_var['model_inputs']['Set Point Profile 1']['Room Temperature Lower Band'] += -controlband
        inputs_dict_var['model_inputs']['Set Point Profile 1']['Room Temperature Upper Band'] += controlband
        inputs_dict_var['model_inputs']['Set Point Profile 2']['Room Temperature Lower Band'] += -controlband
        inputs_dict_var['model_inputs']['Set Point Profile 2']['Room Temperature Upper Band'] += controlband
        inputs_dict_var['model_inputs']['Set Point Profile 3']['Room Temperature Lower Band'] += -controlband
        inputs_dict_var['model_inputs']['Set Point Profile 3']['Room Temperature Upper Band'] += controlband
        hvac_obj = HVAC.HVAC_Internal_Calculation(inputs_dict_var)
        summary_values = hvac_obj.get_summary_values()
        
        for result in results:
            variations[mitigation][result].append(summary_values['total'][result+'_total'])
        
        total  = (summary_values['total']['fans_total']+summary_values['total']['preheat_total']+
                  summary_values['total']['heating_total']+summary_values['total']['cooling_total']+
                  summary_values['total']['dehum_total']+summary_values['total']['hum_total'])
        variations[mitigation]['total'].append(total)
        
    return variations

def humidity_controlband_variations(controlbands, inputs_dict, variations, mitigation):
    for controlband in controlbands:
        inputs_dict_var = deepcopy(inputs_dict)
        for i in range(1,4):
            inputs_dict_var['model_inputs']['Set Point Profile '+str(i)]['Room Humidity Lower Band'] += -controlband
            inputs_dict_var['model_inputs']['Set Point Profile '+str(i)]['Room Humidity Upper Band'] += controlband
        
        hvac_obj = HVAC.HVAC_Internal_Calculation(inputs_dict_var)
        summary_values = hvac_obj.get_summary_values()
        
        for result in results:
            variations[mitigation][result].append(summary_values['total'][result+'_total'])
        
        total  = (summary_values['total']['fans_total']+summary_values['total']['preheat_total']+
                  summary_values['total']['heating_total']+summary_values['total']['cooling_total']+
                  summary_values['total']['dehum_total']+summary_values['total']['hum_total'])
        variations[mitigation]['total'].append(total)
        
    return variations
    

def temperature_setpoint_variations(setpoints, inputs_dict, variations, mitigation):
    for setpoint in setpoints:
        inputs_dict_var = deepcopy(inputs_dict)
        for i in range(1,4):
            inputs_dict_var['model_inputs']['Set Point Profile '+str(i)]['Room Temperature Lower Band'] += setpoint
            inputs_dict_var['model_inputs']['Set Point Profile '+str(i)]['Room Temperature Upper Band'] += setpoint
            inputs_dict_var['model_inputs']['Set Point Profile '+str(i)]['Room Temperature Setpoint']   += setpoint
        
        hvac_obj = HVAC.HVAC_Internal_Calculation(inputs_dict_var)
        summary_values = hvac_obj.get_summary_values()
        
        for result in results:
            variations[mitigation][result].append(summary_values['total'][result+'_total'])
        
        total  = (summary_values['total']['fans_total']+summary_values['total']['preheat_total']+
                  summary_values['total']['heating_total']+summary_values['total']['cooling_total']+
                  summary_values['total']['dehum_total']+summary_values['total']['hum_total'])
        variations[mitigation]['total'].append(total)
        
    return variations


if __name__ == "__main__":  
    
    inputs_path = 'inputs.json'    
    with open(inputs_path, 'r') as infile:
        inputs_dict = json.load(infile)

    hvac_obj = HVAC.HVAC_Internal_Calculation(inputs_dict)
    detailed_result = hvac_obj.get_detailed_results()
 
    summary_values = hvac_obj.get_summary_values()
         
    with open('summary_results.json', 'w') as outfile2:
        json.dump(summary_values, outfile2)
      
    with open('detailed_results.json', 'w') as outfile:
        json.dump(detailed_result, outfile)
    
    ### Create variations ###
    mitigations = ['supply airflow', 'temperature control bands', 'humidity control bands', 'temperature setpoints']
    results     = ['fans', 'preheat', 'heating', 'cooling', 'dehum', 'hum']
    total       = (summary_values['total']['fans_total']+summary_values['total']['preheat_total']+
                   summary_values['total']['heating_total']+summary_values['total']['cooling_total']+
                   summary_values['total']['dehum_total']+summary_values['total']['hum_total'])
    variations = {}
    for mitigation in mitigations:
        variations[mitigation] = {'fans':    [summary_values['total']['fans_total']],
                           'preheat': [summary_values['total']['preheat_total']],
                           'heating': [summary_values['total']['heating_total']],
                           'cooling': [summary_values['total']['cooling_total']],
                           'dehum':   [summary_values['total']['dehum_total']],
                           'hum':     [summary_values['total']['hum_total']],
                           'total':   [total]
                           }

    airflows                 = [0.95, 0.9, 0.85, 0.8, 0.75, 0.7]
    temperature_controlbands = [0.5, 1, 1.5, 2]
    humidity_controlbands    = [0.05, 0.1, 0.15, 0.2]
    temperature_setpoints    = [-1, 1]
    
    variations = supply_airflow_variations(airflows, inputs_dict, variations, mitigations[0])
    variations = temperature_controlband_variations(temperature_controlbands, inputs_dict, variations, mitigations[1])
    variations = humidity_controlband_variations(humidity_controlbands, inputs_dict, variations, mitigations[2])
    variations = temperature_setpoint_variations(temperature_setpoints, inputs_dict, variations, mitigations[3])
    
    
    with open('variations.json', 'w') as outfile:
        json.dump(variations, outfile)
    
    
       
        
        
        
        
        
        
        