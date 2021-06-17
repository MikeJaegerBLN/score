import json
import HVAC_CalcEngine_DictBased as HVAC
from copy import deepcopy


def generate_variation(value, inputs_dict, mitigation, results):
    
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
        
        variations = generate_variation(airflow, inputs_dict_var, mitigation, results)
        
    return variations
    
def temperature_controlband_variations(controlbands, inputs_dict, variations, mitigation, results):
    for controlband in controlbands:
        inputs_dict_var = deepcopy(inputs_dict)
        for i in range(1,4):
            inputs_dict_var['model_inputs']['Set Point Profile '+str(i)]['Room Temperature Lower Band'] += -controlband
            inputs_dict_var['model_inputs']['Set Point Profile '+str(i)]['Room Temperature Upper Band'] += controlband
        
        variations = generate_variation(controlband, inputs_dict_var, mitigation, results)
        
    return variations

def humidity_controlband_variations(controlbands, inputs_dict, variations, mitigation, results):
    for controlband in controlbands:
        inputs_dict_var = deepcopy(inputs_dict)
        for i in range(1,4):
            inputs_dict_var['model_inputs']['Set Point Profile '+str(i)]['Room Humidity Lower Band'] += -controlband
            inputs_dict_var['model_inputs']['Set Point Profile '+str(i)]['Room Humidity Upper Band'] += controlband
        
        variations = generate_variation(controlband, inputs_dict_var, mitigation, results)
        
    return variations
    
def temperature_setpoint_variations(setpoints, inputs_dict, variations, mitigation, results):
    for setpoint in setpoints:
        inputs_dict_var = deepcopy(inputs_dict)
        for i in range(1,4):
            inputs_dict_var['model_inputs']['Set Point Profile '+str(i)]['Room Temperature Lower Band'] += setpoint
            inputs_dict_var['model_inputs']['Set Point Profile '+str(i)]['Room Temperature Upper Band'] += setpoint
            inputs_dict_var['model_inputs']['Set Point Profile '+str(i)]['Room Temperature Setpoint']   += setpoint
        
        variations = generate_variation(setpoint, inputs_dict_var, mitigation, results)
           
    return variations


if __name__ == "__main__":  
    
    inputs_path = 'inputs.json'    
    with open(inputs_path, 'r') as infile:
        inputs_dict = json.load(infile)

    hvac_obj = HVAC.HVAC_Internal_Calculation(inputs_dict)
    detailed_results = hvac_obj.get_detailed_results()
 
    summary_values = hvac_obj.get_summary_values()
         
    
    
    ### Create variations ###
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
        variations[mitigation]['percentage'] = {'fans':  [100],
                                              'preheat': [100],
                                              'heating': [100],
                                              'cooling': [100],
                                              'dehum':   [100],
                                              'hum':     [100],
                                              'total':   [100]
                                              }
        if mitigation==mitigations[0]:
            variations[mitigation]['value'] = [1] 
        if mitigation==mitigations[1] or mitigation==mitigations[2] or mitigation==mitigations[3]:
            variations[mitigation]['value'] = [0] 

    airflows                 = [0.95, 0.9, 0.85, 0.8, 0.75, 0.7]
    temperature_controlbands = [0.5, 1, 1.5, 2, 10]
    humidity_controlbands    = [0.05, 0.1, 0.15, 0.2]
    temperature_setpoints    = [-1, 1]
    
    variations = supply_airflow_variations(airflows, inputs_dict, variations, mitigations[0], results)
    variations = temperature_controlband_variations(temperature_controlbands, inputs_dict, variations, mitigations[1], results)
    variations = humidity_controlband_variations(humidity_controlbands, inputs_dict, variations, mitigations[2], results)
    variations = temperature_setpoint_variations(temperature_setpoints, inputs_dict, variations, mitigations[3], results)
  
    #summary_values.json
    #detailed_results.json
    #variations.json
        
    
    
       
        
        
        
        
        
        
        