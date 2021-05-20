import json
import HVAC_CalcEngine_DictBased as HVAC

if __name__ == "__main__":  
    
    inputs_path = 'inputs.json'    
    with open(inputs_path, 'r') as infile:
        inputs_dict = json.load(infile)
        
    hvac_obj = HVAC.HVAC_Internal_Calculation(inputs_dict)
    detailed_result = hvac_obj.get_detailed_results()
 
    summary_values = hvac_obj.get_summary_values()
        
    with open('detailed_results.json', 'w') as outfile:
        json.dump(detailed_result, outfile)
        
    with open('summary_results.json', 'w') as outfile2:
        json.dump(summary_values, outfile2)
