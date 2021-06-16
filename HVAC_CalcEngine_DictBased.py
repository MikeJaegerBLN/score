import HVAC_PsyEquations as PsyCalc
from matplotlib import pyplot
import numpy

class HVAC_Internal_Calculation(object):    
    
    def __init__(self, inputs_dict):
        
        self.inputs_dict                       = inputs_dict
        self.cp_air                            = 1005    #J/kg/K
        self.rho_air                           = 1.2041  #kg/m^3 @ 20 deg C
        self.ambient_pressure                  = 101326  #Pa
        self.evaporation_enthalpy_water        = 2501000 #J/kg
        
        self.ambient_temperature               = self.inputs_dict['weather_data']['Dry_Bulb_Temperature(°C)']
        self.ambient_rh                        = self.inputs_dict['weather_data']['Relative_Humidity(%)']
        self.ambient_w                         = []
        for i, t in enumerate(self.ambient_temperature):
            self.ambient_w.append(PsyCalc.w_of_tdb_rh(t,self.ambient_rh[i], self.ambient_pressure))
        
        self.chilled_water_temperature         = self.inputs_dict['model_inputs']['General']['Average Cooling Coil Temperature - Tdb']
        self.room_area                         = self.inputs_dict['model_inputs']['General']['Room Area']
        self.room_height                       = self.inputs_dict['model_inputs']['General']['Room Height']
        self.room_height                       = self.inputs_dict['model_inputs']['General']['Room Height']
        
        self.setpoint_profile                  = self.inputs_dict['setpoint_profile']['Setpoint_Profile']
        self.run_schedule_profile              = self.inputs_dict['run_schedule']['Run_Schedule_Profile']
        self.thermal_profile_sensible          = self.inputs_dict['thermal_profile_sensible']['Thermal_Profile']
        self.thermal_profile_latent            = self.inputs_dict['thermal_profile_latent']['Thermal_Profile']
        
        #initialize tdb and rh setpoints according to hour 0, setpoints are updated every hour in the calculation
        self.set_setpoints(0)
               
        self.supply_fan_pressure               = self.inputs_dict['model_inputs']['General']["Supply Fan Total Static Pressure"]
        self.return_fan_pressure               = self.inputs_dict['model_inputs']['General']["Return Fan Total Static Pressure"]
        self.supply_fan_efficiency             = self.inputs_dict['model_inputs']['General']["Supply Fan Efficiency"]
        self.return_fan_efficiency             = self.inputs_dict['model_inputs']['General']["Return Fan Efficiency"]
        self.min_fresh_air_percentage          = self.inputs_dict['model_inputs']['General']["Minimum Fresh Air Percentage"]
        self.active_mixing                     = self.inputs_dict['model_inputs']['General']["Active Mixing"]
        self.heat_recovery_efficiency          = self.inputs_dict['model_inputs']['General']["Heat Recovery Efficiency (S)"]
        self.HVAC_systemtype, self.check_dehum = self.get_HVAC_systemtype(inputs_dict['model_inputs']['General'])
        
        #inital state = room setpoint
        self.current_HVAC_tdb                  = self.room_temperature_setpoint
        self.current_HVAC_twb                  = self.room_temperature_setpoint
        self.current_HVAC_rh                   = self.room_humidity_setpoint 
        self.current_HVAC_w                    = PsyCalc.w_of_tdb_rh(self.room_temperature_setpoint, 
                                                                     self.room_humidity_setpoint, 
                                                                     self.ambient_pressure)   
        
        self.room_temperature                  = self.room_temperature_setpoint
        self.room_humidity                     = self.room_humidity_setpoint
        self.room_w                            = self.current_HVAC_w
        
        self.results_room_tdb                  = []
        self.results_room_rh                   = []
        self.results_room_w                    = []
        self.results_required_supply_tdb       = []
        self.results_required_supply_rh        = []
        
        self.results_heat_load                 = []
        self.results_preheat_load              = []
        self.results_cool_load                 = []
        self.results_dehum_load                = []
        self.results_hum_load                  = []
        self.results_supply_fan_power          = []
        self.results_return_fan_power          = []
        self.results_supply_airflow            = []
        self.results_return_airflow            = []
        self.results_airflow_fresh             = []
        self.results_airflow_fresh_percentage  = [] 
        self.results_room_ratio_line           = [] 
        self.results_final_supply_air_tdb      = []
        self.results_final_supply_air_rh       = []
        self.results_tdb_within                = 0
        self.results_rh_within                 = 0
        self.results_tdb_outside               = 0
        self.results_rh_outside                = 0
        self.results_tdb_rh_within             = 0
        self.results_tdb_rh_outside            = 0
        
        self.results_post_preheat_tdb          = []
        self.results_post_preheat_rh           = []
        self.results_post_HR_air_tdb           = []
        self.results_post_HR_air_rh            = []
        self.results_post_heat_tdb             = []
        self.results_post_heat_rh              = []
        self.results_post_cool_tdb             = []
        self.results_post_cool_rh              = []
        self.results_post_dehum_tdb            = []
        self.results_post_dehum_rh             = []
        self.results_post_hum_tdb              = []
        self.results_post_hum_rh               = []
        self.results_humidification_water      = []
        self.results_ahu_mode                  = []
          
        self.results_run_schedule              = []
        self.results_thermal_profile_sensible  = []
        self.results_thermal_profile_latent    = []
        self.results_setpoint_profile          = []
        
        #for internal tracking of HVAC properties after each step
        self.HVAC_internal_properties = {}  
        
        for hour in range(len(self.ambient_temperature)):
            
            #for internal tracking of HVAC properties after each step
            self.HVAC_internal_properties[hour] = {}  
            self.prepare_for_internal_HVAC_tracking('Step0', hour) 
            
            #update setpoints according to setpoint proifle
            self.set_setpoints(hour)
            
            #update supply / return airflow according to run schedule profile
            self.set_run_schedule(hour)
            
            #update room heat gain according to thermal profile
            self.set_room_heat_gain(hour)
            
            #update required supply tdb and room delta tdb
            self.set_required_room_supply_tdb_and_w(hour)
            
            #for tracking what ahu is doing
            self.ahu_mode = ''
            for step, entry in enumerate(self.HVAC_systemtype):    
            
                #for internal tracking of HVAC properties after each step
                self.prepare_for_internal_HVAC_tracking('Step'+str(step+1), hour) 
              
                #if AHU is not in operation
                if self.supply_airflow==0:
                    self.DoNothing('Step'+str(step+1), hour)
                 
                #else AHU is in operation
                elif 'Fresh' in entry:
                    self.FreshAir('Step'+str(step+1), hour)
                elif 'Mix' in entry:
                    self.Mix('Step'+str(step+1), hour)
                elif 'Plate Heat Exchanger' in entry:
                    self.PHE('Step'+str(step+1), hour)
                elif 'Thermal Wheel' in entry:
                    self.ThermalWheel('Step'+str(step+1), hour)
                elif 'Run Around Coil' in entry:
                    self.RunAroundCoil('Step'+str(step+1), hour)
                elif 'Heat' in entry:
                    self.Heat('Step'+str(step+1), hour)
                elif 'Preheat' in entry:
                    self.Preheat('Step'+str(step+1), hour)
                elif 'Cool' in entry:
                    self.Cool('Step'+str(step+1), hour)
                elif 'Steam' in entry:
                    self.SteamHum('Step'+str(step+1), hour)
                elif 'Spray' in entry:
                    self.SprayHum('Step'+str(step+1), hour)  
               
            #calculate fan power supply / return
            self.calculate_fan_power()
                    
            #update room conditions tdb, rh and w
            self.update_room_conditions('Step'+str(step+1), hour)
            
            #update all loads with 0 which dont occur, e.g. HVAC type is MIX HEAT COOL the HUM and DEHUM load is populated w/ 0 then
            self.update_zero_loads(hour)
            
    def get_detailed_results(self):
        
        detailed_results = {
            'Room_Sensible_Load[kW]' : self.results_thermal_profile_sensible,
            'Room_Latent_Load[kW]' : self.results_thermal_profile_latent,
            'Room_Ratio_Line' : self.results_room_ratio_line,
            'Supply_Fan_Power[kW]' : self.results_supply_fan_power,
            'Return_Fan_Power[kW]' : self.results_return_fan_power,
            'Airflow_Fresh[kg/s]' : self.results_airflow_fresh,
            'Airflow_Return[kg/s]' : self.results_return_airflow,
            'Airflow_Total_Supply[kg/s]' : self.results_supply_airflow,
            'Fresh_Air_Percentage': self.results_airflow_fresh_percentage,
            'Preheat_Load[kW]' : self.results_preheat_load,
            'Heating_Load[kW]' : self.results_heat_load,
            'Cooling_Load[kW]' : self.results_cool_load,
            'Dehumidification_Load[kW]' : self.results_dehum_load,
            'Humidification_Load[kW]' : self.results_hum_load,
            'Humidification_Water[L]' : self.results_humidification_water,
            'Fresh_Air_Tdb' : self.ambient_temperature,
            'Fresh_Air_RH' : self.ambient_rh,
            'Return_Air_Tdb' : self.results_room_tdb,
            'Return_Air_RH' : self.results_room_rh,
            'Post_Preheat_Tdb' : self.results_post_preheat_tdb,
            'Post_Preheat_RH' : self.results_post_preheat_rh,
            'Post_HR_Air_Tdb' : self.results_post_HR_air_tdb,
            'Post_HR_Air_RH' : self.results_post_HR_air_rh,
            'Post_Heat_Tdb' : self.results_post_heat_tdb,
            'Post_Heat_RH' : self.results_post_heat_rh,
            'Post_Cool_Tdb' : self.results_post_cool_tdb,
            'Post_Cool_RH' : self.results_post_cool_rh,
            'Post_Dehum_Tdb' : self.results_post_dehum_tdb,
            'Post_Dehum_RH' : self.results_post_dehum_rh,
            'Post_Hum_Tdb' : self.results_post_hum_tdb,
            'Post_Hum_RH' : self.results_post_hum_rh,
            'Required_Supply_Air_Tdb' : self.results_required_supply_tdb,
            'Required_Supply_Air_RH' : self.results_required_supply_rh,
            'Final_Supply_Air_Tdb' : self.results_final_supply_air_tdb,
            'Final_Supply_Air_RH' : self.results_final_supply_air_rh,
            'Final_Room_Air_Tdb' : self.results_room_tdb,
            'Final_Room_Air_RH' : self.results_room_rh,
            'AHU_Mode' : self.results_ahu_mode
            }
        return detailed_results
    
    def get_summary_values(self):
        total_energy = (numpy.sum(self.results_heat_load) + numpy.sum(self.results_cool_load) + numpy.sum(self.results_hum_load) +
                        numpy.sum(self.results_preheat_load) + numpy.sum(self.results_dehum_load) + numpy.sum(self.results_supply_fan_power) + 
                        numpy.sum(self.results_return_fan_power))
    
        
        summary_values = {'min': {
                'fans_min' : numpy.min(numpy.min(self.results_supply_fan_power)+numpy.min(self.results_return_fan_power)),
                'preheat_min' : numpy.min(self.results_preheat_load), 
                'dehum_min' : numpy.min(self.results_dehum_load),
                'hum_min' : numpy.min(self.results_hum_load),
                'heating_min' : numpy.min(self.results_heat_load),
                'cooling_min' : numpy.min(self.results_cool_load)
                },
            
                          'max': {
                'fans_max' : numpy.max(numpy.max(self.results_supply_fan_power)+numpy.max(self.results_return_fan_power)),
                'preheat_max' : numpy.max(self.results_preheat_load),
                'dehum_max' : numpy.max(self.results_dehum_load), 
                'hum_max' : numpy.max(self.results_hum_load),
                'heating_max' : numpy.max(self.results_heat_load), 
                'cooling_max' : numpy.max(self.results_cool_load)
                },
                          'avg': {
                'fans_avg' : numpy.mean(numpy.mean(self.results_supply_fan_power)+numpy.mean(self.results_return_fan_power)),
                'preheat_avg' : numpy.mean(self.results_preheat_load),      
                'dehum_avg' : numpy.mean(self.results_dehum_load),
                'hum_avg' : numpy.mean(self.results_hum_load),
                'heating_avg' : numpy.mean(self.results_heat_load), 
                'cooling_avg' : numpy.mean(self.results_cool_load)
                },
                          'total': {
                'fans_total' : numpy.sum(self.results_supply_fan_power)+numpy.sum(self.results_return_fan_power),
                'preheat_total' : numpy.sum(self.results_preheat_load), 
                'dehum_total' : numpy.sum(self.results_dehum_load),
                'hum_total' : numpy.sum(self.results_hum_load),
                'heating_total' : numpy.sum(self.results_heat_load), 
                'cooling_total' : numpy.sum(self.results_cool_load)
                },
                          'kpi': {
                'hvac_total_energy_area' : total_energy/self.room_area,
                'hvac_heating' : numpy.sum(self.results_heat_load)/(self.room_area), 
                'hvac_cooling' : numpy.sum(self.results_cool_load)/(self.room_area), 
                'hvac_humidification' : numpy.sum(self.results_hum_load)/(self.room_area), 
                'hvac_dehumidification' : numpy.sum(self.results_dehum_load)/(self.room_area), 
                'hvac_fans' : (numpy.sum(self.results_supply_fan_power)+numpy.sum(self.results_return_fan_power))/(self.room_area)
                },
                          'within': {
                'tdb_hours': float(self.results_tdb_within),
                'rh_hours': float(self.results_rh_within),
                'tdb_rh_hours': float(self.results_tdb_rh_within)
                },
                          'outside': {
                'tdb_hours': float(self.results_tdb_outside),
                'rh_hours': float(self.results_rh_outside),
                'tdb_rh_hours': float(self.results_tdb_rh_outside) 
                    }
                }
        
        for key in summary_values:
            for key_further in summary_values[key]:
                summary_values[key][key_further] = float(summary_values[key][key_further])
                
        
        return summary_values
        
    def set_required_room_supply_tdb_and_w(self, hour):
        
        if self.supply_airflow!=0:
            self.room_delta_tdb = self.room_heat_gain_tdb / (self.supply_airflow * self.rho_air * self.cp_air)
            self.supply_tdb     = self.room_temperature_setpoint - self.room_delta_tdb
            
            h_setpoint          = PsyCalc.enthalpy_moist_air(self.room_temperature_setpoint, self.room_w_setpoint) #kJ/kg
            h_supply            = h_setpoint - self.room_heat_gain_w/(1000 * self.supply_airflow * self.rho_air) #kJ/kg
            self.supply_w       = PsyCalc.w_of_enthalpy_tdb(h_supply, self.room_temperature_setpoint)
            self.room_delta_w   = self.room_w_setpoint - self.supply_w
        else:
            self.room_delta_tdb = 0
            self.supply_tdb     = 0
            self.room_delta_w   = 0
            self.supply_w       = 0
            
        partial_pressure      = PsyCalc.partial_pressure_of_w(self.ambient_pressure, self.supply_w)
        saturation_pressure   = PsyCalc.saturation_pressure(self.supply_tdb)
        supply_rh             = partial_pressure/saturation_pressure
  
        self.results_required_supply_tdb.append(self.supply_tdb)
        self.results_required_supply_rh.append(supply_rh)
      
    def set_room_heat_gain(self, hour):
        
        day                     = int(hour/24)
        profile_hour            = hour - day*24 + 1
        self.room_heat_gain_tdb = self.thermal_profile_sensible[hour]*1000   #unit in Watts
        self.room_heat_gain_w   = self.thermal_profile_latent[hour]*1000    #unit Watts
        
        self.results_thermal_profile_sensible.append(self.room_heat_gain_tdb/1000)
        self.results_thermal_profile_latent.append(self.room_heat_gain_w/1000)
        if self.room_heat_gain_tdb!=0:
            self.results_room_ratio_line.append(self.room_heat_gain_tdb/(self.room_heat_gain_tdb+self.room_heat_gain_w))
        else:
            self.results_room_ratio_line.append(0)
        
    def set_run_schedule(self, hour):
        
        day                     = int(hour/24)
        profile_hour            = hour - day*24 + 1
        run_profile_percentage  = self.run_schedule_profile[hour]
        
        self.supply_airflow     = self.inputs_dict['model_inputs']['General']['Supply Airflow']*run_profile_percentage
        self.return_airflow     = self.inputs_dict['model_inputs']['General']['Return Airflow']*run_profile_percentage
        
        self.results_run_schedule.append(run_profile_percentage)
        self.results_supply_airflow.append(float(self.supply_airflow))
        self.results_return_airflow.append(float(self.return_airflow))
        
      

    def set_setpoints(self, hour):

        day                     = int(hour/24)
        profile_hour            = hour - day*24 + 1
        setpoint_profile_number = self.setpoint_profile[hour]
        
        self.preheat_setpoint                = self.inputs_dict['model_inputs']['General']['Preheat Temperature Setpoint - Tdb']
        self.room_temperature_setpoint       = self.inputs_dict['model_inputs']['Set Point Profile '+str(setpoint_profile_number)]['Room Temperature Setpoint']
        self.room_temperature_lowerband      = self.inputs_dict['model_inputs']['Set Point Profile '+str(setpoint_profile_number)]['Room Temperature Lower Band']
        self.room_temperature_upperband      = self.inputs_dict['model_inputs']['Set Point Profile '+str(setpoint_profile_number)]['Room Temperature Upper Band']
        self.room_humidity_setpoint          = self.inputs_dict['model_inputs']['Set Point Profile '+str(setpoint_profile_number)]['Room Humidity Setpoint']
        self.room_humidity_lowerband         = self.inputs_dict['model_inputs']['Set Point Profile '+str(setpoint_profile_number)]['Room Humidity Lower Band']
        self.room_humidity_upperband         = self.inputs_dict['model_inputs']['Set Point Profile '+str(setpoint_profile_number)]['Room Humidity Upper Band']
        
        self.room_w_setpoint                 = PsyCalc.w_of_tdb_rh(self.room_temperature_setpoint, self.room_humidity_setpoint, self.ambient_pressure) 
        self.room_w_lowerband                = PsyCalc.w_of_tdb_rh(self.room_temperature_setpoint, self.room_humidity_lowerband, self.ambient_pressure)
        self.room_w_upperband                = PsyCalc.w_of_tdb_rh(self.room_temperature_setpoint, self.room_humidity_upperband, self.ambient_pressure)
        
    def update_room_conditions(self, step, hour):
        
        self.results_final_supply_air_tdb.append(self.current_HVAC_tdb)
        self.results_final_supply_air_rh.append(self.current_HVAC_rh)
        self.results_ahu_mode.append(self.ahu_mode[0:-1])
        
        #update room temperature
        if self.supply_airflow!=0:
            #steady state equation
            self.room_temperature = self.current_HVAC_tdb + self.room_delta_tdb 
            
        #no heat gain then room temperature stays constant
        elif self.room_heat_gain_tdb==0:
            self.room_temperature = self.current_HVAC_tdb
            
        #no supply airflow but heat gain in room
        else:
            #transient equation, cut if room_tdb is lower than ambient temperature (physical non-nonsense)
            room_delta_tdb_transient = self.room_heat_gain_tdb / ((self.room_area*self.room_height)*self.rho_air*self.cp_air) * 3600 #K
            if self.current_HVAC_tdb + room_delta_tdb_transient<self.ambient_temperature[hour]:
                self.room_temperature = self.ambient_temperature[hour]
            else:
                self.room_temperature = self.current_HVAC_tdb + room_delta_tdb_transient
        self.results_room_tdb.append(self.room_temperature)
        
                
        #update moisture content w/ load with steamhum backwards
        if self.supply_airflow!=0:
            #steady state equation
            h_start     = PsyCalc.enthalpy_moist_air(self.room_temperature, self.current_HVAC_w)
            h_end       = h_start + self.room_heat_gain_w/(1000 * self.supply_airflow * self.rho_air)
            w_end       = PsyCalc.w_of_enthalpy_tdb(h_end, self.room_temperature)
            self.room_w = w_end
        else:
            #transient eqaution
            room_delta_h_transient = self.room_heat_gain_w / ((self.room_area*self.room_height)*self.rho_air) * 3600 / 1000 #kJ/kg
            h_start     = PsyCalc.enthalpy_moist_air(self.room_temperature, self.room_w) #kJ/kg
            h_end       = h_start + room_delta_h_transient
            w_end       = PsyCalc.w_of_enthalpy_tdb(h_end, self.room_temperature)
            self.room_w = w_end
            
        self.results_room_w.append(self.room_w)
            
        #update room humidity
        partial_pressure      = PsyCalc.partial_pressure_of_w(self.ambient_pressure, self.room_w)
        saturation_pressure   = PsyCalc.saturation_pressure(self.room_temperature)
        self.room_rh          = partial_pressure/saturation_pressure
        self.results_room_rh.append(self.room_rh)
        
        #check hours within and outside
        check_rh, check_tdb = False, False
        if self.room_rh>=self.room_humidity_lowerband and self.room_rh<=self.room_humidity_upperband:
            self.results_rh_within                += 1
            check_rh = True
        else:
            self.results_rh_outside               += 1
            
        if self.room_temperature>=self.room_temperature_lowerband and self.room_temperature<=self.room_temperature_upperband:
            self.results_tdb_within               += 1
            check_tdb = True
        else:
            self.results_tdb_outside              += 1
            
        if check_rh==True and check_tdb==True: 
            self.results_tdb_rh_within            += 1
        else:
            self.results_tdb_rh_outside           += 1
               
    def get_HVAC_systemtype(self, dictionary):
        self.possible_types = ['Mix', 'Heat', 'Cool', 'Steam', 'Spray', 'Preheat', 'Plate Heat Exchanger',
                               'Run Around Coil', 'Thermal Wheel']
        systemtype     = []
        for i in range(1,10):
            for types in self.possible_types:
                if types in str(dictionary['HVAC Module '+str(i)]):
                    systemtype.append(types)
                    break 
        fs = systemtype[0]
        if 'Heat' in fs or 'Cool' in fs or 'Steam' in fs or 'Spray' in fs or 'Preheat' in fs:
            if self.min_fresh_air_percentage!=0:
                systemtype.insert(0,'Fresh')
            else:
                systemtype.insert(0,'Mix')

        #systemtype = ['Mix', 'Cool', 'Heat']
        
        #check if unit can dehumidify
        heat_spot, cool_spot = 0, 0
        for spot, coil in enumerate(systemtype):
            if 'Cool' in coil:
                cool_spot = spot
                heat_spot = spot
            if 'Heat' in coil:
                heat_spot = spot
        check_dehum = True if heat_spot>cool_spot else False
        
        return systemtype, check_dehum 
      
    def prepare_for_internal_HVAC_tracking(self, step, hour):
        self.HVAC_internal_properties[hour][step] = {}
        
        #parameters to track within internal HVAC calculation
        parameters     = ['tdb', 'rh', 'w'] 
        if step=='Step0':
            self.HVAC_internal_properties[hour][step]['type'] = 'Return Air'
            self.current_HVAC_tdb = self.room_temperature
            self.current_HVAC_rh  = self.room_humidity
            self.current_HVAC_w   = self.room_w
            
        default_values = [self.current_HVAC_tdb, self.current_HVAC_rh, self.current_HVAC_w]    
         
        #default values from step before, depending on next step value will be overwritten
        for i, param in enumerate(parameters):
            self.HVAC_internal_properties[hour][step][param] = default_values[i] 
     
    def calculate_fan_power(self):
        
        supply_fan_power = (self.supply_airflow) * (self.supply_fan_pressure) / (self.supply_fan_efficiency)
        self.results_supply_fan_power.append(supply_fan_power/1000)
        
        return_fan_power = (self.return_airflow) * (self.return_fan_pressure) / (self.return_fan_efficiency)  
        self.results_return_fan_power.append(return_fan_power/1000)                                        
        
    def DoNothing(self, step, hour):
        
        #do nothing since supply air flow is zero
        self.HVAC_internal_properties[hour][step]['type'] = 'Zero Supply'

    def update_zero_loads(self, hour):
       
        results = [self.results_heat_load, self.results_preheat_load, self.results_cool_load, self.results_hum_load, 
                   self.results_dehum_load, self.results_airflow_fresh, self.results_post_preheat_tdb, self.results_post_preheat_rh,
                   self.results_post_HR_air_tdb, self.results_post_HR_air_rh, self.results_post_heat_tdb, self.results_post_heat_rh, 
                   self.results_post_cool_tdb, self.results_post_cool_rh, self.results_post_dehum_tdb, self.results_post_dehum_rh, 
                   self.results_post_hum_tdb, self.results_post_hum_rh, self.results_humidification_water, self.results_airflow_fresh_percentage]
        
        for result in results:
            diff = hour+1 - len(result)
            if diff!=0:
                result.append(0)
        
    def FreshAir(self, step, hour):
        
        #updating internal propertes to fresh air conditions
        self.current_HVAC_tdb = self.ambient_temperature[hour]
        self.current_HVAC_rh  = self.ambient_rh[hour]
        self.current_HVAC_w   = PsyCalc.w_of_tdb_rh(self.ambient_temperature[hour], self.ambient_rh[hour], self.ambient_pressure)
        
        #do nothing since Return Air is default for HVAC internal properties Step0
        self.HVAC_internal_properties[hour][step]['type'] = 'Fresh Air'
        self.HVAC_internal_properties[hour][step]['tdb']  = round(self.current_HVAC_tdb,2)   #deg C
        self.HVAC_internal_properties[hour][step]['rh']   = round(self.current_HVAC_rh,4)    #%
        self.HVAC_internal_properties[hour][step]['w']    = self.current_HVAC_w              #g/g
        
        self.results_airflow_fresh.append(float(self.supply_airflow))
        self.results_airflow_fresh_percentage.append(1)
          
    def Mix(self, step, hour):
           
        #check mass imbalance errors
        self.fresh_air_percentage     = self.min_fresh_air_percentage 
        V_air_fresh                   = self.supply_airflow * self.fresh_air_percentage
        required_return_air           = self.supply_airflow - V_air_fresh
        if self.return_airflow<required_return_air:
            required_return_air       = self.return_airflow
            V_air_fresh               = self.supply_airflow - required_return_air
            self.fresh_air_percentage = V_air_fresh / self.supply_airflow
                   
        #adjust fresh air percentage for acitve mixing
        if self.active_mixing=='Yes':
            V_air_fresh                   = self.supply_airflow * self.fresh_air_percentage
            min_possible_air_tdb          = ((V_air_fresh * self.ambient_temperature[hour]) + 
                                            ((self.supply_airflow - V_air_fresh) * self.current_HVAC_tdb)) / (self.supply_airflow)
            
            if ((min_possible_air_tdb<self.supply_tdb and min_possible_air_tdb>self.current_HVAC_tdb) or
                (min_possible_air_tdb>self.supply_tdb and min_possible_air_tdb<self.current_HVAC_tdb)):
                
                ratio                     = ((self.supply_tdb - self.current_HVAC_tdb) / 
                                             (self.ambient_temperature[hour] - self.current_HVAC_tdb))
                
                #fresh air percentage cant be ihigher than 1
                self.fresh_air_percentage = ratio if ratio<=1 else 1
                                
        self.results_airflow_fresh_percentage.append(self.fresh_air_percentage)
                

        #calculate new tdb after mix and update internal tdb of HVAC after this step 
        V_air_fresh           = self.supply_airflow * self.fresh_air_percentage 
        self.current_HVAC_tdb = ((V_air_fresh * self.ambient_temperature[hour]) + 
                                ((self.supply_airflow - V_air_fresh) * self.current_HVAC_tdb)) / (self.supply_airflow)
        
        #calculate new w after mix and update interal w of HVAC after this step
        self.current_HVAC_w   = ((V_air_fresh * self.ambient_w[hour]) + 
                                ((self.supply_airflow - V_air_fresh)  * self.current_HVAC_w)) / (self.supply_airflow)
               
        #update internal rh of HVAC after mix
        partial_pressure      = PsyCalc.partial_pressure_of_w(self.ambient_pressure, self.current_HVAC_w)
        saturation_pressure   = PsyCalc.saturation_pressure(self.current_HVAC_tdb)
        self.current_HVAC_rh  = partial_pressure/saturation_pressure
        
        #since tdb, rh and w change w/ Mix these values are overwritten in internal tracking of HVAC
        self.HVAC_internal_properties[hour][step]['type'] = 'Mix' 
        self.HVAC_internal_properties[hour][step]['tdb']  = round(self.current_HVAC_tdb,2)  #deg C
        self.HVAC_internal_properties[hour][step]['rh']   = round(self.current_HVAC_rh,4)   #%
        self.HVAC_internal_properties[hour][step]['w']    = self.current_HVAC_w             #g/g
        
        self.results_airflow_fresh.append(float(self.supply_airflow*self.fresh_air_percentage))
        self.results_post_HR_air_tdb.append(self.current_HVAC_tdb)
        self.results_post_HR_air_rh.append(self.current_HVAC_rh)
        
    def PHE(self, step, hour):
                
        m_cp              = self.supply_airflow * self.rho_air * self.cp_air
        max_heat_transfer = m_cp * (self.current_HVAC_tdb - self.ambient_temperature[hour])
        
        after_HR_tdb      = self.ambient_temperature[hour] + max_heat_transfer*self.heat_recovery_efficiency / m_cp 
        
        if self.current_HVAC_tdb<self.supply_tdb and after_HR_tdb>self.supply_tdb:
            self.current_HVAC_tdb = self.supply_tdb     
        elif self.current_HVAC_tdb>self.supply_tdb and after_HR_tdb<self.supply_tdb:
            self.current_HVAC_tdb = self.supply_tdb
        else:
            self.current_HVAC_tdb = after_HR_tdb
        self.current_HVAC_w = self.ambient_w[hour]
            
        #update internal rh of HVAC after mix
        partial_pressure      = PsyCalc.partial_pressure_of_w(self.ambient_pressure, self.current_HVAC_w)
        saturation_pressure   = PsyCalc.saturation_pressure(self.current_HVAC_tdb)
        self.current_HVAC_rh  = partial_pressure/saturation_pressure

        #since tdb and rh change w/ Heat thiese values are overwritten in internal tracking of HVAC
        self.HVAC_internal_properties[hour][step]['type'] = 'Plate Heat Exchanger' 
        self.HVAC_internal_properties[hour][step]['tdb']  = round(self.current_HVAC_tdb,2)  #deg C
        self.HVAC_internal_properties[hour][step]['rh']   = round(self.current_HVAC_rh,4)   #%
        self.HVAC_internal_properties[hour][step]['w']    = self.current_HVAC_w             #g/g
        
        self.results_airflow_fresh_percentage.append(1)
        self.results_post_HR_air_tdb.append(self.current_HVAC_tdb)
        self.results_post_HR_air_rh.append(self.current_HVAC_rh)
        
    def ThermalWheel(self, step, hour):
                
        m_cp              = self.supply_airflow * self.rho_air * self.cp_air
        max_heat_transfer = m_cp * (self.current_HVAC_tdb - self.ambient_temperature[hour])
        
        after_HR_tdb      = self.ambient_temperature[hour] + max_heat_transfer*self.heat_recovery_efficiency / m_cp 
        
        if self.current_HVAC_tdb<self.supply_tdb and after_HR_tdb>self.supply_tdb:
            self.current_HVAC_tdb = self.supply_tdb     
        elif self.current_HVAC_tdb>self.supply_tdb and after_HR_tdb<self.supply_tdb:
            self.current_HVAC_tdb = self.supply_tdb
        else:
            self.current_HVAC_tdb = after_HR_tdb
        self.current_HVAC_w = self.ambient_w[hour]
            
        #update internal rh of HVAC after mix
        partial_pressure      = PsyCalc.partial_pressure_of_w(self.ambient_pressure, self.current_HVAC_w)
        saturation_pressure   = PsyCalc.saturation_pressure(self.current_HVAC_tdb)
        self.current_HVAC_rh  = partial_pressure/saturation_pressure

        #since tdb and rh change w/ Heat thiese values are overwritten in internal tracking of HVAC
        self.HVAC_internal_properties[hour][step]['type'] = 'Thermal Wheel' 
        self.HVAC_internal_properties[hour][step]['tdb']  = round(self.current_HVAC_tdb,2)  #deg C
        self.HVAC_internal_properties[hour][step]['rh']   = round(self.current_HVAC_rh,4)   #%
        self.HVAC_internal_properties[hour][step]['w']    = self.current_HVAC_w             #g/g
        
        self.results_airflow_fresh_percentage.append(1)
        self.results_post_HR_air_tdb.append(self.current_HVAC_tdb)
        self.results_post_HR_air_rh.append(self.current_HVAC_rh)
        
    def RunAroundCoil(self, step, hour):
                
        m_cp              = self.supply_airflow * self.rho_air * self.cp_air
        max_heat_transfer = m_cp * (self.current_HVAC_tdb - self.ambient_temperature[hour])
        
        after_HR_tdb      = self.ambient_temperature[hour] + max_heat_transfer*self.heat_recovery_efficiency / m_cp 
        
        if self.current_HVAC_tdb<self.supply_tdb and after_HR_tdb>self.supply_tdb:
            self.current_HVAC_tdb = self.supply_tdb     
        elif self.current_HVAC_tdb>self.supply_tdb and after_HR_tdb<self.supply_tdb:
            self.current_HVAC_tdb = self.supply_tdb
        else:
            self.current_HVAC_tdb = after_HR_tdb
        self.current_HVAC_w = self.ambient_w[hour]
            
        #update internal rh of HVAC after mix
        partial_pressure      = PsyCalc.partial_pressure_of_w(self.ambient_pressure, self.current_HVAC_w)
        saturation_pressure   = PsyCalc.saturation_pressure(self.current_HVAC_tdb)
        self.current_HVAC_rh  = partial_pressure/saturation_pressure

        #since tdb and rh change w/ Heat thiese values are overwritten in internal tracking of HVAC
        self.HVAC_internal_properties[hour][step]['type'] = 'RunAroundCoil' 
        self.HVAC_internal_properties[hour][step]['tdb']  = round(self.current_HVAC_tdb,2)  #deg C
        self.HVAC_internal_properties[hour][step]['rh']   = round(self.current_HVAC_rh,4)   #%
        self.HVAC_internal_properties[hour][step]['w']    = self.current_HVAC_w             #g/g
        
        self.results_airflow_fresh_percentage.append(1)  
        self.results_post_HR_air_tdb.append(self.current_HVAC_tdb)
        self.results_post_HR_air_rh.append(self.current_HVAC_rh)
         
    def Heat(self, step, hour):
        
        #calculate heat load
        if self.current_HVAC_tdb<self.supply_tdb - (self.room_temperature_setpoint - self.room_temperature_lowerband):
            h_start    = PsyCalc.enthalpy_moist_air(self.current_HVAC_tdb, self.current_HVAC_w)
            h_end      = PsyCalc.enthalpy_moist_air(self.supply_tdb, self.current_HVAC_w)
            heat_load  = self.supply_airflow * self.rho_air * (h_end - h_start) * 1000
            
            #internal HVAC tdb equals setpoint temperature now
            self.current_HVAC_tdb = self.supply_tdb
            
            #update internal HVAC rh 
            partial_pressure      = PsyCalc.partial_pressure_of_w(self.ambient_pressure, self.current_HVAC_w)
            saturation_pressure   = PsyCalc.saturation_pressure(self.supply_tdb)
            self.current_HVAC_rh  = partial_pressure/saturation_pressure
        else:
            heat_load = 0
        
        #track current heat load in results
        self.results_heat_load.append(heat_load/1000)
        
        #since tdb and rh change w/ Heat thiese values are overwritten in internal tracking of HVAC
        self.HVAC_internal_properties[hour][step]['type'] = 'Heat' 
        self.HVAC_internal_properties[hour][step]['tdb']  = round(self.current_HVAC_tdb,2)  #deg C
        self.HVAC_internal_properties[hour][step]['rh']   = round(self.current_HVAC_rh,4)   #%
        
        self.results_post_heat_tdb.append(self.current_HVAC_tdb)
        self.results_post_heat_rh.append(self.current_HVAC_rh)       
        
        if heat_load!=0:
            self.ahu_mode += 'HEAT + '
        
    def Preheat(self, step, hour):
    
        #calculate heat load
        if self.current_HVAC_tdb<self.preheat_setpoint:
            h_start       = PsyCalc.enthalpy_moist_air(self.current_HVAC_tdb, self.current_HVAC_w)
            h_end         = PsyCalc.enthalpy_moist_air(self.preheat_setpoint, self.current_HVAC_w)
            preheat_load  = self.supply_airflow * self.rho_air * (h_end - h_start) * 1000
            
            #internal HVAC tdb equals setpoint temperature now
            self.current_HVAC_tdb = self.preheat_setpoint
            
            #update internal HVAC rh 
            partial_pressure      = PsyCalc.partial_pressure_of_w(self.ambient_pressure, self.current_HVAC_w)
            saturation_pressure   = PsyCalc.saturation_pressure(self.preheat_setpoint)
            self.current_HVAC_rh  = partial_pressure/saturation_pressure
        else:
            preheat_load = 0
        
        #track current heat load in results
        self.results_preheat_load.append(preheat_load/1000)
        
        #since tdb and rh change w/ Heat thiese values are overwritten in internal tracking of HVAC
        self.HVAC_internal_properties[hour][step]['type'] = 'Preheat' 
        self.HVAC_internal_properties[hour][step]['tdb']  = round(self.current_HVAC_tdb,2)  #deg C
        self.HVAC_internal_properties[hour][step]['rh']   = round(self.current_HVAC_rh,4)   #%
        
        self.results_post_preheat_tdb.append(self.current_HVAC_tdb)
        self.results_post_preheat_rh.append(self.current_HVAC_rh)  
        
        if preheat_load!=0:
            self.ahu_mode += 'PREHEAT + '                 
     
    def Cool(self, step, hour):
        cooling_type = 'Cool'
        #if process is driven by temperature
        if self.current_HVAC_tdb>self.supply_tdb + (self.room_temperature_upperband - self.room_temperature_setpoint):
            
            #moisture content apparatus dew point
            adp_w                 = PsyCalc.w_of_tdb_rh(self.chilled_water_temperature, 1, self.ambient_pressure)   
            
            #fit a linear line in the psychometric diagram
            y1                    = adp_w 
            y2                    = self.current_HVAC_w 
            
            #if moisture content is higher than appartus dewpoint having linear line in diagram
            if y2>y1:
                x1                    = self.chilled_water_temperature
                x2                    = self.current_HVAC_tdb
                m                     = (y2 - y1) / (x2 - x1)
                n                     = y1 - m * x1
                
                #get value from straight line intersection w/ supply_tdb
                w_end                 = m * self.supply_tdb + n
                w_end = adp_w if w_end<adp_w else w_end
                tdb_end = self.chilled_water_temperature if self.supply_tdb<self.chilled_water_temperature else self.supply_tdb
                
                #calculate cooling load
                h_start               = PsyCalc.enthalpy_moist_air(self.current_HVAC_tdb, self.current_HVAC_w)
                h_end                 = PsyCalc.enthalpy_moist_air(tdb_end, w_end)
                cool_load             = self.supply_airflow * self.rho_air * (h_start - h_end) * 1000
                
                #update interal properties after cooling
                self.current_HVAC_w   = w_end
                partial_pressure      = PsyCalc.partial_pressure_of_w(self.ambient_pressure, self.current_HVAC_w)
                saturation_pressure   = PsyCalc.saturation_pressure(tdb_end)
                self.current_HVAC_rh  = partial_pressure/saturation_pressure
                self.current_HVAC_tdb = tdb_end
                
                #if point lies outside of 100% rh function
                if self.current_HVAC_rh>1:
                    self.current_HVAC_rh = 1 
                    self.current_HVAC_w  = PsyCalc.w_of_tdb_rh(self.current_HVAC_tdb, 1, self.ambient_pressure)    
                    h_end                = PsyCalc.enthalpy_moist_air(tdb_end, self.current_HVAC_w)
                    cool_load            = self.supply_airflow * self.rho_air * (h_start - h_end) * 1000

                cooling_type          = 'Cool'
                
                
                #check if also dehumidification is required and heating coil exists downstreamwise    
                if self.current_HVAC_w>(self.supply_w + (self.room_w_upperband - self.room_w_setpoint)) and self.check_dehum==True:
                
                    #fit a linear line in the psychometric diagram with updated point 2 after cooling
                    x2                    = self.current_HVAC_tdb
                    y2                    = self.current_HVAC_w  
                    m                     = (y2 - y1) / (x2 - x1)
                    n                     = y1 - m * x1
                    
                    tdb_end               = (self.room_w_setpoint - n) / m
                                                           
                    #calculate dehum load
                    h_start               = PsyCalc.enthalpy_moist_air(self.current_HVAC_tdb, self.current_HVAC_w)
                    h_end                 = PsyCalc.enthalpy_moist_air(tdb_end, self.room_w_setpoint)
                    dehum_load            = self.supply_airflow * self.rho_air * (h_start - h_end) * 1000
                    
                    #update interal properties after dehumidifying
                    self.current_HVAC_w   = self.room_w_setpoint
                    partial_pressure      = PsyCalc.partial_pressure_of_w(self.ambient_pressure, self.current_HVAC_w)
                    saturation_pressure   = PsyCalc.saturation_pressure(tdb_end)
                    self.current_HVAC_rh  = partial_pressure/saturation_pressure
                    self.current_HVAC_tdb = tdb_end
                    
                    #if point lies outside of 100% rh function
                    if self.current_HVAC_rh>1:
                        self.current_HVAC_rh = 1 
                        self.current_HVAC_w  = PsyCalc.w_of_tdb_rh(self.current_HVAC_tdb, 1, self.ambient_pressure)    
                        h_end                = PsyCalc.enthalpy_moist_air(tdb_end, self.current_HVAC_w)
                        dehum_load           = self.supply_airflow * self.rho_air * (h_start - h_end) * 1000
                        
                    cooling_type          = 'Cool/Dehum'
                else:
                    dehum_load = 0
                
            #if moisture content is less than apparatus dew point having straight line in diagram
            else:
                tdb_end = self.chilled_water_temperature if self.supply_tdb<self.chilled_water_temperature else self.supply_tdb
                
                h_start               = PsyCalc.enthalpy_moist_air(self.current_HVAC_tdb, self.current_HVAC_w)
                h_end                 = PsyCalc.enthalpy_moist_air(tdb_end, self.current_HVAC_w)
                cool_load             = self.supply_airflow * self.rho_air * (h_start - h_end) * 1000
                dehum_load            = 0
                
                #update interal properties after cooling
                partial_pressure      = PsyCalc.partial_pressure_of_w(self.ambient_pressure, self.current_HVAC_w)
                saturation_pressure   = PsyCalc.saturation_pressure(tdb_end)
                self.current_HVAC_rh  = partial_pressure/saturation_pressure
                self.current_HVAC_tdb = tdb_end
                cooling_type          = 'Cool'
   
        #if process is driven by moisture content and heating coil exists downstreamwise 
        elif self.current_HVAC_w>(self.supply_w + (self.room_w_upperband - self.room_w_setpoint)) and self.check_dehum==True:
            
            #moisture content apparatus dew point
            adp_w                 = PsyCalc.w_of_tdb_rh(self.chilled_water_temperature, 1, self.ambient_pressure)   
            
            #fit a linear line in the psychometric diagram
            y1                    = adp_w 
            y2                    = self.current_HVAC_w 
            
            #if moisture content is higher than appartus dewpoint having linear line in diagram
            if y2>y1:
                x1                    = self.chilled_water_temperature
                x2                    = self.current_HVAC_tdb
                m                     = (y2 - y1) / (x2 - x1)
                n                     = y1 - m * x1
                
                #get value from straight line intersection w/ w_setpoint
                tdb_end               = (self.room_w_setpoint - n) / m
                           
                #calculate dehum load
                h_start               = PsyCalc.enthalpy_moist_air(self.current_HVAC_tdb, self.current_HVAC_w)
                h_end                 = PsyCalc.enthalpy_moist_air(tdb_end, self.room_w_setpoint)
                dehum_load            = self.supply_airflow * self.rho_air * (h_start - h_end) * 1000
                cool_load             = 0
               
                #internal w equals new w now
                self.current_HVAC_w   = self.room_w_setpoint
                partial_pressure      = PsyCalc.partial_pressure_of_w(self.ambient_pressure, self.current_HVAC_w)
                saturation_pressure   = PsyCalc.saturation_pressure(tdb_end)
                self.current_HVAC_rh  = partial_pressure/saturation_pressure
                self.current_HVAC_tdb = tdb_end
                
                #if point lies outside of 100% rh function
                if self.current_HVAC_rh>1:
                    self.current_HVAC_rh = 1 
                    self.current_HVAC_w  = PsyCalc.w_of_tdb_rh(self.current_HVAC_tdb, 1, self.ambient_pressure)    
                    h_end                = PsyCalc.enthalpy_moist_air(tdb_end, self.current_HVAC_w)
                    dehum_load           = self.supply_airflow * self.rho_air * (h_start - h_end) * 1000
                        
                cooling_type          = 'Dehum'
            else:
                pass
                
            
        else:
            cool_load  = 0
            dehum_load = 0
        
        #track current heat load in results
        self.results_cool_load.append(cool_load/1000)
        if self.check_dehum==True:
            self.results_dehum_load.append(dehum_load/1000)
        
        #since tdb, rh and w change w/ Cool these values are overwritten in internal tracking of HVAC
        self.HVAC_internal_properties[hour][step]['type'] = cooling_type 
        self.HVAC_internal_properties[hour][step]['tdb']  = round(self.current_HVAC_tdb,2)  #deg C
        self.HVAC_internal_properties[hour][step]['rh']   = round(self.current_HVAC_rh,4)   #%
        self.HVAC_internal_properties[hour][step]['w']    = self.current_HVAC_w             #g/g
        
        self.results_post_cool_tdb.append(self.current_HVAC_tdb)
        self.results_post_cool_rh.append(self.current_HVAC_rh)  
        
        if self.check_dehum==True and dehum_load!=0:
            self.results_post_dehum_tdb.append(self.current_HVAC_tdb)
            self.results_post_dehum_rh.append(self.current_HVAC_rh)   
            
        if cool_load!=0 and dehum_load!=0:
            self.ahu_mode += 'COOL + DEHUM + '
        if cool_load!=0 and dehum_load==0:
            self.ahu_mode += 'COOL + '
        if cool_load==0 and dehum_load!=0:
            self.ahu_mode += 'DEHUM + '
        
    def SteamHum(self, step, hour):
        
        if self.current_HVAC_w<self.supply_w - (self.room_w_setpoint - self.room_w_lowerband):
            
            #calculate humidification load
            h_start               = PsyCalc.enthalpy_moist_air(self.current_HVAC_tdb, self.current_HVAC_w)
            max_possible_w        = PsyCalc.w_of_tdb_rh(self.current_HVAC_tdb, 1, self.ambient_pressure)
            
            #if relative humidity will exceed 100%
            w_end = max_possible_w if self.room_w_setpoint>max_possible_w else self.room_w_setpoint
            
            #calculate enthalpy at end status
            h_end                 = PsyCalc.enthalpy_moist_air(self.current_HVAC_tdb, w_end)
            hum_load              = abs(self.supply_airflow * self.rho_air * (h_start - h_end) * 1000)
            
            #update iternal properties after Steam Humidification
            self.results_humidification_water.append((w_end - self.current_HVAC_w)/1000) #liter
            self.current_HVAC_w   = w_end
            partial_pressure      = PsyCalc.partial_pressure_of_w(self.ambient_pressure, self.current_HVAC_w)
            saturation_pressure   = PsyCalc.saturation_pressure(self.current_HVAC_tdb)
            self.current_HVAC_rh  = partial_pressure/saturation_pressure
        else:
            hum_load = 0
        
        #track current heat load in results
        self.results_hum_load.append(hum_load/1000)
        
        #since tdb, rh and w change w/ Cool these values are overwritten in internal tracking of HVAC
        self.HVAC_internal_properties[hour][step]['type'] = 'SteamHum' 
        self.HVAC_internal_properties[hour][step]['tdb']  = round(self.current_HVAC_tdb,2)  #deg C
        self.HVAC_internal_properties[hour][step]['rh']   = round(self.current_HVAC_rh,4)   #%
        self.HVAC_internal_properties[hour][step]['w']    = self.current_HVAC_w             #g/g  
        
        self.results_post_hum_tdb.append(self.current_HVAC_tdb)
        self.results_post_hum_rh.append(self.current_HVAC_rh)   
        
        if hum_load!=0:
            self.ahu_mode += 'HUM + '
        
    def SprayHum(self, step, hour):
        
        if self.current_HVAC_w<self.supply_w - (self.room_w_setpoint - self.room_w_lowerband):
            
            #calculate humidification load
            h_start               = PsyCalc.enthalpy_moist_air(self.current_HVAC_tdb, self.current_HVAC_w)
            twb_start             = PsyCalc.twb_of_tdb_w(self.current_HVAC_tdb, self.current_HVAC_w, self.ambient_pressure)
            max_possible_w        = PsyCalc.w_of_tdb_rh(twb_start, 1, self.ambient_pressure)
            
            #if relative humidity will exceed 100%
            w_end = max_possible_w if self.room_w_setpoint>max_possible_w else self.room_w_setpoint
            
            #calculate enthalpy at end status
            h_end                 = PsyCalc.enthalpy_of_twb_w(twb_start, w_end, self.ambient_pressure)
            hum_load              = abs(self.supply_airflow * self.rho_air * (h_start - h_end) * 1000)
            
            #update iternal properties after Steam Humidification
            self.results_humidification_water.append((w_end - self.current_HVAC_w)/1000) #liter
            self.current_HVAC_w   = w_end
            self.current_HVAC_tdb = PsyCalc.tdb_of_enthalpy_w(h_end, self.current_HVAC_w, self.ambient_pressure)
            partial_pressure      = PsyCalc.partial_pressure_of_w(self.ambient_pressure, self.current_HVAC_w)
            saturation_pressure   = PsyCalc.saturation_pressure(self.current_HVAC_tdb)
            self.current_HVAC_rh  = partial_pressure/saturation_pressure
 
        else:
            hum_load = 0
        
        #track current heat load in results
        self.results_hum_load.append(hum_load/1000)
        
        #since tdb, rh and w change w/ Cool these values are overwritten in internal tracking of HVAC
        self.HVAC_internal_properties[hour][step]['type'] = 'SprayHum' 
        self.HVAC_internal_properties[hour][step]['tdb']  = round(self.current_HVAC_tdb,2)  #deg C
        self.HVAC_internal_properties[hour][step]['rh']   = round(self.current_HVAC_rh,4)   #%
        self.HVAC_internal_properties[hour][step]['w']    = self.current_HVAC_w             #g/g 
        
        self.results_post_hum_tdb.append(self.current_HVAC_tdb)
        self.results_post_hum_rh.append(self.current_HVAC_rh) 
        
        if hum_load!=0:
            self.ahu_mode += 'HUM + '
        
 
def ForVisualization(tool):
    
    hour                 = 4906
    hour2                = 2020
    start_axis = 4500
    end_axis = 4520
    fig                  = pyplot.figure(figsize=(10,6), dpi=300)
    ax                   = fig.add_subplot(111)
    results              = tool.HVAC_internal_properties
    ambient_tdb          = tool.ambient_temperature[hour-1]
    ambient_rh           = tool.ambient_rh[hour-1]
    ambient_w            = PsyCalc.w_of_tdb_rh(ambient_tdb, ambient_rh, tool.ambient_pressure)   
    
    no, step, tdb, rh, w, amb, rh_amb, w_amb = [], [], [], [], [], [], [], []
    for i, key in enumerate(results[hour]):
        no.append(i)
        step.append(results[hour-1][key]['type'])
        tdb.append(results[hour-1][key]['tdb'])
        rh.append(results[hour-1][key]['rh'])
        w.append(results[hour-1][key]['w'])
        amb.append(ambient_tdb)
        rh_amb.append(ambient_rh)
        w_amb.append(ambient_w)
    ax.grid(True, linestyle=':')
    ax.plot(no,tdb,'r-')
    ax.plot(no,amb,'b-')
    #ax.plot(no,rh,'b-')
    ax.set_xticks(no) 
    ax.set_xticklabels(step)
    ax.tick_params(labelsize=18)
    #ax.axis([0,2,18,20])
    #ax.set_title('Temperature')
    
    fig3                  = pyplot.figure(figsize=(10,6), dpi=300)
    ax3                   = fig3.add_subplot(111)
    ax3.grid(True, linestyle=':')
    ax3.plot(no,rh,'r-')
    ax3.plot(no,rh_amb,'b-')
    ax3.set_xticks(no) 
    ax3.set_xticklabels(step)
    ax3.tick_params(labelsize=18)
    #ax3.set_title('Relative Humidity')
    
    fig4                  = pyplot.figure(figsize=(10,6), dpi=300)
    ax4                   = fig4.add_subplot(111)
    ax4.grid(True, linestyle=':')
    ax4.plot(no,w,'r-')
    ax4.plot(no,w_amb,'b-')
    ax4.set_xticks(no) 
    ax4.set_xticklabels(step)
    ax4.tick_params(labelsize=18)
    #ax4.set_title('Moisture Content')
    
    fig2 = pyplot.figure(figsize=(10,6), dpi=300)
    ax2  = fig2.add_subplot(111)
    time = numpy.linspace(0,hour-1,hour)
    #ax2.bar('Heat', tool.results_heat_load, color='red')
    #ax2.bar('Cool', tool.results_cool_load, color='blue')
    res = tool.results_final_supply_air_tdb[0:hour]
    res2 = tool.results_required_supply_tdb[0:hour]
    for i in range(len(res)):
        res[i] = res[i] - res2[i]
        
    ax2.plot(time, tool.ambient_temperature[0:hour], 'b--', label='Ambient temperature')
    ax2.plot(time, tool.results_room_tdb[0:hour], 'r', label='Room temperature')
    #ax2.plot(time, tool.results_required_supply_tdb[0:hour], 'k', label='Req Supply', linewidth=0.5)
    #ax2.plot(time, tool.results_final_supply_air_tdb[0:hour], 'r', label='Final Supply', linewidth=0.5)
    #ax2.plot(time, res, 'r', label='Supply2', linewidth=0.5)
    
    ax2.grid(True, linestyle=':')
    ax2.legend(loc='best')
    #ax2.axis([0,8600,-15,40])
    ax2.tick_params(labelsize=18)
    
    fig5 = pyplot.figure(figsize=(10,6), dpi=300)
    ax5  = fig5.add_subplot(111)
    #ax5.plot(time, tool.room_humidity[0:hour], label='Room')
    
        
    ax5.plot(time, tool.results_cool_load[0:hour], label='cooling')
    ax5.plot(time, tool.results_heat_load[0:hour], 'r', label='heating')
    ax5.plot(time, tool.results_dehum_load[0:hour], label='dehum')
    ax5.plot(time, tool.results_hum_load[0:hour], label='hum')
    ax5.plot(time, tool.results_preheat_load[0:hour], label='preheat')
    ax5.grid(True, linestyle=':')
    ax5.legend(loc='best')
    ax5.axis([0,8760,0,100])
    ax5.tick_params(labelsize=18)
   
    fig6 = pyplot.figure(figsize=(10,6), dpi=300)
    ax6  = fig6.add_subplot(111)
    #time2 = numpy.linspace(hour,hour2,hour2-hour)
    #ax6.plot(time2, tool.room_humidity[hour:hour2], label='Room')
    ax6.plot(time, tool.results_room_rh[0:hour], label='Room')
    #ax6.plot(time, tool.ambient_w[0:hour], label='Ambient')
    ax6.grid(True, linestyle=':')
    ax6.tick_params(labelsize=18)
    
    
    fig7 = pyplot.figure(figsize=(10,6), dpi=300)
    ax7  = fig7.add_subplot(111)
    #time2 = numpy.linspace(hour,hour2,hour2-hour)
    #ax6.plot(time2, tool.room_humidity[hour:hour2], label='Room')
    ax7.plot(time, tool.results_room_w[0:hour], label='Room moisture content')
    ax7.plot(time, tool.ambient_w[0:hour], label='Ambient moisture content', linewidth=0.5)
    
    ax7.grid(True, linestyle=':')
    ax7.tick_params(labelsize=18)
    ax7.legend(loc='best')
    
    fig3 = pyplot.figure(figsize=(6,8), dpi=300)
    ax3  = fig3.add_subplot(111)
    ax3.bar('Heat', numpy.sum(tool.results_heat_load)/(tool.room_area*tool.room_height), color='red')
    ax3.bar('Preheat', numpy.sum(tool.results_preheat_load)/(tool.room_area*tool.room_height), color='green')
    ax3.bar('Cool', numpy.sum(tool.results_cool_load)/(tool.room_area*tool.room_height), color='blue')
    ax3.bar('Dehum', numpy.sum(tool.results_dehum_load)/(tool.room_area*tool.room_height), color='black')
    ax3.bar('Hum', numpy.sum(tool.results_hum_load)/(tool.room_area*tool.room_height), color='purple')
    ax3.bar('Fan', (numpy.sum(tool.results_supply_fan_power)+numpy.sum(tool.results_return_fan_power))/(tool.room_area*tool.room_height), color='black')
    ax3.grid(True, linestyle=':')
    ax3.tick_params(labelsize=16)
    
    fig8 = pyplot.figure(figsize=(10,6), dpi=300)
    ax8  = fig8.add_subplot(111)
    #time2 = numpy.linspace(hour,hour2,hour2-hour)
    #ax6.plot(time2, tool.room_humidity[hour:hour2], label='Room')
    ax8.plot(time, tool.results_airflow_fresh_percentage[0:hour], label='Room')
    ax8.grid(True, linestyle=':')
    ax8.tick_params(labelsize=18)
 
   
    
    
    
    