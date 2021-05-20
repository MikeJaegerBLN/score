import math

t_offset = 273.15

def partial_pressure_of_w(pressure, w): # ASHRAE 2017 p. 20 eq. 20 re-arranged
    partial_pressure = (pressure * w)/(0.621945 + w)
    return partial_pressure

def saturation_pressure(tdb): # ASHRAE 2017 p. 19 eq.5 & 6

    K1	=	-5.6745359e3
    K2	=	6.3925247
    K3	=	-9.6778430e-3
    K4	=	6.2215701e-7
    K5	=	2.0747825e-9
    K6	=	-9.4840240e-13
    K7	=	4.1635019

    K8	=	-5.8002206e3
    K9	=	1.3914993
    K10	=	-4.8640239e-2
    K11	=	4.1764768e-5
    K12	=	-1.4452093e-8
    K13	=	6.5459673

    tdb = tdb + t_offset

    if tdb<=0:
        saturation_pressure = math.exp((K1/tdb) + K2 + (K3*tdb) + (K4*tdb**2) + (K5*tdb**3) + (K6*tdb**4) + (K7*math.log(tdb))) # -100 to 0 degrees celcius
    else:
        saturation_pressure = math.exp((K8/tdb) + K9 + (K10*tdb) + (K11*tdb**2) + (K12*tdb**3) + (K13*math.log(tdb))) # 0 to 200 degrees celcius

    return saturation_pressure

def w_of_tdb_rh(tdb, rh, pressure): # ASHRAE 2017 p. 20 eq. 21 with rh
    p_ws = saturation_pressure(tdb)
    w = 0.621945*rh*p_ws/(pressure - rh*p_ws)
    return w

def dew_point_temperature(pressure): #p.21 eq. 37
    pressure = pressure/1000 # for eq. to kPa
    K14 = 6.54
    K15 = 14.562
    K16 = 0.7389
    K17 = 0.09486
    K18 = 0.4569
    
    T1 = K14 + K15*math.log(pressure) + K16*math.log(pressure)**2 + K17*math.log(pressure)**3 + K18*pressure**0.1984 #above 0 deg C
    T2 = 6.09 + 12.608*math.log(pressure) + 0.4959*math.log(pressure)**2 #below 0 deg C

    if T2<0:
        T = T2
    else:
        T = T1    
    return T

def enthalpy_moist_air(tdb, w): # 2017 Fundamentals, eq.30
    h = (1.006*tdb) + (w*(2501 + 1.86*tdb))
    return h

def twb_iterator(twb_n, w_n, tdb, w, pressure, temp_step):  
    while w_n[-1]<w:
        w_twb_saturation = w_of_tdb_rh(twb_n, 1, pressure)
        part_1           = (2501 - 2.326*twb_n)*w_twb_saturation
        part_2           = 1.006*(tdb - twb_n)#
        part_3           = 2501 + 1.86*tdb - 4.186*twb_n
        w_n.append((part_1 - part_2) / part_3)
        twb_n           += temp_step  
    w_n    = [w_n[-2]]
    twb_n += -2*temp_step
    return w_n, twb_n

def twb_of_tdb_w(tdb, w, pressure): #p.21 eq. 33
    twb_n           = -200
    w_n             = [-1]
   
    steps = [10, 1, 0.1, 0.01]
    for step in steps:
        w_n, twb_n = twb_iterator(twb_n, w_n, tdb, w, pressure, temp_step=step)
            
    return twb_n

def enthalpy_of_twb_w(twb, w, pressure): #p.21 eq. 31
    h_twb_w          = 4.186 * twb
    w_twb_saturation = w_of_tdb_rh(twb, 1, pressure)
    h_twb_saturation = enthalpy_moist_air(twb, w_twb_saturation)   
    h                = h_twb_saturation - h_twb_w*(w_twb_saturation - w)
    return h

def tdb_of_enthalpy_w(h, w, pressure): #p.21 eq. 31
    tdb = (h - 2501*w) / (1.006 + w*1.86)
    return tdb

def w_of_enthalpy_tdb(h, tdb):
    w = (h - 1.006*tdb) / (2501 + 1.86*tdb)
    return w
        

