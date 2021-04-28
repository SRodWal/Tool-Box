# -*- coding: utf-8 -*-
"""
Created on Tue Apr 27 19:50:49 2021

@author: serw1
"""
def monthNum(num):
    return {1 : "Jan", 2:"Feb",3:"Mar",4:"Apr",5:"May",6:"Jun",7:"Jul",8:"Aug",9:"Sep",10: "Oct",11:"Nov",12:"Dec"}[num]
 
import datetime
import pandas as pd
import numpy as np
import seaborn as sb
import matplotlib.pyplot as plt
from scipy.interpolate import interp1d

import zipfile
import scipy

from playsound import playsound # Just for fun

data = [ ["May de 2016",	1.4588,	3.6051,	3.9370,	2.3989,	2.2441], 
        ["July de 2016",	1.4588,	3.6051,	3.9370,	2.3989,	2.2441],
        ["November de 2016",	1.4911,	3.6848,	4.0315,	2.4603,	2.2961],
        ["March de 2017",	1.6217,	4.0073,	4.3197,	2.7159,	2.5487],
        ["May de 2017",	1.6214,	4.0065,	4.3114,	2.7068,	2.5401],
        ["August de 2017",	1.6137,	3.9871,	4.2884,	2.6914,	2.5269],
        ["December de 2017",	1.6210,	4.0051,	4.3140,	2.7133,	2.5472],
        ["March de 2018",	1.6427,	4.0588,	4.3140,	2.7133,	2.5472],
        ["June de 2018",	1.6776,	4.1450,	4.3140,	2.7299,	2.5472],
        ["September de 2018",	1.8889,	4.6671,	4.7928,	3.1611,	2.9417],
        ["December de 2018",	3.6430,	4.7404,	4.7373,	3.0883,	2.8755],
        ["March de 2019",	4.0274,	5.2406,	5.3266,	3.4056,	3.1710],
        ["July de 2019",	3.9728,	5.1696,	5.2364,	3.4006,	3.1601],
        ["October de 2019",	3.8678,	5.0330,	5.1195,	3.2387,	2.9952],
        ["January de 2020",	4.0088,	5.2164,	5.1945,	3.3617,	3.1463],
        ["May de 2020",	3.3926,	4.4147,	4.4366,	2.6824,	2.4985],
        ["July de 2020",	3.2679,	4.2524,	4.2868,	2.5040,	2.3356],
        ["October de 2020",	3.3096,	4.3066,	4.3388,	2.5619,	2.3924],
        ["January de 2021",	3.4281,	4.4608,	4.4883,	2.7093,	2.5355],
        ["April de 2021",	3.3657,	4.3796,	4.4082,	2.6437,	2.4725],
        ]
medpow = [246.5493,246.5493,250.7668,252.3153, 252.4918,252.4920,252.6116, 
          253.4864, 257.6204, 260.6356, 262.8404, 287.7614, 295.9551, 297.2365, 
          297.9744, 304.7970, 305.6773, 310.1292, 311.3043, 313.4589]
cols = ["Fecha","Residencial (< 50 kWh)", "Residencial (> 50 kWh)", "Baja Tensión", "Media Tensión", "Alta Tensión"]

# Crear vector de tiempo
yrs = [i for i in range(2016,2022)]
timevec = []
for y in yrs:
    for m in range(1,13):
        for dat in data:
            if (str(y) in dat[0])&(monthNum(m) in dat[0]):            
                timevec.append(datetime.datetime(y, m, 1))
            
df = pd.DataFrame(np.array(data), columns = cols)
df[df.columns[0]] = pd.DataFrame(np.array(timevec))
dftar = df.drop(df.columns[0], axis=1)
dftar.insert(4, "Potencia MT", medpow, True)
for name in dftar.columns:
        df = dftar[name]
        df = pd.to_numeric(df, errors = "coerce")
        dftar[name] = df
powdf = pd.DataFrame(medpow, columns = ["Potencia MT"])
#dftar.describe().to_excel("Tarifas.xlsx")


tar = []
for name in dftar.columns[2:5]:
    tarch = []  
    for i in range(0,len(data)-1):
        tarch.append(2*(dftar[name][i+1]-dftar[name][i])/(dftar[name][i]+dftar[name][i+1])*100)
    tar.append(tarch)

dftarch = pd.DataFrame(np.array(tar).T,columns = dftar.columns[2:5])    
fig = plt.figure(figsize = (8,6))
[sb.distplot(dftarch[i]) for i in dftarch.columns ]
fig.legend(labels = dftar.columns[2:5])
plt.show()

fits = []
for name in dftarch.columns:
    x = np.linspace(-20,20,100)
    mean, var  = scipy.stats.distributions.norm.fit(dftarch[name])
    fitted_data = scipy.stats.distributions.norm.pdf(x, mean, var)
    fits.append(fitted_data)
[plt.plot(x, y) for y in fits]
plt.show    
      
      
    