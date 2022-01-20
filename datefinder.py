import pandas as pd
import numpy as np
from datetime import timedelta
workbook = pd.read_excel('Marketing.xlsx', sheet_name = 'Raw Email Data')

hours = timedelta(hours = 24)
print(hours)
underDay = 0
openedCounter = 0
#print(variants[i])
for k in range(10000):
    sentTime = workbook["Sent"].iloc[k]
    #print(sentTime)
    openTime = workbook["Opened"].iloc[k]
    #print(openTime)
    #print(k)
    if(openTime != 0):
        openedCounter += 1
        diff = openTime - sentTime
        if(diff < hours):
            underDay +=1
            #print(diff)
            #print(openTime)
            #print(sentTime)
            #print("\n")
print(openedCounter)
print(underDay)
print(underDay/openedCounter)