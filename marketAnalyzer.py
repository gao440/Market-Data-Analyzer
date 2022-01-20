# importing packages that we need   
import pandas as pd
import numpy as np
#this has bad coding practices as I completed this in a rush    

# importing the data from the csv file
#reading excel data from the file
workbook = pd.read_excel('Marketing.xlsx', sheet_name = 'Raw Email Data')

#workbook.head()

#print(workbook['Opened'].iloc[2])
countOfOpened = 0
variants = ["A1", "A2","A3", "A4","B1", "B2","B3", "B4",\
        "C1", "C2","C3", "C4", "D1", "D2","D3", "D4", \
        "E1", "E2","E3", "E4", "F1", "F2","F3", "F4", "G1", "G2","G3", "G4", \
            "H1", "H2","H3", "H4"]
#print(len(variants))
#print(variants)

#This is a map that will hold the count of opened emails for each variant
openedMap = {}
#this is a map that will hold the count of clicked emails for each variant
clickedMap = {}
#this is a map that will hold the count of sent emails for each variant
sentMap = {}

for i in range(len(variants)):
    countOfOpened = 0
    countOfClicked = 0
    varianceCount = 0
    #print(variants[i])
    for k in range(10000):
        currentVariant = variants[i]
        #print(currentVariant)
        #searches the excel file for the variant that is being searched
        if(workbook["Variant"].iloc[k] == currentVariant):
            varianceCount += 1
            #print("A1")
            if(workbook["Opened"].iloc[k] != 0):
                countOfOpened += 1
                #print(workbook["Opened"].iloc[i])
            if(workbook["Clicked"].iloc[k] != 0):
                countOfClicked += 1
    openedMap[variants[i]] = countOfOpened
    clickedMap[variants[i]] = countOfClicked
    sentMap[variants[i]] = varianceCount

openconvertMap = {}

for i in range(len(variants)):
    openconvertMap[variants[i]] = round((openedMap[variants[i]]/sentMap[variants[i]])*100,2)

conversionMap = {}
for i in range(len(variants)):
    conversionMap[variants[i]] = round((clickedMap[variants[i]]/openedMap[variants[i]])*100,2)

print("number of sent emails for each variant")
print(sentMap)
print("sent to opened ratio for each variant")
print(openconvertMap)
print("opened")
print(openedMap)
print("clicked")
print(clickedMap)
print("conversion from opened to clicked")
print(conversionMap)    

sentdf = pd.DataFrame(data=sentMap, index=[0])
sentdf = (sentdf.T)

openconvertdf = pd.DataFrame(data=openconvertMap, index=[0])
openconvertdf = (openconvertdf.T)

openeddf = pd.DataFrame(data=openedMap, index=[0])
openeddf = (openeddf.T)

clickeddf = pd.DataFrame(data=clickedMap, index=[0])
clickeddf = (clickeddf.T)

conversiondf =  pd.DataFrame(data=conversionMap, index=[0])
conversiondf = (conversiondf.T)

combined = pd.concat([sentdf, openconvertdf, openeddf, clickeddf, conversiondf], axis=1)
#print(combined)

#print (sentdf)
combined.to_excel('Exported-market-data.xlsx')
