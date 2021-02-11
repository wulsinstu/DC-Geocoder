# -*- coding: utf-8 -*-
"""
Created on Thu Feb 11 12:26:37 2021

@author: Swulsin
"""

import time as t
import requests
import os
from datetime import datetime
import pandas as pd


# Import Excel sheet with addresses. 
# Include addressLine 1 formatted like "3586 6TH STREET NW"
# and addressID to match back once you have your results
#--------------------------------------------------------------

df = pd.read_excel (r'C:\Users\swulsin\Desktop\Addresses.xlsx')  # Update this with the path for your file
#print (df) 


# Main Program ------------------------------------------------------------------------------------------
start = t.time()
length = len(df)
df3 = []

for n in range(length):
    
    address = df._get_value(n, 'addressLine1')
    addressID = df._get_value(n, 'addressID')
   
    result = requests.get("http://citizenatlas.dc.gov/newwebservices/locationverifier.asmx/findLocation2?str={}&f=json".format(address))
    j = result.json()   
    data = j['returnDataset']

    if data == None:
        df3.append({'addressID': addressID,
                        'ADDRESS_ID': None,
                        'FULLADDRESS': None,
                        'WARD': None,
                        'LATITUDE': None,
                        'LONGITUDE': None,
                        'ConfidenceLevel': None
                        })
    else:
        table = data['Table1']
        for item in table:            
            df3.append({'addressID': addressID,
                        'ADDRESS_ID': item['ADDRESS_ID'],
                        'FULLADDRESS': item['FULLADDRESS'],
                        'WARD': item['WARD'],
                        'LATITUDE': item['LATITUDE'],
                        'LONGITUDE': item['LONGITUDE'],
                        'ConfidenceLevel': item['ConfidenceLevel']
                        })
            
new_df3= pd.DataFrame(df3)
print(new_df3)

new_df3.to_excel(r'C:\Users\swulsin\Desktop\AddressesNew.xlsx') # Update this with the path where you want to save your file


end = t.time()
print(end - start)

