#from lib2to3.pgen2.pgen import DFAState
import requests
import json
import pandas as pd
from tqdm import tqdm
import time


# input address for the request
#

df1=pd.read_excel("Taxes - All Locations 20221018.xlsx")
df1=df1.fillna('')
locations=df1.to_numpy().tolist()

locations_taxes_response=[]

i=0
# make the web API request
#
   
for line in tqdm(locations):
    if line[3]!='CA1000'and line[3]!='OOS':
        try:
            ca_request_string = f'https://services.maps.cdtfa.ca.gov/api/taxrate/GetRateByAddress?address={line[1]}&city={line[0]}&zip={line[2]}'
            ca_response = requests.get(ca_request_string)

            # get the individual variable values from the returned data
            #
            json_data = json.loads(ca_response.text)

            rate         = json_data['taxRateInfo'][0]['rate']
            jurisdiction = json_data['taxRateInfo'][0]['jurisdiction']
            r_city       = json_data['taxRateInfo'][0]['city']
            r_county     = json_data['taxRateInfo'][0]['county']
            r_address    = json_data['geocodeInfo']['formattedAddress']
            match_codes  = json_data['geocodeInfo']['matchCodes']
            i=+1

            #print('Acumatica tax rate:',line[3],line[4]/100, 'Web tax rate:',rate, jurisdiction,r_city,r_county,r_address,match_codes)
            if line[4]/100 != rate:
                print('Address ',r_address,'has the wrong rate:',line[4]/100,'. Replace with:',rate)
            new_row = [None] * 7
            new_row[0]= rate
            new_row[1]= line[4]/100
            new_row[2]= jurisdiction
            new_row[3]= r_city
            new_row[4]= r_county
            new_row[5]= r_address
            new_row[6]= match_codes
            locations_taxes_response.append(new_row)
            ca_response.close()
            
        except:
            pass

#print(locations_taxes_response)

df = pd.DataFrame(locations_taxes_response)
df=df.reset_index(drop=True)
df.rename(columns={0:'rate',1: 'Acumatica tax rate',2: 'jurisdiction',3: 'city',4:'county' ,5:'address' ,6: 'match_codes'}, inplace=True)


with pd.ExcelWriter("Taxes from CA.xlsx") as writer:
    df.to_excel(writer)


