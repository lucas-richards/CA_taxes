
from pytz import VERSION
import pickle
import requests
import json
import pandas as pd
from tqdm import tqdm
import time
import AliveDataTools_v103 as AliveDataTools

def main():
    """ main """
    #API
    API=True

    # get_new_data= False
    # inventory_to_process = []
    #     #write files

    # print('Starting Script Version ' + VERSION)
    # print('Starting Inventory')

    # if get_new_data:
    #     inventory_to_process= get_inventory_data()
    #     file= open('invetory.txt','wb')
    #     pickle.dump(inventory_to_process,file)
    #     file.close
    # else:
    #     with open('invetory.txt','rb') as filehandle:
    #         inventory_to_process = pickle.load(filehandle)

    # WriteInventory('WLI', inventory_to_process)
    # WriteInventory('TAG', inventory_to_process)
    # print('Finished Inventory')
    # Read excel and convert to lists
    if API:
        locations=get_data()
        
    else:

        df1=pd.read_excel("OData100 Customer Locations 20221026.xlsx")
        df1=df1.fillna('')
        locations=df1.to_numpy().tolist()


    df2=pd.read_excel("Odata101 Sales TaxZone Rates 20221026.xlsx")
    df2=df2.fillna('')
    taxzone=df2.to_numpy().tolist()


    incorrect_rate=[]
    different_city=[]
    Good=[]
    wrong_zip=[]
    wrong_address=[]
    other_error=[]
    different_city_rate=[]

    # itirate locations and find the correct tax zone
    # make the web API request
    
    for line in tqdm(locations[0:100]):
        for row in taxzone:
            if line[2]!='CA1000'and line[2]!='OOS' and line[2]!='CA0002':
                if row[0] == line[2]:
                    try:
                        new_row = [None] * 8
                        new_row[0]=line[0]
                        new_row[1]=line[1]
                        
                        
                        # address='400 E Oak St'
                        # city='Visalia'
                        # zip='93291'
                        #ca_request_string = f'https://services.maps.cdtfa.ca.gov/api/taxrate/GetRateByAddress?address={address}&city={city}&zip={zip}'
                        
                        ca_request_string = f'https://services.maps.cdtfa.ca.gov/api/taxrate/GetRateByAddress?address={line[10]}&city={line[4]}&zip={line[8]}'
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
                        

                        new_row[2]= rate
                        new_row[3]= row[2]/100
                        new_row[4]= r_city
                        new_row[5]= line[4]
                        new_row[6]= r_address
                        new_row[7]= int(line[8])
                        
                        
                        # For multiple results

                        try:
                            if json_data['taxRateInfo'][1]['city'].lower()==line[4].lower():
                                rate2         = json_data['taxRateInfo'][1]['rate']
                                jurisdiction2 = json_data['taxRateInfo'][1]['jurisdiction']
                                r_city2       = json_data['taxRateInfo'][1]['city']
                                r_county2     = json_data['taxRateInfo'][1]['county']
                                r_address2    = json_data['geocodeInfo']['formattedAddress']
                                match_codes2  = json_data['geocodeInfo']['matchCodes']

                                new_row[2]= rate2
                                new_row[4]= r_city2
                                new_row[6]= r_address2
                                

                        except:
                            pass
                        
                                   
                        # Incorrect rate
                        if abs(row[2]/100 - new_row[2]) > 0.00001:
                            if new_row[4].lower()==line[4].lower():
                                incorrect_rate.append(new_row)
                            else:
                                different_city_rate.append(new_row)
                        elif new_row[4].lower()!=line[4].lower():
                            different_city.append(new_row)
                        else:
                            Good.append(new_row)
                    
                            
                    except:
                        # Wrong zipcode
                        if json_data['errors'][0]['message'][0:13] in ['The Zip field','The field Zip']:
                            new_row[2]= json_data['errors'][0]['message']
                            wrong_zip.append(new_row)
                            

                        # Wrong address
                        elif json_data['errors'][0]['message'] in ['The address could not be geocoded.','The Address field is required.']:
                            new_row[2]= json_data['errors'][0]['message']
                            wrong_address.append(new_row)

                        # Other errors
                        elif json_data['errors'][0]['message'] in other_errors:
                            new_row[2]= json_data['errors'][0]['message']
                            other_error.append(new_row)
                            
                        else:
                            print(json_data)
                        pass
                    
    
    # convert lists into df with column names 
    df_output1=output_df(incorrect_rate,incorrect_rate_columns)
    df_output2=output_df(wrong_address,error_columns)
    df_output3=output_df(wrong_zip,error_columns)
    df_output4=output_df(other_error,error_columns)
    df_output5=output_df(Good,incorrect_rate_columns)
    df_output6=output_df(different_city,incorrect_rate_columns)
    df_output7=output_df(different_city_rate,incorrect_rate_columns)

    # Output dataframes into excel spreadsheet
    writer = pd.ExcelWriter('Taxes analysis results.xlsx', engine='xlsxwriter')
    df_output1.to_excel(writer,sheet_name='Incorrect rate')
    df_output7.to_excel(writer,sheet_name='Incorrect rate and city')
    df_output2.to_excel(writer,sheet_name='Wrong Address')
    df_output3.to_excel(writer,sheet_name='Wrong Zipcode')
    df_output4.to_excel(writer,sheet_name='Other Errors')
    df_output5.to_excel(writer,sheet_name='Good')
    df_output6.to_excel(writer,sheet_name='Good but Different city')
    

    writer.save()


def output_df(list,columns_list):
    df_output1 = pd.DataFrame(list)
    df_output1=df_output1.reset_index(drop=True)
    df_output1.rename(columns=columns_list,inplace=True)
    
    return df_output1

def get_data():
    data_rows = []
    print ("Retrieving Acumatica Customer Location Data")
    a_data = AliveDataTools.OdataQuery(gi="Odata101")
    
    for row in a_data:
        
        c=[None]*11
        c[0] = row[0]
        c[1] = row[1]
        c[2] = row[3]
        c[3] = row[2] 
        c[4] = row[4]
        c[5] = row[5]
        c[6] = row[6]
        c[7] = row[7]
        c[8] = row[8][0:5]
        c[9] = row[9]
        c[10] = row[10]
        
        data_rows.append(c)
        #print(c)
    return data_rows


class CustomerLocation:
    def __init__(self,ac):
        self.accountID = ac
        self.locationID = ''
        self.zone_descr = ''
        self.street = ''
        self.city = ''
        self.tax_reg_no = ''
        self.zip_code = ''


incorrect_rate_columns={
    0:'AccountID',
    1: 'LocationID',
    2: 'CA rate',
    3: 'Acumatica rate',
    4: 'City',
    5: 'Acumatica City',
    6:'Address',
    7: 'Zipcode',
}

error_columns ={
    0:'AccountID',
    1: 'LocationID',
    2: 'Error',
}
    
other_errors= [
"Invalid value: ' Ox'",
"Invalid value: 'P.O. Box'",
"Invalid value: 'Po Box'",
"Invalid value: ' Box'",
"Invalid value: ' Box'",
"Invalid value: 'PO BOX'",
'A problem happened while handling your request.',
"Invalid value: 'PO Box'",
'A tax rate could not be found at the geocoded location.'

]

if __name__ == "__main__": main()

