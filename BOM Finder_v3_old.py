import bisect
import os
import time
import base64
import datetime
import pandas as pd
import logging
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
import zzz as scraper
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.keys import Keys
import requests
from bs4 import BeautifulSoup
import requests
import glob
import smtplib
import ssl
from email.message import EmailMessage
from selenium import webdriver
import sqlite3
con = sqlite3.connect(r"D:\Client Code\New_Db\test.db")
cur = con.cursor()
pd.set_option("display.max_rows", None, "display.max_columns", None)
pd.options.mode.chained_assignment = None


def search_octo_api(ref):
    a = scraper.failures_fun(ref)
    return a[0]


def read_email_from_gmail():
    creds = None

    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)

        with open('token.json', 'w') as token:
            token.write(creds.to_json())

    try:
        input_files = []
        service = build('gmail', 'v1', credentials=creds)
        results = service.users().messages().list(userId='me', labelIds=['INBOX'], q="is:unread").execute()
        messages = results.get('messages', [])

        if not messages:
            return input_files, 'Empty'
        else:
            for message in messages:
                date_time = str(datetime.datetime.now()).split(".")[0].replace(':', '-').replace(" ", "_")
                time.sleep(1)
                msg = service.users().messages().get(userId='me', id=message['id']).execute()
                for header in msg['payload']['headers']:
                    if header['name'] == 'Subject' and 'find bom for' in header['value']:
                        for header in msg['payload']['headers']:
                            if header['name'] == 'From':
                                msg_from = str(header['value']).split(' <')[1].split('>')[0]
                        if 'parts' in msg['payload']:
                            for attachment in msg['payload']['parts']:
                                if 'attachmentId' in attachment['body'] and 'findbom' in attachment['filename']:
                                    file_response = service.users().messages().attachments().get(
                                        userId='me',
                                        messageId=message['id'],
                                        id=attachment['body']['attachmentId']).execute()
                                    file_data = base64.urlsafe_b64decode(file_response.get('data').encode('UTF-8'))
                                    print('From : ' + str(msg_from))
                                    print('Subject : ' + str(header['value']))
                                    print(f'Filename: {msg_from}_{date_time}_{attachment["filename"]}')
                                    input_files.append(f'{msg_from}_{date_time}_{attachment["filename"]}')
                                    with open(f'./Received/{msg_from}_{date_time}_{attachment["filename"]}', 'wb') as fp:
                                    # fp = open(f'./Received/{msg_from}_{date_time}_{attachment["filename"]}', 'wb')
                                        fp.write(file_data)
                                        fp.close()

                                if 'attachmentId' in attachment['body'] and 'direct_insertion' in attachment['filename']:
                                    file_response = service.users().messages().attachments().get(
                                        userId='me',
                                        messageId=message['id'],
                                        id=attachment['body']['attachmentId']).execute()
                                    file_data = base64.urlsafe_b64decode(file_response.get('data').encode('UTF-8'))
                                    print('From : ' + str(msg_from))
                                    print('Subject : ' + str(header['value']))
                                    print(f'Filename: {msg_from}_{date_time}_{attachment["filename"]}')
                                    input_files.append(f'{msg_from}_{date_time}_{attachment["filename"]}')
                                    fp = open(f'./Received/{msg_from}_{date_time}_{attachment["filename"]}', 'wb')
                                    fp.write(file_data)
                                    fp.close()
                                else:
                                    pass

                service.users().messages().modify(userId='me', id=message['id'], body={'removeLabelIds': ['UNREAD']}).execute()

        if len(input_files) > 0:
            return input_files, 'Success'
        else:
            return input_files, 'Empty'

    except HttpError as error:
        print(f'An error occurred: {error}')
        return input_files, 'Failed'


def send_results(user_key):
    print(f'Replying to {user_key.split(".com_")[0]}.com')
    output_file = glob.glob(f'Sent/{user_key}_bom_ht.xlsx')
    tebo_file = glob.glob(f'Sent/{user_key}_bom_tebo.xlsx')
    database_file = glob.glob(f'Sent/{user_key}_bom_database.xlsx')
    failed_file = glob.glob(f'Sent/{user_key}_BOM.xlsx')
    output_parts = pd.read_excel(output_file[0], sheet_name=0)
    parts = output_parts['Manufacturer PART NO.'].tolist()

    datasheets = []
    # for part in parts:
    #     try:
    #         datasheets.append(f'./datasheets/{part.replace("/", "_")}_datasheet.pdf')
    #     except:
    #         pass

    port = 465
    smtp_server = "smtp.gmail.com"
    sender_email = 'Sptc.bom@gmail.com'
    password = 'pbbwyshvptlojsjc'

    msg = EmailMessage()
    msg.set_content('Kindly find attached the bom output and the failed parts to be searched manually.')
    msg['Subject'] = f'BOM Finder {user_key.split(".com_")[1]}'
    msg['From'] = sender_email
    msg['To'] = f'{user_key.split(".com_")[0]}.com'

    with open(output_file[0], 'rb') as f:
        file_data = f.read()
        msg.add_attachment(file_data, maintype="application", subtype="xlsx", filename='bom_ht.xlsx')

    with open(tebo_file[0], 'rb') as f:
        file_data = f.read()
        msg.add_attachment(file_data, maintype="application", subtype="xlsx", filename='bom_tebo.xlsx')

    if failed_file:
        with open(failed_file[0], 'rb') as f:
            file_data = f.read()
            msg.add_attachment(file_data, maintype="application", subtype="xlsx", filename='failed_bom.xlsx')

    with open(database_file[0], 'rb') as f:
        file_data = f.read()
        msg.add_attachment(file_data, maintype="application", subtype="xlsx", filename='bom_database.xlsx')

    

    # for datasheet in datasheets:
    #     try:
    #         with open(datasheet, 'rb') as f:
    #             file_data = f.read()
    #             msg.add_attachment(file_data, maintype="application", subtype="xlsx", filename=f'{datasheet.split("datasheets/")[1]}')
    #     except:
    #         pass
    context = ssl.create_default_context()

    with smtplib.SMTP_SSL(smtp_server, port, context=context) as server:
        server.login(sender_email, password)
        server.send_message(msg)


def normalize_mastersheet():
    master_sheet = pd.read_excel('masterdata.xlsx', sheet_name='data')
    master_sheet = master_sheet[master_sheet['HT'].notna()]
    master_sheet.to_excel('masterdata.xlsx', sheet_name='data', index=False)


def generate_tebo_frame(normal_frame):
    tebo_frame = normal_frame[['VALUE', 'Manufacturer PART NO.', 'HT','REF.']]
    Height =[]
    try:
        for i in tebo_frame['HT']:
            Height.append(calculate_tebo_number(i))
        # tebo_frame['Height'] = tebo_frame.apply(lambda row: calculate_tebo_number(row.HT), axis=1)
    except Exception as e:
        print('Error while calculating tebo: ', e)
        Height.append(None)
        pass
    tebo_frame['Height'] = Height
    print(Height)
    tebo_frame = tebo_frame.drop('HT', axis=1)
    tebo_frame.rename(columns={'Height': 'HT'}, inplace=True)
    return tebo_frame

def calculate_tebo_number(number):
    try:
        if 'mm' in str(number):
            number = str(number).replace('mm','')
        if ' ' in str(number):
            number = str(number).replace(' ','')
        number = str(number).replace('in','')
        if str(number).replace('.','',1).isdigit():
            number = float(number)
            if number - int(number) != 0:
                
                dec = round(number - int(number), 6)
                if dec > 0 and dec <= 0.1:
                    return (int(number) + 0.5)
                elif dec > 0.1 and dec <= 0.6:
                    return (int(number) + 1)
                elif dec > 0.6 and dec <= 0.99:
                    return (int(number) + 1.5)
            else:
                return (round(number) + 0.5)
        else:
            print('Error while calculating tebo: Height isn\'t a number', number)
            return number
    except Exception as e:
        print('Error while calculating tebo: ', e)
        return number
colum = ['VALUE', 'Manufacturer PART NO.', 'HT', 'SMD/Through', 'Plunger Type',
        'Dia', 'Spring Dia', 'Material', 'Pitch', '3d drawing', 'datasheet', 'Source'
        ]
# MrfpPartNo	datasheet	REF	Mounting	VALUE	HT	Dia	Spring_Dia	Material	Pitch	Three_d_drawing	Size_of_componen

def search_result(MrfpPartNo,data_list):
    res = cur.execute("SELECT * FROM OtherProducts WHERE MrfpPartNo='{}'".format(MrfpPartNo))
    # print(res.fetchone())
    # first_db=res.fetchone() is None
    first_db = True
    row = res.fetchone()
    print(row)  
    try:  
        if row:
            first_db = False
            temp=[row[5],row[1],row[6],row[4],'',row[7],row[8],row[9],row[10],row[11],row[2],row[3]]
            if temp not in data_list:
                data_list.append(temp)
    except:
        pass
    res = cur.execute("SELECT * FROM ThroughProducts WHERE MrfpPartNo='{}'".format(MrfpPartNo))
    # print(res.fetchone()) 
    # second_db=res.fetchone() is None
    
    row = res.fetchone()   
    second_db=True
    print(row)
    try:
        if row:
            second_db =False
            temp=[row[5],row[1],row[6],row[4],'',row[7],row[8],row[9],row[10],row[11],row[2],row[3]]
            if temp not in data_list:
                data_list.append(temp)
    except:
        pass
    res = cur.execute("SELECT * FROM SurfaceProducts WHERE MrfpPartNo='{}'".format(MrfpPartNo))
    # print(res.fetchone())
    # third_db=res.fetchone() is None
    third_db=True
    row = res.fetchone()    
    print(row)
    try:
        if row:
            third_db=False
            temp=[row[5],row[1],row[6],row[4],'',row[7],row[8],row[9],row[10],row[11],row[2],row[3]]
            if temp not in data_list:
                data_list.append(temp)
    except:
        pass
        
    print(data_list)
    print(first_db,second_db,third_db)
    if first_db==True and second_db== True and third_db==False:
        df=pd.read_excel('masterdata.xlsx', sheet_name='data')

        return False
    elif first_db==False and second_db== True and third_db==True:
        return False
    elif first_db==True and second_db== False and third_db==True:
        return False
    return True


def enrgtech_scraper(code):
    enrgtechData={}
    enrgtechData[code] = {
            'Enrgtech - Mounting': 'N/A',
            'Enrgtech - Package / Case': 'N/A',
            'Enrgtech - Height': 'N/A'
        }
    if str(code) != 'nan':
        try:
            options = Options()
            options.add_argument(
                "user-agent=Mozilla/5.0 (Windows NT 6.1; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/84.0.4147.125 Safari/537.36")
            options.add_argument('--headless')
            # minimize the window
            options.add_argument("--window-size=1920,1080")
            
            options.add_experimental_option('excludeSwitches', ['enable-logging'])
            driverService = Service('./chromedriver.exe')
            driver = webdriver.Chrome(options=options)
            driver.get(f'https://www.enrgtech.co.uk/search?term={code}')
            time.sleep(5)
            try:
                driver.find_element(By.ID, 'bot-verify')
                print('Captcha appeared in Enrgtech')
            except:
                print('Enrgtech', code)
                response = driver.page_source
                # response = requests.get(f'https://www.enrgtech.co.uk/search?term={code}', headers={'user-agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/99.0.4844.51 Safari/537.36'})
                # if response.status_code == 200:
                soup = BeautifulSoup(response, features="html.parser")
                try:
                    for row_ in soup.find('div', {'class': 'row attributes-row'}).find_all('div'):
                        try:
                            if 'attribute-name' in " ".join(row_['class']) and str(row_.text).strip().lower() == 'mounting type:':
                                enrgtechData[code]['Enrgtech - Mounting'] = str(row_.find_next_sibling('div').text).strip()
                            elif 'attribute-name' in " ".join(row_['class']) and (str(row_.text).strip().lower() == 'height:' or str(row_.text).strip().lower() == 'height - seated (max):'):
                                enrgtechData[code]['Enrgtech - Height'] = str(row_.find_next_sibling('div').text).strip()
                            elif 'attribute-name' in " ".join(row_['class']) and str(row_.text).strip().lower() == 'package / case:':
                                enrgtechData[code]['Enrgtech - Package / Case'] = str(row_.find_next_sibling('div').text).strip()
                        except:
                            continue
                except:
                    pass
        except:
            pass
    return enrgtechData

def mouser_scraper(mrfp_id):
    headers={ 'authority': 'www.mouser.in' ,
   'accept': 'application/json, text/javascript, */*; q=0.01' ,
   'accept-language': 'en-US,en;q=0.9,hi;q=0.8' ,
    'referer': 'https://www.mouser.in/c/?q=ATM15' ,
   'sec-ch-ua': '"Chromium";v="112", "Google Chrome";v="112", "Not:A-Brand";v="99"' ,
   'sec-ch-ua-mobile': '?0' ,
   'sec-ch-ua-platform': '"Windows"' ,
   'sec-fetch-dest': 'empty' ,
   'sec-fetch-mode': 'cors' ,
   'sec-fetch-site': 'same-origin' ,
   'user-agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/112.0.0.0 Safari/537.36' ,
   'x-requested-with': 'XMLHttpRequest' }
    row_list = {
        "Mouting Style": "",
        "Height": "",
        "Mrfp": mrfp_id,
        'Datasheet': ''
    }
    try:
        url= 'https://www.mouser.in/c/?q='+mrfp_id
        respose = requests.get(url, headers=headers)
        print(mrfp_id)
        print('status code',respose.status_code)
        soup = BeautifulSoup(respose.text, 'html.parser')
        table = soup.find('table', {'id': 'SearchResultsGrid_grid'})
        if table is None:
                    try:
                        datasheets = soup.find('div', {'class': 'pdp-product-datasheets'})
                        a = datasheets.find('a')
                        row_list['Datasheet'] = a['href']
                    except:
                        pass
                    try:
                        table = soup.find('table', {'class': 'specs-table'})
                        tr= table.find_all('tr')
                        for i in tr:
                            if 'Mounting Style:' in i.text:
                                row_list["Mouting Style"] = i.text.split(':')[-1].strip()
                            if 'Height' in i.text:
                                row_list["Height"] = i.text.split(':')[-1].strip()
                    except:
                        pass
        else:
            try:
                tbody = table.find('tbody')
                tr = tbody.find_all('tr')
                
                for i in tr[0:1]:
                    td = i.find_all('td')
                    co=0
                    for j in td:
                        co+=1
                        if co==2:
                            button = j.find('button')
                            url = button['onclick'].split('"')[-2]
                            respose = requests.get('https://www.mouser.in'+url, headers=headers)
                            soup = BeautifulSoup(respose.text, 'html.parser')
                            try:
                                datasheets = soup.find('div', {'class': 'pdp-product-datasheets'})
                                a = datasheets.find('a')
                                row_list['Datasheet'] = a['href']
                            except:
                                pass
                            try:
                                table = soup.find('table', {'class': 'specs-table'})
                                tr= table.find_all('tr')
                                for i in tr:
                                    if 'Mounting Style:' in i.text:
                                        row_list["Mouting Style"] = i.text.split(':')[-1].strip()
                                    if 'Height' in i.text:
                                        row_list["Height"] = i.text.split(':')[-1].strip()
                                
                            except:
                                pass
            except:
                pass
    except Exception as e:
        print('error',e)
        return row_list    
    return row_list



def inster_record(df):
    # df = pd.read_excel(path, sheet_name='input')
    for mrfp in df['Manufacturer PART NO.']:
        try:
            # get the data from database for the mrfp 
            cur.execute(f"SELECT * FROM ThroughProducts WHERE MrfpPartNo = '{mrfp}'")
            data = cur.fetchall()
            if len(data) == 0:
                cur.execute('Insert into ThroughProducts (MrfpPartNo, datasheet, REF, Mounting, VALUE, HT, Dia, Spring_Dia, Material, Pitch, Three_d_drawing) values (?,?,?,?,?,?,?  ,?,?,?,?)', (mrfp, df[df['Manufacturer PART NO.'] == mrfp]['datasheet'].values[0] , None , df[df['Manufacturer PART NO.'] == mrfp]['SMD/Through'].values[0] , df[df['Manufacturer PART NO.'] == mrfp]['VALUE'].values[0] , df[df['Manufacturer PART NO.'] == mrfp]['HT'].values[0] , df[df['Manufacturer PART NO.'] == mrfp]['Dia'].values[0] , df[df['Manufacturer PART NO.'] == mrfp]['Spring Dia'].values[0] , df[df['Manufacturer PART NO.'] == mrfp]['Material'].values[0] , df[df['Manufacturer PART NO.'] == mrfp]['Pitch'].values[0] , df[df['Manufacturer PART NO.'] == mrfp]['3d drawing'].values[0] ))
                con.commit()
        except:
            pass

import inspect

class CustomException(Exception):
    def __init__(self, message):
        self.message = message
        self.line_number = inspect.currentframe().f_back.f_lineno

    def __str__(self):
        return f'{self.message} (Line {self.line_number})'

def compare_with_master(mail_boms):
    
    for new_bom in mail_boms:
        try:
            if 'findbom.xls' in str(new_bom) or 'findbom.xlsx' in str(new_bom):
                print(f'Finding parts info for: {new_bom}')
                new_name = str(new_bom).split('_findbom.xlsx')[0]
                normalize_mastersheet()
                #converting all masterpart parts nr into list
                master_parts = pd.read_excel('masterdata.xlsx', sheet_name='data', usecols=['Manufacturer PART NO.'])['Manufacturer PART NO.'].tolist()
                master_sheet = pd.read_excel('masterdata.xlsx', sheet_name='data')
                try:
                    #getting find_bom excel file only parts numbers
                    new_bom_frame = pd.read_excel(f'{str(os.getcwd())}\\Received\\{new_bom}', sheet_name='input')
                    new_bom_frame = new_bom_frame.dropna()
                    new_bom_parts_only = new_bom_frame[['Manufacturer PART NO.']]
                except Exception as e:
                    print(str(e))
                    log.error(f'Please make sure that the file {new_bom} has a sheet name named "input"')
                    status = 'Fail'
                    new_email_name = new_name.split('@')[0].replace('@','').replace(":",' ').replace('-',' ').replace("'",'')
                    current_time = datetime.datetime.now()
                    email_message=str(f'Please make sure that the file {new_email_name} has a sheet name named is input')
                    table1 =f"INSERT INTO mail_status (sendermail,status, message, created_date) VALUES ('{new_email_name}','{status}','{email_message}','{current_time}');"
                    cur.execute(table1)
                    con.commit()
                    continue

                print(f'Comparing {new_bom} with masterdata sheet')
                #converting newbom parts codes into a list
                new_bom_parts = new_bom_parts_only['Manufacturer PART NO.'].tolist()
                difference=[]
                data_list=[]
                for part in new_bom_parts:
                    # print(part,search_result(part,data_list))
                    if search_result(part,data_list):
                        difference.append(part)
                
                # difference = [part for part in new_bom_parts if part not in master_parts]
                print('Difference: ',len(difference))
                print(data_list)
                database_dataframe =  pd.DataFrame(data_list,
                                              columns =colum)
              
                #now we have difference list  which is a list  that conatains all parts number that are not in masterdata
                new_difference=[]
                for val in difference:
                    try:
                        mouser_data = mouser_scraper(val)
                        if mouser_data['Mouting Style'] == '' and mouser_data['Height'] == '' and mouser_data['Datasheet'] == '':
                            print('No data found for', val)
                            new_difference.append(val)
                            continue
                        else:
                            database_dataframe.loc[len(database_dataframe.index)] = ['',mouser_data['Mrfp'],mouser_data['Height'],mouser_data['Mouting Style'],'','','','','','',mouser_data['Datasheet'],'']
                    except Exception as e:
                        print('errorss',str(e))
                        new_difference.append(val)
                 

                visited =[ ]
                new_df = pd.DataFrame(columns = database_dataframe.columns)
                for ids,i in enumerate(database_dataframe['Manufacturer PART NO.']):
                    try:
                        if str(i).strip() not in visited:
                            visited.append( str(i).strip())
                            new_df = new_df.append(database_dataframe.iloc[ids])
                    except Exception as e:
                        print(str(e))
                        continue
                print('new df')
                for ids , i in enumerate(new_df['Manufacturer PART NO.']):
                    for new_ids , j in enumerate(new_bom_frame['Manufacturer PART NO.']):
                        try:
                            if str(i).strip() == str(j).strip():
                                new_df.loc[ids,'VALUE'] = new_bom_frame.loc[new_ids,'VALUE']
                                new_df.loc[ids,'REF.'] = new_bom_frame.loc[new_ids,'REF.']
                                break
                        except Exception as e:
                            print(str(e))
                            continue
                


                difference = new_difference
                print('New Difference: ',len(difference))
                if difference:
                    octopartFrame, arrowFrame, mouserFrame, digikeyFrame, enrgtechFrame, failures, file_paths = scraper.main(difference)
                    #after scraping all sites data is passed through all scraper platforms as df and merged with one and others
                    octopart_merged = pd.merge(new_bom_frame, octopartFrame, left_on='Manufacturer PART NO.', right_index=True, how='left')
                    # mouser_merged = pd.merge(octopart_merged, mouserFrame, left_on='Manufacturer PART NO.', right_index=True,  how='left')
                    # digikey_merged = pd.merge(mouser_merged, digikeyFrame, left_on='Manufacturer PART NO.', right_index=True, how='left')
                    #after merging all scraping data into one df, we save this data  into new excel file
                    final = pd.merge(octopart_merged, mouserFrame, left_on='Manufacturer PART NO.', right_index=True, how='left')
                    final.to_excel(f'Sent/{new_name}_output.xlsx', index=False)
                    #Now we get this new created excel file with all scraping data in it
                    data = pd.read_excel(f'Sent/{new_name}_output.xlsx', sheet_name='Sheet1')
                    #all data of dataframe cells is converted to string

                    # drop a row if Manufacturer PART NO. is not in difference list
                    data = data[data['Manufacturer PART NO.'].isin(difference)]

                   
                    data = data.astype(str)
                    data['HT'] = 'N/A'
                    data['SMD/Through'] = 'N/A'
                    data['3d drawing']  = 'N/A'
                    data['datasheet']  = 'N/A'
                    #iterating through all part number that are saved in new excel file, those rows contains all scraping iformations
                    for index, row in data.iterrows():
                        #here all height for each scraping site is exctracted
                        inputHeights = [data.loc[index, 'Octopart - Height'],
                                        data.loc[index, 'Mouser - Height'],]
                        #checking if all height are empty
                        if not all(x == 'nan' for x in inputHeights):
                            heights = []
                            #getting all height values
                            octoHieght = str(data.loc[index, 'Octopart - Height']).split(' ')
                            mouserHieght = str(data.loc[index, 'Mouser - Height']).split(' ')
                            #all height values are in list , now happens conversion from µm to mm
                            try:
                                if octoHieght[1] == 'µm':
                                    heights.append(float(octoHieght[0]) / 1000)
                                elif octoHieght[1] == 'mm':
                                    heights.append(float(octoHieght[0]))
                            except:
                                pass
                            try:
                                if mouserHieght[1] == 'µm':
                                    heights.append(float(mouserHieght[0]) / 1000)
                                elif mouserHieght[1] == 'mm' or mouserHieght[1] == 'm':
                                    heights.append(float(mouserHieght[0]))
                            except:
                                pass
                            #here we get the  highest  value
                            if heights:
                                data.loc[index, 'HT'] = heights[0]
                        mountTypes = [data.loc[index, 'Octopart - Mount'],
                                    data.loc[index, 'Octopart - Termination'], data.loc[index, 'Mouser - Termination']]
                        #Checking if all mount  types are empty
                        if not all(x == 'nan' for x in mountTypes):
                            if 'SMD/SMT' in mountTypes or 'Surface Mount' in mountTypes:
                                data.loc[index, 'SMD/Through'] = 'SMD'
                            else:
                                data.loc[index, 'SMD/Through'] = 'Through Hole'

                        if str(data.loc[index, 'Manufacturer PART NO.']) in file_paths['3d']:
                            data.loc[index, '3d drawing'] =  file_paths['3d'][str(data.loc[index, 'Manufacturer PART NO.'])]
                        if str(data.loc[index, 'Manufacturer PART NO.']) in file_paths['datasheets']:
                            data.loc[index, 'datasheet'] = file_paths['datasheets'][str(data.loc[index, 'Manufacturer PART NO.'])]
                    #now we have  two new values, the hieight as HT and  mount type as 'SMD/Through'
                    #  saved  into an new df  as output version
                    output_version = data[['VALUE','Manufacturer PART NO.', 'HT', 'SMD/Through']]
                    #with set_index() method each part nr is an index, from which can acces the data through pd
                    output_version.set_index('Manufacturer PART NO.', inplace=True)
                    #now we have old mastersheet combined with new mastersheet and again index is part number
                    output_version.update(master_sheet.set_index('Manufacturer PART NO.')) 
                    output_version.reset_index(inplace=True)
                    #Merging two dataframes old mastersheet and new mastersheet
                    new_master = pd.concat([master_sheet, output_version], ignore_index=True)
                    new_master.drop_duplicates(subset=['Manufacturer PART NO.'], inplace=True)
                    #writing paths for each part
                    for index, row in new_master.iterrows():
                        directory = os.getcwd().replace("\\", "/")
                        if str(new_master.loc[index, 'Manufacturer PART NO.']) in file_paths['3d']:
                            new_master.loc[index, '3d drawing'] =  file_paths['3d'][str(new_master.loc[index, 'Manufacturer PART NO.'])]
                        if str(new_master.loc[index, 'Manufacturer PART NO.']) in file_paths['datasheets']:
                            new_master.loc[index, 'datasheet'] = file_paths['datasheets'][str(new_master.loc[index, 'Manufacturer PART NO.'])]
                    print('Output Version: ',output_version)
                    #trying failured part
                    failures = output_version[output_version['HT'] == 'N/A']
                    print('Failures:', len(failures))
                    for index, row in failures.iterrows():
                        print(f'Searcing octopart API for part #{failures.loc[index, "Manufacturer PART NO."]}')
                        # retries = search_octo_api(failures.loc[index, 'Manufacturer PART NO.'])
                     
                        retries = ['N/A','N/A']
                        if retries:
                            height = retries[0]
                            mount = retries[1]
                        else:
                            height = 'N/A'
                            mount = 'N/A'
                        
                        if height=='N/A' and mount=='N/A':
                            mrfp = failures.loc[index, "Manufacturer PART NO."]
                            enrgtech_data=enrgtech_scraper(mrfp)
                            if enrgtech_data: 
                                if 'Enrgtech - Height' in enrgtech_data[mrfp]:
                                    if 'µm' in enrgtech_data[mrfp]['Enrgtech - Height']:
                                        height = float(enrgtech_data[mrfp]['Enrgtech - Height'].replace('µm','')) / 1000
                                        failures.loc[failures['Manufacturer PART NO.'] == mrfp, 'HT'] = height
                                    else:
                                        failures.loc[failures['Manufacturer PART NO.'] == mrfp, 'HT'] = enrgtech_data[mrfp]['Enrgtech - Height']
                                if 'Enrgtech - Mounting' in enrgtech_data[mrfp]:
                                    
                                    failures.loc[failures['Manufacturer PART NO.'] == mrfp, 'SMD/Through'] = enrgtech_data[mrfp]['Enrgtech - Mounting']
 
                        else:
                            if height:
                                if 'µm' in height:
                                    height = float(height.replace('µm','')) / 1000
                                    failures.loc[index, 'HT'] = height
                                else:
                                    failures.loc[index, 'HT'] = height
                            if mount:
                                if 'SMD/SMT' in mount or 'Surface Mount' in mount:
                                    failures.loc[index, 'SMD/Through'] = 'SMD'
                                elif 'N/A' in mount:
                                    failures.loc[index, 'SMD/Through'] = 'N/A'
                                else:
                                    failures.loc[index, 'SMD/Through'] = 'Through Hole'
                        
                        print('Height:',height,'Mount',mount)
                    #removing N/A Values from master and output version
                    output_version.drop(output_version[output_version['HT'] == 'N/A'].index, inplace=True)
                    print()
                    new_master.drop(new_master[new_master['HT'] == 'N/A'].index, inplace=True)
                    
                    # output_after_api = failures[(failures['HT'] != 'N/A') & (failures['SMD/Through'] != 'N/A')] put or instead of and
                    output_after_api = failures[(failures['HT'] != 'N/A') | (failures['SMD/Through'] != 'N/A')]
                    output_version = pd.concat([output_version, output_after_api], ignore_index=True)
                    new_master = pd.concat([new_master, output_after_api], ignore_index=True)
                    print('Output Version after API:', failures['HT'], failures['SMD/Through'])
                    failures_after_api = failures[(failures['HT'] == 'N/A') & (failures['SMD/Through'] == 'N/A')]
                    print('Failures after API:', len(failures_after_api))
                    
 
                    for index, row in new_master.iterrows():
                        directory = os.getcwd().replace("\\", "/")
                        if str(new_master.loc[index, 'Manufacturer PART NO.']) in file_paths['3d']:
                            new_master.loc[index, '3d drawing'] =  file_paths['3d'][str(new_master.loc[index, 'Manufacturer PART NO.'])]
                        if str(new_master.loc[index, 'Manufacturer PART NO.']) in file_paths['datasheets']:
                            new_master.loc[index, 'datasheet'] = file_paths['datasheets'][str(new_master.loc[index, 'Manufacturer PART NO.'])]

                    # new_master.to_excel('masterdata.xlsx', sheet_name='data', index=False)
                    print(output_version)
                   
                    # new_df = new_df.dropna(subset=['HT'])
                    # merging new_df and output_version
                    output_version = pd.concat([output_version, new_df], ignore_index=True, sort=False)
                    
                    tebo_frame = generate_tebo_frame(output_version)
                    print("<<<<<<<<<",tebo_frame)
                    print('++++++++')
                    print(new_bom_frame)
                    new_bom_frame=pd.read_excel(f'{str(os.getcwd())}\\Received\\{new_bom}', sheet_name='input')
                    # database_dataframe['REF.']= None
                    # tebo_frame = pd.concat([tebo_frame, database_dataframe], ignore_index=True, sort=False)
                    output_version = new_master[new_master['Manufacturer PART NO.'].isin(output_version['Manufacturer PART NO.'].to_list())]  

                    # print(database_dataframe.columns)
                    # for mefpn in output_version['Manufacturer PART NO.']:
                    #     if  mefpn not in database_dataframe['Manufacturer PART NO.'].to_list():
                    #         # add this row in database dataframe
                    #         database_dataframe = database_dataframe.append(output_version[output_version['Manufacturer PART NO.'] == mefpn], ignore_index=True)
                    # print(database_dataframe.columns)
                    # print('_________________________')
                    # for mrfpn in tebo_frame['Manufacturer PART NO.']:
                    #     if  mrfpn not in database_dataframe['Manufacturer PART NO.'].to_list():
                    #         # add this row in database dataframe
                    #         database_dataframe = database_dataframe.append(tebo_frame[tebo_frame['Manufacturer PART NO.'] == mrfpn], ignore_index=True)
                    # print(database_dataframe.columns)
                    # print('_________________________')
                    # for mrfpn,ref, val in zip(database_dataframe['Manufacturer PART NO.'], database_dataframe['REF.'], database_dataframe['VALUE']):
                    #     if str(ref)=='nan' or ref==None:
                    #             new_bom_frame_ref = new_bom_frame[new_bom_frame['Manufacturer PART NO.'] == mrfpn]['REF.'].values[0]
                    #             database_dataframe.loc[database_dataframe['Manufacturer PART NO.'] == mrfpn, 'REF.'] = new_bom_frame_ref
                        
                    #     if str(val)=='nan' or val==None:
                    #             new_bom_frame_value = new_bom_frame[new_bom_frame['Manufacturer PART NO.'] == mrfpn]['VALUE'].values[0]
                            
                    #             database_dataframe.loc[database_dataframe['Manufacturer PART NO.'] == mrfpn, 'VALUE'] = new_bom_frame_value

           

                    tebo_frame.to_excel(f'Sent/{new_name}_bom_tebo.xlsx', sheet_name='input', index=False)

                    output_version.to_excel(f'Sent/{new_name}_bom_ht.xlsx', sheet_name='input', index=False)
                    failures_after_api.to_excel(f'Sent/{new_name}_BOM.xlsx', sheet_name='input', index=False)
                    database_dataframe.to_excel(f'Sent/{new_name}_bom_database.xlsx', sheet_name='input', index=False)
                    
                    inster_record(output_version)

                    
                    # # merge all the dataframes(database_dataframe, output_version, tebo_frame) and create a new dataframe all have one column in common that is Manufacturer PART NO.
                    # new_df = pd.merge(database_dataframe, output_version, on='Manufacturer PART NO.', how='outer')
                    # print(new_df.columns)
                    # # convert the new_df to excel file
                    # new_df.to_excel('bom_ht.xlsx', sheet_name='input', index=False)
                    # new_df = pd.merge(new_df, tebo_frame, on='Manufacturer PART NO.', how='outer')
                    # new_df.to_excel('bom_tebo.xlsx', sheet_name='input', index=False)
                    # print(new_df.columns)
                    
                    
                else:
                    print('All parts exist in the masterdata sheet')
                    data = master_sheet[master_sheet['Manufacturer PART NO.'].isin(new_bom_parts)]
                    data.to_excel(f'Sent/{new_name}_bom_ht.xlsx', sheet_name='input', index=False)
                    tebo_frame = generate_tebo_frame(data)
                    tebo_frame.to_excel(f'Sent/{new_name}_bom_tebo.xlsx', sheet_name='input', index=False)
                add_ref(new_name)

                send_results(new_name)

            else:
                print('Inserting manually searched parts directly')

                master_parts = pd.read_excel('masterdata.xlsx', sheet_name='data', usecols=['Manufacturer PART NO.'])[
                    'Manufacturer PART NO.'].tolist()
                master_sheet = pd.read_excel('masterdata.xlsx', sheet_name='data')
                try:
                    new_bom_frame = pd.read_excel(f'./Received/{new_bom}', sheet_name='input')
                    new_bom_frame = new_bom_frame.dropna()
                    new_bom_parts_only = new_bom_frame[['Manufacturer PART NO.']]
                except:
                    log.error(f'Please make sure that the file {new_bom} has a sheet name named "input"')
                    continue

                print(f'Comparing {new_bom} with masterdata sheet')

                new_bom_parts = new_bom_parts_only['Manufacturer PART NO.'].tolist()
                difference = [part for part in new_bom_parts if part not in master_parts]
                if difference:
                    try:
                        print('Inserting parts', difference)
                        filtered = new_bom_frame.loc[new_bom_frame['Manufacturer PART NO.'].isin(difference)]
                        master_sheet = master_sheet.append(filtered, ignore_index=True)
                        master_sheet.to_excel('masterdata.xlsx', sheet_name='data', index=False)
                    except:
                        print('Error while inserting parts', difference)
                else:
                    print('All parts exist in the masterdata sheet')
        except Exception as e:
            print('fail',str(e))
            print(CustomException(Exception))
            status = 'Fail'
            current_time = datetime.datetime.now()
            new_email_name = new_name.split('@')[0].replace('@','').replace(":",' ').replace('-',' ')
            
            email_message=str(e).replace('@','').replace(":",' ').replace('-',' ').replace("'",'')[:30]
            table1 =f"INSERT INTO mail_status (sendermail,status, message, created_date) VALUES ('{new_email_name}','{status}','{email_message}','{current_time}');"
            print(table1)
            cur.execute(table1)
            con.commit()


def add_ref(new_name):
    output_version = pd.read_excel(f'Sent/{new_name}_output.xlsx',sheet_name=0)
    bom_ht = pd.read_excel(f'Sent/{new_name}_bom_ht.xlsx',sheet_name=0)
    bom_ht['REF.'] = ''
    tebo = pd.read_excel(f'Sent/{new_name}_bom_tebo.xlsx',sheet_name=0)
    tebo['REF.'] = ''
    BOM = pd.read_excel(f'Sent/{new_name}_BOM.xlsx',sheet_name=0)
    BOM['REF.']  = ''
    for index,row in output_version.iterrows():
        part_nr = output_version.loc[index,'Manufacturer PART NO.']
        ref = output_version.loc[index,'REF.']
        #bom_ht
        for index_ht,row_ht in bom_ht.iterrows():
            if str(part_nr) == str(bom_ht.loc[index_ht,'Manufacturer PART NO.']):
                bom_ht.loc[index_ht, 'REF.'] = ref
        #bom_tebo
        for index_tebo,row_tebo in tebo.iterrows():
            if str(part_nr) == str(tebo.loc[index_tebo,'Manufacturer PART NO.']):
                tebo.loc[index_tebo, 'REF.'] = ref
        #BOM
        for  index_bom,row_BOM in BOM.iterrows():
            if  str(part_nr) == str(BOM.loc[index_bom,'Manufacturer PART NO.']):
                    BOM.loc[index_bom, 'REF.'] = ref
        pass
    bom_ht.to_excel(f'Sent/{new_name}_bom_ht.xlsx',index=False)
    tebo.to_excel(f'Sent/{new_name}_bom_tebo.xlsx',index=False)
    BOM.to_excel(f'Sent/{new_name}_BOM.xlsx',index=False)
if __name__ == "__main__":

    log = logging.getLogger('logger')
    formatter = logging.Formatter('%(asctime)s.%(msecs)03d :%(levelname)-8s| %(message)s', "%Y-%m-%d,%H:%M:%S")
    log.setLevel(logging.DEBUG)
    fh = logging.FileHandler('info.log', mode='a', encoding='utf-8')
    fh.setLevel(logging.DEBUG)
    fh.setFormatter(formatter)
    log.addHandler(fh)

    ch = logging.StreamHandler()
    ch.setLevel(logging.DEBUG)
    ch.setFormatter(formatter)
    log.addHandler(ch)

    SCOPES = ['https://www.googleapis.com/auth/gmail.modify']

    while True:
        time.sleep(3)
        print("No Mail")
        new_files, status = read_email_from_gmail()
        if status == 'Success':
            print(f'{len(new_files)} file(s) were retrieved')
            compare_with_master(new_files)
            continue
        elif status == 'Empty':
            print('No new e-mail(s)')
            continue
        else:
            continue

# compare_with_master([r'findbom.xls'])
