import csv
import json
import os
import aiohttp
import os
import os.path
import random
import time
import asyncio
from bs4 import BeautifulSoup
import pandas as pd
from playwright.sync_api import sync_playwright
import requests
threed = {}
datasheets_list= {}
octopartHeaders = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 6.2; Win64; x64)\
        AppleWebKit/537.36 (KHTML, like Gecko)\
        Chrome/78.0.3770.100 Safari/537.36'
    }
arrowHeaders = {"User-Agent": "Mozilla/5.0 (X11; CrOS x86_64 12871.102.0) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/81.0.4044.141 Safari/537.36"}
mouserHeaders = {"User-Agent": 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/99.0.4844.82 Safari/537.36', 'Accept-Encoding': 'gzip, deflate, br', 'Connection': 'keep-alive'}

octopartData = {}
arrowData = {}
mouserData = {}
enrgtechData = {}
digikeyData = {}
failures = []
should_i_scrape = True
def parse_octopart(content,code):
    soup = BeautifulSoup(content, features="html.parser")
    divs = soup.find_all('div')
    for div in divs:
        try:
            if div.find('h3') != None:
                if div.find('h3').text == 'Technical Specifications':
                    for row in div.find_all('tr'):
                        if len(row.find_all('td')) == 2:
                            if row.find_all('td')[0].text == 'Height':
                                octopartData[code]['Octopart - Height'] = row.find_all('td')[1].text
                            if row.find_all('td')[0].text == 'Termination':
                                octopartData[code]['Octopart - Termination'] = row.find_all('td')[1].text
                            if row.find_all('td')[0].text == 'Mount':
                                octopartData[code]['Octopart - Mount'] = row.find_all('td')[1].text
        except:
            pass
    c = 0
    if octopartData[code]['Octopart - Height'] == 'N/A':
        c+=1
    if octopartData[code]['Octopart - Termination'] == 'N/A':
        c+=1
    if octopartData[code]['Octopart - Mount'] == 'N/A':
        c+=1
    if c <= 1 :
        failures.append(code)

###############################################################################################
#CASES WITH ERRORS
# code = FCF03FT01X ------ octopart result = Frontier FCF03FT-01X
# code = RK73H1JTTD-4701F ------ octopart result = KOA Speer RK73H1JTTD4701F
# code = RK73H1JTTD-1002F ------ octopart result = KOA Speer RK73H1JTTD1002F ------ octopart result = KOA RK73H1JTTD1002F
# code = MCR03EZPJ  ===== octopart outputs only partly codes with this nr
def extract_product_link(links,code):
    code = code.replace('-','')
    partly_include_code = []
    for link in links:
        #here i remove - if the code has or the code within product in octopart but the url should contain the code
        if str(code).lower() in str(link.text).lower().replace('-','') and str(link['href']).lower().endswith('r=sp'):
            part_nr_in_href = '-'.join(str(link['href']).lower().split('-')[:-2])[1:]
            if '/' in code:
                part_nr_in_href = part_nr_in_href.replace('%2f', '/')
            url = 'https://octopart.com' + str(link['href'] + str('#Specs'))
            if str(code.lower()) == part_nr_in_href.lower():
                return url
            else:
                # if you find the code in text of all product return, else append the list with product that only partly containng the link and  return the first objet of this list
                partly_include_code.append(url)
    if len(partly_include_code) > 0:
        return partly_include_code[0]
    else:
        return None
async def octopart_scraper_advanced(code):
    octopartData[code] = {
            'Octopart - Mount': 'N/A',
            'Octopart - Termination': 'N/A',
            'Octopart - Height': 'N/A'
        }
    headers = {
        "Accept": "application/json",
        "Content-Type": "application/json",
        "Authorization": "Basic VTAwMDAwODg1Mjc6S2Fsa2kxOTI0IQ=='"
    }
    if str(code) != 'nan' and '/' not in str(code):
        try:
            async with aiohttp.ClientSession() as session:
                # async with session.get(url='https://app.scrapingbee.com/api/v1/',params={'api_key': '75A4Z8NS5Q80JKI2VUIB6WT2301PB51K3Q5OBO9T17NBELEGSBTKSRX10XYAEYJA6NWKV4T80DBIM7G8','url': f'https://octopart.com/search?q={code}&specs=1','premium_proxy': 'true','country_code': 'us'}) as response:
                # async with session.post(url='https://scrape.smartproxy.com/v1/tasks',headers= headers,data={"target": "universal","parse": 'false',"url": f'https://octopart.com/search?q={code}&specs=1'}) as response:
                async with session.post(url='https://scrape.smartproxy.com/v1/tasks',headers= headers,json={        "target": "universal",
                    "parse": 'false',
                    "url": f"https://octopart.com/search?q={code}&specs=1"}) as response:
                    print('Octopart', code, response.status)
                    if response.status == 200:
                        response = await response.json()
                        response = response['results'][0]['content']
                        soup = BeautifulSoup(response, features="html.parser")
                        links = soup.find_all('a')
                        url = extract_product_link(links,code)
                        if url == None:
                            return ['N/A', 'N/A']
                        async with aiohttp.ClientSession() as session1:
                            # async with session1.get(url='https://app.scrapingbee.com/api/v1/',params={'api_key': '75A4Z8NS5Q80JKI2VUIB6WT2301PB51K3Q5OBO9T17NBELEGSBTKSRX10XYAEYJA6NWKV4T80DBIM7G8','url': url,'premium_proxy': 'true','country_code': 'us'}) as data:
                            async with session1.post(url='https://scrape.smartproxy.com/v1/tasks',headers=headers,json={"target": "universal", "parse": False,"url": url}) as data:
                                print('Octopart Second Response: ',data.status)
                                data = await data.json()
                                data = data['results'][0]['content']
                                parse_octopart(data,code)
                        await session1.close()
                        return [octopartData[code]['Octopart - Height'], octopartData[code]['Octopart - Mount']]
                    else:
                        #this doesn't need to get retried
                        failures.append(f'{code}!!')
                        print(f'Failed to scrape {code} from Octopart')
            await session.close()
        except Exception as e:
            print('Error -- ',str(e))
            return ['N/A','N/A']
    elif str(code) != 'nan' and '/' in str(code):
        try:
            async with aiohttp.ClientSession() as session2:
                async with session2.post(url='https://scrape.smartproxy.com/v1/tasks',headers= headers,json={"target": "universal","parse": False,"url": f'https://octopart.com/search?q={code}&specs=1'}) as response:
                    print('Octopart', code, response.status)
                    if response.status == 200:
                        response = await response.json()
                        response = response['results'][0]['content']
                        soup = BeautifulSoup(response, features="html.parser")
                        links = soup.find_all('a')
                        url = extract_product_link(links,code)
                        if url == None:
                            return ['N/A', 'N/A']
                        async with aiohttp.ClientSession() as session3:
                            # async with session3.get(url='https://app.scrapingbee.com/api/v1/', params={'api_key': '75A4Z8NS5Q80JKI2VUIB6WT2301PB51K3Q5OBO9T17NBELEGSBTKSRX10XYAEYJA6NWKV4T80DBIM7G8','url': url, 'premium_proxy': 'true', 'country_code': 'us'}) as data:
                            async with session3.post(url='https://scrape.smartproxy.com/v1/tasks',headers=headers,json={"target": "universal", "parse": False,"url": url}) as data:
                                print('Octopart Second Response: ',data.status)
                                data = await data.json()
                                data = data['results'][0]['content']
                                parse_octopart(data, code)
                        await session3.close()
                        return [octopartData[code]['Octopart - Height'], octopartData[code]['Octopart - Mount']]
                    else:
                        failures.append(f'{code}!!')
                        print(f'Failed to scrape {code} from Octopart')
            await session2.close()
        except Exception as e:
            print('Error -- ', str(e))
            return ['N/A','N/A']
###############################################################################################
async def octopart_scraper(code):
    octopartData[code] = {
            'Octopart - Mount': 'N/A',
            'Octopart - Termination': 'N/A',
            'Octopart - Height': 'N/A'
        }
    headers = {
        "Accept": "application/json",
        "Content-Type": "application/json",
        "Authorization": "Basic VTAwMDAwODg1Mjc6S2Fsa2kxOTI0IQ=='"
    }
    if str(code) != 'nan' and '/' not in str(code):
        try:
            async with aiohttp.ClientSession() as session:
                # async with session.get(url='https://app.scrapingbee.com/api/v1/',params={'api_key': '75A4Z8NS5Q80JKI2VUIB6WT2301PB51K3Q5OBO9T17NBELEGSBTKSRX10XYAEYJA6NWKV4T80DBIM7G8','url': f'https://octopart.com/search?q={code}&specs=1','premium_proxy': 'true','country_code': 'us'}) as response:
                # async with session.post(url='https://scrape.smartproxy.com/v1/tasks',headers= headers,data={"target": "universal","parse": 'false',"url": f'https://octopart.com/search?q={code}&specs=1'}) as response:
                async with session.post(url='https://scrape.smartproxy.com/v1/tasks',headers= headers,json={        "target": "universal",
                    "parse": 'false',
                    "url": f"https://octopart.com/search?q={code}&specs=1"}) as response:
                    print('Octopart', code, response.status)
                    if response.status == 200:
                        response = await response.json()
                        response = response['results'][0]['content']
                        soup = BeautifulSoup(response, features="html.parser")
                        links = soup.find_all('a')
                        for link in links:
                            if str(code).lower() in str(link.text).lower() and str(link['href']).lower().endswith('r=sp') and '-'.join(str(link['href']).lower().split('-')[:-2])[1:] == str(code).lower():
                                url = 'https://octopart.com' + str(link['href'] + str('#Specs'))
                                async with aiohttp.ClientSession() as session1:
                                    # async with session1.get(url='https://app.scrapingbee.com/api/v1/',params={'api_key': '75A4Z8NS5Q80JKI2VUIB6WT2301PB51K3Q5OBO9T17NBELEGSBTKSRX10XYAEYJA6NWKV4T80DBIM7G8','url': url,'premium_proxy': 'true','country_code': 'us'}) as data:
                                    async with session1.post(url='https://scrape.smartproxy.com/v1/tasks',headers=headers,json={"target": "universal", "parse": False,"url": url}) as data:
                                        print('Octopart Second Response: ',data.status)
                                        data = await data.json()
                                        data = data['results'][0]['content']
                                        parse_octopart(data,code)
                                await session1.close()
                                break
                    else:
                        print(f'Failed to scrape {code} from Octopart')
            await session.close()
        except Exception as e:
            print('Error -- ',str(e))
            pass
    elif str(code) != 'nan' and '/' in str(code):
        try:
            async with aiohttp.ClientSession() as session2:
                async with session2.post(url='https://scrape.smartproxy.com/v1/tasks',headers= headers,json={"target": "universal","parse": False,"url": f'https://octopart.com/search?q={code}&specs=1'}) as response:
                    print('Octopart', code, response.status)
                    if response.status == 200:
                        response = await response.json()
                        response = response['results'][0]['content']
                        soup = BeautifulSoup(response, features="html.parser")
                        links = soup.find_all('a')
                        for link in links:
                            if str(code).lower() in str(link.text).lower() and str(link['href']).lower().endswith(
                                    'r=sp') and '-'.join(str(link['href']).lower().split('-')[:-2])[1:].replace('%2f', '/') == str(
                                    code).lower():
                                url = 'https://octopart.com' + str(link['href'] + str('#Specs'))
                                async with aiohttp.ClientSession() as session3:
                                    # async with session3.get(url='https://app.scrapingbee.com/api/v1/', params={'api_key': '75A4Z8NS5Q80JKI2VUIB6WT2301PB51K3Q5OBO9T17NBELEGSBTKSRX10XYAEYJA6NWKV4T80DBIM7G8','url': url, 'premium_proxy': 'true', 'country_code': 'us'}) as data:
                                    async with session3.post(url='https://scrape.smartproxy.com/v1/tasks',headers=headers,json={"target": "universal", "parse": False,"url": url}) as data:
                                        print('Octopart Second Response: ',data.status)
                                        data = await data.json()
                                        data = data['results'][0]['content']
                                        parse_octopart(data, code)
                                await session3.close()
                                break
                    else:
                        print(f'Failed to scrape {code} from Octopart')
            await session2.close()
        except Exception as e:
            print('Error -- ', str(e))
            pass





async def mouser_scraper(code):
    global should_i_scrape
    arrowData[code]  = {
            'Arrow - Mount': 'N/A',
            'Arrow - Package Height': 'N/A',
            'Arrow - Height': 'N/A'
        }
    enrgtechData[code] = {
            'Enrgtech - Mounting': 'N/A',
            'Enrgtech - Package / Case': 'N/A',
            'Enrgtech - Height': 'N/A'
        }
    digikeyData[code] = {
            'Digikey - Mounting': 'N/A',
            'Digikey - Termination': 'N/A',
            'Digikey - Height': 'N/A'
        }
    mouserData[code] = {
            'Mouser - Mount': 'N/A',
            'Mouser - Termination': 'N/A',
            'Mouser - Height': 'N/A'
        }
    octopartData[code] = {
            'Octopart - Mount': 'N/A',
            'Octopart - Termination': 'N/A',
            'Octopart - Height': 'N/A'
        }
    headers = {
        "Accept": "application/json",
        "Content-Type": "application/json",
        "Authorization": "Basic VTAwMDAwODg1Mjc6S2Fsa2kxOTI0IQ=='"
    }
    while True:
        if should_i_scrape:
            print('Scraping Procces Begins')
            break
        if not should_i_scrape:
            await asyncio.sleep(2)
        if should_i_scrape == None:
            print('This part does not exist: ',code)
            failures.append(code)
            return
    if str(code) != 'nan':
        try:
            async with aiohttp.ClientSession() as session:
                async with session.post(url='https://scrape.smartproxy.com/v1/tasks',headers=headers,json={"target": "universal",
                    "parse": 'false',
                    "url": f"https://eu.mouser.com/c/?q={code}"}) as response:
                    print('Mouser', code, response.status)
                    if response.status == 200:
                        response = await response.json()
                        response = response['results'][0]['content']
                        soup = BeautifulSoup(response, features="html.parser")
                        try:
                            #working
                            for tableRow_ in soup.find('form', {'id': 'search-form'}).find('table', {'id': 'SearchResultsGrid_grid'}).find('tbody').find_all('tr'):
                                try:
                                    if code == str(tableRow_.find_all('td')[2].find('div').text).strip().split('No.')[1].strip():
                                        async with aiohttp.ClientSession() as session1:
                                            async with session1.post(url='https://scrape.smartproxy.com/v1/tasks',headers = headers, json={
                                                "target": "universal",
                                                "parse": 'false',
                                                'url': 'https://eu.mouser.com' + str(tableRow_.find_all('td')[2].find('div').a['href'])}) as productResponse:
                                                print('Mouser Second Response:',productResponse.status)
                                                productResponse = await productResponse.json()
                                                productResponse = productResponse['results'][0]['content']
                                                productSoup = BeautifulSoup(productResponse, features="html.parser")
                                                full_table = productSoup.find('table', {'class': 'specs-table'})
                                                for tableRow in full_table.find_all('tr'):
                                                    try:
                                                        if 'Termination Style:' == str(tableRow.find_all('td')[0].text).strip():
                                                            mouserData[code]['Mouser - Termination'] = str(tableRow.find_all('td')[1].text).strip()
                                                        elif 'Height:' in str(tableRow.find_all('td')[0].text).strip():
                                                            mouserData[code]['Mouser - Height'] = str(tableRow.find_all('td')[1].text).strip()
                                                    except Exception as ss:
                                                        continue
                                            await session1.close()
                                except Exception as eee:
                                    continue
                        except Exception as e:

                            for tableRow in soup.find('table', {'class': 'specs-table'}):
                                try:
                                    if 'Termination Style:' == str(tableRow.find_all('td')[0].text).strip():
                                        mouserData[code]['Mouser - Termination'] = str(tableRow.find_all('td')[1].text).strip()
                                    elif 'Height:' in str(tableRow.find_all('td')[0].text).strip():
                                        mouserData[code]['Mouser - Height'] = str(tableRow.find_all('td')[1].text).strip()
                                except :
                                    continue
                    else:
                        pass
                        # log.error(f'Failed to scrape {code} from Mouser')
            await session.close()
        except Exception as e:
            print('Mouser error:',str(e))
            pass

        c = 0
        if mouserData[code]['Mouser - Mount'] != 'N/A':
            c+=1
        if mouserData[code]['Mouser - Termination'] != 'N/A':
            c+=1
        if mouserData[code]['Mouser - Height'] != 'N/A':
            c+=1
        if c < 2:
            await octopart_scraper(code)
            print(octopartData[code])
        else:
            print(mouserData[code])
#PROVIDE A SOLUTION TE FETCH COOKIES AUTOMATICALLY ID THEY DO NOT EXIST
def login_to_snap(browser):
    print('Creating new session')
    ua  ="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/99.0.4844.82 Safari/537.36"
    context =  browser.new_context(user_agent=ua)
    page =  context.new_page()
    url = "https://www.snapeda.com/account/login/"
    page.goto(url,timeout=50000,wait_until='networkidle')
    page.fill('//*[@id="id_username"]','karan@wavepallets.com')
    time.sleep(3)
    page.fill('//*[@id="id_password"]','kkkkkkk')
    time.sleep(3)
    page.click('//*[@id="s-form-login"]/fieldset/div/input')
    page.wait_for_selector(selector='//*[@id="header-search"]/div/h1',timeout=50000)
    context.storage_state(path='storage_state.json')
    context.close()

def parse_cookies():
    if os.path.exists(f'{os.getcwd()}/storage_state.json'):
        pass
    else:
        with sync_playwright() as p:
            chromium = p.chromium
            browser = chromium.launch(headless=False)
            login_to_snap(browser)
    with open('storage_state.json','r') as f:
        data = json.load(f)
        request_session = requests.session()
        for cookie in data['cookies']:
            c = {cookie['name']: cookie['value']}
            request_session.cookies.update(c)
        return request_session

async def downloadFiles_async(partNumber, session_):
    global should_i_scrape
    if str(partNumber) != 'nan':
        try:
            response = session_.get(f'https://www.snapeda.com/api/v1/search_local?q={partNumber}&amp;SEARCH=Search&amp;page=1&amp;search-type=parts').json()
            for part in response['results']:
                if str(part['name']).strip().lower() == str(partNumber).lower():
                    should_i_scrape = True
                    if part['has_3d']:
                        print(f'Retrieving 3D file for {partNumber}')
                        async with aiohttp.ClientSession() as session:
                            async with session.get(url=f'https://www.snapeda.com/parts/{partNumber}/{str(part["manufacturer"]).replace(" ","%20")}/view-part/#download-3d-modal') as file_content:
                                file_content = await file_content.read()
                                file_part_number = str(partNumber).replace('/', '_')
                                threed[str(partNumber)]= f'https://www.snapeda.com/parts/{partNumber}/{str(part["manufacturer"]).replace(" ","%20")}/view-part/#download-3d-modal'
                                open(f'3dModels/{file_part_number}_3dModel.STEP', 'wb').write(file_content)
                                await session.close()

                    if part['has_datasheet']:
                        async with aiohttp.ClientSession() as session1:
                            async with session1.get(url=part['datasheeturl']) as file_content1:
                                print(f'Retrieving datasheet document for {partNumber}')
                                file_content1 = await file_content1.read()
                                file_part_number = str(partNumber).replace('/', '_')
                                open(f'datasheets/{file_part_number}_datasheet.pdf', 'wb').write(file_content1)
                                datasheets_list[str(partNumber)] = part['datasheeturl']
                                await session1.close()
                    break
            should_i_scrape = None
        except Exception as e:
            print(str(e))

            print(f'Error accessing {partNumber} to download datasheet/3dModel')

def main(parts):

    global paths
    event_loop = asyncio.get_event_loop()
    for code in parts:
        try:
            event_loop.run_until_complete(async_run_main_wrapper(code,parse_cookies()))
            print('###################################')
        except:
            pass
    octopartFrame = pd.DataFrame.from_dict(octopartData).transpose()
    arrowFrame = pd.DataFrame.from_dict(arrowData).transpose()
    mouserFrame = pd.DataFrame.from_dict(mouserData).transpose()
    digikeyFrame = pd.DataFrame.from_dict(digikeyData).transpose()
    enrgtechFrame = pd.DataFrame.from_dict(enrgtechData).transpose()
    paths = {
        '3d':threed,'datasheets':datasheets_list}
    return octopartFrame, arrowFrame, mouserFrame, digikeyFrame, enrgtechFrame, failures, paths

async def async_run_main_wrapper(code,_session):
    global should_i_scrape
    should_i_scrape = False
    download = asyncio.ensure_future(downloadFiles_async(code,_session))
    mouser_scraper_async = asyncio.ensure_future(mouser_scraper(code))
    await asyncio.gather(mouser_scraper_async,download)

# #SN74LVC2G34DBVR
global retried_results
def failures_fun(code):
    e_loop = asyncio.get_event_loop()
    e_loop.run_until_complete(main_test(code))
    return retried_results
async def main_test(code):
    global retried_results
    retried_results = await asyncio.gather(asyncio.ensure_future(octopart_scraper_advanced(code)))

#CASES WITH ERRORS
# code = FCF03FT01X ------ octopart result = Frontier FCF03FT-01X
# code = RK73H1JTTD-4701F ------ octopart result = KOA Speer RK73H1JTTD4701F
# code = RK73H1JTTD-1002F ------ octopart result = KOA Speer RK73H1JTTD1002F ------ octopart result = KOA RK73H1JTTD1002F
# code = MCR03EZPJ  ===== octopart outputs only partly codes with this nr

