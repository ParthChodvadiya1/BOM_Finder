import os
import time
from multiprocessing.dummy import Pool
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.keys import Keys
import requests
from bs4 import BeautifulSoup
import pandas as pd
import logging


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


def octopart_scraper(code):
    octopartData[code] = {
            'Octopart - Mount': 'N/A',
            'Octopart - Termination': 'N/A',
            'Octopart - Height': 'N/A'
        }
    if str(code) != 'nan' and '/' not in str(code):
        try:
            options = Options()
            options.add_argument(
                "user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/98.0.4844.51 Safari/537.36")
            options.add_argument('--headless')
            options.add_experimental_option('excludeSwitches', ['enable-logging'])
            driverService = Service('./chromedriver.exe')
            driver = webdriver.Chrome(options=options)
            response = requests.get(f'https://octopart.com/search?q={code}&specs=1', headers=octopartHeaders)
            print('Octopart', code, response.status_code)
            if response.status_code == 200:
                soup = BeautifulSoup(response.content, features="html.parser")
                links = soup.find_all('a')
                for link in links:
                    if str(code).lower() in str(link.text).lower() and str(link['href']).lower().endswith('r=sp') and '-'.join(str(link['href']).lower().split('-')[:-2])[1:] == str(code).lower():
                        url = 'https://octopart.com' + str(link['href'] + str('#Specs'))
                        driver.get(url)
                        WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.ID, 'Specs')))
                        data = driver.execute_script("return __NEXT_DATA__['props']['apolloState']['data']")
                        for item in data:
                            try:
                                if 'Attribute:' in str(item) and data[item]['name'] == 'Mount':
                                    mountID = item
                                elif 'Attribute:' in str(item) and data[item]['name'] == 'Height':
                                    heightID = item
                                elif 'Attribute:' in str(item) and data[item]['name'] == 'Termination':
                                    terminationID = item
                            except:
                                continue

                        for item in data:
                            try:
                                if data[item]['__typename'] == 'Spec' and data[item]['attribute']['id'] == mountID and str(item.split('EG-USD-')[1].split('.')[0]) in url:
                                    octopartData[code]['Octopart - Mount'] = data[item]['display_value']
                                elif data[item]['__typename'] == 'Spec' and data[item]['attribute']['id'] == heightID and str(item.split('EG-USD-')[1].split('.')[0]) in url:
                                    octopartData[code]['Octopart - Height'] = data[item]['display_value']
                                elif data[item]['__typename'] == 'Spec' and data[item]['attribute']['id'] == terminationID and str(item.split('EG-USD-')[1].split('.')[0]) in url:
                                    octopartData[code]['Octopart - Termination'] = data[item]['display_value']
                            except:
                                continue
            else:
                log.error(f'Failed to scrape {code} from Octopart')
                failures.append(code)
        except:
            pass
    elif str(code) != 'nan' and '/' in str(code):
        try:
            options = Options()
            options.add_argument(
                "user-agent=Mozilla/5.0 (Windows NT 6.1; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/84.0.4147.125 Safari/537.36")
            options.add_argument('--headless')
            options.add_experimental_option('excludeSwitches', ['enable-logging'])
            driverService = Service('./chromedriver.exe')
            driver = webdriver.Chrome(options=options)
            response = requests.get(f'https://octopart.com/search?q={code}&specs=1', headers=octopartHeaders)
            print('Octopart', code, response.status_code)
            if response.status_code == 200:
                soup = BeautifulSoup(response.content, features="html.parser")
                links = soup.find_all('a')
                for link in links:
                    try:
                        if str(code).lower() in str(link.text).lower() and str(link['href']).lower().endswith('r=sp') and '-'.join(str(link['href']).lower().split('-')[:-2])[1:].replace('%2f', '/') == str(code).lower():
                            url = 'https://octopart.com' + str(link['href'] + str('#Specs'))
                            driver.get(url)
                            WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.ID, 'Specs')))
                            data = driver.execute_script("return __NEXT_DATA__['props']['apolloState']['data']")
                            for item in data:
                                try:
                                    if 'Attribute:' in str(item) and data[item]['name'] == 'Mount':
                                        mountID = item
                                    elif 'Attribute:' in str(item) and data[item]['name'] == 'Height':
                                        heightID = item
                                    elif 'Attribute:' in str(item) and data[item]['name'] == 'Termination':
                                        terminationID = item
                                except:
                                    continue

                            for item in data:
                                try:
                                    if data[item]['__typename'] == 'Spec' and data[item]['attribute']['id'] == mountID and str(item.split('EG-USD-')[1].split('.')[0]) in url:
                                        octopartData[code]['Octopart - Mount'] = data[item]['display_value']
                                    elif data[item]['__typename'] == 'Spec' and data[item]['attribute']['id'] == heightID and str(item.split('EG-USD-')[1].split('.')[0]) in url:
                                        octopartData[code]['Octopart - Height'] = data[item]['display_value']
                                    elif data[item]['__typename'] == 'Spec' and data[item]['attribute']['id'] == terminationID and str(item.split('EG-USD-')[1].split('.')[0]) in url:
                                        octopartData[code]['Octopart - Termination'] = data[item]['display_value']
                                except:
                                    continue
                    except:
                        continue
            else:
                log.error(f'Failed to scrape {code} from Octopart')
                failures.append(code)
        except:
            pass


def arrow_scraper(code):
    arrowData[code] = {
            'Arrow - Mount': 'N/A',
            'Arrow - Package Height': 'N/A',
            'Arrow - Height': 'N/A'
        }
    if str(code) != 'nan':
        try:
            response = requests.get(f'https://www.arrow.com/en/products/search?q={code}', headers=arrowHeaders)
            print('Arrow', code, response.status_code)
            if response.status_code == 200:
                soup = BeautifulSoup(response.content, features="html.parser")
                try:
                    for tableRow in soup.find('div', {'id': 'Pdp-specifications'}).find('table').find_all('tr'):
                        try:
                            if 'Mounting' == str(tableRow.find_all('td')[0].text).strip():
                                arrowData[code]['Arrow - Mount'] = str(tableRow.find_all('td')[1].text).strip()
                            elif 'Product Height (mm)' in str(tableRow.find_all('td')[0].text).strip():
                                arrowData[code]['Arrow - Height'] = str(tableRow.find_all('td')[1].text).strip()
                            elif 'Package Height' == str(tableRow.find_all('td')[0].text).strip():
                                arrowData[code]['Arrow - Height'] = str(tableRow.find_all('td')[1].text).strip()
                        except:
                            continue
                except:
                    for link in soup.find_all('a', {'class': 'SearchResults-productLink'}):
                        if str(code) == str(link.find('span', {'class': 'SearchResults-productName'}).text).strip():
                            productResponse = requests.get('https://www.arrow.com' + str(link['href']), headers=arrowHeaders)
                            productSoup = BeautifulSoup(productResponse.content, features="html.parser")
                            for tableRow in productSoup.find('div', {'id': 'Pdp-specifications'}).find('table').find_all('tr'):
                                try:
                                    if 'Mounting' == str(tableRow.find_all('td')[0].text).strip():
                                        arrowData[code]['Arrow - Mount'] = str(tableRow.find_all('td')[1].text).strip()
                                    elif 'Product Height (mm)' in str(tableRow.find_all('td')[0].text).strip():
                                        arrowData[code]['Arrow - Height'] = str(tableRow.find_all('td')[1].text).strip()
                                    elif 'Package Height' == str(tableRow.find_all('td')[0].text).strip():
                                        arrowData[code]['Arrow - Height'] = str(tableRow.find_all('td')[1].text).strip()
                                except:
                                    continue
                            break
            else:
                log.error(f'Failed to scrape {code} from Arrow')
                if code not in failures:
                    failures.append(code)
        except:
            pass


def mouser_scraper(code):
    mouserData[code] = {
            'Mouser - Mount': 'N/A',
            'Mouser - Termination': 'N/A',
            'Mouser - Height': 'N/A'
        }
    if str(code) != 'nan':
        try:
            response = requests.get(f'https://eu.mouser.com/c/?q={code}', headers=mouserHeaders)
            print('Mouser', code, response.status_code)
            if response.status_code == 200:
                soup = BeautifulSoup(response.content, features="html.parser")
                try:
                    for tableRow_ in soup.find('form', {'id': 'search-form'}).find('table', {'id': 'SearchResultsGrid_grid'}).find('tbody').find_all('tr'):
                        try:
                            if code == str(tableRow_.find_all('td')[2].find('div').text).strip().split('No.')[1].strip():
                                productResponse = requests.get('https://eu.mouser.com' + str(tableRow_.find_all('td')[2].find('div').a['href']), headers=mouserHeaders)
                                productSoup = BeautifulSoup(productResponse.content, features="html.parser")
                                for tableRow in productSoup.find('table', {'class': 'specs-table'}):
                                    try:
                                        if 'Termination Style:' == str(tableRow.find_all('td')[0].text).strip():
                                            mouserData[code]['Mouser - Termination'] = str(tableRow.find_all('td')[1].text).strip()
                                        elif 'Height:' in str(tableRow.find_all('td')[0].text).strip():
                                            mouserData[code]['Mouser - Height'] = str(tableRow.find_all('td')[1].text).strip()
                                    except:
                                        continue
                        except:
                            continue
                except:
                    for tableRow in soup.find('table', {'class': 'specs-table'}):
                        try:
                            if 'Termination Style:' == str(tableRow.find_all('td')[0].text).strip():
                                mouserData[code]['Mouser - Termination'] = str(tableRow.find_all('td')[1].text).strip()
                            elif 'Height:' in str(tableRow.find_all('td')[0].text).strip():
                                mouserData[code]['Mouser - Height'] = str(tableRow.find_all('td')[1].text).strip()
                        except:
                            continue
            else:
                log.error(f'Failed to scrape {code} from Mouser')
                if code not in failures:
                    failures.append(code)
                # options = Options()
                # options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/99.0.4844.82 Safari/537.36")
                # # options.add_argument('--headless')
                # options.add_experimental_option('excludeSwitches', ['enable-logging'])
                # driverService = Service('./chromedriver.exe')
                # driver = webdriver.Chrome(options=options)
                # driver.get(f'https://eu.mouser.com/c/?q={code}')
                # response = driver.page_source
                # soup = BeautifulSoup(response.content, features="html.parser")
                # try:
                #     for tableRow_ in soup.find('form', {'id': 'search-form'}).find('table', {
                #         'id': 'SearchResultsGrid_grid'}).find('tbody').find_all('tr'):
                #         try:
                #             if code == str(tableRow_.find_all('td')[2].find('div').text).strip().split('No.')[
                #                 1].strip():
                #                 productResponse = requests.get(
                #                     'https://eu.mouser.com' + str(tableRow_.find_all('td')[2].find('div').a['href']),
                #                     headers=mouserHeaders)
                #                 productSoup = BeautifulSoup(productResponse.content, features="html.parser")
                #                 for tableRow in productSoup.find('table', {'class': 'specs-table'}):
                #                     try:
                #                         if 'Termination Style:' == str(tableRow.find_all('td')[0].text).strip():
                #                             mouserData[code]['Mouser - Termination'] = str(
                #                                 tableRow.find_all('td')[1].text).strip()
                #                         elif 'Height:' in str(tableRow.find_all('td')[0].text).strip():
                #                             mouserData[code]['Mouser - Height'] = str(
                #                                 tableRow.find_all('td')[1].text).strip()
                #                     except:
                #                         continue
                #         except:
                #             continue
                # except:
                #     for tableRow in soup.find('table', {'class': 'specs-table'}):
                #         try:
                #             if 'Termination Style:' == str(tableRow.find_all('td')[0].text).strip():
                #                 mouserData[code]['Mouser - Termination'] = str(tableRow.find_all('td')[1].text).strip()
                #             elif 'Height:' in str(tableRow.find_all('td')[0].text).strip():
                #                 mouserData[code]['Mouser - Height'] = str(tableRow.find_all('td')[1].text).strip()
                #         except:
                #             continue
        except:
            pass


def enrgtech_scraper(code):
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


def digikey_scraper(code):
    digikeyData[code] = {
            'Digikey - Mounting': 'N/A',
            'Digikey - Termination': 'N/A',
            'Digikey - Height': 'N/A'
        }
    if str(code) != 'nan':
        try:
            response = requests.get(f'https://www.digikey.com/en/products/result?keywords={code}', headers=mouserHeaders)
            print('digikey', code, response.status_code)
            if response.status_code == 200:
                soup = BeautifulSoup(response.content, features="html.parser")
                try:
                    for row_ in soup.find('table', {'id': 'product-attributes'}).find('tbody').find_all('tr'):
                        try:
                            if 'Height - Seated (Max)' == str(row_.find_all('td')[0].text).strip():
                                digikeyData[code]['Digikey - Height'] = str(row_.find_all('td')[1].text).strip()
                            elif 'Mounting Type' in str(row_.find_all('td')[0].text).strip():
                                digikeyData[code]['Digikey - Mounting'] = str(row_.find_all('td')[1].text).strip()
                        except:
                            continue
                except:
                    try:
                        productLink = soup.find('div', {'class': 'jss124'}).find('div', {'class': 'jss130'}).a['href']
                        productResponse = requests.get(f'https://www.digikey.com/{productLink}', headers=mouserHeaders)
                        productSoup = BeautifulSoup(productResponse.content, features="html.parser")
                        for row_ in productSoup.find('table', {'id': 'product-attributes'}).find('tbody').find_all('tr'):
                            try:
                                if 'Height - Seated (Max)' == str(row_.find_all('td')[0].text).strip():
                                    digikeyData[code]['Digikey - Height'] = str(row_.find_all('td')[1].text).strip()
                                elif 'Mounting Type' in str(row_.find_all('td')[0].text).strip():
                                    digikeyData[code]['Digikey - Mounting'] = str(row_.find_all('td')[1].text).strip()
                            except:
                                continue
                    except:
                        pass
            else:
                log.error(f'Failed to scrape {code} from Digikey')
                if code not in failures:
                    failures.append(code)
        except:
            pass


def login_to_snap():
    options = Options()
    options.add_argument(
        "user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/99.0.4844.82 Safari/537.36")
    options.add_argument('--headless')
    options.add_experimental_option('excludeSwitches', ['enable-logging'])
    driverService = Service('./chromedriver.exe')
    driver = webdriver.Chrome(options=options)
    url = "https://www.snapeda.com/account/login/"
    driver.get(url)
    WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.ID, 's-form-login')))
    username_box = driver.find_element(By.ID, 'id_username')
    username_box.send_keys('karan@wavepallets.com')
    password_box = driver.find_element(By.ID, 'id_password')
    password_box.send_keys('kkkkkkk')
    password_box.send_keys(Keys.ENTER)
    request_session = requests.session()
    for cookie in driver.get_cookies():
        c = {cookie['name']: cookie['value']}
        request_session.cookies.update(c)
    return request_session


def downloadFiles(partNumber, session_):
    if str(partNumber) != 'nan':
        try:
            response = session_.get(f'https://www.snapeda.com/api/v1/search_local?q={partNumber}&amp;SEARCH=Search&amp;page=1&amp;search-type=parts').json()
            for part in response['results']:
                if str(part['name']).strip().lower() == str(partNumber).lower():
                    if part['has_3d']:
                        paths['3d'].append(partNumber)
                        print(f'Retrieving 3D file for {partNumber}')
                        file_content = requests.get(f'https://www.snapeda.com/parts/{partNumber}/{str(part["manufacturer"]).replace(" ","%20")}/view-part/#download-3d-modal').content
                        file_part_number = str(partNumber).replace('/', '_')
                        open(f'3dModels/{file_part_number}_3dModel.STEP', 'wb').write(file_content)

                    if part['has_datasheet']:
                        paths['datasheets'].append(partNumber)
                        print(f'Retrieving datasheet document for {partNumber}')
                        file_content = requests.get(part['datasheeturl']).content
                        file_part_number = str(partNumber).replace('/', '_')
                        open(f'datasheets/{file_part_number}_datasheet.pdf', 'wb').write(file_content)
                    break
        except:
            print(f'Error accessing {partNumber} to download datasheet/3dModel')


def main(parts):
    global paths
    paths = {
        '3d': [],
        'datasheets': []
    }
    p = Pool(5)
    p.map(octopart_scraper, parts)
    p.map(arrow_scraper, parts)
    p.map(digikey_scraper, parts)
    p.map(mouser_scraper, parts)
    p.map(enrgtech_scraper, parts)
    requestSession = login_to_snap()
    sessions = []
    for partt_ in parts:
        sessions.append(requestSession)
    p.starmap(downloadFiles, zip(parts, sessions))


    octopartFrame = pd.DataFrame.from_dict(octopartData).transpose()
    arrowFrame = pd.DataFrame.from_dict(arrowData).transpose()
    mouserFrame = pd.DataFrame.from_dict(mouserData).transpose()
    digikeyFrame = pd.DataFrame.from_dict(digikeyData).transpose()
    enrgtechFrame = pd.DataFrame.from_dict(enrgtechData).transpose()

    return octopartFrame, arrowFrame, mouserFrame, digikeyFrame, enrgtechFrame, failures, paths


if __name__ == "__main__":

    log = logging.getLogger('logger')
    formatter = logging.Formatter('%(asctime)s.%(msecs)03d :%(levelname)-8s| %(message)s' ,"%Y-%m-%d,%H:%M:%S")
    log.setLevel(logging.DEBUG)
    fh = logging.FileHandler('info.log', mode='a', encoding='utf-8')
    fh.setLevel(logging.DEBUG)
    fh.setFormatter(formatter)
    log.addHandler(fh)

    ch = logging.StreamHandler()
    ch.setLevel(logging.DEBUG)
    ch.setFormatter(formatter)
    log.addHandler(ch)






























    # for index, row_ in inputData.iterrows():
    #     octopart_scraper(inputData.loc[index, 'Manufacturer PART NO.'])
    #     inputData.loc[index, 'Octopart - Mount'] = octopartData[inputData.loc[index, 'Manufacturer PART NO.']]['mount']
    #     inputData.loc[index, 'Octopart - Height'] = octopartData[inputData.loc[index, 'Manufacturer PART NO.']]['height']
    #     inputData.loc[index, 'Octopart - Termination'] = octopartData[inputData.loc[index, 'Manufacturer PART NO.']]['termination']
    #     arrow_scraper(inputData.loc[index, 'Manufacturer PART NO.'])
    #     inputData.loc[index, 'Arrow - Mount'] = arrowData[inputData.loc[index, 'Manufacturer PART NO.']]['mount']
    #     inputData.loc[index, 'Arrow - Height'] = arrowData[inputData.loc[index, 'Manufacturer PART NO.']]['height']
    #     inputData.loc[index, 'Arrow - Package Height'] = arrowData[inputData.loc[index, 'Manufacturer PART NO.']]['package_height']
    #     mouser_scraper(inputData.loc[index, 'Manufacturer PART NO.'])
    #     inputData.loc[index, 'Mouser - Termination'] = mouserData[inputData.loc[index, 'Manufacturer PART NO.']]['termination']
    #     inputData.loc[index, 'Mouser - Height'] = mouserData[inputData.loc[index, 'Manufacturer PART NO.']]['height']
    #     enrgtech_scraper(inputData.loc[index, 'Manufacturer PART NO.'])
    #     inputData.loc[index, 'Enrgtech - Mounting'] = enrgtechData[inputData.loc[index, 'Manufacturer PART NO.']]['mount']
    #     inputData.loc[index, 'Enrgtech - Package / Case'] = enrgtechData[inputData.loc[index, 'Manufacturer PART NO.']]['package_case']
    #     inputData.loc[index, 'Enrgtech - Height'] = enrgtechData[inputData.loc[index, 'Manufacturer PART NO.']]['height']
    #     digikey_scraper(inputData.loc[index, 'Manufacturer PART NO.'])
    #     inputData.loc[index, 'Digikey - Mounting'] = digikeyData[inputData.loc[index, 'Manufacturer PART NO.']]['mount']
    #     inputData.loc[index, 'Digikey - Height'] = digikeyData[inputData.loc[index, 'Manufacturer PART NO.']]['height']
    #
    # inputData.to_excel('output.xlsx', index=False)
    #
    # data = pd.read_excel('output.xlsx', sheet_name='Sheet1')
    # data = data.astype(str)
    # data['HT'] = 'N/A'
    # data['SMD/Through'] = 'N/A'
    # for index, row in data.iterrows():
    #     inputHeights = [data.loc[index, 'Octopart - Height'], data.loc[index, 'Arrow - Height'],
    #                     data.loc[index, 'Octopart - Height'], data.loc[index, 'Mouser - Height'],
    #                     data.loc[index, 'Enrgtech - Height'], data.loc[index, 'Arrow - Package Height']]
    #     if not all(x == 'nan' for x in inputHeights):
    #         heights = []
    #         octoHieght = str(data.loc[index, 'Octopart - Height']).split(' ')
    #         arrowHieght = str(data.loc[index, 'Arrow - Height']).split('(')[0]
    #         mouserHieght = str(data.loc[index, 'Mouser - Height']).split(' ')
    #         digikeyHieght = str(data.loc[index, 'Digikey - Height']).split('(')
    #         try:
    #             if octoHieght[1] == 'µm':
    #                 heights.append(float(octoHieght[0]) / 1000)
    #             elif octoHieght[1] == 'mm':
    #                 heights.append(float(octoHieght[0]))
    #         except:
    #             pass
    #
    #         try:
    #             heights.append(float(arrowHieght))
    #         except:
    #             pass
    #
    #         try:
    #             if mouserHieght[1] == 'µm':
    #                 heights.append(float(mouserHieght[0]) / 1000)
    #             elif mouserHieght[1] == 'mm':
    #                 heights.append(float(mouserHieght[0]))
    #         except:
    #             pass
    #
    #         try:
    #             if 'µm' in digikeyHieght[1]:
    #                 heights.append(float(digikeyHieght[1].split('um)')[0]) / 1000)
    #             elif 'mm' in digikeyHieght[1]:
    #                 heights.append(float(digikeyHieght[1].split('mm)')[0]))
    #         except:
    #             pass
    #         data.loc[index, 'HT'] = max(heights)
    #     mountTypes = [data.loc[index, 'Octopart - Mount'], data.loc[index, 'Arrow - Mount'],
    #                   data.loc[index, 'Octopart - Termination'], data.loc[index, 'Mouser - Termination'],
    #                   data.loc[index, 'Enrgtech - Mounting']]
    #     if not all(x == 'nan' for x in mountTypes):
    #         if 'SMD/SMT' in mountTypes or 'Surface Mount' in mountTypes:
    #             data.loc[index, 'SMD/Through'] = 'SMD'
    #         else:
    #             data.loc[index, 'SMD/Through'] = 'Through Hole'
    #
    # output_version = data[['VALUE', 'REF.', 'HT', 'SMD/Through']]
    # data.to_excel('output.xlsx', index=False)
    # output_version.to_excel('bom_output.xlsx', sheet_name='input', index=False)






# except:
#     print('Please make sure that the file name is "BOM" and the sheet name is "input"')
#     print('Terminating..')
#     exit()

# data = pd.read_excel('output.xlsx', sheet_name='Sheet1')
# data = data.astype(str)
# data['HT'] = 'N/A'
# data['SMD/Through'] = 'N/A'
# for index, row in data.iterrows():
#     inputHeights = [data.loc[index, 'Octopart - Height'], data.loc[index, 'Arrow - Height'],
#                   data.loc[index, 'Octopart - Height'], data.loc[index, 'Mouser - Height'],
#                   data.loc[index, 'Enrgtech - Height'], data.loc[index, 'Arrow - Package Height']]
#     print(inputHeights)
#     if not all(x == 'nan' for x in inputHeights):
#         heights = []
#         octoHieght = str(data.loc[index, 'Octopart - Height']).split(' ')
#         arrowHieght = str(data.loc[index, 'Arrow - Height']).split('(')[0]
#         mouserHieght = str(data.loc[index, 'Mouser - Height']).split(' ')
#         digikeyHieght = str(data.loc[index, 'Digikey - Height']).split('(')
#         try:
#             if octoHieght[1] == 'µm':
#                 heights.append(float(octoHieght[0]) / 1000)
#             elif octoHieght[1] == 'mm':
#                 heights.append(float(octoHieght[0]))
#         except:
#             pass
#
#         try:
#             heights.append(float(arrowHieght))
#         except:
#             pass
#
#         try:
#             if mouserHieght[1] == 'µm':
#                 heights.append(float(mouserHieght[0]) / 1000)
#             elif mouserHieght[1] == 'mm':
#                 heights.append(float(mouserHieght[0]))
#         except:
#             pass
#
#         try:
#             if 'µm' in digikeyHieght[1]:
#                 heights.append(float(digikeyHieght[1].split('um)')[0]) / 1000)
#             elif 'mm' in digikeyHieght[1]:
#                 heights.append(float(digikeyHieght[1].split('mm)')[0]))
#         except:
#             pass
#         print(heights)
#         data.loc[index, 'HT'] = max(heights)
#     mountTypes = [data.loc[index, 'Octopart - Mount'], data.loc[index, 'Arrow - Mount'],
#                   data.loc[index, 'Octopart - Termination'], data.loc[index, 'Mouser - Termination'],
#                   data.loc[index, 'Enrgtech - Mounting']]
#     if not all(x == 'nan' for x in mountTypes):
#         if 'SMD/SMT' in mountTypes or 'Surface Mount' in mountTypes:
#             data.loc[index, 'SMD/Through'] = 'SMD'
#         else:
#             data.loc[index, 'SMD/Through'] = 'Through Hole'
#
# data.to_excel('output.xlsx', index=False)
