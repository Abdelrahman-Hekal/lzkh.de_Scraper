from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service as ChromeService 
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait as wait
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver import ActionChains
import undetected_chromedriver as uc
import time
import os
import re
from datetime import datetime
import pandas as pd
import warnings
import sys
import xlsxwriter
from multiprocessing import freeze_support
import calendar 
import shutil
warnings.filterwarnings('ignore')

def initialize_bot():

    # Setting up chrome driver for the bot
    chrome_options = uc.ChromeOptions()
    chrome_options.add_argument('--log-level=3')
    chrome_options.add_argument('--headless')
    chrome_options.add_experimental_option('excludeSwitches', ['enable-logging'])
    # installing the chrome driver
    driver_path = ChromeDriverManager().install()
    chrome_service = ChromeService(driver_path)
    # configuring the driver
    driver = webdriver.Chrome(options=chrome_options, service=chrome_service)
    ver = int(driver.capabilities['chrome']['chromedriverVersion'].split('.')[0])
    driver.quit()
    chrome_options = uc.ChromeOptions()
    chrome_options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/87.0.4280.88 Safari/537.36")
    chrome_options.add_argument('--log-level=3')
    chrome_options.add_argument("--enable-javascript")
    chrome_options.add_argument("--start-maximized")
    chrome_options.add_argument("--disable-gpu")
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--lang=en")
    chrome_options.add_argument("--incognito")
    chrome_options.add_argument('--headless=new')
    chrome_options.page_load_strategy = 'normal'
    driver = uc.Chrome(version_main = ver, options=chrome_options) 
    driver.set_window_size(1920, 1080)
    driver.maximize_window()
    driver.set_page_load_timeout(200)

    return driver

def scrape_lzkh(driver, output1):

    # getting the products list

    driver.get('https://www.lzkh.de/referat-ausbildung-zfa/stellenmarkt-praxispersonal/stellenangebote?tx_lzkhjobmarket_joblist%5Bcategory%5D=7&cHash=36eaa7ceacae37a94091e92410992263')
    time.sleep(2)

    # scraping Products details
    print('-'*75)
    print('Scraping Companies Details...')
    print('-'*75)

    data = pd.DataFrame()
    companies = wait(driver, 4).until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, "div[class='accordion-item']")))
    n = len(companies)
    keys = ['Company Name', 'Salutation', 'Titel', 'First Name', 'Last Name', 'Street', 'House number', 'Zip Code', 'City', 'Country', 'E-Mail', 'Website']
    for i, comp in enumerate(companies):
        try:
            details = {}
            for key in keys:
                details[key] = ''

            print(f'Scraping the details of company {i+1}\{n}')
            
            try:
                table = wait(comp, 4).until(EC.presence_of_element_located((By.CSS_SELECTOR, "dl[class='dl-horizontal']")))
            except Exception as err:
                print(f'warning: failed to access the table data for company {i+1}\{n}')
            
            # mail and website
            try:
                links = wait(table, 4).until(EC.presence_of_all_elements_located((By.TAG_NAME, "a")))
                for link in links:
                    text = link.get_attribute('textContent')
                    if 'www.' in text:
                        details['Website'] = text.strip()
                    elif '(at)' in text:
                        details['E-Mail'] = text.replace('(at)', '@').strip()
            except:
                print(f'Warning: Failed to scrape the website and mail for company {i+1}\{n}')

            # other info
            try:
                text = table.get_attribute('innerHTML')
                elems = text.replace('<br>', '').replace('<dt>', '').replace('<dd>', '').replace('</dt>', ';').replace('</dd>', ';').replace('amp;', '').strip().split(';')
                for j, elem in enumerate(elems):
                    if 'Praxisname:' in elem:
                        if '&' in elems[j+1]:
                            debug = True
                        details['Company Name'] = elems[j+1].strip()             
                    elif 'Adresse:' in elem:
                        try:
                            if elems[j+1].count('\n') > 0:
                                parts = elems[j+1].split('\n')
                                num = parts[0].split()[-1]
                                try:
                                    digits = int(num.replace('-', ''))
                                    if digits > 10000:
                                        details['Zip Code'] = digits
                                        details['City'] = parts[0].replace(details['Zip Code'], '').replace(',', '').replace('.', '').strip()
                                        details['House number'] = re.findall(r'[0-9]+', parts[1])[0]
                                        details['Street'] = parts[1].replace(details['House number'], '').replace(',', '').replace('.', '').strip() 
                                        continue
                                except:
                                    num = re.findall(r'[0-9]+', parts[0])[0]
                                    if int(num) > 10000:
                                        details['Zip Code'] = num
                                        details['City'] = parts[0].replace(details['Zip Code'], '').replace(',', '').replace('.', '').strip()
                                        details['House number'] = re.findall(r'[0-9]+', parts[1])[0]
                                        details['Street'] = parts[1].replace(details['House number'], '').replace(',', '').replace('.', '').strip() 
                                        continue

                                details['House number'] = num
                                details['Street'] = parts[0].replace(details['House number'], '').replace(',', '').replace('.', '').strip() 
                                try:
                                    details['Zip Code'] = re.findall(r'[0-9]+', parts[1])[0]
                                except:
                                    pass
                                details['City'] = parts[1].replace(details['Zip Code'], '').replace(',', '').replace('.', '').strip()
                            else:
                                text = elems[j+1]
                                nums = re.findall(r'[0-9]+', text)
                                for num in nums:
                                    if int(num) > 10000:
                                        details['Zip Code'] = num
                                    else:
                                        details['House number'] = num

                                text = text.replace(details['Zip Code'], '').replace(details['House number'], '').strip()
                                if ',' in text:
                                    details['Street'] = text.split(',')[0].replace(',', '').replace('.', '').strip()
                                    details['City'] = text.split(',')[-1].replace(',', '').replace('.', '').strip()
                                else:
                                    details['City'] = text.replace(',', '').replace('.', '').strip()
                        except:
                            pass

                        break
            except:
                pass

            # appending the output to the datafame       
            data = pd.concat([data, pd.DataFrame([details.copy()])], ignore_index=True)
        except Exception as err:
            print(f'Warning: the below error occurred while scraping the product link: {link}')
            print(str(err))
           
    # output to excel
    data = data.drop_duplicates()
    if data.shape[0] > 0:
        writer = pd.ExcelWriter(output1)
        data.to_excel(writer, index=False)
        writer.close()
    else:
        print('-'*75)
        print('No valid data is scraped')
       
def initialize_output():

    stamp = datetime.now().strftime("%d_%m_%Y_%H_%M")
    path = os.getcwd() + '\\Scraped_Data\\' + stamp
    if os.path.exists(path):
        shutil.rmtree(path)
    os.makedirs(path)

    file1 = f'vorlage_adressdatei_{stamp}.xlsx'

    # Windws and Linux slashes
    if os.getcwd().find('/') != -1:
        output1 = path.replace('\\', '/') + "/" + file1
    else:
        output1 = path + "\\" + file1  

    # Create an new Excel file and add a worksheet.
    workbook1 = xlsxwriter.Workbook(output1)
    workbook1.add_worksheet()
    workbook1.close()    

    return output1

def main():

    print('Initializing The Bot ...')
    start = time.time()
    output1 = initialize_output()
    try:
        driver = initialize_bot()
    except Exception as err:
        print('Failed to initialize the Chrome driver due to the following error:\n')
        print(str(err))
        print('-'*75)
        input('Press any key to exit.')
        sys.exit()

    try:
        scrape_lzkh(driver, output1)
    except Exception as err: 
        print(f'Warning: the below error occurred:\n {err}')
        driver.quit()
        time.sleep(5)
        driver = initialize_bot()

    driver.quit()
    print('-'*75)
    elapsed_time = round(((time.time() - start)/60), 4)
    hrs = round(elapsed_time/60, 4)
    input(f'Process is completed in {elapsed_time} mins ({hrs} hours), Press any key to exit.')
    sys.exit()

if __name__ == '__main__':

    main()

