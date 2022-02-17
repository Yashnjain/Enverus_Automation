from selenium import webdriver
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.common.by import By
import time
from datetime import date
import logging
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.firefox import GeckoDriverManager
import sharepy
import os
from bu_config import get_config
import bu_alerts


today_date=date.today()
# log progress --
logfile = os.getcwd() +"\\logs\\"+'Enverus_Logfile'+str(today_date)+'.txt'

logging.basicConfig(filename=logfile, filemode='w',
                    format='%(asctime)s %(message)s')
logging.basicConfig(
    level=logging.INFO, 
    format='%(asctime)s [%(levelname)s] - %(message)s',
    filename=logfile)

logger = logging.getLogger()
logger.setLevel(logging.INFO)
logging.info('setting paTH TO DOWNLOAD')
path = os.getcwd() + "\\Download"
logging.info('SETTING PROFILE SETTINGS FOR FIREFOX')


profile = webdriver.FirefoxProfile()
profile.set_preference('browser.download.folderList', 2)
profile.set_preference('browser.download.dir', path)
profile.set_preference('browser.download.useDownloadDir', True)
profile.set_preference('browser.download.viewableInternally.enabledTypes', "")
profile.set_preference('browser.helperApps.neverAsk.saveToDisk','Portable Document Format (PDF), application/pdf')
profile.set_preference('pdfjs.disabled', True)
logging.info('Adding firefox profile')
driver=webdriver.Firefox(executable_path=GeckoDriverManager().install(),firefox_profile=profile)


site = 'https://biourja.sharepoint.com'
path1 = "/BiourjaPower/_api/web/GetFolderByServerRelativeUrl"
# path2= "Shared Documents/Vendor Research/Enverus(PRT)"
def remove_existing_files(files_location):
    logger.info("Inside remove_existing_files function")
    try:
        files = os.listdir(files_location)
        if len(files) > 0:
            for file in files:
                os.remove(files_location + "\\" + file)
            logger.info("Existing files removed successfully")
        else:
            print("No existing files available to reomve")

        print("Pause")
    except Exception as e:
        logger.info(e)

        raise e

def login():  
    '''This function downloads log in to the website'''
    try:
        logging.info('Accesing website')
        driver.get("https://outlook.office365.com/owa/biourja.com/")
        logging.info('providing id and passwords')
        WebDriverWait(driver, 90, poll_frequency=1).until(EC.element_to_be_clickable((By.ID, "i0116"))).send_keys(username)
        time.sleep(1)
        WebDriverWait(driver, 90, poll_frequency=1).until(EC.element_to_be_clickable((By.ID, "idSIButton9"))).click()
        time.sleep(1)
        WebDriverWait(driver, 90, poll_frequency=1).until(EC.element_to_be_clickable((By.ID, "i0118"))).send_keys(password)
        time.sleep(1)
        logging.info('click on No Button')
        driver.find_element_by_id("idSIButton9").click()
        time.sleep(1)
        driver.find_element_by_xpath('//*[@id="idBtn_Back"]').click()
        time.sleep(5)
        logging.info('Accessing search box')
        WebDriverWait(driver, 90, poll_frequency=1).until(EC.element_to_be_clickable((By.ID, "searchBoxId-Mail"))).click()
        time.sleep(5)
        logging.info('Clearing Search Bar')
        driver.find_element_by_xpath('/html/body/div[3]/div/div[1]/div/div[1]/div[2]/div/div/div/div/div[1]/div[2]/div/div/div/div/div[1]/div/div[2]/div/input').clear()
    except Exception as e:
        raise e
    

def Pjm_western_hub():
    try:
        logging.info('Searching with keywords')
        driver.find_element_by_xpath('/html/body/div[3]/div/div[1]/div/div[1]/div[2]/div/div/div/div/div[1]/div[2]/div/div/div/div/div[1]/div/div[2]/div/input').send_keys("PRT PJM")
        time.sleep(1)
        logging.info('Hitting search button')
        driver.find_element_by_xpath('/html/body/div[3]/div/div[1]/div/div[1]/div[2]/div/div/div/div/div[1]/div[2]/div/div/div/div/div[1]/button/span/i').click()
        time.sleep(1)
        logging.info('search for mail')
        driver.find_element_by_xpath("/html/body/div[3]/div/div[2]/div[2]/div/div/div/div[3]/div[2]/div/div[1]/div[2]/div/div/div/div/div/div[6]/div/div").click()
        time.sleep(10)
        logging.info('pdf download link')
        driver.find_element_by_partial_link_text('-Western Hub').click()
        time.sleep(30)
        logging.info('Switching tab')
        main_page = driver.window_handles[1] 
        driver.switch_to.window(main_page)
        logging.info('hitting download btn')
        driver.find_element_by_tag_name("a").click()   
        time.sleep(20)
        driver.close()
    except Exception as e:
        raise e

def Prt_ERCOT():
    try:
        logging.info('Switching tab')
        main_page = driver.window_handles[0] 
        driver.switch_to.window(main_page)
        logging.info('Clearing Search Bar')
        driver.find_element_by_xpath('/html/body/div[3]/div/div[1]/div/div[1]/div[2]/div/div/div/div/div[1]/div[2]/div/div/div/div/div[1]/div/div[2]/div/input').clear()
        logging.info('Searching with keywords')
        driver.find_element_by_xpath('/html/body/div[3]/div/div[1]/div/div[1]/div[2]/div/div/div/div/div[1]/div[2]/div/div/div/div/div[1]/div/div[2]/div/input').send_keys("PRT ERCOT")
        time.sleep(1)
        logging.info('Hitting search button')
        driver.find_element_by_xpath('/html/body/div[3]/div/div[1]/div/div[1]/div[2]/div/div/div/div/div[1]/div[2]/div/div/div/div/div[1]/button/span/i').click()
        time.sleep(1)
        logging.info('search for mail')
        driver.find_element_by_xpath("/html/body/div[3]/div/div[2]/div[2]/div/div/div/div[3]/div[2]/div/div[1]/div[2]/div/div/div/div/div/div[6]/div/div").click()        
        time.sleep(10)
        logging.info('pdf download link')
        driver.find_element_by_partial_link_text('-North').click()
        time.sleep(30)
        logging.info('Switching tab')
        main_page = driver.window_handles[1] 
        driver.switch_to.window(main_page)
        logging.info('hitting download btn')
        driver.find_element_by_tag_name("a").click()   
        time.sleep(20)
        driver.close()
    except Exception as e:
        raise e
def Prt_CAISO():
    try:
        logging.info('Switching tab')
        main_page = driver.window_handles[0] 
        driver.switch_to.window(main_page)
        logging.info('Clearing Search Bar')
        driver.find_element_by_xpath('/html/body/div[3]/div/div[1]/div/div[1]/div[2]/div/div/div/div/div[1]/div[2]/div/div/div/div/div[1]/div/div[2]/div/input').clear()
        logging.info('Searching with keywords')
        driver.find_element_by_xpath('/html/body/div[3]/div/div[1]/div/div[1]/div[2]/div/div/div/div/div[1]/div[2]/div/div/div/div/div[1]/div/div[2]/div/input').send_keys("PRT CA 15")
        time.sleep(1)
        logging.info('Hitting search button')
        driver.find_element_by_xpath('/html/body/div[3]/div/div[1]/div/div[1]/div[2]/div/div/div/div/div[1]/div[2]/div/div/div/div/div[1]/button/span/i').click()
        time.sleep(1)
        logging.info('search for mail')
        driver.find_element_by_xpath("/html/body/div[3]/div/div[2]/div[2]/div/div/div/div[3]/div[2]/div/div[1]/div[2]/div/div/div/div/div/div[6]/div/div").click()        
        time.sleep(10)
        logging.info('pdf download link')
        driver.find_element_by_partial_link_text('SP-15').click()
        time.sleep(30)
        logging.info('Switching tab')
        main_page = driver.window_handles[1] 
        driver.switch_to.window(main_page)
        logging.info('hitting download btn')
        driver.find_element_by_tag_name("a").click()   
        time.sleep(20)
        driver.quit()
    except Exception as e:
        raise e
def connect_to_sharepoint():
    try:
        site='https://biourja.sharepoint.com'
        username = os.getenv("user") if os.getenv("user") else sp_username
        password = os.getenv("password") if os.getenv("password") else sp_password
        # Connecting to Sharepoint and downloading the file with sync params
        s = sharepy.connect(site, username, password)
        return s
    except Exception as e:
        raise e



        

def shp_file_upload(s):
    try:
        filesToUpload = os.listdir(os.getcwd() + "\\Download")
        for fileToUpload in filesToUpload:
                
            headers = {"accept": "application/json;odata=verbose",
            "content-type": "Portable Document Format (PDF)"}

            with open(os.path.join(os.getcwd() + "\\Download", f'{fileToUpload}'), 'rb') as read_file:
                    content = read_file.read()
            if "PJM" in fileToUpload:
                    folder="PJMISO"
            elif "SP" in fileToUpload:
                    folder="CAISO"
            elif "ERCOT" in fileToUpload:
                    folder="ERCOT"
            else:
                    print("FILE NOT DOWNLOADED YET")        
            p = s.post(f"{site}{path1}('{share_point_path}/{folder}')/Files/add(url='{fileToUpload}',overwrite=true)", data=content, headers=headers)
                # url = f"https://biourja.sharepoint.com/_api/web/GetFolderByServerRelativeUrl('Shared Documents/Vendor Research/Enverus(PRT)/PJMISO')/Files/add(url='dummy.pdf',overwrite=true)"
                # r = s.post(url.format("C:/Users/Yashn.jain/Desktop/First_Project", "Enverus_PJM 90 Price Forecast 02-02-22T.pdf"), data=content, headers=headers)
            print(f'{fileToUpload} uploaded successfully')
    
        print(f'{job_name} executed succesfully')
    except Exception as e:
        raise e
def main():
    try:
        remove_existing_files(files_location)
        login()
        Pjm_western_hub()
        Prt_ERCOT()
        Prt_CAISO()
        s=connect_to_sharepoint()
        shp_file_upload(s)
        bu_alerts.send_mail(receiver_email = receiver_email,mail_subject =f'JOB SUCCESS - {job_name}',mail_body = f'{job_name} completed successfully, Attached logs',attachment_location = logfile)
    except Exception as e:
        logging.exception(str(e))
        bu_alerts.send_mail(receiver_email = receiver_email,mail_subject =f'JOB FAILED -{job_name}',mail_body = f'{job_name} failed, Attached logs',attachment_location = logfile)
            
    
if __name__ == "__main__":
    logging.info("Execution Started")
    time_start=time.time()
    files_location=os.getcwd() + "\\Download"
    credential_dict = get_config('ENVERUSPRT_EMAIL_FILES_AUTOMATION','ENVERUSPRT_EMAIL_FILES_AUTOMATION')
    username = credential_dict['USERNAME'].split(';')[0]
    password = credential_dict['PASSWORD'].split(';')[0]
    sp_username = credential_dict['USERNAME'].split(';')[1]
    sp_password =  credential_dict['PASSWORD'].split(';')[1]
    share_point_path = '/'.join(credential_dict['API_KEY'].split('/')[4:])
    # share_point_path = credential_dict['API_KEY'].split('/')[4:]
    receiver_email = credential_dict['EMAIL_LIST']
    # receiver_email = 'yashn.jain@biourja.com,mrutunjaya.sahoo@biourja.com'
    job_name='ENVERUSPRT_EMAIL_FILES_AUTOMATION'
    main()
    time_end=time.time()
    logging.info(f'It takes {time_start-time_end} seconds to run')

