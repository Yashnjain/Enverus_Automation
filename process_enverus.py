from selenium import webdriver
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.common.by import By
import time
from datetime import date,datetime
import logging
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.firefox import GeckoDriverManager
import sharepy
import os
from bu_config import get_config
import bu_alerts
import numpy as np


def get_f_list_from_sp(s, path):
    try:
        logging.info("Inside get_f_list_from_sp() function")
        # Get list of all files and folders in library
        # sharepoint_site
        r = s.get(sharepoint_site + """/BiourjaPower/_api/web/GetFolderByServerRelativeUrl('"""+path+"""')/Files""")
        files = r.json()['d']['results']
        f_lst = []
        for file in files:
            # print(file["Name"])
            f_lst.append(file["Name"])
        print("done")
        return f_lst
    except Exception as e:
        logging.info(f"Error in get_f_list_from_sp() function:{e}")
        raise e
def remove_existing_files(files_location):
    """_summary_

    Args:
        files_location (_type_): _description_

    Raises:
        e: _description_
    """           
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
        logging.info('Accesing website')
        driver.get(url)
        logging.info('providing id and passwords')
        WebDriverWait(driver, 90, poll_frequency=1).until(EC.element_to_be_clickable((By.ID, "i0116"))).send_keys(username)
        time.sleep(1)
        WebDriverWait(driver, 90, poll_frequency=1).until(EC.element_to_be_clickable((By.ID, "idSIButton9"))).click()
        time.sleep(1)
        WebDriverWait(driver, 90, poll_frequency=1).until(EC.element_to_be_clickable((By.ID, "i0118"))).send_keys(password)
        time.sleep(1)
        logging.info('click on No Button')
        WebDriverWait(driver, 90, poll_frequency=1).until(EC.element_to_be_clickable((By.ID, "idSIButton9"))).click()
        time.sleep(1)
        WebDriverWait(driver, 90, poll_frequency=1).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="idBtn_Back"]'))).click()
        time.sleep(5)
        retry=0
        while retry < 10:
            try:
                logging.info('closing unwanted overlay')
                try:
                    WebDriverWait(driver, 90, poll_frequency=1).until(EC.element_to_be_clickable((By.CSS_SELECTOR, ".ms-Dialog-button"))).click()
                except:
                    pass
                time.sleep(5)
                logging.info('Accessing search box')
                WebDriverWait(driver, 90, poll_frequency=1).until(EC.element_to_be_clickable((By.ID, "searchBoxId-Mail"))).click()
                time.sleep(5)
                logging.info("setting search for only inbox")
                WebDriverWait(driver, 90, poll_frequency=1).until(EC.element_to_be_clickable((By.XPATH,"//span[@id='searchScopeButtonId-option']"))).click()
                time.sleep(10)
                WebDriverWait(driver, 90, poll_frequency=1).until(EC.element_to_be_clickable((By.XPATH,"//span[@data-automationid='splitbuttonprimary']//span[contains(text(),'Inbox')]"))).click()
                break
            except Exception as e:
                time.sleep(5)
                retry+=1
                if retry ==10:
                    raise e 
        logging.info('Clearing Search Bar')
        search_bar=WebDriverWait(driver, 90, poll_frequency=1).until(EC.element_to_be_clickable((By.CSS_SELECTOR,"input[placeholder='Search']")))
        search_bar.clear()
        return search_bar,driver
    except Exception as e:
        logging.info(f"Error in login() function:{e}")
        raise e
    
def download_files(search_bar,driver):
    """_summary_

    Raises:
        e: _description_
    """    
    try:
        logging.info("Inside download_files() function")
        dict1={"PRT 90-Day PJM Price Forecast":'-Western Hub',"PRT 90-Day ERCOT Price Forecast":'-North',"PRT 90-Day CAISO SP-15":'SP-15'}
        for key, value in dict1.items():
            logging.info('Switching tab')
            main_page = driver.window_handles[0] 
            driver.switch_to.window(main_page)
            logging.info('Clearing Search Bar')
            search_bar.clear()
            logging.info('Searching with keywords')
            search_bar.send_keys(key)
            time.sleep(1)
            logging.info('Hitting search button')
            WebDriverWait(driver, 90, poll_frequency=1).until(EC.element_to_be_clickable((By.XPATH, "//button[@aria-label='Search']//span[@data-automationid='splitbuttonprimary']"))).click()
            time.sleep(1)
            try:
                logging.info('search for mail')
                recent_mail=WebDriverWait(driver, 150, poll_frequency=1).until(EC.element_to_be_clickable(
                    (By.XPATH,'/html/body/div[1]/div/div[2]/div/div[2]/div[2]/div/div/div/div[3]/div/div[2]/div[1]/div[1]\
                    /div/div/div/div/div/div/div/div[5]/div/div[2]/div/div/div[2]')))
                time.sleep(5)
                recent_mail.click()
            except Exception as e:
                logging.info(f"Error in lselecting recent mail")
                raise e
            time.sleep(10)
            logging.info('pdf download link')
            WebDriverWait(driver, 90, poll_frequency=1).until(EC.element_to_be_clickable((By.PARTIAL_LINK_TEXT, value))).click()
            time.sleep(30)
            logging.info(f'checking opened windows for {key} file')
            windows_opened=driver.window_handles
            try:
                if len(windows_opened)==2:
                    logging.info(f'Two windows are opened for {key} file')
                    logging.info(f'Switching tab for downloading {key} pdf file')
                    main_page = driver.window_handles[1] 
                    driver.switch_to.window(main_page)
                    logging.info('hitting download btn')
                    time.sleep(15)
                    WebDriverWait(driver, 90, poll_frequency=1).until(EC.element_to_be_clickable((By.TAG_NAME, "a"))).click()
                    time.sleep(40)
                    logging.info(f'{key} Got downloaded Successfully')
                    driver.close()
                else:
                    logging.info(f'Single window opened for {key} file')
                    logging.info(f'{key} Got downloaded Successfully')
                    print("Only one opened")
            except Exception as e:
                raise e             
    except Exception as e:
        logging.info(f"Error in login() function:{e}")
        raise e
    finally:
        driver.quit()

def connect_to_sharepoint():
    """_summary_

    Raises:
        e: _description_

    Returns:
        _type_: _description_
    """    
    try:
        logging.info("Inside connect_to_sharepoint() function")
        username = os.getenv("user") if os.getenv("user") else sp_username
        password = os.getenv("password") if os.getenv("password") else sp_password
        # Connecting to Sharepoint and downloading the file with sync params
        s = sharepy.connect(sharepoint_site, username, password)
        return s
    except Exception as e:
        logging.info(f"Error in connect_to_sharepoint() function:{e}")
        raise e

def shp_file_check(s):
    try:
        logging.info("Inside shp_file_check() function")
        # temp=False        
        folder_list=["PJMISO","CAISO","ERCOT"]
        count=0  
        for items in folder_list:
            count+=1
            checking_list= f'{sharepoint_path_2}/{items}'          
            f_lst = get_f_list_from_sp(s, checking_list)
            filesToUpload = os.listdir(os.getcwd() + "\\download")
            print("goint into the loop")
            for item in filesToUpload:
                print(item)
                if item in f_lst:
                    global receiver_email
                    receiver_email = credential_dict['EMAIL_LIST'].split(';')[1]  
                    # temp=True         
        # return receiver_email            
    except Exception as e:
        logging.info(f"Error in shp_file_check() function:{e}")
        raise e    
def shp_file_upload(s):
    """_summary_

    Args:
        s (_type_): _description_

    Raises:
        e: _description_

    Returns:
        _type_: _description_
    """    
    filesToUpload = os.listdir(os.getcwd() + "\\download")
    try:
        logging.info("Inside shp_file_upload() function")
        global body
        body = ''
        for fileToUpload in filesToUpload:
            z=path+'\\'+fileToUpload
            locations_list.append(z)    
            headers = {"accept": "application/json;odata=verbose",
            "content-type": "Portable Document Format (PDF)"}

            with open(os.path.join(os.getcwd() + "\\download", f'{fileToUpload}'), 'rb') as read_file:
                    content = read_file.read()
            if "PJM" in fileToUpload:
                folder="PJMISO"
            elif "SP" in fileToUpload:
                folder="CAISO"
            elif "ERCOT" in fileToUpload:
                folder="ERCOT"
            else:
                print("FILE NOT DOWNLOADED YET")    
            p = s.post(f"{sharepoint_site}{sharepoint_path_1}('{sharepoint_path_2}{folder}/')/Files/add(url='{fileToUpload}',overwrite=true)", data=content, headers=headers)   
            #p = s.post(f"{site}{path1}('{share_point_path}/{folder}')/Files/add(url='{fileToUpload}',overwrite=true)", data=content, headers=headers)
            # url = f"https://biourja.sharepoint.com/_api/web/GetFolderByServerRelativeUrl('Shared Documents/Vendor Research/Enverus(PRT)/PJMISO')/Files/add(url='dummy.pdf',overwrite=true)"
            # r = s.post(url.format("C:/Users/Yashn.jain/Desktop/First_Project", "Enverus_PJM 90 Price Forecast 02-02-22T.pdf"), data=content, headers=headers)
            nl = '<br>'
            body += (f'{nl}<strong>{folder}</strong> {nl}{nl} {fileToUpload} successfully uploaded in {folder}, {nl} Attached link for the same=<a href ="{share_point_path}/\\{folder}">{folder}</a>{nl}')
            print(f'{fileToUpload} uploaded successfully')
        print(f'{job_name} executed succesfully')
        return locations_list
    except Exception as e:
        logging.info(f"Error in shp_file_upload() function:{e}")
        raise e

def main():
    try:
        logging.info("Inside main() function")
        no_of_rows=0
        Database=""
        log_json='[{"JOB_ID": "'+str(job_id)+'","CURRENT_DATETIME": "'+str(datetime.now())+'"}]'
        bu_alerts.bulog(process_name=processname,database=Database,status='Started',table_name='',
            row_count=no_of_rows, log=log_json, warehouse='ITPYTHON_WH',process_owner=process_owner)
        remove_existing_files(files_location)
        search_bar,driver=login()
        download_files(search_bar,driver)
        s=connect_to_sharepoint()
        shp_file_check(s)
        shp_file_upload(s) 
        locations_list.append(logfile)
        bu_alerts.bulog(process_name=processname,database=Database,status='Completed',table_name='',
            row_count=no_of_rows, log=log_json, warehouse='ITPYTHON_WH',process_owner=process_owner)  
        bu_alerts.send_mail(receiver_email = receiver_email,mail_subject =f'JOB SUCCESS - {job_name}',mail_body = f'{body}{job_name} completed successfully, Attached PDF and Logs',multiple_attachment_list = locations_list)
    except Exception as e:
        logging.info(f"Error in main() function:{e}")
        log_json='[{"JOB_ID": "'+str(job_id)+'","CURRENT_DATETIME": "'+str(datetime.now())+'"}]'
        bu_alerts.bulog(process_name= processname,database=Database,status='Failed',table_name='',
            row_count=no_of_rows, log=log_json, warehouse='ITPYTHON_WH',process_owner=process_owner)
        logging.exception(str(e))
        locations_list.append(logfile)
        bu_alerts.send_mail(receiver_email = receiver_email,mail_subject =f'JOB FAILED -{job_name}',mail_body = f'{job_name} failed, Attached logs',multiple_attachment_list = locations_list)
               
if __name__ == "__main__": 
    try:
        logging.info("Execution Started")
        time_start=time.time()
        #Global VARIABLES

        locations_list=[]
        body = ''
        today_date=date.today()
        # log progress --
        for handler in logging.root.handlers[:]:
            logging.root.removeHandler(handler)
        # logfile = os.getcwd() +"\\logs\\"+'Enverus_Logfile'+str(today_date)+'.txt'

        logfile = os.getcwd() + '\\' + 'logs' + '\\' + 'Enverus_Log_{}.txt'.format(str(today_date))

        logging.basicConfig(
            level=logging.INFO, 
            format='%(asctime)s [%(levelname)s] - %(message)s',
            filename=logfile)

        logger = logging.getLogger()
        logger.setLevel(logging.INFO)
        logging.info('setting paTH TO download')
        path = os.getcwd() + '\\download'

        credential_dict = get_config('ENVERUSPRT_EMAIL_FILES_AUTOMATION','ENVERUSPRT_EMAIL_FILES_AUTOMATION')
        username = credential_dict['USERNAME'].split(';')[0]
        password = credential_dict['PASSWORD'].split(';')[0]
        sp_username = credential_dict['USERNAME'].split(';')[1]
        sp_password =  credential_dict['PASSWORD'].split(';')[1]
        url=credential_dict['SOURCE_URL']
        sharepoint_site=credential_dict['API_KEY'].split(';')[0]
        sharepoint_path_1=credential_dict['API_KEY'].split(';')[1]
        sharepoint_path_2=credential_dict['API_KEY'].split(';')[2]
        share_point_path = f'{sharepoint_site}/BiourjaPower/{sharepoint_path_2}'
        # receiver_email='enoch.benjamin@biourja.com,bhavana.kaurav@biourja.com'
        receiver_email = credential_dict['EMAIL_LIST'].split(';')[0]
        directories_created=["download","Logs"]
        for directory in directories_created:
            path3 = os.path.join(os.getcwd(),directory)  
            try:
                os.makedirs(path3, exist_ok = True)
                print("Directory '%s' created successfully" % directory)
            except OSError as error:
                print("Directory '%s' can not be created" % directory)       
        files_location=os.getcwd() + "\\download"
        filesToUpload = os.listdir(os.getcwd() + "\\download")
        job_name=credential_dict['PROJECT_NAME']
        job_id=np.random.randint(1000000,9999999)
        processname = credential_dict['PROJECT_NAME']
        process_owner = credential_dict['IT_OWNER']
        main()
        time_end=time.time()
        logging.info(f'It takes {time_start-time_end} seconds to run')
    except Exception as e:
        logging.exception(str(e))
        bu_alerts.send_mail(receiver_email = receiver_email,mail_subject =f'JOB FAILED -{job_name}',mail_body = f'{job_name} failed in __main__, Attached logs',attachment_location = logfile)
    



