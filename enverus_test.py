from selenium import webdriver
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.common.by import By
import time
from datetime import date,datetime
import logging
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.firefox.options import Options
from webdriver_manager.firefox import GeckoDriverManager
import sharepy
import os
from bu_config import get_config
import bu_alerts
import numpy as np


def get_f_list_from_sp(s, path):
  # Get list of all files and folders in library
  r = s.get(site + """/BiourjaPower/_api/web/GetFolderByServerRelativeUrl('"""+path+"""')/Files""")
  files = r.json()['d']['results']
  f_lst = []
  for file in files:
    # print(file["Name"])
    f_lst.append(file["Name"])
  print("done")
  return f_lst
  
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
        print('Exception caught during execution remove_existing_files() : {}'.format(str(e)))
        logging.exception('Exception caught during remove_existing_files() : {}'.format(str(e)))
        raise e

def login():  
    '''This function downloads log in to the website'''
    try:
        logging.info('Accesing website')
        driver.get("https://outlook.office365.com/owa/biourja.com/")
        logging.info('providing id and passwords')
        WebDriverWait(driver, 90, poll_frequency=1).until(EC.element_to_be_clickable((By.ID, "i0116"))).send_keys(username)
        time.sleep(5)
        WebDriverWait(driver, 90, poll_frequency=1).until(EC.element_to_be_clickable((By.ID, "idSIButton9"))).click()
        time.sleep(5)
        WebDriverWait(driver, 90, poll_frequency=1).until(EC.element_to_be_clickable((By.ID, "i0118"))).send_keys(password)
        time.sleep(5)
        logging.info('click on No Button')
        WebDriverWait(driver, 90, poll_frequency=1).until(EC.element_to_be_clickable((By.ID, "idSIButton9"))).click()
        time.sleep(5)
        WebDriverWait(driver, 90, poll_frequency=1).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="idBtn_Back"]'))).click()
        time.sleep(5)
        retry=0
        while retry < 10:
            try:
                logging.info('closing unwanted overlay')
                try:
                    WebDriverWait(driver, 90, poll_frequency=1).until(EC.element_to_be_clickable((By.CSS_SELECTOR, ".ms-Dialog-button"))).click()
                except:
                    pass          #it passses no error in the try block.....#it tells not to raise any exception in execpt block.
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
        #search_bar=WebDriverWait(driver, 90, poll_frequency=1).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[2]/div/div[1]/div/div[1]/div[2]/div/div/div/div/div[1]/div[2]/div/div/div/div/div[1]/div/div[2]/div/input')))
        search_bar=WebDriverWait(driver, 90, poll_frequency=1).until(EC.element_to_be_clickable((By.XPATH,'//*[@id="topSearchInput"]')))
        search_bar.clear()
        return search_bar
    except Exception as e:
        print('Exception caught during execution login() : {}'.format(str(e)))
        logging.exception('Exception caught during login() : {}'.format(str(e)))
        raise e

    
def download_files(search_bar):
    """_summary_

    Raises:
        e: _description_
    """    
    try:
        dict1={"PRT PJM":'-Western Hub',"PRT ERCOT":'-North',"PRT CA 15":'SP-15'}
        for key, value in dict1.items():
            logging.info('Switching tab')
            main_page = driver.window_handles[0]      #driver.window_handles[0] is the method to store the present link in the window handle.If we open another link then we can store that link in another window handle as driver.window_handle[1] 
            driver.switch_to.window(main_page)       #driver.switvh_to.window() switches to the respected tab we need to make operations by passing the variable name in the arguement.
            logging.info('Clearing Search Bar')
            search_bar.clear()
            logging.info('Searching with keywords')
            search_bar.send_keys(key)
            time.sleep(1)
            logging.info('Hitting search button')
            # WebDriverWait(driver, 90, poll_frequency=1).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div/div[1]/div/div[1]\
            #                             /div[2]/div/div/div/div/div/div[1]/div[2]/div/div/div/div/div[1]/div/div[3]/button/span/i/span/i"))).click()
            #WebDriverWait(driver, 90, poll_frequency=1).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[2]/div/div[1]/div/div[1]/div[2]/div/div/div/div/div[1]/div[2]/div/div/div/div/div[1]/button/span/i'))).click()
            time.sleep(1)
            logging.info('search for mail')
            WebDriverWait(driver, 90, poll_frequency=1).until(EC.element_to_be_clickable((By.XPATH, "//button[@aria-label='Search']//span[@data-automationid='splitbuttonprimary']"))).click()        
            time.sleep(10)
            WebDriverWait(driver, 90, poll_frequency=1).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div/div[2]/div[2]/div[2]/div/div/div/div[3]/div/div/div[1]/div[2]/div/div/div/div/div/div[6]/div/div"))).click()
            time.sleep(10)
            logging.info('pdf download link')
            WebDriverWait(driver, 90, poll_frequency=1).until(EC.element_to_be_clickable((By.PARTIAL_LINK_TEXT, value))).click()
            time.sleep(30)
            try:
                logging.info('Switching tab')
                multi_window = driver.window_handles
                length = len(multi_window)
                if length == 2:
                    main_page = driver.window_handles[1] 
                    driver.switch_to.window(main_page)
                    logging.info('hitting download btn')
                    time.sleep(10)
                    WebDriverWait(driver, 90, poll_frequency=1).until(EC.element_to_be_clickable((By.TAG_NAME, "a"))).click()
                    time.sleep(40)
                    driver.close()
                else:
                    print(f'{key} Downloaded')     
            except ValueError as e:
                print(f'{key} Downloaded') 



    except Exception as e:
        print('Exception caught during execution download_files() : {}'.format(str(e)))
        logging.exception('Exception caught during download_files() : {}'.format(str(e)))
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
        site='https://biourja.sharepoint.com'
        username = os.getenv("user") if os.getenv("user") else sp_username
        password = os.getenv("password") if os.getenv("password") else sp_password
        # Connecting to Sharepoint and downloading the file with sync params
        s = sharepy.connect(site, username, password)
        return s
    except Exception as e:
        print('Exception caught during execution connect_to_sharepoint() : {}'.format(str(e)))
        logging.exception('Exception caught during connect_to_sharepoint() : {}'.format(str(e)))
        raise e

def shp_file_check(s):
    # temp=False        
    folder_list=["PJMISO","CAISO","ERCOT"]
    count=0  
    for items in folder_list:
        count+=1
        path3= "Shared%20Documents/Vendor Research/Enverus(PRT)/"
        checking_list= f'{path3}/{items}'          
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
            p = s.post(f"{site}{path1}('{share_point_path}/{folder}')/Files/add(url='{fileToUpload}',overwrite=true)", data=content, headers=headers)
                # url = f"https://biourja.sharepoint.com/_api/web/GetFolderByServerRelativeUrl('Shared Documents/Vendor Research/Enverus(PRT)/PJMISO')/Files/add(url='dummy.pdf',overwrite=true)"
                # r = s.post(url.format("C:/Users/Yashn.jain/Desktop/First_Project", "Enverus_PJM 90 Price Forecast 02-02-22T.pdf"), data=content, headers=headers)
            nl = '<br>'
            body += (f'{nl}<strong>{folder}</strong> {nl}{nl} {fileToUpload} successfully uploaded in {folder}, {nl} Attached link for the same=<a href ="{temp_path}\{folder}">{folder}</a>{nl}')
            print(f'{fileToUpload} uploaded successfully')
        print(f'{job_name} executed succesfully')
        return locations_list
    except Exception as e:
        print('Exception caught during execution shp_file_upload() : {}'.format(str(e)))
        logging.exception('Exception caught during shp_file_upload() : {}'.format(str(e)))
        raise e

def main():
    try:
        no_of_rows=0
        Database=""
        log_json='[{"JOB_ID": "'+str(job_id)+'","CURRENT_DATETIME": "'+str(datetime.now())+'"}]'
        bu_alerts.bulog(process_name=processname,database=Database,status='Started',table_name='',
            row_count=no_of_rows, log=log_json, warehouse='ITPYTHON_WH',process_owner=process_owner)
        remove_existing_files(files_location)
        search_bar=login()
        download_files(search_bar)
        s=connect_to_sharepoint()
        shp_file_check(s)
        shp_file_upload(s) 
        locations_list.append(logfile)
        bu_alerts.bulog(process_name=processname,database=Database,status='Completed',table_name='',
            row_count=no_of_rows, log=log_json, warehouse='ITPYTHON_WH',process_owner=process_owner)  
        
        bu_alerts.send_mail(
                # receiver_email= 'priyanka.solanki@biourja.com,radha.waswani@biourja.com',
                receiver_email = receiver_email,
                mail_subject = f'JOB SUCCESS - {job_name}',
                mail_body = f'{job_name} completed successfully, Attached logs',
                multiple_attachment_list = locations_list
            )
        # bu_alerts.send_mail(receiver_email = receiver_email,
        #                     mail_subject =f'JOB SUCCESS - {job_name}',
        #                     mail_body = f'{body}{job_name} completed successfully, Attached PDF and Logs',
        #                     multiple_attachment_list = locations_list
                            # )
    except Exception as e:
        log_json='[{"JOB_ID": "'+str(job_id)+'","CURRENT_DATETIME": "'+str(datetime.now())+'"}]'
        bu_alerts.bulog(process_name= processname,database=Database,status='Failed',table_name='',
            row_count=no_of_rows, log=log_json, warehouse='ITPYTHON_WH',process_owner=process_owner)
        logging.exception(str(e))
        locations_list.append(logfile)
        bu_alerts.send_mail(receiver_email = receiver_email,mail_subject =f'JOB FAILED -{job_name}',mail_body = f'{job_name} failed, Attached logs',multiple_attachment_list = logfile)
               
if __name__ == "__main__": 
    try:
        logging.info("Execution Started")
        time_start=time.time()
        #Global VARIABLES

        locations_list=[]
        body = ''
        site = 'https://biourja.sharepoint.com'
        path1 = "/BiourjaPower/_api/web/GetFolderByServerRelativeUrl"
        # path2= "Shared Documents/Vendor Research/Enverus(PRT)"

        today_date=date.today()
        # log progress --
        for handler in logging.root.handlers[:]:
            logging.root.removeHandler(handler)
        # logfile = os.getcwd() +"\\logs\\"+'Enverus_Logfile'+str(today_date)+'.txt'

        logfile = os.getcwd() + '\\' + 'logs' + '\\' + 'Enverus_Log_{}.txt'.format(str(today_date))   # { Enverus_Log_2022-10-12.txt  }
        # if os.path.isfile(logfile):
        #     os.remove(logfile)
        logging.basicConfig(
            level=logging.INFO, 
            format='%(asctime)s [%(levelname)s] - %(message)s',
            filename=logfile)

        logger = logging.getLogger()
        logger.setLevel(logging.INFO)
        logging.info('setting paTH TO download')
        path = os.getcwd() + '\\download'
        logging.info('SETTING PROFILE SETTINGS FOR FIREFOX')


        options = Options()
        mime_types = ['application/pdf', 'text/plain', 'application/vnd.ms-excel', 'text/csv', 'application/csv', 'text/comma-separated-values','application/download', 'application/octet-stream', 'binary/octet-stream', 'application/binary', 'application/x-unknown','attachment/csv','application/vnd.openxmlformats-officedocument.spreadsheetml.sheet']
        options.set_preference("browser.download.folderList", 2)
        options.set_preference("browser.download.manager.showWhenStarting", True)
        options.set_preference("browser.download.dir", path)
        options.set_preference("browser.helperApps.neverAsk.saveToDisk",",".join(mime_types))
        options.set_preference("browser.helperApps.neverAsk.openFile", "application/pdf, application/octet-stream, application/x-winzip, application/x-pdf, application/x-gzip")
        options.set_preference("pdfjs.disabled", True)
        # browser = webdriver.Firefox(firefox_profile=fp, options=options, executable_path="D:\\python_practice\geckodriver.exe")
        print(os.getcwd())
        executable_path = 'geckodriver.exe'
        print(executable_path)
        driver = webdriver.Firefox(options=options, executable_path=executable_path)




        # profile = webdriver.FirefoxProfile()
        # profile.set_preference('browser.download.folderList', 2)
        # profile.set_preference('browser.download.dir', path)
        # profile.set_preference('browser.download.useDownloadDir', True)
        # profile.set_preference('browser.download.viewableInternally.enabledTypes', "")
        # profile.set_preference('browser.helperApps.neverAsk.saveToDisk','Portable Document Format (PDF), application/pdf')
        # profile.set_preference('pdfjs.disabled', True)
        # logging.info('Adding firefox profile')
        # driver=webdriver.Firefox(executable_path=GeckoDriverManager().install(),firefox_profile=profile)

        credential_dict = get_config('ENVERUSPRT_EMAIL_FILES_AUTOMATION','ENVERUSPRT_EMAIL_FILES_AUTOMATION')
        username = credential_dict['USERNAME'].split(';')[0]
        password = credential_dict['PASSWORD'].split(';')[0]
        sp_username = credential_dict['USERNAME'].split(';')[1]
        sp_password =  credential_dict['PASSWORD'].split(';')[1]
        share_point_path = '/'.join(credential_dict['API_KEY'].split('/')[4:])
        temp_path = credential_dict['API_KEY']
        #receiver_email = 'enoch.benjamin@biourja.com'
        receiver_email = credential_dict['EMAIL_LIST'].split(';')[0]
        directories_created=["download","Logs"]
        for directory in directories_created:
            path3 = os.path.join(os.getcwd(),directory)  
            try:
                os.makedirs(path3, exist_ok = True)          #makedies: is a method of creating a directory, and exit_ok = True, tells weather the dir is existed or not but to create the new one
                print("Directory '%s' created successfully" % directory)
            except OSError as error:
                print("Directory '%s' can not be created" % directory)       
        files_location=os.getcwd() + "\\download"
        filesToUpload = os.listdir(os.getcwd() + "\\download")        #os.listdir() prints all the files from the cwd and we pass any arguement as folder name from which folder the files should be printed ...here we given "downloads as a folder name"
        # share_point_path = credential_dict['API_KEY'].split('/')[4:]
        
        # receiver_email='yashn.jain@biourja.com'
        job_name='TEST:ENVERUSPRT_EMAIL_FILES_AUTOMATION'#credential_dict['PROJECT_NAME']
        job_id=np.random.randint(1000000,9999999)
        processname = credential_dict['PROJECT_NAME']
        process_owner = credential_dict['IT_OWNER']
        main()
        time_end=time.time()             #gives the current time in seconds.
        logging.info(f'It takes {time_start-time_end} seconds to run')
    except Exception as e:
        logging.exception(str(e))
        bu_alerts.send_mail(receiver_email = receiver_email,mail_subject =f'JOB FAILED -{job_name}',mail_body = f'{job_name} failed in __main__, Attached logs',attachment_location = logfile)
    

