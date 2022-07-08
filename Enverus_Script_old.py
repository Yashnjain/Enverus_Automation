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
import smtplib
import email.mime.multipart
import email.mime.text
import email.mime.base
import email.encoders as encoders


#Global VARIABLES

locations_list=[]
body = ''
site = 'https://biourja.sharepoint.com'
path1 = "/BiourjaPower/_api/web/GetFolderByServerRelativeUrl"
# path2= "Shared Documents/Vendor Research/Enverus(PRT)"

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
path = os.getcwd() + "\\"+"Download"
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

credential_dict = get_config('ENVERUSPRT_EMAIL_FILES_AUTOMATION','ENVERUSPRT_EMAIL_FILES_AUTOMATION')
username = credential_dict['USERNAME'].split(';')[0]
password = credential_dict['PASSWORD'].split(';')[0]
sp_username = credential_dict['USERNAME'].split(';')[1]
sp_password =  credential_dict['PASSWORD'].split(';')[1]
share_point_path = '/'.join(credential_dict['API_KEY'].split('/')[4:])
temp_path = credential_dict['API_KEY']
receiver_email = credential_dict['EMAIL_LIST'].split(';')[0]

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
  
def send_mail(receiver_email: str, mail_subject: str, mail_body: str, attachment_locations: list = None, sender_email: str = None, sender_password: str=None) -> bool:
    """The Function responsible to do all the mail sending logic.

    Args:
        sender_email (str): Email Id of the sender.
        sender_password (str): Password of the sender.
        receiver_email (str): Email Id of the receiver.
        mail_subject (str): Subject line of the email.
        mail_body (str): Message body of the Email.
        attachment_locations (list, optional): Absolute path of the attachment. Defaults to None.

    Returns:
        bool: [description]
    """
    logging.info("INTO THE SEND MAIL FUNCTION")
    done = False
    try:
        logging.info("GIVING CREDENTIALS FOR SENDING MAIL")
        if not sender_email or sender_password:
            sender_email = "biourjapowerdata@biourja.com"
            sender_password = r"Texas08642"
            # sender_email = r"virtual-out@biourja.com"
            # sender_password = "t?%;`p39&Pv[L<6Y^cz$z2bn"
        receivers = receiver_email.split(",")
        msg = email.mime.multipart.MIMEMultipart()
        msg['From'] = "biourjapowerdata@biourja.com"
        msg['To'] = receiver_email
        msg['Subject'] = mail_subject
        body = mail_body
        logging.info("Attaching mail body")
        msg.attach(email.mime.text.MIMEText(body, 'html'))
        logging.info("Attching files in the mail")
        for files_locations in attachment_locations:
            with open(files_locations, 'r+b') as attachment:
                # instance of MIMEBase and named as p
                p = email.mime.base.MIMEBase('application', 'octet-stream')
                # To change the payload into encoded form
                p.set_payload((attachment).read())
                encoders.encode_base64(p)  # encode into base64
                p.add_header('Content-Disposition',
                             "attachment; filename= %s" % files_locations)
                msg.attach(p)  # attach the instance 'p' to instance 'msg'

        # s = smtplib.SMTP('smtp.gmail.com', 587) # creates SMTP session
        s = smtplib.SMTP('smtp.office365.com',
                         587)  # creates SMTP session
        s.starttls()  # start TLS for security
        s.login(sender_email, sender_password)  # Authentication
        text = msg.as_string()  # Converts the Multipart msg into a string

        s.sendmail(sender_email, receivers, text)  # sending the mail
        s.quit()  # terminating the session
        done = True
        logging.info("Email sent successfully")
        print("Email sent successfully.")
    except Exception as e:
        print(
            f"Could not send the email, error occured, More Details : {e}")
    finally:
        return done

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
        WebDriverWait(driver, 90, poll_frequency=1).until(EC.element_to_be_clickable((By.ID, "idSIButton9"))).click()
        time.sleep(1)
        WebDriverWait(driver, 90, poll_frequency=1).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="idBtn_Back"]'))).click()
        time.sleep(5)
        retry=0
        while retry < 10:
            try:
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
        WebDriverWait(driver, 90, poll_frequency=1).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[3]/div/div[1]/div/div[1]/div[2]/div/div/div/div/div[1]/div[2]/div/div/div/div/div[1]/div/div[2]/div/input'))).clear()
        # driver.find_element_by_xpath('/html/body/div[3]/div/div[1]/div/div[1]/div[2]/div/div/div/div/div[1]/div[2]/div/div/div/div/div[1]/div/div[2]/div/input').clear()
    except Exception as e:
        raise e
    
def Pjm_western_hub():
    """_summary_

    Raises:
        e: _description_
    """    
    try:
        logging.info('Searching with keywords')
        WebDriverWait(driver, 90, poll_frequency=1).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[3]/div/div[1]/div/div[1]/div[2]/div/div/div/div/div[1]/div[2]/div/div/div/div/div[1]/div/div[2]/div/input'))).send_keys("PRT PJM")
        # driver.find_element_by_xpath('/html/body/div[3]/div/div[1]/div/div[1]/div[2]/div/div/div/div/div[1]/div[2]/div/div/div/div/div[1]/div/div[2]/div/input').send_keys("PRT PJM")
        time.sleep(1)
        logging.info('Hitting search button')
        WebDriverWait(driver, 90, poll_frequency=1).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[3]/div/div[1]/div/div[1]/div[2]/div/div/div/div/div[1]/div[2]/div/div/div/div/div[1]/button/span/i'))).click()
        # driver.find_element_by_xpath('/html/body/div[3]/div/div[1]/div/div[1]/div[2]/div/div/div/div/div[1]/div[2]/div/div/div/div/div[1]/button/span/i').click()
        time.sleep(1)
        logging.info('search for mail')
        WebDriverWait(driver, 90, poll_frequency=1).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[3]/div/div[2]/div[2]/div[2]/div/div/div/div[3]/div/div/div[1]/div[2]/div/div/div/div/div/div[6]/div/div"))).click()
        # driver.find_element_by_xpath("/html/body/div[3]/div/div[2]/div[2]/div/div/div/div[3]/div[2]/div/div[1]/div[2]/div/div/div/div/div/div[6]/div/div").click()
        time.sleep(10)
        logging.info('pdf download link')
        WebDriverWait(driver, 90, poll_frequency=1).until(EC.element_to_be_clickable((By.PARTIAL_LINK_TEXT, '-Western Hub'))).click()
        # driver.find_element_by_partial_link_text('-Western Hub').click()
        time.sleep(30)
        logging.info('Switching tab')
        main_page = driver.window_handles[1] 
        driver.switch_to.window(main_page)
        logging.info('hitting download btn')
        time.sleep(10)
        WebDriverWait(driver, 90, poll_frequency=1).until(EC.element_to_be_clickable((By.TAG_NAME, "a"))).click()
        # driver.find_element_by_tag_name("a").click()   
        time.sleep(40)
        driver.close()
    except Exception as e:
        raise e

def Prt_ERCOT():
    """_summary_

    Raises:
        e: _description_
    """    
    try:
        logging.info('Switching tab')
        main_page = driver.window_handles[0] 
        driver.switch_to.window(main_page)
        logging.info('Clearing Search Bar')
        WebDriverWait(driver, 90, poll_frequency=1).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[3]/div/div[1]/div/div[1]/div[2]/div/div/div/div/div[1]/div[2]/div/div/div/div/div[1]/div/div[2]/div/input'))).clear()
        # driver.find_element_by_xpath('/html/body/div[3]/div/div[1]/div/div[1]/div[2]/div/div/div/div/div[1]/div[2]/div/div/div/div/div[1]/div/div[2]/div/input').clear()
        logging.info('Searching with keywords')
        WebDriverWait(driver, 90, poll_frequency=1).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[3]/div/div[1]/div/div[1]/div[2]/div/div/div/div/div[1]/div[2]/div/div/div/div/div[1]/div/div[2]/div/input'))).send_keys("PRT ERCOT")
        # driver.find_element_by_xpath('/html/body/div[3]/div/div[1]/div/div[1]/div[2]/div/div/div/div/div[1]/div[2]/div/div/div/div/div[1]/div/div[2]/div/input').send_keys("PRT ERCOT")
        time.sleep(1)
        logging.info('Hitting search button')
        WebDriverWait(driver, 90, poll_frequency=1).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[3]/div/div[1]/div/div[1]/div[2]/div/div/div/div/div[1]/div[2]/div/div/div/div/div[1]/button/span/i'))).click()
        # driver.find_element_by_xpath('/html/body/div[3]/div/div[1]/div/div[1]/div[2]/div/div/div/div/div[1]/div[2]/div/div/div/div/div[1]/button/span/i').click()
        time.sleep(1)
        logging.info('search for mail')
        WebDriverWait(driver, 90, poll_frequency=1).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[3]/div/div[2]/div[2]/div[2]/div/div/div/div[3]/div/div/div[1]/div[2]/div/div/div/div/div/div[6]/div/div"))).click()
        # driver.find_element_by_xpath("/html/body/div[3]/div/div[2]/div[2]/div/div/div/div[3]/div[2]/div/div[1]/div[2]/div/div/div/div/div/div[6]/div/div").click()        
        time.sleep(10)
        logging.info('pdf download link')
        WebDriverWait(driver, 90, poll_frequency=1).until(EC.element_to_be_clickable((By.PARTIAL_LINK_TEXT,'-North' ))).click()
        # driver.find_element_by_partial_link_text('-North').click()
        time.sleep(30)
        logging.info('Switching tab')
        main_page = driver.window_handles[1] 
        driver.switch_to.window(main_page)
        logging.info('hitting download btn')
        time.sleep(10)
        WebDriverWait(driver, 90, poll_frequency=1).until(EC.element_to_be_clickable((By.TAG_NAME, "a"))).click() 
        # driver.find_element_by_tag_name("a").click()   
        time.sleep(40)
        driver.close()
    except Exception as e:
        raise e

def Prt_CAISO():
    try:
        logging.info('Switching tab')
        main_page = driver.window_handles[0] 
        driver.switch_to.window(main_page)
        logging.info('Clearing Search Bar')
        WebDriverWait(driver, 90, poll_frequency=1).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[3]/div/div[1]/div/div[1]/div[2]/div/div/div/div/div[1]/div[2]/div/div/div/div/div[1]/div/div[2]/div/input'))).clear()
        # driver.find_element_by_xpath('/html/body/div[3]/div/div[1]/div/div[1]/div[2]/div/div/div/div/div[1]/div[2]/div/div/div/div/div[1]/div/div[2]/div/input').clear()
        logging.info('Searching with keywords')
        WebDriverWait(driver, 90, poll_frequency=1).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[3]/div/div[1]/div/div[1]/div[2]/div/div/div/div/div[1]/div[2]/div/div/div/div/div[1]/div/div[2]/div/input'))).send_keys("PRT CA 15")
        # driver.find_element_by_xpath('/html/body/div[3]/div/div[1]/div/div[1]/div[2]/div/div/div/div/div[1]/div[2]/div/div/div/div/div[1]/div/div[2]/div/input').send_keys("PRT CA 15")
        time.sleep(1)
        logging.info('Hitting search button')
        WebDriverWait(driver, 90, poll_frequency=1).until(EC.element_to_be_clickable((By.XPATH,'/html/body/div[3]/div/div[1]/div/div[1]/div[2]/div/div/div/div/div[1]/div[2]/div/div/div/div/div[1]/button/span/i'))).click()
        # driver.find_element_by_xpath('/html/body/div[3]/div/div[1]/div/div[1]/div[2]/div/div/div/div/div[1]/div[2]/div/div/div/div/div[1]/button/span/i').click()
        time.sleep(1)
        logging.info('search for mail')
        WebDriverWait(driver, 90, poll_frequency=1).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[3]/div/div[2]/div[2]/div[2]/div/div/div/div[3]/div/div/div[1]/div[2]/div/div/div/div/div/div[6]/div/div"))).click()
        # driver.find_element_by_xpath("/html/body/div[3]/div/div[2]/div[2]/div/div/div/div[3]/div[2]/div/div[1]/div[2]/div/div/div/div/div/div[6]/div/div").click()        
        time.sleep(10)
        logging.info('pdf download link')
        WebDriverWait(driver, 90, poll_frequency=1).until(EC.element_to_be_clickable((By.PARTIAL_LINK_TEXT,'SP-15'))).click()
        # driver.find_element_by_partial_link_text('SP-15').click()
        time.sleep(30)
        logging.info('Switching tab')
        main_page = driver.window_handles[1] 
        driver.switch_to.window(main_page)
        logging.info('hitting download btn')
        time.sleep(10)
        WebDriverWait(driver, 90, poll_frequency=1).until(EC.element_to_be_clickable((By.TAG_NAME,"a"))).click()
        # driver.find_element_by_tag_name("a").click()   
        time.sleep(40)
        driver.quit()
    except Exception as e:
        raise e

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
        raise e

def shp_file_check(s):
    # temp=False        
    folder_list=["PJMISO","CAISO","ERCOT"]
    count=0  
    for items in folder_list:
        count+=1
        path3= "Shared%20Documents/Vendor Research/Enverus(PRT)/"
        Checking_list= f'{path3}/{items}'          
        f_lst = get_f_list_from_sp(s, Checking_list)
        filesToUpload = os.listdir(os.getcwd() + "\\Download")
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
    filesToUpload = os.listdir(os.getcwd() + "\\Download")
    try:
        global body
        body = ''
        for fileToUpload in filesToUpload:
            z=path+'\\'+fileToUpload
            locations_list.append(z)    
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
            nl = '<br>'
            body += (f'{nl}<strong>{folder}</strong> {nl}{nl} {fileToUpload} successfully uploaded in {folder}, {nl} Attached link for the same={temp_path}\{folder}{nl}')
            print(f'{fileToUpload} uploaded successfully')
        print(f'{job_name} executed succesfully')
        return locations_list
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
        shp_file_check(s)
        shp_file_upload(s)
        
        locations_list.append(logfile)
        send_mail(receiver_email = receiver_email,mail_subject =f'JOB SUCCESS - {job_name}',mail_body = f'{body}{job_name} completed successfully, Attached PDF and Logs',attachment_locations = locations_list)
    except Exception as e:
        logging.exception(str(e))
        bu_alerts.send_mail(receiver_email = receiver_email,mail_subject =f'JOB FAILED -{job_name}',mail_body = f'{job_name} failed, Attached logs',attachment_location = logfile)
               
if __name__ == "__main__": 
    logging.info("Execution Started")
    time_start=time.time()
    directories_created=["Download","Logs"]
    for directory in directories_created:
        path3 = os.path.join(os.getcwd(),directory)  
        try:
            os.makedirs(path3, exist_ok = True)
            print("Directory '%s' created successfully" % directory)
        except OSError as error:
            print("Directory '%s' can not be created" % directory)       
    files_location=os.getcwd() + "\\Download"
    filesToUpload = os.listdir(os.getcwd() + "\\Download")
    # share_point_path = credential_dict['API_KEY'].split('/')[4:]
    
    # receiver_email='yashn.jain@biourja.com'
    job_name='ENVERUS_PRT_EMAIL_FILES_AUTOMATION'
    main()
    time_end=time.time()
    logging.info(f'It takes {time_start-time_end} seconds to run')

