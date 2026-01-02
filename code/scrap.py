from selenium import webdriver
from selenium.webdriver.common.by import By
import time , os ,shutil
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.select import Select
from selenium.common.exceptions import NoSuchElementException, TimeoutException
from datetime import datetime , timedelta
import requests , zipfile
import pandas as pd
from dotenv import load_dotenv
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import smtplib , requests , email , imaplib
from email.header import decode_header
from selenium.webdriver.common.alert import Alert
from selenium.webdriver.support.ui import Select
from lxml import html
from retrying import retry
import csv
import logging
Options = webdriver.ChromeOptions()
Options.add_argument('--headless=new')
Options.add_argument('--ignore-certificate-errors')
Options.add_argument('--ignore-ssl-errors')
Options.add_argument('--no-sandbox')
Options.add_argument('--disable-dev-shm-usage')
Options.add_experimental_option("excludeSwitches",["enable-automation"])
Options.page_load_strategy = 'normal'
driver = webdriver.Chrome(service = Service(ChromeDriverManager().install()), options = Options)
driver.maximize_window()
driver.delete_all_cookies()
load_dotenv()
url = os.environ.get('url')
url_user = os.environ.get('url_user')
url_pwd = os.environ.get('url_pwd')
mail_from = os.environ.get('mail_from')
from_addr = os.environ.get('from_addr')
to_addr = os.environ.get('to_addr').split(',')
ccaddr = os.environ.get('ccaddr').split(',')
mail_pwd = os.environ.get('mail_pwd')
server_name = os.environ.get('server_name')
port = os.environ.get('port')
login_mail_id=os.environ.get('login_mail_id')
login_mail_pwd=os.environ.get('login_mail_pwd')
# Logger setup
logging.basicConfig(
	filename='logs/tvsautomation.log',
	level=logging.INFO,
	format='%(asctime)s %(levelname)s:%(message)s'
)
logger = logging.getLogger(__name__)
def login():
    global user
    if url != '':
        response = requests.head(url)
        if response.status_code == 200:
            driver.get(url)
        else:
            import url_issue
            driver.close()
            exit()
    url_user = os.environ.get('url_user')
    url_pwd = os.environ.get('url_pwd')
    users = url_user.split(',')
    passwords = url_pwd #.split(',')
    for i, user in enumerate(users):
        password = passwords #passwords[0] if i < 3 else passwords[1]
        try:
            driver.get(url)
        except ConnectionError:
            import user_issue
            driver.close()
            exit()
        try:
            driver.find_element(By.NAME, 'vendor_id').send_keys(user)
            driver.find_element(By.NAME, 'password').send_keys(password)
            select_role = Select(driver.find_element(By.NAME, 'role'))
            select_role.select_by_visible_text('Vendor')
            driver.find_element(By.XPATH, '/html/body/app-root/app-vlogin/div/div/div[2]/div[1]/div[6]/input').click()
        except NoSuchElementException as e:
            print(f"Login error for user {user}: {e}")
            import login_issue
            driver.close()
            exit()
        time.sleep(5)
        try:
            tvs_fun()
            user_processed = True
        except Exception as e:
            print(f"Error during tvs_fun execution for user {user}: {e}")
            driver.close()
            exit()
        if user_processed:
            print(f"Successfully processed user: {user}")
        else:
            print(f"Failed to process user: {user}")
def tvs_fun():
        try:  
            if driver.find_element(By.XPATH,'//*[@id="mat-dialog-1"]/app-popupvendorpdf/div/div/div/div[1]/button').click():
                print('firstclose')
            elif  driver.find_element(By.XPATH , '//*[@id="mat-dialog-0"]/app-surveypopup/div/div/div/div[1]/button/span').click():
                print('nextpage')
        except Exception as e:
            import login_issue
            exit()
        time.sleep(1)
        driver.find_element(By.LINK_TEXT,'Delivery').click()
        time.sleep(1)
        driver.find_element(By.LINK_TEXT,'Invoice Upload').click()
        time.sleep(1)
        window_handles = driver.window_handles
        for handle in window_handles:
            driver.switch_to.window(handle)
            title_2 = driver.title
        driver.find_element(By.NAME,'Button2').click()
        yesterday = datetime.today() - timedelta(days=1)
        yesterday_day = yesterday.day
        yesterday_month = yesterday.strftime('%B')
        yesterday_year = yesterday.year
        yesterdays_date = yesterday_month + " " + str(yesterday_year)
        time.sleep(3)
        first_date = driver.find_element(By.ID,'sdate').click()
        time.sleep(.25)
        driver.find_element(By.CLASS_NAME, 'datepicker-switch').text
        #if current_month != yesterdays_date :
        date_element = driver.find_element(By.XPATH, f"//td[text()='{yesterday_day}']")
        date_element.click()
        driver.find_element(By.ID,'edate').click()
        time.sleep(.25)
        driver.find_element(By.XPATH,'/html/body/div/div[1]/table/thead/tr[1]/th[2]').text
        yesterday = datetime.today() - timedelta(days=1)
        yesterday_day = yesterday.day
        #if current_month != yesterdays_date :
        driver.find_element(By.XPATH, f"//td[text()='{yesterday_day}']").click()
        time.sleep(.5)
        dropdown = driver.find_element(By.NAME, "drploc")
        select = Select(dropdown)
        select.select_by_visible_text("Successful")
        try:
            download_data()
        except Exception as e:
            print(f"Error at download  ")
                
def download_data():
    time.sleep(4)
    driver.find_element(By.CSS_SELECTOR,'#btnsrch').click()   ##btnsrch
    time.sleep(3)
    """
    alert = Alert(driver)
    confirm_box = driver.switch_to.alert
    confirm_box.accept()
    """
    data = driver.find_element(By.ID,'invoicedata').text
    data = data.split('\n')
    data = data[10]
    try:
        if data == 'No Data':
            print("No Data")
        else:
            table = driver.find_element(By.TAG_NAME, 'table')
            headers = [header.text for header in table.find_elements(By.TAG_NAME, 'th')]
            headers = headers[:-1]
            table_data = table.find_elements(By.TAG_NAME, 'tbody')
            all_data = []
            for tbody in table_data:
                rows = tbody.find_elements(By.TAG_NAME, 'tr')
                for row in rows:
                    cells = row.find_elements(By.TAG_NAME, 'td')
                    row_data = [cell.text for cell in cells]
                    row_data = row_data[:-2]
                    all_data.append(row_data)
            df = pd.DataFrame(all_data, columns=headers)
            now = datetime.now()
            day_no  =  now.strftime("%j") 
            week_no =  now.strftime("%U") 
            year = now.strftime("%Y")
            df.to_csv(fr"\tvsbot\tvs_portal_files\{user}_{day_no}_{week_no}_{year}.csv",index = False)
    except Exception as e:
            print(e)
    time.sleep(0.5)
@retry
def extract_mail_table():
    while True:
        try:
            mail = imaplib.IMAP4_SSL(server_name)
            mail.login(login_mail_id, login_mail_pwd)
            mail.select("inbox")
            result, data = mail.search(None, '(UNSEEN FROM "itd@lgb.co.in" SUBJECT "TVS Yesterdays sales")')
            if result == "OK":
                email_ids = data[0].split()
                for email_id in email_ids:
                    result, message_data = mail.fetch(email_id, "(RFC822)")
                    raw_email = message_data[0][1]
                    msg = email.message_from_bytes(raw_email)
                    subject, encoding = decode_header(msg["Subject"])[0]
                    if isinstance(subject, bytes):
                        subject = subject.decode(encoding or "utf-8")
                    body = ""
                    if msg.is_multipart():
                        for part in msg.walk():
                            content_type = part.get_content_type()
                            content_disposition = str(part.get("Content-Disposition"))
                            if "attachment" not in content_disposition:
                                if content_type == "text/plain":
                                    body += part.get_payload(decode=True).decode('utf-8')
                                elif content_type == "text/html":
                                    body += part.get_payload(decode=True).decode('utf-8')
                    else:
                        body = msg.get_payload(decode=True).decode('utf-8')
                    tree = html.fromstring(body)
                    tables = tree.xpath('//table')
                    all_table_data = []
                    for table in tables:
                        print("Table HTML:", html.tostring(table, pretty_print=True).decode())
                        headers = []
                        header_row = table.xpath('.//tr/th')
                        if header_row:
                            headers = [th.text_content().strip() for th in header_row]
                        table_data = []
                        rows = table.xpath('.//tr')
                        for row in rows:
                            columns = row.xpath('.//td')
                            row_data = [col.text_content().strip() for col in columns]
                            if row_data:
                                table_data.append(row_data)
                        if headers:
                            all_table_data.append(headers)
                        all_table_data.extend(table_data)
                        with open(r'\tvsbot\report\email_table_data.csv', 'w', newline='', encoding='utf-8') as file:
                            writer = csv.writer(file)
                            for table_data in all_table_data:
                                if isinstance(table_data, list) and len(table_data) > 0 and isinstance(table_data[0], list):
                                    writer.writerow(['New Table'])
                                writer.writerow(table_data)
                break
            #mail.close()
            #mail.logout()
        except Exception as e:
            print(e)
            mail.logout()

def compare_files():
    try:
        # Load the DataFrames from CSV files
        folder_path = r'\tvsbot\tvs_portal_files'
        file_list = [file for file in os.listdir(folder_path) if file.endswith('.csv')]
        combined_data = pd.DataFrame()
        for file in file_list:
            file_path = os.path.join(folder_path, file)
            data = pd.read_csv(file_path)
            combined_data = pd.concat([combined_data, data], ignore_index=True)
            combined_data.to_csv(r'\tvsbot\combined_invoice\combined_data.csv', index = False)
        mail_data = pd.read_csv(r'\tvsbot\report\email_table_data.csv' , skiprows = 1)
        col_rename = mail_data.rename(columns = {'CEX/NEX InvoiceNo':'InvoiceNo'})
        mail_data_df = pd.DataFrame(col_rename)
        compare_column = 'InvoiceNo'
        try:
            diff_df = mail_data_df[~mail_data_df[compare_column].isin(combined_data[compare_column])]
            diff_df.to_csv(r'\tvsbot\result_set\difference_records.csv' , index = False)
            if compare_column not in combined_data.columns or compare_column not in mail_data_df.columns:
                raise KeyError(f"Column '{compare_column}' does not exist in both DataFrames.")
            common_df = combined_data.merge(mail_data_df, on=compare_column, how='inner')
            common_df.to_csv(r'\tvsbot\result_set\common_records.csv', index= False)
        except KeyError as e:
            print(f"Error: {e}")
    except Exception as e:
        print(f"An unexpected error occurred: {e}")
        
if __name__ == '__main__':
    if login():
        logger.info("Succeffully Logged TVS Portal")
        if extract_mail_table():
            logger.info("TVS Yesterday's Report EXtracted Succesfully")
            if compare_files():
                logger.info("Data Manipulated")
                def zip_directory(folder_path, output_path):
                    with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
                        for root, dirs, files in os.walk(folder_path):
                            logger.info("files",files)
                            for file in files:
                                file_path = os.path.join(root, file)
                                zipf.write(file_path, os.path.relpath(file_path, folder_path))
                folder_path = r'\tvsbot\tvs_portal_files'
                folder_to_zip = r'\tvsbot\result_set'
                zip_filename = r'\tvsbot\mail_folder\tvs_report.zip'
                zip_directory(folder_to_zip, zip_filename)
                subject = "Reg : TVS Report"
                body = """<!DOCTYPE html>
                    <html lang="en">
                    <head>
                        <meta charset="UTF-8">
                        <meta name="viewport" content="width=device-width, initial-scale=1.0">
                        <title>Welcome Email</title>
                        <style>
                            body {
                                font-family: Arial, sans-serif;
                                background-color: #f4f4f4;
                                margin: 0;
                                padding: 0;
                                color: #333;
                            }
                            .email-container {
                                background-color: #ffffff;
                                padding: 20px;
                                margin: 20px auto;
                                max-width: 600px;
                                border: 1px solid #ddd;
                                border-radius: 5px;
                            }

                            .header {
                                background-color: #4CAF50;
                                color: white;
                                padding: 10px;
                                text-align: center;
                                border-radius: 5px 5px 0 0;
                            }

                            .content {
                                padding: 20px;
                            }

                            .footer {
                                background-color: #f4f4f4;
                                color: #666;
                                text-align: center;
                                padding: 10px;
                                font-size: 12px;
                                border-top: 1px solid #ddd;
                                border-radius: 0 0 5px 5px;
                            }
                        </style>
                    </head>
                    <body>
                        <div class="email-container">
                            <div class="header">
                                <h1>Welcome to IT Service!</h1>
                            </div>
                            <div class="content">
                                <caption>Sir/Madam,</caption><br>
                                    <p> This is common mail report for updation of TVS INVOICE </p><br>
                            </div>
                            <div class="content">
                            <div class="row">
                                <div class="col-12">
                                    <caption>Note:</caption>
                                    <ul class="ms-0">
                                        <li>This is an autogenerated mail. Please do not respond to this.</li>
                                    </ol>
                                </div>
                            </div>
                            <div class="footer">
                                <p>&copy; 2025 LGB. All rights reserved.</p>
                            </div>
                        </div>
                    </body>
                    </html>"""
                
                message = MIMEMultipart()
                message["From"] = mail_from
                message["To"] = ','.join(to_addr)
                message['Cc'] = ','.join(ccaddr)
                message["Subject"] = subject
                message.attach(MIMEText(body, "html"))
                with open(zip_filename, "rb") as attachment:
                    part = MIMEBase("application", "octet-stream")
                    part.set_payload(attachment.read())
                encoders.encode_base64(part)
                part.add_header(
                    "Content-Disposition",
                    f"attachment; filename= {os.path.basename(zip_filename)}",)
                message.attach(part)
                text = message.as_string()
                all_rec = to_addr + ccaddr
                with smtplib.SMTP(server_name, port) as server:
                    server.starttls() 
                    server.login(from_addr, mail_pwd)
                    server.sendmail(mail_from, all_rec , text)
                    server.quit()
                
                def move_files(src_folder1, src_folder2,src_folder3,src_folder4,src_folder5, dest_folder):
                        now = datetime.now()
                        day_no = now.strftime("%j") 
                        week_no = now.strftime("%U") 
                        year = now.strftime("%Y")
                        new_folder_name = f"day_{day_no}_week_{week_no}_year_{year}"
                        new_dest_folder = os.path.join(dest_folder, new_folder_name)
                        if not os.path.exists(new_dest_folder):
                            os.makedirs(new_dest_folder)
                        for filename in os.listdir(src_folder1):
                            src_path = os.path.join(src_folder1, filename)
                            if os.path.isfile(src_path):
                                shutil.move(src_path, os.path.join(new_dest_folder, filename))
                                print(f"Moved: {src_path} to {new_dest_folder}")
                        for filename in os.listdir(src_folder2):
                            src_path = os.path.join(src_folder2, filename)
                            if os.path.isfile(src_path):
                                shutil.move(src_path, os.path.join(new_dest_folder, filename))
                                print(f"Moved: {src_path} to {new_dest_folder}")
                        for filename in os.listdir(src_folder3):
                            src_path = os.path.join(src_folder3, filename)
                            if os.path.isfile(src_path):
                                shutil.move(src_path, os.path.join(new_dest_folder, filename))
                                print(f"Moved: {src_path} to {new_dest_folder}")
                        for filename in os.listdir(src_folder4):
                            src_path = os.path.join(src_folder4, filename)
                            if os.path.isfile(src_path):
                                shutil.move(src_path, os.path.join(new_dest_folder, filename))
                                print(f"Moved: {src_path} to {new_dest_folder}")
                        for filename in os.listdir(src_folder5):
                            src_path = os.path.join(src_folder5, filename)
                            if os.path.isfile(src_path):
                                shutil.move(src_path, os.path.join(new_dest_folder, filename))
                                print(f"Moved: {src_path} to {new_dest_folder}")
                source_folder1 = r'\tvsbot\combined_invoice'
                source_folder2 = r'\tvsbot\tvs_portal_files'
                source_folder3 = r'\tvsbot\report'
                source_folder4 = r'\tvsbot\result_set'
                source_folder5 = r'\tvsbot\mail_folder'
                destination_folder = r'\tvsbot\old_files'
                move_files(source_folder1, source_folder2,source_folder3,source_folder4,source_folder5, destination_folder)
