import os
import time
import csv
import shutil
import logging
import zipfile
import smtplib
import email
import imaplib
import pandas as pd
import requests
from datetime import datetime, timedelta
from dotenv import load_dotenv
from email import encoders
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.header import decode_header
from lxml import html
from retrying import retry

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select, WebDriverWait
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.common.exceptions import WebDriverException, NoSuchElementException
from webdriver_manager.chrome import ChromeDriverManager
import chromedriver_autoinstaller

# -------------------- SELENIUM SETUP ----------------------
chrome_options = Options()
chrome_options.add_argument("--ignore-certificate-errors")
chrome_options.add_argument("--ignore-ssl-errors")
chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
chrome_options.page_load_strategy = "normal"

chromedriver_autoinstaller.install()

driver = webdriver.Chrome(
    service=Service(ChromeDriverManager().install()),
    options=chrome_options
)

driver.maximize_window()
driver.get_cookies()

# -------------------- ENV SETUP ----------------------
load_dotenv()
if "__file__" in globals():
    os.chdir(os.path.dirname(os.path.abspath(__file__)))

load_dotenv(r"D:\tvsbot\.env")

# -------------------- LOGGER ----------------------
logging.basicConfig(
    filename=r'logs\tvsautomation.log',
    level=logging.INFO,
    format='%(asctime)s %(levelname)s:%(message)s'
)
logger = logging.getLogger(__name__)


# ============================================================
#                          tvsbot CLASS
# ============================================================
class tvsbot:
    def __init__(self, driver):
        self.driver = driver
        self.wait = WebDriverWait(self.driver, 25)

        self.url = os.getenv("url")
        self.users = os.getenv("url_user").split(",")
        self.password = os.getenv("url_pwd")

        self._init_mailconfig()
        self._init_directories()

    def _init_directories(self):
        self.result_set = r"D:\tvsbot\result_set"
        self.tvs_portal_files = r"D:\tvsbot\tvs_portal_files"
        self.combined_invoice = r"D:\tvsbot\combined_invoice"
        self.report = r"D:\tvsbot\report"
        self.mail_folder = r"D:\tvsbot\mail_folder"
        self.old_files = r"D:\tvsbot\old_files"

        # Fix wrong path usage
        self.zip_filename = r"D:\tvsbot\report\report.zip"

    def _init_mailconfig(self):
        self.mail_from = os.getenv("mail_from")
        self.from_addr = os.getenv("from_addr")

        self.to_addr = os.getenv("to_addr").split(",")
        self.ccaddr = os.getenv("ccaddr").split(",")

        self.mail_pwd = os.getenv("mail_pwd")
        self.login_mail_id = os.getenv("login_mail_id")
        self.login_mail_pwd = os.getenv("login_mail_pwd")
        self.server_name = os.getenv("server_name")
        self.port = int(os.getenv("port"))

    # =======================================================
    #                           LOGIN
    # =======================================================
    def login(self):
        try:
            res = requests.head(self.url)
            if res.status_code != 200:
                driver.quit()
                return False
        except:
            driver.quit()
            return False

        for user in self.users:
            try:
                driver.get(self.url)
                driver.find_element(By.NAME, 'vendor_id').send_keys(user)
                driver.find_element(By.NAME, 'password').send_keys(self.password)

                Select(driver.find_element(By.NAME, 'role')).select_by_visible_text('Vendor')
                driver.find_element(By.XPATH, '//input[@value="Login"]').click()

                time.sleep(5)
                self.user = user

                self.tvs_portal_filesun()
                print(f"Processed user: {user}")

            except Exception as e:
                print(f"Login Error ({user}): {e}")
                return False

        return True

    # =======================================================
    #                 PORTAL PROCESSING FUNCTION
    # =======================================================
    def tvs_portal_filesun(self):
        try:
            try:
                driver.find_element(By.XPATH, '//*[@id="mat-dialog-1"]//button').click()
            except:
                pass

            try:
                driver.find_element(By.XPATH, '//*[@id="mat-dialog-0"]//button').click()
            except:
                pass

        except Exception:
            return False

        time.sleep(1)
        driver.find_element(By.LINK_TEXT, 'Delivery').click()
        driver.find_element(By.LINK_TEXT, 'Invoice Upload').click()
        time.sleep(1)

        driver.find_element(By.NAME, 'Button2').click()

        yesterday = datetime.today() - timedelta(days=1)
        day = yesterday.day

        # Start date
        driver.find_element(By.ID, 'sdate').click()
        driver.find_element(By.XPATH, f"//td[text()='{day}']").click()

        # End date
        driver.find_element(By.ID, 'edate').click()
        driver.find_element(By.XPATH, f"//td[text()='{day}']").click()

        Select(driver.find_element(By.NAME, "drploc")).select_by_visible_text("Successful")

        self.download_data()

    # =======================================================
    #                       DOWNLOAD DATA
    # =======================================================
    def download_data(self):
        driver.find_element(By.CSS_SELECTOR, '#btnsrch').click()
        time.sleep(2)

        try:
            table = driver.find_element(By.TAG_NAME, 'table')

            headers = [th.text for th in table.find_elements(By.TAG_NAME, 'th')][:-1]

            rows = table.find_elements(By.XPATH, ".//tbody/tr")
            all_rows = []

            for row in rows:
                tds = row.find_elements(By.TAG_NAME, 'td')
                row_data = [x.text for x in tds][:-2]
                all_rows.append(row_data)

            df = pd.DataFrame(all_rows, columns=headers)

            now = datetime.now()
            filename = fr"D:\tvsbot\tvs_portal_files\{self.user}_{now.strftime('%j_%U_%Y')}.csv"
            df.to_csv(filename, index=False)

        except Exception as e:
            print("Download Error:", e)

    # =======================================================
    #                EMAIL TABLE EXTRACTION
    # =======================================================
    @retry(stop_max_attempt_number=3, wait_fixed=3000)
    def extract_mail_table(self):
        try:
            mail = imaplib.IMAP4_SSL(self.server_name)
            mail.login(self.login_mail_id, self.login_mail_pwd)
            mail.select("inbox")

            result, data = mail.search(None,
                '(UNSEEN FROM "itd@lgb.co.in" SUBJECT "TVS Yesterdays sales")'
            )

            if result != "OK":
                return False

            email_ids = data[0].split()
            if not email_ids:
                return False

            email_id = email_ids[-1]

            _, message_data = mail.fetch(email_id, "(RFC822)")
            msg = email.message_from_bytes(message_data[0][1])

            body = ""
            for part in msg.walk():
                if part.get_content_type() == "text/html":
                    body += part.get_payload(decode=True).decode()

            tree = html.fromstring(body)
            tables = tree.xpath("//table")

            csv_path = os.path.join(self.mail_folder, "email_table_data.csv")

            with open(csv_path, "w", newline="", encoding="utf-8") as f:
                writer = csv.writer(f)

                for table in tables:
                    for row in table.xpath(".//tr"):
                        row_data = [td.text_content().strip() for td in row.xpath(".//td")]
                        if row_data:
                            writer.writerow(row_data)

            return True

        except Exception as e:
            print("Mail extract error:", e)
            return False

    # =======================================================
    #                   COMPARE DATA FILES
    # =======================================================
    def compare_files(self):

        try:
            files = [f for f in os.listdir(self.tvs_portal_files) if f.endswith(".csv")]
            combined = pd.concat(
                [pd.read_csv(os.path.join(self.tvs_portal_files, f)) for f in files],
                ignore_index=True
            )

            mail_file = os.path.join(self.mail_folder, "email_table_data.csv")
            mail_df = pd.read_csv(mail_file)

            if "CEX/NEX InvoiceNo" in mail_df.columns:
                mail_df.rename(columns={"CEX/NEX InvoiceNo": "InvoiceNo"}, inplace=True)

            diff = mail_df[~mail_df["InvoiceNo"].isin(combined["InvoiceNo"])]
            diff.to_csv(os.path.join(self.result_set, "diff.csv"), index=False)

            comm = combined.merge(mail_df, on="InvoiceNo", how="inner")
            comm.to_csv(os.path.join(self.result_set, "comm.csv"), index=False)

            return True

        except Exception as e:
            print("Compare Files Error:", e)
            return False

    # =======================================================
    #                       SEND MAIL
    # =======================================================
    def send_mail(self):

        def zip_dir(source, dest):
            with zipfile.ZipFile(dest, "w", zipfile.ZIP_DEFLATED) as z:
                for root, _, files in os.walk(source):
                    for f in files:
                        p = os.path.join(root, f)
                        z.write(p, os.path.relpath(p, source))

        zip_dir(self.result_set, self.zip_filename)

        msg = MIMEMultipart()
        msg["From"] = self.mail_from
        msg["To"] = ",".join(self.to_addr)
        msg["Cc"] = ",".join(self.ccaddr)
        msg["Subject"] = "Reg: TVS Report"

        msg.attach(MIMEText("<p>Attached Report</p>", "html"))

        with open(self.zip_filename, "rb") as f:
            part = MIMEBase("application", "octet-stream")
            part.set_payload(f.read())
            encoders.encode_base64(part)
            part.add_header("Content-Disposition", f"attachment; filename=report.zip")
            msg.attach(part)

        all_recipients = self.to_addr + self.ccaddr

        try:
            with smtplib.SMTP(self.server_name, self.port) as s:
                s.starttls()
                s.login(self.from_addr, self.mail_pwd)
                s.sendmail(self.mail_from, all_recipients, msg.as_string())
            return True
        except Exception as e:
            print("Mail send error:", e)
            return False

    # =======================================================
    #               MOVE OLD FILES TO ARCHIVE
    # =======================================================
    def move_files(self):
        now = datetime.now()
        folder = f"day_{now.strftime('%j')}_week_{now.strftime('%U')}_year_{now.strftime('%Y')}"

        dest = os.path.join(self.old_files, folder)
        os.makedirs(dest, exist_ok=True)

        for path in [
            self.combined_invoice,
            self.tvs_portal_files,
            self.report,
            self.result_set
        ]:
            for f in os.listdir(path):
                shutil.move(os.path.join(path, f), os.path.join(dest, f))


# =========================================================
#                           MAIN
# =========================================================
def main():
    try:
        bot = tvsbot(driver)

        if bot.login():
            logger.info("Login successful")

            if bot.extract_mail_table():
                logger.info("Mail extracted")

                if bot.compare_files():
                    logger.info("Comparison done")

                    if bot.send_mail():
                        logger.info("Mail sent")
                        bot.move_files()

    except WebDriverException as e:
        logger.exception("WebDriver Error: %s", e)

    finally:
        driver.quit()


if __name__ == "__main__":
    main()
