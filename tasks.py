from robocorp.tasks import task
from robocorp import browser
from RPA.Browser.Selenium import Selenium

from RPA.HTTP import HTTP

import csv

from RPA.PDF import PDF

import os
import time
import zipfile

@task
def robot_manufacturing():
    browser.configure(
        slowmo=100,   
        browser_engine="chrome", 
    )

    open_the_other_intranet_website()
    download_excel_file()
    iterate_csv()

    create_zip_from_files_with_prefix('output', 'robot_results_', 'output.zip')
    
def open_the_other_intranet_website():
    browser.goto("https://robotsparebinindustries.com/#/robot-order")

def download_excel_file():
    """Downloads excel file from the given URL"""
    http = HTTP()
    http.download(url="https://robotsparebinindustries.com/orders.csv", overwrite=True)

 
def iterate_csv():
    data = read_csv_to_table()
    for row in data:
        try:
            order_robot(row)
        except Exception as e:
            browser.goto('about:blank')
            open_the_other_intranet_website()
            print(f"An error occurred: {e}")

def read_csv_to_table() -> list:
    csv_file_path = 'orders.csv'
    with open(csv_file_path, mode='r', encoding='utf-8') as file:
        reader = csv.DictReader(file)
        data = [row for row in reader]
        return data
   

def order_robot(row):
    page = browser.page()
    page.locator("xpath=//button[contains(.,'OK')]").click()
    page.locator("xpath=//select[@id='head']").select_option(str(row["Head"]))
    page.locator("xpath=//input[@id='id-body-"+str(row["Body"])+"']").check()
    page.locator("xpath=//label[contains(.,'3. Legs:')]/../input").fill(str(row["Legs"]))
    page.locator("xpath=//input[@id='address']").fill(str(row["Address"]))
    page.locator("xpath=//button[@id='preview']").click()
    collect_results()
    page.locator("xpath=//button[@id='order']").click()
    export_as_pdf(str(row["Order number"]))
    page.locator("xpath=//button[@id='order-another']").click()


def wait_for_download():

    file_path = 'orders.csv'
    timeout = 30
    start_time = time.time()
    while not os.path.exists(file_path):
        time.sleep(1)  
        if time.time() - start_time > timeout:
            print(f"Timed out waiting for file '{file_path}'.")
            break
    else:
        print(f"The file '{file_path}' has been found.")

def collect_results():
    """Take a screenshot of the page"""
    page = browser.page()
    page.locator("xpath=//div[@id='robot-preview-image']").screenshot(path="output/robot_summary.png")

def export_as_pdf(name):
    page = browser.page()
    robot_results_html = page.locator("xpath=//div[@id='receipt']").inner_html()
    pdf = PDF()
    robot_results_html= robot_results_html+"<img src='output/robot_summary.png'>"
    pdf.html_to_pdf(robot_results_html, f"output/robot_results_{name}.pdf")    


def create_zip_from_files_with_prefix(directory, prefix, zip_filename):
    with zipfile.ZipFile(zip_filename, 'w', zipfile.ZIP_DEFLATED) as zipf:
        for root, dirs, files in os.walk(directory):
            for file in files:
                if file.startswith(prefix):
                    full_path = os.path.join(root, file)
                    zipf.write(full_path, arcname=file)

