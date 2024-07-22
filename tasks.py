from RPA.PDF import PDF
from robocorp.tasks import task
from robocorp import browser
from RPA.HTTP import HTTP
from RPA.Excel.Files import Files
from RPA.Tables import Tables
from RPA.Outlook.Application import Application
from robot.api import logger
from RPA.Archive import Archive
import csv
import os

@task
def order_robots_from_RobotSpareBin():
    """
    Orders robots from RobotSpareBin Industries Inc.
    Saves the order HTML receipt as a PDF file.
    Saves the screenshot of the ordered robot.
    Embeds the screenshot of the robot to the PDF receipt.
    Creates ZIP archive of the receipts and the images.
    """
    init()

    open_robot_order_website()
    download_excel_file()
    fill_the_form()
    archive_receipts()

    cleanup()
def open_robot_order_website():
    # 
    page = browser.page()
    browser.goto("https://robotsparebinindustries.com/#/robot-order")

#get order
def get_order():
    data = Tables().read_table_from_csv("orders.csv")
    print(data)
    return data

def read_csv_file(filename):
    data = []
    with open(filename, 'rt', encoding="utf8") as csvfile:
            reader = csv.reader(csvfile)
            for row in reader:
                data.append(row)
    return data

def close_annoying_modal():
    page = browser.page()
    page.click("button:text('Yep')")


def fill_the_form():
    """Fills in the sales data and click the 'Submit' button"""
    page = browser.page()
    worksheet = get_order()

    for row in worksheet:
         fill_and_submit(row)

    archive_receipts()

def fill_and_submit(row):
    page = browser.page()

    page.select_option("#head", str(row["Head"]))
    rad = "#id-body-"+str(row["Body"])
    page.set_checked(rad, True)
    page.get_attribute
    page.keyboard.press('Tab')
    page.keyboard.press(str(row["Legs"]))

    page.fill("#address", row["Address"])    

    page.click("#preview")
   

    is_alert_visible = True

    while is_alert_visible:
        page.click("#order")
        is_alert_visible = page.locator("//div[@class='alert alert-danger']").is_visible()

        if not is_alert_visible:
            break

       
    store_receipt_as_pdf(row["Order number"])
    
    page.click("#order-another")
    close_annoying_modal()

def download_excel_file():
    http = HTTP()
    http.download(url="https://robotsparebinindustries.com/orders.csv", overwrite=True)


def store_receipt_as_pdf(order_number):
    """Saves file to a location in output folder"""
    page = browser.page()
    pdf = PDF()
    pdf_path = "output/receipts/order_"+order_number+".pdf"
   
    order_receipt_html = page.locator("#receipt").inner_html()
    pdf.html_to_pdf(order_receipt_html,  pdf_path)

    #take screenshot
    screenshot_robot(order_number)
   
    #embed screenshot in pdf
    embed_screenshot_to_receipt("output/receipts/order_"+order_number+".png", pdf_path)
    

def screenshot_robot(order_number):
    page = browser.page()
    page.locator(selector="#robot-preview-image").screenshot(path="output/receipts/order_"+order_number+".png")


def embed_screenshot_to_receipt(screenshot, pdf_file):
    pdf = PDF()
    
    pdf.add_files_to_pdf(
        files=[screenshot],
        target_document=pdf_file,
        append=True
    )


def archive_receipts():
    """Downloads excel file from the given URL"""
    lib = Archive()
    lib.archive_folder_with_zip("output/receipts",'receipts.zip',recursive=True)


def init():
    "All functions that need to be done to initialize robot"
    create_folders()

def create_folders():
    "All functions that need to be done to initialize robot"
    os.makedirs("output/receipts", exist_ok=True)

def cleanup():
    "All functions that need to be done to cleanup robot robot run"
    # TODO: delete folders that are more than X days old
    # TODO: delete downloaded orders.csv
