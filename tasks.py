from robocorp.tasks import task
from robocorp import browser

from RPA.HTTP import HTTP
from RPA.Tables import Tables
from RPA.PDF import PDF
from RPA.Archive import Archive


@task
def order_robots_from_RobotSpareBin():
    """
    Orders robots from RobotSpareBin Industries Inc.
    Saves the order HTML receipt as a PDF file.
    Saves the screenshot of the ordered robot.
    Embeds the screenshot of the robot to the PDF receipt.
    Creates ZIP archive of the receipts and the images.
    """
    browser.configure(
        slowmo=250,
    )
    open_robot_order_website()
    download_excel_file()
    get_orders()
    archive_receipts_file()
    

def open_robot_order_website():
    """Navigates to the given URL and click OK"""
    browser.goto("https://robotsparebinindustries.com/#/robot-order")
    
def close_annoying_modal():
    """Close the modal by clicking OK"""
    page=browser.page()
    page.click("text=OK")

def download_excel_file():
    """Downloads excel file from the given URL"""
    http = HTTP()
    http.download(url="https://robotsparebinindustries.com/orders.csv", overwrite=True)

def get_orders():
    """Read data from excel and convert into table"""
    orders=Tables().read_table_from_csv("orders.csv")

    for order in orders:
        fill_and_submit_order_form(order)

def fill_and_submit_order_form(order):
    """Fills in the orders data, 'PREVIEW' the order and click the 'ORDER' button"""
    page = browser.page()
    close_annoying_modal()

    page.select_option("#head", order["Head"] )
    page.click("#id-body-"+order["Body"])
    page.fill("//html/body/div/div/div[1]/div/div[1]/form/div[3]/input",order["Legs"])
    page.fill("#address", order["Address"])
    page.click("text=Preview")
    page.click("text=ORDER")

    while not page.query_selector("#order-another"):
        page.click("#order")
        
    store_receipt_as_pdf(order["Order number"])
    screenshot_robot(order["Order number"])
    page.click("#order-another")


def store_receipt_as_pdf(order_number):
    """Stores receipts to PDF and takes a screenshot of ordered robots"""
    page=browser.page()

    pdf=PDF()
    pdf_file=f"output/receipt/{order_number}.pdf"
    pdf.html_to_pdf(page.locator("#receipt").inner_html(),pdf_file)

def screenshot_robot(order_number):
    """Take a screenshot of the receipt"""
    page=browser.page()
    img=page.query_selector("#robot-preview-image")
    screenshot=f"output/screenshot/{order_number}.png"
    img.screenshot(path=screenshot)

    pdf_file=f"output/receipt/{order_number}.pdf"
    embed_screenshot_to_receipt(screenshot, pdf_file)

def embed_screenshot_to_receipt(screenshot, pdf_file):
    """Embed the robot screenshot to the receipt PDF file"""
    pdf=PDF()
    pdf.add_files_to_pdf(files=[screenshot],target_document=pdf_file,append=True)

def archive_receipts_file(): 
    """Archives the receipts file into folder"""
    folder=Archive()
    folder.archive_folder_with_zip('output/receipt','output/orders.zip',include='*.pdf')
    





