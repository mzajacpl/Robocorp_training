from robocorp.tasks         import task
from RPA.Windows            import Windows
from RPA.Browser.Selenium   import Selenium
from RPA.Excel.Files        import Files
from RPA.HTTP               import HTTP
from RPA.Tables             import Tables
from RPA.PDF                import PDF
from RPA.Archive            import Archive
from datetime               import datetime
import requests
import os

"""Variables"""
browser = Selenium()
http = HTTP()
excel_file = Files()
target_path = "C:\input\Order_Robots_from_RobotSpareBin"
tables = Tables()
pdf = PDF()
lib = Archive()

@task
def order_robots_from_RobotSpareBin():
       """
    Orders robots from RobotSpareBin Industries Inc.
    Saves the order HTML receipt as a PDF file.
    Saves the screenshot of the ordered robot.
    Embeds the screenshot of the robot to the PDF receipt.
    Creates ZIP archive of the receipts and the images.
    """
       #download_excel_file()
       open_robot_order_website()
       get_orders_from_excel()
       archive_receipts()
       browser.close_browser()


       

def open_robot_order_website():
    """open the page and go to order website"""
    browser.open_chrome_browser("https://robotsparebinindustries.com/", maximized=True)
    browser.auto_close = False
    browser.click_element("css:a[href='#/robot-order']")
    #browser.open_chrome_browser("https://robotsparebinindustries.com/#/robot-order", maximized=True)
    

def download_excel_file():
    """Downloads excel file from the given URL"""
    response = requests.get("https://robotsparebinindustries.com/orders.csv")
    response.raise_for_status()
    with open('orders.csv', 'wb') as f:
        f.write(response.content)
    #http.download(url="https://robotsparebinindustries.com/orders.csv", overwrite=True)

def create_new_folder():
    folder_path = '.\\robots'
    if not os.path.exists(folder_path):
        os.makedirs(folder_path)

def get_orders_from_excel():
    """Read data from excel and fill in the sales form"""
    table = tables.read_table_from_csv(
        path = 'orders.csv',
        header = True,
        dialect= "excel"
    )
    create_new_folder()
    for row in table:
         browser.click_button_when_visible("xpath=//button[text()='OK']")
         fill_and_submit_order_form(row)
         name = store_receipt_as_pdf()
         screenshot_robot(name)
         pdf_name = f'{name}.pdf'
         image_name = f'{name}.png'
         embed_screenshot_to_receipt(image_name, pdf_name)
         browser.click_button("xpath=//button[text()='Order another robot']")
         

def fill_and_submit_order_form(row):
    browser.select_from_list_by_value("css:select.custom-select", '1')
    body_value = row['Body']
    browser.click_element(f'css:.form-check-input[type="radio"][value="{body_value}"]')
    legs = browser.find_element("xpath://html/body/div/div/div[1]/div/div[1]/form/div[3]/input")
    browser.input_text(legs, '2')
    browser.input_text("xpath://html/body/div/div/div[1]/div/div[1]/form/div[4]/input", 'Address 123')
    browser.click_button("xpath=//button[text()='Order']")
    x = 0
    for x in range(0,3):
        element = browser.does_page_contain_button("xpath=//button[text()='Order']")
        if element is True:
            #check if element on web exists
            #press button again
            browser.click_button("xpath=//button[text()='Order']")
        x = x+1

def screenshot_robot(name):
    browser.screenshot("xpath=//html/body/div/div/div[1]/div/div[2]", f'{name}.png')

def store_receipt_as_pdf():
    order_html = browser.get_text("css:#receipt")
    #html = browser.execute_javascript('document.querySelector("#receipt").innerHTML')
    robot_name = browser.get_text("css:p.badge.badge-success")
    pdf.html_to_pdf(order_html, f".\\robots\{robot_name}.pdf")
    return robot_name
    
def embed_screenshot_to_receipt(screenshot, pdf_file):
    pdf.add_files_to_pdf(
        files=[screenshot],
        target_document=f'.\\robots\{pdf_file}',
        append=True
    )
def archive_receipts():
    lib.archive_folder_with_zip("./robots", 'robot.zip')