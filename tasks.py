from robocorp.tasks import task
from robocorp import browser
from RPA.HTTP import HTTP
from RPA.Tables import Tables
from RPA.PDF import PDF
from RPA.Archive import Archive
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
    orders = get_orders()
    for order in orders:
        open_robot_order_website()
        fill_the_form(order)
        order_number = order["Order number"]
        store_receipt_as_pdf(order_number)
        go_to_order_another_robot()
    
    archive_receipts()
        
    
def open_robot_order_website():
    """Opens the robot ordering website"""
    browser.goto("https://robotsparebinindustries.com/#/robot-order")
    close_annoying_modal()
    
def close_annoying_modal():
    """Closes annoying modal and signs off my rights to website"""
    page = browser.page()
    page.click("button:text('OK')")

def get_orders():
    """Downloads the orders from the website and generates a Table object for us to work off of"""
    http = HTTP()
    http.download("https://robotsparebinindustries.com/orders.csv", "orders.csv", overwrite=True)
    table = Tables()
    worksheet = table.read_table_from_csv("orders.csv", header=True)
    return worksheet

def fill_the_form(row):
    """Fills out the form using the data provided in orders.csv"""
    head_id = row["Head"]
    body_id = row["Body"]
    leg_id = row["Legs"]
    address = row["Address"]
    
    body_selector = str("#id-body-" + body_id)
    
    page = browser.page()
    page.select_option("#head", head_id)
    page.locator(body_selector).click()
    page.get_by_placeholder("Enter the part number for the legs").fill(leg_id)
    page.fill("#address", address)
    page.click("button:text('Preview')")
    while not page.is_visible("#receipt"):
        page.click("button:text('Order')")
        page.wait_for_timeout(1000)
    
def store_receipt_as_pdf(order_number):
    """Extracts the receipt from the webpage and stores it as a pdf, as well as appending an image of the robot to it."""
    pdf = PDF()
    page = browser.page()
    dir = "output/receipts"
    receipt_path = dir + "/" + order_number + "_receipt.pdf"
    content = page.locator("#receipt").inner_html()
    pdf.html_to_pdf(content, receipt_path)
    pdf.close_all_pdfs()
    screenshot = screenshot_robot(order_number, dir)
    embed_screenshot_to_receipt(screenshot, receipt_path)
    
    

def screenshot_robot(order_number, path):
    """Saves a screenshot of the robot with the order_number as its output"""
    screenshot_path = path + "/robot_" + order_number + ".png"
    page = browser.page()
    page.locator("#robot-preview").screenshot(path=screenshot_path)
    return screenshot_path

def embed_screenshot_to_receipt(screenshot_path, receipt_path):
    """Embeds the screenshot of the robot to the receipt pdf as a watermark, removes the original screenshot after to save on storage space"""
    pdf = PDF()
    pdf.open_pdf(receipt_path)
    pdf.add_watermark_image_to_pdf(
        image_path=screenshot_path,
        source_path=receipt_path,
        output_path=receipt_path
    )
    pdf.close_all_pdfs()
    os.remove(screenshot_path)
    
def go_to_order_another_robot():
    """Navigates to the start-page to handle multiple orders"""
    page = browser.page()
    page.click("#order-another")
    
def archive_receipts():
    """Archives all the receipts as a zip-file, receipts.zip"""
    lib = Archive()
    lib.archive_folder_with_zip("output/receipts", "receipts.zip", exclude="*.png")
    