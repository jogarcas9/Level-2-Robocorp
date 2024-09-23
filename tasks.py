from robocorp.tasks import task
from RPA.HTTP import HTTP
from RPA.Tables import Tables
from RPA.Browser.Selenium import Selenium
from selenium.common.exceptions import NoSuchElementException
from bs4 import BeautifulSoup
import os
from RPA.PDF import PDF
import time
import zipfile
#Variables

file_url = 'https://robotsparebinindustries.com/orders.csv'
browser_url = 'https://robotsparebinindustries.com/#/robot-order'
page = Selenium(auto_close=False)
user_path=r"C:\Users\jogar\AppData\Local\Google\Chrome\User Data\Default"
pdf = PDF()
main_path=r'C:\Users\jogar\OneDrive\Desktop\Robocorp level 2\recipies'


@task
def order_robots_from_RobotSpareBin():
    open_robot_order_website()
    download_file()
    orders = get_orders()
    for order in orders:
        close_annoying_modal()
        order_number = fill_page(order)
        pdf_path=create_pdf(order_number)
        png_path=screenshot_robot(pdf_path)
        combine_png_pdf(pdf_path, png_path)
        click_new_robot()
    archive_receipts(main_path)

def open_robot_order_website():
    page.open_available_browser(browser_url, maximized=True)
    
def download_file():
    http = HTTP()
    http.download(file_url, overwrite=True)

def get_orders():
    library = Tables()
    orders = library.read_table_from_csv(
        'orders.csv', columns=['Order number', 'Head', 'Body','Legs', 'Address']
    )
    return orders

def close_annoying_modal():
    page.click_button('OK')

def fill_page(x):
    Order_number = x['Order number']
    Head = x['Head']
    Body = x['Body']
    Legs = x['Legs']
    Address = x['Address']
    
    page.select_from_list_by_value('id=head', Head)
    radio_button='id=id-body-'+str(Body)
    page.click_button(radio_button)
    id=find_name_selector()
    leg_id='id='+id
    page.input_text(leg_id, str(Legs))
    page.input_text('id=address', str(Address))
    page.click_button('id=order')
    try:
        if page.does_page_contain_element('class=alert-danger'):
            time.sleep(5)
            page.click_button('id=order')
    finally:
        pass
    return Order_number

def find_name_selector():
    html=page.get_source()
    soup=BeautifulSoup(html, 'html.parser')
    label=soup.find('label', text="3. Legs:")
    legs_id=label['for']
    return legs_id

def create_pdf(number):

    pdf_path=r"C:\Users\jogar\OneDrive\Desktop\Robocorp level 2\recipies"
    pdf_path=pdf_path+'\\'+str(number)+'.pdf'
    html_content = """
    <html>
    <head>
        <title><h1>Order Confirmation</h1></title>
    </head>
    <body>
        <h1>Thank you for your order!</h1>
        <p>We will ship your robot to you as soon as our warehouse robots gather the parts you ordered!</p>
        <p>You will receive your robot in no time!</p>
    </body>
    </html>
    """
    pdf.html_to_pdf(html_content, pdf_path)
    return pdf_path

def screenshot_robot(pdf_path):
    jpg_path=pdf_path.replace('.pdf','.jpg')
    page.screenshot('id=robot-preview-image',jpg_path)
    return jpg_path

def combine_png_pdf(pdf_file, png_file):
    html_content = f"""
    <html>
    <body>
        <img src="{png_file}" alt="Screenshot" style="width: 100%;"/>
    </body>
    </html>
    """

    pdf.html_to_pdf(html_content, 'temp.pdf')
    lst_files=[pdf_file,'temp.pdf']
    pdf.add_files_to_pdf(files=lst_files, 
                         target_document=pdf_file)
    os.remove('temp.pdf')
    os.remove(png_file)

def click_new_robot():
    page.click_button('id=order-another')

def archive_receipts(pdf_directory):
    zip_name = 'pdfs.zip'
    with zipfile.ZipFile(zip_name, 'w') as zipf:
        for root, files in os.walk(pdf_directory):
            for file in files:
                if file.endswith('.pdf'):
                    file_path=os.path.join(root,file)
                    zipf.write(file_path,pdf_directory)