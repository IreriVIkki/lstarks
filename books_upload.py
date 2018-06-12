import time
import openpyxl
from json import dumps
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support.select import Select


# ##============Logging in to account
fox = webdriver.Firefox()
fox.get('https://www.createspace.com/Login.do')
fox.find_element_by_id('loginField').send_keys('dezinerjimmy@gmail.com')
fox.find_element_by_id('passwordField').send_keys('67864322')
fox.find_element_by_id('login_button').click()
# ######=================================================######


# ##===============Setting up global elements
def upload_next_book(upload_link):
    return upload_link


def loaded_xpath(xpath):
    wait = WebDriverWait(fox, 45)
    element = wait.until(EC.visibility_of_element_located((By.XPATH, xpath)))
    return element


def loaded_link_text(link_txt):
    wait = WebDriverWait(fox, 45)
    element = wait.until(EC.visibility_of_element_located((By.LINK_TEXT, link_txt)))
    return element


def loaded_id(el_id):
    wait = WebDriverWait(fox, 45)
    element = wait.until(EC.visibility_of_element_located((By.ID, el_id)))
    return element


def loaded_css_selector(css_selector):
    wait = WebDriverWait(fox, 45)
    element = wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, css_selector)))
    return element


def loaded_tag_name(tag_name):
    wait = WebDriverWait(fox, 45)
    element = wait.until(EC.visibility_of_element_located((By.NAME, tag_name)))
    return element


wb = openpyxl.load_workbook(r'C:\Users\Me MySelf I\Desktop\desktop prmng\New folder\moms.xlsx')
sheet = wb.active

list = sheet['a3'].value.split('\n')

# print(list)

for index, i in enumerate(list):
    book_title = f"{i.split(': ')[0]} Notebook"
    print(f'title ----->  {book_title}')

    subtitle = f"{i.split(': ')[1]} Notebook Gift |120 pages ruled With Stethoscope cover"
    print(f'subtitle ----->  {subtitle}')

    last_name = f"{i.split(': ')[0]} Gift"
    print(f'title  ----->  {last_name}')

    book_pdf = '\\'.join([sheet['d2'].value, sheet['e2'].value])
    print(f'book path  ----->  {book_pdf}')

    cover_name = f"page-{21 + index}.pdf"
    book_cover_pdf = '\\'.join([sheet['d2'].value, cover_name])
    print(f'cover path  ----->  {book_cover_pdf}')

    bisac = sheet['h2'].value
    print(f'cover path  ----->  {bisac}')

    description = f"{book_title} \n\n\nA simple gift idea; 120 pages ruled notebook with a glossy finish custom cover. \n\nAn a4 size general purpose notebook."
    print(f'description ----->  {description}')

    keywords = sheet['i2'].value
    print(f'keywords ----->  {keywords}')
    print('\n')

    current = fox.current_url
    destination = ('https://tsw.createspace.com/getstarted/productselection')

    print(fox.current_window_handle)
    print(fox.window_handles)

    if current != destination:
        fox.switch_to_window(fox.window_handles[-1])
        fox.get(destination)

    print(fox.window_handles[-1])
    print(fox.current_window_handle)

    book_title = loaded_id('tsw_productselection_titlename').send_keys(book_title)

    paperback_option = loaded_id('paperback_radio').click()

    advanced_setup = loaded_xpath('//div[@id="advancedSetupOption"]/a/div/div[2]').click()

    subtitle_input = loaded_css_selector('input[class="contentFormField contentTextField subtitle_field formItemNormal"]').send_keys(subtitle)

    last_name_input = loaded_xpath('//div[@class="action PRIMARY_CONTRIBUTOR"]/div[1]/div/div[2]/div[1]/div[5]/input').send_keys(last_name)

    language_selection1 = Select(loaded_xpath('//div[@class="action LANGUAGE"]/div[2]/div[2]/select')).select_by_visible_text('English')

    isbn_choice = loaded_xpath('//div[@class="artifactContainer artifactContainerbook_isbnAjaxContainer"]/div/span/div/div[4]/div[1]/input').click()

    choose_size = loaded_xpath('//div[@class="physPropTrimSize"]/div/a').click()

    size_choice = loaded_xpath('//div[@class="trimSizeLBContent"]/div[3]/div[6]/div[4]').click()

    bleed = WebDriverWait(fox, 45).until(EC.element_to_be_clickable((By.XPATH, '//span[@class="services"]/div/div[4]/div[2]/div/div/div/form/div/div[2]/div/div[2]/div/div[3]/input[1]'))).click()

    if loaded_id("tsw_advanced_book_interior_tsw_book_interior__uploadTextTarget"):
        fox.execute_script('document.getElementById("tsw_advanced_book_interior_tsw_book_interior__uploadTextTarget").disabled = "false"')

        fox.execute_script('document.getElementById("tsw_advanced_book_interior_tsw_book_interior__uploadTextTarget").value =' + dumps(book_pdf)
                           )

    if loaded_id("tsw_advanced_book_cover_tsw_book_cover__uploadTextTarget"):
        fox.execute_script('document.getElementById("tsw_advanced_book_cover_tsw_book_cover__uploadTextTarget").disabled = "false"')

        fox.execute_script('document.getElementById("tsw_advanced_book_cover_tsw_book_cover__uploadTextTarget").value =' + dumps(book_cover_pdf))

    book_description = loaded_xpath('//textarea[@id="tsw_advanced_description_description"]').send_keys(description)

    # ###==============Set bisac category chooser=============#####
    try:
        fox.execute_script('document.getElementById("bisacCategoryChooser").style = "display: none;"')
    finally:
        fox.execute_script('document.getElementById("bisacCodeInput").style = ""')

    bisac_code = fox.find_element_by_xpath('//input[@id="tsw_advanced_description_bisaccode"]').send_keys(bisac)

    # ###=======================######========================#####

    book_language = Select(loaded_id('tsw_advanced_description_language')).select_by_visible_text('English')

    book_country = Select(loaded_id('tsw_advanced_description_publicationcountry')).select_by_visible_text('United States')

    book_keywords = loaded_xpath('//input[@id="tsw_advanced_description_searchkeywords"]').send_keys(keywords)

    save_changes = loaded_xpath('//div[@id="distributeChannels"]/div/div/div[3]/div[3]/div').click()

    time.sleep(4)

    fox.execute_script("window.open('https://tsw.createspace.com/getstarted/productselection', 'new window')")

    print(f'***************\n   Book {index + 1} Uploaded\n**********')
