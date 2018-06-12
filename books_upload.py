import time
import openpyxl
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support.select import Select

fox = webdriver.Firefox()
fox.get('https://www.createspace.com/Login.do')
fox.find_element_by_id('loginField').send_keys('dezinerjimmy@gmail.com')
fox.find_element_by_id('passwordField').send_keys('67864322')
fox.find_element_by_id('login_button').click()

current = fox.current_url
destination = ('https://tsw.createspace.com/getstarted/productselection')

if current != destination:
    new_title = fox.get(destination)


def loaded_xpath(xpath):
    wait = WebDriverWait(fox, 45)
    element = wait.until(EC.visibility_of_element_located((By.XPATH, xpath)))
    return element


def loaded_id(el_id):
    wait = WebDriverWait(fox, 45)
    element = wait.until(EC.visibility_of_element_located((By.ID, el_id)))
    return element


def loaded_css_selector(css_selector):
    wait = WebDriverWait(fox, 45)
    element = wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, css_selector)))
    return element


wb = openpyxl.load_workbook(r'C:\Users\Me MySelf I\Desktop\desktop prmng\New folder\moms.xlsx')
sheet = wb.active

row = 2

book_title = loaded_id('tsw_productselection_titlename').send_keys(sheet['A' + str(row)].value)

paperback_option = loaded_id('paperback_radio').click()

advanced_setup = loaded_css_selector('#advancedSetupOption > a:nth-child(3) > div:nth-child(1)').click()

subtitle_input = loaded_css_selector('input[class="contentFormField contentTextField subtitle_field formItemNormal"]').send_keys(sheet['B' + str(row)].value)

last_name_input = loaded_xpath('//div[@class="action PRIMARY_CONTRIBUTOR"]/div[1]/div/div[2]/div[1]/div[5]/input').send_keys(sheet['C' + str(row)].value)

language_selection1 = Select(loaded_xpath('//div[@class="action LANGUAGE"]/div[2]/div[2]/select')).select_by_visible_text('English')

isbn_choice = loaded_xpath('//div[@class="artifactContainer artifactContainerbook_isbnAjaxContainer"]/div/span/div/div[4]/div[1]/input').click()

choose_size = loaded_xpath('//div[@class="physPropTrimSize"]/div/a').click()

size_choice = loaded_xpath('//div[@class="trimSizeLBContent"]/div[3]/div[6]/div[4]').click()

time.sleep(3)

bleed = loaded_xpath('//span[@class="services"]/div/div[4]/div[2]/div/div/div/form/div/div[2]/div/div[2]/div/div[3]/input[1]').click()

if loaded_id("tsw_advanced_book_interior_tsw_book_interior__uploadTextTarget"):
    fox.execute_script('document.getElementById("tsw_advanced_book_interior_tsw_book_interior__uploadTextTarget").disabled = "false"')

    fox.execute_script('document.getElementById("tsw_advanced_book_interior_tsw_book_interior__uploadTextTarget").value = "C:\\Users\\Me MySelf I\\Desktop\\desktop prmng\\New folder\\moms_book.pdf"')

if loaded_id("tsw_advanced_book_cover_tsw_book_cover__uploadTextTarget"):
    fox.execute_script('document.getElementById("tsw_advanced_book_cover_tsw_book_cover__uploadTextTarget").disabled = "false"')

    fox.execute_script('document.getElementById("tsw_advanced_book_cover_tsw_book_cover__uploadTextTarget").value = "C:\\Users\\Me MySelf I\\Desktop\\desktop prmng\\New folder\\SINGLE MOM.pdf"')

book_description = loaded_xpath('//textarea[@id="tsw_advanced_description_description"]').send_keys(sheet['G' + str(row)].value)

# ###==============Set bisac category chooser=============#####
try:
    fox.execute_script('document.getElementById("bisacCategoryChooser").style = "display: none;"')
finally:
    fox.execute_script('document.getElementById("bisacCodeInput").style = ""')

fox.find_element_by_xpath('//input[@id="tsw_advanced_description_bisaccode"]').send_keys(sheet['H' + str(row)].value)

# ###=======================######========================#####

book_language = Select(loaded_id('tsw_advanced_description_language')).select_by_visible_text('English')

book_country = Select(loaded_id('tsw_advanced_description_publicationcountry')).select_by_visible_text('United States')

book_keywords = loaded_xpath('//input[@id="tsw_advanced_description_searchkeywords"]').send_keys(sheet['I' + str(row)].value)
