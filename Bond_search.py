import config_webdriver_manager
import web_actions
import os
#import msvcrt
import time
import datetime
import csv
import numpy as np
import pandas as pd 
import random
import xlsxwriter
from bs4 import BeautifulSoup
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver import ActionChains
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from locators import Locators
from selenium.webdriver.support import expected_conditions as EC
from datetime import date
from dateutil.relativedelta import relativedelta
import openpyxl
from selenium.common.exceptions import  TimeoutException
from openpyxl import load_workbook
import clipboard as clp
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import NoSuchElementException
config_webdriver_manager.driver.command_executor._commands["send_command"] = ("POST", '/session/$sessionId/chromium/send_command')
params = {'cmd': 'Page.setDownloadBehavior', 'params': {'behavior': 'allow', 'downloadPath': "/path/to/download/dir"}}
command_result = config_webdriver_manager.driver.execute("send_command", params)

#config_webdriver_manager.driver.get("https://raikkulenz.kapsi.fi/myproject/test.htm")
config_webdriver_manager.driver.get("https://www.boerse-stuttgart.de/de-de/tools/produktsuche/anleihen-finder/")

time.sleep(20)
# Navigate to the webpage
# Wait until the cookie consent element is present
button = config_webdriver_manager.driver.find_element(By.CSS_SELECTOR, Locators.cookie_acceptor_css_new)
button.click()


""" cookie_consent_aside = WebDriverWait(config_webdriver_manager.driver, 15).until(
    EC.presence_of_element_located((By.ID, "cookie-consent"))
)

l=cookie_consent_aside.location

# Execute JavaScript to change the z-index of the element
config_webdriver_manager.driver.execute_script("arguments[0].style.zIndex = '50';", cookie_consent_aside)

# Verify the change (optional)
print("z-index changed to:", config_webdriver_manager.driver.execute_script("return arguments[0].style.zIndex;", cookie_consent_aside))


cookie_acceptor_click_button=WebDriverWait(config_webdriver_manager.driver,15).until(EC.presence_of_element_located((By.CSS_SELECTOR, Locators.Cookie_einstellungen_CSS_PATH)))

l1=cookie_acceptor_click_button.location

x1 = l1['x']
y1 = l1['y']
print("Cooke acceptor button coordinates")
print(f'x: {x1}')
print(f'y: {y1}')
s=cookie_acceptor_click_button.size

height = s['height']
width = s['width']

print(f'height: {height}')
print(f'width: {width}')

x_coordinate_cookie_acceptor_button=x1+width/2
y_coordinate_cookie_acceptor_button=y1+height/2

print(f'x coordinate of button: {x_coordinate_cookie_acceptor_button}')
print(f'y coordinate of button: {y_coordinate_cookie_acceptor_button}')


actions = ActionChains(config_webdriver_manager.driver)

actions.move_by_offset(x_coordinate_cookie_acceptor_button, y_coordinate_cookie_acceptor_button)
time.sleep(1)
actions.click().perform()
 """

web_actions.zoom40androlldown()
web_actions.relax()

#print("Click on Zusaetzliche Filter")
config_webdriver_manager.driver.execute_script("document.body.style.zoom='60%'")# 
web_actions.relax()


#web_actions.click_operation_CSS(Locators.Wahrung_click_CSS_SELECTOR)
web_actions.click_operation_xpath(Locators.Wahrung_click_XPATH)

web_actions.click_operation_xpath(Locators.Wahrung_text_field_click_XPATH)

web_actions.write_operation_selector(Locators.Wahrung_text_field_click_CSS_3_SELECTOR,"Euro")

time.sleep(random.randint(2,6))

checkbox = config_webdriver_manager.driver.find_element(By.CSS_SELECTOR, Locators.Wahrung_checkbox_Eur1_CSS)
checkbox.click()



#checkbox = WebDriverWait(config_webdriver_manager.driver, 10).until(
#    EC.element_to_be_clickable((By.XPATH, '//label[text()="Euro"]/preceding-sibling::input[@type="checkbox"]'))
#)
#checkbox.click()


#web_actions.click_operation_xpath(Locators.Wahrung_Euro_small_cross_click_XP)


web_actions.click_operation_xpath(Locators.Wahrung_Anwenden_XP)
#web_actions.click_operation_CSS(Locators.Wahrung_checkbox_Eur_CSS)
print("Milestone Click Anwenden of euro through")

time.sleep(random.randint(2,6))
#web_actions.click_operation_css(Locators.Wahrung_anwenden_css)
#time.sleep(random.randint(2,6))
web_actions.relax()
time.sleep(random.randint(2,6))

web_actions.click_operation_Xpath(Locators.Zusatzliche_Filter_right_XPATH)
web_actions.click_operation_css(Locators.handelbare_Einh_css)
web_actions.click_operation_css(Locators.Gelisteter_Zeitraum_Button_css)
web_actions.relax()
time.sleep(random.randint(2,6))

#web_actions.click_operation_CSS(Locators.Zusatzliche_Filter_Anwenden_CSS)
web_actions.click_operation_Xpath(Locators.Zusatzliche_Filter_Anwenden_Full_XPath)
web_actions.relax()

web_actions.click_operation_Xpath(Locators.Handelbare_einh_Button_xpath)
web_actions.relax()
time.sleep(random.randint(2,6))
web_actions.clear_operation_A_xpath(Locators.min_Littera_xp)
web_actions.write_operation_b_xpath(Locators.min_Littera_xp,"1")
time.sleep(random.randint(2,6))
web_actions.clear_operation_A_xpath(Locators.max_Littera_xp)


input_field = config_webdriver_manager.driver.find_element(By.CSS_SELECTOR, Locators.max_Littera_css)
input_field.clear()
input_field.send_keys('1000', Keys.RETURN)

""" input_field = config_webdriver_manager.driver.find_element(By.CSS_SELECTOR, Locators.max_Littera_css)
actions = ActionChains(config_webdriver_manager.driver)
actions.send_keys()
actions.click(input_field).send_keys('1000').send_keys(Keys.RETURN).perform() """



web_actions.click_operation_Xpath(Locators.Anwenden_littera_xp)
#web_actions.click_operation_css(Locators.Anwenden_littera_css)


#Choosing Bond Listed date from to date 
Smaller_date=  datetime.date.today()+ relativedelta(years=5)
Smaller_date_format = Smaller_date.strftime("%d.%m.%Y")

input_field = WebDriverWait(config_webdriver_manager.driver, 10).until(
    EC.element_to_be_clickable((By.XPATH, '//input[@placeholder="Fälligkeit von:"]'))
)
input_field.send_keys(Smaller_date_format)

web_actions.click_operation_Xpath(Locators.Smaller_date_2_xp)
time.sleep(random.randint(5,10))
web_actions.relax()

""" web_actions.clear_operation_A_xpath(Locators.Smaller_date_2_xp)
Smaller_date=  datetime.date.today()+ relativedelta(years=5)
Smaller_date_format = Smaller_date.strftime("%d.%m.%Y")
web_actions.write_operation_xpath(Locators.Smaller_date_2_xp,Smaller_date_format)
time.sleep(random.randint(2,10))
 """

Bigger_date=  datetime.date.today()+ relativedelta(years=10)
Bigger_date_format = Bigger_date.strftime("%d.%m.%Y")

time.sleep(random.randint(1,10))
web_actions.relax()



web_actions.click_operation_Xpath("/html/body/div[3]/main/div[2]/div/div[1]/div[5]/div/div/div[1]/div/div/input[2]")
#Ei vielä toimii 
time.sleep(random.randint(1,10))
web_actions.relax()


""" input_field_3 = config_webdriver_manager.driver.find_element(By.XPATH, Locators.falligkeit_Bigger_date_fxp)
actions = ActionChains(config_webdriver_manager.driver)
actions.click(input_field_3).send_keys(Bigger_date_format).send_keys(Keys.RETURN).perform() """

input_field_3 = WebDriverWait(config_webdriver_manager.driver, 10).until(
    EC.element_to_be_clickable((By.XPATH, '//input[@placeholder="bis:"]'))
)
input_field_3.clear()
input_field_3.send_keys(Bigger_date_format)
input_field_3.send_keys(Keys.RETURN)


time.sleep(random.randint(5,10))
#web_actions.write_operation_xpath(Locators.faelligkeit_element_bigger_date_FULL_XPATH,Bigger_date_format)

# miksei toimii???

#web_actions.click_and_write_by_actionchain_XPATH(Locators.faelligkeit_element_bigger_date_FULL_XPATH,Bigger_date_format)



#Gelisteter Zeitraum Bond Listed date from to date 
web_actions.click_operation_css(Locators.Zusatzliche_Filter_css)
#web_actions.click_operation_Xpath(Locators.Zusatzliche_Filter_Full_XPath)

# Zusatzliche_Filter_Full_XPath 

""" web_actions.relax()
web_actions.click_operation_css(Locators.Gelisteter_Zeitraum_ruksi_css)

web_actions.relax()
web_actions.click_operation_css(Locators.Anwenden_Gelisteter_Zeitraum_css)
web_actions.relax()
web_actions.click_operation_Xpath(Locators.Anwenden_Gelisteter_Zeitraum_XPATH)
web_actions.relax()
 """
Smaller_date_listed=  datetime.date.today()- relativedelta(years=3)
Smaller_date_format_listed = Smaller_date_listed.strftime("%d.%m.%Y")

web_actions.click_and_write_by_actionchain_XPATH(Locators.gelistet_zeitraum_element_smaller_Date_XPATH,Smaller_date_format_listed)

time.sleep(random.randint(2,16))
#web_actions.click_operation_CSS(Locators.Preselector_Anleihen_Typ_CSS)
web_actions.click_operation_Xpath(Locators.Preselector_Anleihen_Typ_FULL_XPATH)
web_actions.relax()
web_actions.click_operation_Xpath('//*[@id="categoryLevel2"]/div/div/div[1]/div/input')
web_actions.relax()
web_actions.write_operation_xpath('//*[@id="categoryLevel2"]/div/div/div[1]/div/input',"Unternehmensanleihe")
time.sleep(random.randint(2,16))

web_actions.relax()

#web_actions.click_operation_id_A(Locators.Unternehmensanleihe_checkbox_ID)

#Element_ruksi=Locators.Unternehmensanleihe_checkbox_with_text_XPATH




xpath_expression = '/html/body/div[3]/main/div[2]/div/div[1]/div[2]/div/div/ul/li/div/div/span/input'
element = WebDriverWait(config_webdriver_manager.driver, 20).until(
    EC.presence_of_element_located((By.XPATH, xpath_expression))

)
actions = ActionChains(config_webdriver_manager.driver)
time.sleep(random.randint(4,6))

web_actions.relax()
actions.move_to_element(element).click().perform()

# css_selector = '#categoryLevel2 > div > div > ul > li > div > div > span'
# element = WebDriverWait(config_webdriver_manager.driver, 20).until(
#     EC.element_to_be_clickable((By.CSS_SELECTOR, css_selector))
# )

# actions = ActionChains(config_webdriver_manager.driver)
# actions.move_to_element(element)

# time.sleep(1)
# actions.click().perform()

#web_actions.click_operation_CSS(Locators.Anwenden_Unternehmensanleihe_Select_class)

web_actions.click_operation_Xpath(Locators.Anwenden_Unternehmensanleihe_XPATH)
web_actions.relax()
time.sleep(random.randint(2,6))
web_actions.click_operation_Xpath(Locators.Rendite_button_XPATH)

time.sleep(random.randint(2,6))


web_actions.click_operation_Xpath(Locators.Rendite_textbox_FULLXPATH)

time.sleep(random.randint(2,6))
web_actions.relax()
web_actions.clear_operation_A_xpath(Locators.Rendite_textbox_FULLXPATH)
web_actions.write_operation_xpath(Locators.Rendite_textbox_FULLXPATH,'3%')
web_actions.relax()

time.sleep(random.randint(2,6))

web_actions.click_operation_Xpath(Locators.Rendite_Anwenden_XPATH)
web_actions.relax()


# time.sleep(random.randint(2,6))
# web_actions.click_operation_id_A(Locators.Gelisteter_Zeitraum_preselector_Button_id)
# web_actions.click_operation_Xpath(Locators.Zusatzliche_Filter_Anwenden_Full_XPath)
# web_actions.click_operation_Xpath(Locators.Gelisteter_Zeitraum_Button_xp)
# time.sleep(random.randint(2,6))
# web_actions.clear_operation_A_xpath(Locators.Gelisteter_Zeitraum_smaller_date_xp)
# Smaller_date_listed=  datetime.date.today()- relativedelta(years=3)
# Smaller_date_format_listed = Smaller_date_listed.strftime("%d.%m.%Y")
# web_actions.write_operation_xpath(Locators.Gelisteter_Zeitraum_smaller_date_xp,Smaller_date_format_listed)
# web_actions.click_operation_Xpath(Locators.Gelisteter_Zeitraum_Bigger_date_xp)
# time.sleep(random.randint(2,6))
# web_actions.clear_operation_A_xpath(Locators.Gelisteter_Zeitraum_Bigger_date_xp)
# Bigger_date_listed= datetime.date.today().strftime("%d.%m.%Y")
# web_actions.write_operation_xpath(Locators.Gelisteter_Zeitraum_Bigger_date_xp, Bigger_date_listed)
# time.sleep(random.randint(2,6))
# web_actions.click_operation_Xpath(Locators.Anwenden_Gelisteter_Zeitraum_full_XPath)
# #Select Unternehmensanleihe = Bond
# web_actions.click_operation_id_A(Locators.Preselector_Anleihen_Typ_ID)
# web_actions.click_operation_id_A(Locators.Select_Unternehmensanleihe_Checkbox_ID)
# web_actions.click_operation_Xpath(Locators.Anwenden_Unternehmensanleihe_Select_full_XPath)
# # Choose Yield between maximum and minimum
# web_actions.click_operation_Xpath(Locators.Rendite_button_xp)
# web_actions.clear_operation_A_xpath(Locators.Rendite_lower_margin_xp)
# web_actions.write_operation_xpath(Locators.Rendite_lower_margin_xp, "3%")
# web_actions.clear_operation_A_xpath(Locators.Rendite_upper_margin_xp)
# web_actions.write_operation_xpath(Locators.Rendite_upper_margin_xp, "16%")
# time.sleep(random.randint(2,6))
# web_actions.click_operation_Xpath(Locators.Rendite_anwenden_xp)
#Sort by yield
time.sleep(random.randint(2,6))


config_webdriver_manager.driver.execute_script ("window.scrollTo(0,document.body.scrollHeight);")
#show_more_hits
web_actions.click_operation_Xpath(Locators.Weitere_Treffer_Anzeigen_XPATH)
_j=1
_n=5
while _j<_n:
    config_webdriver_manager.driver.execute_script ("window.scrollTo(0,document.body.scrollHeight);")
    time.sleep(random.randint(2,6))
    web_actions.click_operation_Xpath(Locators.Weitere_Treffer_Anzeigen_XPATH)
    _j+=1
    time.sleep(random.randint(2,6))




# Write Results to Excel file
html_source =config_webdriver_manager.driver.page_source
soup=BeautifulSoup(html_source, 'html.parser')
table=soup.table
table_rows = table.find_all('tr')

tables=[]
for tr in table_rows:
    td = tr.find_all('td')
    row=[]
    for i in td:
        word= i.text.strip()
        word=word.replace(".000", "000") #littera remove seperators
        word=word.replace(",000", "") #littera remove zeros after comma
        row.append(word)
    tables.append(row)
print(tables)
df = pd.DataFrame(tables)

#out= df.to_json(orient='records')[1:-1].replace('},{', '} {')
#with open('c:\\temp\\tutorial\\df_to_json.txt', 'w') as f:
#    f.write(out)


wb = load_workbook('c:\\temp\\tutorial\\sample.xlsx')

sheet = wb.worksheets[0]
# this statement inserts a column before column 2

for x in range(0, 10, 1):
    sheet.insert_cols(0)

web_actions.df_to_excel(df, wb.active)
wb.save("c:\\temp\\tutorial\\sample.xlsx")

while True:
    pass
driver.quit()
config_webdriver_manager.driver.quit()

