import config_webdriver_manager
from selenium.common.exceptions import NoSuchElementException
from selenium.common.exceptions import  TimeoutException
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
import openpyxl
from openpyxl import load_workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from selenium.webdriver import ActionChains
from selenium.webdriver.common.action_chains import ActionChains
import random
import time


def do_something(fname):
    print(config_webdriver_manager.driver)
    print(fname)


def do_something_else():
    print("Click")
    print(config_webdriver_manager.driver)
    
#USED 5
def click_operation_id_A(fname):
    try:
        element=config_webdriver_manager.driver.find_element(By.ID, fname)
        config_webdriver_manager.driver.execute_script("arguments[0].click();", element)
    except NoSuchElementException as exception:
        print ('Click on ' +  fname + '  ID A not successful')
    except TimeoutException:
        pass

#USED 1
def click_operation_id_B(fname):
    try:
        element=WebDriverWait(config_webdriver_manager.driver, 20).until(EC.element_to_be_clickable((By.ID, fname))).click()
        config_webdriver_manager.driver.find_elements(By.ID, fname).sendKeys('\uE035')
        #config_webdriver_manager.driver.execute_script("arguments[0].click();", element)
    except NoSuchElementException as exception:
        print ('Click on ' +  fname + '  ID B not successful')
    except TimeoutException:
        pass

#USED 1
def click_operation_B_Xpath(fname):
    WebDriverWait(config_webdriver_manager.driver, 10).until(lambda d: d.execute_script('return document.readyState') == 'complete')
    try:
        element=WebDriverWait(config_webdriver_manager.driver, 20).until(EC.element_to_be_clickable((By.XPATH, fname))).click()
        config_webdriver_manager.driver.find_element("xpath", fname).sendKeys('\uE035')
    except NoSuchElementException as exception:
        print ('Click on ' + fname+ ' B XPath  not successful')
    except TimeoutException:
        pass

#USED 1
def click_operation_CSS(fname):
    WebDriverWait(config_webdriver_manager.driver, 10).until(lambda d: d.execute_script('return document.readyState') == 'complete')
    try:
        element=WebDriverWait(config_webdriver_manager.driver, 20).until(EC.element_to_be_clickable((By.CSS_SELECTOR, fname))).click()
        config_webdriver_manager.driver.find_element("CSS_SELECTOR", fname).sendKeys('\uE035')
    except NoSuchElementException as exception:
        print ('Click on ' + fname+ ' B XPath  not successful')
    except TimeoutException:
        pass



def click_operation_classname(fname):
    WebDriverWait(config_webdriver_manager.driver, 10).until(lambda d: d.execute_script('return document.readyState') == 'complete')
    try:
        element=WebDriverWait(config_webdriver_manager.driver, 20).until(EC.element_to_be_clickable((By.CLASS_NAME, fname))).click()
        config_webdriver_manager.driver.find_element("classname", fname).sendKeys('\uE035')

        config_webdriver_manager.driver.find_element("xpath", fname).sendKeys('\uE035')
    except NoSuchElementException as exception:
        print ('Click on ' + fname+ ' B XPath  not successful')
    except TimeoutException:
        pass

#USED 3
def click_operation_css(fname):
    WebDriverWait(config_webdriver_manager.driver, 10).until(lambda d: d.execute_script('return document.readyState') == 'complete')
    global A
    try:
        element= config_webdriver_manager.driver.find_element(By.CSS_SELECTOR, fname)
        config_webdriver_manager.driver.execute_script("arguments[0].click();", element)
        A=True
    except NoSuchElementException as exception:
        print ('Click on ' + fname+ ' with css  not successful')
        A=False

# new 
def click_operation_xpath(fname):
     WebDriverWait(config_webdriver_manager.driver, 10).until(lambda d: d.execute_script('return document.readyState') == 'complete')
     try:
         element = WebDriverWait(config_webdriver_manager.driver, 20).until(EC.element_to_be_clickable((By.XPATH, fname)))
         config_webdriver_manager.driver.execute_script("arguments[0].click();", element)
     except TimeoutException:
         pass
     except NoSuchElementException:
         print('Click on ' + fname + ' XPath not successful')
xpath_selector = "//button[@class='text-5']/i[@class='icon-outline-close text-5 text-primary-dark']"
click_operation_xpath(xpath_selector)

#USED 4
def write_operation_b_xpath(fname, Input):
    WebDriverWait(config_webdriver_manager.driver, 10).until(lambda d: d.execute_script('return document.readyState') == 'complete')
    global A
    try:
        element=WebDriverWait(config_webdriver_manager.driver, 20).until(EC.element_to_be_clickable((By.XPATH, fname))).send_keys(Input)
        #driver.find_element("xpath", fname).send_keys(Input)
    except NoSuchElementException as exception:
        print ('Writing '+ Input + ' not successful with xpath celector')
        A=False

#NEW
def write_operation_selector(fname, Input):
    WebDriverWait(config_webdriver_manager.driver, 10).until(lambda d: d.execute_script('return document.readyState') == 'complete')
    global A
    try:
        element=WebDriverWait(config_webdriver_manager.driver, 20).until(EC.element_to_be_clickable((By.CSS_SELECTOR, fname))).send_keys(Input)
        #driver.find_element("xpath", fname).send_keys(Input)
    except NoSuchElementException as exception:
        print ('Writing '+ Input + ' not successful with css celector')
        A=False

#USED 18
def click_operation_Xpath(fname):
    WebDriverWait(config_webdriver_manager.driver, 10).until(lambda d: d.execute_script('return document.readyState') == 'complete')
    try:
        element=WebDriverWait(config_webdriver_manager.driver, 20).until(EC.element_to_be_clickable((By.XPATH, fname)))
        config_webdriver_manager.driver.execute_script("arguments[0].click();", element)
    except TimeoutException:
        pass
    except NoSuchElementException:
        print ('Click on ' + fname+ ' XPath not successful')
#USED 8
def clear_operation_A_xpath(fname):
    WebDriverWait(config_webdriver_manager.driver, 20).until(lambda d: d.execute_script('return document.readyState') == 'complete')
    try:
        print(" field name "+fname)
        element=WebDriverWait(config_webdriver_manager.driver, 20).until(EC.element_to_be_clickable((By.XPATH, fname))).send_keys(Keys.CONTROL + 'a', Keys.BACKSPACE)
    except NoSuchElementException as exception:
        print ('Clear  ' +fname+ '   not successful')
    except TimeoutException:
        pass

#USED 8
def clear_operation_A_CSS(fname):
    WebDriverWait(config_webdriver_manager.driver, 20).until(lambda d: d.execute_script('return document.readyState') == 'complete')
    try:  
        print(" field name "+fname)
        element=WebDriverWait(config_webdriver_manager.driver, 20).until(EC.element_to_be_clickable((By.CSS_SELECTOR, fname))).send_keys(Keys.CONTROL + 'a', Keys.BACKSPACE)
    except NoSuchElementException as exception:
        print ('Clear  ' +fname+ '   not successful')
    except TimeoutException:
        pass

#USED 6
def write_operation_xpath(fname, Input):
    WebDriverWait(config_webdriver_manager.driver, 10).until(lambda d: d.execute_script('return document.readyState') == 'complete')
    global A
    try:
        config_webdriver_manager.driver.find_element("xpath", fname).send_keys(Input)
    except NoSuchElementException as exception:
        print ('Writing '+ Input + ' not successful with xpath celector')
        A=False
    except TimeoutException:
        pass


#USED 6
def write_operation_classname(fname, Input):
    WebDriverWait(config_webdriver_manager.driver, 10).until(lambda d: d.execute_script('return document.readyState') == 'complete')
    global A
    try:
        config_webdriver_manager.driver.find_element("class", fname).send_keys(Input)

    except NoSuchElementException as exception:
        print ('Writing '+ Input + ' not successful with class selector')
        A=False
    except TimeoutException:
        pass

#TESTING PREVIOUSLY UNKNOWN SCROLL FUNCTION 
def scroll_operation_B_Xpath(fname):
    WebDriverWait(config_webdriver_manager.driver, 10).until(lambda d: d.execute_script('return document.readyState') == 'complete')
    try:
        element=WebDriverWait(config_webdriver_manager.driver, 20).until(EC.element_to_be_clickable((By.XPATH, fname))).location_once_scrolled_into_view
#        element=WebDriverWait(config_webdriver_manager.driver, 20).until(EC.element_to_be_clickable((By.XPATH, fname))).click()
#        config_webdriver_manager.driver.find_element("xpath", fname).sendKeys('\uE035')
    except NoSuchElementException as exception:
        print ('Scroll ' + fname+ ' B XPath  not successful')
    except TimeoutException:
        pass


def click_and_write_by_actionchain_XPATH(fname,Input):
    try:
        # Locate the element with placeholder "Fälligkeit"

        webelement = WebDriverWait(config_webdriver_manager.driver, 10).until(
            EC.presence_of_element_located((By.XPATH, fname))
        )
    except NoSuchElementException as e:
        print("Element not found:", e)

    relax()
    actions = ActionChains(config_webdriver_manager.driver)
    time.sleep(random.randint(2,6))
    actions.click(webelement).send_keys(Input).send_keys(Keys.RETURN).perform()
    relax()


# for writing the results to Excel
def df_to_excel(df, ws, header=True, index=True, startrow=0, startcol=0):
    """Write DataFrame df to openpyxl worksheet ws"""

    rows = dataframe_to_rows(df, header=header, index=index)

    for r_idx, row in enumerate(rows, startrow + 1):
        for c_idx, value in enumerate(row, startcol + 1):
             ws.cell(row=r_idx, column=c_idx).value = value


#RELAX
def relax():
    WebDriverWait(config_webdriver_manager.driver, 10).until(lambda d: d.execute_script('return document.readyState') == 'complete')
#ZOOM 40%
def zoom40androlldown():
    config_webdriver_manager.driver.execute_script("document.body.style.zoom='40%'")# works with Chrome
    config_webdriver_manager.driver.execute_script ("window.scrollTo(0,document.body.scrollHeight);")
def zoom40():
    config_webdriver_manager.driver.execute_script("document.body.style.zoom='40%'")# works with Chrome


# write to input field Anleihe auslaufzeit
def write_to_input_field_css_auslaufzeit_AI(date,locatorname, number_of_field):
    time.sleep(random.randint(1,3))
    WebDriverWait(config_webdriver_manager.driver, 10).until(lambda d: d.execute_script('return document.readyState') == 'complete')
    inputs = WebDriverWait(config_webdriver_manager.driver, 10).until(
        EC.presence_of_all_elements_located((By.CSS_SELECTOR,locatorname))
    )
    # Click, clear, and send keys to the SECOND input
    input_variable = inputs[number_of_field]  #  Index 0 = first input, Index 1 = second input
    time.sleep(random.randint(1,2))
    input_variable.click()
    time.sleep(random.randint(1,2))
    input_variable.send_keys(Keys.CONTROL + "a")
    time.sleep(random.randint(1,2))
    input_variable.send_keys(Keys.DELETE)
    time.sleep(random.randint(1,2))
    input_variable.send_keys(date, Keys.RETURN)
    WebDriverWait(config_webdriver_manager.driver, 10).until(lambda d: d.execute_script('return document.readyState') == 'complete')
    time.sleep(random.randint(1,3))

