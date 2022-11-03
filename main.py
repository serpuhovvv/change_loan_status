# Install requirements: pip install -r requirements.txt

import time
from selenium import webdriver
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver import ActionChains
from pathlib import Path
import time
from selenium import webdriver
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver import ActionChains
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
import time
from selenium.webdriver.support.wait import WebDriverWait
import pandas as pd
import time
from bs4 import BeautifulSoup
import requests
from datetime import date
from selenium.webdriver.support.select import Select


login = ''
password = ''


def wait_xpath(xpath):
    element = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable(
            (By.XPATH, xpath)
        )
    )
    return element


def wait_id(id):
    element = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable(
            (By.ID, id)
        )
    )
    return element


def wait_class(class_name):
    element = WebDriverWait(driver, 20).until(
        EC.element_to_be_clickable(
            (By.CLASS_NAME, class_name)
        )
    )
    return element


def wait_text(text):
    element = WebDriverWait(driver, 20).until(
        EC.element_to_be_clickable(
            (By.LINK_TEXT, text)
        )
    )
    return element


def wait_frame_id(id):
    WebDriverWait(driver, 30).until(EC.frame_to_be_available_and_switch_to_it((By.ID, id)))


def switch_to_frame(frame_reference):
    driver.switch_to.frame(frame_reference)


def switch_to_parent_frame():
    driver.switch_to.parent_frame()


def switch_to_default_content():
    driver.switch_to.default_content()


driver = webdriver.Chrome()
driver.maximize_window()
driver.get('https://portal.admortgage.com/Default.aspx')

wait_frame_id('contentFrame')

username = wait_xpath('//*[@id="UserName"]')
password = wait_xpath('//*[@id="Password"]')
login = wait_xpath('//*[@id="LoginButton"]')

username.send_keys(login)
password.send_keys(password)
login.click()
time.sleep(5)

df = pd.read_excel('C:/Users/serg.pudikov/QA Files/Change loan status.xlsx', sheet_name='Sheet1')
loannum = df['loannu']
print(loannum)
okloans = []
errloans = []


for each in loannum:
    try:
        switch_to_frame(0)
        switch_to_frame(0)
        wait_id('FilterRadioButtonList_0').click()
        time.sleep(1)
        wait_xpath('//*[@id="SearchTextBox"]').send_keys(each)
        time.sleep(2)
        wait_xpath('//*[@id="SearchButton"]').click()
        switch_to_default_content()
        time.sleep(2)

        switch_to_frame(0)
        switch_to_frame(0)
        switch_to_frame(0)
        time.sleep(2)
        action = ActionChains(driver)
        a = wait_xpath('//*[@id="PipelineRow1"]/td[2]')
        action.double_click(a)
        action.perform()
        switch_to_default_content()
        time.sleep(7)

        wait_id('Change Status').click()
        time.sleep(5)

        wait_frame_id('dialogframe')

        option = Select(wait_id('StatusListBox'))
        option.select_by_visible_text('PC Review Completed')
        time.sleep(2)

        switch_to_default_content()

        wait_xpath('//*[@id="btnOkay"]').click()
        time.sleep(7)

        okloans.append(each)
        df1 = pd.DataFrame(data={'ok loans': okloans})
        df1.to_excel('C:/Users/serg.pudikov/QA Files/okloans.xlsx')

        assert wait_xpath('//*[@id="Row30"]/td[2]').text == 'PC Review Completed'

    except Exception as ex:

        errloans.append(each)
        df2 = pd.DataFrame(data={'error loans': errloans})
        df2.to_excel('C:/Users/serg.pudikov/QA Files/errloans.xlsx')
        print(str(each) + ': Something went wrong')

    finally:
        switch_to_default_content()
        wait_id('ExitLoanli').click()
        time.sleep(5)
        driver.refresh()
