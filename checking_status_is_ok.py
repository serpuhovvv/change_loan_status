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

username.send_keys('spudikov')
password.send_keys('ADMortgage2000!')
login.click()
time.sleep(5)

df = pd.read_excel('C:/Users/serg.pudikov/QA Files/Change loan status.xlsx', sheet_name='Sheet1')
loannum = df['loannu']
print(loannum)
errloans = []

for each in loannum:
    try:
        switch_to_frame(0)
        switch_to_frame(0)
        wait_id('FilterRadioButtonList_0').click()
        time.sleep(1)
        wait_xpath('//*[@id="SearchTextBox"]').send_keys(each)
        time.sleep(1)
        wait_xpath('//*[@id="SearchButton"]').click()
        switch_to_default_content()
        time.sleep(2)

        switch_to_frame(0)
        switch_to_frame(0)
        switch_to_frame(0)
        status = wait_xpath('//*[@id="PipelineRow1"]/td[6]')
        assert status.text == 'PC Review Completed'

    except Exception as ex:
        errloans.append(each)
        dferr = pd.DataFrame(data={'error loans': errloans})
        dferr.to_excel('C:/Users/serg.pudikov/QA Files/errloans_final.xlsx')
        print(each + ': Something went wrong')

    finally:
        switch_to_default_content()
        driver.refresh()
        time.sleep(2)