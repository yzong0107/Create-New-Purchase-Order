__author__ = "Tim Zong (yzong@ualberta.ca)"

import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support import expected_conditions
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.desired_capabilities import DesiredCapabilities
import time
import getpass


class PurchaseOrder():
    def setup_method(self):
        self.driver = webdriver.Chrome()
        self.vars = {}

    def teardown_method(self):
        self.driver.quit()

    def login(self):
        self.driver.get("https://www.aimdemo.ualberta.ca/fmax/screen/WORKDESK")
        self.driver.set_window_size(1900, 1020)
        username = input('Enter your username: ')
        password = getpass.getpass('Enter your password : ')
        self.driver.find_element(By.ID, "username").send_keys(username)
        self.driver.find_element(By.ID, "password").send_keys(password)
        self.driver.find_element(By.ID, "login").click()
        self.driver.find_element(By.ID, "mainForm:menuListMain:PURCHASING").click()


if __name__ == '__main__':
    new_po = PurchaseOrder()
    new_po.setup_method()
    new_po.login()
