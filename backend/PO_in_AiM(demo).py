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

    def log_po(self,po_no,supplier_no,item,line_total,WO,phase,material=True,first_PO=True):
        try:
            if first_PO:
                self.driver.find_element(By.ID, "mainForm:menuListMain:new_PO_VIEW").click()
            else:
                self.driver.find_element(By.ID, "mainForm:buttonPanel:new").click()
            """PO main page"""
            self.driver.find_element(By.ID, "mainForm:PO_EDIT_content:ae_i_poe_e_description").send_keys(item)
            self.driver.find_element(By.ID, "mainForm:PO_EDIT_content:contractorZoom:contractorZoom0").send_keys(supplier_no)
            self.driver.find_element(By.ID, "mainForm:PO_EDIT_content:contractorZoom:contractorZoom1").send_keys("1")
            self.driver.find_element(By.ID, "mainForm:PO_EDIT_content:termsZoom:termsZoom01").send_keys("1")
            self.driver.find_element(By.ID, "mainForm:PO_EDIT_content:poStatusTypeZoom:level0").send_keys("e-pro")
            self.driver.find_element(By.ID, "mainForm:PO_EDIT_content:poStatusZoom:level0").send_keys("open")
            self.driver.find_element(By.ID, "mainForm:PO_EDIT_content:defaultWoZoom:defaultWorkOrder").send_keys(WO)
            self.driver.find_element(By.ID, "mainForm:PO_EDIT_content:defaultWoZoom:defaultPhase").send_keys(phase)
            self.driver.find_element(By.ID, "mainForm:PO_EDIT_content:disbDefaultsLineItem").click()
            dropdown = self.driver.find_element(By.ID, "mainForm:PO_EDIT_content:disbDefaultsLineItem")
            dropdown.find_element(By.XPATH, "//option[. = 'Service']").click()
            self.driver.find_element(By.ID, "mainForm:PO_EDIT_content:disbDefaultsLineItem").click()
            self.driver.find_element(By.CSS_SELECTOR, "#mainForm\\3APO_EDIT_content\\3AtermsZoom\\3AtermsZoom01_button > .halflings").click()
            time.sleep(0.5)
            """Line item"""
            self.driver.find_element(By.ID, "mainForm:PO_EDIT_content:oldPoLineItemsList:addLineItemButton").click()
            self.driver.find_element(By.ID, "mainForm:PO_LINE_ITEM_EDIT_content:ae_i_poe_d_vend_dsc").send_keys(WO+" - "+phase)
            self.driver.find_element(By.ID, "mainForm:PO_LINE_ITEM_EDIT_content:amountValueServices").clear()
            self.driver.find_element(By.ID, "mainForm:PO_LINE_ITEM_EDIT_content:amountValueServices").send_keys(line_total)
            self.driver.find_element(By.ID, "mainForm:PO_LINE_ITEM_EDIT_content:subledgerValue").click()
            dropdown = self.driver.find_element(By.ID, "mainForm:PO_LINE_ITEM_EDIT_content:subledgerValue")
            if material:
                dropdown.find_element(By.XPATH, "//option[. = 'Material']").click()
            else:
                dropdown.find_element(By.XPATH, "//option[. = 'Labor']").click()
            self.driver.find_element(By.ID, "mainForm:PO_LINE_ITEM_EDIT_content:subledgerValue").click()
            self.driver.find_element(By.ID, "mainForm:buttonPanel:done").click()
            """UDF"""
            self.driver.find_element(By.ID, "mainForm:sideButtonPanel:moreMenu_3").click()
            self.driver.find_element(By.ID, "mainForm:PO_UDF_EDIT_content:ae_i_poe_e_udf_custom001").send_keys(po_no)
            self.driver.find_element(By.ID, "mainForm:buttonPanel:done").click()

            self.driver.find_element(By.ID, "mainForm:buttonPanel:save").click()
        except:
            #TODO: deal with exceptions
            time.sleep(100)

if __name__ == '__main__':

    new_po = PurchaseOrder()
    new_po.setup_method()
    new_po.login()

    sheet = pd.read_excel("..\excel input\order_line_list.xlsx", dtype=str)
    for i in range(3,6):
        start_time = time.time()
        po_no,_, supplier_no,_, item, line_total, WO, phase,CP = sheet.iloc[i].values
        if pd.notna(CP):
            #TODO: handle the PO with CP number
            continue
        first_po = True if i==3 else False
        new_po.log_po(po_no, supplier_no, item, line_total, WO, phase,material=True,first_PO=first_po)
        print ("row {} is processed, time: {}s".format(i+1,time.time()-start_time))


