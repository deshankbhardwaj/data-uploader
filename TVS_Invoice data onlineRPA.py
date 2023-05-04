
import selenium
import os

import time
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.select import Select
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

from selenium.webdriver.common.action_chains import ActionChains
import tkinter as tk
from tkinter import filedialog
from pathlib import Path



 

PATH = "chromedriver.exe"
driver = webdriver.Chrome(PATH)

clientname = "bajwatvs"

root = tk.Tk()
root.withdraw()

files = filedialog.askopenfilenames(filetypes=[("Excel files", "*.xlsx")])
print("--------------")
print(files)
l = [Path(x).resolve() for x in files]
print("--------------")
print(type(l))
print("--------------")
print(l)
print("--------------")
print (l)
print("--------------")
print("Number of files is:", len(l))




# -------------------------------------------------------------------------------------------------------------------------------------------------------lp = os.listdir(path='C:\codes\python\webscrapper_amazon\crm\client_name')
# -------------------------------------------------------------------------------------------------------------------------------------------------------l=[x.split('.')[0] for x in lp]

j=len(l)
i=0


def auth():
    uid_ninja = WebDriverWait(driver, 30).until(
    EC.presence_of_element_located((By.XPATH,"/html/body/div[1]/div[4]/form/input[1]"))
    )
    uid_ninja.click()
    uid_ninja.send_keys("uuserID")
    
    
    passwordofninjauploader = driver.find_element(By.XPATH,"/html/body/div[1]/div[4]/form/input[2]")
    passwordofninjauploader.send_keys("password")

    submit_button = driver.find_element(By.XPATH,"/html/body/div[1]/div[4]/form/input[4]")
    submit_button.click()



def upload(*l):

    

    #----------------- making the process more direct so the process of google login is removed ---------------------------------------

    # ninjamenu = WebDriverWait(driver, 30).until(
    # EC.presence_of_element_located((By.XPATH,'//*[@id="gn-menu"]/li[1]/a'))  
    # )
    # ninjamenu.click()

    # time.sleep(2)

    # adminops = WebDriverWait(driver, 30).until(
    # EC.presence_of_element_located((By.XPATH,'/html/body/nav/div/div/ul/li[1]/nav/div/ul/li[1]/a'))
    # )
    # adminops.click()

    # uploadadmin =  WebDriverWait(driver, 30).until(
    # EC.presence_of_element_located((By.XPATH,"/html/body/div[1]/div/div/div/div[2]/div/a"))
    # )
    # uploadadmin.click()

    #----------------------------------------------------------------------------------------------------------------------------------

    file_list = WebDriverWait(driver, 30).until(
    EC.presence_of_element_located((By.ID,"file_type_list"))
    )
    Select(file_list).select_by_value("297")
    time.sleep(3)
    location = driver.find_element(By.XPATH,"/html/body/div[1]/div[2]/div[1]/div[3]/div[2]/div[1]")

    location.click()
    input = driver.find_element(By.XPATH,"/html/body/div[1]/div[2]/div[1]/div[3]/div[2]/div[1]/div/div/input")
    input.send_keys(Keys.ARROW_DOWN)
    input.send_keys(Keys.RETURN)
    date = driver.find_element(By.XPATH,"/html/body/div[1]/div[2]/div[1]/div[3]/div[2]/div[2]/input[1]")
    date.click()

    month_n = driver.find_element(By.XPATH,"/html/body/div[3]/div/div/select[1]")
    month_n.click()
    month_n = driver.find_element(By.XPATH,"/html/body/div[3]/div/div/select[1]/option[2]")
    month_n.click()

    year = driver.find_element(By.XPATH,"/html/body/div[3]/div/div/select[2]")
    year.click()
    year = driver.find_element(By.XPATH,"/html/body/div[3]/div/div/select[2]/option[6]")
    year.click()

    date = driver.find_element(By.XPATH,"/html/body/div[3]/table/tbody/tr[2]/td[2]/a")
    date.click()

        #end date below
    date = driver.find_element(By.XPATH,"/html/body/div[1]/div[2]/div[1]/div[3]/div[2]/div[2]/input[2]")
    date.click()

    month_n = driver.find_element(By.XPATH,"/html/body/div[3]/div/div/select[1]")
    month_n.click()
    month_n = driver.find_element(By.XPATH,"/html/body/div[3]/div/div/select[1]/option[2]")
    month_n.click()

    year = driver.find_element(By.XPATH,"/html/body/div[3]/div/div/select[2]")
    year.click()
    year = driver.find_element(By.XPATH,"/html/body/div[3]/div/div/select[2]/option[32]")
    year.click()

    date = driver.find_element(By.XPATH,"/html/body/div[3]/table/tbody/tr[5]/td[4]/a")
    date.click()

    uploadfile = driver.find_element(By.XPATH,"/html/body/div[1]/div[2]/div[1]/div[3]/div[2]/div[4]/div[1]/input")
    uploadfile.send_keys(str(l[i]))  ##uploadfile.send_keys("C:\codes\python\webscrapper_amazon\crm\client_name\\"+l[i]+".xlsx")
    

    time.sleep(3)

    donehogya = WebDriverWait(driver,3000).until(                             #increasing the wait time for some stubborn data files that take more time to process
    EC.element_to_be_clickable((By.ID,"done_button1"))
    )

    time.sleep(2)

    donehogya.click()

    time.sleep(1)


def main():

    for i in range(j+1):
      if i == 0 :  
        url = 'https://ninjacrm.com/'+clientname+'/'
        driver.get(url)
        auth()
        upload(l[i])
        i+=1
      else:
         upload(l[i])
         i+=1
    

main()
