import selenium
import os

import time
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.action_chains import ActionChains
import tkinter as tk
from tkinter import filedialog
from pathlib import Path
from selenium.webdriver.chrome.service import Service
#from webdriver_manager.chrome import ChromeDriverManager
import sys
import logging
from datetime import datetime



# Set up logging
logging.basicConfig(filename='output.log', level=logging.INFO)



# Start Selenium 
PATH = "chromedriver.exe"
driver = webdriver.Chrome(PATH)
driver.maximize_window()


#driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()))

root = tk.Tk()
root.withdraw()
root.focus_force()
root.deiconify()
root.geometry("10x10")


files = filedialog.askopenfilenames(filetypes=[("Excel files", "*.xlsx")])

root.lower() #pushing the deiconified window back
root.destroy() #destroying it as well

print("--------------")
print(files)
l = [Path(x).name.split('.')[0] for x in files]
print("--------------")
print(type(l))
print("--------------")
print(l)
print("--------------")
print (l)
print("--------------")
print("Number of files is:", len(l))

path_l = [Path(x).resolve() for x in files] 
print(path_l)

# -------------------------------------------------------------------------------------------------------------------------------------------------------lp = os.listdir(path='C:\codes\python\webscrapper_amazon\crm\client_name')
# -------------------------------------------------------------------------------------------------------------------------------------------------------l=[x.split('.')[0] for x in lp]

j=len(l)
i=0

def upload(*l):

    uid_ninja = WebDriverWait(driver, 30).until(
    EC.presence_of_element_located((By.XPATH,"/html/body/div[1]/div[4]/form/input[1]"))
    )
    uid_ninja.click()
    uid_ninja.send_keys("username")
    
    
    passwordofninjauploader = driver.find_element(By.XPATH,"/html/body/div[1]/div[4]/form/input[2]")
    passwordofninjauploader.send_keys("username")

    submit_button = driver.find_element(By.XPATH,"/html/body/div[1]/div[4]/form/input[4]")
    submit_button.click()

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
    Select(file_list).select_by_value("347")
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




def submain():
    for i in range(j+1):
        url = 'https://ninjacrm.com/'+l[i]+'/'
        driver.get(url)

        upload(l[i])
        uploadfile = driver.find_element(By.XPATH,"/html/body/div[1]/div[2]/div[1]/div[3]/div[2]/div[4]/div[1]/input")
        uploadfile.send_keys(str(path_l[i]))  ##uploadfile.send_keys("C:\codes\python\webscrapper_amazon\crm\client_name\\"+l[i]+".xlsx")
        

        time.sleep(3)

        donehogya = WebDriverWait(driver,120).until(                             #increasing the wait time for some stubborn data files that take more time to process
        EC.element_to_be_clickable((By.ID,"done_button1"))
        )

        uploaderror = driver.find_element(By.ID,"error_dialog1")
        
        print(datetime.now().strftime("\n(%H:%M:%S)")+' : Upload status at '+l[i]+' : '+uploaderror.text)

        time.sleep(2)

        donehogya.click()
        print('Data upload at '+l[i]+' completed...')
        root1.update()

        time.sleep(1)
        i+=1
   
def main():
    print(datetime.now().strftime("Date : %d-%m-%Y // Time : %H:%M:%S"))
    print("========RPA Running==========")
    print("--------------")
    print(type(l))
    print("--------------")
    print(l)
    print("--------------")
    print (l)
    print("--------------")
    print("Number of files is:", len(l))
    root1.update()
    try:
        submain()
    except IndexError:
        print("\nOperation completed.")
        driver.quit()
    except KeyboardInterrupt:
        print("Keyboard Interruption detected, RPA stopped")
    except DeprecationWarning:
        print("-")
    except Exception as e:
        print(e)


class OutputRedirector:
    def __init__(self, text_widget):
        self.text_widget = text_widget

    def write(self, message):
        self.text_widget.config(state=tk.NORMAL)
        self.text_widget.insert(tk.END, message)
        self.text_widget.config(state=tk.DISABLED)
        self.text_widget.see(tk.END)
        logging.info(message)
    def flush(self):
        pass

root1 = tk.Tk()
root1.title("Console")
root1.lower()

text_widget = tk.Text(root1, wrap=tk.WORD)
text_widget.pack(expand=True, fill=tk.BOTH)
text_widget.config(state=tk.DISABLED)

sys.stdout = OutputRedirector(text_widget)
root1.update()
main()
root1.update()
root1.mainloop()