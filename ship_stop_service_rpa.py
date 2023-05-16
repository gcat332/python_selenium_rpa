import sys
from pandas import read_csv
from PyQt5.QtCore import *
from PyQt5.QtGui import *
from PyQt5.QtWidgets import *
from time import sleep
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
from tkinter import Tk
from tkinter import filedialog
from getpass import getpass
import time
import os
from os import path
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from subprocess import CREATE_NO_WINDOW
import sys

log = []

def quit():
    sys.exit()

def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.dirname(__file__)
    return os.path.join(base_path, relative_path)


def call_file():
 root = Tk()
 root.withdraw()
 file_path = filedialog.askopenfilename(title = "Select file",filetypes = (("CSV Files","*.csv"),))
 df = read_csv(file_path)
 if(len(df.columns) != 1 ):
  noti_file.setText("Invalid input file please try again!")
 else :
  df.columns = ['SHTT']
  noti_file.setText(file_path)
  return file_path
 
def check_gid_password() :
 if ( gid.text() != '' and pw.text() != '') :
  fp = noti_file.text()
  if (path.isfile(fp)):
   account = gid.text()
   password = pw.text()
   mes = gid.text() + " has logged-in!"
   noti_gid.setStyleSheet(css4)
   noti_gid.setText(mes)
   shtt_release(account,password,fp)
  else :
   noti_gid.setStyleSheet(css5)
   noti_gid.setText("    Invalid File_Path...")
 else :
  noti_gid.setStyleSheet(css5)
  noti_gid.setText("    Please input GID or Password!")

def update_log() :
   log = QMessageBox()
   log.setIcon(QMessageBox.Information)
   log.setText("Update Version 1.2.1 (25/04/2023)")
   log.setWindowTitle("Updated log")
   log.setDetailedText("""
   V1.2.1 (25/04/2023) - add auto name refresh
   V1.2.0 (07/09/2022) - bug fixed
   V1.0.0 (02/09/2022) - full released version""")
   log.setStandardButtons(QMessageBox.Ok)
   retval = log.exec_()
  
def shtt_release(account,password,fp):
  noti_gid.setText("    SHTT releasing...")
  chrome_service = Service(ChromeDriverManager().install())
  chrome_service.creationflags = CREATE_NO_WINDOW
  browser = webdriver.Chrome(service=chrome_service)
  #browser.set_window_position(-3000, 0)
  df = read_csv(fp)
  df.columns = ['SHTT']
  df['Status'] = ''
  browser.get(process.env.SSV_URL)
  if browser.find_elements(By.XPATH,'//*[@id="details-button"]'):
   browser.find_element(By.XPATH,'//*[@id="details-button"]').click()
   browser.find_element(By.XPATH,'//*[@id="proceed-link"]').click()
  else :
   browser.quit()
   noti_gid.setStyleSheet(css5)
   noti_gid.setText("    404-Error")
  browser.find_element(By.ID,'username').send_keys(account)
  browser.find_element(By.ID,'password').send_keys(password)
  browser.find_element(By.ID,'ssoSubmit').click()
  time.sleep(3);
  if browser.find_elements(By.XPATH,'//button[@id="ssoSubmit"]'):  
   noti_gid.setStyleSheet(css5)
   noti_gid.setText("    Invalid GID/Password")
   browser.quit()
  else :
   QMessageBox.information(window,'พร้อมแล้ว!',' กดเพื่อเริ่มได้เลย! ')  
   for i in range(len(df)):
    SHTT = df.iloc[i][0]
    WebDriverWait(browser, 120).until(EC.presence_of_element_located((By.XPATH, '//input[@id="ctl00_cCenter_Sec1_GenInf1_tbSHN"]')))
    browser.find_element(By.XPATH,'//input[@id="ctl00_cCenter_Sec1_GenInf1_tbSHN"]').clear()
    browser.find_element(By.XPATH,'//input[@id="ctl00_cCenter_Sec1_GenInf1_tbSHN"]').send_keys(SHTT)
    browser.find_element(By.XPATH,'//input[@id="ctl00_cCenter_Sec1_GenInf1_bnSearchPN"]').click()
    time.sleep(2);
    link = WebDriverWait(browser, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="ctl00_cCenter_Sec1_GenInf1_gvView_ctl02_hplSHN"]'))).get_attribute('href')
    link = link[-12:]
    if ( link == SHTT) :
     browser.find_element(By.XPATH,'//*[@id="ctl00_cCenter_Sec1_GenInf1_gvView_ctl02_hplSHN"]').click();
    else :
     df['Status'][i] = 'Not Found SH-TT'
     df.to_csv('dispose_log_file.csv',index=False)
     continue
    main_tab= browser.window_handles[0]
    shtt_tab= browser.window_handles[1]
    WebDriverWait(browser, 120)
    browser.switch_to.window(shtt_tab)
    if browser.find_elements(By.XPATH,'//input[@id="UC_Login1_txtGID"]'):
     browser.find_element(By.XPATH,'//input[@id="UC_Login1_txtGID"]').send_keys(account)
     browser.find_element(By.XPATH,'//input[@id="UC_Login1_txtPWD"]').send_keys(password)
     browser.find_element(By.XPATH,'//input[@id="UC_Login1_ImageButton1"]').click()
    SHTT_STATUS = browser.find_element(By.XPATH,'//*[@id="ctl00_cCenter_Sec1_GenInf1_lbStatus"]').text
    if SHTT_STATUS == "Full Released" :
     dispose = browser.find_element(By.XPATH,'//textarea[@id="ctl00_cCenter_Sec6_Close_taComment"]').get_attribute('disabled')
     if dispose == 'true':
      browser.find_element(By.XPATH,'//input[@id="ctl00_cCenter_Sec1_GenInf1_btnRefName"]').click()
     WebDriverWait(browser, 120).until(EC.presence_of_element_located((By.XPATH,'//input[@id="ctl00_cCenter_Sec6_Close_btnSubmitPE"]')))
     browser.find_element(By.XPATH,'//textarea[@id="ctl00_cCenter_Sec6_Close_taComment"]').send_keys("Continue monitor and looking for opportunity for improvement")
     browser.find_element(By.XPATH,'//input[@id="ctl00_cCenter_Sec6_Close_btnSubmitPE"]').click()
     df['Status'][i] = 'Closed Done'
     df.to_csv('dispose_log_file.csv',index=False)	
    else :     
     df['Status'][i] = 'No Submit Button'
     df.to_csv('dispose_log_file.csv',index=False)
    browser.close()
    browser.switch_to.window(main_tab)
   browser.quit()	
   noti_gid.setStyleSheet(css4)
   noti_gid.setText("   Done.. อย่าลืมเช็ค log file นะ!")
   QMessageBox.information(window,'ทำเสร็จแล้ว!','log file ชื่อว่า  :  dispose_log_file.csv') 
   
 
#main_app_window 
qAp = QApplication(sys.argv)
window = QWidget()

qAp.setWindowIcon(QIcon('cat.png'))
window.setWindowTitle('MGR SHTT Disposer v.1.2.1')
window.resize(700,370)

info1 = QLabel("""MGR SHTT Disposer v.1.2.1 (Full Released)""",window)
info1.setGeometry(20,0,660,30)
info2 = QLabel("""*input file must be CSV and contains information as following in order**""",window)
info2.setGeometry(20,30,660,30)
info3 = QLabel("""      SHTT-NUMBER   |  
    ----------------------
   0   TTP-000001     | 
   1   TTK-000002     | 
   2   SH-TT-00003   | """,window)
info3.setGeometry(20,60,660,90)
css1 = '''
color: blue;
font-family: Tahoma;
font-size: 15px;
background-color: #c0d9af;
'''
css2 = '''
color: black;
font-family: Tahoma;
font-size: 12px;
'''
css3 = '''
color: white;
font-family: Tahoma;
font-size: 13px;
background-color: black;
'''
css4 = '''
color: black;
font-family: Tahoma;
font-size: 15px;
background-color: #c2f596;
font-weight: bold;
'''
css5 = '''
color: white;
font-family: Tahoma;
font-size: 15px;
background-color: red;
font-weight: bold;
'''
css6 = '''
color: white;
font-family: Tahoma;
font-size: 15px;
background-color: grey;
font-weight: bold;
'''

cat = resource_path("cat.png")

info1.setStyleSheet(css1)
info2.setStyleSheet(css2)
info3.setStyleSheet(css3)

info4 = QLabel("""Status : """,window)
info4.setGeometry(120,290,80,60)

noti_gid = QLabel('   Stand by...',window) 
noti_gid.setGeometry(200,290,300,60)
noti_gid.setStyleSheet(css6)

noti_file = QLabel(' ',window)
noti_file.setGeometry(180,210,500,30)

gid = QLineEdit(window)
gid.setPlaceholderText('Please enter your GID')
gid.setGeometry(20,170,140,30)

pw = QLineEdit(window)
pw.setPlaceholderText('Please enter your password')
pw.setEchoMode(QLineEdit.Password)
pw.setGeometry(170,170,140,30)

bt1 = QPushButton(QIcon('cat.png'),'เลือก Excel file..',window)
bt1.setGeometry(20,210,150,30)
bt1.clicked.connect(call_file)

bt2 = QPushButton('Submit',window)
bt2.setGeometry(20,250,150,30)
bt2.clicked.connect(check_gid_password)


window.show()
qAp.exec_()