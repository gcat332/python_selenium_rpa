import sys
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from subprocess import CREATE_NO_WINDOW
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
from webdriver_manager.chrome import ChromeDriverManager
from selenium.common.exceptions import NoSuchElementException

sliderRadioID = 'j_id0:formDirect:pageBlock1:pageBlockSection1:j_id65:j_id70:3'
sliderRadioID2 = 'j_id0:formDirect:pageBlock1:pageBlockSection1:j_id65:j_id70:0'
ScrapRadioID2 = 'j_id0:formDirect:pageBlock1:pageBlockSection4:j_id124:j_id128:0'
hgsaRadioID = 'j_id0:formDirect:pageBlock1:pageBlockSection4:j_id92:j_id96:2'
hgsaRadioID2 = 'j_id0:formDirect:pageBlock1:pageBlockSection4:j_id92:j_id96:1'
uaiRadioID = 'j_id0:formDirect:pageBlock1:pageBlockSection1:j_id82:j_id87:3'
screenRadioID = 'j_id0:formDirect:pageBlock1:pageBlockSection1:j_id82:j_id87:4'
scrapRadioID = 'j_id0:formDirect:pageBlock1:pageBlockSection1:j_id82:j_id87:5'
rtvRadioID = 'j_id0:formDirect:pageBlock1:pageBlockSection1:j_id82:j_id87:1'
rtvOptionInternalRadioID = 'j_id0:formDirect:pageBlock1:pageBlockSection4:j_id113:j_id117:0'
SubmitButtonID = '//*[@id="j_id0:formDirect:pageBlock1:j_id34:bottom"]/input[1]'
OptionSliderID = 'j_id0:formDirect:pageBlock1:pageBlockSection4:j_id92:j_id96:0'
approveButtonXpath = '//a[@class="actionLink"][2]'
approve2ButtonXpath = '//input[@title="Approve"]'
QuerySelectwaitUAI = 'div[id="j_id0:formDirect:pageBlock1:pageBlockSection4"] div table tbody tr th.first.last'
saveButtonCss = 'input[name="j_id0:formDirect:pageBlock1:j_id34:j_id35"]'

#<input type="submit" name="j_id0:formDirect:pageBlock1:j_id34:j_id35" value="Save" style="font-size:11px;" class="btn">

log = []

def quit():
    sys.exit()

def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.dirname(__file__)
    return os.path.join(base_path, relative_path)

def mr_dispose(account,password,fp) :
  noti_gid.setText("    Processing...")
  chrome_service = Service(ChromeDriverManager().install())
  chrome_service.creationflags = CREATE_NO_WINDOW
  chrome_options = Options()
  df = read_csv(fp)
  df.columns = ['DIR','DISPOSE','MOVE FROM','MOVE TO','REASON','PLANNER','SCARPCODE']
  df['Status'] = ''
  browser = webdriver.Chrome(service=chrome_service,chrome_options=chrome_options)
  if (cb2.isChecked()):
   browser.set_window_position(-3000, 0)
  browser.get(process.env.SALEFORCE_URL)
  browser.find_element(By.ID,'username').send_keys(account)
  browser.find_element(By.ID,'password').send_keys(password)
  browser.find_element(By.ID,'ssoSubmit').click()

  WebDriverWait(browser,200).until(lambda browser: browser.find_elements(By.XPATH,'//div[@class="validation bg-danger"]') or browser.find_elements(By.ID,"phSearchInput"))
  if browser.find_elements(By.XPATH,'//div[@class="validation bg-danger"]'):  
   noti_gid.setStyleSheet(css5)
   noti_gid.setText("    Invalid ID/Password")
   browser.quit()
  else :
   QMessageBox.information(window,'ยืนยันอีกรอบ ดาต้าถูกแล้วใช่ไหม',' กดเพื่อเริ่มได้เลย! ')
   for i in range(len(df)):
    mr = df.iloc[i][0]
    judge = df.iloc[i][1].upper()
    movefrom = df.iloc[i][2].upper()
    moveto=df.iloc[i][3].upper()
    reason = df.iloc[i][4]
    planner = df.iloc[i][5]
    scrapcode = df.iloc[i][6]
    if len(reason) <= 20:
	    reason = reason+("."*(20-len(reason)))
 
    WebDriverWait(browser, 120).until(EC.presence_of_element_located((By.ID, 'phSearchInput')))
    search = browser.find_element(By.ID,'phSearchInput')
    search.clear()
    search.send_keys(mr)
    search.send_keys(Keys.RETURN) # hit return after you enter search text
    WebDriverWait(browser, 120).until(EC.presence_of_element_located((By.XPATH, '//*[@id="MAT_MR_Direct__c_body"]/table/tbody/tr[2]/th/a')))
    browser.find_element(By.XPATH,'//*[@id="MAT_MR_Direct__c_body"]/table/tbody/tr[2]/th/a').click()
    WebDriverWait(browser, 120).until(EC.presence_of_element_located((By.XPATH, '//*[@id="j_id0:formDirect:pageBlock1"]')))
    if browser.find_elements(By.CSS_SELECTOR,(saveButtonCss)):
       print ("Go!");
       if (judge=='UAI'):
            UAI(browser,moveto,movefrom,reason)
            df['Status'][i] = 'Done UAI'
            df.to_csv('MR_log.csv',index=False)
       elif (judge=='SCREEN'):
            SCN(browser,moveto,movefrom,reason)
            df['Status'][i] = 'Done Screen'
            df.to_csv('MR_log.csv',index=False)
       elif (judge=='SCRAP'):
            SCRAP(browser,moveto,movefrom,reason,planner)
            df['Status'][i] = 'Done Scrap'
            df.to_csv('MR_log.csv',index=False)
       elif (judge =='RTV'):
            RTV(browser,moveto,movefrom,reason,scrapcode)
            df['Status'][i] = 'Done RTV'
            df.to_csv('MR_log.csv',index=False)
       else:
            df['Status'][i] = 'Dispose not in lists (UAI/SCREEN/SCRAP/RTV)'
            df.to_csv('MR_log.csv',index=False)
    else:
       df['Status'][i] = 'No Save Button'
       df.to_csv('MR_log.csv',index=False)
    
   noti_gid.setStyleSheet(css4)
   noti_gid.setText(" Done.. อย่าลืมเช็ค log file นะ!")
   browser.quit()
   #quit()


def Template() :
   tp = QMessageBox()
   tp.setIcon(QMessageBox.Information)
   tp.setText("Please create CSV file with below template. (Click Deatail for view template)")
   
   tp.setWindowTitle("Template MR  :                                                                                                                  ")
   tp.setDetailedText("""    DIR-NUMBER | DISPOSE |  MOVE FROM |  MOVE TO  | DISPOSAL REASON    | Planner(RTV)    |Scarp_Code(Scarp)  |
    ---------------------------------------------------- ---------------------------------------------------------
0   DIR-000001 | UAI         |   23-MRB      |  FG-SDET     | UAI Material use as normal        |                           |                     |
1   DIR-000002 | UAI         |   23-MRB      |  FG-SDET     | UAI Material use as normal        |                           |                     |
2   DIR-000003 | Scarp      |   23-MRB      |  81-MET      | UAI Material use as normal        |                            | GE               |
3   DIR-000004 | RTV         |   23-MRB      |  23-MRB      | RTV Slider issue pattern line       | Sumitra Panit   |                     |""")
   tp.setStandardButtons(QMessageBox.Ok)
   retval = tp.exec_()

def UAI(browser,moveto,movefrom,reason):
    WebDriverWait(browser, 120).until(EC.presence_of_element_located((By.ID, uaiRadioID)))
    browser.find_element(By.ID,uaiRadioID).click()
    WebDriverWait(browser, 5).until(EC.presence_of_element_located((By.CSS_SELECTOR, QuerySelectwaitUAI)))
    # sleep(3)
    browser.find_element(By.ID,sliderRadioID).click()
    browser.find_element(By.ID,hgsaRadioID).click()
    Select(browser.find_element(By.CSS_SELECTOR,'select[name="j_id0:formDirect:pageBlock1:pageBlockSection3:j_id156:j_id159"]')).select_by_visible_text(movefrom)
    Select(browser.find_element(By.CSS_SELECTOR,'select[name="j_id0:formDirect:pageBlock1:pageBlockSection3:j_id162:j_id165"]')).select_by_visible_text(moveto)
    browser.find_element(By.ID,'j_id0:formDirect:pageBlock1:pageBlockSection5:j_id361:j_id364').send_keys(reason)
    browser.find_element(By.CSS_SELECTOR,saveButtonCss).click()
    WebDriverWait(browser, 120).until(EC.presence_of_element_located((By.XPATH, approveButtonXpath)))
    browser.find_element(By.XPATH,approveButtonXpath).click()
    WebDriverWait(browser, 120).until(EC.presence_of_element_located((By.XPATH, approve2ButtonXpath)))
    browser.find_element(By.XPATH,approve2ButtonXpath).click()


def SCN(browser,moveto,movefrom,reason):
    browser.find_element(By.ID,screenRadioID).click()

    WebDriverWait(browser, 5).until(EC.presence_of_element_located(
        (By.CSS_SELECTOR, QuerySelectwaitUAI)))
    # sleep(3)
    browser.find_element(By.ID,sliderRadioID2).click()
    browser.find_element(By.ID,hgsaRadioID2).click()
    Select(browser.find_element(By.CSS_SELECTOR,
        'select[name="j_id0:formDirect:pageBlock1:pageBlockSection3:j_id156:j_id159"]')).select_by_visible_text(movefrom)
    Select(browser.find_element(By.CSS_SELECTOR,
        'select[name="j_id0:formDirect:pageBlock1:pageBlockSection3:j_id162:j_id165"]')).select_by_visible_text(moveto)
    browser.find_element(By.ID,
        'j_id0:formDirect:pageBlock1:pageBlockSection5:j_id361:j_id364').send_keys(reason)
    browser.find_element(By.CSS_SELECTOR,saveButtonCss).click()
    WebDriverWait(browser, 120).until(
        EC.presence_of_element_located((By.XPATH, approveButtonXpath)))
    browser.find_element(By.XPATH,approveButtonXpath).click()
    WebDriverWait(browser, 120).until(
        EC.presence_of_element_located((By.XPATH, approve2ButtonXpath)))
    browser.find_element(By.XPATH,approve2ButtonXpath).click()


def RTV(browser,moveto,movefrom,reason,planner):
        WebDriverWait(browser, 120).until(EC.presence_of_element_located((By.ID, rtvRadioID)))
        browser.find_element(By.ID,rtvRadioID).click()
        WebDriverWait(browser, 120).until(
            EC.element_to_be_clickable((By.ID, rtvOptionInternalRadioID)))
        browser.find_element(By.ID,sliderRadioID).click()
        browser.find_element(By.ID,OptionSliderID).click()
        browser.find_element(By.ID,rtvOptionInternalRadioID).click()
        browser.find_element(By.ID,
            'j_id0:formDirect:pageBlock1:pageBlockSection5:j_id361:j_id364').send_keys(reason)
        browser.find_element(By.CSS_SELECTOR,saveButtonCss).click()
        WebDriverWait(browser, 120).until(EC.presence_of_element_located(
            (By.ID, 'j_id0:formDirect:pageBlock1:pageBlockSection5:MAT_1:reqMaterialMultipicklist:j_id209:multiselectPanel:leftList')))
        Select(browser.find_element(By.CSS_SELECTOR,
            'select[id="j_id0:formDirect:pageBlock1:pageBlockSection5:MAT_1:reqMaterialMultipicklist:j_id209:multiselectPanel:leftList"]')).select_by_visible_text(planner)
        browser.find_element(By.ID,
            'j_id0:formDirect:pageBlock1:pageBlockSection5:MAT_1:reqMaterialMultipicklist:j_id209:btnRight').click()
        browser.find_element(By.CSS_SELECTOR,saveButtonCss).click()
        WebDriverWait(browser, 120).until(
            EC.presence_of_element_located((By.XPATH, approveButtonXpath)))
        browser.find_element(By.XPATH,approveButtonXpath).click()
        WebDriverWait(browser, 120).until(
            EC.presence_of_element_located((By.XPATH, approve2ButtonXpath)))
        browser.find_element(By.XPATH,approve2ButtonXpath).click()


def SCRAP(browser,moveto,movefrom,reason,scrapcode):
    WebDriverWait(browser, 120).until(
        EC.presence_of_element_located((By.ID, scrapRadioID)))
    browser.find_element(By.ID,scrapRadioID).click()
    WebDriverWait(browser, 120).until(
        EC.element_to_be_clickable((By.ID, ScrapRadioID2)))
    browser.find_element(By.ID,ScrapRadioID2).click()
    WebDriverWait(browser, 120).until(EC.presence_of_element_located(
            (By.ID, 'j_id0:formDirect:pageBlock1:pageBlockSection4:j_id132:j_id134')))
    if (scrapcode == 'GE'):
	    scrapdetail = 'GE = Scrap HGA Electrical'
    elif (scrapcode == 'SE'):
        scrapdetail = 'SE =Scrap SDET slider fail ET'
    elif (scrapcode == 'GM'):
        scrapdetail = 'GM = Scrap HGA Mechanical'
    elif (scrapcode == 'SM'):
        scrapdetail = 'SM =Scrap SDET slider fail Mech'
    else:
        browser.quit()
        print("Scrap_Code ERROR")
    Select(browser.find_element(By.CSS_SELECTOR,
        'select[name="j_id0:formDirect:pageBlock1:pageBlockSection3:j_id156:j_id159"]')).select_by_visible_text(movefrom)
    Select(browser.find_element(By.CSS_SELECTOR,
        'select[name="j_id0:formDirect:pageBlock1:pageBlockSection3:j_id162:j_id165"]')).select_by_visible_text(moveto)
    Select(browser.find_element(By.CSS_SELECTOR,'select[name="j_id0:formDirect:pageBlock1:pageBlockSection4:j_id132:j_id135"]')).select_by_visible_text(scrapdetail)
    browser.find_element(By.ID,'j_id0:formDirect:pageBlock1:pageBlockSection5:j_id361:j_id364').send_keys(reason)
    browser.find_element(By.CSS_SELECTOR,saveButtonCss).click()
    WebDriverWait(browser, 120).until(EC.presence_of_element_located((By.XPATH, approveButtonXpath)))
    browser.find_element(By.XPATH,approveButtonXpath).click()
    WebDriverWait(browser, 120).until(EC.presence_of_element_located((By.XPATH, approve2ButtonXpath)))
    browser.find_element(By.XPATH,approve2ButtonXpath).click()	


def call_file():
 root = Tk()
 root.withdraw()
 file_path = filedialog.askopenfilename(title = "Select file",filetypes = (("CSV Files","*.csv"),))
 df = read_csv(file_path)
 if(len(df.columns) != 7 ):
  noti_file.setText("Invalid input file please try again!")
 else :
  df.columns = ['DIR','DISPOSE','MOVE FROM','MOVE TO','REASON','PLANNER','SCARPCODE']
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
   mr_dispose(account,password,fp)
  else :
   noti_gid.setStyleSheet(css5)
   noti_gid.setText("    Invalid File_Path...")
 else :
  noti_gid.setStyleSheet(css5)
  noti_gid.setText("    Please input GID or Password!")

def inform_viwe_mode():
 if (cb2.isChecked()):
  noti_gid.setStyleSheet(css4)
  noti_gid.setText("   Disable browser mode")
 else:
  noti_gid.setStyleSheet(css5)
  noti_gid.setText("   Enable browser mode")


#main_app_window 
qAp = QApplication(sys.argv)
window = QWidget()

qAp.setWindowIcon(QIcon('soju.png'))
window.setWindowTitle('MR_Auto_Disposer_Remake V_1.0')
window.resize(700,330)
#window.adjustSize()

info1 = QLabel("""MR_Auto_Disposer_Remake V_1.0 created by Kasira C.""",window)
info1.setGeometry(20,0,660,30)
info2 = QLabel("""Please provide Seagate account and password, and select excel file which including data followed by template.""",window)
info2.setGeometry(20,30,660,30)
info3 = QLabel("""**input file must be CSV and contains information as following in order**""",window)
info3.setGeometry(20,60,660,30)

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
color: red;
font-family: Tahoma;
font-size: 15px;
background-color: #c0d9af;
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
info4.setGeometry(120,230,80,60)

noti_gid = QLabel(' ',window) 
noti_gid.setGeometry(200,230,300,60)
noti_gid.setStyleSheet(css6)

noti_file = QLabel(' ',window)
noti_file.setGeometry(180,140,500,30)

gid = QLineEdit(window)
gid.setPlaceholderText('Please enter your GID')
gid.setGeometry(20,100,140,30)

pw = QLineEdit(window)
pw.setPlaceholderText('Please enter your password')
pw.setEchoMode(QLineEdit.Password)
pw.setGeometry(170,100,140,30)

bt1 = QPushButton(QIcon('soju.png'),'เลือก Excel file..',window)
bt1.setGeometry(20,140,150,30)
bt1.clicked.connect(call_file)

bt2 = QPushButton('Submit',window)
bt2.setGeometry(20,180,150,30)
bt2.clicked.connect(check_gid_password)

cb2 = QCheckBox("Disable View Mode",window)
cb2.setGeometry(340,180,170,30)
cb2.stateChanged.connect(inform_viwe_mode)

btlog = QPushButton('Template',window)
btlog.setGeometry(625,305,70,20)
btlog.clicked.connect(Template)

window.show()
qAp.exec_()