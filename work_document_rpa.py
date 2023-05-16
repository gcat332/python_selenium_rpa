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


def get_process_code(pc_code):
    if (pc_code == 'F'):
	    pc_code = 'F: Preload / Auto Preload (PLD)'
    elif (pc_code == 'G'):
        pc_code = 'G: SAAM / Auto SAAM (SAAM)'
    elif (pc_code == 'H'):
        pc_code = 'H: Fly Test (FLY)'
    elif (pc_code == 'I'):
        pc_code = 'I: Electrical Test (ET)'
    elif (pc_code == 'J'):
        pc_code = 'J: Sorter / Auto Sorter (SRT)'
    elif (pc_code == 'K'):
        pc_code = 'K: Gap Measure (GAP)'
    elif (pc_code == 'L'):
        pc_code = 'L: Cleaning (CLN)'
    elif (pc_code == 'N'):
        pc_code = 'N: Auto Load Machine (ALULM)'
    elif (pc_code == 'O'):
        pc_code = 'O: Auto Unload Machine (ALULM)'
    elif (pc_code == 'SS'):
        pc_code = 'SS: Solvent For SAMs (SOV)'
    return pc_code

def call_file():
 root = Tk()
 root.withdraw()
 file_path = filedialog.askopenfilename(title = "Select file",filetypes = (("CSV Files","*.csv"),))
 df = read_csv(file_path)
 if(len(df.columns) != 59 ):
  noti_file.setText("Invalid input file please try again!")
 else :
  df.columns = ['PRODUCT','WO_NAME','TAB','WF_3D','TRAY','TMWI','SDET_MARKING','SLD_PN','SLD_AAB','SET_PN','SET_VER','STEP1_TSR','STEP1_PC','STEP2_TSR','STEP2_PC','STEP3_TSR','STEP3_PC','STEP4_TSR','STEP4_PC','STEP5_TSR','STEP5_PC','STEP6_TSR','STEP6_PC','STEP7_TSR','STEP7_PC','STEP8_TSR','STEP8_PC','MARKING_BIN0','MARKING_BIN1','MARKING_BIN2','MARKING_BIN3','MARKING_BIN4','MARKING_BIN5','MARKING_BIN6','MARKING_BIN7','ET_SORT0','ET_SORT1','ET_SORT2','ET_SORT3','ET_SORT4','ET_SORT5','ET_SORT6','ET_SORT7','ET_SORT8','CD_A','CD_G','CD_H','CD_I','CD_M','CD_Q','CD_L','CD_S','CD_J','CD_T','CD_OTHERS','ET_FAIL','ET_NULL','ET_P','DESC']
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
   wo_create(account,password,fp)
  else :
   noti_gid.setStyleSheet(css5)
   noti_gid.setText("    Invalid File_Path...")
 else :
  noti_gid.setStyleSheet(css5)
  noti_gid.setText("    Please input GID or Password!")

def update_log() :
   log = QMessageBox()
   log.setIcon(QMessageBox.Information)
   log.setText("Update Version 0.4 (02/09/2022)")
   log.setWindowTitle("Updated log")
   log.setDetailedText("""
   V0.4.0 (02/09/2022) - fix major bug 
   V0.3.3 (26/08/2022) - add error_log_file
   V0.3.0 (24/08/2022) - add wo create disable function""")
   log.setStandardButtons(QMessageBox.Ok)
   retval = log.exec_()
  
def wo_create(account,password,fp):
  noti_gid.setText("    WO creating...")

  chrome_service = Service(ChromeDriverManager().install())
  chrome_service.creationflags = CREATE_NO_WINDOW
  chrome_options = Options()
  if (cb2.isChecked()):
   chrome_options.add_argument("headless")
  df = read_csv(fp)
  df.columns = ['PRODUCT','WO_NAME','TAB','WF_3D','TRAY','TMWI','SDET_MARKING','SLD_PN','SLD_AAB','SET_PN','SET_VER','STEP1_TSR','STEP1_PC','STEP2_TSR','STEP2_PC','STEP3_TSR','STEP3_PC','STEP4_TSR','STEP4_PC','STEP5_TSR','STEP5_PC','STEP6_TSR','STEP6_PC','STEP7_TSR','STEP7_PC','STEP8_TSR','STEP8_PC','MARKING_BIN0','MARKING_BIN1','MARKING_BIN2','MARKING_BIN3','MARKING_BIN4','MARKING_BIN5','MARKING_BIN6','MARKING_BIN7','ET_SORT0','ET_SORT1','ET_SORT2','ET_SORT3','ET_SORT4','ET_SORT5','ET_SORT6','ET_SORT7','ET_SORT8','CD_A','CD_G','CD_H','CD_I','CD_M','CD_Q','CD_L','CD_S','CD_J','CD_T','CD_OTHERS','ET_FAIL','ET_NULL','ET_P','DESC']
  df['Status'] = ''
  browser = webdriver.Chrome(service=chrome_service,chrome_options=chrome_options)
  browser.get(processs.env.WO_URL)
  browser.switch_to.frame('appDetail')
  browser.find_element(By.NAME,'gid').send_keys(account)
  browser.find_element(By.NAME,'pwd').send_keys(password)
  browser.find_element(By.NAME,'imgSubmit').click()
  WebDriverWait(browser,40)
  if browser.find_elements(By.XPATH,'//html/body/table[2]'):
   noti_gid.setStyleSheet(css5)
   noti_gid.setText("    Invalid ID/Password")
   browser.quit()
#check_login	  
  else :
   WebDriverWait(browser,40).until(lambda browser: browser.find_elements(By.XPATH,'//html/body/table/tbody/tr/td[1]/center') or browser.find_elements(By.CLASS_NAME,"menuUnhighlight"))
   QMessageBox.information(window,'กำลังทำงานจ้า',' กดเพื่อเริ่มได้เลย! ')
#check_param	
   for i in range(len(df)):
    PRODUCT = df.iloc[i][0].upper()
    WO_NAME = df.iloc[i][1].upper()
    TAB = df.iloc[i][2].upper()
    WF_3D = df.iloc[i][3].upper()
    TRAY = df.iloc[i][4]
    TRAY = TRAY.astype(str)
    TMWI = df.iloc[i][5].upper()
    SDET_MARKING = df.iloc[i][6].upper()
    SLD_PN = df.iloc[i][7]
    SLD_PN = SLD_PN.astype(str)
    SLD_AAB = df.iloc[i][8].upper()
    SET_PN = df.iloc[i][9]
    SET_PN = SET_PN.astype(str)
    SET_VER = df.iloc[i][10].upper()
    STEP1_TSR = df.iloc[i][11]
    STEP1_PC = df.iloc[i][12]
    STEP2_TSR = df.iloc[i][13]
    STEP2_PC = df.iloc[i][14]
    STEP3_TSR = df.iloc[i][15]
    STEP3_PC = df.iloc[i][16]
    STEP4_TSR = df.iloc[i][17]
    STEP4_PC = df.iloc[i][18]
    STEP5_TSR = df.iloc[i][19]
    STEP5_PC = df.iloc[i][20]
    STEP6_TSR = df.iloc[i][21]
    STEP6_PC = df.iloc[i][22]
    STEP7_TSR = df.iloc[i][23]
    STEP7_PC = df.iloc[i][24]
    STEP8_TSR = df.iloc[i][25]
    STEP8_PC = df.iloc[i][26]
    MARKING_BIN0 = df.iloc[i][27]
    MARKING_BIN1 = df.iloc[i][28]
    MARKING_BIN2 = df.iloc[i][29]
    MARKING_BIN3 = df.iloc[i][30]
    MARKING_BIN4 = df.iloc[i][31]
    MARKING_BIN5 = df.iloc[i][32]
    MARKING_BIN6 = df.iloc[i][33]
    MARKING_BIN7 = df.iloc[i][34]
    ET_SORT0 = df.iloc[i][35]
    ET_SORT1 = df.iloc[i][36]
    ET_SORT2 = df.iloc[i][37]
    ET_SORT3 = df.iloc[i][38]
    ET_SORT4 = df.iloc[i][39]
    ET_SORT5 = df.iloc[i][40]
    ET_SORT6 = df.iloc[i][41]
    ET_SORT7 = df.iloc[i][42]
    ET_SORT8 = df.iloc[i][43]
    CD_A = df.iloc[i][44]
    CD_G = df.iloc[i][45]
    CD_H = df.iloc[i][46]
    CD_I = df.iloc[i][47]
    CD_M = df.iloc[i][48]
    CD_Q = df.iloc[i][49]
    CD_L = df.iloc[i][50]
    CD_S = df.iloc[i][51]
    CD_J = df.iloc[i][52]
    CD_T = df.iloc[i][53]
    CD_OTHERS = df.iloc[i][54]
    ET_FAIL = df.iloc[i][55]
    ET_NULL = df.iloc[i][56]
    ET_P = df.iloc[i][57]
    DESC = df.iloc[i][58]
#material_setup
    browser.find_element(By.XPATH,'/html/body/table/tbody/tr/td[1]/center/table/tbody/tr[7]/td').click()
    Select(browser.find_element(By.XPATH,'//*[@id="product"]')).select_by_visible_text(PRODUCT)
    Select(browser.find_element(By.XPATH,'//*[@id="tab"]')).select_by_visible_text(TAB)
    Select(browser.find_element(By.XPATH,'//*[@id="sdet_wo_type"]')).select_by_visible_text("Normal (PMR)")
    browser.find_element(By.XPATH,'//*[@id="submit1"]').click()
    WebDriverWait(browser, 20).until(EC.visibility_of_element_located((By.XPATH,'//*[@id="contents"]')))
    browser.find_element(By.XPATH,'//*[@id="woId"]').send_keys(WO_NAME)
    Select(browser.find_element(By.XPATH,'//*[@id="woPrefix"]')).select_by_visible_text("SDET NPD Prime (SNP)")
    browser.find_element(By.XPATH,'//*[@id="waferCode"]').send_keys(WF_3D)
    if browser.find_elements(By.XPATH,'//select[@id="trayType"]/option[@value ="'+TRAY+'"]') :
     Select(browser.find_element(By.XPATH,'//*[@id="trayType"]')).select_by_visible_text(TRAY)
    else :
     df['Status'][i] = 'Tray Type not found'
     df.to_csv('wo_create_log.csv',index=False)
     continue
    browser.find_element(By.XPATH,'//*[@id="description"]').send_keys(DESC)
    browser.find_element(By.XPATH,'//*[@id="tmwi"]').send_keys(TMWI)
    browser.find_element(By.XPATH,'//*[@id="tmwiR"]').send_keys(TMWI)
    browser.find_element(By.XPATH,'//*[@id="marking"]').send_keys(SDET_MARKING)
    WebDriverWait(browser,80).until(EC.visibility_of_element_located((By.XPATH,'//select[@id="sldrPartNum"]/option[@value ="'+SLD_PN+'"]')))
    if browser.find_elements(By.XPATH,'//select[@id="sldrPartNum"]/option[@value ="'+SLD_PN+'"]'):
     Select(browser.find_element(By.XPATH,'//*[@id="sldrPartNum"]')).select_by_visible_text(SLD_PN)
    else :
     df['Status'][i] = 'SLD_PN not found'
     df.to_csv('wo_create_log.csv',index=False)
     continue
    WebDriverWait(browser,80).until(EC.visibility_of_element_located((By.XPATH,'//select[@id="cb_aab_designs"]/option[@value ="'+SLD_AAB+'"]')))	 
    if browser.find_elements(By.XPATH,'//select[@id="cb_aab_designs"]/option[@value ="'+SLD_AAB+'"]'):
     Select(browser.find_element(By.XPATH,'//*[@id="cb_aab_designs"]')).select_by_visible_text(SLD_AAB)
    else :
     df['Status'][i] = 'SLD_AAB not found'
     df.to_csv('wo_create_log.csv',index=False)
     continue
    WebDriverWait(browser,80).until(EC.visibility_of_element_located((By.XPATH,'//select[@id="cb_sets_part_nums"]/option[@value ="'+SET_PN+'"]')))
    if browser.find_elements(By.XPATH,'//select[@id="cb_sets_part_nums"]/option[@value ="'+SET_PN+'"]'):
     Select(browser.find_element(By.XPATH,'//*[@id="cb_sets_part_nums"]')).select_by_visible_text(SET_PN)
    else :
     df['Status'][i] = 'SET_PN not found'
     df.to_csv('wo_create_log.csv',index=False)
     continue
    WebDriverWait(browser,80).until(EC.visibility_of_element_located((By.XPATH,'//select[@id="cb_sets_vers"]/option[@value ="'+SET_VER+'"]')))
    if browser.find_elements(By.XPATH,'//select[@id="cb_sets_vers"]/option[@value ="'+SET_VER+'"]'):
     Select(browser.find_element(By.XPATH,'//*[@id="cb_sets_vers"]')).select_by_visible_text(SET_VER)
    else :
     df['Status'][i] = 'SET_VER not found'
     df.to_csv('wo_create_log.csv',index=False)
     continue
    WebDriverWait(browser,40).until(EC.visibility_of_element_located((By.XPATH,'//*[@id="bt_add_sets"]')))
    browser.find_element(By.XPATH,'//*[@id="bt_add_sets"]').click()
#TSR_setup
    if (type(STEP1_PC) == str) :
     Select(browser.find_element(By.XPATH,'//*[@id="step01"]')).select_by_visible_text(get_process_code(STEP1_PC))
     browser.find_element(By.XPATH,'//*[@id="prog01"]').send_keys(STEP1_TSR)
     if (STEP1_PC == 'G'):
      browser.find_element(By.XPATH,'//*[@id="ySamp01"]').click()
      browser.find_element(By.XPATH,'//*[@id="rOpn01" and @value="HMB"]').click()
    if (type(STEP2_PC) == str) :
     Select(browser.find_element(By.XPATH,'//*[@id="step02"]')).select_by_visible_text(get_process_code(STEP2_PC))
     browser.find_element(By.XPATH,'//*[@id="prog02"]').send_keys(STEP2_TSR)
     if (STEP2_PC == 'G'):
      browser.find_element(By.XPATH,'//*[@id="ySamp02"]').click()
      browser.find_element(By.XPATH,'//*[@id="rOpn02" and @value="HMB"]').click()
    if (type(STEP3_PC) == str) :
     Select(browser.find_element(By.XPATH,'//*[@id="step03"]')).select_by_visible_text(get_process_code(STEP3_PC))
     browser.find_element(By.XPATH,'//*[@id="prog03"]').send_keys(STEP3_TSR)
     if (STEP3_PC == 'G'):
      browser.find_element(By.XPATH,'//*[@id="ySamp03"]').click()
      browser.find_element(By.XPATH,'//*[@id="rOpn03 and @value="HMB""]').click()
    if (type(STEP4_PC) == str) :
     Select(browser.find_element(By.XPATH,'//*[@id="step04"]')).select_by_visible_text(get_process_code(STEP4_PC))
     browser.find_element(By.XPATH,'//*[@id="prog04"]').send_keys(STEP4_TSR)
     if (STEP4_PC == 'G'):
      browser.find_element(By.XPATH,'//*[@id="ySamp04"]').click()
      browser.find_element(By.XPATH,'//*[@id="rOpn04 and @value="HMB""]').click()
    if (type(STEP5_PC) == str) :
     Select(browser.find_element(By.XPATH,'//*[@id="step05"]')).select_by_visible_text(get_process_code(STEP5_PC))
     browser.find_element(By.XPATH,'//*[@id="prog05"]').send_keys(STEP5_TSR)
     if (STEP5_PC == 'G'):
      browser.find_element(By.XPATH,'//*[@id="ySamp05"]').click()
      browser.find_element(By.XPATH,'//*[@id="rOpn05 and @value="HMB""]').click()
    if (type(STEP6_PC) == str) :
     Select(browser.find_element(By.XPATH,'//*[@id="step06"]')).select_by_visible_text(get_process_code(STEP6_PC))
     browser.find_element(By.XPATH,'//*[@id="prog06"]').send_keys(STEP6_TSR)
     if (STEP6_PC == 'G'):
      browser.find_element(By.XPATH,'//*[@id="ySamp06"]').click()
      browser.find_element(By.XPATH,'//*[@id="rOpn06 and @value="HMB""]').click()
    if (type(STEP7_PC) == str) :
     Select(browser.find_element(By.XPATH,'//*[@id="step07"]')).select_by_visible_text(get_process_code(STEP7_PC))
     browser.find_element(By.XPATH,'//*[@id="prog07"]').send_keys(STEP7_TSR)
     if (STEP7_PC == 'G'):
      browser.find_element(By.XPATH,'//*[@id="ySamp07"]').click()
      browser.find_element(By.XPATH,'//*[@id="rOpn07 and @value="HMB""]').click()
    if (type(STEP8_PC) == str) :
     Select(browser.find_element(By.XPATH,'//*[@id="step08"]')).select_by_visible_text(get_process_code(STEP8_PC))
     browser.find_element(By.XPATH,'//*[@id="prog08"]').send_keys(STEP8_TSR)
     if (STEP8_PC == 'G'):
      browser.find_element(By.XPATH,'//*[@id="ySamp08"]').click()
      browser.find_element(By.XPATH,'//*[@id="rOpn08 and @value="HMB""]').click()
#bin_setup
    if (type(ET_SORT0) == str and ET_SORT0.startswith("Bin")) :
     browser.find_element(By.XPATH,'//*[@id="gSort0"]').click()
     Select(browser.find_element(By.XPATH,'//*[@id="et0"]')).select_by_visible_text(ET_SORT0)
    if (type(ET_SORT1) == str and ET_SORT1.startswith("Bin")) :
     browser.find_element(By.XPATH,'//*[@id="gSort1"]').click()
     Select(browser.find_element(By.XPATH,'//*[@id="et1"]')).select_by_visible_text(ET_SORT1)
    if (type(ET_SORT2) == str and ET_SORT2.startswith("Bin")) :
     browser.find_element(By.XPATH,'//*[@id="gSort2"]').click()
     Select(browser.find_element(By.XPATH,'//*[@id="et2"]')).select_by_visible_text(ET_SORT2)
    if (type(ET_SORT3) == str and ET_SORT3.startswith("Bin")) :
     browser.find_element(By.XPATH,'//*[@id="gSort3"]').click()
     Select(browser.find_element(By.XPATH,'//*[@id="et3"]')).select_by_visible_text(ET_SORT3)
    if (type(ET_SORT4) == str and ET_SORT4.startswith("Bin")) :
     browser.find_element(By.XPATH,'//*[@id="gSort4"]').click()
     Select(browser.find_element(By.XPATH,'//*[@id="et4"]')).select_by_visible_text(ET_SORT4)
    if (type(ET_SORT5) == str and ET_SORT5.startswith("Bin")) :
     browser.find_element(By.XPATH,'//*[@id="gSort5"]').click()
     Select(browser.find_element(By.XPATH,'//*[@id="et5"]')).select_by_visible_text(ET_SORT5)
    if (type(ET_SORT6) == str and ET_SORT6.startswith("Bin")) :
     browser.find_element(By.XPATH,'//*[@id="gSort6"]').click()
     Select(browser.find_element(By.XPATH,'//*[@id="et6"]')).select_by_visible_text(ET_SORT6)
    if (type(ET_SORT7) == str and ET_SORT7.startswith("Bin")) :
     browser.find_element(By.XPATH,'//*[@id="gSort7"]').click()
     Select(browser.find_element(By.XPATH,'//*[@id="et7"]')).select_by_visible_text(ET_SORT7)
    if (type(ET_SORT8) == str and ET_SORT8.startswith("Bin")) :
     browser.find_element(By.XPATH,'//*[@id="gSort8"]').click()
     Select(browser.find_element(By.XPATH,'//*[@id="et8"]')).select_by_visible_text(ET_SORT8)
    if (type(CD_A) == str and CD_A.startswith("Bin")) :
     Select(browser.find_element(By.XPATH,'//*[@id="etA"]')).select_by_visible_text(CD_A)
    if (type(CD_G) == str and CD_G.startswith("Bin")) :
     Select(browser.find_element(By.XPATH,'//*[@id="etG"]')).select_by_visible_text(CD_G)
    if (type(CD_H) == str and CD_H.startswith("Bin")) :
     Select(browser.find_element(By.XPATH,'//*[@id="etH"]')).select_by_visible_text(CD_H)
    if (type(CD_I) == str and CD_I.startswith("Bin")) :
     Select(browser.find_element(By.XPATH,'//*[@id="etI"]')).select_by_visible_text(CD_I)
    if (type(CD_M) == str and CD_M.startswith("Bin")) :
     Select(browser.find_element(By.XPATH,'//*[@id="etM"]')).select_by_visible_text(CD_M)
    if (type(CD_Q) == str and CD_Q.startswith("Bin")) :
     Select(browser.find_element(By.XPATH,'//*[@id="etQ"]')).select_by_visible_text(CD_Q)
    if (type(CD_L) == str and CD_L.startswith("Bin")) :
     Select(browser.find_element(By.XPATH,'//*[@id="etL"]')).select_by_visible_text(CD_L)
    if (type(CD_S) == str and CD_S.startswith("Bin")) :
     Select(browser.find_element(By.XPATH,'//*[@id="etS"]')).select_by_visible_text(CD_S)
    if (type(CD_J) == str and CD_J.startswith("Bin")) :
     Select(browser.find_element(By.XPATH,'//*[@id="etJ"]')).select_by_visible_text(CD_J)
    if (type(CD_T) == str and CD_T.startswith("Bin")) :
     Select(browser.find_element(By.XPATH,'//*[@id="etT"]')).select_by_visible_text(CD_T)
    if (type(ET_FAIL) == str and ET_FAIL.startswith("Bin")) :
     Select(browser.find_element(By.XPATH,'//*[@id="etF"]')).select_by_visible_text(ET_FAIL)
    if (type(ET_NULL) == str and ET_NULL.startswith("Bin")) :
     Select(browser.find_element(By.XPATH,'//*[@id="etN"]')).select_by_visible_text(ET_NULL)
    if (type(ET_P) == str and ET_P.startswith("Bin")) :
     Select(browser.find_element(By.XPATH,'//*[@id="etP"]')).select_by_visible_text(ET_P)
    if (type(MARKING_BIN0) == str) :
     browser.find_element(By.XPATH,'//*[@id="new_marking_0"]').send_keys(MARKING_BIN0)
    if (type(MARKING_BIN1) == str) :
     browser.find_element(By.XPATH,'//*[@id="new_marking_1"]').send_keys(MARKING_BIN1)
    if (type(MARKING_BIN2) == str) :
     browser.find_element(By.XPATH,'//*[@id="new_marking_2"]').send_keys(MARKING_BIN2)
    if (type(MARKING_BIN3) == str) :
     browser.find_element(By.XPATH,'//*[@id="new_marking_3"]').send_keys(MARKING_BIN3)
    if (type(MARKING_BIN4) == str) :
     browser.find_element(By.XPATH,'//*[@id="new_marking_4"]').send_keys(MARKING_BIN4)
    if (type(MARKING_BIN5) == str) :
     browser.find_element(By.XPATH,'//*[@id="new_marking_5"]').send_keys(MARKING_BIN5)
    if (type(MARKING_BIN6) == str) :
     browser.find_element(By.XPATH,'//*[@id="new_marking_6"]').send_keys(MARKING_BIN6)
    if (type(MARKING_BIN7) == str) :
     browser.find_element(By.XPATH,'//*[@id="new_marking_7"]').send_keys(MARKING_BIN7)
    if (cb1.isChecked()):
     browser.find_element(By.XPATH,'//*[@id="save"]').click()
     df['Status'][i] = 'WO_Created'
     df.to_csv('wo_create_log.csv',index=False)
    else:
     df['Status'][i] = 'WO_Simulate_Created'
     df.to_csv('wo_create_log.csv',index=False)
   noti_gid.setStyleSheet(css4)
   noti_gid.setText("   Done.. อย่าลืมเช็ค WOก่อนใช้งานจริงนะ!")
   browser.quit()
   
def inform_wo():
 if (cb1.isChecked()):
  noti_gid.setStyleSheet(css4)
  noti_gid.setText("   Enable WO creation")
 else:
  noti_gid.setStyleSheet(css5)
  noti_gid.setText("   Disable WO creation")
 
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

qAp.setWindowIcon(QIcon('cat.png'))
window.setWindowTitle('SDET WO Creator V0.4')
window.resize(700,330)

info1 = QLabel("""SDET_WO_CREATOR v.0.4 (Create New / Normal (PMR)) created by Kasira C.""",window)
info1.setGeometry(20,0,660,30)
info2 = QLabel("""Please provide Seagate account and password, and select excel file which including data followed by template.""",window)
info2.setGeometry(20,30,660,30)
info3 = QLabel("""**This program is under development, no cross-check function. Use as your own risk**""",window)
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
cat = resource_path("cat.png")

info1.setStyleSheet(css1)
info2.setStyleSheet(css2)
info3.setStyleSheet(css3)

noti_gid = QLabel(' ',window) 
noti_gid.setGeometry(200,230,300,60)

noti_file = QLabel(' ',window)
noti_file.setGeometry(180,140,500,30)

gid = QLineEdit(window)
gid.setPlaceholderText('Please enter your GID')
gid.setGeometry(20,100,140,30)

pw = QLineEdit(window)
pw.setPlaceholderText('Please enter your password')
pw.setEchoMode(QLineEdit.Password)
pw.setGeometry(170,100,140,30)

bt1 = QPushButton(QIcon('cat.png'),'เลือก Excel file..',window)
bt1.setGeometry(20,140,150,30)
bt1.clicked.connect(call_file)

bt2 = QPushButton('Submit',window)
bt2.setGeometry(20,180,150,30)
bt2.clicked.connect(check_gid_password)

cb1 = QCheckBox("Enable WO Creation",window)
cb1.setChecked(True)
cb1.setGeometry(180,180,150,30)
cb1.stateChanged.connect(inform_wo)

cb2 = QCheckBox("Disable View Mode",window)
cb2.setChecked(True)
cb2.setGeometry(340,180,170,30)
cb2.stateChanged.connect(inform_viwe_mode)

btlog = QPushButton('Log',window)
btlog.setGeometry(665,305,30,20)
btlog.clicked.connect(update_log)

window.show()
qAp.exec_()