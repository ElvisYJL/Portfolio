#!/usr/bin/env python
# coding: utf-8

# In[1]:


from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from xlrd import xldate_as_tuple
import xlrd, csv, os, datetime, logging, traceback
from time import time
import xlwings as xw

print('\n')
print("--- Driver Release Note Tool ---")
print("--- Authorized by Elvis_Lee (A31SPT_SWD_SA2) ---")
print("--- #Elvis_Lee@compal.com #19893 ---")
print('\n')

print("--- Please key in Driver SWB: ")
Driver_SWB = input()
print("--- Please key in Driver SRV: ")
Driver_SRV = input()

ISOTIMEFORMAT = '%Y-%m-%d %H:%M:%S'
theTime = datetime.datetime.now().strftime(ISOTIMEFORMAT)
print('\n')
print('--- Driver Release Note Tool --- Start Time: ' , theTime + '\n')


#Open Excel
wb = xw.Book('Beta_Released_Note_Sample.xlsx')
    
sheet = wb.sheets['Driver_Release'] 

try:
       
    driver_path = 'chromedriver_v87.exe'
    driver = webdriver.Chrome(executable_path = driver_path)
    driver.get('Dell Issue Website')
    driver.implicitly_wait(30)

    f = open('Account_Password.txt')
    Account_Password = []
    for Account_Password_list in f:
        Account_Password.append(Account_Password_list.replace("\n", ""))
    print('\n')
    print('--- AIC Account: ' , Account_Password[0])
    print('--- AIC Password: ' , Account_Password[1] + '\n')

    #locate account
    elem_user = driver.find_element_by_id('j_username') 
    #clear account
    elem_user.clear() 
    #enter user account 
    elem_user.send_keys(Account_Password[0]) 
    #loacte password
    elem_pass = driver.find_element_by_id('j_password') 
    #clear password
    elem_pass.clear() 
    #enter user password
    elem_pass.send_keys(Account_Password[1]) 
    #locate login 
    elem_login = driver.find_element_by_id('loginspan').click()

    driver.switch_to_window(driver.window_handles[1])

    driver.implicitly_wait(30)

    elem_search = driver.find_element_by_xpath("//input[@id='QUICKSEARCH_STRING']")
    
    print('\n')
    print("--- Driver SWB: ", Driver_SWB, " ---")

    elem_search.send_keys(Driver_SWB)
    
    #//span[@id='top_simpleSearchspan']
    elem_search_button = driver.find_element_by_xpath("//*[@id='top_simpleSearchspan']").click()

    driver.implicitly_wait(30)

    try:
        elem_SWB_list = driver.find_element_by_link_text(Driver_SWB).click()
    except:
        print('--- No SWB list')

    driver.implicitly_wait(30)
    
    Number = driver.find_element_by_xpath("//*[@id='col_1001']").text
    if Number == ' ':
        Number = 'N/A'
    Subsystem = driver.find_element_by_xpath("//*[@id='col_1548']").text
    if Subsystem == ' ':
        Subsystem = 'N/A'
    Dell_Version = driver.find_element_by_xpath("//*[@id='col_1576']").text
    if Dell_Version == ' ':
        Dell_Version = 'N/A'
    Predecessor = driver.find_element_by_xpath("//*[@id='col_283535']").text
    if Predecessor == ' ':
        Predecessor = 'N/A'
    Software_Bundle_Type = driver.find_element_by_xpath("//*[@id='col_1556']").text
    if Software_Bundle_Type == ' ':
        Software_Bundle_Type = 'N/A'
    Language = driver.find_element_by_xpath("//*[@id='col_1565']").text
    if Language == ' ':
        Language = 'N/A'
    Description = driver.find_element_by_xpath("//*[@id='col_1002']").text
    if Description == ' ':
        Description = 'N/A'
    Title_External = driver.find_element_by_xpath("//*[@id='col_2000003306']").text
    if Title_External == ' ':
        Title_External = 'N/A'
    Operating_System_OS = driver.find_element_by_xpath("//*[@id='col_1566']").text
    if Operating_System_OS == ' ':
        Operating_System_OS = 'N/A'
    Factory_Install_MDiag_Instructions = driver.find_element_by_xpath("//*[@id='col_2000003301']").text
    if Factory_Install_MDiag_Instructions == ' ':
        Factory_Install_MDiag_Instructions = 'N/A'
    Vendor_Version = driver.find_element_by_xpath("//*[@id='col_1582']").text
    if Vendor_Version == ' ':
        Vendor_Version = 'N/A'
    Description_External = driver.find_element_by_xpath("//*[@id='col_2000003300']").text
    if Description_External == ' ':
        Description_External = 'N/A'
    Fixes_and_Enhancements_External = driver.find_element_by_xpath("//*[@id='col_2000003302']").text
    if Fixes_and_Enhancements_External == ' ':
        Fixes_and_Enhancements_External = 'N/A'
    Image_Watch_Classification = driver.find_element_by_xpath("//*[@id='col_1563']").text
    if Image_Watch_Classification == ' ':
        Image_Watch_Classification = 'N/A'
    Master_Device_Driver_Version = driver.find_element_by_xpath("//*[@id='col_2000003297']").text
    if Master_Device_Driver_Version == ' ':
        Master_Device_Driver_Version = 'N/A'
    Proj_Platform_Affected = driver.find_element_by_xpath("//*[@id='col_1004']").text
    if Proj_Platform_Affected == ' ':
        Proj_Platform_Affected = 'N/A'
    Software_Version = driver.find_element_by_xpath("//*[@id='col_1585']").text
    if Software_Version == ' ':
        Software_Version = 'N/A'
    
    ISOTIMEFORMAT = '%Y-%m-%d'
    theTime_today = datetime.datetime.now().strftime(ISOTIMEFORMAT) 

    print('- SWB, Number: ', Number)
    print('- SWB, Subsystem: ', Subsystem)
    print('- SWB, Dell_Version: ', Dell_Version)
    print('- SWB, Predecessor: ', Predecessor)
    print('- SWB, Software_Bundle_Type: ', Software_Bundle_Type)
    print('- SWB, Language: ', Language)
    print('- SWB, Description: ', Description)
    print('- SWB, Title_External: ', Title_External)
    print('- SWB, Operating_System_OS: ', Operating_System_OS)
    print('- SWB, Factory_Install_MDiag_Instructions: ', Factory_Install_MDiag_Instructions)
    print('- SWB, Vendor_Version: ', Vendor_Version)
    print('- SWB, Description_External: ', Description_External)
    print('- SWB, Fixes_and_Enhancements_External: ', Fixes_and_Enhancements_External)
    print('- SWB, Image_Watch_Classification: ', Image_Watch_Classification)
    print('- SWB, Master_Device_Driver_Version: ', Master_Device_Driver_Version)
    print('- SWB, Proj_Platform_Affected: ', Proj_Platform_Affected)
    print('- SWB, Software_Version: ', Software_Version)
    print('- SWB, theTime_today: ', theTime_today)
    
    sheet['C10'].value = Driver_SWB
    sheet['D42'].value = Number
    sheet['F11'].value = Subsystem
    sheet['F09'].value = Dell_Version
    sheet['D48'].value = Dell_Version
    sheet['F10'].value = Predecessor
    sheet['C12'].value = Software_Bundle_Type
    sheet['B38'].value = Language
    sheet['C15'].value = Description
    sheet['D43'].value = Description
    sheet['D55'].value = Description
    sheet['D45'].value = Title_External
    sheet['D59'].value = Title_External
    sheet['D46'].value = Operating_System_OS
    sheet['D47'].value = Factory_Install_MDiag_Instructions
    sheet['D49'].value = Vendor_Version
    sheet['C9'].value = Vendor_Version
    sheet['D50'].value = Description_External
    sheet['D51'].value = Fixes_and_Enhancements_External
    sheet['D52'].value = Image_Watch_Classification
    sheet['D53'].value = Master_Device_Driver_Version
    sheet['D44'].value = Proj_Platform_Affected
    sheet['C23'].value = Software_Version
    sheet['F17'].value = theTime_today
    '''
    #Attachment
    print('Switch to "Attachments" ... ')
    elem_search_button = driver.find_element_by_link_text("Attachments").click()
    
    Driver_Package_Name = driver.find_element_by_partial_link_text("MT7XX")
    
    sheet['C8'].value = Driver_Package_Name.text
    print('Finish search "Attachments"')
    '''
    
    #Changes-PNCR
    print('\n')
    print('--- Switch to "Changes" ... ')
    elem_search_button = driver.find_element_by_link_text("Changes").click()
    
    Changes_Number = driver.find_element_by_xpath("//*[@id='ITEMTABLE_REVLIST']/tbody/tr[2]/td[2]/div/div[1]/table/tbody/tr[2]/td[4]/a").text
    if Changes_Number == ' ':
        Changes_Number = 'N/A'
    
    print('- SWB, Changes_Number: ', Changes_Number)
    sheet['C11'].value = Changes_Number
    print('--- Finish search "Changes"')
    

    elem_search.clear()

    #Driver_SRV
    elem_search = driver.find_element_by_xpath("//input[@id='QUICKSEARCH_STRING']")

    print('\n')
    print("--- DRIVER SRV: ", Driver_SRV, " ---")

    elem_search.send_keys(Driver_SRV)
    
    #//span[@id='top_simpleSearchspan']
    elem_search_button = driver.find_element_by_xpath("//*[@id='top_simpleSearchspan']").click()

    driver.implicitly_wait(30)

    try:
        elem_SRV_list = driver.find_element_by_link_text(Driver_SRV).click()
    except:
        print('--- No SRV list')

    driver.implicitly_wait(30)

    SRV_Description = driver.find_element_by_xpath("//*[@id='col_1002']").text
    if SRV_Description == ' ':
        SRV_Description = 'N/A'
    SRV_Proj_Platform_Affected = driver.find_element_by_xpath("//*[@id='col_1004']").text
    if SRV_Proj_Platform_Affected == ' ':
        SRV_Proj_Platform_Affected = 'N/A'
    SRV_PCI_ID = driver.find_element_by_xpath("//*[@id='col_1567']").text
    if SRV_PCI_ID == ' ':
        SRV_PCI_ID = 'N/A'
    SRV_PnP_ID = driver.find_element_by_xpath("//*[@id='col_1568']").text
    if SRV_PnP_ID == ' ':
        SRV_PnP_ID = 'N/A'
    SRV_Component_ID = driver.find_element_by_xpath("//*[@id='col_1577']").text
    if SRV_Component_ID == ' ':
        SRV_Component_ID = 'N/A'
    SRV_LOBs_Affected = driver.find_element_by_xpath("//*[@id='col_2091']").text
    if SRV_LOBs_Affected == ' ':
        SRV_LOBs_Affected = 'N/A'

    print('- SRV, SRV_Description: ', SRV_Description)
    print('- SRV, SRV_Proj_Platform_Affected: ', SRV_Proj_Platform_Affected)
    print('- SRV, SRV_PCI_ID: ', SRV_PCI_ID)
    print('- SRV, SRV_PnP_ID: ', SRV_PnP_ID)
    print('- SRV, SRV_Component_ID: ', SRV_Component_ID)
    print('- SRV, SRV_LOBs_Affected: ', SRV_LOBs_Affected)
    
    sheet['D54'].value = Driver_SRV
    sheet['D55'].value = SRV_Description
    sheet['D56'].value = SRV_Proj_Platform_Affected
    sheet['C5'].value = SRV_Proj_Platform_Affected
    sheet['D57'].value = SRV_PCI_ID
    sheet['D58'].value = SRV_PnP_ID
    sheet['D60'].value = SRV_Component_ID
    sheet['B5'].value = SRV_LOBs_Affected
    
    print('\n')
    print("--- Search Information Finish ---")   

except Exception as e: 
    mode = 'w'
    FORMAT = '%(asctime)s %(levelname)s: %(message)s'
    logging.basicConfig(filename=r'RUSB_LOG.txt', filemode= mode, format=FORMAT)
    print("!!! Something get wrong, please refer to RUSB_LOG.txt !!!")
    logging.error("Main program error: ")
    logging.error(e)
    logging.error(traceback.format_exc())
driver.quit()
print('\n')
print("--- Close Browser ---")    

wb.save('Beta_Released_Note_Finish.xlsx')

print('\n')
print("--- Driver Release Note Excel Finish Write ---")

ISOTIMEFORMAT = '%Y-%m-%d %H:%M:%S'
theTime_finish = datetime.datetime.now().strftime(ISOTIMEFORMAT)
print('\n')
print('--- Driver Release Note Tool --- Finish Time: ', theTime_finish)

input()
