#!/usr/bin/env python
# coding: utf-8

# In[22]:


from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import xlwings as xw
from time import time
from datetime import datetime

print('\n')
print("--- Jira Information Tool ---")
print("--- Authorized by Elvis_Lee (A31SPT_SWD_SA2) ---")
print("--- #Elvis_Lee@compal.com #19893 ---")
print('\n')

print('Press any key to run the tool...')
input()

ISOTIMEFORMAT = '%Y-%m-%d %H:%M:%S'
theTime = datetime.now().strftime(ISOTIMEFORMAT)
print('\n')
print('--- Start Time: ' , theTime + '\n')

driver = webdriver.Edge(executable_path = 'msedgedriver.exe')
driver.get('https://jira.cpg.dell.com/secure/Dashboard.jspa')
driver.implicitly_wait(30)

print('Searching the Account and Password...')
f = open('ACPW.txt')
Account_Password = []
for Account_Password_list in f:
    Account_Password.append(Account_Password_list.replace("\n", ""))
print('\n')
print('--- Jira Account: ' , Account_Password[0])
print('--- Jira Password: ' , Account_Password[1] + '\n')

print('Login Jira website...')
#locate account
user_acc = driver.find_element_by_xpath("//*[@id='login-form-username']")
#enter user account 
user_acc.clear()
user_acc.send_keys(Account_Password[0]) 
#loacte password
user_pw = driver.find_element_by_xpath("//*[@id='login-form-password']")
#enter user password
user_pw.clear()
user_pw.send_keys(Account_Password[1]) 
#Click login button
login = driver.find_element_by_xpath("//*[@id='login-form-submit']").click() 
driver.implicitly_wait(30)
print('\n')

print('Searching PIMS nimber...')
f = open('PIMS.txt')
PIMS_array = []
for PIMS_array_list in f:
    PIMS_array.append(PIMS_array_list.replace("\n", ""))
print('\n')
print("--- Search from the following PIMS list ---")
print(PIMS_array)
print('\n')

Title_array = []
Details_Status_array = []
Details_Resolution_array = []
Details_Security_Level_array = []
Source_Originating_Vendor_array = []
Source_Originating_Group_array = []
Source_LOB_Found_array = []
Source_Platform_Found_array = []
Source_Component_Subcomponent_array = []
Source_Group_Activity_array = []
Source_Group_Location_array = []
Source_Phase_Found_array = []
Source_Hardware_Build_Version_array = []
Source_Discovery_Method_array = []
Source_Test_Case_Number_array = []
Severity_Issue_Severity_array = []
Severity_Product_Impact_array = []
Severity_Customer_Impact_array = []
Severity_Likelihood_array = []
Severity_RPN_array = []
Affected_Affected_OS_array = []
Affected_Affected_Languages_array = []
Affected_Affected_Items_array = []
Description_Steps_to_Reproduce_array = []
Disposition_Disposition_Type_array = []
Disposition_Disposition_Details_array = []
Disposition_Technical_Root_Cause_array = []
Vendor_Vendor_Accesses_array = []
Misc_Discretionary_Field_1_array = []
Misc_Discretionary_Field_2_array = []
Misc_Failed_Verification_Fix_Fail_Count_array = []
Misc_Invalid_Disposition_Fix_Fail_Count_array = []
Desciription_content_array = []
Attachments_array = []
Sub_Tasks_array = []
Activity_comments_array = []
Submitted_Date_array = []
Time_lag_array = []

print('Begin to search the information on Jira...' + '\n')

for i in range(0,len(PIMS_array)):
    print('Searching... --- ', PIMS_array[i], ' ---' + '\n')
    PIMS_number = driver.find_element_by_xpath("//*[@id='quickSearchInput']") 
    PIMS_number.clear() 
    PIMS_number.send_keys(PIMS_array[i])
    PIMS_number.send_keys(Keys.ENTER)
    
    Title = driver.find_element_by_xpath("//*[@id='summary-val']").text
    Details_Status = driver.find_element_by_xpath("//*[@id='status-val']").text
    Details_Resolution = driver.find_element_by_xpath("//*[@id='resolution-val']").text
    Details_Security_Level = driver.find_element_by_xpath("//*[@id='security-val']").text
    print('Title: ', Title)
    print('Details_Status: ', Details_Status)
    print('Details_Resolution: ', Details_Resolution)
    print('Details_Security_Level: ', Details_Security_Level)
    print('\n')

    #Source_Originating_Vendor = driver.find_element_by_xpath("//*[@id='customfield_28415-val']").text
    #Source_Originating_Group = driver.find_element_by_xpath("//*[@id='customfield_28425-val']").text
    #Source_LOB_Found = driver.find_element_by_xpath("//*[@id='customfield_28469-field']").text
    Source_Platform_Found = driver.find_element_by_xpath("//*[@id='customfield_28470-val']").text
    Source_Component_Subcomponent = driver.find_element_by_xpath("//*[@id='customfield_28436-val']").text
    #Source_Group_Activity = driver.find_element_by_xpath("//*[@id='customfield_28476-val']").text
    #Source_Group_Location = driver.find_element_by_xpath("//*[@id='customfield_28426-val']").text
    #Source_Phase_Found = driver.find_element_by_xpath("//*[@id='customfield_18732-val']").text
    #Source_Hardware_Build_Version = driver.find_element_by_xpath("//*[@id='customfield_28454-val']").text
    #try:
        #Source_Discovery_Method = driver.find_element_by_xpath("//*[@id='customfield_18712-val']").text
    #except:
        #Source_Discovery_Method = "N/A"
    Source_Test_Case_Number = driver.find_element_by_xpath("//*[@id='customfield_18728-val']").text
    #print('Source_Originating_Vendor: ', Source_Originating_Vendor)
    #print('Source_Originating_Group: ', Source_Originating_Group)
    #print('Source_LOB_Found: ', Source_LOB_Found)
    print('Source_Platform_Found: ', Source_Platform_Found)
    print('Source_Component_Subcomponent: ', Source_Component_Subcomponent)
    #print('Source_Group_Activity: ', Source_Group_Activity)
    #print('Source_Group_Location: ', Source_Group_Location)
    #print('Source_Phase_Found: ', Source_Phase_Found)
    #print('Source_Hardware_Build_Version: ', Source_Hardware_Build_Version)
    #print('Source_Discovery_Method: ', Source_Discovery_Method)
    print('Source_Test_Case_Number: ', Source_Test_Case_Number)
    print('\n')

    #Detail_Sverity_tag = driver.find_element_by_xpath("//*[@id='aui-uid-3']").click()
    Detail_Sverity_tag = driver.find_element_by_link_text("Severity").click()
    Severity_Issue_Severity = driver.find_element_by_xpath("//*[@id='customfield_28162-val']").text
    #Severity_Product_Impact = driver.find_element_by_xpath("//*[@id='customfield_15206-val']").text
    #Severity_Customer_Impact = driver.find_element_by_xpath("//*[@id='customfield_15208-val']").text
    #Severity_Likelihood = driver.find_element_by_xpath("//*[@id='customfield_15205-val']").text
    #Severity_RPN = driver.find_element_by_xpath("//*[@id='customfield_28439-val']").text
    print('Severity_Issue_Severity: ', Severity_Issue_Severity)
    #print('Severity_Product_Impact: ', Severity_Product_Impact)
    #print('Severity_Customer_Impact: ', Severity_Customer_Impact)
    #print('Severity_Likelihood: ', Severity_Likelihood)
    #print('Severity_RPN: ', Severity_RPN)
    print('\n')
    '''
    #Detail_Affected_tag = driver.find_element_by_xpath("//*[@id='aui-uid-4']").click()
    Detail_Affected_tag = driver.find_element_by_link_text("Affected").click()
    try:
        Affected_Affected_OS = driver.find_element_by_xpath("//*[@id='customfield_28432-val']").text
    except:
        Affected_Affected_OS = "N/A"
    try:
        Affected_Affected_Languages = driver.find_element_by_xpath("//*[@id='customfield_28452-val']").text
    except:
        Affected_Affected_Languages = "N/A"
    Affected_Affected_Items = driver.find_element_by_xpath("//*[@id='customfield_28463-val']").text
    print('Affected_Affected_OS: ', Affected_Affected_OS)
    print('Affected_Affected_Languages: ', Affected_Affected_Languages)
    print('Affected_Affected_Items: ', Affected_Affected_Items)
    print('\n')
    '''
    #Detail_Description_tag = driver.find_element_by_xpath("//*[@id='aui-uid-5']").click()
    Detail_Description_tag = driver.find_element_by_link_text("Description").click()
    Description_Steps_to_Reproduce = driver.find_element_by_xpath("//*[@id='customfield_12707-val']").text
    print('Description_Steps_to_Reproduce: ', Description_Steps_to_Reproduce)
    print('\n')
    
    '''
    tag_6 = driver.find_element_by_xpath("//*[@id='aui-uid-6']").text
    if tag_6 == "Disposition":

        Detail_Disposition_tag = driver.find_element_by_xpath("//*[@id='aui-uid-6']").click()
        Disposition_Disposition_Type = driver.find_element_by_xpath("//*[@id='customfield_18716-val']").text
        Disposition_Disposition_Details = driver.find_element_by_xpath("//*[@id='customfield_18715-val']").text
        Disposition_Technical_Root_Cause = driver.find_element_by_xpath("//*[@id='customfield_26700-val']").text
        print('Disposition_Disposition_Type: ', Disposition_Disposition_Type)
        print('Disposition_Disposition_Details: ', Disposition_Disposition_Details)
        print('Disposition_Technical_Root_Cause: ', Disposition_Technical_Root_Cause)
        print('\n')

        Detail_Vendor_tag = driver.find_element_by_xpath("//*[@id='aui-uid-7']").click()
        Vendor_Vendor_Accesses = driver.find_element_by_xpath("//*[@id='customfield_28421-val']").text
        print('Vendor_Vendor_Accesses: ', Vendor_Vendor_Accesses)
        print('\n')

        Detail_Misc_tag = driver.find_element_by_xpath("//*[@id='aui-uid-8']").click()
        Misc_Discretionary_Field_1 = driver.find_element_by_xpath("//*[@id='customfield_18713-val']").text
        Misc_Discretionary_Field_2 = driver.find_element_by_xpath("//*[@id='customfield_28468-val']").text
        Misc_Failed_Verification_Fix_Fail_Count = driver.find_element_by_xpath("//*[@id='customfield_28479-val']").text
        Misc_Invalid_Disposition_Fix_Fail_Count = driver.find_element_by_xpath("//*[@id='customfield_28437-val']").text
        print('Misc_Discretionary_Field_1: ', Misc_Discretionary_Field_1)
        print('Misc_Discretionary_Field_2: ', Misc_Discretionary_Field_2)
        print('Misc_Failed_Verification_Fix_Fail_Count: ', Misc_Failed_Verification_Fix_Fail_Count)
        print('Misc_Invalid_Disposition_Fix_Fail_Count: ', Misc_Invalid_Disposition_Fix_Fail_Count)
        print('\n')

    if tag_6 == "Vendor":
        Detail_Vendor_tag = driver.find_element_by_xpath("//*[@id='aui-uid-6']").click()
        Vendor_Vendor_Accesses = driver.find_element_by_xpath("//*[@id='customfield_28421-val']").text
        print('Vendor_Vendor_Accesses: ', Vendor_Vendor_Accesses)
        print('\n')

        Detail_Misc_tag = driver.find_element_by_xpath("//*[@id='aui-uid-7']").click()
        try:
            Misc_Discretionary_Field_1 = driver.find_element_by_xpath("//*[@id='customfield_18713-val']").text
        except:
            Misc_Discretionary_Field_1 = "N/A"
        try:
            Misc_Discretionary_Field_2 = driver.find_element_by_xpath("//*[@id='customfield_28468-val']").text
        except:
            Misc_Discretionary_Field_2 = "N/A"
        Misc_Failed_Verification_Fix_Fail_Count = driver.find_element_by_xpath("//*[@id='customfield_28479-val']").text
        Misc_Invalid_Disposition_Fix_Fail_Count = driver.find_element_by_xpath("//*[@id='customfield_28437-val']").text
        print('Misc_Discretionary_Field_1: ', Misc_Discretionary_Field_1)
        print('Misc_Discretionary_Field_2: ', Misc_Discretionary_Field_2)
        print('Misc_Failed_Verification_Fix_Fail_Count: ', Misc_Failed_Verification_Fix_Fail_Count)
        print('Misc_Invalid_Disposition_Fix_Fail_Count: ', Misc_Invalid_Disposition_Fix_Fail_Count)
        print('\n')
    '''
    Desciription_content = driver.find_element_by_xpath("//*[@id='description-val']").text
    print('Desciription_content: ', '\n', Desciription_content)
    print('\n')

    Attachments = driver.find_element_by_xpath("//*[@id='attachment_thumbnails']").text
    if len(Attachments) == 0:
        Attachments = 'There is no attachment uploaded'
    print('Attachments: ', '\n', Attachments)
    print('\n')

    #Sub_Tasks = driver.find_element_by_xpath("//*[@id='view-subtasks']").text
    #print('Sub_Tasks: ', '\n', Sub_Tasks)
    #print('\n')

    Activity_comments = driver.find_element_by_xpath("//*[@id='activitymodule']/div[2]/div[2]").text
    print('Activity_comments: ', '\n',  Activity_comments)
    if Activity_comments == "There are no comments yet on this issue.":
        print("There are no comments yet on this issue.")
    print('\n')
    
    Submitted_Date = driver.find_element_by_xpath("//*[@id='customfield_28429-val']/span/time").text
    print('Submitted_Date: ', '\n', Submitted_Date)
    Submitted_Date_split = Submitted_Date[0:-3]
    PIMS_Submitted_Date = datetime.strptime(Submitted_Date_split, "%d/%b/%y %H:%M%S")
    print(PIMS_Submitted_Date)
    #Date Reference https://www.delftstack.com/zh-tw/howto/python/how-to-convert-string-to-datetime/
    Current_time = datetime.now()
    Time_lag = (Current_time - PIMS_Submitted_Date).days
    print("Time_lag: ", Time_lag)
    print('\n')
    
    print('Collect the information...' + '\n')
    Title_array.append(Title)
    Details_Status_array.append(Details_Status)
    Details_Resolution_array.append(Details_Resolution)
    Details_Security_Level_array.append(Details_Security_Level)
    #Source_Originating_Vendor_array.append(Source_Originating_Vendor)
    #Source_Originating_Group_array.append(Source_Originating_Group)
    #Source_LOB_Found_array.append(Source_LOB_Found)
    Source_Platform_Found_array.append(Source_Platform_Found)
    Source_Component_Subcomponent_array.append(Source_Component_Subcomponent)
    #Source_Group_Activity_array.append(Source_Group_Activity)
    #Source_Group_Location_array.append(Source_Group_Location)
    #Source_Phase_Found_array.append(Source_Phase_Found)
    #Source_Hardware_Build_Version_array.append(Source_Hardware_Build_Version)
    #Source_Discovery_Method_array.append(Source_Discovery_Method)
    Source_Test_Case_Number_array.append(Source_Test_Case_Number)
    Severity_Issue_Severity_array.append(Severity_Issue_Severity)
    #Severity_Product_Impact_array.append(Severity_Product_Impact)
    #Severity_Customer_Impact_array.append(Severity_Customer_Impact)
    #Severity_Likelihood_array.append(Severity_Likelihood)
    #Severity_RPN_array.append(Severity_RPN)
    #Affected_Affected_OS_array.append(Affected_Affected_OS)
    #Affected_Affected_Languages_array.append(Affected_Affected_Languages)
    #Affected_Affected_Items_array.append(Affected_Affected_Items)
    Description_Steps_to_Reproduce_array.append(Description_Steps_to_Reproduce)
    #Disposition_Disposition_Type_array.append(Disposition_Disposition_Type)
    #Disposition_Disposition_Details_array.append(Disposition_Disposition_Details)
    #Disposition_Technical_Root_Cause_array.append(Disposition_Technical_Root_Cause)
    #Vendor_Vendor_Accesses_array.append(Vendor_Vendor_Accesses)
    #Misc_Discretionary_Field_1_array.append(Misc_Discretionary_Field_1)
    #Misc_Discretionary_Field_2_array.append(Misc_Discretionary_Field_2)
    #Misc_Failed_Verification_Fix_Fail_Count_array.append(Misc_Failed_Verification_Fix_Fail_Count)
    #Misc_Invalid_Disposition_Fix_Fail_Count_array.append(Misc_Invalid_Disposition_Fix_Fail_Count)
    Desciription_content_array.append(Desciription_content)
    Attachments_array.append(Attachments)
    Sub_Tasks_array.append(Sub_Tasks)
    Activity_comments_array.append(Activity_comments)
    Submitted_Date_array.append(Submitted_Date)
    Time_lag_array.append(Time_lag)

print('Open the Excel...' + '\n')
wb = xw.Book('Jira_information.xlsx')
    
sheet = wb.sheets['PIMS']

sheet.range('a2:a1048576').clear_contents()
sheet.range('b2:b1048576').clear_contents()
sheet.range('c2:c1048576').clear_contents()
sheet.range('d2:d1048576').clear_contents()
sheet.range('e2:e1048576').clear_contents()
sheet.range('f2:f1048576').clear_contents()
sheet.range('g2:g1048576').clear_contents()
sheet.range('h2:h1048576').clear_contents()
sheet.range('i2:i1048576').clear_contents()
sheet.range('j2:j1048576').clear_contents()
sheet.range('k2:k1048576').clear_contents()
sheet.range('l2:l1048576').clear_contents()
sheet.range('m2:m1048576').clear_contents()
sheet.range('n2:n1048576').clear_contents()
sheet.range('o2:o1048576').clear_contents()

print('Export the information to Excel...' + '\n')
sheet.range('a2').options(transpose=True).value = PIMS_array[0:len(PIMS_array)+1]
sheet.range('b2').options(transpose=True).value = Title_array[0:len(Title_array)+1]
sheet.range('c2').options(transpose=True).value = Details_Status_array[0:len(Details_Status_array)+1]
sheet.range('d2').options(transpose=True).value = Details_Resolution_array[0:len(Details_Resolution_array)+1]
sheet.range('e2').options(transpose=True).value = Details_Security_Level_array[0:len(Details_Security_Level_array)+1]
sheet.range('f2').options(transpose=True).value = Source_Platform_Found_array[0:len(Source_Platform_Found_array)+1]
sheet.range('g2').options(transpose=True).value = Source_Component_Subcomponent_array[0:len(Source_Component_Subcomponent_array)+1]
sheet.range('h2').options(transpose=True).value = Source_Test_Case_Number_array[0:len(Source_Test_Case_Number_array)+1]
sheet.range('i2').options(transpose=True).value = Severity_Issue_Severity_array[0:len(Severity_Issue_Severity_array)+1]
sheet.range('j2').options(transpose=True).value = Description_Steps_to_Reproduce_array[0:len(Description_Steps_to_Reproduce_array)+1]
sheet.range('k2').options(transpose=True).value = Desciription_content_array[0:len(Desciription_content_array)+1]
sheet.range('l2').options(transpose=True).value = Attachments_array[0:len(Attachments_array)+1]
sheet.range('m2').options(transpose=True).value = Activity_comments_array[0:len(Activity_comments_array)+1]
sheet.range('n2').options(transpose=True).value = Submitted_Date_array[0:len(Submitted_Date_array)+1]
sheet.range('o2').options(transpose=True).value = Time_lag_array[0:len(Time_lag_array)+1]

sheet = wb.sheets['Overdue']

sheet.range('a2:a1048576').clear_contents()
sheet.range('b2:b1048576').clear_contents()
sheet.range('c2:c1048576').clear_contents()
sheet.range('d2:d1048576').clear_contents()
sheet.range('e2:e1048576').clear_contents()

Overdue_PIMS = []
Overdue_Title = []
Overdue_Platform_Found = []
Overdue_Submitted_Date = []
Overdue_Time_lag = []

for i in range(0, len(Time_lag_array)):
    if Time_lag_array[i] > 30:
        Overdue_PIMS.append(PIMS_array[i])
        Overdue_Title.append(Title_array[i])
        Overdue_Platform_Found.append(Source_Platform_Found_array[i])
        Overdue_Submitted_Date.append(Submitted_Date_array[i])
        Overdue_Time_lag.append(Time_lag_array[i])

sheet.range('a2').options(transpose=True).value = Overdue_PIMS[0:len(Overdue_PIMS)+1]
sheet.range('b2').options(transpose=True).value = Overdue_Title[0:len(Overdue_Title)+1]
sheet.range('c2').options(transpose=True).value = Overdue_Platform_Found[0:len(Overdue_Platform_Found)+1]
sheet.range('d2').options(transpose=True).value = Overdue_Submitted_Date[0:len(Overdue_Submitted_Date)+1]
sheet.range('e2').options(transpose=True).value = Overdue_Time_lag[0:len(Overdue_Time_lag)+1]
        
print('Finish export and save the Excel...')
wb.save('Jira_information.xlsx')

ISOTIMEFORMAT = '%Y-%m-%d %H:%M:%S'
theTime_finish = datetime.now().strftime(ISOTIMEFORMAT)
print('\n')
print('--- Finish Time: ', theTime_finish)
print('\n')

print('--- Success / Press any key to exit.')
input()


# In[11]:


wb = xw.Book('Jira_information.xlsx')

sheet = wb.sheets['Overdue']

sheet.range('a2:a1048576').clear_contents()
sheet.range('b2:b1048576').clear_contents()
sheet.range('c2:c1048576').clear_contents()
sheet.range('d2:d1048576').clear_contents()
sheet.range('e2:e1048576').clear_contents()

Overdue_PIMS = []
Overdue_Title = []
Overdue_Platform_Found = []
Overdue_Submitted_Date = []
Overdue_Time_lag = []

for i in range(0, len(Time_lag_array)):
    if Time_lag_array[i] >= 30:
        Overdue_PIMS.append(PIMS_array[i])
        Overdue_Title.append(Title_array[i])
        Overdue_Platform_Found.append(Source_Platform_Found_array[i])
        Overdue_Submitted_Date.append(Submitted_Date_array[i])
        Overdue_Time_lag.append(Time_lag_array[i])

sheet.range('a2').options(transpose=True).value = Overdue_PIMS[0:len(Overdue_PIMS)+1]
sheet.range('b2').options(transpose=True).value = Overdue_Title[0:len(Overdue_Title)+1]
sheet.range('c2').options(transpose=True).value = Overdue_Platform_Found[0:len(Overdue_Platform_Found)+1]
sheet.range('d2').options(transpose=True).value = Overdue_Submitted_Date[0:len(Overdue_Submitted_Date)+1]
sheet.range('e2').options(transpose=True).value = Overdue_Time_lag[0:len(Overdue_Time_lag)+1]
        
print('Finish export and save the Excel...')
wb.save('Jira_information.xlsx')


# In[10]:


print(Time_lag_array)


# In[5]:


from datetime import datetime
 
a=datetime.now()
Submitted_Date="11/May/21 8:42 PM"
Submitted_Date_split = Submitted_Date[0:-3]
print(Submitted_Date_split)
b = datetime.strptime(Submitted_Date_split, "%d/%b/%y %H:%M%S")
print(a)
print(b)
print((a-b).days)


# In[45]:





# In[ ]:





# In[ ]:





# In[ ]:




