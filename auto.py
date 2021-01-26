
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
import time
import datetime
from docx import Document
import subprocess


import pyperclip
import os
from tqdm import tqdm

def yes_or_no(question):
    reply = str(input(question+' (y/n): ')).lower().strip()
    if reply[0] == 'y':
        return True
    if reply[0] == 'n':
        return False
    else:
        return yes_or_no("Uhhhh... please enter ")

def JobIDSearch():
    ID = input("Job ID is: ")
    flag = 0
    mytable = browser.find_element_by_xpath("//table[@class='table table-striped table-bordered table-hover gridTable']")
    for row in mytable.find_elements_by_css_selector('tr'):
        if len(row.find_elements_by_tag_name('td')) > 2:
            
            if row.find_elements_by_tag_name('td')[4].text == ID:
                flag = 1
                print("found the match")
                buttons = browser.find_elements_by_xpath("//*[contains(text(), '{0}')]".format(row.find_elements_by_tag_name('td')[4+1].text))
                
                for btn in buttons:
                    btn.click()
                break
    if flag == 0:
        return False
    elif flag == 1:
        return True

def webscraping():
    anothertable = browser.find_elements_by_xpath("//table[@class='table table-bordered']")
    FNAME = ""
    SAL = ""
    LNAME = ""
    JOBTITLE = ""
    COMPANY = ""
    ADDRESS1 = ""
    CITY = ""
    PROVINCE = ""
    POSTAL = ""
    for table in anothertable:

        for ir in table.find_elements_by_css_selector('tr'):

            for ic in tqdm(range(len(ir.find_elements_by_tag_name('td')))):
                # print(ir.find_elements_by_tag_name('td')[ic].text)
                if ir.find_elements_by_tag_name('td')[ic].text == "Job Title:":
                    JOBTITLE = ir.find_elements_by_tag_name('td')[ic+1].text
                                
                if ir.find_elements_by_tag_name('td')[ic].text == "Organization:":
                    COMPANY = ir.find_elements_by_tag_name('td')[ic+1].text
                
                if ir.find_elements_by_tag_name('td')[ic].text == "Job Contact First Name:":
                    FNAME  =  ir.find_elements_by_tag_name('td')[ic+1].text
                
                if ir.find_elements_by_tag_name('td')[ic].text == "Job Contact Last Name:":
                    LNAME = ir.find_elements_by_tag_name('td')[ic+1].text
                if ir.find_elements_by_tag_name('td')[ic].text == "Salutation:":
                    SAL = ir.find_elements_by_tag_name('td')[ic+1].text
                if ir.find_elements_by_tag_name('td')[ic].text == "Address Line One:":
                    ADDRESS1 = ir.find_elements_by_tag_name('td')[ic+1].text
                if ir.find_elements_by_tag_name('td')[ic].text == "Address Line Two:":
                    ADDRESS1 = ADDRESS1 + ", " + ir.find_elements_by_tag_name('td')[ic+1].text   
                if ir.find_elements_by_tag_name('td')[ic].text == "City:":
                    CITY = ir.find_elements_by_tag_name('td')[ic+1].text
                if ir.find_elements_by_tag_name('td')[ic].text == "Province / State:":
                    PROVINCE = ir.find_elements_by_tag_name('td')[ic+1].text
                if ir.find_elements_by_tag_name('td')[ic].text == "Postal Code / Zip Code:":
                    POSTAL = ir.find_elements_by_tag_name('td')[ic+1].text
    if FNAME == "":
        SAL = "Whom It May Concern"
        FNAME = ""
        LNAME = ""
    if (FNAME != "") and SAL == "":
        SAL = "SAL"
    return JOBTITLE, COMPANY, FNAME, LNAME, SAL, ADDRESS1, CITY, PROVINCE, POSTAL

def CLGEN(JOBTITLE, COMPANY, FNAME, LNAME, SAL, ADDRESS1, CITY, PROVINCE, POSTAL):
    doc=Document('TEMPLATE.docx')
    DATE = datetime.datetime.today().strftime('%B %d, %Y')
    TEAM = input("TEAM NAME: ")
    JOBDESCRIP = input("JOB DESCRIPTION: ")

    PERSONAL = "via SFU FAS Co-op Office (fas_coop@sfu.ca)"

    for p in doc.paragraphs:
        if p.text.find("DATE")>=0:
            p.text=p.text.replace("DATE",DATE)
        if p.text.find("COMPANY")>=0:
            p.text=p.text.replace("COMPANY",COMPANY.lstrip())
        if p.text.find("ADDRESS1")>=0:
            p.text=p.text.replace("ADDRESS1",ADDRESS1.lstrip())
        if p.text.find("CITY")>=0:
            p.text=p.text.replace("CITY",CITY.lstrip())
        if p.text.find("PROVINCE")>=0:
            p.text=p.text.replace("PROVINCE",PROVINCE.lstrip())
        if p.text.find("POSTAL")>=0:
            p.text=p.text.replace("POSTAL",POSTAL.lstrip())
        if p.text.find("JOBTITLE")>=0:
            p.text=p.text.replace("JOBTITLE",JOBTITLE.lstrip())
        if p.text.find("TEAM")>=0:
            p.text=p.text.replace("TEAM",TEAM.lstrip())
        if p.text.find("SAL")>=0:
            p.text=p.text.replace("SAL",SAL.lstrip())
        if p.text.find("FNAME")>=0:
            p.text=p.text.replace("FNAME",FNAME.lstrip())
        if p.text.find("LNAME")>=0:
            p.text=p.text.replace("LNAME",LNAME.lstrip())
        if p.text.find("JOBDESCRIP")>=0:
            p.text=p.text.replace("JOBDESCRIP",JOBDESCRIP.lstrip())
        if p.text.find("PERSONAL")>=0:
            p.text=p.text.replace("PERSONAL",PERSONAL)
    JOBTITLE = JOBTITLE.replace("/", " ")
    JOBTITLE = JOBTITLE.replace("&", "and")
    SAVEAS = COMPANY + "-"+ JOBTITLE
    print(SAVEAS)
    doc.save(SAVEAS + ".docx")
    
    pyperclip.copy(SAVEAS+'.pdf')
    spam = pyperclip.paste()


    opendoc = subprocess.run(["libreoffice", SAVEAS+".docx"])
    opendoc.returncode

    if yes_or_no("No Errors in docx?") == True:
        
        convert = subprocess.run(["libreoffice","--headless", "--convert-to", "pdf", SAVEAS+".docx"])
        convert.returncode
        kill = subprocess.run(["killall", "soffice.bin"])
        kill.returncode
    pyperclip.copy(SAVEAS+'.pdf')
    return SAVEAS
def Apply(filename):
    filename = filename + '.pdf'
    browser.find_element_by_tag_name('body').send_keys(Keys.CONTROL + Keys.HOME)
    time.sleep(1)
    buttons = browser.find_element_by_xpath("//button[@class='btn__default--text btn--default  applyButton']") #apply button
    buttons.click()

    

    browser.find_element_by_css_selector("input[type='radio'][value='customPkg']").click()
    time.sleep(1)
    newdoc = browser.find_element_by_xpath("//*[contains(text(), 'Click if you need to upload a new document')]")
    newdoc.click()
    senddoc = browser.find_element_by_xpath("//*[@id='fileUpload_docUpload']")
    senddoc.send_keys(os.getcwd()+"/"+filename)
    doctitle = browser.find_element_by_id("docName")
    doctitle.send_keys(filename)
    select = Select(browser.find_element_by_id('docType'))
    select.select_by_value('14')
    time.sleep(4)
    browser.find_element_by_xpath("//*[contains(text(), 'Upload Document')]").click()
    # time.sleep(3)
    browser.find_element_by_css_selector("input[type='radio'][value='customPkg']").click()
    browser.find_element_by_id("packageName").send_keys(filename)
    browser.find_element_by_xpath("//select[@name='14']/option[contains(text(), '{0}')]".format(filename)).click()
    select = Select(browser.find_element_by_id('requiredInPackage18')) #SIS
    select.select_by_index('1')
    select = Select(browser.find_element_by_id('requiredInPackage15')) #Resume
    select.select_by_index('1')
    select = Select(browser.find_element_by_id('requiredInPackage16')) #Transcript
    select.select_by_index('1')
    if yes_or_no("Good?") == True:
        browser.find_element_by_xpath("//*[@class='btn btn-primary']").click()

# options = webdriver.ChromeOptions()
# options.add_argument("--start-maximized")
browser = webdriver.Chrome()
browser.maximize_window()
# browser.set_window_size(1920,1080)
url = 'https://cas.sfu.ca/cas/login?message=Welcome+to+SFU+myExperience.%20Please+login+with+your+SFU+computing+ID.&allow=student,alumni&renew=true&service=https://myexperience.sfu.ca/sfuLogin.htm%3Faction%3Dlogin'
browser.get(url)
username = browser.find_element_by_id("username")
password = browser.find_element_by_id("password")
username.send_keys("YOURID")
time.sleep(0.5)
password.send_keys("YOURPASSWD")
browser.find_element_by_name("submit").click()
browser.get('https://myexperience.sfu.ca/myAccount/co-op/postings.htm')
time.sleep(1)
buttons = browser.find_elements_by_xpath("//*[contains(text(), 'Advance Search')]")

for btn in buttons:
    btn.click()
datelive = browser.find_element_by_id("dateLiveFrom")
tod = datetime.datetime.today()
daterange = input("date range: ")
d = datetime.timedelta(days = int(daterange))
a = tod - d

datelive.send_keys(a.strftime('%m/%d/%Y'))
buttons = browser.find_elements_by_xpath("//*[contains(text(), 'Search Job Postings')]")

for btn in buttons:
    btn.click()


buttons = browser.find_element_by_xpath("//button[@class='btn__hero btn--info drop-down__btn js--btn-toggle-side-menu has--ripple btn--ripple  has-ripple']")
buttons.click()
browser.execute_script("var all = document.getElementsByClassName('orgDivTitleMaxWidth'); for (var i = 0; i < all.length; i++) { all[i].style.maxWidth = '600px';}")


while JobIDSearch() == False:
    JobIDSearch()
    


browser.switch_to.window(browser.window_handles[1])
if yes_or_no("Proceed?") == True:

    JOBTITLE, COMPANY, FNAME, LNAME, SAL, ADDRESS1, CITY, PROVINCE, POSTAL = webscraping()

    filename = CLGEN(JOBTITLE, COMPANY, FNAME, LNAME, SAL, ADDRESS1, CITY, PROVINCE, POSTAL)



    if yes_or_no("Apply?") == True:
        Apply(filename)
        

while yes_or_no("Another One?") == True:

    try:
        browser.close()
        browser.switch_to.window(browser.window_handles[0])
        browser.refresh()
        time.sleep(1)
        browser.execute_script("var all = document.getElementsByClassName('orgDivTitleMaxWidth'); for (var i = 0; i < all.length; i++) { all[i].style.maxWidth = '600px';}")
    except:
        browser.switch_to.window(browser.window_handles[0])
        browser.refresh()
        time.sleep(1)
        browser.execute_script("var all = document.getElementsByClassName('orgDivTitleMaxWidth'); for (var i = 0; i < all.length; i++) { all[i].style.maxWidth = '600px';}")
    while JobIDSearch() == False:
        JobIDSearch()
    


    browser.switch_to.window(browser.window_handles[1])
    if yes_or_no("Proceed?") == True:
        JOBTITLE, COMPANY, FNAME, LNAME, SAL, ADDRESS1, CITY, PROVINCE, POSTAL = webscraping()

        filename = CLGEN(JOBTITLE, COMPANY, FNAME, LNAME, SAL, ADDRESS1, CITY, PROVINCE, POSTAL)
        if yes_or_no("Apply?") == True:
            Apply(filename)




            
