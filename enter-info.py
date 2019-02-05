#! python3
# enter-info.py - scrapes information from an excel sheet and enters information into zen planner


import time, os, openpyxl, re
from selenium import webdriver
from selenium.webdriver.common.keys import Keys

def zenLogin():
    dashboard = 'https://studio.zenplanner.com/zenplanner/studio/index.html#/main/iframe/zenplanner/studio/welcome/my-dashboard.cfm'

    try:
        browser = webdriver.Firefox()
        browser.get('https://studio.zenplanner.com/zenplanner/studio/index.html#/login')
    except Exception as ex:
        print('error: ' + str(ex))
        before = input()

    # Wait for login
    print('Log into ZenPlanner\n')
    while browser.current_url != dashboard:
        None

    return browser

def addPerson(browser, person):

    # --- Enter personal information ---
    browser.get('https://studio.zenplanner.com/zenplanner/studio/index.html#/main/iframe/zenplanner/studio/newPerson/profile.cfm?tt=1548699725827&tt=1548700656460')
    waitForPage(browser, 'https://studio.zenplanner.com/zenplanner/studio/index.html#/main/iframe/zenplanner/studio/newPerson/profile.cfm?tt=1548699725827&tt=1548700656460')

    try:
        browser.switch_to.frame('idTheIframe')
        # Enter name
        browser.find_element_by_id('idFirstName').send_keys(person.get('firstName', 0))
        browser.find_element_by_id('idLastName').send_keys(person.get('lastName', 0))
        # Click next
        browser.find_element_by_css_selector('input.btn:nth-child(6)').click()
        # Enter email address
        browser.find_element_by_id('idPersonEmailAddress').send_keys(person.get('email', 0))
        # Save and Finish
        browser.find_element_by_css_selector('input.btn:nth-child(6)').click()
    except:
        print('\nERROR ENCOUNTERED: program terminating.')

    # --- Sign waiver ---
    if person.get('docSigned', 0):
        browser.find_element_by_css_selector('.alertBox').click()
        browser.find_element_by_css_selector('tr.item:nth-child(5) > td:nth-child(2) > a:nth-child(1)').click()
        browser.find_element_by_css_selector('.err').click()
        browser.find_element_by_css_selector('#idSignatureType_offline').click()
        browser.find_element_by_css_selector('#idSubmit').click()
    return None

# Get excel document
def getExcelDoc():
    waiverDoc = openpyxl.load_workbook('people.xlsx')
    return waiverDoc

# extract information from document and return a person
def getPersonInfo(sheet):

    # Separate first and last name
    firstName, lastName = sheet['A2'].value.split()

    # get email
    email = sheet['B2'].value
    if email != None:
        email.replace('\s', '')

    # Determine whether waiver has been signed
    signed = True
    if(sheet['C2'].value != None and "X" in sheet['C2'].value):
        signed = False

    person = {'firstName': firstName, 'lastName': lastName, 'email': email, 'docSigned': signed}
    print('Inputting ' + firstName + ' ' + lastName + ' - ' + email)

    return person

def deleteRow(sheet):
    try:
        sheet.delete_rows(2)
        return
    except PermissionError:
        print('---------------------------------')
        print('-- ERROR: CLOSE EXCEL DOCUMENT --')
        print('---------------------------------')
        while(True):
            try:
                sheet.delete_rows(2)
                return
            except:
                None

def saveDoc(doc):
    try:
        doc.save('people.xlsx')
        return
    except:
        print('---------------------------------')
        print('-- ERROR: CLOSE EXCEL DOCUMENT --')
        print('---------------------------------')
        while(True):
            try:
                doc.save('people.xlsx')
                return
            except:
                None

# Waits for webpage to finish loading
def waitForPage(browser, page):
    while browser.current_url != page:
        None
    time.sleep(2)
    return None

# ------- Main code -------
print('Running...\n')

browser = zenLogin()
doc = getExcelDoc()
sheet = doc.active
doc.close()
while(sheet['A2'].value != None):
    addPerson(browser, getPersonInfo(sheet))
    deleteRow(sheet)
saveDoc(doc)

browser.get('https://studio.zenplanner.com/zenplanner/studio/index.html#/main/iframe/zenplanner/studio/welcome/my-dashboard.cfm')
print('Finished')