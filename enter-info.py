#! python3
# enter-info.py - scrapes information from an excel sheet and enters information into zen planner


import time
from selenium import webdriver
from selenium.webdriver.common.keys import Keys

def zenLogin():
    # print('Email: ', end='')
    # email = input()
    # print('Password: ', end='')
    # password = input()
    # print('\nLogging in...')

    email = 'obstacleathletics@gmail.com'
    password = 'Deerpark=0'

    browser = webdriver.Firefox()
    browser.get('https://studio.zenplanner.com/zenplanner/studio/index.html#/login')

    # Enter email address
    emailElem = browser.find_element_by_name('email')
    emailElem.send_keys(email)

    # Enter password
    pwdElem = browser.find_element_by_id('idPassword')
    pwdElem.send_keys(password)

    # Log in
    loginElem = browser.find_element_by_id('idSignIn')
    loginElem.click()

    waitForPage(browser, 'https://studio.zenplanner.com/zenplanner/studio/index.html#/main/iframe/zenplanner/studio/welcome/my-dashboard.cfm')
    return browser

def addPerson(browser, first, last, email, doc_signed):
    browser.get('https://studio.zenplanner.com/zenplanner/studio/index.html#/main/iframe/zenplanner/studio/newPerson/profile.cfm?tt=1548699725827&tt=1548700656460')
    waitForPage(browser, 'https://studio.zenplanner.com/zenplanner/studio/index.html#/main/iframe/zenplanner/studio/newPerson/profile.cfm?tt=1548699725827&tt=1548700656460')

    print('Enter name')
    browser.switch_to.frame('idTheIframe')
    # Enter name
    browser.find_element_by_id('idFirstName').send_keys(first)
    browser.find_element_by_id('idLastName').send_keys(last)
    # Click next
    browser.find_element_by_css_selector('input.btn:nth-child(6)').click()

    return None

def waitForPage(browser, page):
    while browser.current_url != page:
        None
    time.sleep(2)
    return None



# ------- Main code -------

browser = zenLogin()
addPerson(browser, 'Mike', 'Kennedy', 'email@gmail.com', True)