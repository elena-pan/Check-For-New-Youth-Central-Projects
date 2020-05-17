# Checks Youth Central if new projects have been posted since
# the last time the program was run, sends email with info on
# the newly posted project shifts

import requests, bs4, os, smtplib, re, webbrowser
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import oauth2 #oauth2.py
import tkinter
from tkinter import simpledialog
from datetime import datetime

# Put each item in past projects list as list
with open('Updated_project_list.txt', 'r') as f:
    past_projects = [line.strip() for line in f]

# Set up selenium headless browser
option = Options()
option.headless = True
driver = webdriver.Chrome('Path_to_chromedriver.exe', options=option)

project_list = []
project_info = []

email = 'youremail@gmail.com'
GOOGLE_CLIENT_ID = 'your_client_id'
GOOGLE_CLIENT_SECRET = 'your_client_secret'

def get_info():
    # Get login to youth central using Selenium
    driver.get('https://app.betterimpact.com/Login/Volunteer')
    
    driver.implicitly_wait(10)
    username = driver.find_element_by_id('UserName')
    username.clear()
    username.send_keys('your_username')
    
    password = driver.find_element_by_id('Password')
    password.clear()
    # Todo: Don't hardcode this
    password.send_keys('your_password')

    # Navigate to Opportunities List
    driver.find_element_by_id('SubmitLoginForm').click()
    driver.find_element_by_id('OpportunitiesMenu').click()
    wait = WebDriverWait(driver, 10)
    opportunitylist = wait.until(EC.element_to_be_clickable((By.LINK_TEXT, 'Opportunity List')))
    driver.find_element_by_link_text('Opportunity List').click()
    '''except Exception as e:
        # Error procedure
            
        errorFileName = 'Error log ' + datetime.now().strftime('%Y-%m-%d %H;%M;%S') + '.txt'
        errorFile = open(errorFileName, 'w')
        errorFile.write(str(e))
        errorFile.close()

        errorFileLocation = os.path.abspath(errorFileName)
        os.startfile(errorFileLocation)
        return'''

    # Make list of current opportunities and rewrite into past projects file
    projects = driver.find_elements_by_class_name('addMeToHash')
    for project in projects:
        project_list.append(project.text)

    update_projects = open('Updated_project_list.txt', 'w')
    for project in project_list:
        update_projects.write(project + '\n')

    if not past_projects == []:
        new_projects = [item for item in project_list if item not in past_projects]
        if not new_projects == []:
            for project in new_projects:
                driver.find_element_by_link_text(project).click()
                html_source = driver.page_source
                soupSelenium = bs4.BeautifulSoup(html_source,'html.parser')
                shifts = []
                for item in soupSelenium.select('#ShiftTbody > tr'):
                    shift_info = []
                    shift_info.append(item.select('.dateTd')[0].getText())
                    shift_info.append(item.select('.right.startTimeTd')[0].getText())
                    shift_info.append(item.select('.right.endTimeTd')[0].getText())
                    shift_info.append(item.select('.openingsCell,center')[0].getText())
                    shifts.append(shift_info)
                project_info.append({'Project':project, 'Info':shifts})
                driver.back()
                
            send_emails(project_info)
            
def input_code():
    root = tkinter.Tk()
    root.withdraw()

    # Input dialog
    code = simpledialog.askstring(title=' ', prompt='Authentication code: ')
    return code

def authenticate():
    
    # Get auth string using oauth2
    refreshFile = open('your_refresh_token.txt')
    refresh_token = refreshFile.read()
    refreshFile.close()
    if refresh_token == '':

        # If no refresh token has been obtained (new user) yet
        # Direct user to authentication url to get code
        url = oauth2.GeneratePermissionUrl(GOOGLE_CLIENT_ID)
        webbrowser.open_new_tab(url)
        auth_code = input_code()
        if auth_code == None:
            # User pressed cancel
            return 'Cancel'

        # Save refresh token
        try:
            response = oauth2.AuthorizeTokens(GOOGLE_CLIENT_ID, GOOGLE_CLIENT_SECRET, auth_code)
            access_token = response['access_token']
            refresh_token = response['refresh_token']

            refreshFile2 = open('your_refresh_token.txt', 'w')
            refreshFile2.write(refresh_token)
            refreshFile2.close()

            auth_string = oauth2.GenerateOAuth2String(email, access_token)
        except Exception as e:

            # Error procedure
            
            errorFileName = 'Error log ' + datetime.now().strftime('%Y-%m-%d %H;%M;%S') + '.txt'
            errorFile = open(errorFileName, 'w')
            errorFile.write(str(e))
            errorFile.close()

            errorFileLocation = os.path.abspath(errorFileName)
            os.startfile(errorFileLocation)
            return 'Error'

    # Get auth code using access code
    else:
        try:
            response = oauth2.RefreshToken(GOOGLE_CLIENT_ID, GOOGLE_CLIENT_SECRET, refresh_token)

        # If refresh token is revoked
        except Exception as e:
            open('your_refresh_token.txt', 'w').close()
            # Error procedure
            
            errorFileName = 'Error log ' + datetime.now().strftime('%Y-%m-%d %H;%M;%S') + '.txt'
            errorFile = open(errorFileName, 'w')
            errorFile.write(str(e))
            errorFile.close()

            errorFileLocation = os.path.abspath(errorFileName)
            os.startfile(errorFileLocation)
            return 'Error'

        access_token = response['access_token']
        auth_string = oauth2.GenerateOAuth2String(email, access_token)

    return auth_string

def send_emails(project_info):

    auth_string = authenticate()
    if auth_string == 'Cancel':
        return

    elif auth_string == 'Error':
        auth_string = authenticate()
        
    
    # Set up SMTP server and authenticate
    smtpObj = smtplib.SMTP('smtp.gmail.com', 587)
    smtpObj.ehlo()
    smtpObj.starttls()
    smtpObj.ehlo()
    smtpObj.docmd('AUTH', 'XOAUTH2 ' + auth_string)

    # Format email text
    email_text = ''
    for item in project_info:
        email_text += '\n' + item['Project']
        for shift in item['Info']:
            email_text += shift[0]+' '
            email_text += shift[1]+'-'
            email_text += shift[2]+', '
            email_text += 'Openings - ' + shift[3]
        email_text += '\n'

    # Send email

    try:
        smtpObj.sendmail(email, email,
        'Subject: New Youth Central Projects!!\n**New Youth Central Projects**\n%s' %(email_text))
        smtpObj.quit()

    except Exception as e:
        # Error procedure
            
        errorFileName = 'Error log ' + datetime.now().strftime('%Y-%m-%d %H;%M;%S') + '.txt'
        errorFile = open(errorFileName, 'w')
        errorFile.write(str(e))
        errorFile.close()

        errorFileLocation = os.path.abspath(errorFileName)
        os.startfile(errorFileLocation)
        return

get_info()
driver.quit()
