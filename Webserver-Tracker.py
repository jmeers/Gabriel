import urllib.request
import smtplib
import xlrd
import time
import xlsxwriter
import datetime

'''
James Meers
MIT License 2018
-Soli Deo Gloria-
'''

#First Run Check
firstRun = 1
msg = ""

from xlrd import open_workbook
# Open Excel sheet
wb = open_workbook('data.xlsx')
sheet = wb.sheet_by_index(0)

# Make List
names = []
urls = []
states = [True] * (sheet.nrows-1)
descriptions = []

# Pull data from the Excel Sheet and put it into the list
for row in range(1, sheet.nrows):
    names.append(sheet.cell(row, 0).value)
    urls.append(sheet.cell(row, 1).value)
    descriptions.append(sheet.cell(row, 2).value)

# Email engine (just add msg)
def send_Email(msg):
    smtpObj = smtplib.SMTP('smtp.server.com',587)
    smtpObj.ehlo()
    smtpObj.starttls()
    smtpObj.login('from@email.com','password')
    smtpObj.sendmail('from@email.com','to@email.com','subject: ' + msg)
    smtpObj.quit()

# Message at start to know the Monitor has come back online
def start_message(msg):
    if firstRun == 1:
        print("Welcome Message")
        msg = "Web Server Monitoring System Online \n\n Web Server Monitoring System Online \n Time\t" + str(datetime.datetime.now())
        send_Email(msg)

# Check if web server is up
def fetch_urls(urls,names,states,descriptions,msg):
      for i, (url,name,state,description) in enumerate(zip(urls,names,states,descriptions)):
        try:
            # Web Server is up and running
            urllib.request.urlopen(url)
            print(url + "\t" + name + '\tServer Online' + " ", end="")
            # if state was false then server was down but now up. Send email to alert user
            if state == False:
                print(" - Server was DOWN but now UP")
                # SEND EMAIL
                msg = "Web Server Back UP - " + str(name) + "\n\nWeb Server Name:\t" + str(name) + "\nTime:\t\t" + str(datetime.datetime.now()) + "\nIP Address:\t" + str(url) + "\nDescription:\t" + str(description)
                send_Email(msg)
                # Change state to False so no more emails are sent until state changes
                states[i] = True
            print("\n")
             
        except urllib.error.HTTPError:
            # Server is down or not running
            print(url + "\t" + name + '\tServer Offline' + " ", end="")
            # if state was True then the server was up but now down. Send email to alert user
            if state == True:
                # If this is first run then print out this
                if firstRun == 1:
                    print(" - Server is DOWN at START")
                    # SEND EMAIL
                    msg = "Web Server Down at Start - " + str(name) + "\n\nWeb Server Name:\t" + str(name) + "\nTime:\t\t" + str(datetime.datetime.now()) + "\nIP Address:\t" + str(url) + "\nDescription:\t" + str(description)
                    send_Email(msg)
                    # Change state to False so no more emails are sent until state changes
                    states[i] = False
                # If this is NOT first run then print out this
                if firstRun == 0:
                    print(" - Server was UP but now DOWN")
                    # SEND EMAIL
                    msg = "Web Server Down - " + str(name) + "\n\nWeb Server Name:\t" + str(name) + "\nTime:\t\t" + str(datetime.datetime.now()) + "\nIP Address:\t" + str(url) + "\nDescription:\t" + str(description)
                    send_Email(msg)
                    # Change state to False so no more emails are sent until state changes
                    states[i] = False
            print("\n")
                  
        except urllib.error.URLError:
            # Server is down or not running
            print(url + "\t" + name + '\tServer Offline' + " ", end="")
            # if state was True then the server was up but now down. Send email to alert user
            if state == True:
                # If this is first run then print out this
                if firstRun == 1:
                    print(" - Server is DOWN at START")
                    # SEND EMAIL
                    msg = "Web Server Down at Start - " + str(name) + "\n\nWeb Server Name:\t" + str(name) + "\nTime:\t\t" + str(datetime.datetime.now()) + "\nIP Address:\t" + str(url) + "\nDescription:\t" + str(description)
                    send_Email(msg)
                    # Change state to False so no more emails are sent until state changes
                    states[i] = False
                # If this is NOT first run then print out this
                if firstRun == 0:
                    print(" - Server was UP but now DOWN")
                    # SEND EMAIL
                    msg = "Web Server Down - " + str(name) + "\n\nWeb Server Name:\t" + str(name) + "\nTime:\t\t" + str(datetime.datetime.now()) + "\nIP Address:\t" + str(url) + "\nDescription:\t" + str(description)
                    send_Email(msg)
                    # Change state to False so no more emails are sent until state changes
                    states[i] = False
            print("\n")

while True:
    start_message(msg)
    fetch_urls(urls,names,states,descriptions,msg)
    print(states)
    print("==========End Check==========")
    firstRun = 0
    #Time Between Checks
    time.sleep(180)
