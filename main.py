from classes.tauinvcl import Tauinv
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait  # available since 2.4.0
from selenium.webdriver.support import expected_conditions as EC  # available since 2.26.0
from datetime import date
from datetime import datetime

import pyautogui
import time
import pandas as pd

import tkinter as tk


import os
import glob

# ------------------------------------------------------------------------------------------------------------------ #
# ______________  ___   ____   __   _______    _______  __________
#       / /      /   |  |  |  | |  |   __  |  |   __  |    | |
#      / /      / /| |  |  |  | |  |  |_|  |  |   __  |    | |
#     / /      / /_| |  |  |  | |  |   _ --|  |   __  |    | |
#    / /      /  __  |  |  |__| |  |  |_|  |  |   __  |    | |
#   /_/      /__/  \_|  |_______|  |_______|  |_______|    |_|
# ------------------------------------------------------------------------------------------------------------------ #
# ****************************************************************************************************************** #
# *                                                                                                                * #
# *                                                                                                                * #
# *                                                                                                                * #
# ****************************************************************************************************************** #

# ------------------------------------------------- INSTRUCTIONS --------------------------------------------------- #

# This is a script which uploads invoices into Taulia portal to SEMPRA clients for Cordoba Corporation
# The script uses selenium and Firefox webdriver (geckodriver) libraries.
# # The main driver is the data stored in CSVFPATH string variable, contains required data for each invoice
# # Each data required is based on the fields, stored in colNames list array
# # The script after logging in the portal > uploads invoice > loops at invoice level for each row in csv
# # The process flow of the script is:

# 1. Using pandas, load files from csv file ("C:\Users\Victor Song\OneDrive - Cordoba Corp\VSMAIN\_TAUBOTDATA")
# 2. Calls initiateup subprocedure to launch webdriver and log in and select client at TAUPORTAL
# 3. Call uploadtauinv, passes the initiated webdriver into uploadtauinv sub procedure, and check if DRAFT or FINAL
# 4. Invoice PDF are located at ("C:\Users\Victor Song\OneDrive - Cordoba Corp\VSMAIN\_TAUBOTDATA\PDFINV")
# 5. Loops to the next row from csv file (each row one invoice)
# 6. If FINAL, the script will delete the invoice from repository and save a csv copy at PDFINVPATH


# ******************************************To RUN THE SCRIPT********************************************************#

# 1. Verify the CONSTANTS variables, make sure the CSV and PDF invoices reconcile
# 2. Run main.py and do not move the mouse (pyautogui is needed in this script)
# 3. The webdriver closes the process once it finishes w/o/e. (Returns 0)

# -----------------------------------------------------------------------------------------------------------vs------#

# global CONSTANTS  ------------------------------
UPLOADTYPE = 'Final'  # Draft or Final
GECKOFPATH = r'C:\Users\Victor Song\PYP\TAUBOT\webdriver\geckodriver.exe'  # PATH for the .exe geckodriver file
CSVFPATH = r'C:\Users\Victor Song\OneDrive - Cordoba Corp\VSMAIN\_TAUBOTDATA\taupload.csv'  # dir where csv file is located, which contains the fields required for each invoice
PDFINVFPATH = r'C:\Users\Victor Song\OneDrive - Cordoba Corp\VSMAIN\_TAUBOTDATA\PDFINV'  # dir where invoices should be copy/pasted to (make sure you keep a copy, the script deletes the invoice)
RECORDSFPATH = r'C:\Users\Victor Song\OneDrive - Cordoba Corp\VSMAIN\_TAUBOTDATA\RECORDS'  # dir where script saves the records of each upload session


# end of global CONSTANTS ------------------------


# List of TAUBOT sub procedures below ----------------------------

def initiateup(GECKOFPATH):

    # create a browser instance
    driver = webdriver.Firefox(executable_path=GECKOFPATH)

    # use get method from driver and open the taulia webside as a get http request
    driver.get("client portal url")

    #make script wait a bit
    driver.implicitly_wait(15)

    # make sure the site you opened has taulia in the title
    assert "Taulia" in driver.title

    # login page: emailfield and passw and the textfield for login information
    emailfield = driver.find_element_by_class_name("tau-input")
    emailfield.send_keys('email account')

    passw = driver.find_element_by_id("password")
    passw.send_keys('password here')

    buttsub = driver.find_element_by_id("loginSubmitButton")
    buttsub.submit()

    driver.implicitly_wait(15)


    # after login page, select client, use select method to pick client ID per TAULIA website through a dropdown menu
    dropdown = Select(driver.find_element_by_id("buyerId"))
    Select.select_by_value(dropdown, 'df006cdf94d343209ab1709cac4410ab')


    # recycling buttsub button for the dropdown menu submit button
    buttsub = driver.find_element_by_name("Ok")
    buttsub.submit()

    time.sleep(4)

    return driver

def uploadtauinv(driver,taup,PDFINVFPATH, uptype = 'Final'):

    # webdriver update, to hold off a script line until certain HTML element finishes loading
    element = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "homePageHeading")))
    # this part is where the loop starts, python needs to obtain PO number from CSV file, which will match the invoice stored in specific filepath and use it as the upload
    # different PO testing uncomment this one
    # driver.execute_script("window.open('https://portal.na1prd.taulia.com/supplier/purchaseOrder/purchaseOrderDetails?poNumber=5660061088&companyCode.id=2c91816663f53e3b016471c03c420042', 'PO_window')")


    driver.execute_script('window.open("{}","PO_window");'.format(f'{str(taup.pourl)}'))

    # once the new tab is opened, you need to switch the to the opened tab
    driver.switch_to.window(driver.window_handles[1])

    # xpath locates the html createinvoice button
    time.sleep(2)
    createinvoice = driver.find_element_by_xpath("//*[@id='poDetailsMenuButtonCreateInvoice']")
    createinvoice.click()

    # wait submit button is loaded in the submit invoice form
    # elementpage = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "submitInvoice")))
    driver.switch_to.window(driver.window_handles[1])

    # scroll down to find the uploadbutton and for the pyautogui coordinates
    driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")

    # fields for invoice num loop
    invoicenum = driver.find_element_by_xpath("//*[@id='number']")
    invoicenum.send_keys(taup.number)

    # field for client contact loop
    ccontact = driver.find_element_by_xpath("//*[@id='contactPerson']")
    ccontact.send_keys(taup.contact)

    # field for qty, this should be always 1
    iqty = driver.find_element_by_xpath("//*[@id='lineItems[0].quantity']")
    iqty.send_keys("1")

    # field for invoice amount, loop
    iamt = driver.find_element_by_xpath("//*[@id='lineItems[0].pricePerUnit']")
    iamt.send_keys(taup.amt)

    # another scroll down just to make sure
    driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")

    # pyautogui portion to move and hover the mouse over the upload button, once it overs, the input object gets loaded resolution of (2736,1824)
    pyautogui.moveTo(432, 1147)
    pyautogui.move(0, 70)
    pyautogui.move(0, -70)
    pyautogui.move(0, 70)
    pyautogui.move(0, -70)

    # pyautogui portion to move and hover the mouse over the upload button, once it overs, the input object gets loaded resolution of (1920,1200)
    # pyautogui.moveTo(320, 749)
    # pyautogui.move(0, 50)
    # pyautogui.move(0, -53)
    # pyautogui.move(0, 54)
    # pyautogui.move(0, -55)

    # elementpage = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "//*[@id='comment']")))
    driver.implicitly_wait(20)

    # use file path to upload the invoice
    upinv = driver.find_element_by_xpath("/html/body/div[contains(@style,'display')]/input")
    upinv.send_keys(fr"{PDFINVFPATH}\{str(taup.number)}.pdf")

    time.sleep(5)

    elementpage = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CLASS_NAME, "fileDeleteLink")))

    if uptype == 'Final':
        subinv = driver.find_element_by_id("submitInvoice")
        subinv.click()
        time.sleep(1)

        # #for 1920,1200 resolution
        # pyautogui.moveTo(811, 693)
        # pyautogui.click()

        # for 2736,1824 resolution
        pyautogui.moveTo(1141, 1014)
        time.sleep(1)
        pyautogui.click()

        time.sleep(4)
  #      element = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "pdfDownload")))

    elif uptype == 'Draft':
        subinv = driver.find_element_by_id("saveAsDraft")
        subinv.click()

        time.sleep(4)
        element = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "draft_save_here_link_id")))

    # driver.switch_to.alert.accept()

    # finalbut = driver.find_element_by_class_name("tau-button tau-button-no ui-button ui-widget ui-state-default ui-corner-all ui-button-text-only")
    # finalbut.click()

    print('Looping...')
    driver.switch_to.window(driver.window_handles[1])
    driver.close()
    time.sleep(3)
    driver.switch_to.window(driver.window_handles[0])
    driver.refresh()
    time.sleep(2)
    return driver

def savetocsv(df,RECORDSFPATH):
    timeStamp = date.today().strftime("%y%m%d")
    timeHour = datetime.now().strftime("%H%M")

    timeAppend = [timeStamp for i in range(len(df))]
    df['TIMESTAMP'] = timeAppend
    df.to_csv(rf'{RECORDSFPATH}\TAUPLOAD_{timeStamp}_{timeHour}.csv', index=False, header=False)

def get_historical_stats(fpath):

    colStats = ['PRONUM', 'AMT', 'INVNUM', 'PONUM', 'CONTACT', 'CLIENT', 'POURL','TIMESTAMP']
    arrDF = []
    result = []
    finalStr = ''

    rootDir = r'C:\Users\Victor Song\OneDrive - Cordoba Corp\VSMAIN\_TAUBOTDATA\RECORDS'

    for path, subfold, files in os.walk(rootDir):
        for name in files:
            arrDF.append(pd.read_csv(os.path.join(path,name),names=colStats))

    result = pd.concat(arrDF)

    totalUploads = len(result['PRONUM'])
    earliestYear = str(result['TIMESTAMP'].min())[0:2]
    earliestMonth = str(result['TIMESTAMP'].min())[2:4]
    earliestDate = f'{earliestMonth}/{earliestYear}'
    tauTimeTaken = (totalUploads/2)/60
    manualTimeTaken = (totalUploads * 5)/60
    netTimeTaken = "{:.1f}".format(manualTimeTaken - tauTimeTaken)

    finalStr += f'TAUBOT STATS SINCE {earliestDate}:\n'
    finalStr += f'-------------------------------------\n'
    finalStr += f'Total Invoices uploaded: {totalUploads}\n'
    finalStr += f'Total Estimated TAUBOT Upload usage time: {"{:.1f}".format(tauTimeTaken)} hours\n'
    finalStr += f'Total time saved: {netTimeTaken} hours'

    return finalStr

def check_invdata(fpath):

    colNames = ['PRONUM', 'AMT', 'INVNUM', 'PONUM', 'CONTACT', 'CLIENT',
                'POURL']  # Fields names for CSVFPATH, the fields are used to pandas df
    colDataTypes = {'PRONUM': str, 'AMT': str, 'INVNUM': str, 'PONUM': str, 'CONTACT': str, 'CLIENT': str,
                    'POURL': str}  # Fields datatype for CSVFPATH, the fields are used to pandas df

    df = pd.read_csv(fpath, dtype=colDataTypes, names=colNames, header=None)

    return f'{df["PRONUM"].count()}枚 loaded!'

def daily_checkin():

    driver = webdriver.Firefox(executable_path=r'C:\Users\Victor Song\PYP\TAUBOT\webdriver\geckodriver.exe')

    # use get method from driver and open the taulia webside as a get http request
    driver.get("https://forms.office.com/Pages/ResponsePage.aspx?id=-QTU9J4b5ke0fCCQXMFtR9L0rCpt0T5PqoUE6natRixUNk85R1dRTzBPTDRYTkE4NDU5RFIzVklLRC4u")

    # letting the HTTP request load, so it can find the checkbox elements
    time.sleep(4)

    # make the driver scroll down, trying to load remainder elements (maybe)
    driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")

    # sa1 and sa2 are the checkboxes for the Santa Ana locations
    sa1 = driver.find_element_by_xpath(
        "/html/body/div/div/div/div/div[1]/div/div[1]/div[3]/div[2]/div[2]/div/div[2]/div/div[11]/div/label/input")
    sa2 = driver.find_element_by_xpath(
        "/html/body/div/div/div/div/div[1]/div/div[1]/div[3]/div[2]/div[2]/div/div[2]/div/div[12]/div/label/input")

    # click method to checkmark
    sa1.click()
    sa2.click()

    # input name in First Last textbox
    nameBox = driver.find_element_by_xpath(
        "/html/body/div/div/div/div/div[1]/div/div[1]/div[3]/div[2]/div[4]/div/div[2]/div/div/input")
    nameBox.send_keys('Victor Song')
    time.sleep(1)
    # scroll down
    driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")

    # submit button & click
    submitButton = driver.find_element_by_xpath(
        "/html/body/div/div/div/div/div[1]/div/div[1]/div[3]/div[3]/div[1]/button/div")
    submitButton.click()

    # terminate driver once submitbutton is clicked
    driver.close()

def start_taubot():
    # -- CONSTANTS -- Change them as needed

    colNames = ['PRONUM', 'AMT', 'INVNUM', 'PONUM', 'CONTACT', 'CLIENT',
                'POURL']  # Fields names for CSVFPATH, the fields are used to pandas df
    colDataTypes = {'PRONUM': str, 'AMT': str, 'INVNUM': str, 'PONUM': str, 'CONTACT': str, 'CLIENT': str,
                    'POURL': str}  # Fields datatype for CSVFPATH, the fields are used to pandas df

    # -- End of Constants --

    df = pd.read_csv(CSVFPATH, dtype=colDataTypes, names=colNames, header=None)

    print('TAUBOT Invoicer Initializing...')
    print(f'Total Invoice to Upload: {len(df)} 枚')

    driver = initiateup(GECKOFPATH)  # initiates the webdriver
    time.sleep(2)

    # loops for each item(row) in csv file using colNames as fields
    for x in range(0, len(df)):
        # Initiates tauinvcl class and adds 1 row from CSV into __init__
        taup = Tauinv(
            str(df.loc[x]['INVNUM']),
            str(df.loc[x]['AMT']),
            str(df.loc[x]['PONUM']),
            str(df.loc[x]['CONTACT']),
            str(df.loc[x]['CLIENT']),
            str(df.loc[x]['POURL'])
        )

        time.sleep(2)
        uploadtauinv(driver, taup, PDFINVFPATH, UPLOADTYPE)
        print(f'Invoice {taup.number} uploaded successfully!')



    if UPLOADTYPE == 'Final':
        savetocsv(df, RECORDSFPATH)
        for x in glob.glob(fr"{PDFINVFPATH}\*.pdf"):
            os.remove(x)

    print('Job Done!')

    get_historical_stats(RECORDSFPATH)

    driver.close()
    # End of sub procedures list ------------------------------

# END OF TAUBOT sub procedures List  ----------------------------

#  Tkinter subs -------------------------------

def tk_tau_starter():
    start_taubot()

def tk_tau_stats():
    feedLabel = tk.Label(mainW,text=f'{get_historical_stats(RECORDSFPATH)}',anchor='w',justify=tk.LEFT).grid(row=3,pady=10,sticky= tk.W)

def tk_csv_stats():
    invCountLabel = tk.Label(mainW,text=f'{check_invdata(CSVFPATH)}',anchor='w',justify = tk.LEFT).grid(row=4,pady=10,sticky=tk.W)

#  END of Tkinter subs -------------------------------



mainW = tk.Tk()

mainW.title('CORDOBOT')
mainW.geometry('600x1000')
mainW['bg'] = '#7d7d7d'

# logo = tk.PhotoImage(file="corlogo.png")


timeStamp = tk.StringVar()
timeStamp.set(date.today().strftime("%m/%d/%y"))
invCountCheck = tk.StringVar()
invCountCheck.set(check_invdata(CSVFPATH))

# logoLabel = tk.Label(mainW,image=logo).grid(row=1,pady=10,sticky=tk.W)
mainLabel = tk.Label(mainW,text=f'Hello! My name is CORDOBOT\nToday is: {timeStamp.get()} ',anchor='w',justify=tk.LEFT).grid(row=2,pady=10,sticky=tk.W)
settingsLabel = tk.Label(mainW,text=f'----------------------------------------',anchor='w',justify=tk.LEFT).grid(row=3,pady=10,sticky=tk.W)



buttonCheckData = tk.Button(mainW,text='CHECK CSV DATA', command=tk_csv_stats).grid(row=7,pady=10,sticky=tk.W)
buttonStartTAU = tk.Button(mainW,text='RUN TAUBOT',command=tk_tau_starter).grid(row=8,pady=10,sticky=tk.W)
buttonStats = tk.Button(mainW,text='GET STATS',command=tk_tau_stats).grid(row=9,pady=10,sticky=tk.W)
buttonDailysign = tk.Button(mainW, text='Daily Sign in',command=daily_checkin).grid(row=10,pady=10,stick=tk.W)
buttonQuit = tk.Button(mainW,text='QUIT',command=mainW.quit).grid(row=15,pady=10,sticky=tk.W)

tk.mainloop()


