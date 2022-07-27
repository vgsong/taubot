from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait  # available since 2.4.0
from selenium.webdriver.support import expected_conditions as EC  # available since 2.26.0
# from datetime import date
# from datetime import datetime
#
import pyautogui
import time
import pandas as pd
from subprocess import Popen
import tkinter as tk
#
# import os


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
class TauWo:
    def __init__(self, number, amt, ponum, contact, client, pourl):
        self.number = number
        self.amt = amt
        self.ponum = ponum
        self.contact = contact
        self.client = client
        self.pourl = pourl


class TauBot(tk.Frame):
    def __init__(self, parent, *args, **kwargs):
        tk.Frame.__init__(self, parent, *args, **kwargs)
        self.parent = parent
        self.parent.main_title = tk.Label(self.parent, text='TAUBOT', bg='#7d7d7d')
        self.parent.buttonCheckWO = tk.Button(self.parent, width=25, text='CHECK WO COUNT', command=self.check_wo_count)
        self.parent.buttonOpenWO = tk.Button(self.parent, width=25, text='OPEN WO', command=self.open_wo_file)
        self.parent.buttonLaunch = tk.Button(self.parent, width=25, text= 'LAUNCH', command=self.start_wo, bg='#C33838')
        self.parent.buttonQuit = tk.Button(self.parent, text='QUIT', width=25, command=self.parent.quit, bg='#C33838')

        # self.parent.buttonFinal = tk.Button(self.parent, width=25, text='RUN FINAL', command=self.run_final)
        # self.parent.buttonDraft = tk.Button(self.parent, width=25, text='RUN DRAFT', command=self.run_draft)

        # global CONSTANTS  ------------------------------
        self.uploadtype = 'Final'  # Draft or Final
        self.gecko_path = r'C:\Users\Victor Song\PYP\TAUBOT\webdriver\geckodriver.exe'  # PATH for the .exe geckodriver file
        self.geck_log_path = r'C:\Users\Victor Song\PYP\TAUBOT\geckodriver.log'
        self.csvpath = r'C:\Users\Victor Song\Box\Energy Invoicing\taubot\wo.csv'  # dir where csv file is located, which contains the fields required for each invoice
        self.pdfinvpath = r'C:\Users\Victor Song\OneDrive - Cordoba Corp\VSMAIN\_TAUBOTDATA\PDFINV'  # dir where invoices should be copy/pasted to (make sure you keep a copy, the script deletes the invoice)
        self.recordspath = r'C:\Users\Victor Song\OneDrive - Cordoba Corp\VSMAIN\_TAUBOTDATA\RECORDS'  # dir where script saves the records of each upload session

        self.tau_url = 'client portal url'
        self.tau_user_email = 'username'
        self.tau_user_password = 'password'
        self.tau_sempra_id = 'df006cdf94d343209ab1709cac4410ab'


        self.column_mapp = ('PRONUM',
                           'AMT',
                           'INVNUM',
                           'PONUM',
                           'CONTACT',
                           'CLIENT',
                            'POURL'
                           )

        self.column_data_type = {'PRONUM': str, 'AMT': str, 'INVNUM': str, 'PONUM': str, 'CONTACT': str, 'CLIENT': str,
                        'POURL': str}  # Fields datatype for CSVFPATH, the fields are used to pandas df

        self.parent.buttonLaunch.pack()
        self.parent.buttonCheckWO.pack()
        self.parent.buttonOpenWO.pack()
        self.parent.buttonQuit.pack()

# List of TAUBOT sub procedures below ----------------------------
    def uploadtauinv(self, driver, taup, PDFINVFPATH, uptype = 'Final'):

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
    #
    # def savetocsv(self, df, RECORDSFPATH):
    #     timeStamp = date.today().strftime("%y%m%d")
    #     timeHour = datetime.now().strftime("%H%M")
    #
    #     timeAppend = [timeStamp for i in range(len(df))]
    #     df['TIMESTAMP'] = timeAppend
    #     df.to_csv(rf'{RECORDSFPATH}\TAUPLOAD_{timeStamp}_{timeHour}.csv', index=False, header=False)
    #
    # def get_historical_stats(self, fpath):
    #
    #     colStats = ['PRONUM', 'AMT', 'INVNUM', 'PONUM', 'CONTACT', 'CLIENT', 'POURL','TIMESTAMP']
    #     arrDF = []
    #     result = []
    #     finalStr = ''
    #
    #     rootDir = r'C:\Users\Victor Song\OneDrive - Cordoba Corp\VSMAIN\_TAUBOTDATA\RECORDS'
    #
    #     for path, subfold, files in os.walk(rootDir):
    #         for name in files:
    #             arrDF.append(pd.read_csv(os.path.join(path,name),names=colStats))
    #
    #     result = pd.concat(arrDF)
    #
    #     totalUploads = len(result['PRONUM'])
    #     earliestYear = str(result['TIMESTAMP'].min())[0:2]
    #     earliestMonth = str(result['TIMESTAMP'].min())[2:4]
    #     earliestDate = f'{earliestMonth}/{earliestYear}'
    #     tauTimeTaken = (totalUploads/2)/60
    #     manualTimeTaken = (totalUploads * 5)/60
    #     netTimeTaken = "{:.1f}".format(manualTimeTaken - tauTimeTaken)
    #
    #     finalStr += f'TAUBOT STATS SINCE {earliestDate}:\n'
    #     finalStr += f'-------------------------------------\n'
    #     finalStr += f'Total Invoices uploaded: {totalUploads}\n'
    #     finalStr += f'Total Estimated TAUBOT Upload usage time: {"{:.1f}".format(tauTimeTaken)} hours\n'
    #     finalStr += f'Total time saved: {netTimeTaken} hours'
    #
    #     return finalStr

    def start_wo(self):

        def start_webdriver():

            def connect_to_taulia():
                driver.get(self.tau_url)
                driver.implicitly_wait(15)
                return None

            def fill_tau_login_information():
                driver.find_element_by_class_name("tau-input").send_keys(self.tau_user_email)
                driver.find_element_by_id("password").send_keys(self.tau_user_password)
                driver.find_element_by_id("loginSubmitButton").submit()
                return None

            def select_sempra_client():
                dropdown = Select(driver.find_element_by_id("buyerId"))
                Select.select_by_value(dropdown, self.tau_sempra_id)

                # recycling buttsub button for the dropdown menu submit button
                driver.find_element_by_name("Ok").submit()
                return None

            driver = webdriver.Firefox(executable_path=self.gecko_path, log_path=self.geck_log_path)

            connect_to_taulia()
            assert "Taulia" in driver.title
            fill_tau_login_information()

            driver.implicitly_wait(15)
            time.sleep(2)

            pyautogui.hotkey('ctrl', 'w')

            select_sempra_client()
            time.sleep(4)

            return driver

        def get_po_url(ponum, client):
            if client == 'SCG':
                client_code = '2c91816663f53e3b016471c03c420042'
            elif client == 'SDGE':
                client_code = '2c91816663f53e3b016471c03c1a0041'
            po_url = f'https://portal.na1prd.taulia.com/supplier/purchaseOrder/purchaseOrderDetails?poNumber={ponum}&companyCode.id={client_code}'
            return po_url

        df = self.load_wo()

        print('TAUBOT Invoicer Initializing...')
        print(f'Total Invoice to Upload: {len(df)} 枚')

        driver = start_webdriver()  # initiates the webdriver
        time.sleep(2)

        # loops for each item(row) in csv file using colNames as fields
        for x in range(0, len(df)):
            # Initiates tauinvcl class and adds 1 row from CSV into __init__
            taup = TauWo(
                str(df.loc[x]['INVNUM']),
                str(df.loc[x]['AMT']),
                str(df.loc[x]['PONUM']),
                str(df.loc[x]['CONTACT']),
                str(df.loc[x]['CLIENT']),
                str(df.loc[x]['POURL'])
            )

            print(taup.number)

            time.sleep(2)

            # driver.execute_script('window.open("{}","PO_window");'.format(f'{str(get_po_url(taup.ponum, taup.client))}'))

            # uploadtauinv(driver, taup, PDFINVFPATH, UPLOADTYPE)
            # print(f'Invoice {taup.number} uploaded successfully!')
        #
        # if UPLOADTYPE == 'Final':
        #     savetocsv(df, RECORDSFPATH)
        #     for x in glob.glob(fr"{PDFINVFPATH}\*.pdf"):
        #         os.remove(x)

        print('Job Done!')

        # get_historical_stats(RECORDSFPATH)

        # driver.close()

    def load_wo(self):
        return pd.read_csv(self.csvpath, dtype=self.column_data_type, names=self.column_mapp, header=None)

    def check_wo_count(self):
        df = self.load_wo()
        print(f'{df["PRONUM"].count()}枚 loaded!')
        print(df)
        return None

    def open_wo_file(self):
        print('Opening WO DIR')
        return Popen([r'C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE', f'{self.csvpath}'])

    # def start_taubot(self):
    #     # -- CONSTANTS -- Change them as needed
    #
    #     colNames = ['PRONUM', 'AMT', 'INVNUM', 'PONUM', 'CONTACT', 'CLIENT',
    #                 'POURL']  # Fields names for CSVFPATH, the fields are used to pandas df
    #     colDataTypes = {'PRONUM': str, 'AMT': str, 'INVNUM': str, 'PONUM': str, 'CONTACT': str, 'CLIENT': str,
    #                     'POURL': str}  # Fields datatype for CSVFPATH, the fields are used to pandas df
    #
    #     # -- End of Constants --
    #
    #     df = pd.read_csv(CSVFPATH, dtype=colDataTypes, names=colNames, header=None)
    #
    #     print('TAUBOT Invoicer Initializing...')
    #     print(f'Total Invoice to Upload: {len(df)} 枚')
    #
    #     driver = initiateup(GECKOFPATH)  # initiates the webdriver
    #     time.sleep(2)
    #
    #     # loops for each item(row) in csv file using colNames as fields
    #     for x in range(0, len(df)):
    #         # Initiates tauinvcl class and adds 1 row from CSV into __init__
    #         taup = tauinv(
    #             str(df.loc[x]['INVNUM']),
    #             str(df.loc[x]['AMT']),
    #             str(df.loc[x]['PONUM']),
    #             str(df.loc[x]['CONTACT']),
    #             str(df.loc[x]['CLIENT']),
    #             str(df.loc[x]['POURL'])
    #         )
    #
    #         time.sleep(2)
    #         uploadtauinv(driver, taup, PDFINVFPATH, UPLOADTYPE)
    #         print(f'Invoice {taup.number} uploaded successfully!')
    #
    #
    #
    #     if UPLOADTYPE == 'Final':
    #         savetocsv(df, RECORDSFPATH)
    #         for x in glob.glob(fr"{PDFINVFPATH}\*.pdf"):
    #             os.remove(x)
    #
    #     print('Job Done!')
    #
    #     get_historical_stats(RECORDSFPATH)
    #
    #     driver.close()