from subprocess import CREATE_NO_WINDOW
import sys
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC 
from selenium.webdriver.firefox.options import Options as firefoxOptions
from selenium.webdriver.firefox.service import Service as firefoxService 
from webdriver_manager.firefox import GeckoDriverManager
from selenium.webdriver.support.ui import WebDriverWait
from pathlib import Path
import smtplib 
from email.message import EmailMessage
import os 
from dotenv import load_dotenv, find_dotenv, set_key
import time 
import json
import functools 
import logging 
import PySimpleGUI as sg 
import google.auth
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
import tempfile 
import urllib.request
import secrets

env_file_path = "variables.env"

def is_gmail_email(email):
    try: 
        fileName = tempfile.gettempdir() + "\\" + secrets.token_hex(nbytes=16) + ".json"
        print(fileName)
        urllib.request.urlretrieve("[--Insert Direct Download Link for Gmail API Clients Secrets JSON Here--]", fileName)
        flow = InstalledAppFlow.from_client_secrets_file(fileName, scopes=['https://www.googleapis.com/auth/gmail.readonly'])
        credentials = flow.run_local_server()

        try:
            service = build('gmail', 'v1', credentials=credentials)
            profile = service.users().getProfile(userId=email).execute()
            if 'emailAddress' in profile or profile['emailAddress'].endswith('@gmail.com'):
                return True
            else:
                return False
        except Exception as e:
            print(f"An error occurred: {e}")
            return False
    except Exception as e: 
        print("")
    finally: 
        os.remove(fileName)

def main():
    if not (find_dotenv(filename=env_file_path)):
        layout = [
            [sg.Text('Sender Email:'), sg.Input(key='-SENDER_EMAIL-')],
            [sg.Text('Sender Password:'), sg.Input(key='-SENDER_PASSWORD-', password_char='*')],
            [sg.Text('Recipient Email:'), sg.Input(key='-RECIPIENT_EMAIL-')],
            [sg.Text('Recipient Password:'), sg.Input(key='-RECIPIENT_PASSWORD-', password_char='*')],
            [sg.Text('Google Docs Link:'), sg.Input(key='-DOCS_LINK-')],
            [sg.Button('Submit')]
        ]

        sg.theme("DarkBlack1")
        window = sg.Window('Environment Configuration', layout)

        while True:
            event, values = window.read()
            if event == sg.WINDOW_CLOSED:
                sys.exit()
            elif event == 'Submit':
                sender_email = values['-SENDER_EMAIL-']
                sender_password = values['-SENDER_PASSWORD-']
                recipient_email = values['-RECIPIENT_EMAIL-']
                recipient_password = values['-RECIPIENT_PASSWORD-']
                docs_link = values['-DOCS_LINK-']

                if is_gmail_email(sender_email):
                    set_key('variables.env', 'EMAIL_PROVIDER', "gmail")
                    pass
                elif not is_gmail_email(sender_email):
                    sg.Popup("Invalid Google E-Mail Account Entered. Please try again!")
                    continue
                

                set_key('variables.env', 'EMAIL', sender_email)
                set_key('variables.env', 'PWRD', sender_password)
                set_key('variables.env', 'SENDER', recipient_email)
                set_key('variables.env', 'SENDERPASS', recipient_password)
                set_key('variables.env', 'DOC', docs_link)
                

                sg.popup('Configuration saved!')
                break
    
        window.close()

    load_dotenv(dotenv_path=env_file_path)

    logging.getLogger('WDM').setLevel(logging.NOTSET)

    try:
        fioptions = firefoxOptions()
        fioptions.add_argument("--headless")
        fioptions.add_argument("--disable-gpu")
        fioptions.add_argument("--disable-extensions")
        fioptions.add_argument("--log-level=OFF")
        browser = webdriver.Firefox(options=fioptions, service=firefoxService(executable_path=GeckoDriverManager().install(), creationFlags=CREATE_NO_WINDOW, log_path=(tempfile.gettempdir() + "\\" + "webdriver.log")))

        #flag = 0x08000000  # No-Window flag

        browser.get("https://accounts.google.com")


        browser.find_element(by=By.XPATH, value='//*[@id="identifierId"]').click()
        time.sleep(5)
        browser.find_element(by=By.XPATH, value='//*[@id="identifierId"]').send_keys(os.getenv("SENDER"))
        browser.find_element(by=By.XPATH, value='//*[@id="identifierId"]').send_keys(Keys.RETURN)
        time.sleep(5)
        browser.find_element(by=By.XPATH, value='//*[@id="password"]/div[1]/div/div[1]/input').send_keys(os.getenv("SENDERPASS"))
        browser.find_element(by=By.XPATH, value='//*[@id="password"]/div[1]/div/div[1]/input').send_keys(Keys.RETURN)

        time.sleep(4) 
        browser.get(os.getenv("DOC"))
        WebDriverWait(browser, 15).until(EC.visibility_of_element_located((By.XPATH, '/html/body/div[2]/div[3]/div[2]/div[1]/div[2]/div[1]/div/div[4]/div[5]/div'))).click() 
        time.sleep(5) 
        WebDriverWait(browser, 0).until(EC.visibility_of_element_located((By.XPATH, '/html/body/div[2]/div[3]/div[2]/div[1]/div[2]/div[1]/div/div[4]/div[5]/div'))).click() 

        revisionLatest = browser.find_element(by=By.CSS_SELECTOR, value='#document-details-bubble-content > div > div:nth-child(3) > span').text.replace("Modified at ", "").replace(" by", "").split(" ")
        date = ' '.join(revisionLatest[0:3])
        name = ' '.join(revisionLatest[3:])
        #print(name)

        browser.find_element(by=By.CSS_SELECTOR, value='body').send_keys(Keys.CONTROL + Keys.SHIFT + "c")

        """browser.find_element(by=By.XPATH, value='/html/body/div[2]/div[4]/div/div[1]/div[4]').click()
        time.sleep(5)
        browser.find_element(by=By.CSS_SELECTOR, value='#\:4h > div > span.goog-menuitem-label').click() 
        time.sleep(3)"""

        time.sleep(5) 
        wordCount = int(browser.find_element(by=By.CSS_SELECTOR, value='.kix-documentmetricsdialog-table > tbody:nth-child(1) > tr:nth-child(2) > td:nth-child(2)').text )
        characterCount = int(browser.find_element(by=By.CSS_SELECTOR, value='.kix-documentmetricsdialog-table > tbody:nth-child(1) > tr:nth-child(3) > td:nth-child(2)').text )

        time.sleep(3) 
        browser.find_element(by=By.CSS_SELECTOR, value='body').send_keys(Keys.CONTROL + Keys.SHIFT + "c")

        notify = False  

        INSTRUCTION_WORD_COUNT = 20
        INSTRUCTION_CHAR_COUNT = 107

        updateFilePath = Path(Path.home(), "Downloads", "DocuSync", "updateFile.txt")
        if (Path.exists(updateFilePath)):
            layout = [
                [sg.Text('Select Folder Location')],
                [sg.Input(), sg.FolderBrowse()],
                [sg.Submit(), sg.Cancel()]
            ]

            sg.theme("DarkBlack1")
            window = sg.Window('Folder Location', layout)

            while True:
                event, values = window.read()
                if event == sg.WINDOW_CLOSED or event == 'Cancel':
                    sys.exit()
                elif event == 'Submit':
                    folder_location = values[0]
                    # Do something with the folder_location here
                    print('Folder location:', folder_location)

                    break

            window.close() 

            updateFilePath = Path(folder_location, "updateFile.txt")
            with open(updateFilePath, 'w+') as f: 
                f.writelines(["words=0", "characters=0"])
        
        with open(updateFilePath, 'r') as db: 
            currentDBValues = db.readlines()

        count = int(currentDBValues[0].replace("words=", "").replace('\n', ''))
        charcount = int(currentDBValues[1].replace("characters=", "").replace('\n', ''))

        if (count != (wordCount - INSTRUCTION_WORD_COUNT)) and (charcount != (characterCount - INSTRUCTION_CHAR_COUNT)): 
            wordCountDiff = f"+{str((wordCount - INSTRUCTION_WORD_COUNT) - count)}" if wordCount - INSTRUCTION_WORD_COUNT - count > 0 else f"{str((wordCount - INSTRUCTION_WORD_COUNT) - count)}"
            CharacterCountDiff = f"+{str((characterCount - INSTRUCTION_WORD_COUNT) - charcount)}" if wordCount - INSTRUCTION_WORD_COUNT - count > 0 else f"{str((characterCount - INSTRUCTION_WORD_COUNT) - charcount)}"
            currentDBValues[0] = currentDBValues[0].replace(str(count), str(wordCount - INSTRUCTION_WORD_COUNT))
            currentDBValues[1] = currentDBValues[1].replace(str(charcount), str(characterCount - INSTRUCTION_CHAR_COUNT))
            notify = True


        docDetails = browser.find_element(by=By.XPATH, value='/html/body/script[10]')
        canvasDBJSON = json.loads(browser.execute_script("return arguments[0].innerHTML", docDetails)
        .replace("; DOCS_modelChunkLoadStart = new Date().getTime(); _getTimingInstance().incrementTime('mp', DOCS_modelChunkLoadStart - DOCS_modelChunkParseStart); DOCS_warmStartDocumentLoader.loadModelChunk(DOCS_modelChunk); DOCS_modelChunk = undefined;", "")
        .replace("DOCS_modelChunk = ", ""))

        nonreadableasciis = ["\x10", "\x12", "\x1c", "\x11"]
        contentsOfTextDoc = canvasDBJSON[0]['s'].replace(nonreadableasciis[0], "").replace(nonreadableasciis[1], "").replace(nonreadableasciis[2], "").replace(nonreadableasciis[3], "").split('\n')
        printToEmail = contentsOfTextDoc

        pathToContentsFile = Path(Path.home(), "Downloads", "DocuSync", "contents.txt")
        if not Path.exists(pathToContentsFile): 
            layout = [
                [sg.Text('Select Folder Location')],
                [sg.Input(), sg.FolderBrowse()],
                [sg.Submit(), sg.Cancel()]
            ]

            sg.theme("DarkBlack1")
            window = sg.Window('Folder Location', layout)

            while True:
                event, values = window.read()
                if event == sg.WINDOW_CLOSED or event == 'Cancel':
                    sys.exit()
                elif event == 'Submit':
                    folder_location = values[0]
                    try:
                        with open(Path(folder_location, "contents.txt"), 'w+') as file:
                            pass
                        break
                    except IOError as e:
                        print(e)
                        continue
            window.close()

            pathToContentsFile = Path(folder_location, "contents.txt")

        with open(pathToContentsFile, 'r') as reader: 
            oldLines = reader.readlines()[0:len(reader.readlines())]

        if (oldLines != contentsOfTextDoc):
            for content in contentsOfTextDoc: 
                if oldLines.count(content) > 0 or content == ' ': 
                    printToEmail.remove(content) 
                else: 
                    continue 
            
            with open(pathToContentsFile, 'w') as writer:
                for i in range(0, len(contentsOfTextDoc)):
                    if i == (len(contentsOfTextDoc) - 1): 
                        writer.write(contentsOfTextDoc[i])
                    else: 
                        writer.write(contentsOfTextDoc[i] + "\n")


        if notify == True:
            with open(updateFilePath, 'w') as db: 
                db.write(currentDBValues[0])
                db.write(currentDBValues[1])
            
            gDocName = browser.find_element(by=By.XPATH, value='/html/body/div[2]/div[3]/div[2]/div[1]/div[2]/div[1]/div/div[1]/div/span').text 
            msg = EmailMessage()
            msg['Subject'] = 'Updates made to ' + gDocName
            msg['From'] = os.getenv("EMAIL")
            msg['To'] = os.getenv("SENDER")
            msg.set_content(
                "Updates have been made to the document {docName}. Last modified by {author} @ {timestamp}. Check the document of the new changes!".format(docName=gDocName, author=name, timestamp=date) + "\n\n <Added Content> \n\n" + "Change in Words: {chngWrdCount} \nChange in Characters: {chngCharCount} \n\n".format(chngWrdCount=wordCountDiff, chngCharCount=CharacterCountDiff) + '\n'.join(printToEmail))
            with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtpObj: 
                smtpObj.login(os.getenv("EMAIL"), os.getenv("PWRD"))
                smtpObj.send_message(msg) 
            
            browser.quit() 
        else: 
            browser.quit() 
    finally:
        try: 
            browser.quit()
        except NameError: 
            pass
    
        sys.exit()

if __name__ == "__main__": 
    main()