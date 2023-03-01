import pystray
import webbrowser
import PIL.Image
import os
import win32com.client as win32
import pyautogui
import socket
import requests
from PIL import ImageGrab
from functools import partial
ImageGrab.grab = partial(ImageGrab.grab, all_screens=True)


version = 'v1.0.1'
programName = 'GSL Groß GmbH' + ' ' + version

programPath = os.path.dirname(__file__)
logoPath = programPath + "\\logo.png"

imageName = socket.gethostname() + '.png'
logoImage = PIL.Image.open(logoPath)
toMail = "support@gsl-computer.de"
rsURL = "https://www.gsl-computer.de/fernwartung-beauftragen/"
updateURL = "https://api.github.com/repos/v2ray/v2ray-core/releases/latest"
#https://github.com/v2ray/v2ray-core/releases
#https://github.com/v2fly/v2ray-core/releases/tag/v4.31.0


#Funktion: Outlook mit oder ohne Anhang oeffnen
def sendMailTo(attachment):

    #Initialisiert Outlook Mailfenster
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mail.To = toMail

    if attachment == True:
        #Erstellt Screenshot
        attachmentFile = programPath + '\\' + imageName
        screenshot = pyautogui.screenshot()
        screenshot.save(attachmentFile)
        
        #Haengt Anhang an
        mail.Attachments.Add(attachmentFile)

    #Ruft Outlookfenster mit Einstellungen auf
    mail.Display(True)

    if attachment == True:
        #Entfernt den Screenshot wieder
        os.remove(attachmentFile)


def remoteSessionWebsite():
    webbrowser.open(rsURL)


def checkUpdateURL():
    result = False
    try:
        requests.get(updateURL, timeout=5)
        result = True
    except requests.exceptions.Timeout:
        result = False
    return result


def checkGithubVersion():
    if checkUpdateURL():
        try:
            response = requests.get(updateURL)
            return response.json()["name"]
        except:
            return "0.0.0"
    else:
        return False
        

def on_clicked(icon, item):
    if str(item) == "Ticket erstellen (mit Screenshot)":
        sendMailTo(True)
    elif str(item) == "Ticket erstellen":
        sendMailTo(False)
    elif str(item) == "RS-Client":
        remoteSessionWebsite()
    elif str(item) == "Schließen":
        icon.stop()


#SytemTray aufbauen
icon = pystray.Icon("GSLSystray", logoImage, menu=pystray.Menu(
    pystray.MenuItem(programName, None, enabled=False),
    pystray.MenuItem("Ticket erstellen", on_clicked),
    pystray.MenuItem("Ticket erstellen (mit Screenshot)", on_clicked),
    pystray.MenuItem("RS-Client", on_clicked),

    pystray.MenuItem("Schließen", on_clicked)
))


#Pruefe auf Github Version
checkGithubVersion = checkGithubVersion()
if checkGithubVersion == False:
    print("URL nicht gefunden")
elif checkGithubVersion == "0.0.0":
    print("Github Version nicht gefunden")
elif checkGithubVersion == version:
    print("Version aktuell")
elif checkGithubVersion != version:
    print("Version " + checkGithubVersion + " gefunden. Installiert ist " + version)

#Erzeuge SystemTray
icon.run()