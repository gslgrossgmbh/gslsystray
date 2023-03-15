import pystray
import webbrowser
import os.path
import win32com.client as win32
import socket
import requests
import subprocess
import urllib.request
import sys

import pyautogui
import PIL.Image
from PIL import ImageGrab
from functools import partial
ImageGrab.grab = partial(ImageGrab.grab, all_screens=True)

version = 'v1.0.5'
programName = 'GSL Groß GmbH' + ' ' + version

def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

programPath = resource_path("")
logoPath = resource_path("logo.png")

imageName = socket.gethostname() + '.png'
logoImage = PIL.Image.open(logoPath)
toMail = "support@gsl-computer.de"
remoteSessionUrl = "https://www.gsl-computer.de/fernwartung-beauftragen/"
collaborationPdfUrl = "https://www.gsl-computer.de/GSL_IT-Support.pdf"
githubApiURL = "https://api.github.com/repos/gslgrossgmbh/gslsystray/releases/latest"

#Funktion: Outlook mit oder ohne Anhang oeffnen
def sendMailTo(attachment):

    #Initialisiert Outlook Mailfenster
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mail.To = toMail

    if attachment == True:
        #Erstellt Screenshot
        attachmentFile = resource_path("") + imageName
        screenshot = pyautogui.screenshot()
        screenshot.save(attachmentFile)
        
        #Haengt Anhang an
        mail.Attachments.Add(attachmentFile)

    #Ruft Outlookfenster mit Einstellungen auf
    mail.Display(True)

    if attachment == True:
        #Entfernt den Screenshot wieder
        os.remove(attachmentFile)

def loadGithubApi():
    result = False
    try:
        global apiResponse
        apiResponse = requests.get(githubApiURL, timeout=5)
        return apiResponse
    except requests.ConnectionError:
        print("No ethernet connection.")
        return False
    except requests.exceptions.Timeout:
        print("URL request timeout.")
        result = False
    return result

def downloadGithubVersion():
    try:            
        global newDownloadUrl
        newDownloadUrl = apiResponse.json()["assets"][0]["browser_download_url"]
        global newDownloadDirectory
        newDownloadDirectory = resource_path("gslsystray.exe")
        try:
            pass
            os.rename(resource_path("gslsystray.exe"), resource_path("gslsystray_old.exe"))
        except:
            pass
        urllib.request.urlretrieve(newDownloadUrl, newDownloadDirectory)
    except:
        print("Download not possible.")

def deleteOldExe():
    try:
        os.remove(resource_path("gslsystray_old.exe"))
    except:
        print("No old .exe found. Passing file deleting.")

#SysTray on_clicked Funktionen
def on_clicked(icon, item):
    if str(item) == "Ticket erstellen (mit Screenshot)":
        sendMailTo(True)
    elif str(item) == "Ticket erstellen":
        sendMailTo(False)
    elif str(item) == "Zusammenarbeit mit dem GSL-Support":
        webbrowser.open(collaborationPdfUrl)
    elif str(item) == "Download zum Fernwartungsmodul":
        webbrowser.open(remoteSessionUrl)
    elif str(item) == "Schließen":
        icon.stop()

#SytemTray aufbauen
icon = pystray.Icon("GSLSystray", logoImage, menu=pystray.Menu(
    pystray.MenuItem(programName, None, enabled=False),
    pystray.MenuItem("Ticket erstellen", on_clicked),
    pystray.MenuItem("Ticket erstellen (mit Screenshot)", on_clicked),
    pystray.MenuItem("Nüztliche Links", 
        pystray.Menu(pystray.MenuItem("Zusammenarbeit mit dem GSL-Support", on_clicked),
        pystray.MenuItem("Download zum Fernwartungsmodul", on_clicked))
    ),
    pystray.MenuItem("Schließen", on_clicked)
))

deleteOldExe()

#Pruefe Github Version
loadedGitApi = loadGithubApi()
if loadedGitApi != False:
    loadedGitApi_filled = True
    try:
        loadedGitApi.json()["name"]
    except:
        loadedGitApi_filled = False

if loadedGitApi != False:
    if loadedGitApi_filled == True:
        if loadedGitApi.json()["name"] == version:
            print("Version up to date")
        else:
            print("Version " + loadedGitApi.json()["name"] + " found. Downloading new version " + version)
            downloadGithubVersion()

            try:
                print("Starting new .exe.")
                subprocess.Popen(newDownloadDirectory, stdout=subprocess.PIPE, stderr=subprocess.PIPE, stdin=subprocess.PIPE)
            except:
                print("Could not start subprocess")

            sys.exit()
    else:
        print(loadedGitApi.json()["message"])
elif loadedGitApi == False:
    print("Github Url: " + githubApiURL + " not found.")

#Erzeuge SystemTray
icon.run()
