import subprocess
import sys
import os
import time
import json

try:
    from win32com.client import Dispatch
except:
    i = input("You don't have the pywin32 module, which is required for this to work. Press ENTER to install!")
    if i == "" or i != "":
        subprocess.check_call([sys.executable, "-m", "pip", "install", "pywin32"])

while True:
    target = input("Type the file/folder path here: ")
    try: 
        target = target.replace('"', '')
    except: 
        pass

    if os.path.exists(target) == False: 
        print("This file/folder does not exist. Make sure you input the FULL path, not just the name of the file.")
    else: 
        break

try:
    with open("user.json", 'r') as f:
        data = json.load(f)
    user = data[0]["user"]
    
    if os.path.exists("C:/Users/" + user) == False:
        print(0/0)

    id = input(f"Are you {user}? [Y/N]")
    if id.lower() == "n":
        print(0/0)
        
except:
    while True:
        user = input("Type your username (case sensitive) \nDon't know your username? Type /help\n")

        if user.lower() == "/help":
            print(r"Open the Windows search bar (in the bottom of your screen or in Start) and type C:\Users.")
            user = input("Your username will be the name of one of the folders. Now, type in your username: ")

        if os.path.exists("C:/Users/" + user) == False:
            print(r"Not a valid username. Check C:\Users")
            time.sleep(5)
        else:
            with open("user.json", 'w') as f:
                json.dump([{"user": user}], f)
            break
        

tname = os.path.splitext(os.path.basename(target))[0]
path = "C:/Users/" + user + "/AppData/Roaming/Microsoft/Windows/Start Menu/Programs/" + tname + ".lnk"

def err(txt):
    print(txt, "(exiting program in 5 seconds)")
    time.sleep(5)
    exit()

shell = Dispatch('WScript.Shell')

try: shortcut = shell.CreateShortCut(path) 
except: err("Unable to save shortcut. Did you enter your name correctly?")

try: shortcut.Targetpath = target
except: err("Unable to find file/folder. Did you type the path in correctly?")

shortcut.save()

print(tname, "has been added to the Start Menu!")
time.sleep(5)