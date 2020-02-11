# This is a "list of programs" script, by Omar Siddiqi
# Credit due to open source code: http://www.blog.pythonlibrary.org/2010/03/03/finding-installed-software-using-python/
# Requirements:
# Get a list of what is currently installed on your machine
# Must pip install pywin32
# Must run IDE like Pycharm as administrator
#
#
#
#

import wmi
from winreg import (HKEY_LOCAL_MACHINE, KEY_ALL_ACCESS, OpenKey, EnumValue, QueryValueEx)
import os
import re
from difflib import SequenceMatcher

################################################### find all .EXE files in downloads folder and append to exeList
fileName = ""
exeList = []

for root, dirs, files in os.walk(r"C:\Users\abdom\Downloads"):
    for file in files:
        if file.endswith(".exe"):
            #print(os.path.join(root, file))
            fileName = (str(file)[:-4]).replace("_"," ") # str() to convert file to string;
                                                         # [:-4] to remove last 4 '.exe' from string
                                                         # ().replace("_"," ") replace "_" with " "
            #print(fileName) # print file name
            res1 = re.sub(r'(?<=[a-z])(?=[A-Z])', ' ', fileName) # insert white space before any caps
            res2 = re.sub(r'[^a-zA-Z0-9\n\.]', ' ', res1) #res2 = re.findall(r'\w+', res1) # remove all special characters like white space and hyphen
            res2 = res2.replace("Setup", "")
            res2 = res2.replace("Set Up", "")
            #print(res2) # print new filename with all words separated
            exeList.append(res2)

################################################### find all installed programs in registry and append to regList

r = wmi.WMI(namespace='DEFAULT').StdRegProv
result, names = r.EnumKey(hDefKey=HKEY_LOCAL_MACHINE,
                          sSubKeyName=r"Software\Microsoft\Windows\CurrentVersion\Uninstall")
regList = []

keyPath = r"Software\Microsoft\Windows\CurrentVersion\Uninstall"

for subkey in names: # for every subkey (Registry Key) in the registry
    try:
        path = keyPath + "\\" + subkey
        key = OpenKey(HKEY_LOCAL_MACHINE, path, 0, KEY_ALL_ACCESS)
        try:
            temp = QueryValueEx(key, 'DisplayName')
            displayName = str(temp[0])
            #print ('Display Name: ' + displayName)
            # print ('Regkey: ' + subkey)
            fileName = str(displayName).replace("_", " ")
            res1 = re.sub(r'(?<=[a-z])(?=[A-Z])', ' ', fileName)
            res2 = re.sub(r'[^a-zA-Z0-9\n\.]', ' ', res1) # res2 = re.findall(r'\w+', res1)
            regList.append(res2)
        except:
            pass
            #print('Display Name: Not Found')
            #print('Regkey: ' + subkey)

    except wmi.x_wmi as x:
        print ("Exception", x.com_error.hresult)

################################################### find similarities between regList and exeList and store in sameList
sameList = []
for x in range(len(regList)):
    for y in range(len(exeList)):
        if SequenceMatcher(None, regList[x], exeList[y]).ratio() >= 0.5:
            sameList.append(exeList[y])

################################################### remove those similarities from exeList, create a result list
################################################### append resultList to regList
resultList = list(set(exeList) ^ set(sameList))

for x in range(len(resultList)):
    regList.append(resultList[x])

################################################### print the final list of installed and unpackaged software
print('Size of this list:', len(regList))
print('Final list of installed and unpackaged software:\n')
print(*regList, sep = "\n")
