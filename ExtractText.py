import sys
import tkinter as tk
from tkinter import filedialog

root = tk.Tk()
root.withdraw()

file_path = filedialog.askdirectory()
if len(file_path) == 0:
    sys.exit()

outData = dict()
outData["_header"] = "Script,Label,Menu,Variable,Lowest Undefined Comment"


# Function to find first occurrence of multiple string checks
# idk where to put this in the script
def MultiInStr(searchText):
    posResult = -1
    tempResult = -1
    equalityCheckCharacters = [" ", "=", "<", ">"]

    for char in equalityCheckCharacters:
        tempResult = searchText.find(char)
        if tempResult > 0 and (tempResult < posResult or posResult == -1):
            posResult = tempResult
    
    return posResult + 1

import os
for file_name in os.listdir(file_path):
    if not file_name.endswith(".rpy"):
        continue

    count = 0
    lines = []
    indentGroup = []
    with open(file_path + "/" + file_name) as fp:
        for line in fp:
            if len(line.strip()) > 0:
                indentGroup.append((len(line) - len(line.lstrip()))/4)
                lines.append(line.strip())
                count += 1

    fp.close()

    count = 0
    isInMenu = False
    seperator = ","

    labelName = ""
    menuName = ""
    varName = ""
    eqVal = ""

    for line in lines:
        if line[:5] == "label":
            #Save label name
            labelName = line[6:len(line) - 1]
        elif line[:4] == "menu":
            menuGroup = indentGroup[count]
            isInMenu = True
        elif line[0] == "\"" and line[-1] == ":" and isInMenu:
            #Log menu choice
            menuName = line[1:line.rfind("\"")].replace(",", "")
        elif line[:2] == "if":
            #Get variable name
            endPos = MultiInStr(line[2:])
            varName = line[3:endPos]
            #Get conditional equality and value
            eqVal = "'" + line[endPos:len(line) - 1].lstrip()
        elif line.find("not defined") > 0:
            #Check for duplicate record
            tempKey = (file_name, labelName, menuName, varName)
            if not (seperator.join(tempKey) in outData):
                outData[tempKey] = file_name[:len(file_name) - 4] + "," + labelName + "," + menuName + "," + varName + "," + eqVal
        elif isInMenu:
            if indentGroup[count] <= menuGroup:
                isInMenu = False
                menuName = ""
        count += 1

outFile = file_path + "/test.csv"

with open(outFile, 'w', newline='') as csvfile:
    for myValue in outData.values():
        csvfile.write(myValue + "\n")

csvfile.close()
