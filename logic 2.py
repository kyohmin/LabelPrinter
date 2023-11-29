from openpyxl import load_workbook
import xlwings as xw
import os

# OPEN THE WORKBOOK ===============================

# xlsm_path = os.path.abspath('Data.xlsm')
# oneTimePath = os.path.abspath('oneTime.txt')
# serialNamePath = os.path.abspath('serialName.txt')
# serialNumberPath = os.path.abspath('serialNumber.txt')
# countPath = os.path.abspath('count.txt')
username = os.getlogin()
path = "C:\\Users\\" + username
xlsm_path = path + "\\LabelPrinter\\Data.xlsm"
oneTimePath = path + "\\LabelPrinter\\oneTime.txt"
serialNamePath = path + "\\LabelPrinter\\serialName.txt"
serialNumberPath = path + "\\LabelPrinter\\serialNumber.txt"
countPath = path + "\\LabelPrinter\count.txt"

oneTime, serialName, serialNumber, count = [], [], [], []
project, unit, serialDate = "", "", ""

dataBook = load_workbook(filename=xlsm_path, read_only = False, keep_vba = True)
dataBook.active = 0 # index of active sheet
dataSheet = dataBook.active



# Extracting one time entering data
# txtFile = open(oneTimePath, 'r', encoding="utf-8-sig")
txtFile = open(oneTimePath, 'r')

oneTime = txtFile.read().strip()
oneTime = oneTime.split("\n")
project = oneTime[0]
unit = oneTime[1]
serialDate = oneTime[2]

txtFile.close()

# Extracting serial names
# txtFile = open(serialNamePath, 'r', encoding="utf-8-sig")
txtFile = open(serialNamePath, 'r')

serialName = txtFile.read().strip()
serialName = serialName.split("\n")
repeatTime = len(serialName)

txtFile.close()

# Extracting serial numbers
# txtFile = open(serialNumberPath, 'r', encoding="utf-8-sig")
txtFile = open(serialNumberPath, 'r')

serialNumber = txtFile.read().strip()
serialNumber = serialNumber.split("\n")

txtFile.close()

# Extracting counts
# txtFile = open(countPath, 'r', encoding="utf-8-sig")
txtFile = open(countPath, 'r')

count = txtFile.read().strip()
count = count.split("\n")

txtFile.close()


# Inserting data into Data.xlsm ===================

dataSheet['A2'] = project
dataSheet["B2"] = unit
dataSheet["C2"] = serialDate
dataSheet["G2"] = repeatTime

for i in range(0,repeatTime):
    dataSheet["D" + str(i + 2)] = serialName[i]
for i in range(0,repeatTime):
    dataSheet["E" + str(i + 2)] = serialNumber[i]
for i in range(0,repeatTime):
    dataSheet["F" + str(i + 2)] = count[i]

dataBook.save(xlsm_path)
dataBook.close()


# Start Excel VBA macro in Python =================
dataBook = xw.Book(xlsm_path)
copyData = dataBook.macro('DataMovingPrinting')

copyData()

dataBook.save()
dataBook.close()