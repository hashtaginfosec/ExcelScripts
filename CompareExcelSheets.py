#Import XLRD module so that we can interact with Excel files
import xlrd

#Declare and instantiate both workbooks
wbV6 = xlrd.open_workbook('Workbook v6.xlsx')
wbV5 = xlrd.open_workbook('Workbook v5.xlsx', 'r')

#Declare and instantiate sheets in each workbook
sheetV6 = wbV6.sheet_by_name("Template")
sheetV5 = wbV5.sheet_by_name("Risk Assessment Gaps")

#Let us test the workbook and sheet assignment
print(sheetV6.name)
print(sheetV5.name)

#Declare Arrays that will hold Unique ID's
arrayV6UniqueID = []
arrayV5UniqueID = []

#Iterate through 134 rows and copy cell 0 (unique ID #) to arrayV6UniqueID
try:
    for i in range(1,134):
       arrayV6UniqueID.append(sheetV6.cell(i,0).value)
                
except IndexError as e:
    print(e)

#Iterate through 132 rows and copy cell 0 (Unique ID # ) to arrayV5UniqueID
try:
    for i in range(1,132):
       arrayV5UniqueID.append(sheetV5.cell(i,0).value)
                
except IndexError as e:
    print(e)

#This is two part statement.
#From left to right
#Sets are lists with no duplicate entries. Since we want to compare unique ID #s between both workbooks, we create two sets.
#To get a difference of workbooks v5 & v6, we subtract v5 set from v6 worksheet set.
#What we get from subtraction is a tuple which we plug into an array.

diffArrayV6 = tuple(set(arrayV6UniqueID) - set(arrayV5UniqueID))
print("Controls that exist in v6 version but not in V5.")
for item in diffArrayV6:
    print("Unique ID " + item)

#Same as above, only reversed to get the difference between v5 and other worksheet
diffArrayV5 = tuple(set(arrayV5UniqueID) - set(arrayV6UniqueID))
print("Controls that exist in V5 but no in v5 version.")
for item in diffArrayV5:
    print("Unique ID " + item)
