import xlrd
import collections 

#--------------------------------------------------------------------------
# Function    : <Function Name>
# Description : <Description of the function.  Something more useful than
#                the obvious>
# Input       : <Description of each item in the signature>
# Output      : <Description of the return value>
#--------------------------------------------------------------------------
def copyExcelFile(coeffListOne, tableOne, coeffListTwo, tableTwo, file, coeffLineNum, dataLineNum):
    wb = xlrd.open_workbook(file) #Create workbook for er document
    sheetOne = wb.sheet_by_index(1) #Set up the old/coriolis parameters sheet in the excel document
    sheetTwo = wb.sheet_by_index(2) #Set up the new/density viscocity parameters sheet in the excel document
    
    newRow = []

    # Copy all coefficient names in  er doc - first sheet
    numCols = sheetOne.ncols   # Number of columns
    for colIndex in range(0, numCols):  # Iterate through columns
        coeffListOne.append(sheetOne.cell_value(coeffLineNum, colIndex))  # Get cell object by row, col

    # Copy all values in er doc - first sheet
    for rowIndex in range(dataLineNum, sheetOne.nrows):    # Iterate through rows
        for colIndex in range(0, numCols):  # Iterate through columns
            temp = str(sheetOne.cell_value(rowIndex, colIndex))
            temp = temp.replace(' ', '')
            newRow.append(temp)  # Get cell object by row, col
        tableOne.append(newRow)
        newRow = []

    # Copy all coefficient names in er doc - second sheet 
    numCols = sheetTwo.ncols   # Number of columns
    for colIndex in range(0, numCols):  # Iterate through columns (x axis direction)
        coeffListTwo.append(sheetTwo.cell_value(coeffLineNum, colIndex))  # Get cell object by row, col

    # Copy all values in er doc - first sheet  
    for rowIndex in range(dataLineNum, sheetTwo.nrows):    # Iterate through rows (y axis direction)
        for colIndex in range(0, numCols):  # Iterate through columns
            newRow.append(sheetTwo.cell_value(rowIndex, colIndex))  # Get cell object by row, col
        tableTwo.append(newRow)
        newRow = []




def copyGreenErDoc(greenFile, greentableOne, greenTableTwo, greenTableThree, greenTableFour):
    newRow = []
    coeffRow = False
    beforeTables = True
    tableNumber = 0
    
    newSensor = False
    tempSensor = None
    
    tables = greenFile.tables
    
    for table in tables:
        for row in table.rows:
            coeffRow = False
            if(beforeTables == False and len(newRow) != 0):

                if [s for s in newRow if '0.95' in s]:
                    sensor = newRow[0:2]
                else:
                    sensor = newRow[0]
                
                if tempSensor == sensor:
                    newSensor = False
                else:
                    newSensor = True
                    tempSensor = sensor
               
                if tableNumber == 1:
                    greentableOne.append(newRow)

                elif tableNumber == 2:
                    if newSensor == True:
                        newSensor = False
                        greenTableTwo.append(sensor)
                    greenTableTwo.append(newRow[2:4])

                elif tableNumber == 3:
                    if newSensor == True:
                        newSensor = False
                        greenTableTwo.append([sensor])
#                        greenTableThree.append(sensor)
#                    greenTableThree.append(newRow[1:3])
                    greenTableTwo.append(newRow[1:3])

                elif tableNumber == 5: # because of the way the table is setup Sensor Model is read twice and this becomes 5 instead of 4
                    greenTableFour.append(newRow)

                newRow = []
            for cell in row.cells:
                if("Sensor Model") in cell.text:
                    tableNumber = tableNumber + 1
                    coeffRow = True
                    beforeTables = False

                if beforeTables == False and coeffRow == False:  # Add item to row if its after the tables start
#                    newRow = [x for x in newRow if x != '< 0.95'] #Remove blank string items
                    if cell.text != '':  # Remove blank lines
                        newRow.append(cell.text)



def copyPurpleErDoc(purpleFile):
    purpleDocTable = []
    beforeTables = True
    coeffRow = False
    newRow = []

    #Put green ER-20015206_AP docx into list (original)
    tables = purpleFile.tables
    for table in tables:
        for row in table.rows:
            coeffRow = False
            if(beforeTables == False and len(newRow) != 0):
                purpleDocTable.append(newRow)
#                print(newRow)
                newRow = []

            for cell in row.cells:
                if("Sensor Model") in cell.text:
                    coeffRow = True
                    beforeTables = False
                if beforeTables == False and coeffRow == False:
                    if cell.text != '':  # Remove blank lines
                        newRow.append(cell.text)
    purpleDocTable.append(newRow)
    return purpleDocTable



def createSensorComparisonDict(sensorComparisonDict, newBlueCoeffList, newBlueTable, oldBlueCoeffList, oldBlueTable):
    baseModelNum = 0
    IdStringNum = 0

    for titleNum, coeff in enumerate(newBlueCoeffList, 0):
        if "Base Model" in str(coeff):
            baseModelNum = titleNum
        elif "ID String" in str(coeff):
            IdStringNum = titleNum
    
    for row in newBlueTable:
        if row[baseModelNum] not in sensorComparisonDict:
            sensorComparisonDict[row[baseModelNum]] = row[IdStringNum]

    for titleNum, coeff in enumerate(oldBlueCoeffList, 0):
        if "Base Model" in str(coeff):
            baseModelNum = titleNum
        elif "ID String" in str(coeff):
            IdStringNum = titleNum
            
    for row in oldBlueTable:
        if row[baseModelNum] not in sensorComparisonDict:
            sensorComparisonDict[row[baseModelNum]] = row[IdStringNum]
            
    sensorComparisonDict = collections.defaultdict(lambda : 'None')
#    print()
#    print(sensorComparisonDict)
#    print()













