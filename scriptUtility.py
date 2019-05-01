import xlrd
import xlsxwriter
#import math
from math import log10, floor


#--------------------------------------------------------------------------
# Function    : <Function Name>
# Description : <Description of the function.  Something more useful than
#                the obvious>
# Input       : <Description of each item in the signature>
# Output      : <Description of the return value>
#--------------------------------------------------------------------------
def createConstantsArray(line):
    line = line.replace('#define', '')
    line = line.replace('f', '')
    line = line.replace('\n','')
    line = line.split(' ')
    line = [x for x in line if x != ''] #Remove blank items
    del line[2:20]
    return line


def coeffListSetUp(line, coeffList):

    extraCharacters2 = ['{', '\"', '\n', '}', ';', '/', '*', 'COEFF_TABLE,', 'CAT_TABLE,', 'SMV_TABLE,']

    for item in extraCharacters2:  #remove unwanted characters in the line
        line = line.replace(item, '')

    line = line.split(',')   # Split up the code line by comas 
    line = [x for x in line if x != ''] #Remove blank items
    coeffList.append(line)
    
    
    
    
def tableSetUp(line, types):
#    print("pre %s" % repr(line))
    if line == '\n' or line == '':  # Remove blank lines
#        print('exit')
        return [line, types]
    extraCharacters = ['{', ' ', '\"', '\n', 'f','}', ';', '&', '}', 'const', 'S_CAT', 'S_SMV',]
#    extraCharacters = ['{', ' ', '\"', '\n', 'f','}', ';', '/', '&', '}', 'const', 'S_CAT', 'S_SMV',]
    
    for item in extraCharacters:  #remove unwanted characters in the line
        line = line.replace(item, '')
    line = line.split(',')
    
    if(isinstance(types, list)):
        types.append(line[0].split('=')[0]) #copy characters before the '='

    line[0] = line[0].split('=', 1)[-1]  #Remove all characters before the '='
#    print("end %s" % line)
#    print()
    return [line, types]




def copyCodeFile(sensorList, mainCoeffList, coeffTable, catCoeffList, catTable, smvCoeffList, smvTable, constantsTable, catTypes, smvTypes):
#    codeFile   = open("H:\SensorScript\practice.h")
    codeFile   = open("C:\PVCS\ProjectsDB\Kinetis_DB\k2Src\k_src_app\coriolis\sensor.cpp")
#    codeFile   = open("C:\Pegasus\BFSrc800\Coriolis\sensor.cpp")

    inCoeffTable = False
    inSmvTable = False
    inCatTable = False

    line = []
    count = 0
    
    codeFile.seek(0)   # Go to beginning of file

    while count < 3:   # while all 3 tables havent been read (Coeff, Cat, Smv)
        line = codeFile.readline()  # Read a line from the coefficient code file

        if not line:                  # if end of file, stop trying to read the file
            break  # break out of reading file while loop

        if "define" in line:
            line = createConstantsArray(line)
            constantsTable.append(line)
#            print(constantsTable)
#            print()

        if "End" in line:   # end of one of the tables
            count = count + 1
            inCoeffTable = False
            inSmvTable = False
            inCatTable = False

        elif "COEFF_TABLE" in line:      #Create coeff table list
            coeffListSetUp(line, mainCoeffList) # Prepare title line to be put into array
            inCoeffTable = True

        elif inCoeffTable == True:       #Create coeff table
            line, _ = tableSetUp(line, 0)# Prepare coeff data lines to be put into array 
            if line[0] == '' or line[0] == '\n':  # Remove blank lines
                continue
            if not any("--" in s for s in line): #Remove comment rows with additional coeff decriptions
                sensorList.append(line[0])  # Take the sensor types from each line
                coeffTable.append(line)
#                print(line)

        elif "CAT_TABLE" in line:        #Create cat table list
            coeffListSetUp(line, catCoeffList) # Prepare title line to be put into array
            inCatTable = True

        elif inCatTable == True:         #Create cat table
            line, catTypes = tableSetUp(line, catTypes) # Prepare coeff data lines to be put into array 
            if line[0] == '' or line[0] == '\n':  # Remove blank lines
                continue
            catTable.append(line)

        elif "SMV_TABLE" in line:        #Create smv table list
            coeffListSetUp(line, smvCoeffList) # Prepare title line to be put into array
            inSmvTable = True

        elif inSmvTable == True:         #Create smv table
            line, smvTypes =tableSetUp(line, smvTypes) # Prepare coeff data lines to be put into array 
            if line[0] == '' or line[0] == '\n':  # Remove blank lines
                continue
            smvTable.append(line)


            
# Find Flow linearity coeffs
def findFlowLinearityCoeffs(flowLinearityTable):
    codeFile = open("C:\PVCS\ProjectsDB\Kinetis_DB\k2Src\k_src_app\coriolis\massflow.cpp")
    
    insideTable = False
    newRow = []
    codeFile.seek(0)   # Go to beginning of file
    
    while(1):
        line = 0
        line = codeFile.readline()  # Read a line from the coefficient code file

        if not line:                  # if end of file, stop trying to read the file
            break  # break out of reading file while loop

        if "FlowLinearityTables" in line:
            insideTable = True
            
        elif "End" in line:
            break # break out of reading file while loop
            
        elif insideTable == True:
            for item in ['float', 'kg/s', 'comp', 'FL_', '{', ' ', '[]', '\n','}', ';', '}',  'f']:  #remove unwanted characters in the line
                line = line.replace(item, '')
            
            if 'float' in line:
                sensor_line = True
            else: 
                sensor_line = False
            
            if 'LT' in line:
                temp = '< 0.95'
            elif 'GT' in line:
                temp = '> 0.95'
            else:
                temp = None
            
            line = line.split('_',1)[0]  #Remove all characters after the '_'
            line = line.split('/',1)[0]  #Remove all characters after the '/'
            line = line.split('=',1)[0]  #Remove all characters after the '='
            
            if line == '\n' or line == '':  # Remove blank lines
                continue

            line = line.split(',')
            line = [x for x in line if x != ''] #Remove blank string items

#            print(repr(line))
            if temp != None: line.append(temp)
            flowLinearityTable.append(line)
    
#    print(flowLinearityTable)




def addCatAndSmvTablesToCoeffTable(mainCoeffList, coeffTable, catCoeffList, catTable, smvCoeffList, smvTable, catTypes, smvTypes):
    mainCoeffList = mainCoeffList[0]

    for coeffIndex, coeff in enumerate(mainCoeffList, 0):
        if "CAT" in coeff:
            for catCoeff in catCoeffList[0]:
                mainCoeffList.insert(coeffIndex + 1, catCoeff)      # Insert cat titles in coeff title list
            break
    
    for row in coeffTable:
        for itemNum, item in enumerate(row,0):
            if 'CAT' in item:
                if item not in catTypes:
                    for catIndex in range(0, len(catTable[0])):
                        row.insert(itemNum + 1, 'bad cat')
                    break
                for typeNum, type in enumerate(catTypes, 0):
                    if item == type:
                        for catIndex in range(0, len(catTable[0])):
                            row.insert(itemNum + 1, catTable[typeNum][catIndex])
                        break

    for coeffIndex, coeff in enumerate(mainCoeffList, 0):
        if "SMV" in coeff:
            for smv_title in smvCoeffList[0]:
                mainCoeffList.insert(coeffIndex + 1, smv_title)      # Insert smv titles in coeff title list
            break


    for row in coeffTable:
        for itemNum, item in enumerate(row,0):
            if 'SMV' in item:
                if item not in smvTypes:
                    for smvIndex in range(0, len(smvTable[0])):
                        row.insert(itemNum + 1, 'bad smv')
                    break
                for typeNum, type in enumerate(smvTypes, 0):
                    if item == type:
                        for smvIndex in range(0, len(smvTable[0])):
                            row.insert(itemNum + 1, smvTable[typeNum][smvIndex])
                        break



def replaceVariablesWithDefinitions(coeffTable, constantsTable):
    for codeRowNum, codeRow in enumerate(coeffTable, 0):
        for codeItemNum, codeItem in enumerate(codeRow, 0):
            for x in constantsTable:
                if x[0] == codeItem:
                    coeffTable[codeRowNum][codeItemNum] = x[1]
                    break
    return coeffTable



def copyExcelFile(coeffListOne, tableOne, coeffListTwo, tableTwo, file, coeffLineNum, dataLineNum):
    wb = xlrd.open_workbook(file) #Create workbook for excel blue er document
    sheetOne = wb.sheet_by_index(1) #Set up the old parameters sheet in the excel document
    sheetTwo = wb.sheet_by_index(2) #Set up the new parameters sheet in the excel document
    
    newRow = []

    # Copy all coefficient names in blue er doc - old params 
    numCols = sheetOne.ncols   # Number of columns
    for colIndex in range(0, numCols):  # Iterate through columns
        coeffListOne.append(sheetOne.cell_value(coeffLineNum, colIndex))  # Get cell object by row, col

    # Copy all values in blue er doc - old params 
    for rowIndex in range(dataLineNum, sheetOne.nrows):    # Iterate through rows
        for colIndex in range(0, numCols):  # Iterate through columns
            temp = str(sheetOne.cell_value(rowIndex, colIndex))
            temp = temp.replace(' ', '')
#            temp = temp.replace('/', '')  #Some sensor types have the slash
#            newRow.append(sheetOne.cell_value(rowIndex, colIndex))  # Get cell object by row, col
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



def checkDocForCoeff(coeffList, table, coeff, codeRow, workingRow, sensorColumn):
    for coeffIndex, coeffName in enumerate(coeffList, 0): # loop through new sheet doc coeff titles
        if coeff == coeffName: #If these match then this doc contains the coeff title 
            for row in table: #loop through each row of the doc
                if codeRow[0] in str(row[sensorColumn]): # If the sensor in the doc matched the sensor in the code

                    # Exceptions for specific coeffs
                    if coeff == "TubeID":  # A=pi*r^2. r=TubeID(AP)/2 * NumberTubes (AO) 
                        temp = 3.14 * ((float(row[coeffIndex])*float(row[coeffList.index('NumberTubes')])/2)**2)
                        workingRow.append(str(temp))

                    elif coeff == "NominalFlowRate":  #Make the unit conversions between the doc (kg/s) and code (uS)
                        temp = float(row[coeffIndex]) * 1000 / float(row[coeffList.index('FlowCalFactor')]) 
                        workingRow.append(str(temp))
                        
                    elif coeff == "I.D. Resistor":
                        if float(row[coeffIndex]) < -5: 
                            workingRow.append(str("SLOT_BAD"))
                        elif float(row[coeffIndex]) < 4:
                            workingRow.append(str("SLOT0"))
                        elif float(row[coeffIndex]) < 39.7:
                            workingRow.append(str("SLOT_UNKNOWN"))
                        elif float(row[coeffIndex]) < 42.2:
                            workingRow.append(str("SLOT1"))
                        elif float(row[coeffIndex]) < 44.3:
                            workingRow.append(str("SLOT2"))
                        elif float(row[coeffIndex]) < 46.4:
                            workingRow.append(str("SLOT3"))
                        elif float(row[coeffIndex]) < 48.7:
                            workingRow.append(str("SLOT4"))
                        elif float(row[coeffIndex]) < 50.5:
                            workingRow.append(str("SLOT5"))
                        elif float(row[coeffIndex]) < 51.7:
                            workingRow.append(str("SLOT_UNKNOWN"))
                        elif float(row[coeffIndex]) < 702:
                            workingRow.append(str("SLOT12"))

#                    elif coeff == "


                    else:
                        workingRow.append(str(row[coeffIndex]))
                    return workingRow



def compileErDocRow(coeff, codeRow, workingRow, newBlueCoeffList, newBlueTable, oldBlueCoeffList, oldBlueTable,
    coriolisRedCoeffList, coriolisRedTable, densViscRedCoeffList, densViscRedTable, purpleDocTable, greentableOne, 
    greenTableFour):
    
    blueSensorColumn = 3
    redSensorColumn = 1
    
    # Look at new sheet in blue ER doc for coeff 
    if coeff == "A 4":  # There is a space because there is an unrelated A4 in the red document 
        for row in purpleDocTable:
            if codeRow[0] in row[0]:
                workingRow.append(str(row[1]))
                return

    if coeff == "GasFD":
        for row in greenTableFour:
            if codeRow[0] in row[0]:
                workingRow.append(str(row[1]))
                return

    if coeff == "Gas Slope":
        for row in greenTableFour:
            if codeRow[0] in row[0]:
                workingRow.append(str(row[2]))
                return

    if coeff == "Gas Offset":
        for row in greenTableFour:
            if codeRow[0] in row[0]:
                workingRow.append(str(row[3]))
                return

    if coeff == "Liq Slope":
        for row in greenTableFour:
            if codeRow[0] in row[0]:
                workingRow.append(str(row[4]))
                return

    if coeff == "Liq Offset":
        for row in greenTableFour:
            if codeRow[0] in row[0]:
                workingRow.append(str(row[5]))
                return
               
    if coeff == "T03":
        for row in greentableOne:
            if codeRow[0] in row[0]:
                workingRow.append(str(row[1]))
                return
                
          
     # Look at red doc before blue doc for most recent coeff values
    if checkDocForCoeff(coriolisRedCoeffList, coriolisRedTable, coeff, codeRow, workingRow, redSensorColumn) != None:
        return

    if checkDocForCoeff(densViscRedCoeffList, densViscRedTable, coeff, codeRow, workingRow, redSensorColumn) != None:
        return
          
    # Look at new blue doc before old blue doc for most recent coeff values
    if checkDocForCoeff(newBlueCoeffList, newBlueTable, coeff, codeRow, workingRow, blueSensorColumn) != None:
        return

    # Look at old sheet in blue ER doc for coeff
    if checkDocForCoeff(oldBlueCoeffList, oldBlueTable, coeff, codeRow, workingRow, blueSensorColumn) != None:
        return

    workingRow.append('0') # Else, add a '―' for coefficients that cant be populated
    return workingRow


#Using default crapy round function
def formatString(stringVariable):
    try: #If the variable can be made into a float, return the float value
        return "%s"%round(float(stringVariable), 3)
    except: # If the variable cannot be made into a float, then just return the original string
        return stringVariable



def compareString(codeVal, docVal):
    try: #try making the inputs float
        codeVal = float(codeVal)
        docVal = float(docVal)
        tolerance = codeVal * 0.001   #this value determines how lenient the number comparison is
        isNumber = True
    except Exception as e: #if this falls through, then the inputs are strings
        isNumber = False

    if isNumber == True: # if the inputs are numbers, see if they are close enough to match
        if ((docVal - (codeVal + tolerance))*(docVal - (codeVal - tolerance)) <= 0):  # check if 
            return True
    
#    elif coeff == "NominalFlowRate":
    
#    elif coeff == "flags":
#    if codeVal == "50" and "SSA_DRIVESAT" in docVal:
#        return True

    else: #if the inputs arent numbers, see if the strings match
#        if codeVal == docVal:
        if str(docVal) in str(codeVal):
            return True

    return False



def createFinalArray(finalArray, coeffs_to_compare, coeffTable, mainCoeffList, workingRow, finalCodeArray, newBlueCoeffList,
    newBlueTable, oldBlueCoeffList, oldBlueTable, coriolisRedCoeffList, coriolisRedTable, 
    densViscRedCoeffList, densViscRedTable, purpleDocTable, finalDocArray, greentableOne, greenTableFour):

    coeffPopulated = False

    finalArray.append(coeffs_to_compare) # Add coeff titles to first row of final array 

    #Put code coeffs into final code array
    for codeRow in coeffTable:  #put first row of coeffs from code in final array
        for coeff in coeffs_to_compare: # Loop through master coeff list
            coeffPopulated = False
            for code_coeffIndex, code_list_coeff_name in enumerate(mainCoeffList[0], 0):
                if coeff == code_list_coeff_name:
                    coeffPopulated = True
                    workingRow.append(codeRow[code_coeffIndex]) # Add coeff value to list 
                    break

            if coeffPopulated == False:
                workingRow.append("0")  # Add a place holder if value doesnt exist

        finalCodeArray.append(workingRow)
        workingRow = []


        # put er document coeffs into final array 
        for coeff in coeffs_to_compare:  # Loop through master coeff list
            compileErDocRow(coeff, codeRow, workingRow, newBlueCoeffList, newBlueTable, oldBlueCoeffList, oldBlueTable,
                coriolisRedCoeffList, coriolisRedTable, densViscRedCoeffList, densViscRedTable, purpleDocTable, greentableOne,
                greenTableFour)
        finalDocArray.append(workingRow)
        workingRow = []
   
   
   
   #Create and put match or no match row into final array
    for array_num, array in enumerate(finalCodeArray):
        finalArray.append(finalCodeArray[array_num])
        finalArray.append(finalDocArray[array_num])

        for itemNum, item in enumerate(array):

            if compareString(finalCodeArray[array_num][itemNum], finalDocArray[array_num][itemNum]):
                workingRow.append("Ok")

            # TODO move this to the compare section 
            elif ((finalCodeArray[array_num][itemNum] == "0.0") or (finalCodeArray[array_num][itemNum] == "0")) and ((finalDocArray[array_num][itemNum] == 'null') or (finalDocArray[array_num][itemNum] == '―')): 
#            elif ((finalCodeArray[array_num][itemNum] == "0.0") or (finalCodeArray[array_num][itemNum] == "0")) and ((finalDocArray[array_num][itemNum] == 'null')):
                workingRow.append("Ok") 

#            elif ((finalDocArray[array_num][itemNum] == "0.0") or (finalDocArray[array_num][itemNum] == "0")) and ((finalCodeArray[array_num][itemNum] == 'null') or (finalCodeArray[array_num][itemNum] == '―')):
            elif ((finalDocArray[array_num][itemNum] == "0.0") or (finalDocArray[array_num][itemNum] == "0")) and ((finalCodeArray[array_num][itemNum] == 'null')): 
                workingRow.append("Ok")

            else:
                workingRow.append("No Match")

        finalArray.append(workingRow)
        workingRow = []
   

def writeRowDescriptors(worksheet, workbook, rowNum):
    rightBorderFormat = workbook.add_format({'right': 1})
    rightAndBottomBorderFormat = workbook.add_format({'right': 1, 'bottom': 1})
    
    if (rowNum == 0): 
        pass
    elif(rowNum % 3.0 == 1.0): 
        worksheet.write(rowNum, 0, "code", rightBorderFormat)
    elif (rowNum % 3.0 == 2.0): 
        worksheet.write(rowNum, 0, "doc", rightBorderFormat)
    elif(rowNum % 3.0 == 0.0): 
        worksheet.write(rowNum, 0, "compare", rightAndBottomBorderFormat)
   
  
def exportFinalArraytoExcelDocument(finalArray, flowLinearityTable, greenTableTwo, flowLinearityMatch):
    #Transfer final array values to excel document 
    
    finalFlowArray = []

    # Create an new Excel file and add a worksheet.
    workbook = xlsxwriter.Workbook('H:\SensorScript\Comparisons_Spreadsheet.xlsx')
#    workbook = xlsxwriter.Workbook(r'C:\Users\AGrasmeder\Documents\SensorCoeffScript')
    worksheet = workbook.add_worksheet('coeff compare')
    flowLinearity = workbook.add_worksheet('flow linearity')

    # Widen the first colums to make the text clearer.
    worksheet.set_column('B:AZ', 19)
    flowLinearity.set_column('A:AZ', 14)

    defaultFormat = workbook.add_format()   
    underlineFormat = workbook.add_format({'bottom': 1})
    rightBorderFormat = workbook.add_format({'right': 1})
    rightAndBottomBorderFormat = workbook.add_format({'right': 1, 'bottom': 1})
    redFormat = workbook.add_format({'font_color': 'red'})

    # write standard coeff compare worksheet
    for rowNum, row in enumerate(finalArray, 0):
        writeRowDescriptors(worksheet, workbook, rowNum)
        for itemNum, item in enumerate(row, 0):
            if rowNum % 3.0 == 0:
                worksheet.write(rowNum, itemNum+1, finalArray[rowNum][itemNum], underlineFormat)
            else:
                worksheet.write(rowNum, itemNum+1, finalArray[rowNum][itemNum], defaultFormat)


    # write flow linarization table in second sheet
    for rowNum, row in enumerate(flowLinearityTable, 0):
#        writeRowDescriptors(flowLinearity, workbook, rowNum)
        for itemNum, item in enumerate(row, 0):
            if itemNum == 1:
                flowLinearity.write(rowNum, itemNum, flowLinearityTable[rowNum][itemNum], rightBorderFormat)
            else:
                flowLinearity.write(rowNum, itemNum, flowLinearityTable[rowNum][itemNum])

    for rowNum, row in enumerate(greenTableTwo, 0):
#        writeRowDescriptors(flowLinearity, workbook, rowNum)
        for itemNum, item in enumerate(row, 0):
            if itemNum == 1:
                flowLinearity.write(rowNum, itemNum +2, greenTableTwo[rowNum][itemNum], rightBorderFormat)
            else:
                flowLinearity.write(rowNum, itemNum +2, greenTableTwo[rowNum][itemNum])
                
#        for itemNum, item in enumerate(flowLinearityMatch, 0):
#                flowLinearity.write(itemNum, 4, flowLinearityMatch[itemNum])

       
                
    worksheet.conditional_format('A1:ZZ500', 
                                 {'type': 'text',
                                  'criteria': 'containing',
                                  'value': 'No Match',
                                  'format': redFormat})
                                  
    flowLinearity.conditional_format('A1:ZZ500', 
                                 {'type': 'text',
                                  'criteria': 'containing',
                                  'value': 'No Match',
                                  'format': redFormat})

    workbook.close()
   
   
   
   
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

       
       
   
   
def compareflowLinearityTables(flowLinearityTable, greenTableTwo):
    x = 0
    y = 0 
    flowLinearityMatch = []
    for x in range(len(flowLinearityTable)):
        while y < len(greenTableTwo) and x < len(flowLinearityTable):
            try:   #try comparing 2 values for each table
                if formatString(flowLinearityTable[x][0]) in formatString(greenTableTwo[y][0]) and formatString(flowLinearityTable[x][1]) in formatString(greenTableTwo[y][1]):   #If theres a match
                    flowLinearityMatch.append('ok')
                    y = 0
                    break
            except:  # if both tables didnt have 2 entries then do this
                if formatString(flowLinearityTable[x][0]) in formatString(greenTableTwo[y][0]):
                    flowLinearityMatch.append('ok')
                    y = 0
                    break
            if y == len(greenTableTwo) - 1: # If at the end of the greenTable # If there isnt a match
                flowLinearityMatch.append('no match')
                y = 0
                break
            else:
                y = y + 1

    return flowLinearityMatch

   
   

   
   
   
   
   
   
   
   
   
   
   
   
   
   
   
   
   
   
   
   