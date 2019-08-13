import xlsxwriter
import os

#--------------------------------------------------------------------------
# Function    : <Function Name>
# Description : <Description of the function.  Something more useful than
#                the obvious>
# Input       : <Description of each item in the signature>
# Output      : <Description of the return value>
#--------------------------------------------------------------------------

def checkDocForCoeff(coeffList, table, coeff, codeRow, workingRow, sensorColumn, sensorDict):
    for coeffIndex, coeffName in enumerate(coeffList, 0): # loop through the doc coeff titles
        if coeff == coeffName or (coeff == "ID String" and coeffName == "Product"): #If these match then this doc contains the coeff title
            for docRow in table: #loop through each row of the doc
                code = codeRow[0]
                doc = str(docRow[sensorColumn])
                
                dictValues = sensorDict.get(doc)
                listOfKeys = sensorDict.keys()
                releventKeys = []
                for key in listOfKeys:
                    if sensorDict.get(code) == doc :
                        releventKeys.append(key)

                # If the sensor in the doc matched the sensor in the code
                if code == doc or code == dictValues or code in releventKeys and docRow[sensorColumn] != "": 

                    # Exceptions for specific coeffs
                    if coeff == "TubeID":  # A=pi*r^2. r=TubeID(AP)/2 * NumberTubes (AO)
                        tubeID = float(docRow[coeffIndex])
                        numOfTubes = float(docRow[coeffList.index('NumberTubes')])
                        temp = numOfTubes * 3.14 * ((tubeID/2)**2)
                        workingRow.append(str(temp))

                    elif coeff == "NominalFlowRate":  #Make the unit conversions between the doc (kg/s) and code (uS)
                        temp = float(docRow[coeffIndex]) * 1000 / float(docRow[coeffList.index('FlowCalFactor')]) 
                        workingRow.append(str(temp))
                        
                    elif coeff == "I.D. Resistor":
                        if float(docRow[coeffIndex]) < -5: 
                            workingRow.append(str("SLOT_BAD"))
                        elif float(docRow[coeffIndex]) < 4:
                            workingRow.append(str("SLOT0"))
                        elif float(docRow[coeffIndex]) < 39.7:
                            workingRow.append(str("SLOT_UNKNOWN"))
                        elif float(docRow[coeffIndex]) < 42.2:
                            workingRow.append(str("SLOT1"))
                        elif float(docRow[coeffIndex]) < 44.3:
                            workingRow.append(str("SLOT2"))
                        elif float(docRow[coeffIndex]) < 46.4:
                            workingRow.append(str("SLOT3"))
                        elif float(docRow[coeffIndex]) < 48.7:
                            workingRow.append(str("SLOT4"))
                        elif float(docRow[coeffIndex]) < 50.5:
                            workingRow.append(str("SLOT5"))
                        elif float(docRow[coeffIndex]) < 51.7:
                            workingRow.append(str("SLOT_UNKNOWN"))
                        elif float(docRow[coeffIndex]) < 702:
                            workingRow.append(str("SLOT12"))

                    else:
                        workingRow.append(str(docRow[coeffIndex]))

                    return workingRow



def compileErDocRow(coeff, codeRow, workingRow, newBlueCoeffList, newBlueTable, oldBlueCoeffList, oldBlueTable,
    coriolisRedCoeffList, coriolisRedTable, densViscRedCoeffList, densViscRedTable, purpleDocTable, greentableOne, 
    greenTableFour, sensorCompDict):

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
    if checkDocForCoeff(coriolisRedCoeffList, coriolisRedTable, coeff, codeRow, workingRow, redSensorColumn, sensorCompDict) != None:
        return

    if checkDocForCoeff(densViscRedCoeffList, densViscRedTable, coeff, codeRow, workingRow, redSensorColumn, sensorCompDict) != None:
        return

    # Look at new blue doc before old blue doc for most recent coeff values
    if checkDocForCoeff(newBlueCoeffList, newBlueTable, coeff, codeRow, workingRow, blueSensorColumn, sensorCompDict) != None:
        return

    # Look at old sheet in blue ER doc for coeff
    if checkDocForCoeff(oldBlueCoeffList, oldBlueTable, coeff, codeRow, workingRow, blueSensorColumn, sensorCompDict) != None:
        return

    workingRow.append('―') # Else, add a '―' for coefficients that cant be populated  
    return workingRow


#Using default crapy round function
def formatString(stringVariable):
    try: #If the variable can be made into a float, return the float value
        return "%s"%round(float(stringVariable), 3)
    except: # If the variable cannot be made into a float, then just return the original string
        return stringVariable


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


#    fcf = Index of Flow Cal Factor
#    nfr = Index of Nominal Flow Rate
#    k1 = Index of K1
def compareString(codeVal, docVal, coeffIndex, coeffsToCompare, sensorDict):
    try: #try making the inputs float
        codeVal = float(codeVal)
        docVal = float(docVal)
        
        fcf = coeffsToCompare.index('FlowCalFactor')
        nfr = coeffsToCompare.index("NominalFlowRate")
        k1 = coeffsToCompare.index("K1")
        
        tolerance = codeVal * 0.001   #this value determines how lenient the number comparison is
        specialTolerance = codeVal * 0.03 #this value determines how lenient the special number comparison is
        
        isNumber = True
        if coeffIndex == fcf or coeffIndex == nfr or coeffIndex == k1:
            if ((docVal - (codeVal + specialTolerance))*(docVal - (codeVal - specialTolerance)) <= 0):  # check if 
                return True
        
        if ((docVal - (codeVal + tolerance))*(docVal - (codeVal - tolerance)) <= 0):  # check if 
            return True

    except Exception as e: # the inputs are not floats (probably strings)
        codeVal = str(codeVal).strip()
        docVal = str(docVal).strip()

        if docVal in codeVal: #This was done this direction so SLOT0+SLOT1 matches with SLOT0
            return True

        if (((codeVal == "0.0") or (codeVal == "0")) and ((docVal == "null") or (docVal == "―"))): 
            return True

        if (((docVal == "0.0") or (docVal == "0")) and ((codeVal == "null") or (codeVal == "―"))): 
            return True

        try:
            if codeVal == sensorDict.get(docVal):
                return True
        except:
            return False

    return False



def createFinalCodeAndDocArrays(coeffsToCompare, finalCodeArray, finalDocArray, mainCoeffList, coeffTable, newBlueCoeffList, newBlueTable, oldBlueCoeffList, 
    oldBlueTable, coriolisRedCoeffList, coriolisRedTable,     densViscRedCoeffList, densViscRedTable, purpleDocTable, greentableOne, greenTableFour, sensorComparisonDict):
    workingRow = []
    coeffPopulated = False

    #Put code coeffs into final code array
    for codeRow in coeffTable:  #put first row of coeffs from code in final array
        for coeff in coeffsToCompare: # Loop through master coeff list
            coeffPopulated = False
            for codeCoeffIndex, code_list_coeff_name in enumerate(mainCoeffList[0], 0):
                if coeff == code_list_coeff_name:
                    coeffPopulated = True
                    workingRow.append(codeRow[codeCoeffIndex]) # Add coeff value to list 
                    break

            if coeffPopulated == False:
                workingRow.append("―")  # Add a place holder if value doesnt exist 

        finalCodeArray.append(workingRow)
        workingRow = []


        # put er document coeffs into final array 
        for coeff in coeffsToCompare:  # Loop through master coeff list
            compileErDocRow(coeff, codeRow, workingRow, newBlueCoeffList, newBlueTable, oldBlueCoeffList, oldBlueTable,
                coriolisRedCoeffList, coriolisRedTable, densViscRedCoeffList, densViscRedTable, purpleDocTable, greentableOne,
                greenTableFour, sensorComparisonDict)
        finalDocArray.append(workingRow)
        workingRow = []



def createFinalArray(coeffsToCompare, finalArray, finalCodeArray, finalDocArray, sensorDict):
    workingRow = []
    
    finalArray.append(coeffsToCompare) # Add coeff titles to first row of final array 
    

    #Create and put match or no match row into final array
    for array_num, array in enumerate(finalCodeArray):
        finalArray.append(finalCodeArray[array_num])
        finalArray.append(finalDocArray[array_num])

        for itemNum, item in enumerate(array):
            if compareString(finalCodeArray[array_num][itemNum], finalDocArray[array_num][itemNum], itemNum, coeffsToCompare, sensorDict):
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
    filePath = os.getcwd()

    # Create an new Excel file and add a worksheet.
    workbook = xlsxwriter.Workbook(filePath+'\Comparisons_Spreadsheet.xlsx')
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