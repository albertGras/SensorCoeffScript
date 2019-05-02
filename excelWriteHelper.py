import xlsxwriter

#--------------------------------------------------------------------------
# Function    : <Function Name>
# Description : <Description of the function.  Something more useful than
#                the obvious>
# Input       : <Description of each item in the signature>
# Output      : <Description of the return value>
#--------------------------------------------------------------------------


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




def createFinalArray(finalArray, COEFFS_TO_COMPARE, coeffTable, mainCoeffList, workingRow, finalCodeArray, newBlueCoeffList,
    newBlueTable, oldBlueCoeffList, oldBlueTable, coriolisRedCoeffList, coriolisRedTable, 
    densViscRedCoeffList, densViscRedTable, purpleDocTable, finalDocArray, greentableOne, greenTableFour):

    coeffPopulated = False

    finalArray.append(COEFFS_TO_COMPARE) # Add coeff titles to first row of final array 

    #Put code coeffs into final code array
    for codeRow in coeffTable:  #put first row of coeffs from code in final array
        for coeff in COEFFS_TO_COMPARE: # Loop through master coeff list
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
        for coeff in COEFFS_TO_COMPARE:  # Loop through master coeff list
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