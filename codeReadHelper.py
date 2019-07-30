import os

#--------------------------------------------------------------------------
# Function    : <Function Name>
# Description : <Description of the function.  Something more useful than
#                the obvious>
# Input       : <Description of each item in the signature>
# Output      : <Description of the return value>
#--------------------------------------------------------------------------
def createConstantsArray(line):
    for item in ['#define', 'f', '\n', '(', ')']:
        line = line.replace(item, '')

    line = line.split(' ')
    line = [x for x in line if x != ''] #Remove blank items
    del line[2:20]
    return line


def coeffListSetUp(line, coeffList):
    COEFF_LIST_EXTRA = ['{', '\"', '\n', '}', ';', '/', '*', 'COEFF_TABLE,', 'CAT_TABLE,', 'SMV_TABLE,']
    for item in COEFF_LIST_EXTRA:  #remove unwanted characters in the line
        line = line.replace(item, '')

    line = line.split(',')   # Split up the code line by comas 
    line = [x for x in line if x != ''] #Remove blank items
    coeffList.append(line)


def tableSetUp(line, types):
    TABLE_EXTRA = ['{', '\"', '\n', '}', ' ',  'f', ';', '&', '}', 'const', 'S_CAT', 'S_SMV',]

    if line == '\n' or line == '':  # Remove blank lines
        return [line, types]

    for item in TABLE_EXTRA:  #remove unwanted characters in the line
        line = line.replace(item, '')
    line = line.split(',')
    
    if(isinstance(types, list)):  #copy characters before the '='
        types.append(line[0].split('=')[0]) 

    line[0] = line[0].split('=', 1)[-1]  #Remove all characters before the '='
    return [line, types]


def copyCodeFile(codeFile, sensorList, mainCoeffList, coeffTable, catCoeffList, catTable, smvCoeffList, smvTable, constantsTable, catTypes, smvTypes):
#    codeFile   = open("C:\PVCS\ProjectsDB\Kinetis_DB\k2Src\k_src_app\coriolis\sensor.cpp")

    codeFile = open(os.path.join(codeFile, "sensor.cpp"))
    inCoeffTable = False
    inSmvTable = False
    inCatTable = False

    workingRow = []
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
                if line[0] == "CNG/F50P":
                    line[0] = "CNG/F050P"   # Correction: to handle the missing 0 in the code
                sensorList.append(line[0])  # Take the sensor types from each line
                coeffTable.append(line)

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

def removeObsoleteSensors(coeffTable, obsoleteSensors):
    lengthOfCoeffTable = len(coeffTable)-1
    for rowNum, row in enumerate(reversed(coeffTable), 0):
        for sensor in obsoleteSensors:
            if sensor == row[0]:
                del coeffTable[lengthOfCoeffTable-rowNum]
                break
    return coeffTable



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

            if temp != None: line.append(temp)
            flowLinearityTable.append(line)
    




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
                            row.insert(itemNum + 1, catTable[typeNum][catIndex])  # insert cat values into coeff table
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
                            row.insert(itemNum + 1, smvTable[typeNum][smvIndex])  # insert smv values into coeff table
                        break



def replaceVariablesWithDefinitions(coeffTable, constantsTable):
    for codeRowNum, codeRow in enumerate(coeffTable, 0):
        for codeItemNum, codeItem in enumerate(codeRow, 0):
            for x in constantsTable:
                if x[0] == codeItem:
                    coeffTable[codeRowNum][codeItemNum] = x[1]
                    break
    return coeffTable









