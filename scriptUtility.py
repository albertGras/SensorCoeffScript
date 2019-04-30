import xlrd
import xlsxwriter
#import math
from math import log10, floor

def createConstantsArray(line):
    line = line.replace('#define', '')
    line = line.replace('f', '')
    line = line.replace('\n','')
    line = line.split(' ')
    line = [x for x in line if x != ''] #Remove blank elements
    del line[2:20]
    return line


def titleListSetUp(line, title_list):

    extraCharacters2 = ['{', '\"', '\n', '}', ';', '/', '*', 'COEFF_TABLE,', 'CAT_TABLE,', 'SMV_TABLE,']

    for item in extraCharacters2:  #remove unwanted characters in the line
        line = line.replace(item, '')

    line = line.split(',')   # Split up the code line by comas 
    line = [x for x in line if x != ''] #Remove blank elements
    title_list.append(line)
    
    
    
    
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




def copyCodeFile(sensor_list, coeff_title_list, coeff_table, cat_title_list, cat_table, smv_title_list, smv_table, constants_table, cat_types, smv_types):
#    code_file   = open("H:\SensorScript\practice.h")
    code_file   = open("C:\PVCS\ProjectsDB\Kinetis_DB\k2Src\k_src_app\coriolis\sensor.cpp")
#    code_file   = open("C:\Pegasus\BFSrc800\Coriolis\sensor.cpp")

    inCoeffTable = False
    inSmvTable = False
    inCatTable = False

    line = []
    count = 0
    
    code_file.seek(0)   # Go to beginning of file

    while count < 3:   # while all 3 tables havent been read (Coeff, Cat, Smv)
        line = code_file.readline()  # Read a line from the coefficient code file

        if not line:                  # if end of file, stop trying to read the file
            break  # break out of reading file while loop

        if "define" in line:
            line = createConstantsArray(line)
            constants_table.append(line)
#            print(constants_table)
#            print()

        if "End" in line:   # end of one of the tables
            count = count + 1
            inCoeffTable = False
            inSmvTable = False
            inCatTable = False

        elif "COEFF_TABLE" in line:      #Create coeff table list
            titleListSetUp(line, coeff_title_list) # Prepare title line to be put into array
            inCoeffTable = True

        elif inCoeffTable == True:       #Create coeff table
            line, _ = tableSetUp(line, 0)# Prepare coeff data lines to be put into array 
            if line[0] == '' or line[0] == '\n':  # Remove blank lines
                continue
            if not any("--" in s for s in line): #Remove comment rows with additional coeff decriptions
                sensor_list.append(line[0])  # Take the sensor types from each line
                coeff_table.append(line)
#                print(line)

        elif "CAT_TABLE" in line:        #Create cat table list
            titleListSetUp(line, cat_title_list) # Prepare title line to be put into array
            inCatTable = True

        elif inCatTable == True:         #Create cat table
            line, cat_types = tableSetUp(line, cat_types) # Prepare coeff data lines to be put into array 
            if line[0] == '' or line[0] == '\n':  # Remove blank lines
                continue
            cat_table.append(line)

        elif "SMV_TABLE" in line:        #Create smv table list
            titleListSetUp(line, smv_title_list) # Prepare title line to be put into array
            inSmvTable = True

        elif inSmvTable == True:         #Create smv table
            line, smv_types =tableSetUp(line, smv_types) # Prepare coeff data lines to be put into array 
            if line[0] == '' or line[0] == '\n':  # Remove blank lines
                continue
            smv_table.append(line)


            
# Find Flow linearity coeffs
def findFlowLinearityCoeffs(flowLinearityTable):
    code_file = open("C:\PVCS\ProjectsDB\Kinetis_DB\k2Src\k_src_app\coriolis\massflow.cpp")
    
    insideTable = False
    new_row = []
    code_file.seek(0)   # Go to beginning of file
    
    while(1):
        line = 0
        line = code_file.readline()  # Read a line from the coefficient code file

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
            line = [x for x in line if x != ''] #Remove blank string elements

#            print(repr(line))
            if temp != None: line.append(temp)
            flowLinearityTable.append(line)
    
#    print(flowLinearityTable)




def addCatAndSmvTablesToCoeffTable(coeff_title_list, coeff_table, cat_title_list, cat_table, smv_title_list, smv_table, cat_types, smv_types):
    coeff_title_list = coeff_title_list[0]

    for coeff_num, coeff_title in enumerate(coeff_title_list, 0):
        if "CAT" in coeff_title:
            for cat_title in cat_title_list[0]:
                coeff_title_list.insert(coeff_num + 1, cat_title)      # Insert cat titles in coeff title list
            break
    
    for row in coeff_table:
        for element_num, element in enumerate(row,0):
            if 'CAT' in element:
                if element not in cat_types:
                    for index in range(0, len(cat_table[0])):
                        row.insert(element_num + 1, 'bad cat')
                    break
                for type_num, type in enumerate(cat_types, 0):
                    if element == type:
                        for index in range(0, len(cat_table[0])):
                            row.insert(element_num + 1, cat_table[type_num][index])
                        break

    for coeff_num, coeff_title in enumerate(coeff_title_list, 0):
        if "SMV" in coeff_title:
            for smv_title in smv_title_list[0]:
                coeff_title_list.insert(coeff_num + 1, smv_title)      # Insert smv titles in coeff title list
            break


    for row in coeff_table:
        for element_num, element in enumerate(row,0):
            if 'SMV' in element:
                if element not in smv_types:
                    for index in range(0, len(smv_table[0])):
                        row.insert(element_num + 1, 'bad smv')
                    break
                for type_num, type in enumerate(smv_types, 0):
                    if element == type:
                        for index in range(0, len(smv_table[0])):
                            row.insert(element_num + 1, smv_table[type_num][index])
                        break



def replaceVariablesWithDefinitions(coeff_table, constants_table):
    for code_row_num, code_row in enumerate(coeff_table, 0):
        for code_element_num, code_element in enumerate(code_row, 0):
            for x in constants_table:
                if x[0] == code_element:
                    coeff_table[code_row_num][code_element_num] = x[1]
                    break
    return coeff_table



def copyExcelFile(excel_title_list_one, excel_table_one, excel_title_list_two, excel_table_two, excel_file, title_line_num, data_line_num):
    wb = xlrd.open_workbook(excel_file) #Create workbook for excel blue er document
    sheet_one = wb.sheet_by_index(1) #Set up the old parameters sheet in the excel document
    sheet_two = wb.sheet_by_index(2) #Set up the new parameters sheet in the excel document
    
    new_row = []

    # Copy all coefficient names in blue er doc - old params 
    num_cols = sheet_one.ncols   # Number of columns
    for col_idx in range(0, num_cols):  # Iterate through columns
        excel_title_list_one.append(sheet_one.cell_value(title_line_num, col_idx))  # Get cell object by row, col

    # Copy all values in blue er doc - old params 
    for row_idx in range(data_line_num, sheet_one.nrows):    # Iterate through rows
        for col_idx in range(0, num_cols):  # Iterate through columns
            temp = str(sheet_one.cell_value(row_idx, col_idx))
            temp = temp.replace(' ', '')
#            temp = temp.replace('/', '')  #Some sensor types have the slash
#            new_row.append(sheet_one.cell_value(row_idx, col_idx))  # Get cell object by row, col
            new_row.append(temp)  # Get cell object by row, col
        excel_table_one.append(new_row)
        new_row = []

    # Copy all coefficient names in er doc - second sheet 
    num_cols = sheet_two.ncols   # Number of columns
    for col_idx in range(0, num_cols):  # Iterate through columns (x axis direction)
        excel_title_list_two.append(sheet_two.cell_value(title_line_num, col_idx))  # Get cell object by row, col

    # Copy all values in er doc - first sheet  
    for row_idx in range(data_line_num, sheet_two.nrows):    # Iterate through rows (y axis direction)
        for col_idx in range(0, num_cols):  # Iterate through columns
            new_row.append(sheet_two.cell_value(row_idx, col_idx))  # Get cell object by row, col
        excel_table_two.append(new_row)
        new_row = []



def checkDocForCoeff(doc_title_list, doc_table, coeff, code_row, working_row, sensor_col):
    for doc_coeff_num, doc_coeff_name in enumerate(doc_title_list, 0): # loop through new sheet doc coeff titles
        if coeff == doc_coeff_name: #If these match then this doc contains the coeff title 
            for doc_row in doc_table: #loop through each row of the doc
                if code_row[0] in str(doc_row[sensor_col]): # If the sensor in the doc matched the sensor in the code

                    if coeff == "TubeID":  # A=pi*r^2. r=TubeID(AP)/2 * NumberTubes (AO) 
                        temp = 3.14 * ((float(doc_row[doc_coeff_num])*float(doc_row[doc_title_list.index('NumberTubes')])/2)**2)
                        working_row.append(str(temp))

                    elif coeff == "NominalFlowRate":  #Make the unit conversions between the doc (kg/s) and code (uS)
                        temp = float(doc_row[doc_coeff_num]) * 1000 / float(doc_row[doc_title_list.index('FlowCalFactor')]) 
                        working_row.append(str(temp))

                    else:
                        working_row.append(str(doc_row[doc_coeff_num]))
                    return working_row



def compileErDocRow(coeff, code_row, working_row, new_blue_doc_title_list, new_blue_doc_table, old_blue_doc_title_list, old_blue_doc_table,
    coriolis_red_doc_title_list, coriolis_red_doc_table, dens_visc_red_doc_title_list, dens_visc_red_doc_table, purpleDocTable, greenTableOne, 
    greenTableFour):
    
    blue_sensor_col = 3
    red_sensor_col = 1
    
    # Look at new sheet in blue ER doc for coeff 
    if coeff == "A 4":  # There is a space because there is an unrelated A4 in the red document 
        for doc_row in purpleDocTable:
            if code_row[0] in doc_row[0]:
                working_row.append(str(doc_row[1]))
                return

    if coeff == "GasFD":
        for doc_row in greenTableFour:
            if code_row[0] in doc_row[0]:
                working_row.append(str(doc_row[1]))
                return

    if coeff == "Gas Slope":
        for doc_row in greenTableFour:
            if code_row[0] in doc_row[0]:
                working_row.append(str(doc_row[2]))
                return

    if coeff == "Gas Offset":
        for doc_row in greenTableFour:
            if code_row[0] in doc_row[0]:
                working_row.append(str(doc_row[3]))
                return

    if coeff == "Liq Slope":
        for doc_row in greenTableFour:
            if code_row[0] in doc_row[0]:
                working_row.append(str(doc_row[4]))
                return

    if coeff == "Liq Offset":
        for doc_row in greenTableFour:
            if code_row[0] in doc_row[0]:
                working_row.append(str(doc_row[5]))
                return
               
    if coeff == "T03":
        for doc_row in greenTableOne:
            if code_row[0] in doc_row[0]:
                working_row.append(str(doc_row[1]))
                return
                
          

    # Look at new blue doc before old blue doc for most recent coeff values
    if checkDocForCoeff(new_blue_doc_title_list, new_blue_doc_table, coeff, code_row, working_row, blue_sensor_col) != None:
        return

    # Look at old sheet in blue ER doc for coeff
    if checkDocForCoeff(old_blue_doc_title_list, old_blue_doc_table, coeff, code_row, working_row, blue_sensor_col) != None:
        return

    if checkDocForCoeff(coriolis_red_doc_title_list, coriolis_red_doc_table, coeff, code_row, working_row, red_sensor_col) != None:
        return

    if checkDocForCoeff(dens_visc_red_doc_title_list, dens_visc_red_doc_table, coeff, code_row, working_row, red_sensor_col) != None:
        return

    working_row.append('―') # Else, add a '―' for coefficients that cant be populated
    return working_row


#Using default crapy round function
def formatString(stringVariable):
    try: #If the variable can be made into a float, return the float value
        return "%s"%round(float(stringVariable), 3)
    except: # If the variable cannot be made into a float, then just return the original string
        return stringVariable



def compareString(codeVal, DocVal):
    try: #try making the inputs float
        codeVal = float(codeVal)
        DocVal = float(DocVal)
        tolerance = codeVal * 0.0001
        isNumber = True
    except Exception as e: #if this falls through, then the inputs are strings
        isNumber = False

    if isNumber == True: # if the inputs are numbers, see if they are close enough to match
        if ((DocVal - (codeVal + tolerance))*(DocVal - (codeVal - tolerance)) <= 0):  # check if 
            return True
    
#    elif coeff == "NominalFlowRate":
        

    else: #if the inputs arent numbers, see if the strings match
        if codeVal == DocVal:
            return True

    return False



def createFinalArray(final_array, coeffs_to_compare, coeff_table, coeff_title_list, working_row, final_code_array, new_blue_doc_title_list,
    new_blue_doc_table, old_blue_doc_title_list, old_blue_doc_table, coriolis_red_doc_title_list, coriolis_red_doc_table, 
    dens_visc_red_doc_title_list, dens_visc_red_doc_table, purpleDocTable, final_doc_array, greenTableOne, greenTableFour):

    coeffPopulated = False

    final_array.append(coeffs_to_compare) # Add coeff titles to first row of final array 

    #Put code coeffs into final code array
    for code_row in coeff_table:  #put first row of coeffs from code in final array
        for coeff in coeffs_to_compare: # Loop through master coeff list
            coeffPopulated = False
            for code_coeff_num, code_list_coeff_name in enumerate(coeff_title_list[0], 0):
                if coeff == code_list_coeff_name:
                    coeffPopulated = True
                    working_row.append(code_row[code_coeff_num]) # Add coeff value to list 
                    break

            if coeffPopulated == False:
                working_row.append("―")  # Add a place holder if value doesnt exist

        final_code_array.append(working_row)
        working_row = []


        # put er document coeffs into final array 
        for coeff in coeffs_to_compare:  # Loop through master coeff list
            compileErDocRow(coeff, code_row, working_row, new_blue_doc_title_list, new_blue_doc_table, old_blue_doc_title_list, old_blue_doc_table,
                coriolis_red_doc_title_list, coriolis_red_doc_table, dens_visc_red_doc_title_list, dens_visc_red_doc_table, purpleDocTable, greenTableOne,
                greenTableFour)
        final_doc_array.append(working_row)
        working_row = []
   
   
   
   #Create and put match or no match row into final array
    for array_num, array in enumerate(final_code_array):
        final_array.append(final_code_array[array_num])
        final_array.append(final_doc_array[array_num])

        for element_num, element in enumerate(array):

            if compareString(final_code_array[array_num][element_num], final_doc_array[array_num][element_num]):
                working_row.append("Ok")

            # TODO move this to the compare section 
#            elif ((final_code_array[array_num][element_num] == "0.0") or (final_code_array[array_num][element_num] == "0")) and ((final_doc_array[array_num][element_num] == 'null') or (final_doc_array[array_num][element_num] == '―')): 
            elif ((final_code_array[array_num][element_num] == "0.0") or (final_code_array[array_num][element_num] == "0")) and ((final_doc_array[array_num][element_num] == 'null')):
                working_row.append("Ok") 

#            elif ((final_doc_array[array_num][element_num] == "0.0") or (final_doc_array[array_num][element_num] == "0")) and ((final_code_array[array_num][element_num] == 'null') or (final_code_array[array_num][element_num] == '―')):
            elif ((final_doc_array[array_num][element_num] == "0.0") or (final_doc_array[array_num][element_num] == "0")) and ((final_code_array[array_num][element_num] == 'null')): 
                working_row.append("Ok")

            else:
                working_row.append("No Match")

        final_array.append(working_row)
        working_row = []
   

def writeRowDescriptors(worksheet, workbook, row_num):
    right_border_format = workbook.add_format({'right': 1})
    right_and_bottom_border_format = workbook.add_format({'right': 1, 'bottom': 1})
    
    if (row_num == 0): 
        pass
    elif(row_num % 3.0 == 1.0): 
        worksheet.write(row_num, 0, "code", right_border_format)
    elif (row_num % 3.0 == 2.0): 
        worksheet.write(row_num, 0, "doc", right_border_format)
    elif(row_num % 3.0 == 0.0): 
        worksheet.write(row_num, 0, "compare", right_and_bottom_border_format)
   
  
def exportFinalArraytoExcelDocument(final_array, flowLinearityTable, greenTableTwo, flowLinearityMatch):
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

    default_format = workbook.add_format()   
    underline_format = workbook.add_format({'bottom': 1})
    right_border_format = workbook.add_format({'right': 1})
    right_and_bottom_border_format = workbook.add_format({'right': 1, 'bottom': 1})
    red_format = workbook.add_format({'font_color': 'red'})

    # write standard coeff compare worksheet
    for row_num, row in enumerate(final_array, 0):
        writeRowDescriptors(worksheet, workbook, row_num)
        for element_num, element in enumerate(row, 0):
            if row_num % 3.0 == 0:
                worksheet.write(row_num, element_num+1, final_array[row_num][element_num], underline_format)
            else:
                worksheet.write(row_num, element_num+1, final_array[row_num][element_num], default_format)


    # write flow linarization table in second sheet
    for row_num, row in enumerate(flowLinearityTable, 0):
#        writeRowDescriptors(flowLinearity, workbook, row_num)
        for element_num, element in enumerate(row, 0):
            if element_num == 1:
                flowLinearity.write(row_num, element_num, flowLinearityTable[row_num][element_num], right_border_format)
            else:
                flowLinearity.write(row_num, element_num, flowLinearityTable[row_num][element_num])

    for row_num, row in enumerate(greenTableTwo, 0):
#        writeRowDescriptors(flowLinearity, workbook, row_num)
        for element_num, element in enumerate(row, 0):
            if element_num == 1:
                flowLinearity.write(row_num, element_num +2, greenTableTwo[row_num][element_num], right_border_format)
            else:
                flowLinearity.write(row_num, element_num +2, greenTableTwo[row_num][element_num])
                
#        for element_num, element in enumerate(flowLinearityMatch, 0):
#                flowLinearity.write(element_num, 4, flowLinearityMatch[element_num])

       
                
    worksheet.conditional_format('A1:ZZ500', 
                                 {'type': 'text',
                                  'criteria': 'containing',
                                  'value': 'No Match',
                                  'format': red_format})
                                  
    flowLinearity.conditional_format('A1:ZZ500', 
                                 {'type': 'text',
                                  'criteria': 'containing',
                                  'value': 'No Match',
                                  'format': red_format})

    workbook.close()
   
   
   
   
def copyGreenErDoc(er_doc_green, greenTableOne, greenTableTwo, greenTableThree, greenTableFour):
    new_row = []
    title_row = False
    beforeTables = True
    tableNumber = 0
    
    newSensor = False
    tmp_sensor = None
    
    tables = er_doc_green.tables
    
    for table in tables:
        for row in table.rows:
            title_row = False
            if(beforeTables == False and len(new_row) != 0):

                if [s for s in new_row if '0.95' in s]:
                    sensor = new_row[0:2]
                else:
                    sensor = new_row[0]
                
                if tmp_sensor == sensor:
                    newSensor = False
                else:
                    newSensor = True
                    tmp_sensor = sensor
               
                if tableNumber == 1:
                    greenTableOne.append(new_row)

                elif tableNumber == 2:
                    if newSensor == True:
                        newSensor = False
                        greenTableTwo.append(sensor)
                    greenTableTwo.append(new_row[2:4])

                elif tableNumber == 3:
                    if newSensor == True:
                        newSensor = False
                        greenTableTwo.append([sensor])
#                        greenTableThree.append(sensor)
#                    greenTableThree.append(new_row[1:3])
                    greenTableTwo.append(new_row[1:3])

                elif tableNumber == 5: # because of the way the table is setup Sensor Model is read twice and this becomes 5 instead of 4
                    greenTableFour.append(new_row)

                new_row = []
            for cell in row.cells:
                if("Sensor Model") in cell.text:
                    tableNumber = tableNumber + 1
                    title_row = True
                    beforeTables = False

                if beforeTables == False and title_row == False:  # Add item to row if its after the tables start
#                    new_row = [x for x in new_row if x != '< 0.95'] #Remove blank string elements
                    if cell.text != '':  # Remove blank lines
                        new_row.append(cell.text)



def copyPurpleErDoc(er_doc_purple):
    purpleDocTable = []
    beforeTables = True
    title_row = False
    new_row = []

    #Put green ER-20015206_AP docx into list (original)
    tables = er_doc_purple.tables
    for table in tables:
        for row in table.rows:
            title_row = False
            if(beforeTables == False and len(new_row) != 0):
                purpleDocTable.append(new_row)
#                print(new_row)
                new_row = []

            for cell in row.cells:
                if("Sensor Model") in cell.text:
                    title_row = True
                    beforeTables = False
                if beforeTables == False and title_row == False:
                    if cell.text != '':  # Remove blank lines
                        new_row.append(cell.text)
    purpleDocTable.append(new_row)
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

   
   

   
   
   
   
   
   
   
   
   
   
   
   
   
   
   
   
   
   
   
   