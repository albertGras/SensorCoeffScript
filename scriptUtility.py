import xlrd

def copyCodeFile(sensor_list, coeff_title_list, coeff_table, cat_title_list, cat_table, smv_title_list, smv_table):
    code_file   = open("H:\SensorScript\practice.h")
#    code_file   = open("C:\PVCS\ProjectsDB\Kinetis_DB\k2Src\k_src_app\coriolis\sensor.cpp")

    inCoeffTable = False
    inSmvTable = False
    inCatTable = False

    line = []
    count = 0

    extraCharacters = ['{', ' ', '\"', '\n', 'f','}', ';', '/',]
    extraCharacters2 = ['{', '\"', '\n', 'f','}', ';', '/',]

    code_file.seek(0)   # Go to beginning of file

    while count < 3:   # while all 3 tables havent been read (Coeff, Cat, Smv)
        line = code_file.readline()
        if not line:                  # if end of file, stop reading file
            break
        if "End" in line:   # end of one of the tables
            count = count + 1
            inCoeffTable = False
            inSmvTable = False
            inCatTable = False

        elif "COEFF_TABLE" in line:      #Create coeff table list
            for item in extraCharacters2:  #remove unwanted characters in the line
                line = line.replace(item, '')
            line = line.split(',')
            coeff_title_list.append(line)
            inCoeffTable = True
        elif inCoeffTable == True:       #Create coeff table
            for item in extraCharacters:  #remove unwanted characters in the line
                line = line.replace(item, '')
            line = line.split(',')
            if line[0] == '':  # Remove blank lines
                continue
            if not any("--" in s for s in line): #Remove comment rows with additional coeff decriptions
                sensor_list.append(line[0])  # Take the sensor types from each line
                coeff_table.append(line)

        elif "CAT_TABLE" in line:        #Create cat table list
            for item in extraCharacters2:  #remove unwanted characters in the line
                line = line.replace(item, '')
            line = line.split(',')
            cat_title_list.append(line)
            inCatTable = True
        elif inCatTable == True:         #Create cat table
            for item in extraCharacters:  #remove unwanted characters in the line
                line = line.replace(item, '')
            line = line.split(',')
            if line[0] == '':  # Remove blank lines
                continue
            line[0] = line[0].split('=', 1)[-1]  #Remove all characters before the '='
            cat_table.append(line)

        elif "SMV_TABLE" in line:        #Create smv table list
            for item in extraCharacters2:  #remove unwanted characters in the line
                line = line.replace(item, '')
            line = line.split(',')
            smv_title_list.append(line)
            inSmvTable = True
        elif inSmvTable == True:         #Create smv table
            for item in extraCharacters:  #remove unwanted characters in the line
                line = line.replace(item, '')
            line = line.split(',')
            if line[0] == '':  # Remove blank lines
                continue
            line[0] = line[0].split('=', 1)[-1]  #Remove all characters before the '='
            smv_table.append(line)







def copyBlueErDocFile(new_blue_doc_title_list, new_blue_doc_table, old_blue_doc_title_list, old_blue_doc_table):
    #Read in blue ER document into array
    er_doc_blue   = "H:\SensorScript\ER docs\ER-20018334_AK.xlsx"
    
    #Create workbooks for excel document
    wb_blue = xlrd.open_workbook(er_doc_blue)

    #Get the correct sheet in the excel document
    blue_old_params = wb_blue.sheet_by_index(1)
    blue_new_params = wb_blue.sheet_by_index(2)
    
    blue_title_line_number = 15
    blue_data_line_number = 21
    
    new_row = []

    # Copy all coefficient names in blue er doc - old params 
    num_cols = blue_old_params.ncols   # Number of columns
    for col_idx in range(0, num_cols):  # Iterate through columns
        old_blue_doc_title_list.append(blue_old_params.cell_value(blue_title_line_number, col_idx))  # Get cell object by row, col
#    print (old_blue_doc_title_list)
#    print()

    # Copy all values in blue er doc - old params 
    for row_idx in range(blue_data_line_number, blue_old_params.nrows):    # Iterate through rows
        for col_idx in range(0, num_cols):  # Iterate through columns
            new_row.append(blue_old_params.cell_value(row_idx, col_idx))  # Get cell object by row, col
#        print (new_row)
        old_blue_doc_table.append(new_row)
        new_row = []

#    print()

    # Copy all coefficient names in blue er doc - new params 
    num_cols = blue_new_params.ncols   # Number of columns
    for col_idx in range(0, num_cols):  # Iterate through columns (x axis direction)
        new_blue_doc_title_list.append(blue_new_params.cell_value(blue_title_line_number, col_idx))  # Get cell object by row, col
#    print (new_blue_doc_title_list)
#    print()


    # Copy all values in blue er doc - new params 
    for row_idx in range(blue_data_line_number, blue_new_params.nrows):    # Iterate through rows (y axis direction)
        for col_idx in range(0, num_cols):  # Iterate through columns
            new_row.append(blue_new_params.cell_value(row_idx, col_idx))  # Get cell object by row, col
#        print (new_row)
        new_blue_doc_table.append(new_row)
        new_row = []
    
    
    
def compileErDocRow(coeff, code_row, working_row, new_blue_doc_title_list, new_blue_doc_table, old_blue_doc_title_list, old_blue_doc_table):
    match = False
    
    for doc_coeff_num, doc_coeff_name in enumerate(new_blue_doc_title_list, 0): # loop through new sheet doc coeff titles
        if coeff == doc_coeff_name: #If these match then this doc contains the coeff title 
            match = True
            for doc_row in new_blue_doc_table: #loop through each row of the blue doc
                if code_row[0] in doc_row[3]: # If the sensor in the doc matched the sensor in the code
                    working_row.append(str(doc_row[doc_coeff_num]))
                    match = False
                    return working_row
                    
    if match == False: #Didnt find the coeff in the new sheet
        for doc_coeff_num, doc_coeff_name in enumerate(old_blue_doc_title_list, 0): # loop through old sheet doc coeff titles
            if coeff == doc_coeff_name: #If these match then this doc contains the coeff title
                match = True
                for doc_row in old_blue_doc_table: #loop through each row of the blue doc to find the right sensor
                    if code_row[0] in doc_row[3]: # If the sensor in the doc matched the sensor in the code
                        working_row.append(str(doc_row[doc_coeff_num]))
                        match = False
                        return working_row
    
    
    
    




    