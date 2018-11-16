import xlrd
import docx
import numpy

print("Hello World!")



#maybe use below ord function instead
#number = ord(character) - 97


# "Base Model"
# "FCF"
# "K1"
# "K2"
# "ID resistor"
# "ID string"
# "Drive Target"
#  P,I *************
# "A3 Limit"
# "FD Limit"
# "Drive Overshoot Limit"
# "GasFD Algorithm Constant"
# "Tube ID Cross Sectional Area"
# "VOS Gas Slope"
# "VOS Gas Offset"
# "VOS Liquid Slope"
# "VOS Liquid Offset"
# "Drive Saturation Algorithm"
# "Minimum Drive"
# "Fixed PGA"
# "Filtered PO "
# "Flow Pressure Comp (Liquid)"
# "Flow Pressure Comp (Gas)"
# "Density Pressure Comp"
# "Zero Stability"
# "A4 Density Case Temp Coeff"
# "Temp Effect on Density"
# "Temp Effect on Flow"
# "Meter Verification: Tone Level"
# "Meter Verification: Ramp Time"
# "Meter Verification: BL Tempco"
# "Meter Verification: Drv SP FCF"
# "Meter Verification: Puck P FCF"
# "Meter Verification: Tone Spacing"
# "Meter Verification: Freq Diff Limit"
# "Meter Verification: Max Current"
# "Meter Verification: Min Stiffness"
# "Meter Verification: Max Stiffness"
# ""Min Flow Multipler""
# "2-Rate Zero correction ("T03")"
# "Liquid Flow Performance"
# "Gas Flow Performance"
# "Density Performance"
# "Piecewise Linearization"


FCF = 4
K1 = 5
K2 = 6
ID_RESISTOR = 4
ID_STRING = 3

#The location on the excel spreadsheet where the names
#of the coefficients begin. Currently "Sensor Model Code"
first_coeff_loc = 13 
last_coeff_loc = 18

sensor_type = "CMF350"
print(sensor_type)
print()

#Read in ER documents
coeff_names   = "H:\SensorScript\Sensor_Worksheet_v23.xlsx"
er_doc_blue   = "H:\SensorScript\ER docs\ER-20018334_AK.xlsx"
er_doc_red    = "H:\SensorScript\ER docs\ER-20015860_CF.xlsx"
er_doc_green  = docx.Document("H:\SensorScript\ER docs\ER-20015206_AP.docx")
r_doc_purple  = docx.Document("H:\SensorScript\ER docs\ER-20027172_AD.docx")

#Create workbooks for excel documents
wb_coeff = xlrd.open_workbook(coeff_names)
wb_blue = xlrd.open_workbook(er_doc_blue)
wb_red = xlrd.open_workbook(er_doc_red)

#Get the correct sheet in the excel document
coeff_sheet = wb_coeff.sheet_by_index(0)
blue_old_params = wb_blue. sheet_by_index(1)
blue_new_params = wb_blue. sheet_by_index(2)

ERdocList = []
new_row = []
beforeTables = True
title_row = False

#Put docx ER docs into list for easier access 
tables = er_doc_green.tables
for table in tables:
    for row in table.rows:
        title_row = False
        if(beforeTables == False and len(new_row) != 0):
            ERdocList.append(new_row)
            new_row = []
        for cell in row.cells:
            if("Sensor Model") in cell.text:
                title_row = True
                beforeTables = False
            elif(beforeTables == False and title_row == False):
                new_row.append(cell.text)
            else:
                print(cell.text)

# Look for sensor type in docx ER docs
for index in range(0, len(ERdocList)):
    if(sensor_type in ERdocList[index][0]):
        print("MATCH")
    print(ERdocList[index])
print()

# Populate coefficient names for title row
#for row in range(first_coeff_loc, coeff_sheet.nrows):
data = []
new_row = []
for row in range(first_coeff_loc, last_coeff_loc):
    new_row.append(coeff_sheet.cell_value(row ,0))
data.append(new_row)

num_of_coeff = len(new_row)
#print(data)


# Finding sensor type location in ER doc
for row in range(blue_old_params.nrows):
    if sensor_type in blue_old_params.cell_value(row, 0):
        blue_old_sensor_type_loc = row
        break
for row in range(blue_new_params.nrows):
    if sensor_type in blue_new_params.cell_value(row, 0):
        blue_new_sensor_type_loc = row
        break
        

print("Old params location: ", blue_old_sensor_type_loc)
print("New params location: ", blue_new_sensor_type_loc)
print()

#Populate ER Doc row 
new_row = []
new_row.append(blue_old_params.cell_value(blue_old_sensor_type_loc, FCF))
new_row.append(blue_old_params.cell_value(blue_old_sensor_type_loc, K1))
new_row.append(blue_old_params.cell_value(blue_old_sensor_type_loc, K2))
new_row.append(blue_new_params.cell_value(blue_new_sensor_type_loc, ID_RESISTOR))
new_row.append(blue_new_params.cell_value(blue_new_sensor_type_loc, ID_STRING))
data.append(new_row)
#print(data)


#Populate Code row 
new_row = []
code_line = []
code_file = open("C:\PVCS\ProjectsDB\Kinetis_DB\k2Src\k_src_app\coriolis\sensor.cpp")
for line in code_file:      #find string in file
    if "\""+ sensor_type in line:
        code_line = line
        break

code_line = code_line.replace(' ','')
code_line = code_line.replace('f','')
code_array = code_line.split(',')
#print(code_array)
new_row = []
new_row.append(code_array[2])
new_row.append(code_array[3])
new_row.append("NaN")
new_row.append("NaN")
new_row.append(code_array[1])
data.append(new_row)


# Compare ER docs and code to see if they match
# Makes the 4th row, status
new_row = []
for index in range(0, num_of_coeff):
    if str(data[1][index]) == data[2][index]:
        new_row.append("Match")
    else:
        new_row.append("Error")

data.append(new_row)

print(data[0])
print(data[1])
print(data[2])
print(data[3])


print("End")












