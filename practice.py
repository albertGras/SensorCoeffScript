#import docx
#import numpy
from scriptUtility import *
import pandas as pd
import xlsxwriter
import openpyxl


print("Hello World!")

#All the coeffs that need comparison
#coeffs_to_compare = ["ID String", "FCF", "K1", "GasFD", "NominalFlowRate", "TubeID", "A4"] #All the coeffs that need comparison
coeffs_to_compare = ["ID String", "NominalFlowRate", "FCF", "K1", "I.D. Resistor", "Drive Target", "Proportional Gain800", 
    "Integral Gain800", "Proportional Gain2200", "Integral Gain2200", "Proportional Gain700", "Integral Gain700", 
    "FD Limit", "Overshoot", "GasFD", "TubeID", "mp,   p < 0.4", "bp,   p < 0.4", "mp,   p > 0.4", "bp,   p > 0.4",
    "DriveSaturationAlgorithm800", "PressureEffect_Flow_Liquid", "PressureEffect_Flow_Gas", "PressureEffect_Density", 
    "Zero Stability", "A4", "TemperatureEffect_Density", "TemperatureEffect_Flow", "Tone Level", "Ramp Time", 
    "BL Temp Coeff*", "Drive SP FCF", "Puck P FCF", "dF Tone Spacing", "Freq. Drift Limit*", "Max Sensor Current", 
    "Minimum Flow Multiplier", "T03", "MassFlowAccuracy_Liquid", "MassFlowAccuracyMVD_Gas", "DensityAccuracy_Liquid", ] 

#Lists for copying the code file
sensor_list = [] #All the sensors listed in the code
coeff_title_list = [] 
coeff_table = [] #Table of all regular coefficients listed in the code 
cat_title_list = []
cat_table = [] #Table of all Category coefficients 
smv_title_list = []
smv_table = [] #Table of all Smart Meter Verification coefficients 
constants_table = []

#Lists for copying the Blue ER document
new_blue_doc_title_list = []
new_blue_doc_table = []
old_blue_doc_title_list =[]
old_blue_doc_table =[]

#Lists for "compileErDocRow"
working_row = []

final_array = []
final_code_array = []
final_doc_array = []


coeffPopulated = False

copyCodeFile(sensor_list, coeff_title_list, coeff_table, cat_title_list, cat_table, smv_title_list, smv_table, constants_table)
copyBlueErDocFile(new_blue_doc_title_list, new_blue_doc_table, old_blue_doc_title_list, old_blue_doc_table)
addCatAndSmvTablesToCoeffTable(coeff_title_list, coeff_table, cat_title_list, cat_table, smv_title_list, smv_table)

# Replace variables with their values 
#for array_num, array in enumerate(coeff_table, 0):
for code_row_num, code_row in enumerate(coeff_table, 0):
    for code_element_num, code_element in enumerate(code_row, 0):
        for x in constants_table:
            if x[0] == code_element:
                coeff_table[code_row_num][code_element_num] = x[1]


# Create final, results array
final_array.append(coeffs_to_compare)

#Put code coeffs into final array
for code_row in coeff_table:  #put first row of coeffs from code in final array
    for coeff in coeffs_to_compare: # Loop through master coeff list
        coeffPopulated = False
        for code_coeff_num, code_list_coeff_name in enumerate(coeff_title_list[0], 0):
            if coeff == code_list_coeff_name:
                coeffPopulated = True
                working_row.append(code_row[code_coeff_num]) # Add coeff value to list 
        if coeffPopulated == False:
            working_row.append("---")  # Add a place holder if value doesnt exist
    final_code_array.append(working_row)
#    print(working_row)
    working_row = []

    
    
    # put er document coeffs into final array 
    for coeff in coeffs_to_compare:  # Loop through master coeff list
        compileErDocRow(coeff, code_row, working_row, new_blue_doc_title_list, new_blue_doc_table, old_blue_doc_title_list, old_blue_doc_table)
    final_doc_array.append(working_row)
#    print (working_row)
    print()
    working_row = []
    

    
for array_num, array in enumerate(final_code_array):
    final_array.append(final_code_array[array_num])
    print(final_code_array[array_num])
    final_array.append(final_doc_array[array_num])
    print(final_doc_array[array_num])
    
    for element_num, element in enumerate(array):
        if formatString(final_code_array[array_num][element_num]) == formatString(final_doc_array[array_num][element_num]):
#        print( final_code_array[array_num][element_num])
#        print( final_doc_array[array_num][element_num])
            working_row.append("Ok")
        else:
            working_row.append("No Match")
    final_array.append(working_row)
    print(working_row)
    print()
    working_row = []



# Create an new Excel file and add a worksheet.
workbook = xlsxwriter.Workbook('H:\SensorScript\demo.xlsx')
worksheet = workbook.add_worksheet()

# Widen the first column to make the text clearer.
worksheet.set_column('A:AZ', 19)


default_format = workbook.add_format()   
underline_format = workbook.add_format({'bottom': 1})
red_format = workbook.add_format({'font_color': 'red'})


for row_num, row in enumerate(final_array, 0):
    for element_num, element in enumerate(row, 0):
        if row_num % 3.0 == 0:
            worksheet.write(row_num, element_num, final_array[row_num][element_num], underline_format)
        else:
            worksheet.write(row_num, element_num, final_array[row_num][element_num], default_format)



worksheet.conditional_format('A1:AZ20',
                              {'type': 'text',
                              'criteria': 'containing',
                              'value': 'No Match',
                              'format': red_format})

workbook.close()


print()
print()
print()


print("End")


