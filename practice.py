import xlrd
import docx
import numpy
from scriptUtility import *


print("Hello World!")


final_array = []

coeffs_to_compare = ["ID String", "FCF", "K1"] #, "GasFD", "NominalFlowRate", "TubeID", "A4"] #All the coeffs that need comparison

#Lists for copying the code
sensor_list = [] #All the sensors listed in the code
coeff_title_list = [] 
coeff_table = [] #Table of all regular coefficients listed in the code 
cat_title_list = []
cat_table = [] #Table of all Category coefficients 
smv_title_list = []
smv_table = [] #Table of all Smart Meter Verification coefficients 

#Lists for copying the Blue ER document
new_blue_doc_title_list = []
new_blue_doc_table = []
old_blue_doc_title_list =[]
old_blue_doc_table =[]

#Lists for "compileErDocRow"
working_row = []

copyCodeFile(sensor_list, coeff_title_list, coeff_table, cat_title_list, cat_table, smv_title_list, smv_table)
copyBlueErDocFile(new_blue_doc_title_list, new_blue_doc_table, old_blue_doc_title_list, old_blue_doc_table)



# Create final, results array

#Put code coeffs into final array
for code_row in coeff_table:  #put first row of coeff from code in final array
    for coeff in coeffs_to_compare: # Loop through master coeff list
        for code_coeff_num, code_list_coeff_name in enumerate(coeff_title_list[0], 0):
            if coeff == code_list_coeff_name:
                working_row.append(code_row[code_coeff_num -1])
    final_array.append(working_row)
    print(working_row)
    working_row = []

    
    
    # put er document coeffs into final array 
    for coeff in coeffs_to_compare:  # Loop through master coeff list
    
        compileErDocRow(coeff, code_row, working_row, new_blue_doc_title_list, 
            new_blue_doc_table, old_blue_doc_title_list, old_blue_doc_table)
    final_array.append(working_row)
    print (working_row)
    print()
    working_row = []





print("End")







