import xlrd
import docx
import numpy
from scriptUtility import *


print("Hello World!")


working_row = []
final_array = []
beforeTables = True
title_row = False
match = False


coeffs_to_compare = ["ID String", "FCF", "K1", "GasFD", "NominalFlowRate", "TubeID", "A4"] #All the coeffs that need comparison

#Lists for "copyCodeFile"
sensor_list = [] #All the sensors listed in the code
coeff_title_list = [] 
coeff_table = [] #Table of all regular coefficients listed in the code 
cat_title_list = []
cat_table = [] #Table of all Category coefficients 
smv_title_list = []
smv_table = [] #Table of all Smart Meter Verification coefficients 

#Lists for "copyBlueErDocFile"
new_blue_doc_title_list = []
new_blue_doc_table = []
old_blue_doc_title_list =[]
old_blue_doc_table =[]



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
        for doc_coeff_num, doc_coeff_name in enumerate(new_blue_doc_title_list, 0): # loop through new sheet doc coeff titles
            if coeff == doc_coeff_name: #If these match then this doc contains the coeff title 
#                print(coeff)
#                print(doc_coeff_name)
                match = True
#                print("Match new")
                for doc_row in new_blue_doc_table: #loop through each row of the blue doc
                    if code_row[0] in doc_row[0]: # If the sensor in the doc matched the sensor in the code
#                        print("Match 2 new")
#                        print(doc_coeff_name)
#                        print(doc_coeff_num)
#                        print(doc_row[doc_coeff_num])
                        working_row.append(doc_row[doc_coeff_num])
                        match = False
                        break;
                        #return 
                        
        if match == False: #Didnt find the coeff in the new sheet
            for doc_coeff_num, doc_coeff_name in enumerate(old_blue_doc_title_list, 0): # loop through old sheet doc coeff titles
                if coeff == doc_coeff_name: #If these match then this doc contains the coeff title
                    match = True
#                    print("Match old")
                    for doc_row in old_blue_doc_table: #loop through each row of the blue doc to find the right sensor
                        if code_row[0] in doc_row[0]: # If the sensor in the doc matched the sensor in the code
                            print(doc_coeff_num)
                            print(doc_coeff_name)
                            if doc_coeff_name == "ID String":
                                print("repeat!")
                                break;
                            else:
#                                print("Match 2 old")
#                                print(doc_coeff_name)
#                                print(doc_coeff_num)
#                                print(doc_row[doc_coeff_num])
                                working_row.append(doc_row[doc_coeff_num])
                                match = False
                                break;
    
    print(working_row)
    working_row = []
    print()
    
    
    
    
    
    
    
    
    
    
    
    
    
    
#    final_array.append(working_row)
#    print(working_row)
#    working_row = []
#        for doc_coeff_num, doc_coeff_name in enumerate(blue_doc_title_list_new, 0):
#            if doc_coeff == doc_coeff_name:
#                new_row.append(blue_doc_table[doc_coeff_num])
#                match = True
#        if match == False:
#            for doc_coeff_num, doc_coeff_name in enumerate(blue_doc_title_list_old, 0):
#                if doc_coeff == doc_coeff_name:
#                    new_row.append(blue_doc_table[doc_coeff_num])
#                    match = True
#    continue

    
print()
print()
print()
#print(coeff_title_list)











# Look for sensor type in docx ER docs
#for index in range(0, len(ERdocList)):
#    if(sensor_type in ERdocList[index][0]):
#        print("MATCH")
#    print(ERdocList[index])
#print()




#trying something new
#Put green ER-20015206_AP docx into list
#tables = er_doc_green.tables
#for table in tables:
#    for row in table.rows:
#        title_row = False
#        if(beforeTables == False and len(new_row) != 0):
#            ERdocList.append(new_row)
#            new_row = []
#        for cell in row.cells:
#            if("Sensor Model") in cell.text:
#                title_row = True
#                beforeTables = False
#            if beforeTables == False:
#                if cell.text != '':  # Remove blank lines
#                    new_row.append(cell.text)
#    print(table_num)

#for index in range(0, len(ERdocList)):
#    print(ERdocList[index])


















# Compare ER docs and code to see if they match
# Makes the 4th row, status
#new_row = []
#for index in range(0, num_of_coeff):
#    if str(data[1][index]) == data[2][index]:
#        new_row.append("Match")
#    else:
#        new_row.append("Error")

#data.append(new_row)

#print(data[0])
#print(data[1])
#print(data[2])
#print(data[3])




print("End")












