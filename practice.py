import xlrd
import docx
from scriptUtility import *

#Read in ER documents
er_doc_red    = "H:\SensorScript\ER docs\ER-20015860_CF.xlsx"
er_doc_purple  = docx.Document("H:\SensorScript\ER docs\ER-20027172_AD.docx")

#Create workbooks for excel documents
wb_red = xlrd.open_workbook(er_doc_red)



print("Hello World!")
print()

new_row = []
new_blue_doc_title_list = []
new_blue_doc_table = []
old_blue_doc_title_list =[]
old_blue_doc_table =[]



copyBlueErDocFile(new_blue_doc_title_list, new_blue_doc_table, old_blue_doc_title_list, old_blue_doc_table)

#print(new_blue_doc_title_list[0])

#print(old_blue_doc_table[0])

    
# Finding sensor type location in ER doc
#for row in range(blue_old_params.nrows):
#    if sensor_type in blue_old_params.cell_value(row, 0):
#        blue_old_sensor_type_loc = row
#        break
#for row in range(blue_new_params.nrows):
#    if sensor_type in blue_new_params.cell_value(row, 0):
#        blue_new_sensor_type_loc = row
#        break


#print("Old params location: ", blue_old_sensor_type_loc)
#print("New params location: ", blue_new_sensor_type_loc)
#print()

#Populate ER Doc row 
#new_row = []
#new_row.append(blue_old_params.cell_value(blue_old_sensor_type_loc, FCF))
#new_row.append(blue_old_params.cell_value(blue_old_sensor_type_loc, K1))
#new_row.append(blue_old_params.cell_value(blue_old_sensor_type_loc, K2))
#new_row.append(blue_new_params.cell_value(blue_new_sensor_type_loc, ID_RESISTOR))
#new_row.append(blue_new_params.cell_value(blue_new_sensor_type_loc, ID_STRING))
#data.append(new_row)
#print(data)




print("End")