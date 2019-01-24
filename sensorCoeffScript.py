from scriptUtility import *


print("Hello World!")

#All the coeffs that need comparison
coeffs_to_compare = ["ID String", "FCF", "K1", "GasFD", "NominalFlowRate", "TubeID", "A4"]
#coeffs_to_compare = ["ID String", "NominalFlowRate", "FCF", "K1", "I.D. Resistor", "Drive Target", "Proportional Gain800", 
#    "Integral Gain800", "Proportional Gain2200", "Integral Gain2200", "Proportional Gain700", "Integral Gain700", 
#    "FD Limit", "Overshoot", "GasFD", "TubeID", "mp,   p < 0.4", "bp,   p < 0.4", "mp,   p > 0.4", "bp,   p > 0.4",
#    "DriveSaturationAlgorithm800", "PressureEffect_Flow_Liquid", "PressureEffect_Flow_Gas", "PressureEffect_Density", 
#    "Zero Stability", "A4", "TemperatureEffect_Density", "TemperatureEffect_Flow", "Tone Level", "Ramp Time", 
#    "BL Temp Coeff*", "Drive SP FCF", "Puck P FCF", "dF Tone Spacing", "Freq. Drift Limit*", "Max Sensor Current", 
#    "Minimum Flow Multiplier", "T03", "MassFlowAccuracy_Liquid", "MassFlowAccuracyMVD_Gas", "DensityAccuracy_Liquid", ] 

#Lists for copying the code file
sensor_list = [] #All the sensors listed in the code
coeff_title_list = [] 
coeff_table = [] #Table of all regular coefficients listed in the code 
cat_title_list = []
cat_table = [] #Table of all Category coefficients 
smv_title_list = []
smv_table = [] #Table of all Smart Meter Verification coefficients 
constants_table = []

#Blue ER document variables
#blue_excel_file   = "H:\SensorScript\ER docs\ER-20018334_AK.xlsx" #Read in blue ER document into an array
blue_excel_file   = r"C:\Users\AGrasmeder\Documents\SensorCoeffScript\ER docs\ER-20018334_AK.xlsx"
blue_excel_title_line_num = 15 # The row number of where the title names are located
blue_excel_data_line_num = 21 # The row number of where the coeff data begins
new_blue_doc_title_list = []
new_blue_doc_table = []
old_blue_doc_title_list =[]
old_blue_doc_table =[]

#Red ER document variables
#red_excel_file   = "H:\SensorScript\ER docs\ER-20015860_CF.xlsx" #Read in red ER document into an array
red_excel_file   = r"C:\Users\AGrasmeder\Documents\SensorCoeffScript\ER docs\ER-20015860_CF.xlsx"
red_excel_title_line_num = 6 # The row number of where the title names are located
red_excel_data_line_num = 10 # The row number of where the coeff data begins
coriolis_red_doc_title_list = []
coriolis_red_doc_table = []
dens_visc_red_doc_title_list = []
dens_visc_red_doc_table = []

#Lists for "compileErDocRow"
working_row = []

#Arrays for compiling final lists
final_array = []
final_code_array = []
final_doc_array = []




copyCodeFile(sensor_list, coeff_title_list, coeff_table, cat_title_list, cat_table, smv_title_list, smv_table, constants_table)

addCatAndSmvTablesToCoeffTable(coeff_title_list, coeff_table, cat_title_list, cat_table, smv_title_list, smv_table)

replaceVariablesWithDefinitions(coeff_table, constants_table)


#ER-20018334  /  Blue
copyExcelFile(new_blue_doc_title_list, new_blue_doc_table, old_blue_doc_title_list, old_blue_doc_table, blue_excel_file, 
    blue_excel_title_line_num, blue_excel_data_line_num) 

#ER-20015860  /  Red
copyExcelFile(coriolis_red_doc_title_list, coriolis_red_doc_table, dens_visc_red_doc_title_list, dens_visc_red_doc_table, red_excel_file, 
    red_excel_title_line_num, red_excel_data_line_num) 



createFinalArray(final_array, coeffs_to_compare, coeff_table, coeff_title_list, working_row, final_code_array, new_blue_doc_title_list,
    new_blue_doc_table, old_blue_doc_title_list, old_blue_doc_table, final_doc_array)

#exportFinalArraytoExcelDocument(final_array)




print()
print()
print()


print("End")


