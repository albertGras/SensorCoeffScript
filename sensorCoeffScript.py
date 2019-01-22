#import docx
#import numpy
from scriptUtility import *


print("Hello World!")

#All the coeffs that need comparison
#coeffs_to_compare = ["ID String", "FCF", "K1", "GasFD", "NominalFlowRate", "TubeID", "A4"]
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




copyCodeFile(sensor_list, coeff_title_list, coeff_table, cat_title_list, cat_table, smv_title_list, smv_table, constants_table)

copyBlueErDocFile(new_blue_doc_title_list, new_blue_doc_table, old_blue_doc_title_list, old_blue_doc_table)

er_doc_red   = "H:\SensorScript\ER docs\ER-20015860_CF.xlsx" #Read in red ER document into an array
wb_red = xlrd.open_workbook(er_doc_red) #Create workbook for excel red er document

new_row = []

# title = row 7 (so 6?)
#data = 12
# data is 11 on denisty viscosity sheet

addCatAndSmvTablesToCoeffTable(coeff_title_list, coeff_table, cat_title_list, cat_table, smv_title_list, smv_table)

replaceVariablesWithDefinitions(coeff_table, constants_table)



createFinalArray(final_array, coeffs_to_compare, coeff_table, coeff_title_list, working_row, final_code_array, new_blue_doc_title_list,
    new_blue_doc_table, old_blue_doc_title_list, old_blue_doc_table, final_doc_array)

exportFinalArraytoExcelDocument(final_array)




print()
print()
print()


print("End")


