import xlrd
import docx
from transitions import Machine



print("End")



print("Hello World!")
print()

#code_file   = open("H:\SensorScript\practice.h")
code_file   = open("C:\PVCS\ProjectsDB\Kinetis_DB\k2Src\k_src_app\coriolis\sensor.cpp")

coeff_list = ["ID String", "FCF", "K1", "GasFD", "TubeID"] #All the coeffs that need comparison

sensor_list = [] #All the sensors listed in the code
coeff_title_list = []
coeff_table = [] #All the coeffs listed in the code 
cat_title_list = []
cat_table = []
smv_title_list = []
smv_table = []

inCoeffTable = False
inSmvTable = False
inCatTable = False

line = []
count = 0

extraCharacters = ['{', ' ', '\"', '\n', '*', 'f', '/','}', ';']
#code_file.seek(0)   # Go back to beginning of file

code_file.seek(0)

#if not line:

while count < 3:
#start read line funtion
    line = code_file.readline()
    if not line:
        break
    for item in extraCharacters:  #remove unwanted characters in the line
        line = line.replace(item, '')
    line = line.split(',')
#end read line function
    if "End" in line:
        count = count + 1
        inCoeffTable = False
        inSmvTable = False
        inCatTable = False
    elif "COEFF_TABLE" in line:
        print("in coeff table")
        coeff_title_list.append(line)
        inCoeffTable = True
    elif inCoeffTable == True:
        sensor_list.append(line[0])  # Take the sensor types from each line
        coeff_table.append(line)
    elif "CAT_TABLE" in line:
        print("in cat table")
        cat_title_list.append(line)
        inCatTable = True
    elif inCatTable == True:
        cat_table.append(line)
    elif "SMV_TABLE" in line:
        print("in smv table")
        smv_title_list.append(line)
        inSmvTable = True
    elif inSmvTable == True:
        smv_table.append(line)


print("End")


print(sensor_list)
print()
print()
print(coeff_title_list)
print()
print(coeff_table[0])
print(coeff_table[1])
print(coeff_table[2])
print(coeff_table[3])
print(coeff_table[4])
print()
print()
print(cat_table[0])
print(cat_table[1])
print()
print()
print(smv_table[0])
print(smv_table[1])



