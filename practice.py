import xlrd
import docx

print("Hello World!")
print()

#code_file   = open("H:\SensorScript\practice.h")
code_file   = open("C:\PVCS\ProjectsDB\Kinetis_DB\k2Src\k_src_app\coriolis\sensor.cpp")

coeff_list = ["ID String", "FCF", "K1", "GasFD", "TubeID"] #All the coeffs that need comparison

sensor_list = [] #All the sensors listed in the code
coeff_title_list = [] 
coeff_table = [] #Table of all regular coefficients listed in the code 
cat_title_list = []
cat_table = [] #Table of all Category coefficients 
smv_title_list = []
smv_table = [] #Table of all Smart Meter Verification coefficients 

inCoeffTable = False
inSmvTable = False
inCatTable = False

line = []
count = 0

extraCharacters = ['{', ' ', '\"', '\n', 'f','}', ';', '/',]

code_file.seek(0)   # Go to beginning of file


while count < 3:   # while all 3 tables havent been read (Coeff, Cat, Smv)
    line = code_file.readline()
    if not line:                  # if end of file, stop reading file
        break
    for item in extraCharacters:  #remove unwanted characters in the line
        line = line.replace(item, '')
    line = line.split(',')
    if "End" in line:   # end of one of the tables
        count = count + 1
        inCoeffTable = False
        inSmvTable = False
        inCatTable = False

    elif "COEFF_TABLE" in line:      #Create coeff table list
        coeff_title_list.append(line)
        inCoeffTable = True
    elif inCoeffTable == True:       #Create coeff table
        if not any("--" in s for s in line):
            sensor_list.append(line[0])  # Take the sensor types from each line
            coeff_table.append(line)

    elif "CAT_TABLE" in line:        #Create cat table list
        cat_title_list.append(line)
        inCatTable = True
    elif inCatTable == True:         #Create cat table
        cat_table.append(line)

    elif "SMV_TABLE" in line:        #Create smv table list
        smv_title_list.append(line)
        inSmvTable = True
    elif inSmvTable == True:         #Create smv table
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
print(coeff_table[5])
print(coeff_table[6])
print(coeff_table[7])
print(coeff_table[8])
print(coeff_table[9])
print(coeff_table[10])
print(coeff_table[11])
print(coeff_table[12])
print(coeff_table[13])
print(coeff_table[14])
print(coeff_table[15])
print(coeff_table[16])
print(coeff_table[17])
print(coeff_table[18])
print(coeff_table[19])
print()
print()
print(cat_title_list)
print()
print(cat_table[0])
print(cat_table[1])
print(cat_table[2])
print(cat_table[3])
print(cat_table[4])
print(cat_table[5])
print(cat_table[6])
print(cat_table[7])
print(cat_table[8])
print(cat_table[9])
print()
print()
print(smv_title_list)
print()
print(smv_table[0])
print(smv_table[1])
print(smv_table[2])
print(smv_table[3])
print(smv_table[4])
print(smv_table[5])
print(smv_table[6])
print(smv_table[7])
print(smv_table[8])
print(smv_table[9])



