import xlrd
import docx
import numpy

print("Hello World!")




# "K2"
# "ID resistor"

# "Drive Target"

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


#sensor_type = "CMF350"
#print(sensor_type)
print()







code_file   = open("H:\SensorScript\practice.h")

# Make a list of all ID Strings


#for line in code_file:      #find string in file
#    if "ID String" in line:
#        code_line = line
#        break
#code_array = code_line.split(',')

coeff_list = ["ID String", "FCF", "K1", "GasFD", "TubeID"] #All the coeffs that need comparison
title_list = [] #All the coeffs listed in the code 
sensor_list = [] #All the sensors listed in the code

new_row = []
full_report = []

correctTable = False
match_found = False
CAT_found = False
SMV_found = False

extraCharacters = ['{',' ','\"','}','\n','/*', 'f']
code_file.seek(0)   # Go back to beginning of file



#Get sensor list and coefficient titles from code
for line in code_file:
    if "};" in line:            #If its the end of the table
        correctTable = False
    if correctTable == True: #Dont read the file until in the right spot
        for item in extraCharacters:  #remove unwanted characters in the line
            line = line.replace(item, '')
        current_line = line.split(',')
        sensor_list.append(current_line[0])  # Take the sensor types from each line
        
        for coeff in coeff_list:  #Match up the coefficients and titles
            for title_num, title in enumerate(title_list, 0):
                if title == coeff:                           #Put the code number coefficient
                    new_row.append(current_line[title_num])  # into the full data report
                    match_found = True
                    break #No need to keep looking after the matching title is found 
            if CAT_found == True: # CAT variables located in another table
                print() # Search through all the CAT table variables
            if SMV_found == True: # SMV variables located in another table
                print() # Search through all the SMV table variables
            if match_found == False:
                print()
            
            match_found = False

        print(new_row)
        full_report.append(new_row)
        new_row = [] #Clear data for next sensor type 
        
    if "ID String" in line: # Take the code coefficient names and make a list
        correctTable = True
        line = line.replace('/', '')
        line = line.replace('\n', '')
        title_list = line.split(',')  # Makes the title list 
        #print(title_list)
        #print()
        #print()


    
#print(full_report[0])
#print(full_report[1])
#print(full_report[2])
        
#print()
#print()
#print(sensor_list)
#print(full_report)


print("End")










