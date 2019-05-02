from scriptUtility import *
from codeReadHelper import *
from docReadHelper import *
from excelWriteHelper import *
from docx import Document


print("Hello World!")

#All the coeffs that need comparison
COEFFS_TO_COMPARE = ["ID String", "FlowCalFactor", "K1", "flags", "Drive Saturation Algorithm 800", "T03",
    "I.D. Resistor", "TubeID", "NominalFlowRate",
    "PressureEffect_Flow_Liquid", "PressureEffect_Density", "ZeroStability", "A 4", 
    "Drive Target", "Proportional Gain 800", "Integral Gain 800", "FD Limit", "Overshoot", 
    "TemperatureEffect_Density", "TemperatureEffect_Flow", "Tone Level", "Ramp Time", "BL Temp Coeff", "Drive SP FCF", 
    "Puck P FCF", "dF Tone Spacing", "Freq. Drift Limit", "Max Sensor Current", "Minimum Flow Multiplier", 
    "MassFlowAccuracy_Liquid", "MassFlowAccuracyMVD_Gas", "DensityAccuracy_Liquid", "Drive SP",   
] 

#coeffs not in 5700 
# "Proportional Gain 2200", "Integral Gain 2200",

#Lists for copying the code file
sensorList = [] #All the sensors listed in the code
mainCoeffList = [] 
coeffTable = [] #Table of all regular coefficients listed in the code 
catCoeffList = []
catTable = [] #Table of all Category coefficients 
smvCoeffList = []
smvTable = [] #Table of all Smart Meter Verification coefficients 
constantsTable = []

#Blue ER document variables
blueFile   = "H:\ER docs\ER-20018334_AK.xlsx" #Read in blue ER document into an array
#blueFile   = r"C:\Users\AGrasmeder\Documents\SensorCoeffScript\ER docs\ER-20018334_AK.xlsx"
blueTitleLineNum = 15 # The row number of where the title names are located
blueDataLineNum = 21 # The row number of where the coeff data begins
newBlueCoeffList = []
newBlueTable = []
oldBlueCoeffList =[]
oldBlueTable =[]

#Red ER document variables
redFile   = "H:\ER docs\ER-20015860_CH.xlsx" #Read in red ER document into an array
#redFile   = r"C:\Users\AGrasmeder\Documents\SensorCoeffScript\ER docs\ER-20015860_CF.xlsx"
redCoeffLineNum = 6 # The row number of where the title names are located
redCoeffDataNum = 10 # The row number of where the coeff data begins
coriolisRedCoeffList = []
coriolisRedTable = []
densViscRedCoeffList = []
densViscRedTable = []


#Green ER document variables
greenFile   = Document("H:\ER docs\ER-20015206_AP.docx") #Read in green ER document into an array
greentableOne = []
greenTableTwo = []
greenTableThree = []
greenTableFour = []

#Purple ER document variables
purpleFile   = Document("H:\ER docs\ER-20027172_AD.docx") #Read in purple ER document into an array
purpleDocTable = []

#Lists for "compileErDocRow"
workingRow = []

#Arrays for compiling final lists
finalArray = []
finalCodeArray = []
finalDocArray = []

catTypes = []
smvTypes = []

flowLinearityTable = [] 




copyCodeFile(sensorList, mainCoeffList, coeffTable, catCoeffList, catTable, smvCoeffList, smvTable, constantsTable, catTypes, smvTypes)

addCatAndSmvTablesToCoeffTable(mainCoeffList, coeffTable, catCoeffList, catTable, smvCoeffList, smvTable, catTypes, smvTypes)

replaceVariablesWithDefinitions(coeffTable, constantsTable)


#print(mainCoeffList)
#for z in range(0, len(coeffTable)):
#    print(coeffTable[z])

#ER-20018334  /  Blue
copyExcelFile(oldBlueCoeffList, oldBlueTable, newBlueCoeffList, newBlueTable, blueFile, 
    blueTitleLineNum, blueDataLineNum) 

#ER-20015860  /  Red
copyExcelFile(coriolisRedCoeffList, coriolisRedTable, densViscRedCoeffList, densViscRedTable, 
    redFile, redCoeffLineNum, redCoeffDataNum) 


#ER-20015206 / Green
copyGreenErDoc(greenFile, greentableOne, greenTableTwo, greenTableThree, greenTableFour)

findFlowLinearityCoeffs(flowLinearityTable)


print()

#for z in range(0, len(flowLinearityTable)):
#    print(flowLinearityTable[z])

print()
print()

#ER-20027172 / Purple
purpleDocTable = copyPurpleErDoc(purpleFile)

#print(purpleDocTable)

createFinalArray(finalArray, COEFFS_TO_COMPARE, coeffTable, mainCoeffList, workingRow, finalCodeArray, newBlueCoeffList,
    newBlueTable, oldBlueCoeffList, oldBlueTable, coriolisRedCoeffList, coriolisRedTable, 
    densViscRedCoeffList, densViscRedTable, purpleDocTable, finalDocArray, greentableOne, greenTableFour)

flowLinearityMatch = compareflowLinearityTables(flowLinearityTable, greenTableTwo)

exportFinalArraytoExcelDocument(finalArray, flowLinearityTable, greenTableTwo, flowLinearityMatch)

print()

#for z in range(0, len(finalArray)):
#    print(finalArray[z])


print()
print()
print()


print("End")


