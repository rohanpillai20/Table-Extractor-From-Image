#Import Packages
import os
import copy
import pandas as pd
from pandas import ExcelWriter
from openpyxl import load_workbook

#Defining Class Structure
class struct:
    def __init__(self):
        self.param = []
        self.text = None

#Initialization
#Input Directory. Place the JSON files here.
inpDir = r'C:\Users\Rohan\Desktop\JSON to Excel (Bounding Box)\Input - JSON'
#Output Path. The Excel will be generated here.
exportPath = r'C:\Users\Rohan\Desktop\JSON to Excel (Bounding Box)\Output - Excel\PythonExport.xlsx'

keyword = "boundingBox"
keyword1 = "words"
difference = 10
differenceVertical = 45

#Reading file by file
def readfile() :
    try:
        for root, dirs, filenames in os.walk(inpDir):
            for f in filenames :
                print('\033[1m' + "For", f + '\033[0m')
                file1 = (os.path.join(inpDir, f))            
                readtext(file1,f)
                print("")
    except Exception as e:
        print("Problem in readfile():", e)

#Reading a file line by line            
def readtext(file1,f):
    try:
        
        allParams = []

        with open(file1) as ff:        
            lines = ff.readlines()        
            for i in range(0, len(lines)):
                line = lines[i]
                
                #Check for 'boundingBox and words in line & 11th line
                if (keyword in line) and (keyword1 in lines[i+11]):

                    list = []
                    index = 1
                    tempStruct = struct()

                    while True:
                        if lines[i+index].strip().replace(",","").isdecimal():
                            list.append(int(lines[i+index].strip().replace(",","")))
                        else:
                            break
                        index = index + 1

                    tempStruct.param = list;
                    tempStruct.text = lines[i+10].strip().replace("\"text\": ","").replace(",","")
                    allParams.append(tempStruct)
                    
    except Exception as e:
        print("Error in readtext():", e)

    horizontalAccess(allParams,f)

#Segregate into rows    
def horizontalAccess(allParams,f):
    try:
        
        #Print the list of values read by Microsoft Azure Computer Vision API
        for i in allParams:
            print(i.param," : ", i.text)
        
        #The total number of values read
        totalCount = len(allParams)
        print("Number of Parameters read: ", totalCount)
        
        #List that will store all values row-wise
        relationList=[]
        
        #Create a temporary list and store it with the first element
        tempList = [allParams[0]]
        
        #Loop check for relation in y-coordinates
        for index in range(0, len(allParams)-1):
            relation = 0

            if (allParams[index].param[1]<=allParams[index+1].param[1] and allParams[index].param[3]>=allParams[index+1].param[3]) or  (allParams[index].param[1]<=allParams[index+1].param[1] and allParams[index].param[3]>=allParams[index+1].param[3]):
                relation = relation + 1
            elif max(abs(allParams[index].param[1]-allParams[index+1].param[1]), abs(allParams[index].param[3]-allParams[index+1].param[3])) < difference:
                relation = relation + 1

            if relation>0:
                tempList.append(allParams[index+1])

                if index == len(allParams)-2:
                    relationList.append(tempList)
            else:
                relationList.append(tempList)

                tempList = [allParams[index+1]]

    except Exception as e:
        print("Exception in horizontalAccess():", e)
        
    verticalAccess(relationList,f)

#Segregate into columns
def verticalAccess(relationList,f):
    try:
        
        #Variable stores the count of number of rows
        noOfRows = len(relationList)
        
        #Null in case a value does not exist
        null = struct()
        null.param = [0 for i in range(0, len(relationList[0][0].param))]
        null.text="-"
        
        #Loop for counting the maximum number of columns in the table
        noOfMaxCols = 0
        for i in range(0, noOfRows):
            if len(relationList[i])>noOfMaxCols:
                noOfMaxCols = len(relationList[i])
        
        #Initialize a list with max number of columns and number of rows
        dataFrame = [[null for x in range(0,noOfMaxCols)] for y in range(0,noOfRows)]
        for i in range(0, len(relationList[0])):
            dataFrame[0][i] = relationList[0][i]
        
        #Check for relation in x-coordinates
        index = 0
        for t in range(1, noOfRows):
            index2 = 0
            for index1 in range(0, len(relationList[0])):
                for index2 in range(0, len(relationList[index+t])):
                    relation = 0

                    if (relationList[index][index1].param[0]<=relationList[index+t][index2].param[0] and relationList[index][index1].param[2]>=relationList[index+t][index2].param[2]) or  (relationList[index+t][index2].param[0]<=relationList[index][index1].param[0] and relationList[index+t][index2].param[2]>=relationList[index][index1].param[2]):
                        relation = relation + 1
                    elif min(abs(relationList[index][index1].param[0]-relationList[index+t][index2].param[0]), abs(relationList[index][index1].param[2]-relationList[index+t][index2].param[2])) < differenceVertical:
                        relation = relation + 1

                    if relation>0:
                        del dataFrame[t][relationList[0].index(relationList[index][index1])]
                        dataFrame[t].insert(relationList[0].index(relationList[index][index1]) ,relationList[index+t][index2])
                        break
                        
    except Exception as e:
        print("Exception in verticalAccess():", e)
    
    exportInExcel(dataFrame,noOfRows,f)

#Print the output in an Excel file
def exportInExcel(dataFrame,noOfRows,f):
    try:
        #Copy the text in a list
        a = copy.copy(dataFrame)
        for i in range(0, noOfRows):
            for j in range(0, len(dataFrame[i])):
                if(dataFrame[i][j]=='-'):
                    a[i][j]='-'
                else:
                    a[i][j]=dataFrame[i][j].text
        print(' ')
        #Convert the list into a pandas.DataFrame
        df = pd.DataFrame(a)
        print(df)
        
        if os.path.exists(exportPath):
            #If any excel sheet exists
            book = load_workbook(exportPath)
            writer = pd.ExcelWriter(exportPath, engine='openpyxl') 
            writer.book = book
            writer.sheets = dict((ws.title, ws) for ws in book.worksheets)
            df.to_excel(writer, f)
            writer.save()
        else:
            #If it doesn't, then create a new file
            writer = ExcelWriter(exportPath)
            df.to_excel(writer,f)
            writer.save()
            
    except Exception as e:
        print("Exception in exportInExcel():", e)


#Main Function
if __name__ == '__main__':
    try:
        readfile()
    except Exception as e:
        print("Problem in Main:", e)   
