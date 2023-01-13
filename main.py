import pandas as pd
import os
import glob

import numpy
import xlsxwriter
# use glob to get all the csv files
# in the folder

path = os.getcwd()+"\DE1"
xls_de1 = glob.glob(os.path.join(path, "DE1*.xls"))
path = os.getcwd()+"\DE2"
xls_de2 = glob.glob(os.path.join(path, "DE2*.xls"))

def calculatedata(xls_de1,engine):
    summarized_data = pd.DataFrame()

    percylinderdata = pd.DataFrame()
    initialnaming=0
    labelname=pd.DataFrame()
    ringindex = ["Ring 1","Ring 2","Ring 3","Ring 4"]
    cylinderlabels = ["Cyl " + str(i) for i in range(1,13)]
    # loop over the list files for Diesel 1

    for f in xls_de1:
        
        # read the xls file
        df = pd.read_excel(f, 'Sheet1', skiprows = 23, nrows=4,  usecols= 'G:R',header=None,names=cylinderlabels, index_col=None)
        df.index =ringindex #name the index based on ring number
        hours = pd.read_excel(f, 'Sheet1', index_col=None, usecols = "I", header = 4, nrows=0)      
        hours=hours.columns.values[0]
        
        date = pd.read_excel(f, 'Sheet1', index_col=None, usecols = "E", header = 4, nrows=0)      
        date=date.columns.values[0]
        labelentry=[]
        labelentry.append("Date:")
        labelentry.append(date)
        labelentry.append("Hours:")
        labelentry.append(hours)
        labelentry.append("Engine:")
        labelentry.append(engine)
        for i in range(6):
            labelentry.append("")
        emptyentry=[]
        for i in range(12):
            emptyentry.append([""])

        labelentry = pd.DataFrame(labelentry).T
        labelentry.columns=cylinderlabels
        labelentry.dropna(inplace = True)
        emptyentry = pd.DataFrame(emptyentry).T
        emptyentry.columns=cylinderlabels
        emptyentry.dropna(inplace = True)
        entry1 = pd.concat([labelentry])
        entry1 = pd.concat([entry1,df])
        entry1 = pd.concat([entry1,emptyentry])
        summarized_data = pd.concat([summarized_data,entry1]) # εμφανιση συγκεντρωτικων δεδομενων για καθε μετρηση

        historicaldata = df.values.tolist()
        cylinder=[]
        column=pd.DataFrame()
        for j in range(12):
            cylinder.append([])
        for i in range(4):
            for j in range(12):
                cylinder[j].append(historicaldata[i][j])
        for j in range(12):
            i = j+1
            timelabel = pd.Series(hours)
            if initialnaming == 0:
                labelname = pd.Series(cylinderlabels[j])
            else:
                labelname = pd.Series("")
            column = pd.concat([column,labelname,timelabel,df["Cyl " + str(i)]],axis=0)
        initialnaming=1
        percylinderdata= pd.concat([percylinderdata,column],axis=1)
    return percylinderdata,summarized_data
    
            
    
#print(summarized_data)
# loop over the list files for Diesel 2
percylinderdata,summarized_data=calculatedata(xls_de1,"DE 1")
percylinderdata1,summarized_data1=calculatedata(xls_de2,"DE 2")      
bwriter = pd.ExcelWriter('Exported_Data.xlsx', engine='xlsxwriter')
summarized_data.to_excel(bwriter, sheet_name='DE1 Δεδομένα ανά μέτρηση')
percylinderdata.to_excel(bwriter, sheet_name='DE1 Δεδομένα ανά κύλινδρο')
summarized_data1.to_excel(bwriter, sheet_name='DE2 Δεδομένα ανά μέτρηση')
percylinderdata1.to_excel(bwriter, sheet_name='DE2 Δεδομένα ανά κύλινδρο')
bwriter.save()
