import pandas as pd
import numpy as np
import nltk
import os
import nltk.corpus
from datetime import date
inputfile=input("Enter input file name:")

df = pd.read_excel(inputfile, sheet_name=0)
casenumber=list(df['Case Number'])
caseowner=list(df['Case Owner'])
internaltitle=list(df['Internal Title'])
updatedate=list(df['Updated On'])
#casetitle=list(df['Title'])
updatedtimestamp=list(df['Last Outgoing Communication'])
#sap=list(df['Support Area Path'])

totalcasecount=dict((x,caseowner.count(x)) for x in set(caseowner))
idlecasescount={}
idlecases={}
percentage={}
metainfo={}

for k in range(len(caseowner)):
    idlecasescount[caseowner[k]]=0
    percentage[caseowner[k]]=0
    idlecases[caseowner[k]]=[]
   
    

#idlenesstrack
dateonly=[]
pastdate=[]
for i in range(len(updatedtimestamp)):
    temp=updatedtimestamp[i].strftime('%Y-%m-%d')
    dateonly.append(temp.split('-'))
    pastdate.append(date(int(dateonly[i][0]),int(dateonly[i][1]),int(dateonly[i][2])))

lastupdatedays=[]
today=date.today()


for i in range(len(casenumber)):
    lastupdatedays.append(today-pastdate[i])
    
    #print(today-pastdate[i])
cases10days=[]
cases5days=[]
cases15days=[]

for i in range(len(lastupdatedays)):
    if lastupdatedays[i].days >= 5 and lastupdatedays[i].days<10:
        #print(str(lastupdatedays[i])[0])
        cases5days.append(i)
        idlecases[caseowner[i]].append(casenumber[i])
        #idlecases[caseowner[i]].append(internaltitle[i])
        #idlecases[caseowner[i]].append(lastupdatedays[i].days)
        idlecasescount[caseowner[i]]=idlecasescount[caseowner[i]]+1
        
    if lastupdatedays[i].days >= 10 and lastupdatedays[i].days<15:
        #print(str(lastupdatedays[i])[0])
        cases10days.append(i)
        idlecases[caseowner[i]].append(casenumber[i])
        
        #idlecases[caseowner[i]].append(internaltitle[i])
        #idlecases[caseowner[i]].append(lastupdatedays[i].days)
        idlecasescount[caseowner[i]]=idlecasescount[caseowner[i]]+1
    if lastupdatedays[i].days >= 15:
        #print(str(lastupdatedays[i])[0])
        cases15days.append(i)
        #print(i)
        idlecases[caseowner[i]].append(casenumber[i])
        idlecasescount[caseowner[i]]=idlecasescount[caseowner[i]]+1
        


for key in totalcasecount:
    if key in idlecasescount:
        percentage[key]=((totalcasecount[key]-idlecasescount[key])/totalcasecount[key])*100
    else:
        idlecasescount[key]=0
        


import xlsxwriter

workbook = xlsxwriter.Workbook('result.xlsx')
table = workbook.add_worksheet("Table")
format1= workbook.add_format()
format1.set_bold()
format1.set_font_color('purple')
table.write(0,0,"Case Owner",format1)
table.write(0,1,"Number of idle cases",format1)
table.write(0,2,"Idle case SR",format1)
table.write(0,3,"Internal title",format1)
table.write(0,4,"Last outgoing email update",format1)
table.write(0,5,"Bin health",format1)
table.write(0,6,"Case updated on ",format1)


row=1
col=0
for key,value in idlecases.items():
    table.write(row,col,str(key))
    table.write(row,col+1,len(value))
    table.write(row,col+5,str(percentage[key]))
    #idlesr=value
    for i in range(len(value)):
        index=casenumber.index(value[i])
        table.write(row,col+2,str(value[i]))
        table.write(row,col+3,str(internaltitle[index]))
        #table.write(row,col+4,str(lastupdatedays[index].days))
        table.write(row,col+4,str(updatedtimestamp[index]))
        table.write(row,col+6,str(updatedate[index]))
        row=row+1

charts=workbook.add_worksheet('Charts')

charts.write(0,0,"Case Age")
charts.write(0,1,"No.of cases")
charts.write(1,0,">15 days")
charts.write(1,1,int(len(cases15days)))
charts.write(2,0,">10 days")
charts.write(2,1,int(len(cases10days)))
charts.write(3,0,">5 days")
charts.write(3,1,int(len(cases5days)))

agechart=workbook.add_chart({'type':'pie'})
agechart.add_series({
    'name':       'Age category',
    'categories': ['Charts', 1, 0, 3, 0],
    'values':     ['Charts', 1, 1, 3, 1],
})
agechart.set_title ({'name': 'Case Age Numbers'})
 

# Set an Excel chart style.
agechart.set_style(3)
charts.insert_chart('F3', agechart,
     {'x_offset': 20, 'y_offset': 5})


#idlelessengineerwise
charts.write(0,3,"Engineer")
#charts.write(0,4,"Total cases")
charts.write(0,5,"Idle cases")

charts.write_column("D2",idlecasescount.keys())
charts.write_column("F2",idlecasescount.values())
c=1
for key,value in idlecasescount.items():
    charts.write(c,4,totalcasecount[key])
    c=c+1

cochart=workbook.add_chart({'type':'bar'})

"""cochart.add_series({
    'name':       '=Charts!$E$1',
    'categories': '=Charts!$D$2:$D$'+str(len(idlecasescount.keys())+1),
    'values':     '=Charts!$E$2:$E$'+str(len(idlecasescount.keys())+1),
})"""
cochart.add_series({
    'name':       '=Charts!$F$1',
    'categories': '=Charts!$D$2:$D$'+str(len(idlecasescount.keys())+1),
    'values':     '=Charts!$F$2:$F$'+str(len(idlecasescount.keys())+1),
})
cochart.set_title ({'name': 'Idle cases Engineer wise'})
 
# Add x-axis label
cochart.set_x_axis({'name': 'Engineer'})
 
# Add y-axis label
cochart.set_y_axis({'name': 'Cases count'})
 
# Set an Excel chart style.
cochart.set_style(4)
 
# add chart to the worksheet
# the top-left corner of a chart
# is anchored to cell E2 .
charts.insert_chart('A15', cochart)

workbook.close()

"""fivedays=len(cases5days)
tendays=len(cases10days)
fifteendays=len(cases15days)

print(fivedays)
print(tendays)
print(fifteendays)
"""




