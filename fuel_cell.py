#This code is for the excel data obtained from channel-6 of arbin.
import xlrd,sys
import xlwt
import time
#input the excel file from arbin
print "input the name of excel file needed to be treated"
data_file=raw_input('>')

#open excel file and get voltage and curret value
data=xlrd.open_workbook(data_file+'.xls')
table=data.sheet_by_name(u'Channel_1-006')
voltage=table.col_values(7)
current=table.col_values(6)

# initilization
voltage_average=[0]*1000
current_average=[0]*1000
abs_current_density=[0]*1000
voltage_eR=[0]*1000
voltage_fR=[0]*1000
ocv_eR=[0]*1000
ocv_fR=[0]*1000
OCV=0
j=0

#the total number of data in the current/voltage
NJ=len(current)

#select useful data
for i in range(20,NJ):
    if abs(voltage[i]-voltage[i-1])>0.005:
        voltage_average[j]=(voltage[i-1]+voltage[i-2]+voltage[i-3])/3
        current_average[j]=(current[i-1]+current[i-2]+current[i-3])/3
        j+=1

#create excel file of data_treated,sheet and write the header of every column
data_treated=xlwt.Workbook(encoding = 'ascii')
table_treated=data_treated.add_sheet('voltage_vs_current')
table_treated.write(0,0,'current')
table_treated.write(0,1,'voltage')
table_treated.write(0,2,'abs_current_area')
table_treated.write(0,3,'voltage_eR')
table_treated.write(0,4,'voltage_fR')
table_treated.write(0,5,'overpotential_eR')
table_treated.write(0,6,'overpotential_fR')
table_treated.write(0,7,'electronic_resistance')
table_treated.write(0,8,'full_resistance')

#calculate the electronic and iR corrected voltage
print "Please input the electronic resistance,unit is ohm"
eR=float(raw_input('>'))
print "Please input the full resistance,unit is ohm"
fR=float(raw_input('>'))
for i in range(j+1):
    voltage_eR[i]=voltage_average[i]-eR*current_average[i]
    voltage_fR[i]=voltage_average[i]-fR*current_average[i]

#calculate the absolute value for the current density
print "please input the area of the electrode,unit is cm2"
area=float(raw_input('>'))
for i in range(j+1):
    abs_current_density[i]=abs(current_average[i])/area

#calculate the overpotential
i=0
while True and i<=j:
    if current_average[i]==0:
        OCV=voltage_average[i]
    ocv_eR[i]=voltage_eR[i]-OCV
    ocv_fR[i]=voltage_fR[i]-OCV
    i+=1
    
#input the data to the excel
for i in range(j+1):
    table_treated.write(i+1,0,current_average[i])
    table_treated.write(i+1,1,voltage_average[i])
    table_treated.write(i+1,2,abs_current_density[i])
    table_treated.write(i+1,3,voltage_eR[i])
    table_treated.write(i+1,4,voltage_fR[i])
    table_treated.write(i+1,5,ocv_eR[i])
    table_treated.write(i+1,6,ocv_fR[i])
table_treated.write(1,7,eR)
table_treated.write(2,7,'ohm')
table_treated.write(1,8,fR)
table_treated.write(2,8,'ohm')

#comment for the experiment
table_comment=data_treated.add_sheet('comment')
print "Please input the comment for the expriment"
comment=raw_input('>')
table_comment.write(0,0,comment)

#save the treated file in the name you want
print '''please input the file name you want to use,the default one will be
date of today'''
file_name_after_treated=raw_input('>')
date=time.strftime('%Y-%m-%d',time.localtime(time.time()))
data_treated.save(file_name_after_treated+'_'+date+'.xls')
