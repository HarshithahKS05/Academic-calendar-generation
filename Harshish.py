import datetime
import pandas as pd
import sys
import os.path

print('*****************************************************************************************************')
print('                             SRI VENKATESWARA COLLEGE OF ENGINEERING                                 ')
print('                                     SRIPERUMBUDUR-602117                                            ')
print('*****************************************************************************************************')
print()
print()
print(" Hi! This is HARSHISH, I am here to help you with Academic Calendar Generation.  ")
print()
print('======================================================================================================')
print('                                  INSTRUCTIONS                                                        ')
print('======================================================================================================')
print()
print('=> The program requires two files (Harshish.py and calendar.xls).')
print('=> Please ensure that your system has Python with pandas and xlrd modules installed.')
print('=> Ensure calendar.xls file exits in the same folder as this program file. ')
print('=> Please do not change the sheet names in the Excel file.')
print('=> Subject code in the sheets - Time_Table_ and Static_Data  can be changed according to Departments.')
print('=> For each subject enter the required time in minutes in the sheet Static_Data.')
print()
print('======================================================================================================')
print()
varTemp=input('Press any key to continue...')
print()
print()

#getting section name of class from the user  
curPath=os.getcwd()
excelFile=curPath + '\calendar.xlsx'
#excelFile='C:\Ishwarya\calendar\calendar.xlsx'       
if(os.path.exists(excelFile)):
    #print(excelFile,' file exist')
	varTemp=0
else:
    print(excelFile,' does not exist')
    sys.exit()


#getting section from the user
section_str=input('Enter the Section Name (Example: A or B or C ) : ')
#print(section_str)
secName=section_str.upper()
if(secName.isalpha()==True):
    if(len(secName.strip())>1):
        print('Invalid data. Section name should be in 1 character in length (Example: A or B or C ). Quitting.. ')
        sys.exit()
else:
    print('Invalid data. Section name should be in alphabet. Quitting..')
    sys.exit()       
   
       

shHolidays='Holidays'
shTimeTable='Time_Table_'+secName
shStaticData='Static_Data'
xls = pd.ExcelFile(excelFile)
sheetNames=xls.sheet_names

   
   
#checking for the sheets in the file.    
    
sampleVar1=0
sampleVar2=0
sampleVar3=0
for sheetName in sheetNames:
    if (sheetName==shStaticData):
        #print('The Static Data sheet exist')
        sampleVar1=1
    elif(sheetName==shTimeTable):
        #print('The Time Table sheet  exist')
        sampleVar2=1
    elif(sheetName==shHolidays):
        #print('The Holiday sheet exist')
        sampleVar3=1    

if (sampleVar1==0):
    print(shStaticData,' sheet does not exist')
    sys.exit()
if(sampleVar2==0):
    print(shTimeTable,' sheet does not exist')
    sys.exit()
if(sampleVar3==0):
    print(shHolidays,' sheet does not exist')
    sys.exit()

# Open holiday sheet.
sheet1=xls.parse(shHolidays)
# count the number of rows in holiday sheets.
countHrow=sheet1.shape[0]
# Open Timetable sheet.
sheet2=xls.parse(shTimeTable)
# count the number of rows in timetable sheet.
countTTrow=sheet2.shape[0]
#count the number of columns in timetable sheet
countTTcol=sheet2.shape[1]
# Open static data sheet.
sheet3=xls.parse(shStaticData)
# count the number of rows in static data sheet.
countSD=sheet3.shape[0]

#declare variables for maximum class hours for all 7 subjects and set them to zero value initially
minPerClass=0
maxS0=0
maxS1=0
maxS2=0
maxS3=0
maxS4=0
maxS5=0
maxS6=0
maxLib=0
maxLibrary=0
maxL1=0
maxL2=0
totWorkingDays=0
totWeekEnds=0
totHolidays=0
totDays=0
totCAT1days=0
totCAT2days=0
totCAT3days=0
minPerClassStr='PerPeriod'    
    
#create a list for verifying subjects.
subjectList=[]
varTemp=0
#x is a looping variable for subjects.
x=0
while x < countSD:
    sub_xx=sheet3['Subject'][x]
    if(sub_xx ==minPerClassStr ):
        #print(sub_xx , ' is not an actual subject')
        varTemp=0
    else:
    
        subjectList.insert(x,sub_xx)
    x+=1
#print('List of Subjects: ',subjectList)    

# loop through each subject in Time Table for each day and verify whether it exists in SubjectList(). If not, throw error and quit

j=0
varTemp=0
while j < countTTrow:
    l=1
    # loop through each column in the time table sheet for the given day name (Wednesday)
    while l < countTTcol:
        # j = day name row which is like Wednesday
        # l = column for each subject for that period. You need to loop through each column
        sub = sheet2.loc[j][l]
        #match for each subject in the time table sheet
        if(sub in subjectList):
            #print(sub,' exists. No problem.')
            varTemp=1
        else:
            print('Subject [', sub, '] does not exist in the sheet [', shStaticData , ']')
            sys.exit()
        l+=1
    j+=1

#Getting the start date from the user.   
       
date_format = '"%d/%m/%Y'
userInput = input("Enter the Start Date in DD/MM/YYYY format: ") 
try: # strptime throws an exception if the input doesn't match the pattern
    start_date = datetime.datetime.strptime(userInput, "%d/%m/%Y")
    orgStartDate = start_date
except:
    print ("Invalid Date Format. Quiting... \n")
    raise ValueError("Invalid Date Format. Quiting... \n")

    
#Getting the end date from the user.
    
userInput = input ("Enter the End Date in DD/MM/YYYY format: ")
try: # strptime throws an exception if the input doesn't match the pattern
    end_date = datetime.datetime.strptime(userInput, "%d/%m/%Y")
except:
    print ("Invalid Date Format. Quiting... \n")
    raise ValueError("Invalid Date Format. Quiting... \n")
   
    
# get the start date for CAT 1 exam
print()
userInput = input("Enter the Start Date for CAT1 Exams in DD/MM/YYYY format: ") 
try: # strptime throws an exception if the input doesn't match the pattern
    cat1StartDate = datetime.datetime.strptime(userInput, "%d/%m/%Y")
except:
    print ("Invalid Date Format. Quiting... \n")
    raise ValueError("Invalid Date Format. Quiting... \n")
# get the end date for CAT 1 exam
userInput = input("Enter the End Date for CAT1 Exams in DD/MM/YYYY format: ") 
try: # strptime throws an exception if the input doesn't match the pattern
    cat1EndDate = datetime.datetime.strptime(userInput, "%d/%m/%Y")
except:
    print ("Invalid Date Format. Quiting... \n")
    raise ValueError("Invalid Date Format. Quiting... \n")
    
    
# get the start date for CAT 2 exam
print()
userInput = input("Enter the Start Date for CAT 2 Exams in DD/MM/YYYY format: ") 
try: # strptime throws an exception if the input doesn't match the pattern
    cat2StartDate = datetime.datetime.strptime(userInput, "%d/%m/%Y")
except:
    print ("Invalid Date Format. Quiting... \n")
    raise ValueError("Invalid Date Format. Quiting... \n")

    
# get the end date for CAT 2 exam

userInput = input("Enter the End Date for CAT 2 Exams in DD/MM/YYYY format: ") 
try: # strptime throws an exception if the input doesn't match the pattern
    cat2EndDate = datetime.datetime.strptime(userInput, "%d/%m/%Y")
except:
    print ("Invalid Date Format. Quiting... \n")
    raise ValueError("Invalid Date Format. Quiting... \n")


# Variable h contains the total contact hours.
h=0
while h < countSD:
    sd=sheet3['Subject'][h]
    if (sd==minPerClassStr):
        #find the minmum minutes required for each class - i.e. 50 mins stored in static_data sheet
        #.loc[row][col]  -- here row =k and col=1 => required_minutes
        minPerClass=sheet3.loc[h][1]
    elif (sd==subjectList[0]):
        max0 = sheet3.loc[h][1]
        #print(max1)
    elif (sd==subjectList[1]):
        max1 = sheet3.loc[h][1]
        #print(max2)
    elif (sd==subjectList[2]):
        max2 = sheet3.loc[h][1]
        #print(max3)
    elif (sd==subjectList[3]):
        max3 = sheet3.loc[h][1]
        #print(max4)
    elif (sd==subjectList[4]):
        max4 = sheet3.loc[h][1]
        #print(max5)
    elif (sd==subjectList[5]):
        max5 = sheet3.loc[h][1]
        #print(max6)
    elif (sd==subjectList[6]):
        max6 = sheet3.loc[h][1]
        #print(max7)
    elif (sd==subjectList[7]):
        maxLab1 = sheet3.loc[h][1]
        #print(maxLab1)
    elif (sd==subjectList[8]):
        maxLab2 = sheet3.loc[h][1]
        #print(maxLab2)
    elif (sd==subjectList[9]):
        maxLibrary = sheet3.loc[h][1]
        #print(maxLibrary)    
    else:
        print('Ignoring minimum mins required for subject ' ,sd)
    # max hours for each subject is assigned to the respective variables like maxS1, maxS2, etc
    h+=1
    continue    

# variable k is for Static Data sheet. Store minimum required time minutes for each subject from the statis data sheet
k=0
while k < countSD:
    sd=sheet3['Subject'][k]
    if (sd==minPerClassStr):
        #find the minmum minutes required for each class - i.e. 50 mins stored in static_data sheet
        #.loc[row][col]  -- here row =k and col=1 => required_minutes
        minPerClass=sheet3.loc[k][1]
    elif (sd==subjectList[0]):
        maxS0 = sheet3.loc[k][1]
    elif (sd==subjectList[1]):
        maxS1 = sheet3.loc[k][1]
    elif (sd==subjectList[2]):
        maxS2 = sheet3.loc[k][1]
    elif (sd==subjectList[3]):
        maxS3 = sheet3.loc[k][1]
    elif (sd==subjectList[4]):
        maxS4 = sheet3.loc[k][1]
    elif (sd==subjectList[5]):
        maxS5 = sheet3.loc[k][1]
    elif (sd==subjectList[6]):
        maxS6 = sheet3.loc[k][1]
    elif (sd==subjectList[7]):
        maxL1 = sheet3.loc[k][1]
    elif (sd==subjectList[8]):
        maxL2 = sheet3.loc[k][1]
    elif (sd==subjectList[9]):
        maxLib = sheet3.loc[k][1]    
    else:
        varTemp=1
        #print('Ignoring minimum mins required for subject ' ,sd)
    # max hours for each subject is assigned to the respective variables like maxS1, maxS2, etc
    k+=1
    continue    

    
#create a list for the holidays

holidayList=[]

# loop through the Holidays sheet and insert all subjects into the list created above
i=0
while i < countHrow:
    hol=sheet1['Date'][i]
    # holiday date should be within semester start date and end date
    if((hol >= start_date) and (hol <= end_date)):
        # holiday date should be outside CAT 1 start and end date. Otherwise, it is counted as CAT1 days
        if(hol >= cat1StartDate and hol <= cat1EndDate ):
            varTemp=1
            #print('Holiday [', hol, '] falls within CAT1 start and end dates')
        else:
            # holiday date should be outside CAT 2 start and end date. Otherwise, it is counted as CAT2 days
            if(hol >= cat2StartDate and hol <= cat2EndDate ):
                varTemp=1
                #print('Holiday [', hol, '] falls within CAT2 start and end dates')
            else:
                holidayList.insert(i,hol)
                totHolidays= totHolidays+1
                #print(holidayList)
    else:
        varTemp=1
        #print('Holiday [', hol, '] falls outside Semester start and end dates')
    i+=1

holidayList.sort()
#print('Holiday List = ',holidayList)

# Loop in  for weekends, weekdays and CAT exams

while start_date <= end_date:
    strday_name = start_date.strftime("%A")
    if(start_date in holidayList):
        varTemp=1
        #print(start_date, '(', strday_name, ') is in Holiday List')
    #name of the current day being processed.variable name is start_date.    
    elif (strday_name=='Saturday'or strday_name=='Sunday'):
        totWeekEnds= totWeekEnds + 1
        #print(start_date, '(', strday_name, ') is Weekend')    
    elif (start_date >= cat1StartDate and start_date <= cat1EndDate ):
        totCAT1days=totCAT1days + 1
        #print(start_date, ' (',strday_name, ') falls within CAT 1 Exam Dates')
    elif (start_date >= cat2StartDate and start_date <= cat2EndDate ):
        totCAT2days=totCAT2days + 1
        #print(start_date,  ' (',strday_name, ') falls within CAT 2 Exam Dates')
    else:
        #print(start_date, ' (', strday_name, ') is a working day' )
        totWorkingDays = totWorkingDays+1
        # Variable j is loop through the days of a week. For ex..Monday, Tuesday..etc.
        j=0
        while j < countTTrow:
            # in the sheet time table, look for column Day and find the particular day matching the processing Date
            tt=sheet2['Day'][j]
            if (tt==strday_name):
                # it is not a holiday but a working day so that you can exit the holiday loop
                # if you find matching row for the day, then find out the subjects for 7 periods in 7 columns for the day
                # Variable l is to loop for periods.
                l=0
                # loop through each column in the time table sheet for the given day name (Wednesday)
                while l < countTTcol:
                        # j = day name row which is like Wednesday
                        # l = column for each subject for that period. You need to loop through each column
                        sub = sheet2.loc[j][l]
                        #match for each subject in the time table sheet
                        if (sub==subjectList[0]):
                            #print(start_date,' subtracting S1 -  (', maxS1 , ' - ', minPerClass, ')')
                            maxS0 = maxS0 - minPerClass
                        elif (sub==subjectList[1]):
                            #print(start_date,' max=',maxS2,' minperclass=',minPerClass)
                            maxS1 = maxS1 - minPerClass
                        elif (sub==subjectList[2]):
                            maxS2 = maxS2 - minPerClass
                        elif (sub==subjectList[3]):
                            maxS3 = maxS3 - minPerClass
                        elif (sub==subjectList[4]):
                            maxS4 = maxS4 - minPerClass
                        elif (sub==subjectList[5]):
                            maxS5 = maxS5 - minPerClass
                        elif (sub==subjectList[6]):
                            #print(start_date,' max=',maxS7,' minperclass=',minPerClass)
                            maxS6 = maxS6 - minPerClass
                        elif (sub==subjectList[7]):
                            maxL1 = maxL1 - minPerClass
                        elif (sub==subjectList[8]):
                            maxL2 = maxL2 - minPerClass
                        elif (sub==subjectList[9]):
                            maxLib = maxLib - minPerClass    
                        l+=1
            j+=1
    start_date+= datetime.timedelta(days=1)
    totDays = totDays + 1
    #print('done for ',start_date)
print()
print()
print('totWorkingDays=',totWorkingDays)
print()
print('==============================================================')
print('                          REPORT for Section ', secName, '                  ')
print('==============================================================')
print('Start Date: ',orgStartDate)
print('End Date  : ',end_date )
print('Total Days : ',totDays)
print()
print('Total Days of Week ends (Sat & Sun) : ',totWeekEnds)
print('Total Declared Holidays : ', totHolidays)
print()
print('Total Exam Days for CAT 1 : ', totCAT1days)
print('Total Exam Days for CAT 2 : ', totCAT2days)
print()
print('Total Working Days = ',(totDays-(totWeekEnds+totHolidays+totCAT1days+totCAT2days)))
print('==============================================================')
print()
print()
print('Number of classes scheduled as per Time Table for each subject:')
print('---------------------------------------------------------------')
print(subjectList[0],'=',(max0-maxS0)/50)
print(subjectList[1],'=',(max1-maxS1)/50)
print(subjectList[2],'=',(max2-maxS2)/50)
print(subjectList[3],'=',(max3-maxS3)/50)
print(subjectList[4],'=',(max4-maxS4)/50)
print(subjectList[5],'=',(max5-maxS5)/50)
print(subjectList[6],'=',(max6-maxS6)/50)
print(subjectList[7],'=',(maxLab1-maxL1)/50)
print(subjectList[8],'=',(maxLab2-maxL2)/50)
print(subjectList[9],'=',(maxLibrary-maxLib)/50)
print('--------------------------------------------------------------')
print()
print()
print('=============================================================')
print('                     Summary for Section ',secName,'                       ')
print('=============================================================')
print()
print('Subject:          Target:       Current:             Balance:')
print('--------          -------       --------             --------')
print(subjectList[0].ljust(15),  max0,'\t\t', max0-maxS0 ,'\t\t\t',  maxS0)   
print(subjectList[1].ljust(15),  max1,'\t\t',  max1-maxS1,'\t\t\t',  maxS1)
print(subjectList[2].ljust(15),  max2,'\t\t',  max2-maxS2,'\t\t\t',  maxS2)   
print(subjectList[3].ljust(15),  max3,'\t\t',  max3-maxS3,'\t\t\t',  maxS3)   
print(subjectList[4].ljust(15),  max4,'\t\t',  max4-maxS4,'\t\t\t',  maxS4 )   
print(subjectList[5].ljust(15),  max5,'\t\t',  max5-maxS5,'\t\t\t',  maxS5 )   
print(subjectList[6].ljust(15),  max6,'\t\t',  max6-maxS6,'\t\t\t',  maxS6)  
print(subjectList[7].ljust(15),  maxLab1,'\t\t',      maxLab1-maxL1,'\t\t\t',      maxL1)  
print(subjectList[8].ljust(15),  maxLab2,'\t\t',      maxLab2-maxL2,'\t\t\t',      maxL2)
print(subjectList[9].ljust(15),  maxLibrary,'\t\t',   maxLibrary-maxLib,'\t\t\t',  maxLib)

print()
#print('Note: Time for each class is', minPerClass, " minutes")
print('===============================================================')
print('End of Report')
print('Excel File used - ',excelFile)
print('Thank you for using this program - By HarshIsh')
