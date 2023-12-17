import openpyxl
import datetime
# Load the Excel workbook
workbook = openpyxl.load_workbook('Attendance.xlsx')

# Select the worksheet
worksheet = workbook['Students Attendance']
count=0
# Traverse the first column and print out the values
for cell in worksheet.iter_cols(min_col=3, max_col=5, values_only=True):
    for value in cell:
        count+=1
        if value=='Kathan Gabani':
            print(value)
            print(count)
            d=datetime.date.today()
            print(d.strftime("%d"))
            count1=5
            while(True):
                if str(count1-4)==d.strftime("%d"):
                    print(count1)
                    print(count,count1)
                    asc=chr(65+count1)+str(count)
                    print(asc)
                    worksheet[asc]="P"
                    workbook.save('Attendance.xlsx')
                    break
                else:
                    count1+=1

                
                
                
            

