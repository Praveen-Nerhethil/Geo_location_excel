# importing geopy library
from geopy.geocoders import Nominatim
import xlrd


import xlsxwriter
 
workbook1= xlsxwriter.Workbook('RealOutput.xlsx') # getting output excel name
worksheet = workbook1.add_worksheet()
workbook2= xlsxwriter.Workbook('Reamins.xlsx') # not getting output excel name
worksheet1 = workbook2.add_worksheet()

row = 0
col = 0

row1=0
col1=0
fpath=(r"path_of_input_excel") # path of input excel name

workbook = xlrd.open_workbook(fpath)
excel_sheet = workbook.sheet_by_index(0)
excel_sheet.cell_value(0, 0)
real_count=1
tot_count =0
for i in range(excel_sheet.nrows):
    tot_count+=1
    print(tot_count)
    try:
       
        loc = Nominatim(user_agent="GetLoc")
        branch = excel_sheet.cell_value(i, 1)
        getLoc = loc.geocode(branch)
        print("Latitude = ", getLoc.latitude, "\n")
        print("Longitude = ", getLoc.longitude)
        print("Get")
        worksheet.write(row, col, branch)
        worksheet.write(row, col + 1, getLoc.latitude)
        worksheet.write(row, col + 2, getLoc.longitude)
        row += 1
        real_count +=1
        print(real_count)
              
    except Exception as e:
        branch = excel_sheet.cell_value(i, 0)
        print(branch)
        worksheet1.write(row1, col1, branch)
        row1 += 1
   
workbook1.close()
workbook2.close()