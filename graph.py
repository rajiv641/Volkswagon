import openpyxl
from datetime import date
import calendar
import xlrd as xl    
from openpyxl.chart import BarChart, Reference

  
wb_car = xl.open_workbook('LUN-Raw-Data.xlsx')                    #opening & reading the excel file
s1 = wb_car.sheet_by_index(1)                     #extracting the worksheet
s1.cell_value(0,0)                                   #initializing cell from the excel file mentioned through the cell position

todays_date = date.today()

wb = openpyxl.load_workbook('LUN-Raw-Data.xlsx')
sheet = wb['509-datastores-final']
#sheet = wb.active

# Data for plotting
values = Reference(sheet,
                    min_col=2,
                    max_col=2,
                    min_row=1,
                    max_row=s1.nrows)

cats = Reference(sheet, min_col=8, max_col=8, min_row=2, max_row=s1.nrows)

# Create object of BarChart class
chart = BarChart()
chart.add_data(values, titles_from_data=True)
chart.set_categories(cats)

# set the title of the chart
chart.title = "SAN Allocation by Device  " + str(calendar.month_name[todays_date.month]) + " " + str(todays_date.year)

# set the title of the x-axis
#chart.x_axis.title = "Products"

# set the title of the y-axis
#chart.y_axis.title = "Inventory per product"

# the top-left corner of the chart
# is anchored to cell F2 .
sheet.add_chart(chart,"A11")
values_stacked = Reference(sheet,
                            min_col=7,
                            max_col=7,
                            min_row=1,
                            max_row=s1.nrows)

cats_stacked = Reference(sheet, min_col=8, max_col=8, min_row=2, max_row=s1.nrows)

# Create object of BarChart class
chart = BarChart()
chart.type = 'col'
chart.grouping = "stacked"
chart.overlap = 100
chart.add_data(values_stacked, titles_from_data=True)
chart.set_categories(cats_stacked)

# set the title of the chart
chart.title = "SAN Capacity & Usage by Device - " + str(calendar.month_name[todays_date.month]) + " " + str(todays_date.year)



# the top-left corner of the chart
# is anchored to cell G2 .
sheet.add_chart(chart,"A47")
# save the file 
wb.save("LUN-Raw-Data.xlsx")
