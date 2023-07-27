import xlrd
import re
import os
import os.path
import xlsxwriter
import pandas as pd
import xlwt
from xlwt import easyxf, Workbook
import csv
from xls2xlsx import XLS2XLSX
from datetime import date

colour_code = {'green':0x0B,'yellow':0x0D,'red':0x0A}
agency_code = {'397542':"Haymarket",'397587':"Haymarket",'504601':"VW-IT-Shared",'1003635':'skoda','530145':'VWPC','648554':'Crimson','444767':'ReadingRoom','469634':'Blueprint','471519':'Driver Connex'}

def style_pattern(colour):
    style = xlwt.XFStyle()
    font = xlwt.Font()
    font.bold = True
    style.font = font
    pattern = xlwt.Pattern()
    pattern.pattern = xlwt.Pattern.SOLID_PATTERN
    pattern.pattern_fore_colour =  colour_code[colour] #yellow 0x0D ##0x0B (bright green)##
    style.pattern = pattern
    return style

book_car = xlwt.Workbook(encoding = "utf-8")
def workbook_car(name):
    style = style_pattern('yellow')
    sheet_car = book_car.add_sheet(name)
    sheet_car.write(0,0,"DatastoreName",style=style)
    sheet_car.write(0,1,"CapacityGB",style=style)
    sheet_car.write(0,2,"FreeSpaceGB",style=style)
    sheet_car.write(0,3,"PercentReserved",style=style)
    sheet_car.write(0,4,"AvailableSpaceGB",style=style)
    sheet_car.write(0,6,"Datastore",style=style)
    sheet_car.write(0,7,"Utilisation",style=style)
    sheet_car.write(0,8,"Capacity (GB)",style=style)
    sheet_car.write(0,9,"Used (GB)",style=style)
    sheet_car.write(0,10,"Raw Free",style=style)
    sheet_car.write(0,11,"Reserved",style=style)
    sheet_car.write(0,12,"Availabled (GB)",style=style)
    return sheet_car

def workbook_graph(name):
    style = style_pattern('yellow')
    sheet_graph = book_car.add_sheet(name + "-final")
    sheet_graph.write(0,0,"Datastore",style=style)
    sheet_graph.write(0,1,"Utilisation",style=style)
    sheet_graph.write(0,2,"Capacity (GB)",style=style)
    sheet_graph.write(0,3,"Used (GB)",style=style)
    sheet_graph.write(0,4,"Raw Free",style=style)
    sheet_graph.write(0,5,"Reserved",style=style)
    sheet_graph.write(0,6,"Availabled (GB)",style=style)
    sheet_graph.write(0,7,"Agency/App",style=style)
    return sheet_graph


def worksheet_data(worksheet_car,excel_count,datastorename,capacity,freespace,PercentReserved,AvailableSpaceGB,usedgb,utilization,reserved,availablegb):
    style_percent = easyxf(num_format_str='0.00%')
    worksheet_car.write(excel_count,0,datastorename)
    worksheet_car.write(excel_count,1,float(capacity))
    worksheet_car.write(excel_count,2,freespace)
    worksheet_car.write(excel_count,3,float(PercentReserved))
    worksheet_car.write(excel_count,4,float(AvailableSpaceGB))
    worksheet_car.write(excel_count,6,datastorename)
    worksheet_car.write(excel_count,8,float(capacity))
    worksheet_car.write(excel_count,9,usedgb)
    worksheet_car.write(excel_count,10,freespace)
    worksheet_car.write(excel_count,7,utilization,style_percent)
    worksheet_car.write(excel_count,11,reserved)
    worksheet_car.write(excel_count,12,round(availablegb,2))

    excel_count = excel_count + 1
    return excel_count

def worksheet_graph_data(worksheet_graph,excel_graph_count,datastorename,utilization,capacitygb,usedgb,freespace,reserved,availablegb,agency):
    style_percent = easyxf(num_format_str='0.00%')
    worksheet_graph.write(excel_graph_count,0,datastorename)
    worksheet_graph.write(excel_graph_count,1,utilization,style_percent)
    worksheet_graph.write(excel_graph_count,2,float(capacitygb))
    worksheet_graph.write(excel_graph_count,3,float(usedgb))
    worksheet_graph.write(excel_graph_count,4,float(freespace))
    worksheet_graph.write(excel_graph_count,5,reserved)
    worksheet_graph.write(excel_graph_count,6,round(availablegb,2))
    worksheet_graph.write(excel_graph_count,7,agency)
    excel_graph_count = excel_graph_count + 1
    return excel_graph_count


#device_file = input('Enter the Excel DeviceList file :  ')
#if (os.path.exists(device_file) != True):
#    print (format(device_file) + " doesn't exist, Exiting.....")
#    exit(-1)

device_file = '509-datastores.csv'
#x1 = pd.ExcelFile(loc)
#sheet_devicelist = x2.parse()

#for row in sheet_devicelist:
#    print (row)

with open(device_file) as csv_file:
    nam = device_file.split('.')
    print (nam[0])
    worksheet_car = workbook_car(nam[0])
    worksheet_graph = workbook_graph(nam[0])
    csv_reader = csv.reader(csv_file, delimiter=',')
    line_count = 0
    excel_count = 1
    excel_graph_count = 1
    usedgb = 0
    utilization = 0
    reserved = 0
    capacitygb = 0
    temp_agency_code = []
    for row in csv_reader:
        if (line_count != 0): 
            if (re.search("local",row[3])):
                pass
            else:
                capacitygb = float(row[4])
                freespace = float(row[5])
                #print ("capacitygb freespace : " + str(capacitygb) + " " + str(freespace))
                usedgb = capacitygb - freespace
                utilization = round((usedgb/(float(row[4])-float(row[4])*(float(row[6])/100))),5)
                if (capacitygb*(float(row[6])/100) < freespace ):
                    reserved = (capacitygb*(float(row[6])/100))
                    availablegb = freespace - reserved
                    excel_count = worksheet_data(worksheet_car,excel_count,row[3],capacitygb,freespace,row[6],row[7],usedgb,utilization,reserved,availablegb)
                    temp_agency_code = row[3].split('-')
                    a_code = temp_agency_code[0]
                    print (a_code)
                    excel_graph_count = worksheet_graph_data(worksheet_graph,excel_graph_count,row[3],utilization,capacitygb,usedgb,freespace,reserved,availablegb,agency_code[a_code])    
                else:
                    reserved = freespace
                    availablegb= freespace - reserved
                    excel_count = worksheet_data(worksheet_car,excel_count,row[3],row[4],freespace,row[6],row[7],usedgb,utilization,reserved,availablegb)
                    excel_graph_count = worksheet_graph_data(worksheet_graph,excel_graph_count,row[3],utilization,capacitygb,usedgb,freespace,reserved,availablegb,agency,agency_code[a_code])    
                #print (row[3])
        line_count = line_count + 1

book_car.save('LUN-Raw-Data.xls')

x2x = XLS2XLSX("LUN-Raw-Data.xls")
x2x.to_xlsx("LUN-Raw-Data.xlsx")

ret = os.system("python3 graph.py")
if (ret == 0):
    print ("Graphs generated successfully")
