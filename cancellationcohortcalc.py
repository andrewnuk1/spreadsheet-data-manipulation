import xlrd
import xlwt
from xlutils.copy import copy
import datetime

"""
For each housebuilder:
Opens an excel file with weekly unit data (data in columns)
Reads the units for sale
Compares each weekly list to determine those units entering the sales process
Looks for such units that are relisted (i.e. cancelled and then relisted in later time periods)
Builds a cancellation rate for each week and the cancellation cohort for each week
Therefore we can see how often buyers are cancelling their purchases - idea of health of the housing market
"""

row_count = 0

def calculate_cancellations(path, col_adjust, homebuilder):
    global row_count
    workbook = xlrd.open_workbook(path)
    sheet_open = workbook.sheet_by_index(2)

    #detect number and location of columns to read
    #should always be under the date in the excel spreadsheet
    count = 0
    i = 0
    columns_to_read = []
    while True:
        if sheet_open.cell_value(0,i)== "end":
            break
        elif sheet_open.cell_value(0,i) <> "":
            columns_to_read.append(i)
            count = count +1
        i = i + 1

    column_distance = columns_to_read[1] - columns_to_read[0]

    #go through each column and read the data
    i = 0
    data_dict = {}
    for i in range(len(columns_to_read)):
        ii = 0
        data_to_read = []

        for rownum in range(0,sheet_open.nrows):
            if sheet_open.cell_type(rowx=rownum, colx=(column_distance)*i) in (xlrd.XL_CELL_EMPTY, xlrd.XL_CELL_BLANK):
                break
            data_to_read.append(sheet_open.cell_value(rownum,(column_distance)*i + col_adjust))
        #the date of the column is always at position 0 in the list
        #hence, need to force it in if the excel data has the date in another column
        if col_adjust == 1:
            data_to_read[0] = sheet_open.cell_value(0,(column_distance)*i)
        data_dict[i] = data_to_read

        
    # build the list of items sold in all weeks
    sold = []
    for i in range(len(columns_to_read)-1):
        sold.append([item for item in data_dict[i] if item not in data_dict[i+1]])


    # for the items sold in week x, trace how many are left in each following week
    sold_week_lists = []

    for weeki in range(len(columns_to_read)-2):
        sold_week_now = []
        sold_temp = sold[weeki]
        for i in range(weeki,len(columns_to_read)-1-1):
            sold_week_now.append([item for item in sold_temp if item not in data_dict[i+1+1]])
            sold_temp = sold_week_now[i-weeki]


        # store the result in a new list of lists
        sold_week_lists.append(sold_week_now)


        # print on screen - not needed for the program to work
        print(len(sold_week_lists))
        print(len(sold_week_lists[weeki][-1]))
        print(sold_week_lists[weeki][0][0])
        print((round(len(sold[weeki]),2) - len(sold_week_lists[weeki][-1]))*100/(len(sold[weeki])))



    #write the output to a spreadsheet
    if homebuilder == 0:
        workbook = xlwt.Workbook()
        sheet = workbook.add_sheet('cohorts')
    else:
        rb = xlrd.open_workbook('outputcohorts20180607.xls')
        workbook = copy(rb)
        sheet = workbook.get_sheet(0)
    
    sheet.write(0 +(row_count),0,path)
    sheet.write(1 +(row_count),0,"cohorts")
    for i in range(len(columns_to_read)-2):
        sheet.write(2 + (row_count),i, ((datetime.datetime(1899,12,30) + datetime.timedelta(days=int(sold_week_lists[i][0][0]))).strftime('%d-%m-%y'))*1)    

    for i in range(len(columns_to_read)-2):
        for ii in range(len(columns_to_read)-2-i):
            sheet.write(3 + (row_count)+i,ii+i,round((round(len(sold[i]),2) - len(sold_week_lists[i][ii]))*100/(len(sold[i])),2))

    sheet.write(3 + (row_count)+len(columns_to_read)-2 + 1,0,"Final cancellation rate")
    for i in range(len(columns_to_read)-2):
        sheet.write(3 + (row_count)+len(columns_to_read)-2 + 2,i, ((datetime.datetime(1899,12,30) + datetime.timedelta(days=int(sold_week_lists[i][0][0]))).strftime('%d-%m-%y'))*1)
        sheet.write(3 + (row_count)+len(columns_to_read)-2 + 3,i,round((round(len(sold[i]),2) - len(sold_week_lists[i][-1]))*100/(len(sold[i])),2))
        
    row_count = 3 + (row_count)+len(columns_to_read)-2 + 3 + 4
    workbook.save('outputcohorts20180607.xls')



#location of excel spreadsheet
path = [
"C:\\Users\\andrew\\Documents\\Big Data Project\\Persimmon\\SummaryPersimmon1.xlsx",
"C:\\Users\\andrew\\Documents\\Big Data Project\\CharlesChurch\\SummaryCharlesChurch1.xlsx",
"C:\\Users\\andrew\\Documents\\Big Data Project\\CrestNicholson\\SummaryCrestNicholson1.xlsx",
"C:\\Users\\andrew\\Documents\\Big Data Project\\TaylorWimpey\\SummaryTaylorWimpey1.xlsx",]

col_adjust = [0,0,0,1]


#go through each homebuilder to run the calculation
for homebuilder in range(len(path)):
    calculate_cancellations(path[homebuilder],col_adjust[homebuilder], homebuilder)
    
