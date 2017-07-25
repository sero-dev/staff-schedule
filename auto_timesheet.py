import openpyxl
import os

path = os.getcwd();

def getStart(period):
    if(int(period) < 12):
        return(' AM')
    else:
        return(' PM')

def getEnd(period):
    if(int(period) > 9 and int(period) < 12):
        return (' AM')
    else:
        return (' PM')

def convertTime(week, from):
      if(not(week[from].value == None or week[from].value == "OFF")):
     = (week[from].value.rsplit('-', 1)[0]
          + getPeriod1(week[from].value.rsplit('-', 1)[0].rsplit(':', 1)[0]))
     = (week[from].value.rsplit('-', 1)[1]                            # Change 'to' in a way that moves over two columns
          + getPeriod2(week[from].value.rsplit('-', 1)[1].rsplit(':', 1)[0]))

w1 = openpyxl.load_workbook('week.xlsx')
w1 = w1.get_sheet_by_name('Sheet1')

convertTime(w1,'C16')
convertTime(w1,'E16')
convertTime(w1,'G16')
convertTime(w1,'I16')
convertTime(w1,'K16')
convertTime(w1,'M16')
convertTime(w1,'O16')
