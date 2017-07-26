import insertEvent
import openpyxl
import os

path = os.getcwd();

months = {  "January" : "01",
            "Feburary" : "02",
            "March" : "03",
            "April" : "04",
            "May" : "05",
            "June" : "06",
            "July" : "07",
            "August" : "08",
            "September" : "09",
            "October" : "10",
            "November" : "11",
            "December" : "12" }

def getStart(period):
    hour = int(period.rsplit(':', 1)[0])

    if(hour > 7 and hour < 13):
        return str(hour) + ':' + period.rsplit(':', 1)[1] + ':00'
    else:
        return str(hour + 12) + ':' + period.rsplit(':', 1)[1] + ':00'

def getEnd(period):
    hour = int(period.rsplit(':', 1)[0])

    if(hour > 9 and hour < 13):
        return str(hour) + ':' + period.rsplit(':', 1)[1] + ':00'
    else:
        return str(hour + 12) + ':' + period.rsplit(':', 1)[1] + ':00'

def convertTime(week, day, weekday):
      if not (week[day].value == None or week[day].value == "OFF"):

          date = week[day[0] + '5'].value.rsplit(". ", 1)[1]
          if(date[0] == " ") :
              date = date[1:3]
          print(date)

          start = getStart(week[day].value.rsplit('-', 1)[0])
          end = getEnd(week[day].value.rsplit('-', 1)[1])

          start = year + "-" + months[month1] + "-" + date + "T" + start + "-04:00"
          end = year + "-" + months[month1] + "-" + date + "T" + end + "-04:00"

          print(start)
          print(end)

          eventInfo = {
            'summary': 'QBPL Cyber Center',
            'location': 'Queens Library (Central) 89-11 Merrick Blvd, Jamaica, NY 11432',
            'start': {
              'dateTime': '2017-07-26T' + start + '-04:00',
              'timeZone': 'America/New_York',
            },
            'end': {
              'dateTime': '2017-07-26T' + end + '-04:00',
              'timeZone': 'America/New_York',
            },
            'reminders': {
              'useDefault': False,
              'overrides': [
                {'method': 'popup', 'minutes': 60},
              ],
            },
          }

          insertEvent.bookEvent(eventInfo)

week = openpyxl.load_workbook('week.xlsx')
week = week.get_sheet_by_name('Sheet1')

year = week['A3'].value.rsplit(', ', 1)[1]
month1 = week['A3'].value.rsplit(' - ', 1)[0]
month2 = week['A3'].value.rsplit(' - ', 1)[1].rsplit(', ', 1)[0]

month1 = month1[:-3]
month2 = month2[:-3]

convertTime(week,'C16', 0)
convertTime(week,'E16', 1)
convertTime(week,'G16', 2)
convertTime(week,'I16', 3)
convertTime(week,'K16', 4)
convertTime(week,'M16', 5)
convertTime(week,'O16', 6)
