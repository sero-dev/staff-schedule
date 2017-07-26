import openpyxl
import os

path = os.getcwd();

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

def convertTime(week, day):
      if not (week[day].value == None or week[day].value == "OFF"):
          start = getStart(week[day].value.rsplit('-', 1)[0])
          end = getEnd(week[day].value.rsplit('-', 1)[1])

          print(start, '-', end)

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

          print(eventInfo)

week = openpyxl.load_workbook('week.xlsx')
week = week.get_sheet_by_name('Sheet1')

convertTime(week,'E16')
convertTime(week,'G16')
convertTime(week,'I16')
convertTime(week,'K16')
convertTime(week,'M16')
convertTime(week,'O16')
