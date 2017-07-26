import insertEvent
import openpyxl

months = {  "January" : "01",
            "February" : "02",
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

# Converts start time to military time and adds seconds to the end
def getStart(period):
    hour = int(period.rsplit(':', 1)[0])

    if(hour > 7 and hour < 13):
        return str(hour) + ':' + period.rsplit(':', 1)[1] + ':00'
    else:
        return str(hour + 12) + ':' + period.rsplit(':', 1)[1] + ':00'

# Converts end time to military time and adds seconds to the end
def getEnd(period):
    hour = int(period.rsplit(':', 1)[0])

    if(hour > 9 and hour < 13):
        return str(hour) + ':' + period.rsplit(':', 1)[1] + ':00'
    else:
        return str(hour + 12) + ':' + period.rsplit(':', 1)[1] + ':00'

# Converts time in standard time format for Google Calendar API, then books event
def convertTime(week, day):
      if week[day].value != None and week[day].value != "OFF":
          date = week[day[0] + '5'].value.rsplit(". ", 1)[1]
          if(date[0] == " ") :
              date = date[1:3]

          start = getStart(week[day].value.rsplit('-', 1)[0])
          end = getEnd(week[day].value.rsplit('-', 1)[1])

          # Adds 0 in front of single digit hours
          if(start[0] != '1') :
              start = "0" + start

          # Adds 0 in front of single digit hours
          if(end[0] != '1') :
              end = "0" + end

          # Decides whether to use the first month or the second month
          if(int(date) < 7) :
              month = month2
          else :
              month = month1

          # Sets up correct time format for Google Calendar API JSON file
          start = year + "-" + months[month] + "-" + date + "T" + start + "-04:00"
          end = year + "-" + months[month] + "-" + date + "T" + end + "-04:00"

          # Event JSON file
          eventInfo = {
            'summary': 'QBPL Cyber Center',
            'location': 'Queens Library (Central) 89-11 Merrick Blvd, Jamaica, NY 11432',
            'start': {
              'dateTime': start,
              'timeZone': 'America/New_York',
            },
            'end': {
              'dateTime': end,
              'timeZone': 'America/New_York',
            },
            'reminders': {
              'useDefault': False,
              'overrides': [
                {'method': 'popup', 'minutes': 60},
              ],
            },
          }

          # Creates event for Google Calendar
          insertEvent.bookEvent(eventInfo)
          print("Event created: Success!")

week = openpyxl.load_workbook('week.xlsx')
week = week.get_sheet_by_name('Sheet1')

# Gets year and months from workbook
year = week['A3'].value.rsplit(', ', 1)[1]
month1 = week['A3'].value.rsplit(' - ', 1)[0]
month2 = week['A3'].value.rsplit(' - ', 1)[1].rsplit(', ', 1)[0]

# Trims numbers at the end of the month
month1 = month1[:-3]
month2 = month2[:-3]

convertTime(week,'C16')
convertTime(week,'E16')
convertTime(week,'G16')
convertTime(week,'I16')
convertTime(week,'K16')
convertTime(week,'M16')
convertTime(week,'O16')
