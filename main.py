import insertEvent

eventInfo = {
  'summary': 'QBPL Cyber Center',
  'location': 'Queens Library (Central) 89-11 Merrick Blvd, Jamaica, NY 11432',
  'start': {
    'dateTime': '2017-07-26T08:30:00-04:00',
    'timeZone': 'America/New_York',
  },
  'end': {
    'dateTime': '2017-07-26T16:30:00-04:00',
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
