import datetime
import time
import json

import creds as exchange_credentials
import exchange_interface

exchange = exchange_interface.Exchange(
    username=exchange_credentials.username,
    password=exchange_credentials.password,
    # proxyAddress='172.17.16.79',
    # proxyPort='3128',
    # impersonation='rnchallwaysignage1@extron.com'
    # impersonation='rnchallwaysignage1@extron.com'
)


items = exchange.GetNextCalItems()


def NewCallback(cal, item):
    print('NewCallback(', cal, dict(item))


def ChangeCallback(cal, item):
    print('ChangeCallback(', cal, item)


def DeletedCallback(cal, item):
    print('DeletedCallback(', cal, item)


exchange.NewCalendarItem = NewCallback
exchange.CalendarItemChanged = ChangeCallback
exchange.CalendarItemDeleted = DeletedCallback

# print('nowItems=', items)

#exchange.UpdateCalendar()

print('\n\n\n********************************************\n\n\n')
# time.sleep(30)
print('37')
# exchange.UpdateCalendar()
# print('39')

startDT = datetime.datetime.now()
endDT = startDT + datetime.timedelta(days=7)
exchange.UpdateCalendar(startDT=startDT, endDT=endDT)
events = exchange.GetEventsInRange(startDT, endDT)

ret = []
for evt in events.copy():
    for key in ['Start', 'End']:
        evt = dict(evt)
        evt[key] = evt[key].isoformat()

    print('evt=', evt)
    ret.append(evt)

print('*************************\n\n')

print(json.dumps(ret, indent=2, sort_keys=True))

print('*************************\n\n')

print('end main.py')