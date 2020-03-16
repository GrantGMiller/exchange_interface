import datetime
import time

import creds as exchange_credentials
import exchange_interface

exchange = exchange_interface.Exchange(
    # username=exchange_credentials.username,
    # password=exchange_credentials.password,
    # impersonation='rnchallwaysignage1@extron.com'

    # proxyAddress='172.17.16.79',
    # proxyPort='3128',

    # username='NIDA3WFNConference@nih.gov',#working
    # password='Summertime2020!',
    # impersonation='NIDA-3WFN-HR-03C44@nih.gov'

    # username='CA_Roomscheduler@fresenius-kabi.com',#working
    # password='Winter2020!',
    # impersonation='CA_Galaxy_TOR_Pelee@fresenius-kabi.com'

    username='z-touchpanelno-confrm1.9@extron.com',
    password='Extron1025'
)


def NewCallback(cal, item):
    print('NewCallback(', cal, dict(item))


def ChangeCallback(cal, item):
    print('ChangeCallback(', cal, item)


def DeletedCallback(cal, item):
    print('DeletedCallback(', cal, item)


exchange.Connected = lambda _, state: print(state)
exchange.Disconnected = lambda _, state: print(state)
exchange.NewCalendarItem = NewCallback
exchange.CalendarItemChanged = ChangeCallback
exchange.CalendarItemDeleted = DeletedCallback

exchange.UpdateCalendar()

print('Now Events:', exchange.GetNowCalItems())

# exchange.CreateCalendarEvent(
#     subject='Test {})'.format(time.asctime()),
#     body='Test Body',
#     startDT=datetime.datetime.now()+datetime.timedelta(hours=1),
#     endDT=datetime.datetime.now() + datetime.timedelta(hours=2)
# )

while True:
    time.sleep(10)
    print('UpdateCalendar()')
    exchange.UpdateCalendar()
