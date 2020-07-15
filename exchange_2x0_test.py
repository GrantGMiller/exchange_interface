import datetime
import time
import requests
import creds as exchange_credentials
import exchange_interface_2x0

exchange = exchange_interface_2x0.EWS(
    # username=exchange_credentials.username,
    # password=exchange_credentials.password,
    # impersonation='rnchallwaysignage1@extron.com',

    # proxyAddress='172.17.16.79',
    # proxyPort='3128',

    # username='NIDA3WFNConference@nih.gov',#working
    # password='Summertime2020!',
    # impersonation='NIDA-3WFN-HR-03C44@nih.gov'

    # username='CA_Roomscheduler@fresenius-kabi.com', # working
    # username='CA_RoomerExtron@fresenius.onmicrosoft.com', # TTP Error 401: Unauthorized error
    #
    # username='CA_Roomscheduler@fresenius.onmicrosoft.com',
    # password='Winter2020!',
    # #
    # impersonation='CA_Galaxy_TOR_Pelee@fresenius-kabi.com',

    # impersonation='CA_Galaxy_TOR_Scotia@fresenius-kabi.com',
    # impersonation='CA_Galaxy_TOR_Laurentian@fresenius-kabi.com',
    # impersonation='abastos@calea.ca',
    # impersonation='CA_Galaxy_TOR_Pelee@fresenius.onmicrosoft.com',
    #
    # username='z-touchpanelno-confrm1.9@extron.com',
    # password='Extron1025',

    # username='CA_Galaxy_TOR_Laurentian@fresenius-kabi.com',

    # username='CA_Galaxy_TOR_Laurentian@fresenius.onmicrosoft.com',#working no impersonation
    # password='Soj44139',

    # username='CA_RoomerExtron_EU@fresenius.onmicrosoft.com',
    # password='Cac95373',
    # impersonation='CA_Galaxy_TOR_Laurentian@fresenius-kabi.com',
    #
    # username=exchange_credentials.dev_username,
    # password=exchange_credentials.dev_password,
    # impersonation='roomagenttestaccount@ExtronDev.com',
    #
    # username=exchange_credentials.username,
    # password=exchange_credentials.password,
    # impersonation='rnchallwaysignage1@extron.com',

    username=exchange_credentials.username,
    password=exchange_credentials.password,
    impersonation='usa-conf-pm2@extron.com',
)


def NewCallback(cal, item):
    print('NewCallback(', cal, dict(item))


def ChangeCallback(cal, item):
    print('ChangeCallback(', cal, item)


def DeletedCallback(cal, item):
    print('DeletedCallback(', cal, item)


def HandleConnected(_, state):
    # requests.post(
    #     url='https://hooks.slack.com/services/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',
    #     json={'text': state},
    # )
    pass


exchange.Connected = HandleConnected
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
