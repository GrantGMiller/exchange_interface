import creds as exchange_credentials
import exchange_interface

exchange = exchange_interface.Exchange(
    username=exchange_credentials.username,
    password=exchange_credentials.password,
)

# exchange.UpdateCalendar('rnchallwaysignage1@extron.com')

items = exchange.GetNextCalItems()


def NewCallback(cal, item):
    print('NewCallback(', cal, item)


def ChangeCallback(cal, item):
    print('ChangeCallback(', cal, item)


def DeletedCallback(cal, item):
    print('DeletedCallback(', cal, item)


exchange.NewCalendarItem = NewCallback
exchange.CalendarItemChanged = ChangeCallback
exchange.CalendarItemDeleted = DeletedCallback

print('nowItems=', items)

exchange.UpdateCalendar()