import creds
import exchange_interface

exchange = exchange_interface.Exchange(
    username=creds.username,
    password=creds.password,
)

exchange.UpdateCalendar()#'rnchallwaysignage1@extron.com')

items = exchange.GetNextCalItems()
print('GetNextCalItems=', items)

items = exchange.GetNowCalItems()
print('GetNowCalItems', items)