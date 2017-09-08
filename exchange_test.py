import exchange_credentials
import exchange_interface

exchange = exchange_interface.Exchange(
    username=exchange_credentials.username,
    password=exchange_credentials.password,
)

exchange.UpdateCalendar('rnchallwaysignage1@extron.com')

items = exchange.GetNextCalItems()
print('items=', items)