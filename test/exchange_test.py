import creds
import exchange_interface



for resource in ['rnchallwaysignage1@extron.com', 'rnchallwaysignage2@extron.com']:
    exchange = exchange_interface.Exchange(
        username=creds.username,
        password=creds.password,
    )
    exchange.UpdateCalendar(resource)

    items = exchange.GetNextCalItems()
    print(resource, 'GetNextCalItems=', items)

    items = exchange.GetNowCalItems()
    print(resource, 'GetNowCalItems', items)
    print()