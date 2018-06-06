import creds
import exchange_interface
import time

print('time.asctime=', time.asctime())


#for resource in [ 'rnchallwaysignage1@extron.com']: #'rnchallwaysignage1@extron.com',
for resource in [ 'rnchallwaysignage2@extron.com']: #'rnchallwaysignage1@extron.com',
    exchange = exchange_interface.Exchange(
        username=creds.username,
        password=creds.password,
    )
    exchange.UpdateCalendar(resource)

    nextItems = exchange.GetNextCalItems()
    print(resource, 'GetNextCalItems=', nextItems)

    nowItems = exchange.GetNowCalItems()
    print(resource, 'GetNowCalItems', nowItems)
    print()

    for item in nowItems:
        if item.HasAttachments():
            for a in item.GetAttachments():
                print('a=', a)
                print('a.GetFilename=', a.GetFilename())