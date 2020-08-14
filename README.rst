Extron Global Scripter allows access to O365 accounts.

Example Script
==============

::

    import requests
    from exchange_interface import EWS
    import datetime

    ews = EWS(
        username='rf_a105@extrondemo.com',
        password='Extron123!',
    )
    ews.Connected = lambda _, state: print('EWS', state)
    ews.Disconnected = lambda _, state: print('EWS', state)
    ews.NewCalendarItem = lambda _, item: print('NewCalendarItem(', item)
    ews.CalendarItemChanged = lambda _, item: print('CalendarItemChanged(', item)
    ews.CalendarItemDeleted = lambda _, item: print('CalendarItemDeleted(', item)

    ews.UpdateCalendar()

    print('Events happending now=', ews.GetNowCalItems())

    print('Event(s) happening next=', ews.GetNextCalItems())

    nowDT = datetime.datetime.now()
    nowPlus24hrs = nowDT + datetime.timedelta(days=1)

    print('Events happening in the next 24 hours=', ews.GetEventsInRange(startDT=nowDT, endDT=nowPlus24hrs))

    # You can create a new event like this:
    ews.CreateCalendarEvent(
        subjec='Test Subject',
        body='Test Body',
        startDT=datetime.datetime.now(),
        endDT=datetime.datetime.now() + datetime.timedelta(hours=1),
        )



