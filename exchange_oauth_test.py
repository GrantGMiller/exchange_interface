import time

import creds as exchange_credentials
import exchange_interface

exchange = exchange_interface.Exchange(
    # username=exchange_credentials.username,
    # password=exchange_credentials.password,
    # proxyAddress='172.17.16.79',
    # proxyPort='3128',
    # impersonation='mpower@extron.com'
    accessToken='eyJ0eXAiOiJKV1QiLCJub25jZSI6Im16ZWZma3lPVlZRcVZqblA4Zm5MT0NaVTlGRTczeEk4ZTZ4bU9qcWJzWFkiLCJhbGciOiJSUzI1NiIsIng1dCI6ImllX3FXQ1hoWHh0MXpJRXN1NGM3YWNRVkduNCIsImtpZCI6ImllX3FXQ1hoWHh0MXpJRXN1NGM3YWNRVkduNCJ9.eyJhdWQiOiJodHRwczovL291dGxvb2sub2ZmaWNlLmNvbSIsImlzcyI6Imh0dHBzOi8vc3RzLndpbmRvd3MubmV0LzMwZjE4Yzc4LWY3YWItNDg1MS05NzU5LTg4OTUwZTY1ZGM0Yi8iLCJpYXQiOjE1NjY0ODE4ODIsIm5iZiI6MTU2NjQ4MTg4MiwiZXhwIjoxNTY2NDg1NzgyLCJhY2N0IjowLCJhY3IiOiIxIiwiYWlvIjoiQVZRQXEvOE1BQUFBZkYrT21QclJ3VG95cm0rTTRKYjcrWldyVE94Mk03UDJvdHR5OWdvNGIwblU3UnZrWllZSXBickpxQytGeGVtRERndng3WWdGTmU3bHRIKzVkamtEZ1BtNjcvdjZEWmU0ZG5MUUFKRUMvNzg9IiwiYW1yIjpbInB3ZCIsIm1mYSJdLCJhcHBfZGlzcGxheW5hbWUiOiJFeHRyb25HUyIsImFwcGlkIjoiNDU5YWMyMWMtYWRkZS00NWE1LWFiZjMtODVkNzU3YmEyZmRmIiwiYXBwaWRhY3IiOiIwIiwiZW5mcG9saWRzIjpbXSwiZmFtaWx5X25hbWUiOiJNaWxsZXIiLCJnaXZlbl9uYW1lIjoiR3JhbnQiLCJpcGFkZHIiOiIxMi4yMTkuMTEzLjIiLCJuYW1lIjoiR3JhbnQgTWlsbGVyIiwib2lkIjoiZDdlYzIxZWMtNDFhMS00ZmExLTlmOTEtNTc1NDg5NjBmNjFkIiwib25wcmVtX3NpZCI6IlMtMS01LTIxLTQ4NTIxNTA2LTE1MDYzNTg3NTQtMjA5NDY4NjQyMi01MjkyMiIsInB1aWQiOiIxMDAzM0ZGRjhBQzBGRDQ4Iiwic2NwIjoiQ2FsZW5kYXJzLlJlYWRXcml0ZSBFV1MuQWNjZXNzQXNVc2VyLkFsbCBNYWlsLlJlYWQiLCJzaWQiOiI1ZmMwNTVlMC1jYzU1LTRlMzEtOWNiOC01Yzg0MzIwMTRhMGYiLCJzaWduaW5fc3RhdGUiOlsiaW5rbm93bm50d2siXSwic3ViIjoiMnY0MkZBeEpzbFM3WEpMWV9HcU5Kd3hGM3BuMkQxd1l2TTRWMFMxbXNuayIsInRpZCI6IjMwZjE4Yzc4LWY3YWItNDg1MS05NzU5LTg4OTUwZTY1ZGM0YiIsInVuaXF1ZV9uYW1lIjoiZ21pbGxlckBleHRyb24uY29tIiwidXBuIjoiZ21pbGxlckBleHRyb24uY29tIiwidXRpIjoiUzJaN1dUUS1rRW1iUU1lV3hHOFJBUSIsInZlciI6IjEuMCJ9.D1_OwEq5DU6Ih1d8ZlXdt6lpJqPkOipezJvhhTKzQscfMr9IOar33EgupKgL2rxEjfUto5auray9sa-6InSfMNuEXaP1mxXY53Fbia5vPKrfTEkSWltUMB-eLG43ss5rzzUgCp4sfkGcBmKnd7Y8OAmhl-7A40nQQ2dOdnUYtZiTSTkLmIrwbH6e9u36P1w4fv-uZhuI5q_ie5M5AO4YYAZyQuapOSCLHhA5Gp6R9oTPl3T4Aki2L0ShIkIr8tWMhMLApM2hVVT0lp0F_YrqGHpymNtqZwD18rmvVihzFWApr1EONFt_T9L4nVf4sA2zXr4EwX5M-Y9P3g2SddCqxA'
)

# exchange.UpdateCalendar('rnchallwaysignage1@extron.com')

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

print('nowItems=', items)

exchange.UpdateCalendar()
print('\n\n\n********************************************\n\n\n')
time.sleep(30)
print('37')
# exchange.UpdateCalendar()
# print('39')