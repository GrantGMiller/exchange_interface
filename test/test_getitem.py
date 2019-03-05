import creds as exchange_credentials
import exchange_interface
import datetime
exchange = exchange_interface.Exchange(
    username=exchange_credentials.username,
    password=exchange_credentials.password,
)
# exchange.UpdateCalendar(
#     startDT=datetime.datetime(year=2019, month=3, day=9),
#     endDT=datetime.datetime(year=2019, month=3, day=11),
# )
item = exchange.GetItem(
    'AAMkADkzMzBkZDQ5LTBlYjYtNDM1Yy05MjgwLTA0OTU0MDJiMTU1ZQBGAAAAAAAlV3FfZvuWS4G0Ngplked/BwDH6imC4of+QY0K3wQH6l6QAJ3UpDxVAABDb8fIgOIhSab0Z/mFRldyAAKMf7/dAAA='
)

print('item=', item)

exchange.DeleteEvent(item)