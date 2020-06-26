import datetime

import time

from exchange_interface_2x0 import EWS
import webbrowser
import ctypes
from oauth_tools import AuthManager

MessageBox = ctypes.windll.user32.MessageBoxW

ewsServiceAccount = EWS(

    # gm has ApplicationImpersonation
    username='gm_service_account@extrondemo.com',
    impersonation='rf_a101@extrondemo.com',
    password='Extron1025',

)

thisUser = dict()
thisUser['clientID'] = '33f9b88c-2182-4e7e-b2eb-660111ebe6c2'
thisUser['tenantID'] = '79c56eae-3556-4ad9-b028-02f1a46cdf6e'
thisUser['email'] = 'gm_service_account@extrondemo.com'
thisUser['impersonation'] = 'rf_a101@extrondemo.com'
thisUser['type'] = 'Microsoft'

authManager = AuthManager(
    microsoftClientID=thisUser['clientID'],
    microsoftTenantID=thisUser['tenantID'],
)

user = authManager.GetUserByID(thisUser['email'])
print('user=', user)
if user is None:
    print('No user exists for ID="{}"'.format(thisUser['email']))

    d = authManager.CreateNewUser(thisUser['email'], authType=thisUser['type'])
    webbrowser.open(d.get('verification_uri'))

    MessageBox(None, f'{d["user_code"]}\r\n\r\nUse account {thisUser["email"]}', 'Enter this Code', 0)
    print('User Code=', d.get('user_code'))

while True:
    user = authManager.GetUserByID(thisUser['email'])
    if user is None:
        time.sleep(1)
    else:
        break
print('user=', user)

ewsServiceAccountOauth = EWS(

    # gm has ApplicationImpersonation
    username='gm_service_account@extrondemo.com',
    impersonation='rf_a101@extrondemo.com',
    password='Extron1025',
    authType='Oauth',
    oauthCallback=user.GetAcessToken

)

ewsInstances = [ewsServiceAccount, ewsServiceAccountOauth]


def DoToAllEWS(methodName, *args, **kwargs):
    ret = {}
    for ews in ewsInstances:
        method = getattr(ews, methodName)
        ret[ews] = method(*args, **kwargs)

    return ret


def test_CreateEvent():
    newEventSubject = f'Subject {int(time.time())}'

    startDT = datetime.datetime.now()
    endDT = startDT + datetime.timedelta(minutes=15)

    body = f'Body {int(time.time())}'

    DoToAllEWS(
        'CreateCalendarEvent',
        subject=newEventSubject,
        startDT=startDT,
        endDT=endDT,
        body=body,
    )

    DoToAllEWS(
        'UpdateCalendar',
    )

    results = DoToAllEWS(
        'GetAllEvents',
    )

    for ews, res in results.items():
        for event in res:
            if event.Get('Subject') == newEventSubject:
                break # found the event, pass
        else:
            raise Exception('Could not find newly created event')
