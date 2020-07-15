import datetime
import re
import time
from requests_ntlm import HttpNtlmAuth
import requests
from calendar_base import _BaseCalendar, _CalendarItem

# debug requests - https://duckduckgo.com/?q=python+requests+debugging&t=ffab&atb=v137-1&ia=web&iax=qa
import logging

import http.client as http_client

http_client.HTTPConnection.debuglevel = 1
logging.basicConfig()
logging.getLogger().setLevel(logging.DEBUG)
requests_log = logging.getLogger("requests.packages.urllib3")
requests_log.setLevel(logging.DEBUG)
requests_log.propagate = True

TZ_NAME = time.tzname[0]
if TZ_NAME == 'EST':
    TZ_NAME = 'Eastern Standard Time'
elif TZ_NAME == 'PST':
    TZ_NAME = 'Pacific Standard Time'
elif TZ_NAME == 'CST':
    TZ_NAME = 'Central Standard Time'
########################################

RE_CAL_ITEM = re.compile('<t:CalendarItem>[\w\W]*?<\/t:CalendarItem>')
RE_ITEM_ID = re.compile(
    '<t:ItemId Id="(.*?)" ChangeKey="(.*?)"/>'
)  # group(1) = itemID, group(2) = changeKey #within a CalendarItem
RE_SUBJECT = re.compile('<t:Subject>(.*?)</t:Subject>')  # within a CalendarItem
RE_HAS_ATTACHMENTS = re.compile('<t:HasAttachments>(.{4,5})</t:HasAttachments>')  # within a CalendarItem
RE_ORGANIZER = re.compile(
    '<t:Organizer>.*<t:Name>(.*?)</t:Name>.*</t:Organizer>'
)  # group(1)=Name #within a CalendarItem
RE_START_TIME = re.compile('<t:Start>(.*?)</t:Start>')  # group(1) = start time string #within a CalendarItem
RE_END_TIME = re.compile('<t:End>(.*?)</t:End>')  # group(1) = end time string #within a CalendarItem
RE_HTML_BODY = re.compile('<t:Body BodyType="HTML">([\w\W]*)</t:Body>', re.IGNORECASE)

RE_EMAIL_ADDRESS = re.compile('.*?\@.*?\..*?')

RE_ERROR_CLASS = re.compile('ResponseClass="Error"', re.IGNORECASE)
RE_ERROR_MESSAGE = re.compile('<m:MessageText>([\w\W]*)</m:MessageText>')


class EWS(_BaseCalendar):
    def __init__(
            self,
            username=None,
            password=None,
            impersonation=None,
            myTimezoneName=None,
            serverURL=None,
            authType='Basic',  # also accept "NTLM" and "Oauth"
            ntlmDomain=None,
            oauthCallback=None,  # callable, takes no args, returns Oauth token
            apiVersion='Exchange2007_SP1',  # TLS uses "Exchange2007_SP1"
            verifyCerts=True,
            debug=False,
    ):
        super().__init__()
        self._username = username
        self._password = password
        self._impersonation = impersonation
        self._serverURL = serverURL
        self._authType = authType
        self._oauthCallback = oauthCallback
        self._apiVersion = apiVersion
        self._verifyCerts = verifyCerts
        self._debug = debug

        thisMachineTimezoneName = time.tzname[0]
        if thisMachineTimezoneName == 'EST':
            thisMachineTimezoneName = 'Eastern Standard Time'
        elif thisMachineTimezoneName == 'PST':
            thisMachineTimezoneName = 'Pacific Standard Time'
        elif thisMachineTimezoneName == 'CST':
            thisMachineTimezoneName = 'Central Standard Time'

        self._myTimezoneName = myTimezoneName or thisMachineTimezoneName
        print('myTimezoneName=', self._myTimezoneName)

        self._session = requests.session()

        self._session.headers['Content-Type'] = 'text/xml'

        if authType == 'Basic':
            self._session.auth = requests.auth.HTTPBasicAuth(self._username, self._password)
        elif authType == 'NTLM':
            self._session.auth = HttpNtlmAuth(f'{ntlmDomain}\\{self._username.split("@")[0]}', self._password)
        elif authType == 'Oauth':
            pass  # we will put the accessToken into each DoRequest
        else:
            raise TypeError('Unknown Authorization Type')

        self._useImpersonationIfAvailable = True
        self._useDistinguishedFolderMailbox = False

    @property
    def Impersonation(self):
        return self._impersonation

    @Impersonation.setter
    def Impersonation(self, newImpersonation):
        self._impersonation = newImpersonation

    def GetEvents(self, startDT=None, endDT=None):
        # Default is to return events from (now-1days) to (now+7days)
        startDT = startDT or datetime.datetime.utcnow() - datetime.timedelta(days=1)
        endDT = endDT or datetime.datetime.utcnow() + datetime.timedelta(days=7)

        startTimestring = ConvertDatetimeToTimeString(startDT)
        endTimestring = ConvertDatetimeToTimeString(endDT)

        parentFolder = f'''
            <t:DistinguishedFolderId Id="calendar"/>
        '''

        soapBody = f'''
            <m:FindItem Traversal="Shallow">
            <m:ItemShape>
                <t:BaseShape>IdOnly</t:BaseShape>
                <t:AdditionalProperties>
                    <t:FieldURI FieldURI="item:Subject" />
                    <t:FieldURI FieldURI="calendar:Start" />
                    <t:FieldURI FieldURI="calendar:End" />
                    <t:FieldURI FieldURI="item:Body" />
                    <t:FieldURI FieldURI="calendar:Organizer" />
                    <t:FieldURI FieldURI="calendar:RequiredAttendees" />
                    <t:FieldURI FieldURI="calendar:OptionalAttendees" />
                    <t:FieldURI FieldURI="item:HasAttachments" />
                    <t:FieldURI FieldURI="item:Sensitivity" />
                </t:AdditionalProperties>
            </m:ItemShape>
            <m:CalendarView 
                MaxEntriesReturned="100" 
                StartDate="{startTimestring}" 
                EndDate="{endTimestring}" 
                />
            <m:ParentFolderIds>
                {parentFolder}
            </m:ParentFolderIds>
        </m:FindItem>
        '''
        self._DoRequest(soapBody)

    def _DoRequest(self, soapBody):
        # API_VERSION = 'Exchange2013'
        # API_VERSION = 'Exchange2007_SP1'

        if self._impersonation and self._useImpersonationIfAvailable:
            # Note: Don't add a namespace to the <ExchangeImpersonation> and <ConnectingSID> tags
            # This will cause a "You don't have permission to impersonate this account" error.
            # Don't ask my why.
            # UPDATE: removing the namespace makes this work for licensed accounts, but not for service accounts with impersonation, so now i really dont understand whats going on

            # https://docs.microsoft.com/en-us/exchange/client-developer/exchange-web-services/how-to-identify-the-account-to-impersonate
            soapHeader = f'''
                <t:RequestServerVersion Version="{self._apiVersion}" />
                <t:ExchangeImpersonation>
                    <t:ConnectingSID>
                        <t:PrimarySmtpAddress>{self._impersonation}</t:PrimarySmtpAddress> <!-- Needs to be in a single line -->
                    </t:ConnectingSID>
                </t:ExchangeImpersonation>
            '''
        else:
            soapHeader = f'<t:RequestServerVersion Version="{self._apiVersion}" />'

        soapEnvelopeOpenTag = '''
            <soap:Envelope 
                xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" 
                xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages" 
                xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types" 
                xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/"
            >'''

        xml = f'''<?xml version="1.0" encoding="utf-8"?>
                    {soapEnvelopeOpenTag}
                        <soap:Header>
                            {soapHeader}
                        </soap:Header>
                        <soap:Body>
                            {soapBody}
                        </soap:Body>
                    </soap:Envelope>
        '''

        print('xml=', xml)
        if self._serverURL:
            url = self._serverURL + '/EWS/exchange.asmx'
        else:
            url = 'https://outlook.office365.com/EWS/exchange.asmx'

        if self._authType == 'Oauth':
            self._session.headers['authorization'] = f'Bearer {self._oauthCallback()}'

        print('session.headers=', self._session.headers)
        resp = self._session.request(
            method='POST',
            url=url,
            data=xml,
            verify=self._verifyCerts,
        )
        print('resp.status_code=', resp.status_code)
        print('resp.reason=', resp.reason)
        print('resp.text=', resp.text)

        if resp.ok and RE_ERROR_CLASS.search(resp.text) is None:
            self._NewConnectionStatus('Connected')
        else:
            for match in RE_ERROR_MESSAGE.finditer(resp.text):
                print('Error Message:', match.group(1))
            self._NewConnectionStatus('Disconnected')

            if 'The account does not have permission to impersonate the requested user.' in resp.text:
                if self._useImpersonationIfAvailable is True:
                    print('Switching impersonation mode')

                    self._useImpersonationIfAvailable = not self._useImpersonationIfAvailable
                    self._useDistinguishedFolderMailbox = not self._useDistinguishedFolderMailbox

                    print('self._useImpersonationIfAvailable=', self._useImpersonationIfAvailable)
                    print('self._useDistinguishedFolderMailbox=', self._useDistinguishedFolderMailbox)

        return resp

    def UpdateCalendar(self, calendar=None, startDT=None, endDT=None):
        # Default is to return events from (now-1days) to (now+7days)
        startDT = startDT or datetime.datetime.utcnow() - datetime.timedelta(days=1)
        endDT = endDT or datetime.datetime.utcnow() + datetime.timedelta(days=7)

        startTimestring = ConvertDatetimeToTimeString(startDT)
        endTimestring = ConvertDatetimeToTimeString(endDT)

        calendar = calendar or self._impersonation or self._username

        if self._useDistinguishedFolderMailbox:
            parentFolder = f'''
                <t:DistinguishedFolderId Id="calendar">
                    <t:Mailbox>
                        <t:EmailAddress>{self._impersonation}</t:EmailAddress>
                    </t:Mailbox>
                </t:DistinguishedFolderId>
            '''
        else:
            parentFolder = f'''
                <t:DistinguishedFolderId Id="calendar"/>
            '''

        soapBody = f'''
            <m:FindItem Traversal="Shallow">
                <m:ItemShape>
                    <t:BaseShape>IdOnly</t:BaseShape>
                    <t:AdditionalProperties>
                        <t:FieldURI FieldURI="item:Subject" />
                        <t:FieldURI FieldURI="calendar:Start" />
                        <t:FieldURI FieldURI="calendar:End" />
                        <t:FieldURI FieldURI="item:Body" />
                        <t:FieldURI FieldURI="calendar:Organizer" />
                        <t:FieldURI FieldURI="calendar:RequiredAttendees" />
                        <t:FieldURI FieldURI="calendar:OptionalAttendees" />
                        <t:FieldURI FieldURI="item:HasAttachments" />
                        <t:FieldURI FieldURI="item:Sensitivity" />
                    </t:AdditionalProperties>
                </m:ItemShape>
                <m:CalendarView 
                    MaxEntriesReturned="100" 
                    StartDate="{startTimestring}" 
                    EndDate="{endTimestring}" 
                    />
                <m:ParentFolderIds>
                     {parentFolder}
                </m:ParentFolderIds>
            </m:FindItem>
        '''
        resp = self._DoRequest(soapBody)
        calItems = self._CreateCalendarItemsFromResponse(resp.text)
        self.RegisterCalendarItems(calItems=calItems, startDT=startDT, endDT=endDT)

    def _CreateCalendarItemsFromResponse(self, responseString):
        '''

        :param responseString:
        :return: list of calendar items
        '''
        ret = []
        for matchCalItem in RE_CAL_ITEM.finditer(responseString):
            print('matchCalItem=', matchCalItem)
            # go thru the resposne and find any CalendarItems.
            # parse their data and findMode CalendarItem objects
            # store CalendarItem objects in self

            # print('\nmatchCalItem.group(0)=', matchCalItem.group(0))

            data = {}
            startDT = None
            endDT = None

            matchItemId = RE_ITEM_ID.search(matchCalItem.group(0))
            data['ItemId'] = matchItemId.group(1)
            data['ChangeKey'] = matchItemId.group(2)
            data['Subject'] = RE_SUBJECT.search(matchCalItem.group(0)).group(1)
            data['OrganizerName'] = RE_ORGANIZER.search(matchCalItem.group(0)).group(1)

            bodyMatch = RE_HTML_BODY.search(matchCalItem.group(0))
            if bodyMatch:
                print('bodyMatch=', bodyMatch)
                data['Body'] = bodyMatch.group(1)

            res = RE_HAS_ATTACHMENTS.search(matchCalItem.group(0)).group(1)
            if 'true' in res:
                data['HasAttachments'] = True
            elif 'false' in res:
                data['HasAttachments'] = False
            else:
                data['HasAttachments'] = 'Unknown'

            startTimeString = RE_START_TIME.search(matchCalItem.group(0)).group(1)
            endTimeString = RE_END_TIME.search(matchCalItem.group(0)).group(1)

            startDT = ConvertTimeStringToDatetime(startTimeString)
            endDT = ConvertTimeStringToDatetime(endTimeString)

            calItem = _CalendarItem(startDT, endDT, data, self)
            ret.append(calItem)

        return ret

    def CreateCalendarEvent(self, subject, body, startDT, endDT):

        startTimeString = ConvertDatetimeToTimeString(startDT)
        endTimeString = ConvertDatetimeToTimeString(endDT)

        calendar = self._impersonation or self._username

        parentFolder = f'''
                    <t:DistinguishedFolderId Id="calendar"/>
                '''

        soapBody = f'''
            <m:CreateItem SendMeetingInvitations="SendToNone">
                <m:SavedItemFolderId>
                    {parentFolder}
                </m:SavedItemFolderId>
                <m:Items>
                    <t:CalendarItem>
                        <t:Subject>{subject}</t:Subject>
                        <t:Body BodyType="Text">{body}</t:Body>
                        <t:Start>{startTimeString}</t:Start>
                        <t:StartTimeZone Id="{self._myTimezoneName}" />
                        <t:End>{endTimeString}</t:End>
                        <t:EndTimeZone Id="{self._myTimezoneName}" />
                    </t:CalendarItem>
                </m:Items>
            </m:CreateItem>
        '''
        resp = self._DoRequest(soapBody)

    def ChangeEventTime(self, calItem, newStartDT=None, newEndDT=None):

        timeUpdateXML = ''

        if newStartDT is not None:
            startTimeString = ConvertDatetimeToTimeString(newStartDT)

        if newEndDT is not None:
            endTimeString = ConvertDatetimeToTimeString(newEndDT)

        soapBody = f'''
            <m:UpdateItem MessageDisposition="SaveOnly" ConflictResolution="AlwaysOverwrite" SendMeetingInvitationsOrCancellations="SendToNone">
              <m:ItemChanges>
                <t:ItemChange>
                  <t:ItemId 
                    Id="{calItem.Get('ItemId')}" 
                    ChangeKey="{calItem.Get('ChangeKey')}" 
                    />
                  <t:Updates>
                    <t:SetItemField>
                      <t:FieldURI FieldURI="calendar:End" />
                      <t:CalendarItem>
                        <t:End>{endTimeString}</t:End>
                        <t:EndTimeZone Id="{self._myTimezoneName}" />
                      </t:CalendarItem>
                    </t:SetItemField>
                  </t:Updates>
                </t:ItemChange>
              </m:ItemChanges>
            </m:UpdateItem>
        '''
        resp = self._DoRequest(soapBody)

        xmlBody = """<?xml version="1.0" encoding="utf-8"?>
                       <soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages"
                              xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types" xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">
                         <soap:Header>
                           {0}
                         </soap:Header>
                         <soap:Body>
                           <m:UpdateItem MessageDisposition="SaveOnly" ConflictResolution="AlwaysOverwrite" SendMeetingInvitationsOrCancellations="SendToNone">
                             <m:ItemChanges>
                               <t:ItemChange>
                                 <t:ItemId Id="{1}" ChangeKey="{2}" />
                                 <t:Updates>
                                   <t:SetItemField>
                                     <t:FieldURI FieldURI="calendar:End" />
                                     <t:CalendarItem>
                                       <t:End>{3}</t:End>
                                       <t:EndTimeZone>{3}</t:EndTimeZone>
                                     </t:CalendarItem>
                                   </t:SetItemField>
                                 </t:Updates>
                               </t:ItemChange>
                             </m:ItemChanges>
                           </m:UpdateItem>
                         </soap:Body>
                       </soap:Envelope> """.format(
            self._soapHeader,
            calItem.Get('ItemId'),
            calItem.Get('ChangeKey'),
            timeUpdateXML,
            TZ_NAME

        )

        self._SendHttp(xmlBody)

    def ChangeEventBody(self, calItem, newBody):
        print('ChangeEventBody(', calItem, newBody)

        soapBody = f"""
            <m:UpdateItem MessageDisposition="SaveOnly" ConflictResolution="AlwaysOverwrite" SendMeetingInvitationsOrCancellations="SendToNone">
              <m:ItemChanges>
                <t:ItemChange>
                  <t:ItemId 
                    Id="{calItem.Get('ItemId')}"
                    ChangeKey="{calItem.Get('ChangeKey')}" 
                    />
                  <t:Updates>
                    <t:SetItemField>
                      <t:FieldURI FieldURI="item:Body" />
                      <t:CalendarItem>
                        <t:Body BodyType="HTML">{newBody}</t:Body>
                        <t:Body BodyType="Text">{newBody}</t:Body>
                      </t:CalendarItem>
                    </t:SetItemField>
                  </t:Updates>
                </t:ItemChange>
              </m:ItemChanges>
            </m:UpdateItem>
            """
        resp = self._DoRequest(soapBody)

    def DeleteEvent(self, calItem):
        soapBody = f"""
                <m:DeleteItem DeleteType="HardDelete" SendMeetingCancellations="SendToNone">
                  <m:ItemIds>
                    <t:ItemId 
                        Id="{calItem.Get('ItemId')}"
                        ChangeKey="{calItem.Get('ChangeKey')}" 
                    />
                  </m:ItemIds>
                </m:DeleteItem>
            """
        resp = self._DoRequest(soapBody)


def ConvertDatetimeToTimeString(dt):
    return dt.strftime('%Y-%m-%dT%H:%M:%SZ')


def ConvertTimeStringToDatetime(string):
    return datetime.datetime.strptime(string, '%Y-%m-%dT%H:%M:%SZ')


if __name__ == '__main__':
    import creds

    ews = EWS(

        # gm has ApplicationImpersonation
        username='gm_service_account@extrondemo.com',
        impersonation='rf_a101@extrondemo.com',
        password='Extron1025',

        # username='rf_a101@extrondemo.com',
        # password='Extron123!',

        # username='gm_service_account@extrondemo.com',
        # password='Extron1025',

        # username='impersonation-onprem@extron.com',
        # impersonation='Test-pm4@extron.com',
        # password='Extron1025',

        # username=creds.username,
        # password=creds.password,

    )

    ews.Connected = lambda _, state: print('EWS', state)
    ews.Disconnected = lambda _, state: print('EWS', state)
    ews.NewCalendarItem = lambda _, item: print('NewCalendarItem(', item)
    ews.CalendarItemChanged = lambda _, item: print('CalendarItemChanged(', item)
    ews.CalendarItemDeleted = lambda _, item: print('CalendarItemDeleted(', item)

    ews.UpdateCalendar()
    ews.CreateCalendarEvent(
        subject='Test Subject ' + time.asctime(),
        body='Test Body ' + time.asctime(),
        startDT=datetime.datetime.utcnow(),
        endDT=datetime.datetime.utcnow() + datetime.timedelta(minutes=15),
    )
