'''
Based on the module created by David Gonzalez (dgonzalez@extron.com)
Re-worked by Grant Miller (gmiller@extron.com)
'''

import urllib.request, re
from base64 import b64encode, b64decode
import datetime
import time
import requests
import xml

DEBUG = True
oldPrint = print
if not DEBUG:
    print = lambda *a, **k: None

try:
    from extronlib.system import File, ProgramLog
except:
    File = open
    ProgramLog = print

offsetSeconds = time.timezone if (time.localtime().tm_isdst == 0) else time.altzone
offsetHours = offsetSeconds / 60 / 60 * -1
MY_TIME_ZONE = offsetHours

TZ_NAME = time.tzname[0]
if TZ_NAME == 'EST':
    TZ_NAME = 'Eastern Standard Time'
elif TZ_NAME == 'PST':
    TZ_NAME = 'Pacific Standard Time'
elif TZ_NAME == 'CST':
    TZ_NAME = 'Central Standard Time'

print('MY_TIME_ZONE= UTC {}'.format(MY_TIME_ZONE))
print('CURRENT SYSTEM TIME=', time.asctime())

RE_CAL_ITEM = re.compile('<t:CalendarItem>[\w\W]*?<\/t:CalendarItem>')
RE_ITEM_ID = re.compile(
    '<t:ItemId Id="(.*?)" ChangeKey="(.*?)"/>')  # group(1) = itemID, group(2) = changeKey #within a CalendarItem
RE_SUBJECT = re.compile('<t:Subject>(.*?)</t:Subject>')  # within a CalendarItem
RE_HAS_ATTACHMENTS = re.compile('<t:HasAttachments>(.{4,5})</t:HasAttachments>')  # within a CalendarItem
RE_ORGANIZER = re.compile(
    '<t:Organizer>.*<t:Name>(.*?)</t:Name>.*</t:Organizer>')  # group(1)=Name #within a CalendarItem
RE_START_TIME = re.compile('<t:Start>(.*?)</t:Start>')  # group(1) = start time string #within a CalendarItem
RE_END_TIME = re.compile('<t:End>(.*?)</t:End>')  # group(1) = end time string #within a CalendarItem
RE_HTML_BODY = re.compile('<t:Body BodyType="HTML">([\w\W]*)</t:Body>', re.IGNORECASE)

RE_EMAIL_ADDRESS = re.compile('.*?\@.*?\..*?')


def ConvertTimeStringToDatetime(string):
    # print('48 ConvertTimeStringToDatetime\nstring=', string)
    year, month, etc = string.split('-')
    day, etc = etc.split('T')
    hour, minute, etc = etc.split(':')
    second = etc[:-1]
    dt = datetime.datetime(
        year=int(year),
        month=int(month),
        day=int(day),
        hour=int(hour),
        minute=int(minute),
        second=int(second),
    )

    dt = AdjustDatetimeForTimezone(dt, fromZone='Exchange')
    # print('63 dt=', dt)
    return dt


def ConvertDatetimeToTimeString(dt):
    dt = AdjustDatetimeForTimezone(dt, fromZone='Mine')
    return dt.strftime('%Y-%m-%dT%H:%M:%SZ')


def AdjustDatetimeForTimezone(dt, fromZone):
    delta = datetime.timedelta(hours=abs(MY_TIME_ZONE))

    ts = time.mktime(dt.timetuple())
    lt = time.localtime(ts)
    dtIsDST = lt.tm_isdst > 0

    nowIsDST = time.localtime().tm_isdst > 0

    print('nowIsDST=', nowIsDST)
    print('dtIsDST=', dtIsDST)

    if fromZone == 'Mine':
        dt = dt + delta
        if dtIsDST and not nowIsDST:
            dt -= datetime.timedelta(hours=1)
        elif nowIsDST and not dtIsDST:
            dt += datetime.timedelta(hours=1)

    elif fromZone == 'Exchange':
        dt = dt - delta
        if dtIsDST and not nowIsDST:
            dt += datetime.timedelta(hours=1)
        elif nowIsDST and not dtIsDST:
            dt -= datetime.timedelta(hours=1)

    return dt


class _CalendarItem:
    def __init__(self, startDT, endDT, data, parentExchange):
        if data is None:
            data = {}
        # print('_CalendarItem data=', data)
        self._data = data.copy()  # dict like {'ItemId': 'jasfsd', 'Subject': 'SuperMeeting', ...}
        self._startDT = startDT
        self._endDT = endDT
        self._attachments = []
        self._parentExchange = parentExchange

    def AddData(self, key, value):
        self._data[key] = value

    def _CalculateDuration(self):
        # Returns float in seconds
        delta = self.Get('End') - self.Get('Start')
        duration = delta.total_seconds()
        self.AddData('Duration', duration)

    def Get(self, key, default=None):
        if key == 'Start':
            return self._startDT
        elif key == 'End':
            return self._endDT
        elif key == 'Duration':
            self._CalculateDuration()
            return self._data.get(key, default)
        elif key == 'Body':
            self._UpdateFromServer()
            return self._data.get(key, default)
        else:
            return self._data.get(key, default)

    def get(self, key):
        return self.Get(key)

    def __contains__(self, dt):
        '''
        allows you to compare _CalendarItem object like you would compare datetime objects

        Example:
        if datetime.datetime.now() in calItem:
            print('the datetime is within the CalendarItem start/end')

        :param dt:
        :return:
        '''
        # Note: isinstance(datetime.datetime.now(), datetime.date.today()) == True
        # Because the point in time exist in that date
        if isinstance(dt, datetime.datetime):
            if self._startDT <= dt <= self._endDT:
                return True
            else:
                return False

        elif isinstance(dt, datetime.date):
            if self._startDT.year == dt.year and \
                    self._startDT.month == dt.month and \
                    self._startDT.day == dt.day:
                return True

            elif self._endDT.year == dt.year and \
                    self._endDT.month == dt.month and \
                    self._endDT.day == dt.day:
                return True

            else:
                return False

    def GetAttachments(self):
        self._attachments = []
        for attachmentID in self._parentExchange._GetAttachmentIDs(self):
            print('122 attachmentID=', attachmentID)
            newAttachmentObject = _Attachment(attachmentID, self._parentExchange)
            self._attachments.append(newAttachmentObject)

        return self._attachments.copy()

    def HasAttachments(self):

        if len(self._attachments) > 0:
            return True
        else:
            if self._data['HasAttachments'] is True:
                return True
            else:
                return False

        return False

    def Update(self, key, value):
        if key == 'Body':
            self._parentExchange.ChangeEventBody(self, value)
            self._UpdateFromServer()
        else:
            raise KeyError('Only "Body" can be updated at this time')

    @property
    def Data(self):
        return self._data.copy()

    def __iter__(self):
        for k, v in self._data.items():
            yield k, v

        for key in ['Start', 'End', 'Duration']:
            yield key, self.Get(key)

    def __str__(self):
        return '<CalendarItem object: Start={}, End={}, Subject={}, HasAttachements={}, ItemId[-7:]={}>'.format(
            self.Get('Start'),
            self.Get('End'),
            self.Get('Subject'),
            self.HasAttachments(),
            self.Get('ItemId')[-7:],
        )

    def __repr__(self):
        return str(self)

    def __eq__(self, other):
        # oldPrint('188 __eq__ self.Data=', self.Data, ',\nother.Data=', other.Data)
        return self.Get('ItemId') == other.Get('ItemId') and \
               self.Get('ChangeKey') == other.Get('ChangeKey')

    def __lt__(self, other):
        # print('192 __lt__', self, other)
        if isinstance(other, datetime.datetime):
            return self._startDT < other

        elif isinstance(other, _CalendarItem):
            return self._startDT < other._startDT

        else:
            raise TypeError('unorderable types: {} < {}'.format(self, other))

    def __le__(self, other):
        # print('203 __le__', self, other)
        if isinstance(other, datetime.datetime):
            return self._startDT <= other

        elif isinstance(other, _CalendarItem):
            return self._startDT <= other._startDT

        else:
            raise TypeError('unorderable types: {} < {}'.format(self, other))

    def __gt__(self, other):
        # print('214 __gt__', self, other)
        if isinstance(other, datetime.datetime):
            return self._endDT > other
        elif isinstance(other, _CalendarItem):
            return self._endDT > other._endDT

        else:
            raise TypeError('unorderable types: {} < {}'.format(self, other))

    def __ge__(self, other):
        # print('223 __ge__', self, other)
        if isinstance(other, datetime.datetime):
            return self._endDT >= other
        elif isinstance(other, _CalendarItem):
            return self._endDT >= other._endDT

        else:
            raise TypeError('unorderable types: {} < {}'.format(self, other))

    def _UpdateFromServer(self):
        calItem = self._parentExchange.GetItem(self.Get('ItemId'))
        self._parentExchange.RegisterCalendarItem(calItem)
        self._data.update(dict(calItem))


class _Attachment:
    def __init__(self, AttachmentId, parentExchange):
        self.Filename = None
        self.AttachmentId = AttachmentId
        self._parentExchange = parentExchange
        self._content = None

    def _Update(self):
        self._parentExchange._UpdateAttachmentData(self)

    def GetContent(self):
        if self._content is None:
            self._Update()

        return self._content

    def GetContentSize(self):
        # return size of content in KB
        return len(self.GetContent()) / 1024

    def GetID(self):
        return self.AttachmentId

    def GetFilename(self):
        if self.Filename is None:
            self._Update()
        return self.Filename

    def SaveToPath(self, path):
        with File(path, mode='wb') as file:
            file.write(self.GetContent())


class Exchange:
    # Exchange methods
    def __init__(self,
                 username=None,
                 password=None,
                 server='outlook.office365.com',
                 impersonation=None,
                 proxyAddress=None,
                 proxyPort=None,
                 accessTokenCallback=None,  # None or callback that accepts no params and returns an access token
                 ):

        self._username = username
        self._password = password
        self._accessTokenCallback = accessTokenCallback  # used for oauth

        if accessTokenCallback is None:
            if username is None or password is None:
                raise PermissionError('Please provide a username/password or accessTokenCallback')

        self._proxyAddress = proxyAddress
        self._proxyPort = proxyPort

        if self._proxyAddress or self._proxyPort:
            proxyHandler = urllib.request.ProxyHandler({
                'http': 'http://{}:{}'.format(
                    self._proxyAddress,
                    self._proxyPort if self._proxyPort else '3128',  # default proxy port is 3128
                ),
                'https': 'https://{}:{}'.format(
                    self._proxyAddress,
                    self._proxyPort if self._proxyPort else '3128',  # default proxy port is 3128
                ),
            })
            newOpener = urllib.request.build_opener(
                proxyHandler,
                urllib.request.ProxyBasicAuthHandler()
            )
            urllib.request.install_opener(newOpener)

        self.httpURL = 'https://{0}/EWS/exchange.asmx'.format(server)
        print('self.httpURL=', self.httpURL)
        # self.httpURL = 'http://{0}/EWS/exchange.asmx'.format(server) #testing only
        self.encode = b64encode(bytes('{0}:{1}'.format(self._username, self._password), "ascii"))
        self.login = str(self.encode)[2:-1]
        self._impersonation = impersonation or None
        if self._accessTokenCallback:
            self.header = {
                'content-type': 'text/xml',
                'authorization': 'Bearer {}'.format(self._accessTokenCallback())
            }
        else:
            self.header = {
                'content-type': 'text/xml',
                'authorization': 'Basic {}'.format(self.login)
            }
        self._calendarItems = []

        self._startOfWeek = None
        self._endOfWeek = None
        self._soapHeader = None

        self._folderID = None
        self._changeKey = None

        self._connectionStatus = None
        self._Connected = None
        self._Disconnected = None

        self._CalendarItemDeleted = None  # callback for when an item is deleted
        self._CalendarItemChanged = None
        self._NewCalendarItem = None

        self._UpdateFolderIdAndChangeKey()

    @property
    def NewCalendarItem(self):
        return self._NewCalendarItem

    @NewCalendarItem.setter
    def NewCalendarItem(self, func):
        self._NewCalendarItem = func

    ##############
    @property
    def CalendarItemChanged(self):
        return self._CalendarItemChanged

    @CalendarItemChanged.setter
    def CalendarItemChanged(self, func):
        self._CalendarItemChanged = func

    ############
    @property
    def CalendarItemDeleted(self):
        return self._CalendarItemDeleted

    @CalendarItemDeleted.setter
    def CalendarItemDeleted(self, func):
        self._CalendarItemDeleted = func

    ############
    @property
    def Connected(self):
        return self._Connected

    @Connected.setter
    def Connected(self, func):
        self._Connected = func

    #############
    @property
    def Disconnected(self):
        return self._Disconnected

    @Disconnected.setter
    def Disconnected(self, func):
        self._Disconnected = func

    def _NewConnectionStatus(self, state, forceNotification=True):
        print('378 _NewConnectionStatus(', state, ', self._connectionStatus=', self._connectionStatus)
        if state != self._connectionStatus or forceNotification:
            # the connection status has changed
            self._connectionStatus = state
            if state == 'Connected':
                if callable(self._Connected):
                    self._Connected(self, state)
            elif state == 'Disconnected':
                if callable(self._Disconnected):
                    self._Disconnected(self, state)

    # ----------------------------------------------------------------------------------------------------------------------
    # --------------------------------------------------EWS Services--------------------------------------------------------
    # ----------------------------------------------------------------------------------------------------------------------
    def _GetSoapHeader(self, emailAddress):
        # This should only need to be called once to findMode the header that will be used in the XML request from now on
        if emailAddress is None:
            xmlAccount = """<t:RequestServerVersion Version="Exchange2013" />"""
        else:
            # replace = '<t:PrincipalName>{0}</t:PrincipalName>'.format(emailAddress)
            # replace = '<t:SID>{0}</t:SID>'.format(emailAddress)
            # replace = '<t:PrimarySmtpAddress>{0}</t:PrimarySmtpAddress>'.format(emailAddress)
            replace = '<t:SmtpAddress>{0}</t:SmtpAddress>'.format(emailAddress)

            # more info on impersonation at this link: https://docs.microsoft.com/en-us/exchange/client-developer/exchange-web-services/how-to-add-appointments-by-using-exchange-impersonation

            xmlAccount = """<t:RequestServerVersion Version="Exchange2013" />
                            <ExchangeImpersonation>
                                <ConnectingSID>
                                   {0}
                                </ConnectingSID>
                            </ExchangeImpersonation>""".format(replace)
        return xmlAccount

    def _UpdateStartEndOfWeek(self):
        if self._startOfWeek is None or self._endWeekDT < datetime.datetime.now():
            todayDT = datetime.date.today()
            weekday = todayDT.weekday()
            startWeekDT = todayDT - datetime.timedelta(days=weekday)
            self._startWeekDT = datetime.datetime(
                year=startWeekDT.year,
                month=startWeekDT.month,
                day=startWeekDT.day,
            )
            self._startOfWeek = ConvertDatetimeToTimeString(startWeekDT)

        if self._endOfWeek is None or self._endWeekDT < datetime.datetime.now():
            todayDT = datetime.date.today()
            weekday = todayDT.weekday()
            endWeekDT = todayDT + datetime.timedelta(days=6 - weekday)
            self._endWeekDT = datetime.datetime(
                year=endWeekDT.year,
                month=endWeekDT.month,
                day=endWeekDT.day,
                hour=23,
                minute=59,
            )
            self._endOfWeek = ConvertDatetimeToTimeString(endWeekDT)

    def _UpdateFolderIdAndChangeKey(self):
        # Requests Service for ID of calendar folder and change key
        self._soapHeader = self._GetSoapHeader(self._impersonation)

        self._UpdateStartEndOfWeek()

        regExFolderInfo = re.compile(r't:FolderId Id=\"(.{1,})\" ChangeKey=\"(.{1,})\"\/')

        if self._impersonation:
            distinguishedFolderID = '''<t:DistinguishedFolderId Id="calendar">
                          <t:Mailbox>
                            <t:SmtpAddress>{}</t:SmtpAddress>
                          </t:Mailbox>
                        </t:DistinguishedFolderId>'''.format(self._impersonation)
        else:
            distinguishedFolderID='<t:DistinguishedFolderId Id="calendar" />'


        xmlbody = """<?xml version="1.0" encoding="utf-8"?>
                    <soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
                           xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages"
                           xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types"
                           xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">
                      <soap:Header>
                        {0}
                      </soap:Header>
                      <soap:Body>
                        <m:GetFolder>
                          <m:FolderShape>
                            <t:BaseShape>IdOnly</t:BaseShape>
                            
                            <t:AdditionalProperties>
                                <t:FieldURI FieldURI="folder:DisplayName" />
                                <t:FieldURI FieldURI="folder:ParentFolderId" />
                                
                                <t:FieldURI FieldURI="folder:FolderId" />
                                <t:FieldURI FieldURI="folder:ParentFolderId" />
                                <t:FieldURI FieldURI="folder:TotalCount" />
                                <t:FieldURI FieldURI="folder:ChildFolderCount" />
                                <!--<t:FieldURI FieldURI="folder:ExtendedProperty" /> Causing 500 error-->
                                <t:FieldURI FieldURI="folder:FolderClass" />
                                <t:FieldURI FieldURI="folder:ManagedFolderInformation" />
                                <t:FieldURI FieldURI="folder:EffectiveRights" />
                                <t:FieldURI FieldURI="folder:PermissionSet" />
                                <t:FieldURI FieldURI="folder:EffectiveRights" />
                                <t:FieldURI FieldURI="folder:SharingEffectiveRights" />
                                
                            </t:AdditionalProperties>
                            
                          </m:FolderShape>
                          
                          <m:FolderIds>
                            {1}
                          </m:FolderIds>
                          
                        </m:GetFolder>
                      </soap:Body>
                    </soap:Envelope>""".format(
            self._soapHeader,
            distinguishedFolderID
        )

        # Request for ID and Key
        response = self._SendHttp(xmlbody)

        if isinstance(response, str):
            matchFolderInfo = regExFolderInfo.search(response)
            # Set FolderId and ChangeKey
            if matchFolderInfo:
                self._folderID = matchFolderInfo.group(1)
                self._changeKey = matchFolderInfo.group(2)
                print('526 _folderID=', self._folderID)
                print('527 _changeKey=', self._changeKey)

    @property
    def Impersonation(self):
        return self._impersonation

    @Impersonation.setter
    def Impersonation(self, newImpersonation):
        self._impersonation = newImpersonation

    def _GetParentFolderTag(self, calendar=None, mode='find'):
        print('_GetParentFolderTag(', calendar, mode)
        calendar = calendar or self._impersonation or None

        if calendar is None:
            if self._impersonation is None:
                parentFolder = '''
                    <m:ParentFolderIds>
                        <t:FolderId Id="{}" ChangeKey="{}" />
                    </m:ParentFolderIds>
                    '''.format(
                    self._folderID,
                    self._changeKey
                )
            else:
                parentFolder = '''
                    <m:ParentFolderIds>
                        <t:DistinguishedFolderId Id="calendar">
                          <t:Mailbox>
                            <t:SmtpAddress>{}</t:SmtpAddress>
                          </t:Mailbox>
                        </t:DistinguishedFolderId>
                      </m:ParentFolderIds>
                    '''.format(self._impersonation)

        elif RE_EMAIL_ADDRESS.search(calendar) is not None:  # email address
            print('520 emailRegex matched')

            if mode == 'find':
                parentFolder = '''
                    <m:ParentFolderIds>
                        <t:DistinguishedFolderId Id="calendar">
                          <t:Mailbox>
                            <t:EmailAddress>{}</t:EmailAddress>
                          </t:Mailbox>
                        </t:DistinguishedFolderId>
                      </m:ParentFolderIds>
                    '''.format(calendar)
            elif mode == 'create':
                parentFolder = '''
                    <m:ParentFolderId>
                        <t:DistinguishedFolderId Id="{}" ChangeKey="{}">
                          <t:Mailbox>
                            <t:SmtpAddress>{}</t:SmtpAddress>
                          </t:Mailbox>
                        </t:DistinguishedFolderId>
                      </m:ParentFolderId>
                    '''.format(
                    self._folderID,
                    self._changeKey,
                    calendar
                )
        else:  # name
            print('531 name')
            parentFolder = '''
                <m:ParentFolderIds>
                    <t:DistinguishedFolderId Id="calendar">
                      <t:Mailbox>
                        <t:Name>{}</t:Name>
                      </t:Mailbox>
                    </t:DistinguishedFolderId>
                  </m:ParentFolderIds>
                '''.format(calendar)
        print('581 parentFolder=', parentFolder)
        return parentFolder

    def UpdateCalendar(self, calendar=None, startDT=None, endDT=None):
        print('UpdateCalendar(', calendar, startDT, endDT)
        # see all <CalendarItem> options here: https://docs.microsoft.com/en-us/exchange/client-developer/web-service-reference/calendaritem

        # FieldURI's : https://docs.microsoft.com/en-us/exchange/client-developer/web-service-reference/fielduri

        # gets the latest data for this week from exchange and stores it
        # if calendar is not None, this will check another users calendar
        # if calendar is None, it will check your own calendar

        if startDT is None:
            startDTstring = self._startOfWeek
            startDT = self._startWeekDT
        else:
            startDTstring = ConvertDatetimeToTimeString(startDT)

        if endDT is None:
            endDTstring = self._endOfWeek
            endDT = self._endWeekDT
        else:
            endDTstring = ConvertDatetimeToTimeString(endDT)

        if self._endWeekDT < datetime.datetime.now():
            self._UpdateFolderIdAndChangeKey()

        parentFolder = self._GetParentFolderTag(calendar)

        # Note: All FieldURI's located here:
        # https://docs.microsoft.com/en-us/exchange/client-developer/web-service-reference/fielduri
        xmlbody = """<?xml version="1.0" encoding="utf-8"?>
                    <soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
                           xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages"
                           xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types"
                           xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">
                      <soap:Header>
                        {0}
                      </soap:Header>
                      <soap:Body>
                        <m:FindItem Traversal="Shallow">
                          <m:ItemShape>
                            <t:BaseShape>IdOnly</t:BaseShape>
                            <t:AdditionalProperties>
                              <t:FieldURI FieldURI="item:Subject" />
                              
                              <t:FieldURI FieldURI="calendar:Start" />
                              <t:FieldURI FieldURI="calendar:End" />
                              
                              <t:FieldURI FieldURI="item:Body" />
                                <!--
                                <t:FieldURI FieldURI="item:NormalizedBody" />
                                <t:FieldURI FieldURI="item:UniqueBody" />
                                <t:FieldURI FieldURI="item:TextBody" />
                                -->
                              
                              <t:FieldURI FieldURI="calendar:Organizer"/>
                            
                              
                              <t:FieldURI FieldURI="calendar:RequiredAttendees" /> 
                              <t:FieldURI FieldURI="calendar:OptionalAttendees" /> 
                              
                              <t:FieldURI FieldURI="item:HasAttachments" />
                              
                              <t:FieldURI FieldURI="item:Sensitivity" /> <!-- Private Meeting Flag -->
                              
                              
                              
                            </t:AdditionalProperties>
                          </m:ItemShape>
                          <m:CalendarView 
                            MaxEntriesReturned="100" 
                            StartDate="{1}" 
                            EndDate="{2}" />
                            {3}
                        </m:FindItem>
                        
                        
                        
                      </soap:Body>
                    </soap:Envelope>
                    """.format(
            self._soapHeader,  # 0
            startDTstring,  # 1
            endDTstring,  # 2
            parentFolder  # 3
        )

        response = self._SendHttp(xmlbody)
        print('responseString=', response)
        if response:
            exchangeItems = self._CreateCalendarItemsFromResponse(response)

            # check all calitems for changes
            # do callbacks if something changes

            for exchangeItem in exchangeItems:
                if not exchangeItem.Get('Subject').startswith('Canceled:'):  # ignore cancelled items
                    selfItem = self.GetCalendarItemByID(exchangeItem.Get('ItemId'))
                    if selfItem is None:
                        # this is a new item do callback
                        self._calendarItems.append(exchangeItem)
                        if callable(self._NewCalendarItem):
                            self._NewCalendarItem(self, exchangeItem)

                    elif selfItem != exchangeItem:
                        # the item has changed somehow, do callback
                        self._calendarItems.remove(selfItem)
                        self._calendarItems.append(exchangeItem)
                        if callable(self._CalendarItemChanged):
                            self._CalendarItemChanged(self, exchangeItem)

            # check all the items to make sure nothings been deleted
            for selfItem in self._calendarItems.copy():
                # print('startDT=', startDT)
                # print('endDT=', endDT)
                if startDT <= selfItem <= endDT:
                    if selfItem not in exchangeItems:
                        # a event was deleted from the exchange server
                        self._calendarItems.remove(selfItem)
                        if callable(self._CalendarItemDeleted):
                            self._CalendarItemDeleted(self, selfItem)

        else:
            oldPrint(response)
            self._NewConnectionStatus('Disconnected', forceNotification=True)
            raise Exception('UpdateCalendar failed. Check ProgramLog. ' + str(response))

    def _CreateCalendarItemsFromResponse(self, response):
        '''

        :param response:
        :return: list of calendar items
        '''
        ret = []
        for matchCalItem in RE_CAL_ITEM.finditer(response):
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

    def RegisterCalendarItem(self, calItem):
        '''
        This method will add the calendar item to self._calendarItems
        Making sure it is not duplicated and replacing any old data with new data
        Applicable callbacks will be executed
        :param calItem:
        :return:
        '''

        # Remove any CalendarItems that have ended in the past
        # nowDT = datetime.datetime.now()
        # for sub_calItem in self._calendarItems.copy():
        #     endDT = sub_calItem.Get('End')
        #     if endDT < nowDT:
        #         if sub_calItem in self._calendarItems:
        #             self._calendarItems.remove(sub_calItem)

        # Remove any old nowItems that have the same ItemId
        itemId = calItem.Get('ItemId')

        for sub_calItem in self._calendarItems.copy():
            if sub_calItem.Get('ItemId') == itemId:

                if sub_calItem != calItem:
                    # something has changed about this item, do callback
                    if self.CalendarItemChanged:
                        self.CalendarItemChanged(self, calItem)

                self._calendarItems.remove(sub_calItem)

        # Add CalItem to self
        self._calendarItems.append(calItem)

    def CreateCalendarEvent(self, subject, body, startDT=None, endDT=None):
        print('CreateCalendarEvent(', subject, body, startDT, endDT)

        startTimeString = ConvertDatetimeToTimeString(startDT)
        endTimeString = ConvertDatetimeToTimeString(endDT)

        xmlBody = """<?xml version="1.0" encoding="utf-8"?>
                    <soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
                           xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages"
                           xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types"
                           xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">
                      <soap:Header>
                        {0}
                      </soap:Header>
                      <soap:Body>
                        <!-- see link: https://docs.microsoft.com/en-us/dotnet/api/microsoft.exchange.webservices.data.sendinvitationsmode?view=exchange-ews-api -->
                        <m:CreateItem SendMeetingInvitations="SendOnlyToAll">
                        <!--<m:CreateItem SendMeetingInvitations="SendToAllAndSaveCopy">-->
                        <!--<m:CreateItem SendMeetingInvitations="SendToNone">-->
                          <m:Items>
                            <t:CalendarItem>
                              <t:Subject>{1}</t:Subject>
                              <t:Body BodyType="HTML">{2}</t:Body>
                              <t:Body BodyType="Text">{2}</t:Body>
                              <t:Start>{3}</t:Start>
                              <t:End>{4}</t:End>
                              <t:MeetingTimeZone TimeZoneName="{5}" />

                              <!-- Add an attendee, needed for the impersonation room to accept the meeting into their calendar -->
                              <!--
                              <t:Location>{6}</t:Location>
                              <t:RequiredAttendees>
                                <t:Attendee>
                                    <t:Mailbox>
                                        <t:EmailAddress>{6}</t:EmailAddress>
                                    </t:Mailbox>
                                </t:Attendee>
                                HTTPretty   
                              </t:RequiredAttendees>


                              <t:Resources>
                                <t:Attendee>
                                    <t:Mailbox>
                                        <t:EmailAddress>{6}</t:EmailAddress>
                                    </t:Mailbox>
                                </t:Attendee>
                            </t:Resources>

                            -->
                            <!-- nope dont work -->
                            <Organizer>
                                <t:Mailbox>
                                    <t:EmailAddress>{6}</t:EmailAddress>
                                </t:Mailbox>
                            </Organizer>

                            </t:CalendarItem>
                          </m:Items>
                        </m:CreateItem>
                      </soap:Body>
                    </soap:Envelope>""".format(
            self._soapHeader,
            subject,
            body,
            startTimeString,
            endTimeString,
            TZ_NAME,
            self._impersonation
        )

        self._SendHttp(xmlBody)

    def _SendHttp_requests(self, body):
        resp = requests.post(
            url=self.httpURL,
            headers={
                'content-type': 'text/xml',
                'authorization': 'Bearer {}'.format(self._accessTokenCallback())
            },
            data=body,
        )
        print('soapBody=\r\n', body)
        print('resp.text=', resp.text)
        print('resp.status_code=', resp.status_code)
        print('resp.reason=', resp.reason)

        ret = resp.text
        print('965 ret=', ret)
        return ret

    def CreateMeeting_WIP(self, subject, body, startDT=None, endDT=None):
        # https://docs.microsoft.com/en-us/exchange/client-developer/exchange-web-services/how-to-access-a-calendar-as-a-delegate-by-using-ews-in-exchange
        # https://docs.microsoft.com/en-us/exchange/client-developer/exchange-web-services/how-to-access-a-calendar-as-a-delegate-by-using-ews-in-exchange
        print('CreateMeeting(', subject, body, startDT, endDT)

        startTimeString = ConvertDatetimeToTimeString(startDT)
        endTimeString = ConvertDatetimeToTimeString(endDT)

        xmlBody = """<?xml version="1.0" encoding="utf-8"?>
            <soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
                    xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages"
                    xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types"
                    xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">
                <soap:Header>
                    {0}
                </soap:Header>
                <soap:Body>
                    <!-- see link: https://docs.microsoft.com/en-us/dotnet/api/microsoft.exchange.webservices.data.sendinvitationsmode?view=exchange-ews-api -->
                    
                    <m:CreateItem SendMeetingInvitations="SendOnlyToAll">
                    <!--<m:CreateItem SendMeetingInvitations="SendToAllAndSaveCopy">-->
                    <!--<m:CreateItem SendMeetingInvitations="SendToNone">-->
                    
                        <m:SavedItemFolderId>
                            <t:DistinguishedFolderId Id="calendar"/>
                            <t:Mailbox>
                                <t:EmailAddress>{6}</t:EmailAddress>
                            </t:Mailbox>
                        </m:SavedItemFolderId>
                        
                        <m:Items>
                            <t:CalendarItem>
                                <t:Subject>{1}</t:Subject>
                                <t:Start>{3}</t:Start>
                                <t:End>{4}</t:End>
                                <t:MeetingTimeZone TimeZoneName="{5}" />
                                
                            </t:CalendarItem>
                        </m:Items>
                    </m:CreateItem>
                </soap:Body>
            </soap:Envelope>""".format(
            self._soapHeader,
            subject,
            body,
            startTimeString,
            endTimeString,
            TZ_NAME,
            self._impersonation,  # 6
            '',  # self._GetParentFolderTag(self._impersonation, mode='create')
        )

        self._SendHttp_requests(xmlBody)

    def GetCalendarItemByID(self, itemId):
        for calItem in self._calendarItems:
            if calItem.Get('ItemId') == itemId:
                return calItem

    # This function updates the end time of an event. Can be modified to update other functions
    # Was built to update end time for RoomAgent GS needs only
    def ChangeEventTime(self, calItem, newStartDT=None, newEndDT=None):

        timeUpdateXML = ''

        if newStartDT is not None:
            startTimeString = ConvertDatetimeToTimeString(newStartDT)
            timeUpdateXML += '<t:Start>{}</t:Start>'.format(startTimeString)

        if newEndDT is not None:
            endTimeString = ConvertDatetimeToTimeString(newEndDT)
            timeUpdateXML += '<t:End>{}</t:End>'.format(endTimeString)

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
                                  <t:FieldURI FieldURI="item:Body" />
                                  <t:CalendarItem>
                                    <t:Body BodyType="HTML">{3}</t:Body>
                                    <t:Body BodyType="Text">{3}</t:Body>
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
            newBody
        )

        self._SendHttp(xmlBody)

    def DeleteEvent(self, calItem):
        print('565 DeleteEvent(', calItem)
        xmlBody = """<?xml version="1.0" encoding="utf-8"?>
                    <soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
                           xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages"
                           xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types"
                           xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">
                      <soap:Header>
                        {0}
                      </soap:Header>
                      <soap:Body>
                        <m:DeleteItem DeleteType="HardDelete" SendMeetingCancellations="SendToNone">
                          <m:ItemIds>
                            <t:ItemId Id="{1}" ChangeKey="{2}" />
                          </m:ItemIds>
                        </m:DeleteItem>
                      </soap:Body>
                    </soap:Envelope>""".format(
            self._soapHeader,
            calItem.Get('ItemId'),
            calItem.Get('ChangeKey')
        )

        request = self._SendHttp(xmlBody)

    def _UpdateAttachmentData(self, attachmentObject):
        # sets the filename and content of attachment object
        attachmentID = attachmentObject.GetID()

        regExReponse = re.compile(r'<m:ResponseCode>(.+)</m:ResponseCode>')
        regExName = re.compile(r'<t:Name>(.+)</t:Name>')
        regExContentType = re.compile(r'<t:ContentType>(.+)</t:ContentType>')
        regExContent = re.compile(r'<t:Content>(.+)</t:Content>')
        attachmentObjects = []

        xmlBody = """<?xml version="1.0" encoding="utf-8"?>
                            <soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
                                           xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages"
                                           xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types"
                                           xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">
                              <soap:Header>
                                {0}
                              </soap:Header>
                              <soap:Body>
                                <m:GetAttachment>
                                  <m:AttachmentIds>
                                    <t:AttachmentId Id="{1}" />
                                  </m:AttachmentIds>
                                </m:GetAttachment>
                              </soap:Body>
                            </soap:Envelope>""".format(self._soapHeader, attachmentID)

        request = self._SendHttp(xmlBody)

        responseCode = regExReponse.search(request).group(1)
        if responseCode == 'NoError':  # Handle errors sent by the server
            itemName = regExName.search(request).group(1)
            itemName = itemName.replace(' ', '_')  # remove ' ' chars because dont work on linux
            itemContent = regExContent.search(request).group(1)

            attachmentObject._content = b64decode(itemContent)
            attachmentObject.Filename = itemName

    def _GetAttachmentIDs(self, calItem):
        # returns a list of attachment IDs

        itemId = calItem.Get('ItemId')
        regExAttKey = re.compile(r'AttachmentId Id=\"(\S+)\"')
        xmlBody = """<?xml version="1.0" encoding="utf-8"?>
                        <soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
                                       xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages"
                                       xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types"
                                       xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">
                          <soap:Header>
                            {0}
                          </soap:Header>
                          <soap:Body>
                            <m:GetItem>
                              <m:ItemShape>
                                <t:BaseShape>IdOnly</t:BaseShape>
                                <t:AdditionalProperties>
                                  <t:FieldURI FieldURI="item:Attachments" />
                                  <t:FieldURI FieldURI="item:HasAttachments" />
                                </t:AdditionalProperties>
                              </m:ItemShape>
                              <m:ItemIds>
                                <t:ItemId Id="{1}" />
                              </m:ItemIds>
                            </m:GetItem>
                          </soap:Body>
                        </soap:Envelope>""".format(self._soapHeader, itemId)

        response = self._SendHttp(xmlBody)
        attachmentIdList = regExAttKey.findall(response)
        return attachmentIdList

    # ----------------------------------------------------------------------------------------------------------------------
    # -----------------------------------------------Time Zone Handling-----------------------------------------------------
    # ----------------------------------------------------------------------------------------------------------------------

    def _SendHttp(self, body):
        print('647 _SendHttp soapBody=\r\n', FormatXML(body))
        body = body.encode()

        # The access token needs to be "refreshed" periodically.
        # The accessTokenCallback should return a valid token and refresh it if needed
        if self._accessTokenCallback:
            self.header = {
                'content-type': 'text/xml',
                'authorization': 'Bearer {}'.format(self._accessTokenCallback())
            }

        print('908 header=', self.header)

        try:
            resp = requests.request(
                method='POST',
                url=self.httpURL,
                data=body,
                headers=self.header
            )
            print('resp.reason=', resp.reason)
            print('resp.headers=', resp.headers)
            print('resp.text=', resp.text)
            resp.raise_for_status()

            # request = urllib.request.Request(self.httpURL, soapBody, self.header, method='POST')
            # responseString = urllib.request.urlopen(request)
            # if responseString:
            #     ret = responseString.read().decode()
            #     print('655 _SendHttp ret=', ret)
            #     self._NewConnectionStatus('Connected')
            #     return ret

            return resp.text
        except Exception as e:
            self._NewConnectionStatus('Disconnected')
            print('1293 _SendHttp Exception:\n', e)
            ProgramLog('1294 exchange_interface.py Error:' + str(e), 'error')
            if self._password:
                oldPrint('username=', self._username, ', password[-3:]=', self._password[-3:])
            # raise e

    def GetAllEvents(self):
        return self._calendarItems.copy()

    def GetEventAtTime(self, dt=None):
        # dt = datetime.date or datetime.datetime
        # return a list of events that occur on datetime.date or at datetime.datetime

        if dt is None:
            dt = datetime.datetime.now()

        events = []

        for calItem in self._calendarItems.copy():
            if dt in calItem:
                events.append(calItem)

        return events

    def GetEventsInRange(self, startDT, endDT):
        ret = []
        for item in self._calendarItems:
            if startDT <= item <= endDT:
                ret.append(item)

        return ret

    def GetNowCalItems(self):
        # returns list of calendar nowItems happening now

        returnCalItems = []

        nowDT = datetime.datetime.now()

        for calItem in self._calendarItems.copy():
            if nowDT in calItem:
                returnCalItems.append(calItem)

        return returnCalItems

    def GetNextCalItems(self):
        # return a list CalendarItems
        # will not return events happening now. only the nearest future event(s)
        # if multiple events start at the same time, all CalendarItems will be returned

        nowDT = datetime.datetime.now()

        nextStartDT = None
        for calItem in self._calendarItems.copy():
            startDT = calItem.Get('Start')
            if startDT > nowDT:  # its in the future
                if nextStartDT is None or startDT < nextStartDT:  # its sooner than the previous soonest one. (Wha!?)
                    nextStartDT = startDT

        if nextStartDT is None:
            return []  # no events in the future
        else:
            returnCalItems = []
            for calItem in self._calendarItems.copy():
                if nextStartDT == calItem.Get('Start'):
                    returnCalItems.append(calItem)
            return returnCalItems

    def GetItem(self, itemID):
        print('758 GetItem(itemID=', itemID)

        xmlBody = """<?xml version="1.0" encoding="utf-8"?>
                                <soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
                                               xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages"
                                               xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types"
                                               xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">
                                  <soap:Header>
                                    {0}
                                  </soap:Header>
                                  <soap:Body>
                                    <m:GetItem>
                                      <m:ItemShape>
                                        <t:BaseShape>IdOnly</t:BaseShape>
                                        <t:AdditionalProperties>
                                            <t:FieldURI FieldURI="item:Subject" />
                                            <t:FieldURI FieldURI="calendar:Start" />
                                            <t:FieldURI FieldURI="calendar:End" />
                                            <t:FieldURI FieldURI="calendar:Organizer"/>
                                            
                                            
                                            <t:FieldURI FieldURI="item:Body" />
                                            
                                            <t:FieldURI FieldURI="item:HasAttachments" />
                                        </t:AdditionalProperties>
                                      </m:ItemShape>
                                      <m:ItemIds>
                                        <t:ItemId Id="{1}" />
                                      </m:ItemIds>
                                    </m:GetItem>
                                  </soap:Body>
                                </soap:Envelope>""".format(
            self._soapHeader,
            itemID
        )
        response = self._SendHttp(xmlBody)
        calItems = self._CreateCalendarItemsFromResponse(response)
        print('792 calItems=', calItems)
        return calItems[0] if calItems else None


def FormatXML(xml_string):
    # todo
    return xml_string

    # xml.dom.minidom.parseString(xml_string)
    # pretty_xml_as_string = xml.toprettyxml()
    # return pretty_xml_as_string


if __name__ == '__main__':

    exchange = Exchange(
        username='gm_service_account@extrondemo.com',
        impersonation='rf_a120@extrondemo.com',
        password='Extron1025',
    )

    exchange.Connected = lambda _, state: oldPrint('Exchange', state)
    exchange.Disconnected = lambda _, state: oldPrint('Exchange', state)
    exchange.NewCalendarItem = lambda _, item: oldPrint('NewCalendarItem', item)
    exchange.CalendarItemChanged = lambda _, item: oldPrint('CalendarItemChanged', item)
    exchange.CalendarItemDeleted = lambda _, item: oldPrint('CalendarItemDeleted', item)

    while True:
        oldPrint('UpdateCalendar()')
        exchange.UpdateCalendar()
        time.sleep(10)
