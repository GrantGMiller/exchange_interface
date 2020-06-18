from lxml.builder import ElementMaker
import lxml.etree as ET
import datetime

MSG_NS = 'http://schemas.microsoft.com/exchange/services/2006/messages'
TYPE_NS = 'http://schemas.microsoft.com/exchange/services/2006/types'
SOAP_NS = 'http://schemas.xmlsoap.org/soap/envelope/'
NAMESPACES = {'m': MSG_NS, 't': TYPE_NS, 's': SOAP_NS}
M = ElementMaker(namespace=MSG_NS, nsmap=NAMESPACES)
T = ElementMaker(namespace=TYPE_NS, nsmap=NAMESPACES)
EXCHANGE_DATE_FORMAT = '%Y-%m-%dT%H:%M:%SZ'


class Event:
    pass


event = Event()
event.subject = 'Subject'
event.body = 'Body'
event.location = 'location@email.com'

start = datetime.datetime.now()
end = datetime.datetime.now() + datetime.timedelta(hours=1)
start_date = datetime.date.today().isoformat()
end_date = (datetime.date.today() + datetime.timedelta(days=1)).isoformat()

calendar = 'calendar@email.com'
timezone = 'Eastern Standard Time'
invitation = 'SendToNone'
impersonate = 'impersonate@email.com'

impersonation = T.ExchangeImpersonation(T.ConnectingSID(T.PrimarySmtpAddress(impersonate)))

root = M.CreateItem(
    M.SavedItemFolderId(T.DistinguishedFolderId(calendar and T.Mailbox(T.EmailAddress(calendar)) or '', Id='calendar')),
    M.Items(T.CalendarItem(T.Subject(event.subject), T.Body(event.body or '', BodyType='HTML'),
                           T.Start(start.strftime(EXCHANGE_DATE_FORMAT)), T.End(end.strftime(EXCHANGE_DATE_FORMAT)),
                           T.Location(event.location or ''), T.MeetingTimeZone(TimeZoneName=timezone))),
    SendMeetingInvitations=invitation)

event_id = 'event_id123'
body_type='Text'
getItems = M.GetItem(
    M.ItemShape(T.BaseShape('IdOnly'), T.BodyType(body_type), T.AdditionalProperties(T.FieldURI(FieldURI='item:Body'))),
    M.ItemIds(T.ItemId(Id=event_id)))

fid = 'folderID123'
findItems = M.FindItem(M.ItemShape(T.BaseShape('IdOnly'), T.AdditionalProperties(T.FieldURI(FieldURI='item:Subject'), T.FieldURI(FieldURI='calendar:Organizer'), T.FieldURI(FieldURI='calendar:Location'), T.FieldURI(FieldURI='calendar:Start'), T.FieldURI(FieldURI='calendar:End'))), M.CalendarView(StartDate=start_date, EndDate=end_date), M.ParentFolderIds(T.FolderId(Id=fid)), Traversal='Shallow')

findFolder = M.FindFolder(M.FolderShape(T.BaseShape('IdOnly'), T.AdditionalProperties(T.FieldURI(FieldURI='folder:DisplayName'))), M.IndexedPageFolderView(MaxEntriesReturned='100', Offset='0', BasePoint='Beginning'), M.ParentFolderIds(T.DistinguishedFolderId(Id='msgfolderroot')), Traversal='Deep')

print('*********************')
print()
# print(ET.tostring(impersonation))
# print(ET.tostring(root))
# print(ET.tostring(getItems))
# print(ET.tostring(findItems))
print(ET.tostring(findFolder))
