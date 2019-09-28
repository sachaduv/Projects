import win32com.client
import xlwt
import re
import datetime

outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
# 3  Deleted Items
# 4  Outbox
# 5  Sent Items
# 6  Inbox
# 9  Calendar
# 10 Contacts
# 11 Journal
# 12 Notes
# 13 Tasks
# 14 Drafts
inbox = outlook.GetDefaultFolder(6) 
messages = inbox.Items
message = messages.GetFirst()
row =1
wk_date = message.CreationTime.strftime("%Y %m %d").split(" ")
year = int(wk_date[0])
month = int(wk_date[1])
day = int(wk_date[2])
week = datetime.date(year,month,day).isocalendar()[1]


wb = xlwt.Workbook()
sheetname = "WSR_WK"+str(week)
filename = "WSR_WK"+str(week)+".xls"
ws = wb.add_sheet(sheetname)

ws.write(0,0,'Requestor Name')
ws.write(0,1,'Release name')
ws.write(0,2,'Rel No')
ws.write(0,3,'Tasks')
ws.write(0,4,'Clarity Task')
ws.write(0,5,'Functionality')
ws.write(0,6,'Defect')
ws.write(0,7,'Support to Systems')
ws.write(0,8,'Environment')
ws.write(0,9,'Planned Start Date')
ws.write(0,10,'Planned End Date')
ws.write(0,11,'Actual Start Date')
ws.write(0,12,'Actual End Date')
ws.write(0,13,'STATUS')
ws.write(0,14,'Test Owner')
ws.write(0,15,'Remark')

while message:
    msg_dict = {}
    to = message.To
    if(to.find(';')==-1):
        pass
    else:
        to = to.split(";")[0]
    external_mail = re.compile(r'\(External.*|\(Capgemini.*')
    to = external_mail.sub("",to)
    msg_dict['To']=to

    msg_dict['Subject']=message.Subject

    date=message.CreationTime
    msg_dict['Date']=date.strftime("%m/%d/%Y")
    dt = msg_dict['Date'].split("/")
    wk = datetime.date(int(dt[2]),int(dt[0]),int(dt[1])).isocalendar()[1]
    if(wk!=week):
        break
    sender = message.SenderName
    sender = external_mail.sub("",sender)
    msg_dict['Sender']=sender

    ws.write(row,0,msg_dict['To'])
    ws.write(row,3,msg_dict['Subject'])
    if(message.Subject.find("SOF")==-1):
        ws.write(row,4,'Release Verification')
    else:
        ws.write(row,4,'Testing SOF')
    defect = re.compile('[Dd]efect.{,3}[0-9]+')
    defect = defect.findall(message.Subject)
    ws.write(row,6,defect)
    ws.write(row,9,msg_dict['Date'])
    ws.write(row,10,msg_dict['Date'])
    ws.write(row,11,msg_dict['Date'])
    ws.write(row,12,msg_dict['Date'])
    ws.write(row,13,'completed')
    ws.write(row,14,msg_dict['Sender'])

    body = message.Body.lower()
    body_list = body.splitlines()

    def filter_mail(data):
        eliminate = [""," "]
        if(data not in eliminate):
            return True
        else:
            return False

    body_list = list(filter(filter_mail,body_list))
    body_list = body_list[:10]
    body = "\n".join(body_list)
    env_regex = r'environment\s*:\s*(.*)'
    try:
        env_match = re.search(env_regex,body).group(1)
        ws.write(row,1,env_match.upper())
        ws.write(row,8,env_match.upper())
    except:
        pass

    
    ord_regex = re.compile(r'orders are processed|order is processed|order is not found|orders are not found|order is cancelled|orders are cancelled|executable order id')
    ord_match = ord_regex.findall(body)
    sync_regex = re.compile(r'stock picture|sync|stock report')
    sync_match = sync_regex.findall(body)
    stock_regex = re.compile(r'inventory|stock details|stock update|stock updates')
    stock_match = stock_regex.findall(body)
    adjustment_regex = re.compile(r'adjustments|adjustment')
    adjustment_match = adjustment_regex.findall(body)
    supplier_regex = re.compile(r'supplier details')
    supplier_match = supplier_regex.findall(body)
    cancel_regex = re.compile(r'cancellation')
    cancel_match = cancel_regex.findall(body)
    reservation_regex = re.compile(r'reservation|transaction|transactions|modification|back ordered|reservations')
    reservation_match = reservation_regex.findall(body)
    allocation_regex = re.compile(r'allocation|allocations')
    allocation_match = allocation_regex.findall(body)
    basedata_regex = re.compile(r'base data|base load|data load')
    basedata_match = basedata_regex.findall(body)

    if(len(ord_match)>0):
        ws.write(row,5,"Order fulfillment")
        ws.write(row,2,'ISOM')
        ws.write(row,7,'ISOM')
        ws.write(row,15,'Confirmed to ISOM')
    elif(len(sync_match)>0):
        ws.write(row,5,"Full sync")
        ws.write(row,2,'ISOM')
        ws.write(row,7,'ISOM')
        ws.write(row,15,'Confirmed to ISOM')
    elif(len(stock_match)>0):
        ws.write(row,5,"Inventory details")
        ws.write(row,2,'ISOM')
        ws.write(row,7,'ISOM')
        ws.write(row,15,'Confirmed to ISOM')
    elif(len(adjustment_match)>0):
        ws.write(row,5,"Stock adjustment")
        ws.write(row,2,'ASTRO')
        ws.write(row,7,'ASTRO')
        ws.write(row,15,'Confirmed to ASTRO')
    elif(len(supplier_match)>0):
        ws.write(row,5,"Supplier details")
        ws.write(row,2,'ISOM')
        ws.write(row,7,'ISOM')
        ws.write(row,15,'Confirmed to ISOM')
    elif(len(reservation_match)>0):
        ws.write(row,5,"Stock Reservation")
        ws.write(row,2,'ISOM')
        ws.write(row,7,'ISOM')
        ws.write(row,15,'Confirmed to ISOM')
    elif(len(cancel_match)>0):
        ws.write(row,5,"Cancellation")
        ws.write(row,2,'ISOM')
        ws.write(row,7,'ISOM')
        ws.write(row,15,'Confirmed to ISOM')
    elif(len(allocation_match)>0):
        ws.write(row,5,"Stock Allocation")
        ws.write(row,2,'GEMINI')
        ws.write(row,7,'GEMINI')
        ws.write(row,15,'Confirmed to GEMINI')
    elif(len(basedata_match)>0):
        ws.write(row,5,"Base data")
        ws.write(row,2,'ISOM')
        ws.write(row,7,'ISOM')
        ws.write(row,15,'Confirmed to ISOM')
    else:
        ws.write(row,5,"NA")
        ws.write(row,2,'NA')
        ws.write(row,7,'NA')
        ws.write(row,15,'NA')

    with open("mail_body.txt", 'a') as out:
        out.write(body)
        out.write("\n")
        out.write("-----------------------Next--------------------------------")
        out.write("\n")
    message = messages.GetNext()
    row=row+1

wb.save(filename)