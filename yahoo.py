import gspread
import imaplib
import email
import json
from decimal import Decimal
from datetime import date, datetime
import os

def logging(logFile, logMessage):
# Get the current date and time in UTC
    current_time = datetime.utcnow()

# Convert the datetime object to a string
    time_string = current_time.strftime("%Y-%m-%d %H:%M:%S")

# Open the text file in append mode
    with open(logFile, "a") as file:
    # Write the time and "Started" to the file
        file.write(f'{time_string} {logMessage}\n')

log = 'logfile.txt'
logging(log, 'Started')

#connecting to GSeets and selecting the file
sa = gspread.service_account(filename="sa_creds.json")

#read credentials from json
with open("ym_creds.json", "r") as f:
    jmCreds = json.load(f)

#Receipt Summary Data:
receiptSummaryDataFULL = []
receiptSummaryData = []
def readMaximaReceiptSummary(EmailId: str, receipt):
    receiptSummaryData.clear()
    for i in range(0, 12):
        receiptSummaryData.append('')
    receiptSummaryData[0] = EmailId
    for i in receipt:
        #Receipt_id
        if '<pre>Kvitas ' in i:
            receiptSummaryData[1] = i.split()[1]
        #BankReceipt_id
        if 'KVITO NR' in i or 'Kvito nr' in i:
            receiptSummaryData[2] = i.split()[2]
        #Receipt_Document#
        if 'DOKUMENTO NR' in i:
            receiptSummaryData[3] = i.split()[2]
        #RRN
        if 'RRN' in i:
            receiptSummaryData[4] = i.split()[1]
        #Date
        if 'Inv. Nr.' in i:
            receiptSummaryData[5] = i.split()[2]
        if 'LTF ' in i:
            receiptSummaryData[6] = f'{i.split()[3]}-{i.split()[4]}-{i.split()[5]}'
        #Time
        if 'LTF ' in i:
            receiptSummaryData[7] = i.split()[6]
        #GroceryBrand
        if 'MAXIMA LT' in i:
            receiptSummaryData[8] = 'MAXIMA LT'
        #Address
        if 'MAXIMA LT, UAB' in i:
            receiptSummaryData[9] = i.split('<br />')[1].split(' Kasa Nr.')[0][:-1]
        #TotalAmount
        if 'Apsipirkimo suma:' in i:
            receiptSummaryData[10] = i.split(': ')[1][:-4]
        #TotalDiscount
        if 'Kvito nuolaidų suma:' in i:
            receiptSummaryData[11] = i.split(': ')[1][:-4]
    receiptSummaryDataFULL.append(list(receiptSummaryData))

#Processed Emails:
processedEmails = []
processedEmail = []
def emailProcessLog(GroceryBrand: str, Email_ids):
    processDate = date.today().strftime("%Y-%m-%d")
    for i in Email_ids:
        status, email_data = imap_server.fetch(i, "(RFC822)")
        email_message = email.message_from_bytes(email_data[0][1])        
        
        encoded_subject = email_message['Subject']
        decoded_subject = email.header.decode_header(encoded_subject)[0][0]
        if isinstance(decoded_subject, bytes):
            decoded_subject = decoded_subject.decode('utf-8')
        
        if decoded_subject == 'Jūsų apsipirkimo MAXIMOJE kvitas':
            processedEmail = [GroceryBrand, i, processDate]
            processedEmails.append(list(processedEmail))
        else:
            processedEmail = [f'{GroceryBrand} Other', i, processDate]
            processedEmails.append(list(processedEmail))

#Receipt items data:
items = []
def readMaximaReceiptItems(EmailId: str, receipt):
    
    receiptMainInfo = receipt.decode().split('\r\n')
    for i in receiptMainInfo:
        #Receipt_id
        if '<pre>Kvitas ' in i:
            receiptID = i.split()[1]
        #Date
        if 'LTF ' in i:
            receiptDate = f'{i.split()[3]}-{i.split()[4]}-{i.split()[5]}'
        #Time
        if 'LTF ' in i:
            receiptTime = i.split()[6]


    receiptDepic = receipt.decode().split('<pre>Kvitas')
    receiptDepic = receiptDepic[1].split('======================================================')
    receiptDepic = receiptDepic[0].replace(' N\r\n', ' N A\r\n')
    receiptDepic = receiptDepic.replace(' A\r\n','|')
    receiptDepic = receiptDepic.replace('\r\n','|',1)
    receiptDepic = receiptDepic.replace('&#160;',' ')
    receiptDepic = ' '.join(receiptDepic.split())
    receiptDepic = receiptDepic.split('|')
    receiptDepic = receiptDepic[1:len(receiptDepic)-1]

    for i in receiptDepic:
        item = [receiptDate,receiptTime,EmailId,receiptID]
        itemSplit = i.split()
        fullPrice = itemSplit[len(itemSplit)-1]

        if 'X' in itemSplit:
            xIndex = itemSplit.index('X')
            theItem = ' '.join(itemSplit[:xIndex-1])
            unitPrice = itemSplit[xIndex-1].replace(',','.')
            quantity = itemSplit[xIndex+1].replace(',','.')
            measure = itemSplit[xIndex+2]
        else:
            theItem = ' '.join(itemSplit[:-1])
            unitPrice = fullPrice.replace(',','.')
            quantity = ''
            measure = ''

        if fullPrice == 'N':
            fullPrice = itemSplit[len(itemSplit)-2].replace(',','.')
            items[-1][-2] = fullPrice
        elif '-' in fullPrice:
            fullPrice = fullPrice[1:].replace(',','.')
            items[-1][9] = fullPrice
            items[-1][10] = f'{str(round(float(fullPrice)/float(items[-1][8])*100,1))}%'
            items[-1][11] = str(round(float(items[-1][8])-float(fullPrice),2))
        else:
            fullPrice = fullPrice.replace(',','.')
            item.append(theItem)
            item.append(unitPrice)
            item.append(quantity)
            item.append(measure)
            item.append(fullPrice)
            item.append('0.00')
            item.append('0.0%')
            item.append(fullPrice)
            item.append('0.00')
            item.append(i)
            #adding item to items
            items.append(list(item))

groceryBrandDoneEmails = [] #list to write to GSheets
def groceryBrandEmails(GroceryBrand: str, emailList):
    for i in emailList:
        f = [GroceryBrand]
        f.append(i)
        groceryBrandDoneEmails.append(f)

sh = sa.open("Grocery")
#selecting sheets:
wksEmails = sh.worksheet("Emails")
wksSummaryMaxima = sh.worksheet("MaximaSummarized")
wksItemsMaxima = sh.worksheet("MaximaItems")

#getting data from wksEmails sheet
# wksEmails.get()
maximaAlreadyIn = [i[1] for i in wksEmails.get() if 'Maxima' in i[0]]

# Connect to the Yahoo Mail IMAP server
imap_server = imaplib.IMAP4_SSL("imap.mail.yahoo.com")
# Login to the account
imap_server.login(jmCreds["email"], jmCreds["password"])
# Select the "Inbox" folder
imap_server.select("Inbox")

# Search for all emails: status, email_ids = imap_server.search(None, "ALL")
# Search for all emails from MAXIMA
status, email_ids = imap_server.search(None, "FROM noreply.code.provider@maxima.lt")
maximaEmails = email_ids[0].decode().split()

maximaToAdd = [i for i in maximaEmails if i not in maximaAlreadyIn]

#writing email summary
for i in maximaToAdd:
    status, email_data = imap_server.fetch(i, "(RFC822)")
    email_message = email.message_from_bytes(email_data[0][1])
    
    encoded_subject = email_message['Subject']
    decoded_subject = email.header.decode_header(encoded_subject)[0][0]
    if isinstance(decoded_subject, bytes):
        decoded_subject = decoded_subject.decode('utf-8')
    
    if decoded_subject == 'Jūsų apsipirkimo MAXIMOJE kvitas':
        msg1 = email_message.get_payload()[0]
        msg1body = msg1.get_payload(decode=True)
        singleEmail = msg1body.decode().split('\r\n')
        readMaximaReceiptSummary(i, singleEmail)
        readMaximaReceiptItems(i, msg1body)

#Write Summary data
wksSummaryMaxima.append_rows(receiptSummaryDataFULL)
#Write detailed receipt data
wksItemsMaxima.append_rows(items)

#writing processed emails to Emails sheet
emailProcessLog('Maxima',maximaToAdd)
wksEmails.append_rows(processedEmails)

# Disconnect from the server
imap_server.close()
imap_server.logout()

logging(log, 'Finished')