#start
import gspread #type:ignore NOQA
import imaplib
import email
import json

#connecting to GSeets and selecting the file
sa = gspread.service_account(filename="sa_creds.json")

#read credentials from json
with open("ym_creds.json", "r") as f:
    jmCreds = json.load(f)

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

receiptSummaryDataFULL = []
receiptSummaryData = []
def readMaximaReceiptSummary(EmailId: str, receipt):
    receiptSummaryData.clear()
    for i in range(0, 11):
        receiptSummaryData.append('')
    receiptSummaryData[0] = EmailId
    for i in receipt:
        #Receipt_id
        if '<pre>Kvitas ' in i:
            receiptSummaryData[1] = i.split()[1]
        #BankReceipt_id
        if 'KVITO NR' in i:
            receiptSummaryData[2] = i.split()[2]
        #Receipt_Document#
        if 'DOKUMENTO NR' in i:
            receiptSummaryData[3] = i.split()[2]
        #RRN
        if 'RRN:' in i:
            receiptSummaryData[4] = i.split()[1]
        #Date
        if 'LTF NM ' in i:
            receiptSummaryData[5] = f'{i.split()[3]}-{i.split()[4]}-{i.split()[5]}'
        #Time
        if 'LTF NM ' in i:
            receiptSummaryData[6] = i.split()[6]
        #GroceryBrand
        if 'MAXIMA LT' in i:
            receiptSummaryData[7] = 'MAXIMA LT'
        #Address
        if 'MAXIMA LT, UAB' in i:
            receiptSummaryData[8] = i.split('<br />')[1].split(' Kasa Nr.')[0][:-1]
        #TotalAmount
        if 'Apsipirkimo suma:' in i:
            receiptSummaryData[9] = i.split(': ')[1][:-4]
        #TotalDiscount
        if 'Kvito nuolaidų suma:' in i:
            receiptSummaryData[10] = i.split(': ')[1][:-4]
    receiptSummaryDataFULL.append(list(receiptSummaryData))

processedEmails = []
processedEmail = []
def emailProcessLog(GroceryBrand: str, Email_ids):
    for i in Email_ids:
        processedEmail = [GroceryBrand, i]
        processedEmails.append(list(processedEmail))

#writing email summary
for i in maximaToAdd:
    status, email_data = imap_server.fetch(i, "(RFC822)")
    email_message = email.message_from_bytes(email_data[0][1])
    
    msg1 = email_message.get_payload()[0]
    msg1body = msg1.get_payload(decode=True)
    singleEmail = msg1body.decode().split('\r\n')
    readMaximaReceiptSummary(i, singleEmail)
wksSummaryMaxima.append_rows(receiptSummaryDataFULL)

#writing processed emails to Emails sheet
emailProcessLog('Maxima',maximaToAdd)
wksEmails.append_rows(processedEmails)

# Disconnect from the server
imap_server.close()
imap_server.logout()