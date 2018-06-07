import email, getpass, imaplib, os
from xlwt import Workbook

book = Workbook()
sheet1 = book.add_sheet('Sheet 1', cell_overwrite_ok=True)


detach_dir = '.' # directory where to save attachments (default: current)
user = 'YOUR_USERNAME'
pwd = 'YOUR_PASSWORD'

# connecting to the gmail imap server
m = imaplib.IMAP4_SSL("imap.gmail.com")
m.login(user,pwd)
m.select("[Gmail]/All Mail") # here you a can choose a mail box like INBOX instead
# use m.list() to get all the mailboxes

resp, items = m.search(None, "ALL") # you could filter using the IMAP rules here (check http://www.example-code.com/csharp/imap-search-critera.asp)
items = items[0].split() # getting the mails id

i = 0
y = 0



for emailid in items:
    resp, data = m.fetch(emailid, "(RFC822)") # fetching the mail, "`(RFC822)`" means "get the whole stuff", but you can ask for headers only, etc
    email_body = data[0][1] # getting the mail content
    mail = email.message_from_string(email_body) # parsing the mail content to get a mail object
    

  

    #print "["+mail["From"]+"] :" + mail["Subject"]
    #in excel we need the following output:
    #Name, From (email address), email, company (email.split('@')), Account name
    
    t= mail["From"]
    conditions = [t.find('indeed')== -1 and t.find('ziprecruiter')== -1 and t.find('support')== -1 and t.find('noreply')== -1 and t.find('no-reply') and \
                  t.find('cybercoders')== -1 and t.find('alerts')== -1 and t.find('linkedin')== -1 and t.find('monster')== -1 and t.find('dice')== -1 and \
                  t.find('reply')== -1 and t.find('stackoverflow')== -1 and t.find('CyberCoders')== -1 and t.find('topresume')== -1 and t.find('ihire')== -1 and t.find('example')== -1]
    
    if all(conditions):
        i = i+1
        y = y+1
        sheet1.write(i,0,t)
        sheet1.write(y,1,mail["Subject"])

book.save('vlad-testmail3.xls')

