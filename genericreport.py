# disabled username / password logon for use in our Exchange environment

######### Setup your stuff here #######################################

host = 'webmail.whatever.com' # specify port, if required, using a colon and port number following the hostname

fromaddr = 'donotreply@whatever.com' # must be a vaild 'from' address in your environment
toaddr  = ['donotreply@whatever.com'] # list of email addresses
ccaddr  = [''] # list of email addresses
bccaddr  = ['bccaddr@whatever.com'] # list of email addresses
replyto = fromaddr # unless you want a different reply-to

# username = 'username' # not used in our Exchange environment
# password = 'password' # not used in our Exchange environment

msgsubject = 'Subject of email'
htmlmsgtext = "<h2>Name of Report for "  # text with appropriate HTML tags
connstring = 'DRIVER={SQL Server};SERVER=srvname;DATABASE=dbname;UID=user;PWD=password'

######### In normal use nothing changes below this line ###############

import smtplib, pyodbc, sys
from email.MIMEMultipart import MIMEMultipart
from email.MIMEText import MIMEText
from email.Utils import COMMASPACE
from HTMLParser import HTMLParser
import datetime as dt
yesterday = dt.datetime.now() - dt.timedelta(days=1)

date = "'" + yesterday.strftime('%m-%d-%Y') + "'"
htmlmsgtext = htmlmsgtext + date + "</h2>"
conn=pyodbc.connect(connstring)
cursor=conn.cursor()
cursor.execute("exec dbo.spStoredProcedure " + date)
rows = cursor.fetchall()
column_names = [d[0] for d in cursor.description]
cursor.close()
del cursor
if len(rows) == 0:
    sys.exit
htmlmsgtext = htmlmsgtext + '<table style="border:2px solid black">\n'
htmlmsgtext = htmlmsgtext + '<tr>\n'
for column_name in column_names:
    htmlmsgtext = htmlmsgtext + '<th style="border:1px solid black; text-align:center">' + column_name + '</th>'
htmlmsgtext = htmlmsgtext + '\n</tr>\n'
for row in rows:    
    for column in row:
        htmlmsgtext = htmlmsgtext + '<td style="border:1px solid black; text-align:center">' + str(column) + '</td>'
    htmlmsgtext = htmlmsgtext + '\n</tr>\n'
htmlmsgtext = htmlmsgtext + '</table>\n'

# A snippet - class to strip HTML tags for the text version of the email

class MLStripper(HTMLParser):
    def __init__(self):
        self.reset()
        self.fed = []
    def handle_data(self, d):
        self.fed.append(d)
    def get_data(self):
        return ''.join(self.fed)

def strip_tags(html):
    s = MLStripper()
    s.feed(html)
    return s.get_data()

########################################################################


try:
    # Make text version from HTML - First convert tags that produce a line break to carriage returns
    msgtext = htmlmsgtext.replace('</br>',"\r").replace('<br />',"\r").replace('</p>',"\r")
    # Then strip all the other tags out
    msgtext = strip_tags(msgtext)

    # necessary mimey stuff
    msg = MIMEMultipart()
    msg.preamble = 'This is a multi-part message in MIME format.\n'
    msg.epilogue = ''

    body = MIMEMultipart('alternative')
    body.attach(MIMEText(msgtext))
    body.attach(MIMEText(htmlmsgtext, 'html'))
    msg.attach(body)   
    
    msg['From'] = fromaddr
    msg['To'] = COMMASPACE.join(toaddr)
    msg['CC'] = COMMASPACE.join(ccaddr)
    msg['Subject'] = msgsubject
    msg['Reply-To'] = replyto
    
    print 'To addresses follow:'
    print toaddr

    # The actual email sendy bits
    server = smtplib.SMTP(host)
    server.set_debuglevel(False) # set to True for verbose output
    recipients = toaddr + ccaddr + bccaddr
    # Comment this block and uncomment the below try/except block if TLS or user/pass is required.
    server.sendmail(fromaddr, recipients, msg.as_string())
    print 'Email sent.'
    server.quit() # bye bye
	
    # try:
        # # If TLS is used
        # server.starttls()
        # server.login(username,password)
        # server.sendmail(fromaddr, recipients, msg.as_string())
        # print 'Email sent.'
        # server.quit() # bye bye
    # except:
        # # if tls is set for non-tls servers you would have raised an exception, so....
        # server.login(username,password)
        # server.sendmail(fromaddr, recipients, msg.as_string())
        # print 'Email sent.'
        # server.quit() # bye bye        
        	
except:
    print "Email NOT sent to %s successfully. ERR: %s %s %s " % (str(toaddr), str(sys.exc_info()[0]), str(sys.exc_info()[1]), str(sys.exc_info()[2]) )
    #just in case   
