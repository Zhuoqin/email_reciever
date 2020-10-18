import posixpath as path
from urlparse import urlparse, parse_qs, urlunparse
import email
import imaplib
import re
import webbrowser
import threading
import HTMLParser
import urllib

EMAIL = 'your email address'
PASSWORD = 'your password'
SERVER = "outlook.office365.com"

# connect to the server and go to its inbox
mail = imaplib.IMAP4_SSL(SERVER)
mail.login(EMAIL, PASSWORD)
# we choose the inbox but you can select others
mail.select('inbox')

def dealWithMailConten(mail_content):
    foundUrls = re.findall(r'<https?://.+>', mail_content, re.DOTALL)
    for url in foundUrls:
        cleanURL = url.replace("=\r\n", '')
        cleanerURLs = re.findall(r'https%3A%2F%2F.+%2F&amp;', cleanURL, re.M)
        for pureURL in cleanerURLs:
            # unescaped = HTMLParser.HTMLParser().unescape(pureURL)
            filteredURL = pureURL.replace('%2F&amp;','')
            unquoteURL = urllib.unquote(filteredURL)
            print 'opening ' + unquoteURL
            webbrowser.open(unquoteURL)

def checkEmail():
    status, data = mail.search(None, 'NEW')
    if not data or not data[0]:
        print 'No new email received...'
    else:
        print 'New email recieved...'
        mail_ids = []
        for block in data:
            mail_ids += block.split()
        for i in mail_ids:
            status, data = mail.fetch(i, '(RFC822)')
            for response_part in data:
                if isinstance(response_part, tuple):
                    message = email.message_from_string(response_part[1])
                    mail_from = message['from']
                    mail_subject = message['subject']
                    if message.is_multipart():
                        mail_content = ''
                        for part in message.get_payload():
                            if part.get_content_type() == 'text/plain':
                                mail_content += part.get_payload()
                    else:
                        mail_content = message.get_payload()

                    dealWithMailConten(mail_content)

def printit():
    threading.Timer(3, printit).start()
    checkEmail()

printit()
