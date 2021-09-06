#!/usr/bin/env python
#
# Adapted from https://gist.github.com/robulouski/7442321
#Python Script to download RedFire reciept emails that are
#auto-forwarded to a gmail account

import sys
import imaplib
import getpass
import email
import os
from bs4 import BeautifulSoup
import re
from datetime import date

TODAY = date.today()
TOT_MATCH = "Net Total $"
SUBJECT_MATCH = "Batch Receipt: "
HEADER = 'Date,Batch ID,Charge Transactions,Charge Total,Refund Transactions,Refund Total,Net Total'

IMAP_SERVER = 'imap.gmail.com'

# Set email here or pass it as first argument
EMAIL_ACCOUNT = ""
if not EMAIL_ACCOUNT:
    EMAIL_ACCOUNT = str(sys.argv[1])

# Chose folder below
# EMAIL_FOLDER = "INBOX"
# EMAIL_FOLDER = '"[Gmail]/Sent Mail"' # note the double quoting
EMAIL_FOLDER = '"X-BB"'

# Set password here or the script will get it from cmd line
PASSWORD = ""
if not PASSWORD:
    PASSWORD = getpass.getpass()

def append_csv(entry):
    """
    append the data set to a file in csv format
    """
    v = []
    for value in entry.values():
        v.append(value)
    
    csv_row = f"{v[0]},{v[1]},{v[2]},{v[3]},{v[4]},{v[5]},{v[6]}"
    with open("revenue.csv", mode='a+') as file_obj:
        file_obj.seek(0, os.SEEK_SET)
        firstline = file_obj.readline()
        if re.match("^Date", firstline):
             file_obj.write(csv_row + '\n')
        else:
            file_obj.seek(0, os.SEEK_END)
            file_obj.write(HEADER + '\n')
            file_obj.write(csv_row + '\n')

def process_mailbox(M):
    """
   Process each text/html section of the multipart email
    """
    rv, data = M.search(None, "ALL")
    if rv != 'OK':
        print ("No messages found!")
        return
    
    for num in data[0].split():
        csv_dict = {"Date":[],"Batch ID":[],"Charge Transactions":[],"Charge Total":[],"Refund Transactions":[],"Refund Total":[],"Net Total":[]};
        parts = []
        data1 = []
        rv, data = M.fetch(num, '(RFC822)')
        msg = email.message_from_bytes(data[0][1])
        for part in msg.walk():
            if part.get_content_type() == 'text/html':
                parts.append(part.get_payload())
    
            for p in parts:
                soup = BeautifulSoup(p,features='html5lib')
                table = soup.find('table')
                t = table.findAll('tr')
            for row in t:
                cols = row.findAll('td')
                cols = [ele.text.strip() for ele in cols]
                data1.append([ele for ele in cols if ele])

            '''Each needed element must be parsed individually and saved to dict'''                
            f1 = str(data1[4])
            b,c = f1.split(',')
            c = b.strip('[\'')
            e,f = c.split('=\\n')
            csv_dict["Date"] = e+f

            f2 = str(data1[5])
            fb,fc = f2.split(',')
            fe,ff = fc.split('=\\n')
            fes = fe.strip(' \'')
            fef = ff.strip('\']')
            csv_dict["Batch ID"] = fes+fef

            f3 = str(data1[6])[-3]
            csv_dict["Charge Transactions"] = int(f3)

            f4 = str(data1[7])
            fb,fc = f4.split('$')
            fc = fc.strip('\']')
            csv_dict["Charge Total"] = float(fc)

            f5 = str(data1[8])
            fb,fc = f5.split(',')
            if "d" in fc:
                fc = fc.strip('<=\\n/td\>\']')
                fd, fe = fc.split('\'')
                csv_dict["Refund Transactions"] = int(fe)
            elif "=" not in fc:
                fc = fc.strip('\=\']')
                fd, fe = fc.split('\'')
                csv_dict["Refund Transactions"] = int(fe)
            else:
                fc = int(fc[-4])
                csv_dict["Refund Transactions"] = fc

            f6 = str(data1[9])
            fb,fc = f6.split('$')
            fc = fc.strip('\']')
            csv_dict["Refund Total"] = float(fc)

            f7 = str(data1[10])
            fb,fc = f7.split('$')
            fc = fc.strip('\']')
            csv_dict["Net Total"] = float(fc)
        
        append_csv(csv_dict)
        #Remove email after processing
        M.store(num, '+FLAGS', r'(\Deleted)')

def main():
    M = imaplib.IMAP4_SSL(IMAP_SERVER)
    M.login(EMAIL_ACCOUNT, PASSWORD)
    rv, data = M.select(EMAIL_FOLDER)
    if rv == 'OK':
        print("Processing mailbox: ", EMAIL_FOLDER)
        csv_line = process_mailbox(M)
        M.close()
    else:
        print("ERROR: Unable to open mailbox ", rv)
    M.logout()

if __name__ == "__main__":
    main()
