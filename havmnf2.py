import email
import re
import os
import sys
import imaplib
import smtplib
import ctypes
import getpass
import quopri
import telegram
import pdfkit
import datetime
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.utils import formatdate
from email import encoders
import pandas as pd
import numpy as np
from bs4 import BeautifulSoup
from email.header import decode_header
from time import sleep

mail = imaplib.IMAP4_SSL('',993) #add the imap server here
unm = '' #add the username here
pwd = '' #add the password here
mail.login(unm,pwd)
mail.select("INBOX")
tst_contacts=[''] #add email accounts to send pdfs in test environment to.

aua_contacts=['      '] #add email accounts in procution to send emails here

mailtothis = ''# have a single email account to send the email to.
#mailtothis = 'pnl_aru@airline-choice.com'
server = smtplib.SMTP('') # add smtp server here
server.starttls()
server.login(unm,pwd)
#bot = telegram.Bot(token="394211218%3AAAF7EfA7GHw6nccUJQIUhIGlegQTWXgkBMg")
chat_id = "14185356"
def loop():
    mail.select('INBOX')
    n=0
    seat=""
    pnr=""
    bags=""
    (retcode ,messages)=mail.search(None,'(UNSEEN)')
    if retcode == 'OK':
        #bot.send_message(chat_id=chat_id,text="OK", parse_mode=telegram.ParseMode.MARKDOWN)
        for num in messages[0].split():
            n +=1
            print(n)
            typ, data = mail.fetch(num, '(RFC822)')
            raw_email = data[0][1]
            raw_email_string = raw_email.decode('utf-8')
            email_message = email.message_from_string(raw_email_string)
            email_from = email.utils.parseaddr(email_message['From'])[1]
            subject = str(email.header.make_header(email.header.decode_header(email_message['Subject'])))
            if "PNL AG" in subject :
                print(subject)
                splitted = subject.split()
                print (splitted[3])
                just_number = splitted[3].split()
                if just_number.isdigit():
                    print(just_number)
            if "PRL" in subject  :
                print("detectado")
                print(subject)
                pre,fnp = subject.split("PRL", 1)
                fnp=fnp.split()
                fname = fnp[0].replace('/','_')
                vuelo,fecha =re.split('_',fname)
                fecha = fecha+str(datetime.datetime.now().year)
                for part in email_message.walk():
                    if part.get_content_maintype() == 'multipart':
                        continue
                    if part.get('Content-Disposition') is None:
                        continue
                    fileName = part.get_filename()
                    #print(fileName)
                    if bool(fileName):
                        filePath = os.path.join('.',fileName)
                        if not os.path.isfile(filePath) :
                            fp = open(filePath, 'wb')
                            fp.write(part.get_payload(decode=True))
                            #print(type(fp))
                            fp.close()
                            list_apis = []
                            #print(textodeemail)
                            #testprl = testprl.replace('HAV','AUA')
                            #send_body ='\n'.join(textodeemail.split('\n')[3:-20])
                            with open(filePath) as f:
                                for line in f:
                                    if line[0]=="1":
                                        bagsearch = re.search("\.W/K/(.+?) ",line)
                                        if bagsearch:
                                            bags = bagsearch.groups()[0]
                                        else:
                                            bags = ""
                                        pnrsearch = re.search("\.L/(.+?) ",line)
                                        if pnrsearch:
                                            pnr = pnrsearch.groups()[0]
                                            #print (pnr)
                                        else:
                                            pnr=""
                                        seatsearch = re.search("\.R/SEAT HK1 (.+?) ",line)
                                        if seatsearch:
                                            seat = seatsearch.groups()[0]
                                            #print (seat)
                                        else:
                                            seatline = "%s" % next(f) 
                                            seatagain = re.search("\.R/SEAT HK1 (.+?) ",seatline)
                                            if seatagain:
                                               seat=seatagain.groups()[0]

                                    if line.startswith(".R/DOCS"):
                                        param, value = line.split("DOCS HK1",1)
                                        value = re.split(r'/', value)
                                        names = "%s" % next(f)         
                                        if names.startswith(".RN/"):
                                            namess = names.split(".RN/",1)
                                        elif names.startswith("1"):
                                            namess = names.split("1",1)
                                        namessplit = re.split(r'/',namess[1])
                                        #print(value)
                                        list_apis.append([pnr,namessplit[0],namessplit[1].strip(),seat,bags,value[4],value[3],value[5]])
                                    elif line.startswith(".R/PSPT"):
                                        nextline = next(f)
                                        if nextline.startswith('.RN/'):
                                            line = "%s%s" % (line.strip('\n'),nextline.strip('.RN/'))
                                        print(line)
                                        param, value = line.split("PSPT HK1",1)
                                        value = re.split(r'/',value)
                                        print(value)
                                        list_apis.append(["**INF**",value[3],value[4],'**INF**','**INF**',value[1],value[0],value[2]])
                            data_to_send = pd.DataFrame(list_apis,columns=['PNR','LASTNAME','FIRST NAME','SEAT','BAGS','NATIONALITY','DOCS','DOB'])
                            data_to_send.index += 1
                            data_to_send.replace(np.nan, '', regex=True)
                            #data_to_send['letter'] = data_to_send['SEAT'].str.extract('([A-Z]\w{0,})', expand=True)
                            #data_to_send['num'] = data_to_send['SEAT'].str.extract('([0-9])', expand=True)
                            #data_to_send['num']=data_to_send['num'].astype(int)
                            #data_to_send.to_excel(r'APIS %s.xlsx' % fname, index=None, header=True)
                            #data_to_send = data_to_send.sort_values(['num','letter'], ascending=[True,True])
                            data_to_send = data_to_send.sort_values('LASTNAME', ascending=True)
                            #print (data_to_send.num.dtype)
                            #data_to_send.drop(['num', 'letter'], axis=1, inplace=True)
                            data_to_send.reset_index()
                            data_to_send.to_html('test.html')
                            os.remove(fileName)
                            with open('test.html', 'r') as original: data = original.read()
                            with open('test.html', 'w') as modified: modified.write ("<html><head><link rel='stylesheet' type='text/css' href='table.css' media='screen' /> <body><header><div class='logoag'></div><div class='flightinfo'><h2>Passenger Manifest</h2><h2>Flight:"+vuelo+" Date:"+fecha+"  </h2></div></header>"+data+"</body></html>") 
                            PdfFilename=r'paxlist.pdf'
                            pdfkit.from_file('test.html', PdfFilename)
                            documentofinal=r'Manifiesto de Pasajeros.pdf' #% (vuelo,fechavuelo)
                            #break
                            print (data_to_send.shape[0])

                            server = smtplib.SMTP('smtp.gmail.com:587')
                            server.ehlo()
                            server.starttls()
                            server.login(unm,pwd)
                            msg=MIMEMultipart()
                            msg['From'] = "apis@arubaairlines.aw"
                            msg['To'] = "apis@arubaairlines.aw"
                            msg['Date'] = formatdate(localtime = True)
                            msg['Subject'] = "PAX Manifest %s %s " % (vuelo,fecha)
                            msg.attach(MIMEText("Attached Pax Manifest, \n Flight  %s %s" % (vuelo,fecha)))
                            
                            part2 = MIMEBase('application', "octet-stream")
                            part2.set_payload(open(PdfFilename, "rb").read())
                            encoders.encode_base64(part2)
                            part2.add_header('Content-Disposition', 'attachment; filename="PAX Manifest %s.pdf"' % fname)
                            msg.attach(part2)
                            msgit = "\r\n".join([
                                "From: apis@arubaairlines.aw",
                                "To: efrain.valles@arubaairlines.aw",
                                "Subject: APIS Test",
                                "",
                                "test",
                                ])
                            for i in aua_contacts:
                                server.sendmail('apis@arubaairlines.aw', i, msg.as_string())
                            server.quit()
                            #os.remove(fileName)
                            break 
            #for part in email_message.walk():
                #body = part.get_payload()
                #print(body)    
            #print(part)
            #replaced = body.replace('HAV  AUA  MGA','AUA  MGA')
            #cleantext = BeautifulSoup(body).text
            #body = "\n".join(filter(None,cleantext.split("\n")[3:]))
                #textodeemail = "%s" % body[0]
                #print(textodeemail)
                #textodeemail = textodeemail.replace(' HAV  AUA  MGA',' AUA  MGA')
                #textodeemail = textodeemail.replace('HAV','AUA')
                #send_body ='\n'.join(textodeemail.split('\n')[3:-20])
                #print(send_body)
                #break
            #print (send_body)

                #server = smtplib.SMTP('smtp.gmail.com:587')
                #server.ehlo()
                #server.starttls()
                #server.login(unm,pwd)
                #msgit = "\r\n".join([
                #   "From: apis@arubaairlines.aw",
                #   "To: pnl_aru@airline-choice.com",
                #   "Subject: APIS Test",
                #   "",
                #   "%s" % send_body,
                #   ])
                #server.sendmail('apis@arubaairlines.aw', mailtothis, msgit)
                #server.quit()
                #break
                #texto = """*%s*\n%s""" % (subject,body)
                #print(body),subject,email_from)
            	#bot.send_message(chat_id=chat_id,text=texto,parse_mode=telegram.ParseMode.MARKDOWN)
            for respone_part in data:
                 if isinstance (respone_part, tuple):
                     original = email.message_from_string(respone_part[1].decode())
                     #print original['From']
                     data = original['Subject']
                     #print quopri.decodestring(data)
                     typ, data = mail.store(num,'+FLAGS','\\Seen')
            continue
    #print(n)


if __name__ == '__main__':
    try:
        while True:
            loop()
    finally:
        print ("Thanks")

