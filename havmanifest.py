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
mga_contacts=[''] #add email accounts to send the PDF to here. separate them with comas


mailtothis = '' #add single account to send EMAIL TO HERE
#mailtothis = 'pnl_aru@airline-choice.com'
server = smtplib.SMTP('') #add the smtp server here
server.starttls()
server.login(unm,pwd)
#bot = telegram.Bot(token="394211218%3AAAF7EfA7GHw6nccUJQIUhIGlegQTWXgkBMg")
chat_id = "14185356"
def loop():
    mail.select('TEST')
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
                for part in email_message.walk():
                    if part.get_content_maintype() == 'multipart':
                        continue
                    if part.get('Content-Disposition') is None:
                        continue
                    fileName = part.get_filename()
                    print(fileName)
                    
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
                                        list_apis.append([pnr,namessplit[1].strip(),namessplit[0],seat,bags,value[4],value[3],value[5]])
                                    elif line.startswith(".R/PSPT"):
                                        continue
                                        #param, value = line.split("PSPT HK1",1)
                                        #value = re.split(r'/',value)
                                        #list_apis.append([value[4],value[3],value[1],value[0],value[2]])
                            data_to_send = pd.DataFrame(list_apis,columns=['PNR','FIRTS NAME','LAST NAME','SEAT','BAGS','NATIONALITY','DOCS','DOB'])
                            data_to_send.index += 1
                            data_to_send.replace(np.nan, '', regex=True)
                            #data_to_send.to_excel(r'APIS %s.xlsx' % fname, index=None, header=True)
                            data_to_send.to_html('test.html')
                            with open('test.html', 'r') as original: data = original.read()
                            with open('test.html', 'w') as modified: modified.write ("<html><head><link rel='stylesheet' type='text/css' href='table.css' media='screen' /><body><h2> Lista de Pasajeros</h2>"+data+"</body></html>") 
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
                            msg['Subject'] = "Apis Final %s " % fname
                            msg.attach(MIMEText("Saludos, \n Anexo APIS Finales vuelo %s" % fname))
                            
                            part2 = MIMEBase('application', "octet-stream")
                            part2.set_payload(open(PdfFilename, "rb").read())
                            encoders.encode_base64(part2)
                            part2.add_header('Content-Disposition', 'attachment; filename="Manifiesto %s.pdf"' % fname)
                            msg.attach(part2)
                            msgit = "\r\n".join([
                                "From: apis@arubaairlines.aw",
                                "To: efrain.valles@arubaairlines.aw",
                                "Subject: APIS Test",
                                "",
                                "test",
                                ])
                            #server.sendmail('apis@arubaairlines.aw', 'efrain.valles@arubaairlines.aw', msg.as_string())
                            server.quit()
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

