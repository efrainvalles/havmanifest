import re
import os
import sys
import ctypes
import quopri
import telegram
import pdfkit
import datetime
import pandas as pd
import numpy as np
from time import sleep
import platform
from colorama import init, Fore, Back, Style
from tinydb import TinyDB, Query
db = TinyDB('test.json')
init()
filePath = str(sys.argv[1])
list_pnrdata=[]
#print("El archivo es %s" % filePath)

#mail = imaplib.IMAP4_SSL('imap.gmail.com',993)
#unm = 'apis@arubaairlines.aw'
#pwd = 'aruba297!!'
#mail.login(unm,pwd)
#mail.select("INBOX")
#tst_contacts=['efrain.valles@arubaairlines.aw',
#              'effie.jayx@gmail.com',
#              'mvazquez@alterisd.com',
#              'mariachvp@gmail.com']

#aua_contacts=['efrain.valles@arubaairlines.aw',
#              'max.figuiera@arubaairlines.aw',
#              'ITStaff@arubaairlines.aw',
#              'franklyn.krosendijk@arubaairlines.aw',
#              'evalles@arubaairlines.aw',
#              'carlos.parra@arubaairlines.aw',
#              'edwin.gonzalez@arubaairlines.aw',
#              'francisco.arendsz@arubaairlines.aw',
#              'occ@arubaairlines.aw',
#              'jose.otano@arubaairlines.aw',
#              'jacqueline.kaersenhout@arubaairlines.aw',
#              'violeta.rustenberg@arubaairlines.aw',
#              'jasmin.wilson@arubaairlines.aw',
#              ]

#mailtothis = 'efrain.valles@arubaairlines.com'
#mailtothis = 'pnl_aru@airline-choice.com'
#server = smtplib.SMTP('smtp.gmail.com:587')
#server.starttls()
#server.login(unm,pwd)
#bot = telegram.Bot(token="394211218%3AAAF7EfA7GHw6nccUJQIUhIGlegQTWXgkBMg")
#chat_id = "14185356"
global bagi

def getprldata(prlfile):
    with open(prlfile) as f:
        for line in f:
            if line.startswith('.L/'):
                pnrsearch = line.strip("\.L/")
                if pnrsearch:
                    pnr = pnrsearch
                #    print (pnr)
                else:
                    pnr=""
            if line.startswith(".R/DOCS"):
                param, value = line.split("DOCS HK1",1)
                nextline = next(f)
                print (nextline)
                names = "%s" % nextline
                #print(names)
                namess = names.split(".RN/")
                testit = "".join(namess)
                print(testit)
                #print(namess)
                newvalue = "%s%s" % (value.strip('\n'),testit)
                print(newvalue)
                #print(truevalue)
                finalvalue = re.split(r'/', newvalue)
                #print (finalvalue)
                #print(value)
            if line.startswith('.W/K/'):
                bagsearch = line.strip("\.W/K/")
                bagit=re.search("\.W/K/(.+?)/0$",line)
                if bagit.groups()[0]:
                    bagi=bagit.groups()[0]
                    print(bagi)
                else:
                    bagi="000"
                    print("esto esta asi %s" % (bagi))
            if line.startswith('.R/SEAT'):
                seatsearch = re.search("\.R/SEAT (.+?)$",line)
                if seatsearch:
                    seat = seatsearch.groups()[0]
                    seat=seat.strip('0')
                list_pnrdata.append([pnr.strip('\n'),finalvalue[8].strip('\n'),finalvalue[9].strip('\n'),finalvalue[4],finalvalue[3],finalvalue[5],seat, bagi])
    return list_pnrdata


def printlist(listpnr):
    print(listpnr)


lospnrs=getprldata(filePath)
printlist(lospnrs)

#print (len(list_apis))
#for pax in list_apis:
#    print (pax)


