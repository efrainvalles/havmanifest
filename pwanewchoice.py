from pywinauto.application import Application

import sys
import re
import platform
from time import sleep
from colorama import init, Fore, Back, Style
import pandas as pd
from tinydb import TinyDB, Query
db = TinyDB('checker.json')
init()


print ( Back.BLUE+Fore.YELLOW+Style.BRIGHT+"\n\nKIU Checker - CHOICE TO KIU Script Prototipo por Efrain Valles\n\n\
Se inicio KIU, Ingrese sus datos de inicio en KIU en \n\n\
Para iniciar precione esta ventana (s) \n"+Style.RESET_ALL)

windows=platform.platform()


#Iniciando KIU
try:
    if windows.startswith("Windows-7"):
        app = Application().connect(title_re=".*KIU*", class_name="TApplication")
        dlg = app.TApplication
    elif windows.startswith("Windows-10"):
        app = Application().connect(title_re=".*KIU*", class_name="KIU")
        print("Es Windows-10")
        dlg = app.KIU
except:    
    app = Application().start("C:\Reservas\clienteres.exe")
    if windows.startswith("Windows-7"):
        dlg = app.TApplication
    elif windows.startswith("Windows-10"):
        print("Es Windows-10")
        dlg = app.KIU

try:
    if windows.startswith("Windows-7"):
        app2 = Application().connect(title_re=".*python*", class_name="ConsoleWindowClass")
    elif windows.startswith("Windows-10"):
        app2 = Application().connect(title_re="Command Prompt - python*", class_name="ConsoleWindowClass")
except:    
    print("Error en Capturar")

dlg.window()
dlg.move_window(x=600, y=0, width=600, height=400, repaint=True)

dlg2 = app2.ConsoleWindowClass
dlg2.move_window(x=600, y=0, width=600, height=400, repaint=True)


filePath = str(sys.argv[1])
list_apis=[]
list_checked=[]
print("El archivo es %s" % filePath)
try:
    df = pd.read_csv(filePath)
except UnicodeDecodeError:
#    df = pd.read_csv(filePath,encoding = "ISO-8859-1")
    df = pd.read_csv(filePath,engine='python')
    pass

vuelo = df['Unnamed: 7'][1][:3]
fechavuelo = df['Unnamed: 7'][3][:11]
dep_airp = df['Unnamed: 7'][2][:3]
dest_airp = df['Unnamed: 13'][2][:3]
df.drop(df.head(5).index,inplace=True)
df.drop(df.tail(2).index,inplace=True)
clsdf=df.drop(['Unnamed: 1','Unnamed: 3','Unnamed: 5','Unnamed: 7','Passenger Manifest','Unnamed: 13','Unnamed: 16','Unnamed: 17','Unnamed: 18'], axis=1)
clsdf=clsdf.rename(columns={'Unnamed: 0':'chk order','Unnamed: 2': 'pnr','Unnamed: 4': 'last','Unnamed: 6': 'first','Unnamed: 9': 'middle','Unnamed: 10': 'dob','Unnamed: 11': 'gen','Unnamed: 12': 'dest','Unnamed: 14': 'seat','Unnamed: 15':'bags',})
clsdf.reset_index(drop=True,inplace=True)
print(clsdf)
vueloinfo= "Vuelo: AG %s - %s %s %s" % (vuelo,fechavuelo,dep_airp,dest_airp)
print("para el vuelo %s" % (vueloinfo))
breaklist=""
#check if it is in the database

FLTQ=Query()
if db.search(FLTQ.archivo==filePath):
    print("The File is in")
    resultdb = db.search(FLTQ.archivo == filePath)
    print(resultdb)
    breaklist= resultdb[0].get("last_checkin")
else:
    db.insert({'archivo': filePath,'last_checkin':'0'})


def searchpax(rowitem,orden):
    print ("Orden llego asi" + orden)
    try:
        if orden == "R":
            dlg.type_keys("PF-%s{ENTER}" % (rowitem['last'][:5]))
        if orden =="":
            dlg.type_keys("PF-%s{ENTER}" % (rowitem['last'].replace(" ","{VK_SPACE}")))
    except:
        print("Error", sys.exc_info())
        pass
    sleep(1)
    print(Style.BRIGHT)
    print("Nombre Completo:")
    print(Back.BLUE+Fore.YELLOW+"%s/%s" % (rowitem['last'],rowitem['first']))
    print(Style.RESET_ALL)
    print(Style.BRIGHT)
    dlg2.set_focus()
    ordenit = input("\nIngrese el orden ([#],[s](saltar pax), [c] si ya estaba chequeado [r] para repetir con menos letras:")
    ordenit=ordenit.upper()
    return (ordenit)

def checkit(listofpax):
    list_pending=pd.DataFrame(columns = ['chk order' , 'pnr', 'last' , 'first','middle' , 'dob','gen', 'dest' , 'seat', 'bags'])
    
    for index, row in listofpax.iterrows():
        orden=""
        print (row)
        db.update({'last_checkin': index}, FLTQ.archivo == filePath)
        #indi= listofpax.index(i)
        #if i[0] == "**INF**":
        #    print("pax es infante, chequee manual \n %s" % (i))
        #    continue
        sleep(1)
        while orden == "R" or  orden == "":
            print("orden es" + orden)
            orden=orden.upper()
            orden = searchpax(row,orden)
            print ("La orden al final loop es " + orden)
            
        print("orden es AQUI CHIQUI " + orden)
        if orden == "S":
            print("*** SALTANDO PASAJERO ***")
            sleep(1)
            data = {'chk order': row['chk order'],
                    'pnr': row['pnr'],
                    'last': row['last'],
                    'first': row['first'],
                    'middle': row['middle'],
                    'dob': row['dob'],
                    'gen': row['gen'],
                    'dest': row['dest'],
                    'seat': row['seat'],
                    'bags': row['bags']}
            list_pending = list_pending.append(data, ignore_index=True)
            
            print (list_pending)
            print (len(list_pending))
            dlg2.set_focus()
            continue
        if orden == "C":
            print(Back.WHITE+Fore.RED+Style.BRIGHT+"*** pax ya estaba chequeado ***"+Style.RESET_ALL)
            list_checked.append(row)
            dlg2.set_focus()
        if orden =="R":
            print(Back.BLUE+Fore.WHITE+Style.BRIGHT+"*** buscando con menos caracteres ***"+Style.RESET_ALL)
            searchpax(row,orden)
        if orden == "NO":
            break
            continue
        if row['bags'] is not "0" and orden[0].isdigit():
            try:
                dlg.type_keys("PU%s%s,ST%s{ENTER}" % (orden,row['gen'],row['seat']))
            except:
                print('error')
                pass
            sleep(1)
            list_checked.append(row)
            try:
                dlg2.set_focus()
            except:
                print("error")
                pass
        elif row['bags'] is "0" and orden[0].isdigit():
            dlg.type_keys("PU%s%s,ST%s{ENTER}" % (orden,row['gen'],row['seat']))
            list_checked.append(row, ignore_index(True))
            sleep(1)
            try:
                dlg2.set_focus()
            except:
                print("error")
                pass
    return list_pending

inicio = input("desea iniciar? (s/n) :")
if inicio == "s": 
    print("Iniciando checkin backoffice - %s" % (vueloinfo))
else:
    sys.exit()
sleep(1)

if breaklist is "":
    breaklist = input("desea iniciar desde un orden especifico #?  :")


    
print (len(list_apis))
if breaklist is not "":
    print ("Esta es  %d" % (int(breaklist))) 
    #del list_apis[0:int(breaklist)]
    clsdf.drop(clsdf.index[0:(int(breaklist))],inplace=True)
    print(clsdf)
else:
    print("iniciando desde el principio")


print("\n\n Ingrese el Header del Vuelo en KIU y presione s para iniciar")
iniciar = input("iniciar (s/n): ")

if iniciar == "s":
    listofpaxtocheck = checkit(clsdf)

while len(listofpaxtocheck) > 0:
    print("Estos son los pasajeros pendientes")
    for index, row in listofpaxtocheck.iterrows():
        print(row)
    chequear=input("desea chequearlos?")
    if chequear == "s" or chequear == "S":
	    listofpaxtocheck=checkit(listofpaxtocheck)