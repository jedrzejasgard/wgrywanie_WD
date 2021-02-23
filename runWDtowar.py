#-*- coding: utf-8 -*-
import requests
import json
from vendoasg.vendoasg import Vendo
import pandas as pd
import os

vendoApi = Vendo("http://vendo.asgard.pl:5560")
vendoApi.logInApi("esklep","e12345")
vendoApi.loginUser("jpawlewski","jp12345")

#************ nazwa pliku do wgrania **************
nPlik = input('Wpisz nazwę pliku do wgrania:  ')
formatPliku = input('Wpisz format pliku do wgrania (xls/xlsx):  ')
with open(f"{nPlik}.{formatPliku}",'rb')as tabelaDane:
    plik = pd.read_excel(tabelaDane).fillna('')

#print(plik)
#plik = xlrd.open_workbook(f"{nPlik}.{formatPliku}")
total_cols = plik.shape[1]
total_rows = plik.shape[0]
print(total_cols,total_rows)
cols = 1
# tworzenie sownika nazw wartosci dowolnych
dictWD = {}
req = vendoApi.getJson ('/DB/WartosciDowolne', {"Token":vendoApi.USER_TOKEN,"Model":{"ObiektTypDanych":"towar","ObiektyID":[19061],"ZwrocZawartoscPlikow":True,"ZwrocPusteWartosci":True}})
req = req['Wynik']['Rekordy'][0]['Wartosci']
for item in req:
    op = item['Opis']
    na = item['Nazwa']
    if item['Opis'] in dictWD:
        pass
    else:
        dictWD[op] = na
#print(dictWD)


for column in plik.columns[1:]:
    i = 0
    nazwa_col = column
    print(nazwa_col)
    nazwa_WD = dictWD[nazwa_col]
    print(nazwa_WD)
    for index,row in plik.iterrows():
        try:
            kod = plik.loc[i,'Kod']
            kod = str(kod)
            if "." in kod:
                kod = kod[:-2]
            if len(kod) == 4:
                kod = "0" + kod
            print(kod)
            kod_query = vendoApi.getJson ('/Magazyn/Towary/Towar', {"Token":vendoApi.USER_TOKEN,"Model":{"Towar":{"Kod":kod}}})
            #print(f"kod - {kod_query}")
            numerID = kod_query["Wynik"]["Towar"]["ID"]
            wartosc = plik.loc[i,nazwa_col]
            try:
                print(wartosc,len(wartosc))
                if len(wartosc)==0:
                    print('############         Brak wartości ide dalej')
                    pass
            except:
                pass
            
            else:
                response_data = vendoApi.getJson ('/json/reply/Magazyn_Towary_Aktualizuj', {"Token":vendoApi.USER_TOKEN,"Model":{"ID":numerID,"PolaUzytkownika":{"NazwaWewnetrzna":nazwa_WD ,"Wartosc":wartosc}}})
                print("Dodaję KOD: ", kod, " nazwaWD: ", nazwa_WD, " wartość: ", wartosc)
            try:
                dodane = open(f"dodane{nPlik}.txt", 'a')
                dodane.write("Dodaję KOD: " + kod + " | ")
                dodane.write("nazwaWD: " + nazwa_WD + " | ")
                dodane.write("wartość: " + wartosc + "\n")
                dodane.close()
            except UnicodeEncodeError:
                pass
            i += 1
        except KeyError:
            print(f"Błąd dodawania: ", kod)
            err = open(f"errors{nPlik}.txt", 'a')
            err.write(kod + "\n")
            err.close()
            i += 1
            
    cols += 1

wgrany_plik = f"C:\\Users\\asgard_48\\Documents\Skrypty\\Wgrywanie wartośći dowolnych\\{nPlik}.xlsx"
plik_do_archiwum = f'C:\\Users\\asgard_48\\Documents\Skrypty\\Wgrywanie wartośći dowolnych\\Archiwum\\{nPlik}.xlsx'
os.replace( wgrany_plik, plik_do_archiwum )

