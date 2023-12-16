"""
SMK ROZKURWIATOR 0.7
@author: Samuel Mazur $ KH
"""

import os
import sys
from time import sleep
import selenium.common.exceptions
from selenium import webdriver
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
import pandas as pd
import numpy as np
import tkinter
from dateutil.relativedelta import relativedelta
from dateutil.parser import parse as parseDate

def dzialanie(pacjenci, konfiguracja, driver):
    WBwait = WebDriverWait(driver, 50, poll_frequency=1)
    #znajdz przycisk
    przyciskDodaj = None
    for przycisk in driver.find_elements(By.XPATH, '//tbody/tr/td[1]/button[text()="Dodaj"]'):
        if przycisk.is_displayed():
            przyciskDodaj = przycisk
            break
    if przyciskDodaj is None:
        print("Could not find clickable button! Make sure it is visible in the browser")
        return

    for pacjent in pacjenci:
        #dodaj
        WBwait.until(EC.element_to_be_clickable(przyciskDodaj)).click()
        #data
        WBwait.until(EC.element_to_be_clickable((By.XPATH, "//tbody/tr[1]/td[2]/div[1]/input[1]"))).send_keys(pacjent.data)
        #rok szkolenia
        rokSzkolenia = Select(WBwait.until(EC.element_to_be_clickable((By.XPATH,"//tbody/tr[1]/td[4]/div[1]/select[1]"))))
        rokSzkolenia.select_by_value(obliczRokSzkolenia(konfiguracja.rok, pacjent.data))
        #kod zabiegu
        kodZabiegu = Select(WBwait.until(EC.element_to_be_clickable((By.XPATH,"//tbody/tr[1]/td[5]/div[1]/select[1]"))))
        kodZabiegu.select_by_index(str(int(konfiguracja.kod)-1))
        #nazwisko
        WBwait.until(EC.element_to_be_clickable((By.XPATH, "//tbody/tr[1]/td[6]/div[1]/input[1]"))).send_keys(konfiguracja.osoba)
        #miejsce
        miejscestazu_element = WBwait.until(EC.element_to_be_clickable((By.XPATH, "//tbody/tr[1]/td[7]/div[1]/select[1]")))
        miejscestazu = Select(miejscestazu_element)
        try:
            miejscestazu.select_by_index(konfiguracja.miejsce)
        except selenium.common.exceptions.WebDriverException:
            for j in range(int(konfiguracja.miejsce)):
                miejscestazu_element.send_keys(Keys.ARROW_DOWN)
        #nazwastazu
        nazwaStazu_element = WBwait.until(EC.element_to_be_clickable((By.XPATH, "//tbody/tr[1]/td[8]/div[1]/select[1]")))
        nazwaStazu = Select(nazwaStazu_element)
        try:
            nazwaStazu.select_by_index(konfiguracja.nazwa)
        except selenium.common.exceptions.WebDriverException:
            for j in range(int(konfiguracja.nazwa)):
                nazwaStazu_element.send_keys(Keys.ARROW_DOWN)
        #inicjaly
        WBwait.until(EC.element_to_be_clickable((By.XPATH, "//tbody/tr[1]/td[9]/div[1]/input[1]"))).send_keys(pacjent.inicjaly)
        #plec
        plec = Select(WBwait.until(EC.element_to_be_clickable((By.XPATH,"//tbody/tr[1]/td[10]/div[1]/select[1]"))))
        plec.select_by_value(pacjent.plec)
        #asysta
        WBwait.until(EC.element_to_be_clickable((By.XPATH, "//tbody/tr[1]/td[11]/div[1]/input[1]"))).send_keys(pacjent.asysta)
        #nazwaproc
        WBwait.until(EC.element_to_be_clickable((By.XPATH, "//tbody/tr[1]/td[12]/div[1]/input[1]"))).send_keys(pacjent.usluga)

def obliczRokSzkolenia(rokSzkoleniaLubData, data):
    if len(rokSzkoleniaLubData) <= 2:
        #tekst to prawdopodobnie tylko rok
        return rokSzkoleniaLubData

    dataSzkolenia = parseDate(rokSzkoleniaLubData)
    dataProcedury = parseDate(data)
    return str(relativedelta(dataProcedury, dataSzkolenia).years+1)

class Konfiguracja:
    def __init__(self):
        self.rok = None
        self.kod = None
        self.osoba = None
        self.miejsce = None
        self.nazwa = None
        self.lokalizacja = None

class Okno(tkinter.Tk):
    def __init__(self, parent):
        tkinter.Tk.__init__(self,parent)
        self.parent = parent
        self.initialize()

    def initialize(self):
        self.grid()
        stepOne = tkinter.LabelFrame(self, text=" Wypełnij zgodnie z instrukcją ")
        stepOne.grid(row=0, columnspan=7, sticky='W',padx=5, pady=5, ipadx=5, ipady=5)
        self.RokLbl = tkinter.Label(stepOne, anchor="e", justify="right", text="Rok szkolenia\n(Lub data rozpoczecia w przypadku trybu automatcznego)")
        self.RokLbl.grid(row=0, column=0, sticky='E', padx=5, pady=2)
        self.RokTxt = tkinter.Entry(stepOne)
        self.RokTxt.grid(row=0, column=1, columnspan=3, pady=2, sticky='WE')

        self.KodLbl = tkinter.Label(stepOne,text="Kod zabiegu (które miejsce na liscie)")
        self.KodLbl.grid(row=1, column=0, sticky='E', padx=5, pady=2)
        self.KodTxt = tkinter.Entry(stepOne)
        self.KodTxt.grid(row=1, column=1, columnspan=3, pady=2, sticky='WE')

        self.OsobaLbl = tkinter.Label(stepOne,text="Osoba wykonująca")
        self.OsobaLbl.grid(row=2, column=0, sticky='E', padx=5, pady=2)
        self.OsobaTxt = tkinter.Entry(stepOne)
        self.OsobaTxt.grid(row=2, column=1, columnspan=3, pady=2, sticky='WE')

        self.MiejsceLbl = tkinter.Label(stepOne,text="Miejsce wykonania (które miejsce na liscie)")
        self.MiejsceLbl.grid(row=3, column=0, sticky='E', padx=5, pady=2)
        self.MiejsceTxt = tkinter.Entry(stepOne)
        self.MiejsceTxt.grid(row=3, column=1, columnspan=3, pady=2, sticky='WE')

        self.NazwaLbl = tkinter.Label(stepOne,text="Nazwa stażu (które miejsce na liscie)")
        self.NazwaLbl.grid(row=4, column=0, sticky='E', padx=5, pady=2)
        self.NazwaTxt = tkinter.Entry(stepOne)
        self.NazwaTxt.grid(row=4, column=1, columnspan=3, pady=2, sticky='WE')

        self.LokalizacjaLbl = tkinter.Label(stepOne,text="Lokalizacja arkuszy")
        self.LokalizacjaLbl.grid(row=6, column=0, sticky='E', padx=5, pady=2)
        self.LokalizacjaTxt = tkinter.Entry(stepOne)
        self.LokalizacjaTxt.grid(row=6, column=1, columnspan=3, pady=2, sticky='WE')

        self.konfiguracja = Konfiguracja()

        self.GuzikWysylania = tkinter.Button(stepOne, text="Wyslij",command=self.wyslij)
        self.GuzikWysylania.grid(row=7, column=3, sticky='W', padx=5, pady=2)

        self.protocol("WM_DELETE_WINDOW", sys.exit)

    def wyslij(self):
        self.konfiguracja.rok = self.RokTxt.get().strip()
        if self.konfiguracja.rok == "":
            self.RokTxt.config({"background": "Red"})
            return
        else:
            self.RokTxt.config({"background": "White"})

        self.konfiguracja.kod = self.KodTxt.get().strip()
        if self.konfiguracja.kod == "":
            self.KodTxt.config({"background": "Red"})
            return
        else:
            self.KodTxt.config({"background": "White"})

        self.konfiguracja.osoba = self.OsobaTxt.get().strip()
        if self.konfiguracja.osoba == "":
            self.OsobaTxt.config({"background": "Red"})
            return
        else:
            self.OsobaTxt.config({"background": "White"})

        self.konfiguracja.miejsce = self.MiejsceTxt.get().strip()
        if self.konfiguracja.miejsce == "":
            self.MiejsceTxt.config({"background": "Red"})
            return
        else:
            self.MiejsceTxt.config({"background": "White"})

        self.konfiguracja.nazwa = self.NazwaTxt.get().strip()
        if self.konfiguracja.nazwa == "":
            self.NazwaTxt.config({"background": "Red"})
            return
        else:
            self.NazwaTxt.config({"background": "White"})

        self.konfiguracja.lokalizacja = self.LokalizacjaTxt.get().strip()
        if self.konfiguracja.lokalizacja == "":
            self.LokalizacjaTxt.config({"background": "Red"})
            return
        else:
            self.LokalizacjaTxt.config({"background": "White"})

        self.LokalizacjaTxt.delete(0, tkinter.END)
        self.LokalizacjaTxt.insert(0, "")
        self.quit()

class Pacjent:
    def wypelnij(self, data, imie, nazwisko, plec = None, usluga = "", asysta = ""):
        self.data = data
        self.inicjaly = f"{imie.strip()[0].upper()}.{nazwisko.strip()[0].upper()}."
        if plec is None:
            #identyfikacja plci
            if nazwisko.strip().upper().endswith("A"):
                self.plec = "K"
            else:
                self.plec = "M"
        else:
            self.plec = plec.strip()[0].upper()
        self.asysta = asysta
        self.usluga = usluga

    def wypelnijInicjalami(self, data, inicjaly, plec, usluga = "", asysta = ""):
        self.data = data
        inicjalyUjednolicone = inicjaly.strip().upper().replace(".", "")
        self.inicjaly = f"{inicjalyUjednolicone[0]}.{inicjalyUjednolicone[1]}."
        self.plec = plec.strip()[0].upper()
        self.asysta = asysta
        self.usluga = usluga

def zaladujPacjentow(lokalizacja):
    pacjenci = []

    try:
        lokalizacjaUjednolicona = os.path.abspath(lokalizacja)
        lista = os.listdir(lokalizacjaUjednolicona)
    except:
        print("Incorrect path to files!")
        return
    for plik in range(len(lista)):
        try:
            xls_file = pd.ExcelFile(os.path.join(lokalizacjaUjednolicona, lista[plik]))
        except ValueError:
            print(f"Cannot read file {lista[plik]}", file=sys.stderr)
            continue
        df = xls_file.parse()
        #usun nan i inf
        df = df.replace([np.nan, -np.inf], "")

        nazwiskoKol = None
        imieKol = None
        dataKol = None
        plecKol = None
        asystaKol = None
        inicjalyKol = None
        uslugaKol = None
        for kolumna in df.columns:
            nazwaUjednolicona = str(kolumna).strip().lower()
            if dataKol is None and "data" in nazwaUjednolicona:
                dataKol = df[kolumna].dt.strftime("%Y-%m-%d")
            elif imieKol is None and "imi" in nazwaUjednolicona:
                imieKol = df[kolumna]
            elif nazwiskoKol is None and "nazwisko" in nazwaUjednolicona:
                nazwiskoKol = df[kolumna]
            elif plecKol is None and ("plec" in nazwaUjednolicona or "płeć" in nazwaUjednolicona):
                plecKol = df[kolumna]
            elif asystaKol is None and "asysta" in nazwaUjednolicona:
                asystaKol = df[kolumna]
            elif inicjalyKol is None and ("inicjaly" in nazwaUjednolicona or "inicjały" in nazwaUjednolicona):
                inicjalyKol = df[kolumna]
            elif uslugaKol is None and ("usluga" in nazwaUjednolicona or "usługa" in nazwaUjednolicona):
                uslugaKol = df[kolumna]

        if dataKol is None:
            print(f"Date column not found in file {lista[plik]}", file=sys.stderr)
            continue

        if imieKol is None or nazwiskoKol is None:
            if inicjalyKol is None:
                print(f"Patient data not found in file {lista[plik]}", file=sys.stderr)
                continue
            elif plecKol is None:
                print(f"Patient gender data not found in file {lista[plik]}", file=sys.stderr)
                continue

        for i in range(df.shape[0]):
            pacjent = Pacjent()
            if asystaKol is not None:
                aktualnaAsysta = asystaKol[i]
            else:
                aktualnaAsysta = ""
            if uslugaKol is not None:
                aktualnaUsluga = uslugaKol[i]
            else:
                aktualnaUsluga = ""
            #zaladowanie danych
            if inicjalyKol is None:
                if plecKol is not None:
                    aktualnaPlec = plecKol[i]
                else:
                    aktualnaPlec = None
                pacjent.wypelnij(data=dataKol[i], imie=imieKol[i], nazwisko=nazwiskoKol[i],
                                                plec=aktualnaPlec, usluga=aktualnaUsluga, asysta=aktualnaAsysta)
            else:
                pacjent.wypelnijInicjalami(data=dataKol[i], inicjaly=inicjalyKol[i],
                                                plec=plecKol[i], usluga=aktualnaUsluga, asysta=aktualnaAsysta)

            pacjenci.append(pacjent)

    return pacjenci

def main():
    try:
        driver = webdriver.Chrome()
    except selenium.common.exceptions.WebDriverException:
        print("Fallback to Firefox")
        driver = webdriver.Firefox()
    driver.maximize_window()
    driver.get("https://smk.ezdrowie.gov.pl/login.jsp?locale=pl")
    okno = Okno(None)
    okno.title("SMK Rozkurwiator 0.7")
    while True:
        okno.mainloop()
        if okno.konfiguracja.lokalizacja is None:
            sleep(1)
            continue
        pacjenci = zaladujPacjentow(okno.konfiguracja.lokalizacja)
        if pacjenci:
            dzialanie(pacjenci, okno.konfiguracja, driver)
            okno.konfiguracja.lokalizacja = None

if __name__ == "__main__":
    main()
