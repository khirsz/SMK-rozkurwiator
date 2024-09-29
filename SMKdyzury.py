"""
SMK Rozkurwiator dyżurów
@author: Samuel Mazur & KH
"""

import os
import sys
from time import sleep
import selenium.common.exceptions
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
import pandas as pd
import numpy as np
import tkinter

def dzialanie(dyzury, driver):
    WBwait = WebDriverWait(driver, 50, poll_frequency=1)
    #znajdz przycisk
    przyciskDodaj = None
    for przycisk in driver.find_elements(By.XPATH, '//button[text()="Dodaj"]'):
        if przycisk.is_displayed():
            przyciskDodaj = przycisk
            break
    if przyciskDodaj is None:
        print("Could not find clickable button! Make sure it is visible in the browser")
        return

    for dyzur in dyzury:
        if dyzur.data is np.nan:
            continue
        #dodaj
        WBwait.until(EC.element_to_be_clickable(przyciskDodaj)).click()
        #godziny
        WBwait.until(EC.element_to_be_clickable((By.XPATH, f"//tbody/tr[last()]/td[4]/div[1]/input[@value='']"))).send_keys(dyzur.godziny)
        #minuty
        WBwait.until(EC.element_to_be_clickable((By.XPATH, f"//tbody/tr[last()]/td[5]/div[1]/input[@value='']"))).send_keys(dyzur.minuty)
        #data
        WBwait.until(EC.element_to_be_clickable((By.XPATH, f"//tbody/tr[last()]/td[6]/div[1]/input[@value='']"))).send_keys(dyzur.data)
        #Nazwa komórki organizacyjnej
        WBwait.until(EC.element_to_be_clickable((By.XPATH, f"//tbody/tr[last()]/td[8]/div[1]/input[@value='']"))).send_keys(dyzur.nazwa)

class Konfiguracja:
    def __init__(self):
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
        self.LokalizacjaLbl = tkinter.Label(stepOne,text="Lokalizacja arkuszy")
        self.LokalizacjaLbl.grid(row=6, column=0, sticky='E', padx=5, pady=2)
        self.LokalizacjaTxt = tkinter.Entry(stepOne)
        self.LokalizacjaTxt.grid(row=6, column=1, columnspan=3, pady=2, sticky='WE')

        self.konfiguracja = Konfiguracja()

        self.GuzikWysylania = tkinter.Button(stepOne, text="Wyslij",command=self.wyslij)
        self.GuzikWysylania.grid(row=7, column=3, sticky='W', padx=5, pady=2)

        self.protocol("WM_DELETE_WINDOW", sys.exit)

    def wyslij(self):
        self.konfiguracja.lokalizacja = self.LokalizacjaTxt.get()
        if self.konfiguracja.lokalizacja == "":
            self.LokalizacjaTxt.config({"background": "Red"})
            return
        else:
            self.LokalizacjaTxt.config({"background": "White"})

        self.LokalizacjaTxt.delete(0, tkinter.END)
        self.LokalizacjaTxt.insert(0, "")
        self.quit()

class Dyzur:
    def __init__(self, data, nazwa, godziny = "24", minuty = ""):
        self.data = data
        self.nazwa = nazwa
        self.godziny = godziny
        self.minuty = minuty

def zaladujDyzury(lokalizacja):
    dyzury = []

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

        godzinyKol = None
        minutyKol = None
        dataKol = None
        nazwaKol = None
        for kolumna in df.columns:
            nazwaUjednolicona = str(kolumna).strip().lower()
            if dataKol is None and "data" in nazwaUjednolicona:
                dataKol = pd.to_datetime(df[kolumna], errors='coerce').dt.strftime("%Y-%m-%d")
            elif minutyKol is None and "minut" in nazwaUjednolicona:
                minutyKol = df[kolumna]
            elif godzinyKol is None and "godzin" in nazwaUjednolicona:
                godzinyKol = df[kolumna]
            elif nazwaKol is None and "nazwa" in nazwaUjednolicona:
                nazwaKol = df[kolumna]

        if dataKol is None or nazwaKol is None:
            print(f"Incorrect data in file {lista[plik]}", file=sys.stderr)
            continue

        for i in range(df.shape[0]):
            try:
                aktualneGodziny = str(int(godzinyKol[i]))
            except:
                aktualneGodziny = "24"

            try:
                aktualneMinuty = str(int(minutyKol[i]))
            except:
                aktualneMinuty = ""

            #zaladowanie danych
            dyzury.append(Dyzur(dataKol[i], nazwaKol[i], aktualneGodziny, aktualneMinuty))

    return dyzury

def main():
    try:
        driver = webdriver.Chrome()
    except selenium.common.exceptions.WebDriverException:
        print("Fallback to Firefox")
        driver = webdriver.Firefox()
    driver.maximize_window()
    driver.get("https://smk.ezdrowie.gov.pl/login.jsp?locale=pl")
    okno = Okno(None)
    okno.title("SMK Rozkurwiator Dyżurów 0.2")
    while True:
        okno.mainloop()
        if okno.konfiguracja.lokalizacja is None:
            sleep(1)
            continue
        dyzury = zaladujDyzury(okno.konfiguracja.lokalizacja)
        if dyzury:
            dzialanie(dyzury, driver)
            okno.konfiguracja.lokalizacja = None

if __name__ == "__main__":
    main()
