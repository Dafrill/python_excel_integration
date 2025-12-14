# -*- coding: utf-8 -*-
import time
from math import sqrt
import sys
import pywintypes
from win32com.server.exception import COMException

import win32com.client
import tkinter as tk
from tkinter import Tk, Button, Label, filedialog, messagebox
from tkinter import filedialog, messagebox
from tkinter import ttk as ttk
from tkinter.ttk import Combobox

sheet_names = []
chosen_sheet_name = False
root = Tk()
root.withdraw()
file = ''
worksheet = None
workbook = None
excel = None
sheet_index = 0
sheet_names_only = None
should_hide_windows = False
should_reset_float_values = True
def click_fun(wn, _ml):
    print("wykonanie funkcji click_fun")
    top = tk.Toplevel(wn)
    top.geometry("600x240")
    top.resizable(False, False)

    style = ttk.Style()
    style.configure("My.TButton", font=('Segoe UI', 11), padding=(6, 8))

    Mlabel = ttk.Label(top, text="Projekt nr 10", font=('Segoe UI', 16, 'bold'))
    Mlabel.grid(row=0, column=0, columnspan=4, pady=(10, 15))

    grid_opt = {'padx': 10, 'pady': 5, 'sticky': "ew"}

    for i in range(4):
        top.grid_columnconfigure(i, weight=1)

    # PRZYCISKI GÓRNE
    ttk.Button(top, text="Perform calculations", width=18, style="My.TButton", command=lambda: window()).grid(row=1, column=0, **grid_opt)

    ttk.Button(top, text="About the programme", width=18, style="My.TButton", command=lambda: messagebox.showinfo("Info", "Authors: Magdalena Tałaj && Viktoria Toman")).grid(row=2, column=1, pady=10)
    ttk.Button(top, text="Close", width=18, style="My.TButton", command=top.destroy).grid(row=2, column=2, pady=10)

    top.bind("<Escape>", lambda e: top.destroy())

def mediana(lista):
    n = len(lista)
    if n==0:
        raise ValueError("The list cannot be empty")
    lista.sort()
    if n%2==0:
        m=((lista[n//2-1])+(lista[(n)//2]))/2.0
    if n%2==1:
        m=float(lista[n//2])
    return m

def srednia(lista):
    suma = 0
    for i in range(len(lista)):
        suma=suma+lista[i]
    s=suma/len(lista)
    return s

def odchylenie(lista):
    x=srednia(lista)
    n= len(lista)
    suma = 0
    for i in range(len(lista)):
        suma=suma+((lista[i]-x)**2)
    m = sqrt(suma/(n-1))
    return m

def oblicz(lista):
    wynik1 = mediana(lista)
    wynik2 = srednia(lista)
    wynik3 = odchylenie(lista)
    return f"median: {wynik1:.2f}\naverage: {wynik2:.2f}\nstandard deviation: {wynik3:.2f}"


def window():
    global root
    global sheet_names
    global chosen_sheet_name
    root.deiconify()
    root.title("Input")

    load_button = Button(root, text="Insert Excel file", command=insert_file)
    load_button.pack(pady=5)

    root.mainloop()



def get_sheet_index(sheets, sheet_name):
    for index, name in sheets.items():
        if name == sheet_name:
            return index
    return None  # Zwróć None, jeśli arkusz o podanej nazwie nie istnieje


def add_combobox(root, options):
    global worksheet
    global chosen_sheet_name
    combobox = Combobox(root, values=options)
    combobox.set("Choose sheet")  # Ustawienie wartości domyślnej
    combobox.pack(pady=10)
    button = Button(root, text="Close", command=root.destroy, font=("Arial", 12))
    button.pack(pady=10)


    # Funkcja do śledzenia wyboru użytkownika
    def on_select(event):
        global chosen_sheet_name
        global file
        global worksheet
        global excel
        global workbook
        global root
        global should_hide_windows
        float_values = []
        chosen_sheet_name = combobox.get()

        #print(f"Wybrano arkusz: {chosen_sheet_name}")
        root.withdraw()


        # Otwórz Excel i ustaw widoczność aplikacji
        excel = win32com.client.Dispatch("Excel.Application", ExcelEventHandler)
        excel.Visible = True  # Ustawienie Excela jako widocznego

        # Otwórz skoroszyt na podstawie ścieżki
        workbook = excel.Workbooks.Open(file)

        # Przejdź do wybranego arkusza i aktywuj go
        worksheet = workbook.Sheets(get_sheet_index(sheet_names, chosen_sheet_name))  # Zakładamy, że arkusz 2 to drugi arkusz w skoroszycie
        worksheet.Activate()  # Aktywuj arkusz
        selection = excel.Selection

        can_check_float_values = True
        should_break = False
        def range_is_chosen():
            root.destroy()
            root3.attributes("-topmost", True)
            root3.destroy()
            print("hello")
            print(float_values)
            w = oblicz(float_values)
            print(w)

            global should_hide_windows
            global should_break
            global should_reset_float_values
            should_reset_float_values = False
            should_break = True
            should_hide_windows = True
            result_window = Tk()
            result_window.title("Outcome")

            label = Label(result_window, text=w, font=("Arial", 14), fg="green")
            label.pack(pady=20)

            button = Button(result_window, text="Close", command=sys.exit, font=("Arial", 12))
            button.pack(pady=10)

            result_window.mainloop()



        try:
            root3 = Tk()
            root3.title("Choose range")
            root3.attributes("-topmost", True)
            root3.withdraw()
            label = Label(root3,
                          text=f"Chosen range: {selection.Address}\n ",
                          font=("Arial", 16), fg="blue")
            label2 = Label(root3, text="Close the Excel window and click the button to confirm.",font=("Arial", 16), fg="red")
            label.pack(pady=20)  # Ustawienie odstępu w pionie
            label2.pack(pady=20)
            button = Button(root3, text="Confirm", command=lambda:range_is_chosen(), font=("Arial", 14))
            button.pack(pady=20)
            if should_hide_windows:
                #print("Let's hide windows")
                root3.withdraw()
            while not should_break:


                try:
                    if can_check_float_values:
                        float_values = []
                        try:
                            for row in selection.Value:
                                for cell in row:


                                    if cell is None:
                                        can_check_float_values = False
                                        continue
                                    else:
                                        try:
                                            float_values.append(float(cell))


                                        except:
                                            can_check_float_values = False
                                            continue
                                        finally:
                                            can_check_float_values = False


                        except TypeError as e:
                            float_values.append(float(selection.Value))
                            pass

                except Exception as e:
                    can_check_float_values = False
                    print(e)
                    #print (f"selection value: {selection.Value}")
                time.sleep(0.1)


                if not selection.Address == excel.Selection.Address:

                    # if can_destroy_window:
                    #     root3.destroy()
                    selection = excel.Selection
                    label["text"] = f"Chosen range: {selection.Address}\n "
                    root3.deiconify()
                    #label["text"]=f"Zaznaczono obszar: {selection.Address}"
                    #print(f"Zaznaczono obszar: {selection.Address}")
                    can_check_float_values = True
                    root3.update()
                    can_destroy_window = True


        except pywintypes.com_error:
            pass
        except Exception as e:
            #print("The worksheet has been closed")
            print(e)
            #print(f"float_values: {float_values}")

        #worksheet.Selection.Copy()  # Może być potrzebne, by wykonać jakąś akcję w Excelu
        #print(f"Adres zaznaczonego zakresu: {worksheet.Selection.Address}")




    combobox.bind("<<ComboboxSelected>>", on_select)


# Klasa do obsługi zdarzeń Excela
class ExcelEventHandler:
    def OnSheetSelectionChange(self, sh, target):
        return target.Address


# Funkcja do wstawiania pliku Excel
def insert_file():
    global file
    global sheet_names
    global chosen_sheet_name
    global workbook
    global excel
    global sheet_names_only
    try:
        file = filedialog.askopenfilename(title="Choose Excel file",
                                          filetypes=[("Excel Files", "*.xlsx *.xls *.xlsm")])
        if file:
            excel = win32com.client.DispatchWithEvents("Excel.Application", ExcelEventHandler)
            workbook = excel.Workbooks.Open(file)
            excel.Visible = False
            # Pobranie nazw arkuszy
            sheet_names = {index + 1: sheet.Name for index, sheet in enumerate(workbook.Sheets)}
            sheet_names_only = [sheet.Name for sheet in workbook.Sheets]

            # Tworzenie nowego okna do wyboru arkusza
            root2 = Tk()
            root2.title("Choose sheet")
            add_combobox(root2, sheet_names_only)

            # Poczekaj na wybór użytkownika, zanim przejdziesz dalej
            root2.mainloop()

            # Sprawdzenie, który arkusz został wybrany

            # Wyświetl aktualny adres zaznaczenia w Excelu

            # Zamknij workbook bez zapisywania
            workbook.Close(False)
            excel.Quit()

    except pywintypes.com_error:
        pass




# Uruchomienie aplikacji
if __name__ == "__main__":
        #click_fun()
        window()