# -*- coding: utf-8 -*-
"""
Created on Tue Jul 11 15:18:40 2023

@author: cail
"""

import requests
import os
import time
from PIL import Image
import pandas as pd
import openpyxl
from openpyxl import load_workbook
import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox

df = pd.read_excel('PrefixeSocoda (003).xlsx')
supplier_to_brands = {}
for supplier, brand in zip(df['FABRICANT'], df['MARQUE']):
    if supplier not in supplier_to_brands:
        supplier_to_brands[supplier] = []
    if brand not in supplier_to_brands[supplier]:
        supplier_to_brands[supplier].append(brand)

# Mettre à jour la liste des marques en fonction du fournisseur sélectionné
def update_brand_list(*args):
    selected_supplier = supplier_var.get()
    brand_list['menu'].delete(0, 'end')
    brands = df[df['FABRICANT'] == selected_supplier]['MARQUE'].unique()
    for brand in brands:
        brand_list['menu'].add_command(label=brand, command=tk._setit(brand_var, brand))

def get_trigram(fabricant, marque):
    trigram = df.loc[(df['FABRICANT'] == fabricant) & (df['MARQUE'] == marque), 'PREFIXE'].values
    if len(trigram) == 0:
        return "Trigramme non trouvé"
    else:
        return trigram[0]
    
def show_error_popup(message):
    root = tk.Tk()
    root.withdraw()
    messagebox.showerror("Error", message)

def printProgressBar (iteration, total, prefix = '', suffix = '', decimals = 1, length = 100, fill = '█', printEnd = "\r"):

    percent = ("{0:." + str(decimals) + "f}").format(100 * (iteration / float(total)))
    filledLength = int(length * iteration // total)
    bar = fill * filledLength + '-' * (length - filledLength)

    print(f'\r{prefix} |{bar}| {percent}% {suffix}', end = printEnd)
    # Print New Line on Complete
    if iteration == total:
        print()


def create_request(wb, four_name, marq_name, trigramme):
    start_time = time.perf_counter()
    workbook = load_workbook(wb)
    sheet = workbook['01_COMMERCE']
    z = 1 
    sheet.insert_cols(idx = sheet.max_column+1, amount=1)
    
    coltest = sheet.max_column+1
    sheet.cell(row = 1, column = coltest, value="test")
    workbook.save(wb)
    
    
    
    
    end_time = time.perf_counter()
    execution_time = end_time - start_time
    print("----------------------------------------------------------")
    print("Temps d'exécution : {:.2f} seconds".format(execution_time))
    print("----------------------------------------------------------") 
    root.destroy()

def recuperer_ltre(valeur, plage):
    for row in plage:
        for cell in row:
            if cell.value == valeur :
                column = cell.column
                ltr_column = openpyxl.utils.get_column_letter(column)
                return ltr_column
        
def open_file():
    folder_path = filedialog.askopenfilename()
    folder_entry.delete(0, tk.END)
    folder_entry.insert(0, folder_path)

def update_progressbar(pbar, value):
    pbar['value'] = value
    pbar.update()

def clear_entry():
    folder_entry.delete(0, tk.END)
    valider_button.config(state='normal')
    folder_entry.config(state='normal')
    folder_button.config(state='normal')
    
def new_tarif():
    folder_entry.delete(0, tk.END)
    valider_button.config(state='normal')
    folder_entry.config(state='normal')
    folder_button.config(state='normal')
    clear_button.config(state='normal')

def OnValidation():
    folder_path = folder_entry.get()
    fourn_name = supplier_var.get()
    marq_name = brand_var.get()
    if folder_path != "":
        print("Chemin du dossier : "+folder_path)       
        if str(fourn_name) != "" or str(fourn_name) != "-Choisir un fournisseur-":
            print("Nom du fournisseur : "+fourn_name)
            if str(marq_name) != "" or str(marq_name) !="-Choisir une marque-":
                print("Nom de la marque : "+marq_name)
            else:
                print("Le nom de la marque n'est pas spécifié")
        else:
            print("Le nom du fournisseur n'est pas spécifié")
    else:
        print("Chemin du dossier n'est pas spécifié")
    
    if folder_path != "" and fourn_name != "" and marq_name != "":
        valider_button.config(state='disabled')
        folder_entry.config(state='disabled')
        folder_button.config(state='disabled')
        clear_button.config(state = "disabled")
        
        trigramme = get_trigram(fourn_name, marq_name).lower()
        create_request(folder_path, fourn_name, marq_name, trigramme)
         
def quit_app():
    root.destroy()
    
root = tk.Tk()
root.title("REQUETATOR23.07.11")
root.geometry("450x160")

texte_param = tk.Label(root, text="Initialisation de Requetator")
texte_param.grid(row=0, column=1)

folder_label = tk.Label(root, text="Fichier Tarif :")
folder_label.grid(row=1, column=0)

folder_entry = tk.Entry(root, width=30)
folder_entry.grid(row=1, column=1)

folder_button = tk.Button(root, text="   ...   ", command=open_file)
folder_button.grid(row=1, column=2)

supplier_label = tk.Label(root, text="Fournisseur")
supplier_label.grid(row=3, column=0)

supplier_var = tk.StringVar(root)
supplier_var.set("-Choisir un fournisseur-")

supplier_list = tk.OptionMenu(root, supplier_var, *df['FABRICANT'].unique())
supplier_list.grid(row=3, column=1)

brand_label = tk.Label(root, text="Marque")
brand_label.grid(row=4, column=0)

brand_var = tk.StringVar(root)
brand_var.set("-Choisir une marque-")

brand_list = tk.OptionMenu(root, brand_var, "")
brand_list.grid(row=4, column=1)

clear_button = tk.Button(root, text="Effacer", command=clear_entry, state= 'normal')
clear_button.grid(row=6, column=1)

valider_button = tk.Button(root, text="Valider",state='normal', command=OnValidation)
valider_button.grid(row=6, column=2)

quit_button = tk.Button(root, text="Quitter", command=quit_app)
quit_button.grid(row=7, column=2)

new_button = tk.Button(root, text="Nouveau ! ", command=new_tarif)
new_button.grid(row=7, column=0)

supplier_var.trace('w', update_brand_list)

root.mainloop()
