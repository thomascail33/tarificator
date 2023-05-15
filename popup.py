# -*- coding: utf-8 -*-
"""
Created on Tue Mar 21 14:16:18 2023

@author: cail
"""
import os
import sys 
import tkinter as tk

def quit_app():
    root.destroy()
    
def OnValidation():
    chemin_final = sys.argv[1]
    os.system('start excel.exe "{}"'.format(chemin_final))
    root.destroy()
    
root = tk.Tk()
root.title("Voulez vous ouvrir le fichier de travail? ")
root.geometry("180x80")

main_label = tk.Label(root, text="Fichier tarif\ngénéré avec succés !", justify="center")
main_label.config(font=("Arial", 10))
main_label.grid(row=0, column=0, padx=30)

valider_button = tk.Button(root, text="Ouvrir",state='normal', justify="center", command=OnValidation)
valider_button.grid(row=1, column=0, sticky=tk.W, padx= 45)

quit_button = tk.Button(root, text="Fermer", command=quit_app, justify="center")
quit_button.grid(row=1, column=0, sticky=tk.E, padx= 45)


root.mainloop()

