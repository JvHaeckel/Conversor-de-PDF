# Converter PDF para excel usando Python 

import pandas as pd
import tabula
import tkinter as tk 
from tkinter import filedialog, messagebox

arquivo = r("")
dataframe = tabula



################ Interface gráfica ################

janela = tk.Tk()
janela.geometry("500x250")
janela.title("Conversor de Pdf")

# Título maior em negrito

# Título maior e em negrito
tk.Label(janela, text="Converter PDF para Excel", font=("Helvetica", 14, "bold")).pack(pady=(10, 15))


tk.Label(janela, text="Escolha o arquivo apenas em PDF:", font=("Helvetica", 12, "bold")).pack(pady=(20, 5))
botao = tk.Button(janela, text="Calcular Avos", command=processar, font=("Helvetica", 10, "bold"))
botao.pack(pady=10)

janela.mainloop()

