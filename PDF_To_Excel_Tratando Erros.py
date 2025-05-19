# Converter PDF para excel usando Python 

import pandas as pd
import tabula
import tkinter as tk 
from tkinter import filedialog, messagebox, ttk

################ Barra de Progresso Funções ################

def start_progress():
    progressbar['value'] = 0
    progressbar.pack(pady=10)
    botao.config(state='disabled')
    janela.update_idletasks()

def stop_progress():
    progressbar.pack_forget()
    botao.config(state='normal')
    janela.update_idletasks()

################ Pegando Documento da Interface + LENDO O PDF ################

def processamento():
    start_progress()

    ################ Pegando Documento da Interface  ################
    caminho_arquivo = filedialog.askopenfilename(
        title="Selecione arquivo em PDF: ",
        filetypes=[("Arquivo PDF", "*.pdf")]
    )
     
    if not caminho_arquivo:
         messagebox.showwarning("Aviso", "Nenhum arquivo foi selecionado.")
         stop_progress()
         return
    
    ################ LENDO O PDF  ################
    try:
        # A função read_pdf do Tabulas lê tabelas do PDF e retorna um DataFrame
        dataframe = tabula.read_pdf(
            caminho_arquivo, 
            pages='all', 
            multiple_tables = True
            )
    
        # Verifica se tem tabelas no PDF
        if not dataframe:
            messagebox.showinfo("Informação", "Nenhuma tabela foi encontrada no PDF.")
            stop_progress()
            return
        
        # Combina todas as tabelas em um único DataFrame
        dataframe_complete = pd.concat(dataframe)
        
         # Abre a janela para salvar o arquivo, forçando extensão .xlsx
        caminho_saida = filedialog.asksaveasfilename(
            defaultextension=".xlsx" ,
            filetypes=[("Arquivos Excel", "*.xlsx")],
            title="Salvar arquivo como:"
        )
        
        if caminho_saida:
           # Converte o Dataframe para Excel
            dataframe_complete.to_excel(caminho_saida, index=False)
            messagebox.showinfo("Sucesso", f"Arquivo salvo com sucesso em:\n{caminho_saida}")
        else:
            messagebox.showwarning("Cancelado", "Operação cancelada pelo usuário.")
        
        # Tratando erro de deixar arquivo do excel a ser salvo em Aberto.
    except PermissionError:
        messagebox.showerror("Permisssão negada", "Você deve estar mantendo arquivo em excel de destino em aberto")    
        
    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao processar o arquivo:\n{e}")

    stop_progress()

################ Tkinter Básico ################

janela = tk.Tk()
janela.geometry("500x250")
janela.title("Conversor de Pdf")

# Título maior e em negrito
tk.Label(janela, text="Converter PDF para Excel", font=("Helvetica", 14, "bold")).pack(pady=(10, 15))

tk.Label(janela, text="Escolha o arquivo apenas em PDF:", font=("Helvetica", 12, "bold")).pack(pady=(20, 5))
botao = tk.Button(janela, text="Converter PDF", command=processamento, font=("Helvetica", 10, "bold"))
botao.pack(pady=10)

progressbar = ttk.Progressbar(janela, orient='horizontal', length=300, mode='determinate')
progressbar.pack_forget()

janela.mainloop()