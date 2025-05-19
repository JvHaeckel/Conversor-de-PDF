# Converter PDF para excel usando Python 

import pandas as pd
import tabula
import tkinter as tk 
from tkinter import filedialog, messagebox

def processamento():
    
    ################ Pegando Documento da Interface  ################
    caminho_arquivo = filedialog.askopenfilename(
        title="Selecione arquivo em PDF: ",
        filetypes=[("Arquivo PDF", "*.pdf")]
    )
     
    if not caminho_arquivo:
         messagebox.showwarning("Aviso", "Nenhum arquivo foi selecionado.")
         return
    
    ################ LENDO O PDF  ################
    try:
        
        # A função read_pdf do Tabulas lê tabelas do PDF e retorna um DataFrame
        dataframe = tabula.read_pdf(caminho_arquivo, pages='all', multiple_tables = True)
    
        # Verifica se tem tabelas no PDF
        if not dataframe:
            messagebox.showinfo("Informação", "Nenhuma tabela foi encontrada no PDF.")
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
        
    
    
    except Exception as e:
            messagebox.showerror("Erro", f"Erro ao processar o arquivo:\n{e}")
    

################ Interface gráfica ################

janela = tk.Tk()
janela.geometry("500x250")
janela.title("Conversor de Pdf")


# Título maior e em negrito
tk.Label(janela, text="Converter PDF para Excel", font=("Helvetica", 14, "bold")).pack(pady=(10, 15))

tk.Label(janela, text="Escolha o arquivo apenas em PDF:", font=("Helvetica", 12, "bold")).pack(pady=(20, 5))
botao = tk.Button(janela, text="Calcular Avos", command=processamento, font=("Helvetica", 10, "bold"))
botao.pack(pady=10)

janela.mainloop()
