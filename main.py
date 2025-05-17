import os
import shutil
import tkinter as tk
from tkinter import filedialog, messagebox
from datetime import datetime

def selecionar_arquivo():
    caminho = filedialog.askopenfilename(filetypes=[("Arquivos Excel", "*.xlsx *.xls")])
    if caminho:
        entrada_arquivo.delete(0, tk.END)
        entrada_arquivo.insert(0, caminho)

def salvar_backup():
    nome_personalizado = entrada_nome.get().strip()
    caminho_arquivo = entrada_arquivo.get().strip()

    if not nome_personalizado or not caminho_arquivo:
        messagebox.showwarning("Atenção", "Preencha todos os campos!")
        return

    nome_original = os.path.basename(caminho_arquivo)
    data_hora = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")

    # Nome no formato: NomePersonalizado_NomeDoArquivo_Data
    nome_pasta = f"{nome_personalizado}_{nome_original}_{data_hora}"
    pasta_destino = os.path.join(os.getcwd(), nome_pasta)

    try:
        os.makedirs(pasta_destino)
        shutil.copy2(caminho_arquivo, pasta_destino)
        messagebox.showinfo("Sucesso", f"Backup salvo na pasta:\n{nome_pasta}")
    except Exception as e:
        messagebox.showerror("Erro", f"Não foi possível salvar o backup:\n{e}")

# Interface gráfica
janela = tk.Tk()
janela.title("Backup de Arquivo Excel")
janela.geometry("500x200")

tk.Label(janela, text="Nome personalizado da pasta:").pack(pady=5)
entrada_nome = tk.Entry(janela, width=50)
entrada_nome.pack()

tk.Label(janela, text="Arquivo Excel:").pack(pady=5)
frame_arquivo = tk.Frame(janela)
frame_arquivo.pack()
entrada_arquivo = tk.Entry(frame_arquivo, width=40)
entrada_arquivo.pack(side=tk.LEFT)
btn_selecionar = tk.Button(frame_arquivo, text="Selecionar", command=selecionar_arquivo)
btn_selecionar.pack(side=tk.LEFT, padx=5)

btn_salvar = tk.Button(janela, text="Salvar Backup", command=salvar_backup, bg="green", fg="white")
btn_salvar.pack(pady=15)

janela.mainloop()
