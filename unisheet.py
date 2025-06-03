import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import os

def selecionar_arquivos():
    arquivos = filedialog.askopenfilenames(
        title="Selecione os arquivos Excel",
        filetypes=[("Arquivos Excel", "*.xlsx *.xls")]
    )
    if arquivos:
        lista_arquivos.set("\n".join(arquivos))
        btn_unir.config(state=tk.NORMAL)
        global arquivos_selecionados
        arquivos_selecionados = arquivos

def unir_arquivos():
    try:
        dfs = []
        for arquivo in arquivos_selecionados:
            df = pd.read_excel(arquivo)
            df["Arquivo_Origem"] = os.path.basename(arquivo)  # adiciona uma coluna de origem opcional
            dfs.append(df)
        
        resultado = pd.concat(dfs, ignore_index=True)

        salvar_como = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Arquivo Excel", "*.xlsx")],
            title="Salvar arquivo unido como"
        )
        if salvar_como:
            resultado.to_excel(salvar_como, index=False)
            messagebox.showinfo("Sucesso", f"Arquivo salvo com sucesso:\n{salvar_como}")
        else:
            messagebox.showwarning("Cancelado", "A operação foi cancelada.")
    except Exception as e:
        messagebox.showerror("Erro", f"Ocorreu um erro ao unir os arquivos:\n{e}")

# Criação da interface
janela = tk.Tk()
janela.title("Unisheet")
janela.geometry("600x400")
janela.configure(bg="#f0f0f0")

# Título
titulo = tk.Label(janela, text="Unisheet", font=("Arial", 18, "bold"), bg="#f0f0f0", fg="#333")
titulo.pack(pady=20)

# Botão de seleção
btn_selecionar = tk.Button(janela, text="Selecionar arquivos", font=("Arial", 12), command=selecionar_arquivos, bg="#4caf50", fg="white", padx=10, pady=5)
btn_selecionar.pack()

# Área de exibição dos arquivos selecionados
lista_arquivos = tk.StringVar()
label_arquivos = tk.Label(janela, textvariable=lista_arquivos, font=("Arial", 10), justify="left", bg="#f0f0f0", fg="#555")
label_arquivos.pack(pady=10)

# Botão de unir
btn_unir = tk.Button(janela, text="Unir e Salvar", font=("Arial", 12, "bold"), command=unir_arquivos, bg="green", fg="white", padx=10, pady=5, state=tk.DISABLED)
btn_unir.pack(pady=20)

# Rodapé
footer = tk.Label(janela, text="Desenvolvido por Felipe (maidenzinho)", font=("Arial", 9), bg="#f0f0f0", fg="#999")
footer.pack(side="bottom", pady=10)

# Execução da interface
janela.mainloop()