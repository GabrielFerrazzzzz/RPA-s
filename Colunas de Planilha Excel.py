import tkinter as tk
from tkinter import filedialog
import pandas as pd
import os

# CARREGUE O CSV AQUI OU O XLSX

def load_original_csv():
    # Agora permite selecionar ambos, arquivos CSV e Excel
    filepath = filedialog.askopenfilename(
        filetypes=[("CSV files", "*.csv"), ("Excel files", "*.xlsx")]
    )
    if filepath:
        original_file_path_var.set(filepath)

# ELE IRÁ SALVAR O ARQUIVO EDITADO COMO CSV

def save_new_csv():
    filepath = filedialog.asksaveasfilename(
        defaultextension=".csv",
        filetypes=[("CSV files", "*.csv")]
    )
    if filepath:
        new_file_path_var.set(filepath)

# IRÁ EXPORTAR AS COLUNAS QUE PRECISAR SER PEGAS

def export_columns():
    original_file_path = original_file_path_var.get()
    new_file_path = new_file_path_var.get()
    
    if original_file_path and new_file_path:
        try:
            if original_file_path.endswith('.csv'):
                data = pd.read_csv(original_file_path, delimiter=';', encoding='utf-8-sig')
            elif original_file_path.endswith('.xlsx'):
                data = pd.read_excel(original_file_path)
            
            data['NOME DA NOVA COLUNA AQUI'] = '' # ESCREVA UM NOVO DE NOVAS COLUNAS CASO TENHAM NO ARQUIVO, IRÁ APARECER NA PRÓXIMA COLUNA EM BRANCO
            data['Status'] = '' #EXEMPLOS
            data['Ocorrência'] = '' #EXEMPLOS


            # AQUI VOCÊ COLOCA O NOME DAS COLUNAS QUE PRECISA DAQUELA PLANILHA (VOCÊ PODE TER QUANTAS COLUNAS QUISER NO ARQUIVO NOVO)
            selected_columns = data[["NOME DA COLUNAS DA PLANILHA ORIGINAL AQUI", "NOME DA COLUNAS DA PLANILHA ORIGINAL AQUI", "NOME"]] #"NOME" É UM EXEMPLO APENAS

            # Verificar se o arquivo já existe para decidir se o cabeçalho deve ser adicionado
            header = not os.path.isfile(new_file_path)
            # Adicionar dados ao arquivo, anexando se já existe (modo 'a')
            selected_columns.to_csv(new_file_path, mode='a', index=False, sep=';', encoding='utf-8-sig', header=header)
            result_label.config(text=f"Arquivo atualizado com sucesso em: {new_file_path}")
        except Exception as e:
            result_label.config(text=f"Ocorreu um erro: {e}")
    else:
        result_label.config(text="Por favor, selecione os caminhos dos arquivos.")


# (O restante do código permanece o mesmo)
# Criando a janela principal
root = tk.Tk()
root.title("Exportar Colunas CSV")

# Definindo as variáveis de caminho dos arquivos
original_file_path_var = tk.StringVar()
new_file_path_var = tk.StringVar()

# Adicionando os widgets
tk.Label(root, text="Caminho do arquivo CSV original:").pack(fill='x', padx=5, pady=5)
original_path_entry = tk.Entry(root, textvariable=original_file_path_var, state='readonly')
original_path_entry.pack(fill='x', padx=5, pady=2)
tk.Button(root, text="Selecionar", command=load_original_csv).pack(fill='x', padx=5, pady=2)

tk.Label(root, text="Salvar novo arquivo CSV em:").pack(fill='x', padx=5, pady=5)
new_path_entry = tk.Entry(root, textvariable=new_file_path_var, state='readonly')
new_path_entry.pack(fill='x', padx=5, pady=2)
tk.Button(root, text="Selecionar", command=save_new_csv).pack(fill='x', padx=5, pady=2)

tk.Button(root, text="Exportar Colunas", command=export_columns).pack(fill='x', padx=5, pady=10)

# Label para mostrar resultados ou erros
result_label = tk.Label(root, text="")
result_label.pack(fill='x', padx=5, pady=5)

# Iniciando o loop da GUI
root.mainloop()