# Arquivo: quebra_sped.py
# Autor: Rogério N
# Descrição: Este é um script Python simples para quebrar SPED por sheets em excel (ainda em andamento)

import pandas as pd
import re
import os
import tkinter as tk
from tkinter import filedialog#, messagebox
from pathlib import Path
# Importando customtkinter como ctk
import customtkinter as ctk
# Função para processar um arquivo SPED
def process_sped(file_paths, output_folder, status_label):
    # Lista de registros fornecida (remover duplicatas usando set)
    record_list = list(set([
        '0000', '0001', '0005', '0100', '0150', '0190', '0200', '0205', '0400',
        'C001', 'C100', 'C170', 'C190', 'C400', 'C405', 'C420', 'C460',
        'D001', 'D100', 'D500', 'E001', 'E100', 'E110', 'E200', 'E210',
        'G001', 'G110', 'G125', 'G140', 'H001', 'H005', 'H010', 'H020',
        'K001', 'K100', 'K200', 'K220', 'K230', 'K250', 'K290', 'K300',
        'K315', '1001', '1010', '1100', '1900', '1990', '9001', '9900',
        '9990', '9999', '0110', '0140', '0500', 'A001', 'A100', 'A170',
        'A180', 'C001', 'C100', 'C170', 'C180', 'C190', 'D001', 'D100',
        'D170', 'D180', 'D190', 'F001', 'F100', 'F170', 'F180', 'F190',
        'M001', 'M100', 'M105', 'M200', 'M500', 'M505', 'M600', '1001',
        '1010', '1990'
    ]))
    # Definição de padrões para identificar cada tipo de registro
    regex_blocks = {record: rf'\|{record}\|' for record in record_list}
    # Função para identificar o tipo de bloco
    def identify_block(line):
        for key, regex_pattern in regex_blocks.items():
            if re.search(regex_pattern, line):
                return key
        return None
    # Loop sobre todos os arquivos SPED selecionados
    for file_path in file_paths:
        # Lista para armazenar os dados de cada bloco
        data_blocks = {key: [] for key in regex_blocks}
        try:
            # Tentativa de leitura do arquivo SPED TXT e separação dos blocos
            with open(file_path, 'r', encoding='latin-1') as file:
                for line in file:
                    block_type = identify_block(line)
                    if block_type:
                        data_blocks[block_type].append(line.strip())
            # Criar um nome de arquivo Excel baseado no nome do arquivo .txt selecionado
            file_name = Path(file_path).stem  # Obtém o nome base sem extensão
            output_excel = os.path.join(output_folder, f'{file_name}.xlsx')
            # Criação de um DataFrame para cada tipo de bloco
            dfs = {key: pd.DataFrame([re.split(r'\|', row.strip('|')) for row in data_blocks[key]]) for key in data_blocks}
            # Ordenar as chaves do dicionário para que as sheets fiquem em ordem numérica e alfabética
            sorted_keys = sorted(dfs.keys())
            # Criação do arquivo Excel com cada bloco em uma sheet, se não estiver vazio
            with pd.ExcelWriter(output_excel, engine='openpyxl') as writer:
                for key in sorted_keys:
                    df = dfs[key]
                    if not df.empty:
                        df.to_excel(writer, sheet_name=key, index=False, header=False)
            status_label.configure(text=f'Arquivo Excel "{output_excel}" criado com sucesso!')
        except UnicodeDecodeError as e:
            status_label.configure(text=f"Erro de decodificação: {e}")
        except Exception as e:
            status_label.configure(text=f"Ocorreu um erro: {e}")
# Funções de interface gráfica
def select_sped_files():
    file_paths = filedialog.askopenfilenames(filetypes=[("Text files", "*.txt")])
    if file_paths:
        sped_entry.delete(0, tk.END)
        sped_entry.insert(0, ", ".join(file_paths))
def select_output_folder():
    folder_path = filedialog.askdirectory()
    if folder_path:
        output_entry.delete(0, tk.END)
        output_entry.insert(0, folder_path)
def start_processing(status_label):
    file_paths = sped_entry.get().split(", ")
    output_folder = output_entry.get()
    if file_paths and output_folder:
        process_sped(file_paths, output_folder, status_label)
    else:
        status_label.configure(text='Por favor, selecione os arquivos SPED e a pasta de destino.')
# Criação da interface gráfica utilizando customtkinter (ctk)
app = ctk.CTk()
app.title("Processador de Arquivo SPED")
app.geometry("770x520")  # Aumentando o tamanho da janela em cerca de 30%
frame = ctk.CTkScrollableFrame(app, width=800, height=520)
frame.pack(pady=30, padx=30)
fonte_label = ("Arial Black", 17)  # Definindo a fonte Arial Black com tamanho 14 para uso frequente
sped_label = ctk.CTkLabel(frame, text="Selecione os arquivos SPED:", font=fonte_label)
sped_label.pack(pady=15)
sped_entry = ctk.CTkEntry(frame, width=130, font=fonte_label)
sped_entry.pack(pady=10)
sped_button = ctk.CTkButton(frame, text="Selecionar Arquivos", command=select_sped_files, font=fonte_label)
sped_button.pack(pady=10)
output_label = ctk.CTkLabel(frame, text="Selecione a pasta de destino:", font=fonte_label)
output_label.pack(pady=15)
output_entry = ctk.CTkEntry(frame, width=130, font=fonte_label)
output_entry.pack(pady=10)
output_button = ctk.CTkButton(frame, text="Selecionar Pasta", command=select_output_folder, font=fonte_label)
output_button.pack(pady=10)
process_button = ctk.CTkButton(frame, text="GERAR ARQUIVO EM EXCEL", command=lambda: start_processing(status_label), font=fonte_label)
process_button.pack(pady=15)
status_label = ctk.CTkLabel(frame, text="", font=fonte_label)
status_label.pack(pady=15)
app.mainloop()


#GERA EXECUTAVEL -- pyinstaller --onefile -w quebra_sped.py
#pyinstaller --clean quebra_sped.py
