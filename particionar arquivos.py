import tkinter as tk
from tkinter import filedialog
import pandas as pd
import os

def dividir_arquivo():
    # Abre a janela de seleção de arquivo
    arquivo_origem = filedialog.askopenfilename(title="Selecione o arquivo Excel de origem")

    if arquivo_origem:
        # Carrega o arquivo Excel
        dados_excel = pd.read_excel(arquivo_origem)

        # Divide os dados em pedaços de 999 linhas cada
        tamanho_pedaço = 999
        total_linhas = len(dados_excel)
        num_pedaços = total_linhas // tamanho_pedaço + (1 if total_linhas % tamanho_pedaço != 0 else 0)

        # Abre a janela de seleção de pasta de destino
        pasta_destino = filedialog.askdirectory(title="Selecione a pasta de destino")

        if pasta_destino:
            # Cria e salva os novos arquivos
            cabecalho = dados_excel.columns.tolist()

            for i in range(num_pedaços):
                inicio = i * tamanho_pedaço
                fim = min((i + 1) * tamanho_pedaço, total_linhas)
                novo_arquivo = os.path.join(pasta_destino, f'Arquivo_particionado_{i+1}.xlsx')
                dados_pedaço = dados_excel.iloc[inicio:fim]
                with pd.ExcelWriter(novo_arquivo, engine='xlsxwriter') as writer:
                    dados_pedaço.to_excel(writer, index=False, header=cabecalho)
            
            # Mensagem de conclusão
            tk.messagebox.showinfo("Concluído", f"Arquivo dividido em {num_pedaços} partes e salvos na pasta de destino.")
            root.quit()  # Encerra o programa após a conclusão

# Cria a janela principal
root = tk.Tk()
root.withdraw()  # Esconde a janela principal

# Chama a função para dividir o arquivo
dividir_arquivo()

# Mantém a janela principal aberta
root.mainloop()
