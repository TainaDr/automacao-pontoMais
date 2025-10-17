import pandas as pd
import os
import glob
import re

pasta_planilhas = '../planilhas'
nome_arquivo_saida = 'turnos_extraidos.xlsx'

caminho_arquivo = os.path.join(pasta_planilhas, nome_arquivo_saida)

if not os.path.exists(caminho_arquivo):
    print(f"ERRO: O arquivo '{nome_arquivo_saida}' não foi encontrado em '{pasta_planilhas}'.")
else:
    print(f"Lendo o arquivo '{caminho_arquivo}'...")
    try:
        # Lê a planilha Excel com os dados dos turnos
        df = pd.read_excel(caminho_arquivo, engine='openpyxl')

        # Define um padrão que captura horários no formato HH:MM
        padrao_horario = re.compile(r'(\d{2}:\d{2})')

        entradas = []  # armazena a primeira entrada de cada descrição
        for descricao in df['descrição']:
            if isinstance(descricao, str):
                match = padrao_horario.search(descricao)  # procura o primeiro horário na descrição
                if match:
                    entradas.append(match.group(1))  # adiciona o horário encontrado
                else:
                    entradas.append(None)  # Se não encontrar, adiciona None
            else:
                entradas.append(None)  # Caso a descrição não seja string

        # Adiciona a coluna 'primeira_entrada_prevista' com os horários extraídos
        df['primeira_entrada_prevista'] = entradas

        # Salva a planilha atualizada
        df.to_excel(caminho_arquivo, index=False)
        print(f"Coluna 'primeira_entrada_prevista' adicionada e arquivo atualizado com sucesso em '{caminho_arquivo}'.")

    except Exception as e:
        print(f"Erro ao processar o arquivo: {e}")
