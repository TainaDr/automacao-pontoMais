import pandas as pd
import os
import re
import glob

base_dir = os.path.dirname(os.path.abspath(__file__))
padrao_arquivo = os.path.join(base_dir, '..', 'planilhas',  'Pontomais_-_Turnos_-*.xlsx')

# Lista todos os arquivos que correspondem ao padrão
arquivos_encontrados = glob.glob(padrao_arquivo)

if not arquivos_encontrados:
    print(f"ERRO: Nenhum arquivo encontrado com o padrão 'Pontomais_-_Turnos_-*' em 'planilhas'.")
else:
    # Seleciona o primeiro arquivo encontrado (ou você pode usar max para o mais recente)
    caminho_arquivo_entrada = max(arquivos_encontrados, key=os.path.getctime)
    print(f"Lendo o arquivo '{caminho_arquivo_entrada}'...")

    try:
        # Lê a planilha Excel
        df = pd.read_excel(caminho_arquivo_entrada, engine='openpyxl')

        # Coluna C (índice 1)
        coluna_turnos = df.iloc[:, 1]

        # Filtra apenas os valores que começam com '0'
        valores_filtrados = coluna_turnos.dropna()
        valores_filtrados = valores_filtrados[valores_filtrados.astype(str).str.startswith('0')].reset_index(drop=True)

        # Cria DataFrame apenas com os turnos filtrados
        df_filtrado = pd.DataFrame({'raw': valores_filtrados})

        # Separa o código inicial da descrição usando '-'
        codigos = []
        descricoes = []
        for valor in df_filtrado['raw']:
            partes = valor.split('-', 1)
            if len(partes) == 2:
                codigo = partes[0].strip()
                descricao = partes[1].strip()
            else:
                codigo = partes[0].strip()
                descricao = ''
            codigos.append(codigo)
            descricoes.append(descricao)

        df_filtrado['codigo'] = codigos
        df_filtrado['descrição'] = descricoes

        # Extrai o primeiro horário de cada descrição
        padrao_horario = re.compile(r'(\d{2}:\d{2})')
        entradas = []
        for descricao in df_filtrado['descrição']:
            if isinstance(descricao, str):
                match = padrao_horario.search(descricao)
                if match:
                    entradas.append(match.group(1))
                else:
                    entradas.append(None)
            else:
                entradas.append(None)

        df_filtrado['primeira_entrada_prevista'] = entradas

        # Mantém apenas as colunas desejadas na ordem: código, descrição, primeira_entrada_prevista
        df_final = df_filtrado[['codigo', 'descrição', 'primeira_entrada_prevista']]

        # Salva o arquivo final
        nome_arquivo_saida = 'turnos_extraidos.xlsx'
        caminho_arquivo_saida = os.path.join(base_dir, '..', 'planilhas', nome_arquivo_saida)
        df_final.to_excel(caminho_arquivo_saida, index=False)
        print(f"Arquivo '{nome_arquivo_saida}' gerado com sucesso com colunas separadas e horário extraído.")

    except Exception as e:
        print(f"Erro ao processar o arquivo: {e}")
