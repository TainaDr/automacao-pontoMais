import pandas as pd
import os
import datetime
from openpyxl.utils import get_column_letter
import openpyxl
import glob

diretorio_planilhas = '../planilhas'
diretorio_relatorio = '../relatorio'

os.makedirs(diretorio_relatorio, exist_ok=True)

linhas_pular = 2
linhas_pular_ponto = 3
indice_cabecalho = 0

colunas_a_remover_mapa = {
    'colaboradores': [
        'PIS', 'Cargo', 'CPF', 'E-mail', 'Centro de custo', 'Data de admissão'
    ],
    'turnos': [
        'Limite de horas extras', 'Tipo', 'Virada do turno'
    ],
    'registros_de_ponto': [
        'Equipe', 'Turno'
    ]
}

data_atual = datetime.datetime.now()
timestamp_diario = data_atual.strftime("%Y-%m-%d")
nome_arquivo = f'Pontomais-unificado_{timestamp_diario}.xlsx'
arquivo_final = os.path.join(diretorio_relatorio, nome_arquivo)

def obter_arquivos_para_processar(diretorio, quantidade):
    print(f"Buscando os {quantidade} arquivos mais recentes em '{diretorio}'...")
    caminho_completo = os.path.join(diretorio, '*.xlsx')
    lista_de_arquivos = glob.glob(caminho_completo)

    if len(lista_de_arquivos) < quantidade:
        print(f"Atenção: Foram encontrados apenas {len(lista_de_arquivos)} arquivos. São necessários {quantidade}.")
        return []

    lista_de_arquivos.sort(key=os.path.getmtime)
    return lista_de_arquivos[:quantidade]

def gerar_nomes_relatorios(lista_arquivos):
    nomes = []
    for arquivo in lista_arquivos:
        nome_base = os.path.basename(arquivo)
        nome_limpo = nome_base.replace('.xlsx', '')
        nomes.append(f'Relatório {nome_limpo}')
    return nomes

def unificar_e_formatar_excel(lista_arquivos, chaves_cabecalho, header_row_index, arquivo_saida):
    lista_de_dados = []
    print("\nIniciando o processo de unificação...")

    for arquivo in lista_arquivos:
        try:
            nome_arquivo_base = os.path.basename(arquivo).lower()

            if 'registros_de_ponto' in nome_arquivo_base:
                print(f"   - Arquivo '{os.path.basename(arquivo)}' (Relatório de Ponto) -> Pulando {linhas_pular_ponto} linhas.")
                df = pd.read_excel(arquivo, skiprows=linhas_pular_ponto, header=header_row_index, engine='openpyxl')
            else:
                print(f"   - Arquivo '{os.path.basename(arquivo)}' -> Pulando {linhas_pular} linhas.")
                df = pd.read_excel(arquivo, skiprows=linhas_pular, header=header_row_index, engine='openpyxl')

            lista_de_dados.append(df)

        except Exception as e:
            print(f"Erro ao ler o arquivo '{arquivo}': {e}")
            return False

    if not lista_de_dados:
        print("\nNenhum dado foi lido. O programa será encerrado.")
        return False

    print("\nJuntando os dados com cabeçalhos de múltiplos níveis...")
    dados_unificados = pd.concat(lista_de_dados, axis=1, keys=chaves_cabecalho)

    print("\nRemovendo colunas desnecessárias com base em palavras-chave...")
    colunas_a_remover_final = []
    
    # Itera sobre as colunas que REALMENTE existem no arquivo unificado
    for col_tupla in dados_unificados.columns:
        report_name = str(col_tupla[0]).lower() 
        column_name = str(col_tupla[1])     

        # Itera sobre o nosso mapa de palavras-chave e colunas
        for keyword, column_list in colunas_a_remover_mapa.items():
            if keyword in report_name and column_name in column_list:
                colunas_a_remover_final.append(col_tupla)
                break # Sai do loop interno assim que encontra uma correspondência

    if colunas_a_remover_final:
        dados_unificados = dados_unificados.drop(columns=colunas_a_remover_final)
        print(f"   - {len(colunas_a_remover_final)} coluna(s) removida(s) com sucesso.")
    else:
        print("   - Nenhuma coluna para remover foi encontrada.")

    print(f"\nSalvando e formatando o arquivo final: '{arquivo_saida}'...")
    try:
        with pd.ExcelWriter(arquivo_saida, engine='openpyxl') as writer:
            dados_unificados.to_excel(writer, sheet_name='Dados Unificados')
            worksheet = writer.sheets['Dados Unificados']
            for i, column in enumerate(dados_unificados.columns):
                header_level_1 = len(str(column[0]))
                header_level_2 = len(str(column[1]))
                max_len_data = 0
                if not dados_unificados[column].empty:
                    max_len_data = dados_unificados[column].astype(str).map(len).max()

                column_width = max(header_level_1, header_level_2, max_len_data) + 2
                column_letter = get_column_letter(i + 2)
                worksheet.column_dimensions[column_letter].width = column_width

        print(f"\n-> Sucesso! Arquivo '{arquivo_saida}' finalizado e formatado corretamente.")
        return True

    except Exception as e:
        print(f"\n-> Erro ao salvar ou formatar o arquivo final: {e}")
        return False

def limpar_arquivos_originais(lista_arquivos):
    print("\nIniciando limpeza dos arquivos de origem...")
    for arquivo in lista_arquivos:
        try:
            os.remove(arquivo)
            print(f"   - Arquivo '{os.path.basename(arquivo)}' removido com sucesso.")
        except OSError as e:
            print(f"   - Erro ao remover o arquivo '{os.path.basename(arquivo)}': {e}")

if __name__ == "__main__":
    arquivos_para_processar = obter_arquivos_para_processar(diretorio_planilhas, 3)

    if arquivos_para_processar:
        print("\nArquivos encontrados (do mais antigo para o mais novo):")
        for f in arquivos_para_processar:
            print(f"   - {os.path.basename(f)}")

        nomes_relatorios = gerar_nomes_relatorios(arquivos_para_processar)

        sucesso = unificar_e_formatar_excel(
            lista_arquivos=arquivos_para_processar,
            chaves_cabecalho=nomes_relatorios,
            header_row_index=indice_cabecalho,
            arquivo_saida=arquivo_final
        )

        if sucesso:
            print("\nProcesso finalizado. Limpando arquivos originais...")
            limpar_arquivos_originais(arquivos_para_processar)
            print("Arquivos da pasta 'planilhas' excluídos com sucesso.")

    else:
        print("\nProcesso não iniciado devido à falta de arquivos.")