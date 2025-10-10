import pandas as pd
import os
import glob
#Michael revisar aqui
pasta_planilhas = '../planilhas'
nome_arquivo_saida = 'turnos_extraidos.xlsx'

print(f"Procurando pelo relatório mais recente na pasta '{pasta_planilhas}'...")
caminho_busca = os.path.join(pasta_planilhas, 'Pontomais_-_Turnos_-_*.xlsx')
arquivo_final = os.path.join(pasta_planilhas, nome_arquivo_saida)
lista_arquivos = glob.glob(caminho_busca)

if not lista_arquivos:
    print(f" ERRO: Nenhum arquivo começando com 'Pontomais_-_Turnos_-_' foi encontrado na pasta '{pasta_planilhas}'.")
else:
    nome_arquivo_entrada = max(lista_arquivos, key=os.path.getmtime)
    print(f" Arquivo mais recente encontrado: '{nome_arquivo_entrada}'")

    try:
        df = pd.read_excel(nome_arquivo_entrada, header=None, engine='openpyxl', dtype={2: str})
        
        valores_encontrados = []
        for valor in df[1].dropna():
            if str(valor).strip().startswith('0'):
                valores_encontrados.append(valor)

        if not valores_encontrados:
            print("\nNenhum valor começando com '0' foi encontrado na Coluna C. Nenhum arquivo foi salvo.")
        else:
            print(f"\nEncontrados {len(valores_encontrados)} valores. Limpando e organizando os dados...")
            
            dados_processados = []
            for valor_completo in valores_encontrados:
                partes = valor_completo.split('-', 1)
                if len(partes) == 2:
                    codigo_bruto, descricao = partes
                    codigo_limpo = codigo_bruto.strip().lstrip('0')
                    dados_processados.append({
                        'codigo': codigo_limpo,
                        'descrição': descricao.strip()
                    })
            
            if dados_processados:
                df_final = pd.DataFrame(dados_processados)
                df_final.to_excel(arquivo_final, index=False)
                print(f"Os dados foram processados e salvos no arquivo '{arquivo_final}'.")
            else:
                print("\nNenhum dos valores encontrados pôde ser dividido corretamente (faltava o '-').")

    except Exception as e:
        print(f"\n Ocorreu um erro ao processar o arquivo: {e}")