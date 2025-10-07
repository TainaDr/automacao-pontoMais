import subprocess
import sys
import time

def script_sequencial(caminho):
    print(f"INICIANDO: {caminho}...")
    
    try:
        processo = subprocess.run(
            [sys.executable, caminho], 
            check=True, 
            capture_output=True,
            text=True,
            encoding='utf-8'
        )

        print(f"Saída de '{caminho}':\n{processo.stdout}")
        print(f"[SUCESSO] O script {caminho} foi concluído com êxito.\n")
        return True # Retorna True em caso de sucesso

    except FileNotFoundError:
        print(f"[ERRO FATAL] O arquivo '{caminho}' não foi encontrado.")
        return False # Retorna False em caso de erro

    except subprocess.CalledProcessError as e:
        print(f"[ERRO FATAL] O script {caminho} falhou.")
        print(f"Código de saída: {e.returncode}")
        print(f"Saída de erro:\n{e.stderr}")
        return False # Retorna False em caso de erro
        
    except Exception as e:
        print(f"[ERRO INESPERADO] Ocorreu um erro não identificado ao executar {caminho}: {e}")
        return False

def main():
    # Define aqui a lista e a ordem dos scripts 
    scripts_para_executar = [
        'extracao/extracao_colaborador.py',
        'extracao/extracao_ponto.py',
        'extracao/extracao_turno.py',
        'transformacao/transformacao.py'
    ]

    print("INICIANDO PROCESSO DE EXECUÇÃO SEQUENCIAL")
    
    inicio_total = time.time()

    for script in scripts_para_executar:
        if not script_sequencial(script):
            print("\nExecução interrompida devido a um erro no script anterior.")
            break  # Interrompe o loop se um script falhar
    
    fim_total = time.time()
    tempo_total = fim_total - inicio_total
    
    print(f"PROCESSO FINALIZADO em {tempo_total:.2f} segundos.")

if __name__ == "__main__":
    main()