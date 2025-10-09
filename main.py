import subprocess
import sys
import time
import os

def script_sequencial(caminho):
    print(f"INICIANDO: {caminho}...")
    
    try:
        caminho_absoluto = os.path.abspath(caminho)
        pasta_do_script = os.path.dirname(caminho_absoluto)

        processo = subprocess.run(
            [sys.executable, caminho_absoluto],
            check=True,
            capture_output=True,
            text=True,
            encoding='cp1252',
            cwd=pasta_do_script  
        )

        print(f"Saída de '{caminho}':\n{processo.stdout}")
        print(f"O script {caminho} foi concluído com êxito.\n")
        return True

    except FileNotFoundError:
        print(f"O arquivo '{caminho}' não foi encontrado.")
        return False

    except subprocess.CalledProcessError as e:
        print(f"O script {caminho} falhou.")
        print(f"Código de saída: {e.returncode}")
        print(f"Saída de erro:\n{e.stderr}")
        return False
        
    except Exception as e:
        print(f"Ocorreu um erro ao executar {caminho}: {e}")
        return False

def main():
    scripts_para_executar = [
        'extracao/extracao.py',
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