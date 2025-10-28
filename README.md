# Automação de Relatórios

### Visão Geral

Este projeto tem como objetivo automatizar a geração e tratamento de relatórios do sistema **PontoMais**, incluindo:

* Relatório de Colaboradores
* Relatório de Registro de Pontos
* Relatório de Turnos

A automação visa economizar tempo, evitar erros manuais e permitir análise rápida dos dados de ponto, folgas, afastamentos e horas extras.

### Estrutura do Projeto

* **extracao/** → Contém os scripts de extração dos relatórios.
* **transformacao/** → Contém os scripts de transformação dos relatórios.
* **planilhas/** → Armazena arquivos baixados do PontoMais.
* **relatorio/** → Diretório de saída com relatórios processados.

### Execução

1. Configurar parâmetros de filtros e credenciais em um arquivo .env.
    ```bash
    # Arquivo local
    LOGIN_EMAIL="SEU-LOGIN"
    LOGIN_SENHA="SUA-SENHA"
    CAMINHO_ARQUIVO=CAMINHO-DIRETORIO
   ```
2. Executar o script correspondente ao relatório desejado:

   ```bash
   python main.py
   ```
3. Os relatórios gerados são salvos na pasta `relatorio/`.
