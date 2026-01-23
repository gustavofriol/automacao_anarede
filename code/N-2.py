from os.path import exists
import pandas as pd
from __pwfWriter__ import Writer
import __diretorios__ as diretorios
import __relatoriosAnarede__ as Rela
import subprocess
import os


def executar_anarede(path_pwf):
    print(f"Executando Anarede com o arquivo: {path_pwf}")
    subprocess.run(
        [diretorios.anarede_exe_path, path_pwf, "/b"],  # [programa, argumento1]
        cwd=diretorios.dir_pwf,
        capture_output=True,
        text=True,
        check=True,
        encoding='latin-1'  # Adicionado para compatibilidade com a saída do Anarede
    )
    print("Execução do Anarede para o caso base finalizada.")


# Configuração da planilha base
dir_ctgs = diretorios.dir_banco_de_dados  # Diretório onde se encontra a planilha base
file_path_ctgs = diretorios.excel_ativos  # Junção diretório + nome da planilha base

#  --------------------------------------------------- ATENÇÃO: ---------------------------------------------------
"""
A variável "name_dir_n_1" abaixo deve ser o mesmo nome que consta em "sheet_xlsx_ctgs" na execução do arquivo N-1. 
Aqui, esse nome é o nome da pasta onde os relatórios de variação n-1 estão localizados.
"""
name_dir_n_1 = 'PMD-NOV'

dir_reports = diretorios.dir_relatorios
dir_reports_n_1 = fr'{dir_reports}\{name_dir_n_1}'

########################################################################################################################
print("Etapa 1: Extraindo informações sobre as contingências N-1 executadas")


def escolher_criterio():
    """
    Função para escolher o critério de seleção para eleger N-2.
    Opções:
    - Var. MVA/V
    - Var. Carregamento %
    """
    prefix_reports = 'Relatório de Variações - '
    n_1_list = []
    dfs_n_1 = {}
    criterio_n_2 = 0  # Inicializa a variável

    # Loop para garantir que o usuário digite uma opção válida (1 ou 2)
    while criterio_n_2 not in [1, 2]:
        print("""
Critério para eleger N-2.
Pegar 20 elementos com maior variação de:

[1] MVA/V
[2] Carregamento % 
        """)
        try:
            criterio_n_2 = int(input('\nSelecione o critério para eleger as contingências N-2: '))
            if criterio_n_2 not in [1, 2]:
                print("\n--- ERRO: Opção inválida. Por favor, selecione 1 ou 2. ---\n")
        except ValueError:
            print("\n--- ERRO: Entrada inválida. Por favor, digite um número. ---\n")

    # Agora que temos um critério válido, processamos os arquivos
    if os.path.exists(dir_reports_n_1):
        for file in os.listdir(dir_reports_n_1):
            path = os.path.join(dir_reports_n_1, file)
            if os.path.isfile(path) and file.endswith('.xlsx'):
                nome_sem_prefixo = file.replace(prefix_reports, '')
                nome_da_linha = nome_sem_prefixo.rsplit('.', 1)[0]
                n_1_list.append(nome_da_linha)

                df_ctg = pd.read_excel(path)

                if criterio_n_2 == 1:
                    df_ctg.sort_values(by='Var. MVA/V', ascending=False, inplace=True) # Precisa reordener index após o sort??

                else:  # Se não for 1, com certeza é 2
                    df_ctg.sort_values(by='Var. Carregamento %', ascending=False, inplace=True)

                df_ctg.reset_index(drop=True, inplace=True)
                dfs_n_1[file] = df_ctg.head(20)
    else:
        print(f"Erro: O diretório '{dir_reports_n_1}' não foi encontrado.")
        print("Verifique se o script N-1 já foi executado.")

    return n_1_list, dfs_n_1

n_1_list, dfs_n_1 = escolher_criterio()

# n_1_list: Nome das pastas
# dfs_n_1: Dataframes contendo as 20 contingências a se fazer N-2
# Falta: Lista DE-PARA-NC de N-1 para gerar DLIN no PWF

########################################################################################################################
