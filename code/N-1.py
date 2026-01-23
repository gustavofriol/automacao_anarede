from os.path import exists
import pandas as pd
from __pwfWriter__ import Writer
import __diretorios__ as diretorios
import __relatoriosAnarede__ as Rela
import subprocess
import time
import shutil
import os
import keyboard
import pathlib


def executar_anarede(path_pwf, diretorio_sinal_esperado, timeout_segundos=43200):
    """
    Executa o Anarede, monitora a criação de um DIRETÓRIO DE SINAL e força
    o encerramento do processo.

    Args:
        path_pwf (str): Caminho para o arquivo .pwf a ser executado.
        diretorio_sinal_esperado (str): Caminho completo do DIRETÓRIO de SINAL
                                        que o Anarede deve criar no final.
        timeout_segundos (int): Tempo máximo de espera em segundos (Default: 24h).
    """
    print(f"Executando Anarede com o arquivo: {path_pwf}")
    print(f"Aguardando a criação do DIRETÓRIO de sinal: {diretorio_sinal_esperado}")

    # 1. Garante que o DIRETÓRIO de "sinal" não exista
    try:
        if os.path.isdir(diretorio_sinal_esperado):
            # shutil.rmtree remove um diretório e seu conteúdo
            shutil.rmtree(diretorio_sinal_esperado)
            print("Diretório de sinal antigo removido.")
    except OSError as e:
        print(f"Aviso: Não foi possível remover o diretório de sinal antigo. {e}")

    # 2. Inicia o processo Anarede de forma não bloqueante (Popen)
    processo = subprocess.Popen(
        [diretorios.anarede_exe_path, path_pwf, "/b"],
        cwd=diretorios.dir_pwf,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        text=True,
        encoding='latin-1'
    )

    tempo_inicio = time.time()
    sinal_detectado = False

    # 3. Loop de monitoramento
    while True:
        # Verifica se o tempo de espera excedeu
        if time.time() - tempo_inicio > timeout_segundos:
            print(f"ERRO: Timeout de {timeout_segundos}s atingido. O Anarede não gerou o diretório sinal.")
            processo.kill()
            processo.wait()
            raise TimeoutError("Execução do Anarede falhou por timeout.")

        # Verifica se o DIRETÓRIO de sinal foi criado
        if os.path.isdir(diretorio_sinal_esperado):  # <--- VERIFICAÇÃO PRINCIPAL
            print("Diretório de sinal detectado. Anarede concluiu a tarefa.")
            sinal_detectado = True
            break

        # Verifica se o processo morreu inesperadamente
        if processo.poll() is not None:
            stdout, stderr = processo.communicate()
            print("ERRO: Processo do Anarede terminou inesperadamente.")
            print(f"STDOUT: {stdout}\nSTDERR: {stderr}")
            raise ChildProcessError("Execução do Anarede falhou.")

        # Pausa curta para não sobrecarregar a CPU
        time.sleep(2)

    # 4. Força o encerramento do Anarede
    if sinal_detectado and processo.poll() is None:  # Se o processo ainda estiver rodando
        print("Forçando o encerramento do Anarede para continuar o script...")
        try:
            processo.kill()
            processo.wait()
        except Exception as e:
            print(f"Aviso: Não foi possível matar o processo do Anarede. {e}")

    print("Execução do Anarede finalizada com sucesso.")
    time.sleep(1)  # Pequena pausa para garantir que o OS libere os arquivos


# Instanciando objetos e definções prévias
writer = Writer(pwf_dir=diretorios.dir_pwf, casos_dir=diretorios.dir_casos_referencia)
escolhas = writer.select_pwf()
caso_sav = escolhas[1]

# Configuração da planilha base
caminho_planilha_base = diretorios.planilha_base_n1  # Junção diretório + nome da planilha de contingências
excel_base = pd.ExcelFile(caminho_planilha_base)
lista_abas = excel_base.sheet_names
nro_abas = 0

print("\nListas de contingências na planilha basse:")
for i, aba in enumerate(lista_abas):
   nro_abas = i+1
   print(fr"[{i+1}] {aba}")

indice_aba = int(input(f"\nSelecione pelo índice ({0}-{nro_abas}) a lista de contingências à executar: ")) - 1

aba_planilha_base =  lista_abas[indice_aba] # Aba da planilha de contingências selecionada

arquivo_regime = os.path.join(diretorios.dir_pwf, diretorios.file_txt_regime)
df_ativos_esul = pd.read_excel(caminho_planilha_base, sheet_name=aba_planilha_base)


dir_output = os.path.join(diretorios.dir_relatorios, 'N-1', caso_sav, aba_planilha_base)

print(f"\nDiretório de saída definido como: {dir_output}")
if not exists(dir_output):
    print(f"Diretório não encontrado. Criando diretório: {dir_output}")
    os.makedirs(dir_output)
else:
    print("Diretório de saída já existe.")

########################################################################################################################


print("--- Início da Execução do Script ---")
print("Etapa 1: Escrevendo o deck para o caso base (Regime).")

print(f"Escrevendo arquivo de caso base: {diretorios.file_pwf_regime}")
writer.write_caso_zero(diretorios.file_pwf_regime, diretorios.file_txt_regime, escolhas)

########################################################################################################################

pwf_regime = fr'{diretorios.dir_pwf}/{diretorios.file_pwf_regime}'
sinal_regime = os.path.join(diretorios.dir_pwf, 'regime.signal')
executar_anarede(pwf_regime, sinal_regime)  # Passa o sinal para a função

########################################################################################################################


print("\nEtapa 3: Processando os resultados do caso base e comparando com a lista de ativos.")

print(f"Lendo o relatório de saída do Anarede: '{arquivo_regime}'...")
relatorio_regime = Rela.Dadl(arquivo_entrada=arquivo_regime)
df_regime = relatorio_regime.read_dadl()

# Condicionando dataframes
df_regime['de-para-nc'] = df_regime.apply(
    lambda x: relatorio_regime.criar_de_para_nc(row=x, col_de='de', col_para='para', col_nc='nc'), axis=1)
df_ativos_esul['de-para-nc'] = df_ativos_esul.apply(
    lambda x: relatorio_regime.criar_de_para_nc(row=x, col_de='Núm. Barra DE', col_para='Núm. Barra PARA', col_nc='NC'),
    axis=1)

# Verificando quais ctgs estão na planilha base mas não estão no caso
print("Verificando quais contingências da planilha não foram encontradas no caso base...")
lista_diferenca = list(set(df_ativos_esul['de-para-nc']) - set(df_regime['de-para-nc']))

# Filtrar df_ativos_esul, removendo os itens da lista_diferenca
print("\nFiltrando a lista da planilha base para remover os itens não encontrados...")
df_ativos_esul_filtrado = df_ativos_esul[~df_ativos_esul['de-para-nc'].isin(lista_diferenca)]
df_ativos_esul_filtrado.reset_index(drop=True,
                                    inplace=True)  # Reset no índice do df para deixar ele sequencial, após a remoção dos itens faltantes no caso
print(f"Filtragem concluída. {len(df_ativos_esul_filtrado)} ativos serão considerados.")

########################################################################################################################


print("\nEtapa 4: Escrevendo e executando o deck de contingências.")
writer = Writer(pwf_dir=diretorios.dir_pwf, casos_dir=diretorios.dir_casos_referencia)

print(f"Escrevendo o deck de contingências: {diretorios.file_pwf_ctgs}")
writer.write_contingencias(diretorios.file_pwf_ctgs, diretorios.file_txt_ctgs, escolhas, df_ativos_esul_filtrado)


pwf_ctgs = fr'{diretorios.dir_pwf}/{diretorios.file_pwf_ctgs}'
sinal_ctgs = os.path.join(diretorios.dir_pwf, 'ctgs.signal')
executar_anarede(pwf_ctgs, sinal_ctgs)  # Passa o sinal para a função


########################################################################################################################


print("\nEtapa 5: Processando os resultados das contingências.")

print("\nExtraindo dataframes de fluxo (MOSF) do relatório de contingências...")
mosf_ctgs = Rela.Mosf(arquivo_entrada=diretorios.path_ctgs_txt)
dfs_ctgs_originais = mosf_ctgs.extrair_dfs_contingencias(df_ativos_esul_filtrado)

print("Extraindo dataframe de fluxo (MOSF) do relatório do caso em regime...")
mosf_zero = Rela.Mosf(arquivo_entrada=diretorios.path_regime_txt)
df_zero_original = mosf_zero.extrair_df_caso_zero()

ctgs_keys = mosf_ctgs.extrair_lista_ctgs(
    df_ctgs=df_ativos_esul_filtrado)  # Nessa lista serão adicionados os nomes 'DE-PARA-NC' das contingências
nro_ctgs = len(dfs_ctgs_originais)
print(f"\nTotal de {nro_ctgs} contingências serão processadas.")

dfs_ctgs_comp = [[]] * nro_ctgs

ctgs_sem_relatorio = []
for i in range(nro_ctgs):
    if len(dfs_ctgs_originais[i]) <2:
        ctgs_sem_relatorio.append(df_ativos_esul_filtrado.iloc[i]['Loc.instalação'])
        dfs_ctgs_comp[i] = dfs_ctgs_originais[i][['DE-PARA-NC', 'MVA/V', 'Carregamento %']].copy()
        continue
    dfs_ctgs_originais[i]['DE-PARA-NC'] = dfs_ctgs_originais[i].apply(lambda x: mosf_ctgs.criar_de_para_nc(row=x, col_de='Núm. Barra DE', col_para='Núm. Barra PARA', col_nc='NC'),
    axis=1)
    dfs_ctgs_comp[i] = dfs_ctgs_originais[i][['DE-PARA-NC', 'MVA/V', 'Carregamento %']].copy()

print('CONTINGÊNCIAS SEM RELATÓRIO EXPORTADO PELO ANAREDE:\n')
print(ctgs_sem_relatorio)
# Exportar os dados para um arquivo .txt
with open(fr"{dir_output}\Desligamentos sem relatório.txt", "w", encoding="utf-8") as arquivo:
    arquivo.write(
        'As contingências abaixo estão na planilha base mas não foram encontradas no relatório exportado pelo Anarede. Rodar manualmente, se necessário.\n')
    arquivo.write('\nRelatórios excel vazios:\n\n')
    for item in ctgs_sem_relatorio:
        arquivo.write(item + "\n")


df_zero_original['DE-PARA-NC'] = df_zero_original.apply(lambda x: mosf_zero.criar_de_para_nc(row=x, col_de='Núm. Barra DE', col_para='Núm. Barra PARA', col_nc='NC'),
    axis=1)
df_zero = df_zero_original[['DE-PARA-NC', 'MVA/V', 'Carregamento %']].copy()

dfs_zero_comp = []
desligamentos_totais = []

for i, df_ctg in enumerate(dfs_ctgs_comp):
    print(f"Processando contingência {i + 1}/{nro_ctgs}...")
    # Identifica a row ausente no df_ctg comparado com df_zero
    merged_df = pd.merge(df_zero, df_ctg, on='DE-PARA-NC', how='left', indicator=True)
    row_to_remove = merged_df[merged_df['_merge'] == 'left_only'].index

    # Extrai os valores da coluna 'DE-PARA-NC' onde houve 'left_only' e converte para lista
    valores_ausentes = merged_df['DE-PARA-NC'][merged_df['_merge'] == 'left_only'].values.tolist()
    desligamentos_totais.extend(valores_ausentes)  # Usamos extend para adicionar os valores diretamente à lista

    # Cria uma cópia do df_zero e remove a row identificada
    df_zero_comp_atual = df_zero.copy()
    df_zero_comp_atual = df_zero_comp_atual.drop(row_to_remove)

    # Adiciona o dataframe resultante à lista dfs_zero_comp
    dfs_zero_comp.append(df_zero_comp_atual)

print("\nCalculando as variações de fluxo e tensão para cada contingência...")
dfs_var_fluxo = [pd.DataFrame(columns=['DE-PARA-NC', 'var', 'Carregamento %'])] * nro_ctgs
for i in range(nro_ctgs):
    dfs_var_fluxo[i] = pd.merge(dfs_zero_comp[i], dfs_ctgs_comp[i], on='DE-PARA-NC', suffixes=('_Pré', '_Pós'),
                                how='inner')
for i in range(len(dfs_var_fluxo)):
    dfs_var_fluxo[i]['Var. MVA/V'] = dfs_var_fluxo[i]['MVA/V_Pós'] - dfs_var_fluxo[i]['MVA/V_Pré']
    dfs_var_fluxo[i]['Var. Carregamento %'] = dfs_var_fluxo[i]['Carregamento %_Pós'] - dfs_var_fluxo[i]['Carregamento %_Pré']

mvt = Rela.Mvt(arquivo_entrada=diretorios.path_ctgs_txt)
dfs_var_tensao = mvt.extrair_dfs_contingencias(df_ativos_esul_filtrado)

print("\nEtapa 6: Organizando e exportando os resultados finais.")
dict_var_tensao, dict_var_fluxo = {}, {}

# Cria um dataframe de referência com os nomes das barras para consulta posterior
df_branch_names = df_zero_original[['DE-PARA-NC', 'Nome Barra DE', 'Nome Barra PARA', 'NC']].copy()

for i in range(nro_ctgs):
    ctg_key = ctgs_keys[i]

    # --- Processa e armazena o DataFrame de Variação de Fluxo ---
    df_fluxo = dfs_var_fluxo[i]
    # Mescla com a referência de nomes para obter as colunas de nome
    df_fluxo_named = pd.merge(df_fluxo, df_branch_names, on='DE-PARA-NC', how='left')

    # Cria a nova coluna 'LT-Circuito' concatenando os nomes e o número do circuito
    df_fluxo_named['LT-Circuito'] = df_fluxo_named['Nome Barra DE'] + '/' + df_fluxo_named['Nome Barra PARA'] + '/' + \
                                    df_fluxo_named['NC'].astype(str)

    # Reordena as colunas para posicionar 'LT-Circuito' à esquerda de 'DE-PARA-NC'
    cols = df_fluxo_named.columns.tolist()
    # Define as colunas que não são mais necessárias individualmente
    cols_to_remove = ['Nome Barra DE', 'Nome Barra PARA', 'NC']
    new_col = 'LT-Circuito'

    # Monta a lista de colunas na nova ordem
    remaining_cols = [c for c in cols if c not in cols_to_remove and c != new_col]
    de_para_nc_index = remaining_cols.index('DE-PARA-NC')
    new_order = remaining_cols[:de_para_nc_index] + [new_col] + remaining_cols[de_para_nc_index:]
    df_fluxo_final = df_fluxo_named[new_order].sort_values(by='Var. Carregamento %', ascending = False)

    dict_var_fluxo[ctg_key] = df_fluxo_final

    # --- Armazena o DataFrame de Variação de Tensão ---
    dict_var_tensao[ctg_key] = dfs_var_tensao[i]

# Cria um mapeamento da chave da contingência ('de-para-nc') para o nome do arquivo ('Loc.instalação')
name_map = pd.Series(df_ativos_esul_filtrado['Loc.instalação'].values,
                     index=df_ativos_esul_filtrado['de-para-nc']).to_dict()

df_tensao_vazio = pd.DataFrame(
    columns=['Núm_Barra', 'Nome_Barra', 'Tensão_Pré', 'Tensão_Pós', 'Variação_PU', 'Variação_%'])

def write_formated_sheet(writer, dataframe, nome_aba):
    """
    Escreve um DataFrame em uma aba de um arquivo Excel e aplica formatação.
    Args:
        writer (pd.ExcelWriter): O objeto ExcelWriter para o arquivo.
        dataframe (pd.DataFrame): O DataFrame a ser escrito.
        nome_aba (str): O nome da aba na qual o DataFrame será escrito.
    """
    dataframe.to_excel(writer, sheet_name=nome_aba, index=False)

    workbook = writer.book
    worksheet = writer.sheets[nome_aba]

    if not dataframe.empty:
        # Adiciona a Tabela do Excel para habilitar filtros e estilo
        (num_linhas, num_colunas) = dataframe.shape
        # O intervalo precisa cobrir todos os dados, incluindo o cabeçalho
        intervalo_tabela = f'A1:{chr(ord("A") + num_colunas - 1)}{num_linhas + 1}'
        worksheet.add_table(intervalo_tabela, {'columns': [{'header': col} for col in dataframe.columns]})

        # Ajusta a largura de cada coluna para se adequar ao conteúdo
        for idx, col in enumerate(dataframe.columns):
            # Calcula a largura máxima entre o cabeçalho e o maior valor na coluna
            largura_maxima = max(dataframe[col].astype(str).map(len).max(), len(col)) + 2
            letra_coluna = chr(ord("A") + idx)
            worksheet.set_column(f'{letra_coluna}:{letra_coluna}', largura_maxima)


print("\nIniciando a exportação dos relatórios para arquivos Excel...")
for i, ctg_key in enumerate(ctgs_keys):
    # Obtém o nome base para o arquivo da coluna 'Loc.instalação'
    file_name_base = name_map.get(ctg_key, ctg_key)

    # Remove caracteres inválidos do nome do arquivo para evitar erros
    sanitized_name = "".join(c for c in str(file_name_base) if c.isalnum() or c in (' ', '-', '_')).strip()

    output_filename = os.path.join(dir_output, f'Relatorio_{sanitized_name}.xlsx')
    print(f"Escrevendo arquivo: 'Relatorio_{sanitized_name}.xlsx'")

    # Utiliza o ExcelWriter para criar o arquivo
    with pd.ExcelWriter(output_filename, engine='xlsxwriter') as writer:

        # --- Aba 1: Variações de Fluxo ---
        df_fluxo = dict_var_fluxo[ctg_key]
        write_formated_sheet(writer, df_fluxo, 'Variações de Fluxo')

        # --- Aba 2: Variações de Tensão ---
        df_tensao_atual = dfs_var_tensao[i]

        # Verifica se o DataFrame de tensão é válido. Se não for, usa o DataFrame vazio.
        if isinstance(df_tensao_atual, pd.DataFrame):
            write_formated_sheet(writer, df_tensao_atual, 'Variações de Tensão')
        else:
            write_formated_sheet(writer, df_tensao_vazio, 'Variações de Tensão')



########################################################################################################################

print("\nEtapa 7: Criando o Relatório Geral.")

# 1. Definir o prefixo e sufixo para extrair nomes
prefixo_relatorio = "Relatorio_"
sufixo_relatorio = ".xlsx"

# 2. Obter o valor do 'Caso' da Etapa 1
try:
    # A variável 'escolhas' foi definida na Etapa 1
    pwf_selecionado = escolhas[2]
except (NameError, IndexError):
    print("AVISO: Não foi possível determinar o 'Caso' a partir da variável 'escolhas'. Usando 'N/D'.")
    pwf_selecionado = "N/D"  # Fallback, caso 'escolhas' não esteja definida como esperado

# 3. Preparar os dados para o DataFrame
dados_relatorio_geral = []

print(f"Lendo relatórios individuais da pasta: {dir_output}")

# Lista de arquivos no diretório de saída
try:
    arquivos_no_diretorio = os.listdir(dir_output)
except FileNotFoundError:
    print(f"ERRO: O diretório de saída '{dir_output}' não foi encontrado. Relatório Geral não será gerado.")
    arquivos_no_diretorio = []

for filename in arquivos_no_diretorio:
    # Verificar se o arquivo corresponde ao padrão
    if filename.startswith(prefixo_relatorio) and filename.endswith(sufixo_relatorio):
        # Extrair 'Local Instalação' do nome do arquivo
        local_instalacao = filename[len(prefixo_relatorio):-len(sufixo_relatorio)]

        # Montar o texto do link (como visto na imagem de exemplo)
        link_text = f"Relatório de Variações - {local_instalacao}"

        # Obter o caminho absoluto para o arquivo (para o hyperlink)
        full_path = os.path.abspath(os.path.join(dir_output, filename))

        # Converter o caminho em um URI (formato file://...)
        # Isso garante que o hyperlink funcione corretamente
        path_uri = pathlib.Path(full_path).as_uri()

        # Adicionar à lista de dados
        dados_relatorio_geral.append({
            "Local Instalação": local_instalacao,
            "Caso": pwf_selecionado,
            "Relatório_Texto": link_text,
            "Relatório_URI": path_uri
        })

if not dados_relatorio_geral:
    print("Nenhum relatório individual encontrado. O Relatório Geral não será gerado.")
else:
    # 4. Criar o DataFrame
    df_geral = pd.DataFrame(dados_relatorio_geral)

    # 5. Definir o caminho para o novo arquivo Excel
    # Salva o relatório geral um nível acima, na pasta 'dir_relatorios'
    nome_arquivo_geral = f"Relatorio Geral {aba_planilha_base}.xlsx"
    path_geral = os.path.join(dir_output, nome_arquivo_geral)

    print(f"Escrevendo Relatório Geral em: {path_geral}")

    # 6. Escrever o arquivo Excel usando XlsxWriter
    try:
        with pd.ExcelWriter(path_geral, engine='xlsxwriter') as writer_geral:
            # Escrever o DataFrame (apenas as colunas visíveis)
            df_geral_output = df_geral[["Local Instalação", "Caso", "Relatório_Texto"]]
            df_geral_output.to_excel(
                writer_geral,
                sheet_name='Resumo',
                index=False,
                header=["Local Instalação", "Caso", "Relatório"]
                # Renomeia 'Relatório_Texto' para 'Relatório' no header
            )

            workbook_geral = writer_geral.book
            worksheet_geral = writer_geral.sheets['Resumo']

            # 7. Adicionar os hyperlinks
            # Obter o índice da coluna 'Relatório' (que é 2, ou 'C')
            col_idx_relatorio = df_geral_output.columns.get_loc("Relatório_Texto")

            # Formato padrão para o link
            url_format = workbook_geral.add_format({'color': 'blue', 'underline': 1})

            # Escrever os URLs na coluna 'Relatório'
            # (start_row=1 para pular o cabeçalho)
            for row_num, (link_uri, link_text) in enumerate(zip(df_geral['Relatório_URI'], df_geral['Relatório_Texto']),
                                                            start=1):
                worksheet_geral.write_url(
                    row_num,
                    col_idx_relatorio,
                    link_uri,
                    cell_format=url_format,
                    string=link_text
                )

            # 8. Adicionar formatação de tabela (como na função original)
            (num_linhas, _) = df_geral_output.shape
            intervalo_tabela = f'A1:C{num_linhas + 1}'  # Colunas A, B, C
            worksheet_geral.add_table(intervalo_tabela, {
                'columns': [
                    {'header': 'Local Instalação'},
                    {'header': 'Caso'},
                    {'header': 'Relatório de Variações'}
                ]
            })

            # 9. Ajustar a largura das colunas
            try:
                largura_local = max(df_geral['Local Instalação'].astype(str).map(len).max(),
                                    len("Local Instalação")) + 2
                largura_caso = max(df_geral['Caso'].astype(str).map(len).max(), len("Caso")) + 2
                largura_relatorio = max(df_geral['Relatório_Texto'].astype(str).map(len).max(), len("Relatório")) + 2

                worksheet_geral.set_column('A:A', largura_local)
                worksheet_geral.set_column('B:B', largura_caso)
                worksheet_geral.set_column('C:C', largura_relatorio)
            except Exception as e:
                print(f"Aviso: Não foi possível ajustar largura das colunas. {e}")

        print(f"Relatório Geral '{nome_arquivo_geral}' criado com sucesso na pasta '{dir_output}'.")

    except PermissionError:
        print(f"ERRO: Permissão negada. O arquivo '{path_geral}' pode estar aberto. Feche o arquivo e tente novamente.")
    except Exception as e:
        print(f"ERRO: Falha ao escrever o Relatório Geral. {e}")

print("\n--- Fim da Execução do Script ---")
print(fr"Relatórios exportados para a pasta {dir_output}")
print("Pressione Esc para sair.")
while True:
    if keyboard.is_pressed('esc'):  # Detecta a tecla Esc
        print("Tecla Esc pressionada. Saindo...")
        break
