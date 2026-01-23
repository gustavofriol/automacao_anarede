import pandas as pd
from __pwfWriter__ import Writer
import __diretorios__ as diretorios
import __relatoriosAnarede__ as Rela
import subprocess
import os
import time
import shutil
import traceback


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


# Definição planilha de contingências
caminho_planilha_base = diretorios.planilha_base_perda_barra  # Junção diretório + nome da planilha de contingências
excel_base = pd.ExcelFile(caminho_planilha_base)
lista_abas = excel_base.sheet_names
nro_abas = 0

print("\nListas de contingências na planilha basse:")
for i, aba in enumerate(lista_abas):
    nro_abas = i + 1
    print(fr"[{i + 1}] {aba}")

indice_aba = int(
    input(f"\nSelecione pelo índice ({1}-{nro_abas}) a lista de contingências à executar: ")) - 1  # Corrigido para 1

aba_planilha_base = lista_abas[indice_aba]  # Aba da planilha de contingências selecionada
coluna_base = 'Num'  # Coluna da planilha onde se encontra o número das barras que serão desligadas

######################################## Escrever Deck Regime ##########################################################

writer = Writer(pwf_dir=diretorios.dir_pwf, casos_dir=diretorios.dir_casos_referencia)
sav_pwfs = writer.get_pwf_list()
sav = sav_pwfs[0]
pwf_files = sav_pwfs[1]

for pwf in pwf_files:
    # Obtendo PWF para gerar DADB do caso em regime
    writer.write_dadb_sav(
        nome_arquivo_pwf=diretorios.file_pwf_dadb,
        nome_relatorio_saida=diretorios.file_txt_dadb,
        sav=sav,
        pwf_files=pwf_files
    )

# --- CHAMADA 1 MODIFICADA ---
pwf_regime = fr'{diretorios.dir_pwf}/{diretorios.file_pwf_dadb}_{sav}.pwf'
# Define o nome do SINAL (diretório) que esperamos que o PWF crie
sinal_regime = os.path.join(diretorios.dir_pwf, 'regime.signal')
executar_anarede(pwf_regime, sinal_regime)  # Passa o sinal para a função
# --- FIM DA MODIFICAÇÃO ---

############################################# Escrever Decks CTGS ######################################################

pwf_files = [item.removesuffix('.PWF') for item in pwf_files]

# Itera sobre cada DADB e atualiza a lista de barras para fazer o DCTG
for nome_pwf in pwf_files:
    txt_path = fr"{diretorios.dir_pwf}\{diretorios.file_txt_dadb}_{nome_pwf}.txt"

    try:
        dadb = Rela.Dadb(arquivo_entrada=txt_path)
    except FileNotFoundError as e:
        print(e)
        print(
            "O arquivo txt não foi encontrado na pasta onde estão os PWFS.\n Certifique-se de que os cartões PWF já foram rodados no Anarede ")
        dadb = None

    if dadb is None:
        print(f"ERRO: DADB Nulo para {nome_pwf}. Pulando esta iteração.")
        continue  # Pula para o próximo pwf

    df_dadb = dadb.read_dadb()

    # Atualiza dados de barra com base no DADB do caso
    lista_barras = dadb.get_num_barras(dir_xlsx=caminho_planilha_base, nome_coluna=coluna_base, sheet=aba_planilha_base)
    lista_barras_int = [int(b) for b in lista_barras]
    lista_barras_atualizado = [b for b in lista_barras_int if b in set(df_dadb["Num_Barra"])]

    # Escreve DCTG com dados de barra atualizados
    writer = Writer(pwf_dir=diretorios.dir_pwf, casos_dir=diretorios.dir_casos_referencia)
    writer.write_dctg_perda_barra_sav(
        nome_arquivo_pwf=diretorios.file_pwf_perda_barra,
        nome_relatorio_saida=diretorios.file_txt_perda_barra,
        lista_barras=lista_barras_atualizado,
        sav=sav,
        pwf_files=pwf_files
    )

# --- CHAMADA 2 MODIFICADA ---
pwf_ctgs = fr'{diretorios.dir_pwf}/{diretorios.file_pwf_perda_barra}_{sav}.pwf'
# Define o nome do SINAL (diretório) que esperamos que o PWF crie
sinal_ctgs = os.path.join(diretorios.dir_pwf, 'ctgs.signal')
executar_anarede(pwf_ctgs, sinal_ctgs)  # Passa o sinal para a função
# --- FIM DA MODIFICAÇÃO ---

#################################### Exportar Resultados ###############################################################

# Dicionário que irá armazenar os resultados finais de todos os casos
dict_valores_geral = {}

# Critérios para a análise
CRITERIO_ILHAMENTO = 250
CRITERIO_SUBTENSAO = 300
CRITERIO_FLUXO = 150

for pwf in pwf_files:
    print(f"\n=======================================================")
    print(f"Processando Caso: '{pwf}'")
    print(f"=======================================================")

    nome_rela_pb = f'{diretorios.file_txt_perda_barra}_{pwf}.txt'
    nome_rela_dadb = f'{diretorios.file_txt_dadb}_{pwf}.txt'
    path_perda_barra = os.path.join(diretorios.dir_pwf, nome_rela_pb)
    path_dadb = os.path.join(diretorios.dir_pwf, nome_rela_dadb)

    # --- Bloco de try/except robustecido ---
    try:
        ilha = Rela.Ilha(arquivo_entrada=path_perda_barra)
        mosf = Rela.Mosf(arquivo_entrada=path_perda_barra)
        most = Rela.Most(arquivo_entrada=path_perda_barra)
        perda_barra = Rela.Dadb(arquivo_entrada=path_perda_barra)
        dadb = Rela.Dadb(arquivo_entrada=path_dadb)
    except FileNotFoundError as e:
        print(f"ERRO: Arquivos de relatório não encontrados para o caso {pwf}. Pulando este caso.")
        print(f"Arquivo faltante: {e.filename}")
        continue
    except Exception as e:
        print(f"Erro ao inicializar relatórios para o caso {pwf}: {e}")
        traceback.print_exc()
        continue
    # --- Fim do bloco ---

    df_ctgs = pd.read_excel(caminho_planilha_base, sheet_name=aba_planilha_base)
    dfs_ctgs_ilhas = ilha.extrair_dfs_contingencias(df_ctgs)
    dfs_ctgs_fluxo = mosf.extrair_dfs_contingencias(df_ctgs=df_ctgs)

    dfs_dadb = perda_barra.extrair_dfs_contingencias(df_ctgs)

    # --- Robustez adicionada para lidar com 'False' na lista ---
    for df in dfs_dadb:
        if isinstance(df, pd.DataFrame):
            df['Tensao_mod'] = pd.to_numeric(df['Tensao_mod'], errors='coerce')
            df['area'] = pd.to_numeric(df['area'], errors='coerce')
            df['area'] = df['area'].astype('Int64')

    dfs_dadb_filtrado = []

    limite_tensao = 0.9

    intervalos_area = []
    for start, end in [(1, 6), (51, 54), (101, 105), (401, 405)]:
        intervalos_area.extend(range(start, end + 1))

    intervalos_area_set = set(intervalos_area)

    for df in dfs_dadb:
        if isinstance(df, pd.DataFrame):  # Só processa se for um DataFrame
            condicao_tensao = df['Tensao_mod'] < limite_tensao
            condicao_area = df['area'].isin(intervalos_area_set)
            df_filtrado = df.loc[condicao_tensao & condicao_area]

            if not df_filtrado.empty:
                dfs_dadb_filtrado.append(df_filtrado)
            else:
                dfs_dadb_filtrado.append(False)
        else:
            dfs_dadb_filtrado.append(False)  # Adiciona False se o df original era False

    for i, df in enumerate(dfs_ctgs_fluxo):
        if df is False:
            dfs_ctgs_fluxo[i] = pd.DataFrame(
                columns=["Núm. Barra DE", "Nome Barra DE", "Núm. Barra PARA", "Nome Barra PARA", "NC", "MW", "MVAR",
                         "MVA/V", "Violação MVA", "Carregamento %"])

    for i, df in enumerate(dfs_dadb_filtrado):
        if df is False:
            dfs_dadb_filtrado[i] = pd.DataFrame(
                columns=["Num_Barra", "Nome_Barra", "area", "Tensao_mod", "GeraMW_atual", "Gera_Mvar_atual", "Carga_MW",
                         "Carga_Mvar", "Estado"])

    lista_barras = dadb.get_lista_barras(dir_xlsx=caminho_planilha_base, nome_coluna=coluna_base,
                                         sheet=aba_planilha_base)
    ctgs_keys = perda_barra.get_ctgs_keys(lista_barras=lista_barras)

    dict_ilhas, dict_mosf, dict_most = {}, {}, {}

    for nome, df in zip(ctgs_keys, dfs_ctgs_ilhas):
        dict_ilhas[nome] = df

    for nome, df in zip(ctgs_keys, dfs_ctgs_fluxo):
        dict_mosf[nome] = df

    for nome, df in zip(ctgs_keys, dfs_dadb_filtrado):
        dict_most[nome] = df

    # --- INÍCIO DA NOVA LÓGICA INTEGRADA ---
    print(f"  Calculando valores agregados para o relatório geral do caso '{pwf}'...")
    dados_caso_atual = []
    for ctg_key in ctgs_keys:
        df_ilhamento = dict_ilhas.get(ctg_key)
        df_tensao = dict_most.get(ctg_key)
        df_fluxo = dict_mosf.get(ctg_key)

        # Robustez adicionada para lidar com None/False nos dicts
        valor_ilhamento = df_ilhamento['Carga_MW'].sum() if isinstance(df_ilhamento, pd.DataFrame) else 0
        valor_subtensao = df_tensao['Carga_MW'].sum() if isinstance(df_tensao, pd.DataFrame) else 0
        valor_fluxo = df_fluxo['Carregamento %'].max() if (
                    isinstance(df_fluxo, pd.DataFrame) and not df_fluxo.empty) else 0

        dados_caso_atual.append({
            'Barra': ctg_key,
            'Ilhamento (MW)': valor_ilhamento,
            'Subtensão (MW)': valor_subtensao,
            'Fluxo (%)': valor_fluxo
        })
    # --- Fim da Lógica ---

    if dados_caso_atual:
        df_caso = pd.DataFrame(dados_caso_atual)
        dict_valores_geral[pwf] = df_caso
    else:
        print(f"  AVISO: Nenhuma barra processada para o caso '{pwf}'.")

    dir_sav = os.path.join(diretorios.dir_relatorios, 'Perda de Barra', sav)
    dir_caso = os.path.join(dir_sav, pwf)

    if not os.path.exists(dir_caso):
        print(f"  Criando diretório: {dir_caso}")
        os.makedirs(dir_caso)

    print(f"  Exportando relatórios individuais para o caso '{pwf}'...")
    for i in ctgs_keys:
        path_excel_individual = os.path.join(dir_caso, f'Perda da Barra - {i}.xlsx')
        with pd.ExcelWriter(path_excel_individual) as writer_individual:
            if isinstance(dict_ilhas.get(i), pd.DataFrame):
                dict_ilhas[i].to_excel(writer_individual, sheet_name='Ilhamentos', index=False)
            if isinstance(dict_most.get(i), pd.DataFrame):
                dict_most[i].to_excel(writer_individual, sheet_name='Monitoramento de Tensão', index=False)
            if isinstance(dict_mosf.get(i), pd.DataFrame):
                dict_mosf[i].to_excel(writer_individual, sheet_name='Monitoramento de Fluxo', index=False)

# --- INÍCIO DA FINALIZAÇÃO E EXPORTAÇÃO DO RELATÓRIO GERAL ---
print("\n\n--- Realizando análise dos critérios para o relatório geral... ---")
# ... (O resto do seu script, que já estava correto, continua aqui) ...
for caso_nome, df_caso in dict_valores_geral.items():
    cond_ilhamento = df_caso['Ilhamento (MW)'] > CRITERIO_ILHAMENTO
    cond_subtensao = df_caso['Subtensão (MW)'] > CRITERIO_SUBTENSAO
    cond_fluxo = df_caso['Fluxo (%)'] > CRITERIO_FLUXO
    mascara_final_nok = cond_ilhamento | cond_subtensao | cond_fluxo
    df_caso['Análise'] = 'OK'
    df_caso.loc[mascara_final_nok, 'Análise'] = 'NOK'

dfs_valores_para_unir = []
dfs_analises_para_unir = []
for caso_nome, df_caso in sorted(dict_valores_geral.items()):
    df_temp_valores = df_caso.drop(columns=['Análise']).set_index('Barra')
    df_temp_valores.columns = pd.MultiIndex.from_product([[caso_nome], df_temp_valores.columns])
    dfs_valores_para_unir.append(df_temp_valores)
    df_temp_analise = df_caso[['Barra', 'Análise']].set_index('Barra')
    df_temp_analise = df_temp_analise.rename(columns={'Análise': caso_nome})
    dfs_analises_para_unir.append(df_temp_analise)

if dfs_valores_para_unir and dfs_analises_para_unir:
    df_valores_final = pd.concat(dfs_valores_para_unir, axis=1).fillna(0)
    for col in df_valores_final.columns:
        if 'Ilhamento' in col[1] or 'Subtensão' in col[1]:
            df_valores_final[col] = df_valores_final[col].astype(int)
    df_valores_final.index.name = 'Barra'
    df_valores_final.columns.names = ['Caso', None]
    df_analises_final = pd.concat(dfs_analises_para_unir, axis=1).fillna('OK').reset_index()

    dir_raiz_perda_barra = os.path.join(diretorios.dir_relatorios, 'Perda de Barra')
    path_relatorio_final = os.path.join(dir_raiz_perda_barra, f'Relatorio Geral Perda de Barra - {sav}.xlsx')

    print(f"\n--- Gerando relatório geral com duas abas em: {path_relatorio_final} ---")

    try:
        with pd.ExcelWriter(path_relatorio_final) as writer_final:
            df_valores_final.to_excel(writer_final, sheet_name='valores')
            df_analises_final.to_excel(writer_final, sheet_name='análises', index=False)
        print("\nRelatório geral com as abas 'valores' e 'análises' gerado com sucesso!")
    except PermissionError:
        print(f"\nERRO: Permissão negada. O arquivo '{path_relatorio_final}' está aberto.")
        print("Feche o arquivo Excel e tente novamente.")
    except Exception as e:
        print(f"\nERRO inesperado ao salvar o relatório geral: {e}")
        traceback.print_exc()
else:
    print("\nNenhum dado foi processado para gerar o relatório final.")