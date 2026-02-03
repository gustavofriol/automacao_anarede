from pathlib import Path

anarede_exe_path = Path(r"C:\Cepel\Anarede\V120100\ANAREDE.exe")

# Definição das pastas
dir_banco_de_dados = fr"C:\Users\0341642\OneDrive - eletrobras.com\Área de Trabalho\automacao_anarede\exe\data"
dir_pwf = r"C:\Users\0341642\OneDrive - eletrobras.com\Área de Trabalho\automacao_anarede\exe\anarede"
dir_casos_referencia = r"C:\Users\0341642\OneDrive - eletrobras.com\Área de Trabalho\automacao_anarede\exe\data\Casos de Referência"
dir_relatorios = r"C:\Users\0341642\OneDrive - eletrobras.com\Área de Trabalho\automacao_anarede\relatorios"

# Nome dos relatórios txt e xlsx exportados pelo Anarede
file_txt_ctgs = 'RELATORIO_CTGS.txt'
file_txt_regime = 'RELATORIO_REGIME.txt'
file_txt_dadb = 'DADB'
file_txt_perda_barra = 'CTGS_PERDA_BARRA'

file_xlsx_ctgs = 'FLUXOS_CTGS.xlsx'
file_xlsx_regime = 'FLUXOS_CASO_ZERO.xlsx'
file_xslx_num_barras = 'barrasTensaoNumANA.xlsx'

file_pwf_ctgs = 'DADOS_CTGS.pwf'
file_pwf_regime = 'DADOS_REGIME.pwf'
file_pwf_dadb = 'DADB_REGIME'
file_pwf_perda_barra = 'PERDA_DE_BARRA'


# Caminho completo para os arquivos
path_ctgs_txt = fr'{dir_pwf}\{file_txt_ctgs}'
path_regime_txt = fr'{dir_pwf}\{file_txt_regime}'
path_dadb_txt = fr'{dir_pwf}\{file_txt_dadb}'

path_ctgs_xlsx = fr'{dir_pwf}\{file_xlsx_ctgs}'   # REVISAR
path_zero_xlsx = fr'{dir_pwf}\{file_xlsx_regime}' # REVISAR

file_num_barras = fr'{dir_banco_de_dados}\{file_xslx_num_barras}'


# Planilha de ativos da Eletrosul
planilha_base_n1 = fr"{dir_banco_de_dados}\Planilha Base - Automação Anarede - N-1.xlsx"
planilha_base_perda_barra = fr"{dir_banco_de_dados}\Planilha Base - Automação Anarede - Perda de Barra.xlsx"