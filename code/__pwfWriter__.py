from pathlib import Path
import re
import __diretorios__ as diretorios
import os


class Writer:
    def __init__(self, pwf_dir: str, casos_dir: str):
        self.pwf_dir = pwf_dir
        self.casos_dir = casos_dir

    def get_pwf_list(self) -> tuple:
        """
        Acessa uma pasta onde os arquivos pwf para cada caso do ONS estão salvos e retorna o nome do SAV selecionado e
        uma lista com o nome de cada arquivo.
        """
        pasta_casos = Path(self.casos_dir)
        arquivos = [file.name for file in pasta_casos.glob("*")]
        indices = list(range(len(arquivos)))

        lista_casos = "Arquivos SAV disponíveis:\n"

        # Exibindo arquivos e pedindo um input
        for i in range(len(indices)):
            lista_casos = lista_casos + f"[{indices[i] + 1}] {arquivos[i]}\n"

        sav_selected = arquivos[int(input(f"{lista_casos}\n Selecione o arquivo: ")) - 1]

        # Armazenando casos do SAV e retornando em uma lista
        pasta_pwfs = Path(self.casos_dir + fr"\{sav_selected}")
        arquivos_pwf = [file.name for file in pasta_pwfs.glob("*.PWF")]

        return sav_selected, arquivos_pwf

    ########################################################################################################################

    def select_pwf(self) -> tuple:

        ############################# Lógica para selecionar caso #############################

        pasta_casos = Path(self.casos_dir)
        arquivos = [file.name for file in pasta_casos.glob("*")]
        indices = list(range(len(arquivos)))

        lista_casos = "Casos disponíveis:\n"

        for i in range(len(indices)):
            lista_casos = lista_casos + f"[{indices[i] + 1}] {arquivos[i]}\n"

        caso_selecionado = arquivos[int(input(f"{lista_casos}\nSelecione o caso: ")) - 1]

        ############################# Lógica para selecionar PWF #############################

        pasta_pwfs = Path(self.casos_dir + fr"\{caso_selecionado}")
        arquivos_pfw = [file.name for file in pasta_pwfs.glob("*.PWF")]
        indices_arquivos = list(range(len(arquivos_pfw)))

        lista_pwfs = "\nEstudos Elétricos:\n"

        for i in range(len(indices_arquivos)):
            lista_pwfs = lista_pwfs + f"[{indices_arquivos[i] + 1}] {arquivos_pfw[i]}\n"

        indice_selecionado = int(input(f"{lista_pwfs}\n Agora escolha como quer reestabelecer o caso: "))
        pwf_selecionado = arquivos_pfw[indice_selecionado - 1]

        return indice_selecionado, caso_selecionado, pwf_selecionado

    ########################################################################################################################

    def write_caso_zero(self, nome_arquivo_pwf: str, nome_relatorio_saida: str, escolhas: tuple):

        # Definir local onde arquivo do caso zero vai ficar
        caso_zero = fr"{self.pwf_dir}\{nome_arquivo_pwf}"

        # --- INÍCIO DA MODIFICAÇÃO ---
        # 1. Criar o caminho ABSOLUTO para o relatório .txt
        path_relatorio_abs = os.path.join(self.pwf_dir, nome_relatorio_saida)

        # 2. Criar o caminho ABSOLUTO para o diretório de sinal
        sinal_dir_abs = os.path.join(self.pwf_dir, 'regime.signal')
        sinal_dir_dos = f'"{sinal_dir_abs}"'  # Adiciona aspas para o MKDIR
        # --- FIM DA MODIFICAÇÃO ---

        conteudo = rf"""
ULOG                                                                 
2
{self.casos_dir}\{escolhas[1]}\{escolhas[1]}.SAV

ARQV REST
{escolhas[0]}                                                                                                                                          

DOPC
FILE L
RCVG D
RMON L
MFCT L
EMRG D
99999

DMTE
(tp) (no ) C (tp) (no ) C (tp) (no ) C (tp) (no ) O F
AREA 1     A AREA 6
AREA 51    A AREA 54
AREA 101   A AREA 105
AREA 401   A AREA 405
99999
DMFL
(tp) (no ) C (tp) (no ) C (tp) (no ) C (tp) (no ) O I
AREA 1     A AREA 6
AREA 51    A AREA 54
AREA 101   A AREA 105
AREA 401   A AREA 405
99999

ULOG
4
{path_relatorio_abs}

EXLF MOSF MOST CPER DADL
(% F)(% T)
0.001 0.1
RELA

ULOG
4

CASO

DOSC
MKDIR {sinal_dir_dos}
99999

FIM 
            """

        with open(caso_zero, "w") as v0:
            v0.write(conteudo)

    ########################################################################################################################

    @staticmethod
    def write_dctg(df):
        dctg = ""
        for i in range(len(df)):
            num_barra_de = str(df.loc[i, 'Núm. Barra DE'])
            num_barra_para = str(df.loc[i, 'Núm. Barra PARA'])
            nc = str(df.loc[i, 'NC'])

            dctg += f"""
(Nc) O Pr (       IDENTIFICACAO DA CONTINGENCIA        )
{i + 1:>{4}}    1 {num_barra_de}-{num_barra_para}

(Tp) (El ) (Pa ) Nc (Ext) (DV1) (DV2) (DV3) (DV4) (DV5) (DV6) (DV7) Gr Und
CIRC {num_barra_de:>{5}} {num_barra_para:>{5}} {nc:>{2}}                                             
FCAS"""
        return dctg

    ########################################################################################################################

    def write_contingencias(self, nome_arquivo_pwf: str, nome_relatorio_saida: str, escolhas: tuple, dataframe):
        dctg = self.write_dctg(dataframe)

        pwf_contingencias = fr"{self.pwf_dir}\{nome_arquivo_pwf}"

        # --- INÍCIO DA MODIFICAÇÃO ---
        # 1. Criar o caminho ABSOLUTO para o relatório .txt
        path_relatorio_abs = os.path.join(self.pwf_dir, nome_relatorio_saida)

        # 2. Criar o caminho ABSOLUTO para o diretório de sinal
        sinal_dir_abs = os.path.join(self.pwf_dir, 'ctgs.signal')
        sinal_dir_dos = f'"{sinal_dir_abs}"'  # Adiciona aspas para o MKDIR
        # --- FIM DA MODIFICAÇÃO ---

        conteudo = rf"""
ULOG                                                                 
2
{self.casos_dir}\{escolhas[1]}\{escolhas[1]}.SAV

ARQV REST
{escolhas[0]}                                                                                                                                          

DOPC
FILE L
RCVG D
RMON L
MFCT L
EMRG D
99999

DMTE
(tp) (no ) C (tp) (no ) C (tp) (no ) C (tp) (no ) O F
AREA 1     A AREA 6
AREA 51    A AREA 54
AREA 101   A AREA 105
AREA 401   A AREA 405
99999
DMFL
(tp) (no ) C (tp) (no ) C (tp) (no ) C (tp) (no ) O I
AREA 1     A AREA 6
AREA 51    A AREA 54
AREA 101   A AREA 105
AREA 401   A AREA 405
99999

ULOG
4
{path_relatorio_abs}

DCTG
{dctg}
99999

EXCT MOSF MOST CPER
(% F)(% T)
0.001 0.1

(P Pr Pr Pr Pr Pr Pr Pr Pr Pr Pr Pr
 1

ULOG
4

CASO

DOSC
MKDIR {sinal_dir_dos}
99999

FIM 
            """

        with open(pwf_contingencias, "w") as v0:
            v0.write(conteudo)

    ########################################################################################################################

    def write_dadb_sav(self, nome_arquivo_pwf: str, nome_relatorio_saida: str, sav: str, pwf_files: list):

        pwf_files = [item.removesuffix('.PWF') for item in pwf_files]
        pwf_text = ''
        pwf_path = fr"{self.pwf_dir}\{nome_arquivo_pwf}_{sav}.pwf"

        casos_posicoes = []

        padrao = r'^\d+'

        for texto in pwf_files:
            match = re.search(padrao, texto)

            if match:
                numero = match.group(0)
                casos_posicoes.append(numero)

        dict_sav = dict(zip(casos_posicoes, pwf_files))

        for posicao, nome_pwf in dict_sav.items():
            text_caso = rf"""
ULOG                                                                 
2
{self.casos_dir}\{sav}\{sav}.SAV

ARQV REST
{posicao}                                                                                                                                          

DOPC
FILE L
RCVG D
RMON L
MFCT L
EMRG D
99999

ULOG
4
{diretorios.dir_pwf}/{nome_relatorio_saida}_{nome_pwf}.txt

EXLF DADB
RELA

ULOG
4

CASO
                        """
            pwf_text = pwf_text + text_caso + '\n'


        sinal_dir_abs = os.path.join(self.pwf_dir, 'regime.signal')
        sinal_dir_dos = f'"{sinal_dir_abs}"'

        pwf_text = pwf_text + rf"""
DOSC
MKDIR {sinal_dir_dos}
99999
"""

        pwf_text = pwf_text + '\n' + 'FIM'

        with open(pwf_path, "w") as v0:
            v0.write(pwf_text)


    ########################################################################################################################

    def write_dctg_perda_barra_sav(self, nome_arquivo_pwf: str, nome_relatorio_saida: str, lista_barras: list, sav: str,
                                   pwf_files: list):

        pwf_files = [item.removesuffix('.PWF') for item in pwf_files]
        pwf_path = fr"{self.pwf_dir}\{nome_arquivo_pwf}_{sav}.pwf"

        casos_posicoes = []

        padrao = r'^\d+'
        for texto in pwf_files:
            match = re.search(padrao, texto)
            if match:
                numero = match.group(0)
                casos_posicoes.append(numero)

        dict_sav = dict(zip(casos_posicoes, pwf_files))

        dctg = ''
        for i, barra in enumerate(lista_barras, start=1):
            dctg = dctg + f'{str(i).rjust(4)}    1 Desligamento Barra: {str(barra)}\n' + f'BARD {str(barra).rjust(5)}\nFCAS\n'

        dctg = dctg + '99999\n'

        pwf_txt = ''
        for posicao, nome_pwf in dict_sav.items():
            pre_dctg = rf"""
ULOG                                                                 
2
{self.casos_dir}\{sav}\{sav}.SAV

ARQV REST
{posicao}                                                                                                                                          

DOPC
FILE L
RMON L
MFCT L
EMRG D
99999

DMTE
(tp) (no ) C (tp) (no ) C (tp) (no ) C (tp) (no ) O F
AREA 1     A AREA 6
AREA 51    A AREA 54
AREA 101   A AREA 105
AREA 401   A AREA 405
99999
DMFL
(tp) (no ) C (tp) (no ) C (tp) (no ) C (tp) (no ) O I
AREA 1     A AREA 6
AREA 51    A AREA 54
AREA 101   A AREA 105
AREA 401   A AREA 405
99999

DCAR
(tp) (no ) C (tp) (no ) C (tp) (no ) C (tp) (no ) O (A) (B) (C) (D) (Vfl)
AREA  101  A AREA  105                                  100     100
AREA  1    A AREA  6                                    100     100      
AREA 51    A AREA 54                                    100     100
AREA 401   A AREA 405                                   100     100
99999 

EXLF

ULOG
4
{diretorios.dir_pwf}/{nome_relatorio_saida}_{nome_pwf}.txt

DCTG
"""
            pos_dctg = """
EXCT DADB MOSF RCVG CPER
(% F)(% T)
115


(P Pr Pr Pr Pr Pr Pr Pr Pr Pr Pr Pr
 1

ULOG
4

CASO  
"""

            pwf_txt = pwf_txt + pre_dctg + dctg + pos_dctg


        sinal_dir_abs = os.path.join(self.pwf_dir, 'ctgs.signal')
        sinal_dir_dos = f'"{sinal_dir_abs}"'

        pwf_txt = pwf_txt + rf"""
DOSC
MKDIR {sinal_dir_dos}
99999
    """

        pwf_txt = pwf_txt + 'FIM'

        with open(pwf_path, "w") as f:
            f.write(pwf_txt)

########################################################################################################################