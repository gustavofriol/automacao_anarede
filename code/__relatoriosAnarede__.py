import pandas as pd
import numpy as np
import re, os, shutil


class RelatorioAnarede:
    def __init__(self, arquivo_entrada: str):
        self.arquivo_entrada = arquivo_entrada
        self.linhas = []

    def read(self):
        try:
            with open(self.arquivo_entrada, 'r') as arquivo:
                self.linhas = arquivo.readlines()
        except Exception as e:
            print(f'Erro ao ler o arquivo: {e}')
            return []
        return self.linhas

    @staticmethod
    def criar_de_para_nc(row, col_de, col_para, col_nc):
        """
        row: linha que a função recebe quando chamada com "apply"
        col_de: Nome da coluna DE do dataframe
        col_para: Nome da coluna PARA do dataframe
        col_nc: Nome da coluna NC do dataframe
        """
        de = row[col_de]
        para = row[col_para]
        nc = row[col_nc]
        menor = min(de, para)
        maior = max(de, para)
        return f'{menor}-{maior}-{nc}'


class Mosf(RelatorioAnarede):
    @staticmethod
    def _extrair_secoes(linhas, num_ctgs=1):
        secoes_extraidas = [[] for _ in range(num_ctgs)]
        dentro_secao = False
        num_ctg = 0

        for linha in linhas:
            match = re.search(r" CONTINGENCIA\s+(\d+)", linha)
            if match:
                num_ctg = int(match.group(1)) - 1

            if linha.startswith(" X------------X------------X--X-------X-------X-------X--------X---------------X"):
                dentro_secao = True
            elif ("CEPEL" in linha and dentro_secao) or ("IND SEVER" in linha and dentro_secao) or (
                    linha.startswith(" ONS") and dentro_secao):
                dentro_secao = False
            elif len(linha) < 2:
                continue
            elif dentro_secao:
                secoes_extraidas[num_ctg].append(linha)

        for i in range(num_ctgs):
            if secoes_extraidas[i] == []:
                secoes_extraidas[i] = [False] # Problema: cria false que não permite que o apply seja aplicado na criar de-para-nc

        return secoes_extraidas


    @staticmethod
    def extrair_lista_ctgs(df_ctgs):
        """ Extrai lista de contingências a partir do dataframe que faz o dctg"""
        lista_dctgs = []
        for i in range(len(df_ctgs)):
            de = int(df_ctgs.loc[i, 'Núm. Barra DE'])
            para = int(df_ctgs.loc[i, 'Núm. Barra PARA'])
            nc = df_ctgs.loc[i, 'NC']
            menor = min(de, para)
            maior = max(de, para)
            de_para_nc = f'{menor}-{maior}-{nc}'
            lista_dctgs.append(de_para_nc)
        return lista_dctgs

    @staticmethod
    def _processar_secao(secao):
        if secao == [False]:
            colunas = ['Núm. Barra DE', 'Nome Barra DE', 'Núm. Barra PARA', 'Nome Barra PARA',
                       'NC', 'MW', 'MVAR', 'MVA/V', 'Violação MVA', 'Carregamento %', 'DE-PARA-NC'
                       ]
            df = pd.DataFrame(columns=colunas)
            return df

        tabela_1 = []
        tabela_2 = []

        for i, linha in enumerate(secao):
            if linha == False:
                return False
            if i % 2 == 0:
                tabela_1.append(linha.strip().split())
            else:
                tabela_2.append(linha.strip().split())

        for tabela in [tabela_1, tabela_2]:
            for linha in tabela:
                if linha:
                    linha.pop()

        df1 = pd.DataFrame(tabela_1, columns=['Núm. Barra DE', 'Núm. Barra PARA', 'Carregamento %'])
        df2 = pd.DataFrame(tabela_2, columns=['Nome Barra DE', 'Nome Barra PARA', 'NC', 'MW', 'MVAR',
                                              'MVA/V', 'Violação MVA'])

        if not df1.empty and not df2.empty:
            df_nome_barra_de = df2.loc[:, 'Nome Barra DE']
            df_nome_barra_para = df2.loc[:, 'Nome Barra PARA']
            df_num_barra_de = df1.loc[:, 'Núm. Barra DE']
            df_num_barra_para = df1.loc[:, 'Núm. Barra PARA']
            df_carregamento = df1.loc[:, 'Carregamento %']
            df2_resto = df2.loc[:, 'NC':]

            df_final = pd.concat([df_num_barra_de, df_nome_barra_de, df_num_barra_para, df_nome_barra_para,
                                  df2_resto, df_carregamento], axis=1)

            df_final['MW'] = pd.to_numeric(df_final['MW'], errors='coerce')
            df_final['MVAR'] = pd.to_numeric(df_final['MVAR'], errors='coerce')
            df_final['MVA/V'] = pd.to_numeric(df_final['MVA/V'], errors='coerce')
            df_final['Violação MVA'] = pd.to_numeric(df_final['Violação MVA'], errors='coerce')
            df_final['Carregamento %'] = pd.to_numeric(df_final['Carregamento %'], errors='coerce')
            df_final['Núm. Barra DE'] = pd.to_numeric(df_final['Núm. Barra DE'], errors='coerce')
            df_final['Núm. Barra PARA'] = pd.to_numeric(df_final['Núm. Barra PARA'], errors='coerce')

            return df_final

    def extrair_df_caso_zero(self):
        linhas = self.read()
        if not linhas:
            return None
        secao_extraida = self._extrair_secoes(linhas, num_ctgs=1)[0]
        return self._processar_secao(secao_extraida)

    def extrair_dfs_contingencias(self, df_ctgs):
        linhas = self.read()
        if not linhas:
            return []
        secoes_extraidas = self._extrair_secoes(linhas, num_ctgs=len(df_ctgs))
        dfs = [self._processar_secao(secao) for secao in secoes_extraidas if self._processar_secao(secao) is not None]
        return dfs


class Most(RelatorioAnarede):
    @staticmethod
    def _extrair_secoes(linhas, num_ctgs=1):
        secoes_extraidas = [[] for _ in range(num_ctgs)]
        dentro_secao = False
        num_ctg = 0

        for linha in linhas:
            match = re.search(r" CONTINGENCIA\s+(\d+)", linha)
            if match:
                num_ctg = int(match.group(1)) - 1

            if linha.startswith(" X-----X------------X---X------X------X------X--------X--------X--------X---------------X"):
                dentro_secao = True
            elif ("CEPEL" in linha and dentro_secao) or ("IND SEVER" in linha and dentro_secao) or (
                    linha.startswith(" ONS") and dentro_secao):
                dentro_secao = False
            elif len(linha) < 2:
                continue
            elif dentro_secao:
                secoes_extraidas[num_ctg].append(linha)
        return secoes_extraidas

    @staticmethod
    def _processar_secao(secao):

        dados = []
        for linha in secao:
            linha = linha.strip()  # Remove leading/trailing whitespace
            if not linha:  # Skip empty lines
                continue
            partes = linha.split()
            if not partes:  # Skip lines that become empty after splitting
                continue
            try:
                num_barra = partes[0]
                nome_barra_parts = []
                indice = 1
                while indice < len(partes) and not partes[indice].replace('.', '', 1).isdigit():
                    nome_barra_parts.append(partes[indice])
                    indice += 1
                nome_barra = " ".join(nome_barra_parts)
                area = partes[indice]
                min_tensao = partes[indice + 1]
                tensao_mod = partes[indice + 2]
                max_tensao = partes[indice + 3]
                violacao = partes[indice + 4]
                shunt_info = " ".join(partes[indice + 5:])
                shunt_parts = shunt_info.split()
                shunt_bar = shunt_parts[0] if shunt_parts else None
                shunt_lin = shunt_parts[1] if len(shunt_parts) > 1 else None
                severidade = " ".join(shunt_parts[2:]) if len(shunt_parts) > 2 else None

                dados.append(
                    [num_barra, nome_barra, area, min_tensao, tensao_mod, max_tensao, violacao, shunt_bar, shunt_lin,
                     severidade])
            except IndexError:
                print(f"Erro ao processar a linha: '{linha}'")
                print("A linha não possui o número esperado de elementos.")
                # Decide como lidar com a linha com erro: ignorar, registrar, levantar uma exceção, etc.
                # For now, we'll just skip it.
                continue
        df = pd.DataFrame(dados, columns=['Núm. Barra', 'Nome Barra', 'Área', 'Mín.', 'Tensão MOD.', 'Máx.',
                                          'Violação(PU)', 'SHUNTBAR (Mvar)', 'SHUNTLIN (Mvar)', 'Severidade'])

        if not df.empty:

            df['Núm. Barra'] = pd.to_numeric(df['Núm. Barra'], errors='coerce')
            df['Mín.'] = pd.to_numeric(df['Mín.'], errors='coerce')
            df['Tensão MOD.'] = pd.to_numeric(df['Tensão MOD.'], errors='coerce')
            df['Máx.'] = pd.to_numeric(df['Máx.'], errors='coerce')
            df['Violação(PU)'] = pd.to_numeric(df['Violação(PU)'], errors='coerce')
            df['SHUNTBAR (Mvar)'] = pd.to_numeric(df['SHUNTBAR (Mvar)'], errors='coerce')
            df['SHUNTLIN (Mvar)'] = pd.to_numeric(df['SHUNTLIN (Mvar)'], errors='coerce')

            return df
        return None

    def extrair_df_caso_zero(self):
        linhas = self.read()
        if not linhas:
            return None
        secao_extraida = self._extrair_secoes(linhas, num_ctgs=1)[0]
        return self._processar_secao(secao_extraida)

    def extrair_dfs_contingencias(self, df_ctgs):
        linhas = self.read()
        if not linhas:
            return []
        secoes_extraidas = self._extrair_secoes(linhas, num_ctgs=len(df_ctgs))
        dfs = [self._processar_secao(secao) for secao in secoes_extraidas if self._processar_secao(secao) is not None]
        return dfs


class Mvt(RelatorioAnarede):
    @staticmethod
    def _extrair_secoes(linhas, num_ctgs=1):
        secoes_extraidas = [[] for _ in range(num_ctgs)]
        dentro_secao = False
        num_ctg = 0

        for linha in linhas:
            match = re.search(r" CONTINGENCIA\s+(\d+)", linha)
            if match:
                num_ctg = int(match.group(1)) - 1

            if linha.startswith(" X-----X------------X------X------X--------X--------X"):
                dentro_secao = True

            elif ("CEPEL" in linha and dentro_secao) or (linha.startswith(" ONS") and dentro_secao) or (
                "CONTINGENCIA" in linha and dentro_secao) or ("MONITORACAO" in linha and dentro_secao):
                dentro_secao = False
            elif len(linha) < 2:
                continue
            elif dentro_secao:
                secoes_extraidas[num_ctg].append(linha)
        return secoes_extraidas

    @staticmethod
    def _processar_secao(secao):
        linhas_proc = []
        aviso_sem_variacao = ' Nao foram encontradas variacoes de tensao acima do percentual informado entre as barras monitoradas.\n'

        for linha in secao:
            if linha == aviso_sem_variacao:
                return False
            linhas_proc.append(linha.strip().split())

        df = pd.DataFrame(linhas_proc, columns=['Núm_Barra', 'Nome_Barra', 'Tensão_Pré',
                                                'Tensão_Pós', 'Variação_PU', 'Variação_%'])

        df['Tensão_Pré'] = pd.to_numeric(df['Tensão_Pré'], errors='coerce')
        df['Tensão_Pós'] = pd.to_numeric(df['Tensão_Pós'], errors='coerce')
        df['Variação_PU'] = pd.to_numeric(df['Variação_PU'], errors='coerce')
        df['Variação_%'] = pd.to_numeric(df['Variação_%'], errors='coerce')

        return df

    def extrair_dfs_contingencias(self, df_ctgs):
        linhas = self.read()
        if not linhas:
            return []
        secoes_extraidas = self._extrair_secoes(linhas, num_ctgs=len(df_ctgs))
        dfs = [self._processar_secao(secao) for secao in secoes_extraidas if self._processar_secao(secao) is not None]
        return dfs


class Dadb(RelatorioAnarede):
    @staticmethod
    def _extrair_secoes(linhas, num_ctgs=1):
        secoes_extraidas = [[] for _ in range(num_ctgs)]
        dentro_secao = False
        num_ctg = 0

        for linha in linhas:
            match = re.search(r" CONTINGENCIA\s+(\d+)", linha)
            if match:
                num_ctg = int(match.group(1)) - 1

            if linha.startswith(" X-----X------------X--X---X------X-----X------X-----X---X--X--X-------X-------X-------X-------X-------X-------X-------X---X---X---X---X---X---X---X---X---X---X---X---X"):
                dentro_secao = True
            elif ("CEPEL" in linha and dentro_secao) or (linha.startswith(" ONS") and dentro_secao):
                dentro_secao = False
            elif len(linha) < 2:
                continue
            elif dentro_secao:
                secoes_extraidas[num_ctg].append(linha)
        return secoes_extraidas

    @staticmethod
    def _processar_secao(secao):
        linhas_proc = []

        # Nomes das colunas correspondentes
        nomes_colunas = [
            "Num_Barra", "Nome_Barra", "area", "Tensao_mod", "GeraMW_atual", "Gera_Mvar_atual", "Carga_MW", "Carga_Mvar", "Estado"
        ]

        for linha in secao:
            num_barra = linha[0:8].strip().split()
            nome_barra = linha[8:21].strip().split()
            area = linha[24:28].strip().split()
            tensao_mod = linha[35:41].strip().split()
            geracao_mw = linha[64:72].strip().split()
            geracao_mvar = linha[80:88].strip().split()
            carga_mw = linha[98:104].strip().split()
            carga_mvar = linha[104:112].strip().split()
            estado = linha[120:124].strip().split()

            linha_final = num_barra+nome_barra+area+tensao_mod+geracao_mw+geracao_mvar+carga_mw+carga_mvar+estado

            linhas_proc.append(linha_final)

        linhas_drop = []
        for linha in linhas_proc:
            if len(linha) == 10:
                linhas_proc.remove(linha)
                linhas_drop.append(linha)

        df = pd.DataFrame(linhas_proc, columns=nomes_colunas)

        # CONVERSÃO PARA NUMÉRICO (LINHAS ADICIONADAS/MODIFICADAS)
        df["Num_Barra"] = pd.to_numeric(df["Num_Barra"], errors="coerce")
        df["Tensao_mod"] = pd.to_numeric(df["Tensao_mod"], errors="coerce")
        df["GeraMW_atual"] = pd.to_numeric(df["GeraMW_atual"], errors="coerce")
        df["Gera_Mvar_atual"] = pd.to_numeric(df["Gera_Mvar_atual"], errors="coerce")
        df["Carga_MW"] = pd.to_numeric(df["Carga_MW"], errors="coerce")
        df["Carga_Mvar"] = pd.to_numeric(df["Carga_Mvar"], errors="coerce")
        df["area"] = pd.to_numeric(df["area"], errors="coerce")

        return df, linhas_drop


    def extrair_dfs_contingencias(self, df_ctgs):
        linhas = self.read()
        if not linhas:
            return []
        secoes_extraidas = self._extrair_secoes(linhas, num_ctgs=len(df_ctgs))

        dfs =[]
        linhas_drop = []

        for n, secao in enumerate(secoes_extraidas):
            print(f'Processando contingência {n}.')
            df, linhas_drop = self._processar_secao(secao)
            dfs.append(df)
            linhas_drop.append(linhas_drop)

        print(f'Linhas não processadas devido à erro:\n{linhas_drop}\nAnalisar manualmente se for o caso.')

        return dfs

    def read_dadb(self):
        linhas = self.read()
        if not linhas:
            return None
        secao_extraida = self._extrair_secoes(linhas, num_ctgs=1)[0]
        df, _ = self._processar_secao(secao_extraida)
        return df

    @staticmethod
    def get_num_barras(dir_xlsx, nome_coluna, sheet):
        """
        Leitura dos números das barras a partir de um excel.
        Retorna uma lista de int com os números das barras.
        'nome_coluna' deve receber o nome da coluna do excel onde estão armazenados os dados de barra.
        """
        if not os.path.exists(dir_xlsx):
            print("Arquivo não encontrado.")
            return None

        try:
            df = pd.read_excel(dir_xlsx, dtype={nome_coluna: str}, sheet_name=sheet)
            if nome_coluna not in df.columns:
                print(f"Coluna '{nome_coluna}' não encontrada no arquivo.")
                return None

            return sorted(df[nome_coluna].dropna().tolist(), key=lambda x: int(x))

        except Exception as e:
            print(f"Erro ao ler o arquivo: {e} ")
            return None

    def get_lista_barras(self, dir_xlsx, nome_coluna, sheet):
        """
        Deve ser usado com arquivo de entrada igual ao DADB.
        """
        df_dadb = self.read_dadb()
        lista_barras = self.get_num_barras(dir_xlsx=dir_xlsx, nome_coluna=nome_coluna, sheet=sheet)
        lista_barras_int = [int(b) for b in lista_barras]
        lista_barras_atualizado = [b for b in lista_barras_int if b in set(df_dadb["Num_Barra"])]
        return lista_barras_atualizado

    def get_ctgs_keys(self, lista_barras):
        """
        Deve ser usado com arquivo de entrada igual ao RELATORIO_PERDA_BARRA
        """
        num_ctgs = len(lista_barras)
        # LINHA CORRIGIDA: Inicializa com um valor de string padrão.
        ctgs_keys = [f"CTG_NAO_PROCESSADA_PARA_BARRA_{barra}" for barra in lista_barras]

        linhas = self.read()
        num_ctg = 0
        num_barras = []

        for n, linha in enumerate(linhas):
            match_ctg = re.search(r" CONTINGENCIA\s+(\d+)", linha)
            match_info = re.search(r"Numero da BARRA:\s+(\d+)\s+Nome da BARRA:\s+([\w-]+)", linha)
            if match_ctg:
                num_ctg = int(match_ctg.group(1)) - 1

            if match_info:
                numero_barra = match_info.group(1)
                num_barras.append(int(numero_barra))
                nome_barra = match_info.group(2)

                # Garante que não estamos tentando acessar um índice fora dos limites
                if 0 <= num_ctg < len(ctgs_keys):
                    ctgs_keys[num_ctg] = f"{numero_barra}-{nome_barra}"

        # Para otimizar a busca, ainda é uma boa prática converter a segunda lista para um conjunto
        set_num_barras = set(num_barras)

        # A list comprehension itera sobre lista_barras e inclui o item apenas se ele não estiver no conjunto
        num_barras_excluidos = [barra for barra in lista_barras if barra not in set_num_barras]

        print(f"\nBarras do Excel não encontradas no relatório: {num_barras_excluidos}")
        return ctgs_keys


class Ilha(RelatorioAnarede):
    """
    Ao instanciar essa classe e chamar a função 'extrair_dfs_contingencias', a ideia é que seja armazenada no objeto
    os dataframes com ilhamento referente a cada contingência.
    """
    @staticmethod
    def _extrair_secoes(linhas, num_ctgs=1):
        secoes_extraidas = [[] for _ in range(num_ctgs)]
        dentro_secao = False
        num_ctg = 0

        for n, linha in enumerate(linhas):
            match = re.search(r" CONTINGENCIA\s+(\d+)", linha)
            if match:
                num_ctg = int(match.group(1)) - 1

            if linha.startswith(" X---------------------------------X-----X-------------X----------X------------X") and 'Barras Desligadas por Ilhamento:' not in linhas[n+1]:
                dentro_secao = True
            elif ("CEPEL" in linha and dentro_secao) or ("TOTAL:" in linha and dentro_secao) or (
                    linha.startswith(" ONS") and dentro_secao):
                dentro_secao = False
            elif len(linha) < 2:
                continue
            elif dentro_secao:
                secoes_extraidas[num_ctg].append(linha)
        return secoes_extraidas

    @staticmethod
    def _processar_secao(secao):
        cols_1_2 = slice(0, 56)
        col_3 = slice(57, 67)
        col_4 = slice(68, 80)

        linhas_proc = []

        for linha in secao:
            # 1. Extrai as partes da string usando as fatias predefinidas
            partes_1_e_2 = linha[cols_1_2].strip().split()

            # Usa .strip() para remover espaços. Se não houver nada, o resultado é ''
            parte_3 = linha[col_3].strip()
            parte_4 = linha[col_4].strip()

            # 2. Constrói a linha final, substituindo strings vazias por NaN
            # Uma string vazia ('') é considerada 'False' em um contexto booleano
            coluna_3_final = parte_3 if parte_3 else np.nan
            coluna_4_final = parte_4 if parte_4 else np.nan

            # 3. Monta a lista completa da linha
            nova_linha = partes_1_e_2 + [coluna_3_final, coluna_4_final]

            linhas_proc.append(nova_linha)


        df = pd.DataFrame(linhas_proc, columns=['Num', 'Nome', 'Carga_MW', 'Geracao_MW'])

        df['Carga_MW'] = pd.to_numeric(df['Carga_MW'], errors='coerce')
        df['Geracao_MW'] = pd.to_numeric(df['Geracao_MW'], errors='coerce')

        return df

    def extrair_dfs_contingencias(self, df_ctgs):
        linhas = self.read()
        if not linhas:
            return []
        secoes_extraidas = self._extrair_secoes(linhas, num_ctgs=len(df_ctgs))
        dfs = [self._processar_secao(secao) for secao in secoes_extraidas if self._processar_secao(secao) is not None]
        return dfs

class Dadl(RelatorioAnarede):
    @staticmethod
    def _extrair_secoes(linhas):
        secoes_extraidas = []
        dentro_secao = False

        for linha in linhas:

            if linha.startswith(" X-----X-----X--X---X---X---X-----X-------X-------X-------X------X------X------X-----X------X-X------X------X------X------------X------------X---X---X---X---X---X---X---X---X---X---X---X---X"):
                dentro_secao = True
            elif ("CEPEL" in linha and dentro_secao) or (linha.startswith(" ONS") and dentro_secao):
                dentro_secao = False
            elif len(linha) < 2:
                continue
            elif dentro_secao:
                secoes_extraidas.append(linha)
        return secoes_extraidas

    @staticmethod
    def _processar_secao(secao):
        linhas_proc = []

        # Nomes das colunas correspondentes
        nomes_colunas = ['de', 'para', 'nc']

        for linha in secao:
            de = linha[0:8].strip().split()
            para = linha[8:14].strip().split()
            nc = linha[14:17].strip().split()
            linha_final = de+para+nc

            linhas_proc.append(linha_final)


        df = pd.DataFrame(linhas_proc, columns=nomes_colunas)
        df["de"] = pd.to_numeric(df["de"], errors="coerce")
        df["para"] = pd.to_numeric(df["para"], errors="coerce")
        df['nc'] = pd.to_numeric(df['nc'], errors="coerce")

        return df

    def read_dadl(self):
        linhas = self.read()
        if not linhas:
            return None
        secao_extraida = self._extrair_secoes(linhas)
        df = self._processar_secao(secao_extraida)
        return df





