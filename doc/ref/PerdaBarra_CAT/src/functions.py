import pandas as pd
import os

#-----------------------------------------------------------------------------------------------------------------
# Leitura dosnúmeoros da barras a partir de um excel
def le_num_barras(arquivo, nomeColuna):
    if not os.path.exists(arquivo):
        print("Arquivo não encontrado.")
        return None
    
    try:
    
        df = pd.read_excel(arquivo, dtype={nomeColuna: str})
    
        if nomeColuna not in df.columns:
            print(f"Coluna '{nomeColuna}' não encontrada no arquivo.")
            return None
        
        return sorted(df[nomeColuna].dropna().tolist(), key=lambda x: int(x))
    
    except Exception as e:
        print(f"Erro ao ler o arquivo: {e} ")
        return None

    
#-----------------------------------------------------------------------------------------------------------------
# Escreve DCTG a partir dos numeros das barras
def escreveDCTG(listaBarras, nomeArquivo="dctg_barras.pwf"):
    
    with open(nomeArquivo, "w") as f:
        
        f.write("DCTG\n")
        
        for i, barra in enumerate(listaBarras, start=1):
            f.write(f"{str(i).rjust(4)}    1 Desligamento Barra: {str(barra)}\n")
            f.write(f"BARD {str(barra).rjust(5)}\nFCAS\n")
            
        f.write("99999\nFIM")
    
#-----------------------------------------------------------------------------------------------------------------
# Lê o relatório emitido pelo DADB

def leDADB(nomeArquivo):
    
    # Definição das colunas com base na estrutura observada (ajustar conforme necessário)
    colunas = [
        (1, 7),    # Número da barra
        (8, 20),   # Nome da barra
        (21, 23),  # Tipo da Barra
        (24, 27),  # Área
        (35, 40),  # Módulo Tensão
        (41, 47),  # Ângulo Tensão
        (58, 60),  # Grupo Base Tensão
        (61, 63),  # Grupo Limite Tensão
        (64, 71),  # Geração MW Atual
        (72, 79),  # Geração Mvar Mínima
        (80, 87),  # Geração Mvar Atual
        (88, 95),  # Geração Mvar Máxima
        (96, 103),  # Carga MW
        (104, 111),  # Carga Mvar
        (112, 119),  # Shunt Mvar
        (120, 123),  # Estado
        (124, 128)  # TipoLigação
    ]

    # Nomes das colunas correspondentes
    nomes_colunas = [
        "Num_Barra", "Nome_Barra", "TipoBarra", "Area", "Tensao_mod", "Tensao_angulo",
        "GrpBaseTensao", "GrpLimiteTensao", "GeraMW_atual", "GeraMvar_min", "GeraMvar_atual", "GeraMvar_max",
        "Carga_MW", "Carga_Mvar", "Shunt_Mvar", "Estado", "TipoLigacao"
    ]

    # Lendo o arquivo com pandas
    df = pd.read_fwf(nomeArquivo, colspecs=colunas, header=None, names=nomes_colunas, encoding = "latin-1")
    
    # Forçar conversão da coluna "Num_Barra" para inteiro e remover linhas inválidas
    df["Num_Barra"] = pd.to_numeric(df["Num_Barra"], errors="coerce")  # Converte para número (NaN se falhar)
    df = df.dropna(subset=["Num_Barra"])  # Remove linhas onde a conversão falhou
    
    # Converter colunas numéricas para float
    colunas_float = ["Tensao_mod", "Tensao_angulo", "GeraMW_atual", "GeraMvar_min", "GeraMvar_atual",
                     "GeraMvar_max", "Carga_MW", "Carga_Mvar", "Shunt_Mvar"]
    
    df[colunas_float] = df[colunas_float].apply(pd.to_numeric, errors="coerce")  # Converte para float
    df = df.dropna(subset=colunas_float)  # Remove linhas com valores inválidos nas colunas float

    

    return df

#-----------------------------------------------------------------------------------------------------------------
# Lê relatório de saída MOCT

def leMOST(nomeArquivo):
    
    df = open(nomeArquivo).readlines()
    df = pd.DataFrame(df, columns=['linha'])
    
    textoBusca = "  Barras Desligadas por Ilhamento:"
    
    indIlha = df['linha'].str.startswith(textoBusca)
    indIlha = df.index[indIlha].tolist()
    
    
    textoBusca = "                            TOTAL:"
    
    indTotal = df['linha'].str.startswith(textoBusca)
    indTotal = df.index[indTotal].tolist()
    
    
    textoBusca = " ------ IND SEVER."
    
    indSev = df['linha'].str.startswith(textoBusca)
    indSev = df.index[indSev].tolist()
    
    textoBusca = " MONITORACAO DE TENSAO SELECIONADA"
    
    indMOST = df['linha'].str.startswith(textoBusca)
    indMOST = df.index[indMOST].tolist()
    
    return indMOST
 
 
    
saida = leMOST("../data/BARRASCONTINGENCIAS.txt")   
    

print(saida[1])
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    