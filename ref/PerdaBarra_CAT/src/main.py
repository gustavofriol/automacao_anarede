from functions import le_num_barras, escreveDCTG, leDADB

arquivo = "../data/barrasTensaoNumANA.xlsx"
nomeColuna = "Num"

numBarras = le_num_barras(arquivo, nomeColuna)

dadb = leDADB("../data/BARRASREGIME.txt")


# Garantir que Num_Barra em dadb seja do tipo inteiro
# dadb["Num_Barra"] = dadb["Num_Barra"].astype(int)

# Converter numBarras (string) para inteiros
numBarras = [int(b) for b in numBarras]

# Filtrar numBarras para manter apenas os valores que estão em dadb["Num_Barra"]
numBarras = [b for b in numBarras if b in set(dadb["Num_Barra"])]


escreveDCTG(numBarras,  nomeArquivo="../data/dctg_barras.pwf")

