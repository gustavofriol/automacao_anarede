import pandas as pd
import os

def le_num_barras(arquivo, nomeColuna):
    if not os.path.exists(arquivo):
        print("Arquivo não encontrado.")
        return None
    
    try:
    
        df = pd.read_excel(arquivo, dtype={nomeColuna: str})
    
        if nomeColuna not in df.columns:
            print(f"Coluna '{nomeColuna}' não encontrada no arquivo.")
            return None
        
        return df[nomeColuna].astype(str).tolist()
    
    except Exception as e:
        print(f"Erro ao ler o arquivo: {e} ")
        return None

    



