import shutil
from pathlib import Path

# 1. Configuração dos Caminhos
PASTA_TRABALHO = Path(r"C:\Users\0341642\OneDrive - eletrobras.com\Área de Trabalho\automacao_anarede\code")
PASTA_EXECUCAO = Path(r"L:\DOS\Gustavo\Automação Anarede")
ARQUIVO_IGNORADO = "__diretorios__.py"


def realizar_deploy():
    print(f"Iniciando deploy para: {PASTA_EXECUCAO}")

    # Validação de segurança: verifica se a origem existe
    if not PASTA_TRABALHO.exists():
        print(f"Erro: Pasta de trabalho não encontrada em {PASTA_TRABALHO}")
        return

    # Cria a pasta de execução caso ela não exista (primeira vez)
    PASTA_EXECUCAO.mkdir(parents=True, exist_ok=True)

    # --- PASSO 1: Limpeza da pasta de execução ---
    print("Limpando scripts antigos da rede...")
    for arquivo_dest in PASTA_EXECUCAO.glob("*.py"):
        if arquivo_dest.name != ARQUIVO_IGNORADO:
            try:
                arquivo_dest.unlink()  # Deleta o arquivo
                print(f"  Removido: {arquivo_dest.name}")
            except Exception as e:
                print(f"  Erro ao remover {arquivo_dest.name}: {e}")

    # --- PASSO 2: Cópia dos novos scripts ---
    print("\nCopiando novos scripts para a rede...")
    arquivos_copiados = 0

    for arquivo_orig in PASTA_TRABALHO.glob("*.py"):
        if arquivo_orig.name != ARQUIVO_IGNORADO:
            try:
                destino_final = PASTA_EXECUCAO / arquivo_orig.name
                shutil.copy2(arquivo_orig, destino_final)
                print(f"  Copiado: {arquivo_orig.name}")
                arquivos_copiados += 1
            except Exception as e:
                print(f"  Erro ao copiar {arquivo_orig.name}: {e}")

    print(f"\n--- Deploy concluído com sucesso! ---")
    print(f"Total de arquivos atualizados: {arquivos_copiados}")
    print(f"O arquivo {ARQUIVO_IGNORADO} foi preservado em ambos os locais.")


if __name__ == "__main__":
    realizar_deploy()