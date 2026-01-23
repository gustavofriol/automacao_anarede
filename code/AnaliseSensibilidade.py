import pandas as pd
import os
import __diretorios__ as direorios

########################################################################################################################

dir_savs = fr'{direorios.dir_relatorios}\N-1'
savs = [entrada.name for entrada in os.scandir(dir_savs) if entrada.is_dir()]

while True:
    try:
        print('Casos disponíveis:')

        if savs:
            for n, sav in enumerate(savs):
                print(f"[{n}] {sav}")
        else:
            print(f"Nenhum relatório N-1 foi encontrado em {dir_savs}.")

        entry = input(f"\nSelecione um caso [0-{len(savs) - 1}]: ")
        choice = int(entry)

        if 0 <= choice <= len(savs) - 1:
            break
        else:
            print(f"Erro: O número {choice} está fora do intervalo. Digite um valor entre 0 e {len(savs) - 1}.\n")

    except ValueError:
        print(f"Erro: A entrada '{entry}' não é um número inteiro válido. Tente novamente.\n")

sav = savs[choice]

########################################################################################################################

dir_relarorios = rf'{dir_savs}\{sav}'
relatorios = [entrada.name for entrada in os.scandir(dir_relarorios) if entrada.is_dir()]

while True:
    try:
        print('\nRelatórios disponíveis:')
        if relatorios:
            for n, folder in enumerate(relatorios):
                print(f"[{n}] {folder}")
        else:
            print(f"Nenhum relatório N-1 foi encontrado em {dir_relarorios}.")

        entry = input(f"\nSelecione um relatório [0-{len(relatorios) - 1}]: ")
        choice = int(entry)

        if 0 <= choice <= len(relatorios) - 1:
            break
        else:
            print(f"Erro: O número {choice} está fora do intervalo. Digite um valor entre 0 e {len(relatorios) - 1}.\n")

    except ValueError:
        print(f"Erro: A entrada '{entry}' não é um número inteiro válido. Tente novamente.\n")

relatorio = relatorios[choice]

########################################################################################################################

dir_files = fr'{dir_relarorios}\{relatorio}'
file_names = []
final_dict = {}

for file in os.listdir(dir_files):

    file_path = os.path.join(dir_files, file)

    if os.path.isfile(file_path) and file.lower().endswith('.xlsx'):
        file_name = file[:-5]
        file_names.append(file_name)

        #df = pd.read_excel(file_path)

print(file_names)


