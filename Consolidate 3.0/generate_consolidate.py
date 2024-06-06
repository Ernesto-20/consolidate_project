import os

from src.consolidate import consolidate

# if __name__ == "__main__":
files = []
file_names = []
name_model = None
file_model = None
file_stock = None

for file in os.listdir('./'):
    if (file.__contains__('.xlsx') or file.__contains__('.xlsb')) and file[0] != '~' and file.__contains__('Reporte') == False and file.__contains__('Cons-sem') == False and file.__contains__('Stock') == False:
        init = file.find('-') + 1
        end = file.find('-') + 4
        file_names.append(file[init: end])
        files.append(file)
    elif file.__contains__('Cons-sem__.mgzn'):
        name_model = 'Cons-sem__.mgzn.'
        file_model = file
    elif file.__contains__('Stock-sem__.mgzn'):
        file_stock = file

consolidate(files, file_names, file_model, name_model, file_stock)
