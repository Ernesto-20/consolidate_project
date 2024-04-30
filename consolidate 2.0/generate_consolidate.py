import os

from src.consolidate import consolidate

# if __name__ == "__main__":
file_list = []
files_names = []
name_model = None
file_model = None

for file in os.listdir('./'):
    if file.__contains__('.xlsx') and file[0] != '~' and file.__contains__('Reporte') == False and file.__contains__('Cons-sem') == False:
        init = file.find('-') + 1
        end = file.find('-') + 4
        files_names.append(file[init: end])
        file_list.append(file)
    elif file.__contains__('Cons-sem__.mgzn'):
        name_model = 'Cons-sem__.mgzn.'
        file_model = file

# print('****')
# print(file_list)
consolidate(file_list, files_names, file_model, name_model)