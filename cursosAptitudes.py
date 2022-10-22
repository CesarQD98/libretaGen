import openpyxl
import pprint

wb = openpyxl.load_workbook('./data/notas.xlsx')
sheet = wb['1°E']

# TODO: Save all courses and data into a txt or a dict/list/tuple (need to decide)
tuplaCursos = tuple(sheet['C8':'AB8'])
tuplaAptitudes = tuple(sheet['C9':'AB9'])

cursosAptitudes = []

for rowOfCellObjects in sheet['C8':'AB8']:
    for cellObj in rowOfCellObjects:
        if cellObj.value != None:
            cursosAptitudes.append(cellObj.value)

# print(tuplaCursos)
print(tuplaAptitudes)

print(cursosAptitudes)
print(len(cursosAptitudes))