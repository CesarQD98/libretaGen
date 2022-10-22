from operator import index
import openpyxl
import pprint

# Open and read the Workbook
print('Opening workbook...')
wb = openpyxl.load_workbook('./data/notas.xlsx')
sheet = wb['1°E']
alumnoData = {}

# TODO: Create a dict or list that contains every course and its aptitudes
GRADO = sheet['B3'].value.strip('°')
SECCION = sheet['B5'].value.upper()
CURSOSAPTITUDES = {'Desarrollo Personal, Ciudadanía y Cívica': ['Construye su identidad', 'Convive y participa democráticamente en la búsqueda del bien común'], 'Ciencias Sociales': [
    'Construye interpretaciones históricas', 'Gestiona responsablemente el espacio y el ambiente', 'Gestiona responsablemente los recursos económicos']}

# Fill in alumnoData with each student info.
print('Reading rows...')
fila = 10
while (True):
    notasCol = 3
    name = sheet['B' + str(fila)].value

    if name == None:
        break

    aptitud = sheet['C9'].value
    nota = sheet['C' + str(fila):'G' + str(fila)]
    nroOrden = fila - 9

    # Adds a new student to the dict
    alumnoData.setdefault(name, {})

    # Adds 'grado', 'seccion', 'nroOrden' and 'cursos'
    alumnoData[name].setdefault('grado', GRADO)
    alumnoData[name].setdefault('seccion', SECCION)
    alumnoData[name].setdefault('nroOrden', nroOrden)
    alumnoData[name].setdefault('cursos', {})

    for k, v in CURSOSAPTITUDES.items():
        alumnoData[name]['cursos'].setdefault(k, {})
        for i in range(len(v)):
            alumnoData[name]['cursos'][k].setdefault(v[i], sheet.cell(row=fila, column=(notasCol + i)).value)
        
    fila += 1


# Expected output for each student:
# alumnoData = {'AGUIRRE PRADO...': {'grado': 1, 'seccion': 'e',
#                                    'cursos': {'desarrollo...': {'aptitud1': 'A', 'aptitud2': 'B'}}}}

print(pprint.pformat(alumnoData))
