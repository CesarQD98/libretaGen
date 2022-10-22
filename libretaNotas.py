import openpyxl
import pprint

# Open and read the Workbook
print('Opening workbook...')
wb = openpyxl.load_workbook('./data/notas.xlsx')
# print(wb.sheetnames)
sheet = wb['1°E']
alumnoData = {}

GRADO = sheet['B3'].value.strip('°')
SECCION = sheet['B5'].value.upper()

# TODO: Fill in alumnoData with each student info.
print('Reading rows...')
row = 10
while(True):
    name = sheet['B' + str(row)].value

    if name == None:
        break

    curso = sheet['C8'].value
    aptitud = sheet['C9'].value
    nota = sheet['C' + str(row)].value
    nroOrden = row - 9
    
    # Adds a new student to the dict
    alumnoData.setdefault(name, {})

    # TODO: Adds 'notas', more than course.
    # Adds 'grado', 'seccion', 'nroOrden' and 'cursos'
    alumnoData[name].setdefault('grado', GRADO)
    alumnoData[name].setdefault('seccion', SECCION)
    alumnoData[name].setdefault('nroOrden', nroOrden)
    alumnoData[name].setdefault('cursos', {})
    alumnoData[name]['cursos'].setdefault('curso1', curso)

    print(row)
    row += 1



# Expected output for each student:
# alumnoData = {'AGUIRRE PRADO...': {'grado': 1, 'seccion': 'e',
#                                    'cursos': {'desarrollo...': {'aptitud1': 'A', 'aptitud2': 'B'}}}}

    



# print('B47: ' + str(sheet['B47'].value))
# print('C8: ' + sheet['C8'].value)
# print('D8: ' + sheet['C8'].value)
# print('GRADO: ' + str(GRADO))
# print('SECCION: ' + str(SECCION))
# print('max_row: ' + str(sheet.max_row))
# print('max_column: ' + str(sheet.max_column))
# print()
print(pprint.pformat(alumnoData))


