#creacion del libro de excel y la importacion de openpyxl
import openpyxl
miarchivo=openpyxl.Workbook()
mihoja=miarchivo.active
mihoja["A1"]=5
mihoja["A2"]=7
mihoja['B1']='hola'
mihoja['B2']="python"
miarchivo.save('hoja_excel.xlsx')
miarchivo=openpyxl.load_workbook('hoja_excel.xlsx')

#imprimo cada valor de la hoja de excel, por separado
print('valor de la posicion A1 es:',mihoja['A1'].value)
print('valor de la posicion A2 es:', mihoja['A2'].value)
print('valor de la posicion B1 es: ', mihoja['B1'].value)
print('valor de la posi                      cion B2 es :', mihoja['B2'].value)

#imprimir cada fila  la matriz de excel en una lista.
lmatriz=mihoja['A1':'B2']
for valor in lmatriz:
        print([lmatriz.value for lmatriz  in valor])
#agregar varias lista 
filas=[[4,6,4],[4,4,7],['print1',5,'excel']]
mihoja.append(['col1','col2','col3'])
for i in filas:
        mihoja.append(i)
#cambiar el valor a una celda
mihoja['A1']='DIANA'
miarchivo.save('hoja_excel.xlsx')
print('El nuevo valor de A1 es : ' , mihoja['A1'].value)


