import geocoder
import requests
import xlrd
import xlwt
from xlutils.copy import copy

#Pasa la dirección a Geocoder y regresa la latitud y longitud si la encuentra, en caso contrario regresa vacio.
def getCoords(address):
	g = geocoder.google(address)
	return g.latlng

#Recorre el archivo con direcciones y regresa un archivo con la dirección más las coordenadas.
def recorreArchivo():
    #Libro con direcciones
    workbook = xlrd.open_workbook("librocondirecciones.xls")
    #Copia el libro para crear otro y colocar las coordenadas
	wb = copy(workbook)
    #Obtiene la hoja del libro donde se encuentran las direcciones, en este caso solo hay una
	rdSheet = workbook.sheet_by_index(0)
    #Selecciona la hoja del libro donde se guardaran las direcciones
	wbSheet = wb.get_sheet(0)
    #recorre cada fila del archivo con direcciones
	for row in range(rdSheet.nrows):
		data = rdSheet.row_values(row)
        #Obtiene el string del valor de la columna donde se encuentra la dirección del libro que contiene las direcciones, en este caso es la columna B
        address = data[1]
        #Pasa la dirección al método getCoords para obtener las coordenadas
		latlong = getCoords(address)
        #Escribe las direcciones en el libro nuevo
		wbSheet.write(row, 2, str(latlong))
		print row
	wb.save("libroconcoordenadas.xls")
    #Termina
	print "nothing to do"

recorreArchivo()