# coding=utf-8
from Levenshtein import ratio

#Devuelve un valor entre 0-1 para porcentaje de similaridad de cadenas
def comparar_cad(cadena1, cadena2):
   if ratio(cadena1,cadena2)>=0.8: return True
   return False

#Devuelve un valor entre 0 y 1 segun la igualdad entre ambos arreglos
def comparar_arreglos(arr1, arr2):
   if arr1==arr2: return 1
   if len(arr1)==0 or len(arr2)==0: return float(0)
   cont = 0
   for element2 in arr2:
      str2 = element2.encode('utf-8')
      for element1 in arr1:
	     str1 = element1.encode('utf-8')
	     if comparar_cad(str1,str2):
	        cont=cont+1
	        break
   porcentaje = float(cont)/float(len(arr1))
   return porcentaje
   
#Devuelve el puntaje segun el rango correspondiente al porcentaje
def get_puntaje_rango(rangos,porcentaje):
   for r in rangos:
      [(key,val)]=r.items()
      rango = key.split('-')
      min = float(rango[0])
      max = float(rango[1])
      if porcentaje>=min and porcentaje<=max:
         return val
   return 0
#methodList = [method for method in dir(columna) if callable(getattr(columna, method))]


#Devuelve un arreglo con la suma de los items de la misma posicion
def sumar_arrs(arr1,arr2):
   result = []
   n = len(arr1)
   for i in range(n): result.append(0)
   for i in range(n):
      result[i]=float(arr1[i])+float(arr2[i])
   return result
   
def get_texto(arr):
   texto=''
   for e in arr[1:-1]:
      texto=texto+e.text
   return texto