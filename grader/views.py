#!/usr/bin/env python
# -*- coding: utf-8 -*-

import StringIO
import tempfile
import json
from random import *
from django.shortcuts import render
from django.http import HttpResponse
from django.template import RequestContext, loader

from funciones import *
# Create your views here.


def index(request):
	ruta = 'C:/Users/Saul/Desktop/calificador/'
	template = loader.get_template('grader/index.html')
	nota_tc,nota_notapie,nota_letracap,nota_saltos,nota_vinetas,nota_columnas,nota_piepagina,notas_formato,notas_bordes= None,None,None,None,None,None,None,None,None
	total = 0
	arr = []
	token = ''
	if request.POST:
		if request.FILES['modelo']:
			modelo_doc = request.FILES['modelo']
			document = Document(modelo_doc)
			
			doc_name = modelo_doc.name
			doc_modelo = ruta + doc_name
			directorio = request.FILES.getlist('directorio')
			token = generar_token()
			for d in directorio:
				doc_respuesta = d.name
				#print doc_respuesta
				doc_respuesta = ruta + 'respuestas/' + d.name
				modelo = document.element
				respuesta = Document(d).element
				modelo2 = document #cargar_documet(doc_modelo) 
				respuesta2 = Document(d) #cargar_documet(doc_respuesta)
				nota_tc = validar_tc(modelo2,respuesta2)
				nota_notapie = validar_notapie(modelo_doc, d)
				nota_letracap = validar_letracap(modelo, respuesta)
				nota_saltos = validar_saltos(modelo, respuesta)
				nota_vinetas = validar_vinetas(modelo, respuesta)
				nota_columnas = validar_columnas(modelo, respuesta)
				nota_piepagina = validar_piepagina(modelo_doc, d)
				notas_formato = validar_formato(modelo2, respuesta2)
				notas_bordes = validar_bordes(modelo2, respuesta2)
				total = sum(nota_tc)+nota_notapie+nota_letracap+nota_saltos+nota_vinetas+nota_columnas+nota_piepagina+sum(notas_formato)+sum(notas_bordes)
				calif = {'nota_tc':nota_tc,'tot_tc':sum(nota_tc),'nota_notapie':nota_notapie,'nota_letracap':nota_letracap,'nota_saltos':nota_saltos,'nota_vinetas':nota_vinetas,'notas_bordes':notas_bordes,'tot_bordes':sum(notas_bordes),
											'nota_columnas':nota_columnas,'nota_piepagina':nota_piepagina,'notas_formato':notas_formato,'tot_formato':sum(notas_formato),'total':total,'archivo':d.name[:30]}
				arr.append(calif)
				#print calif
				
	context = RequestContext(request,{'notas':arr,'token':token}) #{'nota_tc':nota_tc,'nota_notapie':nota_notapie,'nota_letracap':nota_letracap,'nota_saltos':nota_saltos,'nota_vinetas':nota_vinetas,'notas_bordes':notas_bordes,
										#'nota_columnas':nota_columnas,'nota_piepagina':nota_piepagina,'notas_formato':notas_formato,'total':total,'notas':arr})
	guardar_notas_tmp(arr,token)
	return HttpResponse(template.render(context))
	
def get_excel(notas):
	import xlsxwriter
	from xlsxwriter.workbook import Workbook
	output = StringIO.StringIO()
	workbook = Workbook(output)
	worksheet = workbook.add_worksheet("NotasWord")
	#Estilos
	estilos = {}
	estilos['titulo'] = workbook.add_format({'font_name': 'Arial','bold':1, 'font_size': '10','bg_color':'#ffd966'})
	estilos['titulo_gr'] = workbook.add_format({'font_name': 'Arial','bold':1, 'font_size': '9','bg_color':'#c6e0b4','text_wrap':1,'align':'center','bottom':2,'left':2,'right':2})
	estilos['titulo_min'] = workbook.add_format({'font_name': 'Arial','bold':1, 'font_size': '8','bg_color':'#c6e0b4','text_wrap':1,'align':'center','bottom':2})
	estilos['normal'] = workbook.add_format({'font_name': 'Arial', 'font_size': '11'})
	estilos['normal_neg'] = workbook.add_format({'font_name': 'Arial', 'font_size': '11','bold':1})
	worksheet.set_column('A:A', 5)
	worksheet.set_column('B:B', 25)
	worksheet.set_column('C:S', 8.5)
	worksheet.set_column('T:T', 10)
	#Cabecera
	worksheet.merge_range('A1:A2','No.', estilos['titulo'])
	worksheet.merge_range('B1:B2','Estudiante', estilos['titulo'])
	worksheet.merge_range('C1:E1','Tabla de Contenido', estilos['titulo_gr'])
	worksheet.write('C2', 'Genera TC', estilos['titulo_min'])
	worksheet.write('D2', 'Aplica Estilos', estilos['titulo_min'])
	worksheet.write('E2', 'Asocia Multinivel', estilos['titulo_min'])
	worksheet.merge_range('F1:J1','Formato del documento', estilos['titulo_gr'])
	worksheet.write('F2', 'Interlineado', estilos['titulo_min'])
	worksheet.write('G2', 'Espaciado', estilos['titulo_min'])
	worksheet.write('H2', 'Fuente', estilos['titulo_min'])
	worksheet.write('I2', 'Tamano', estilos['titulo_min'])
	worksheet.write('J2', 'Color', estilos['titulo_min'])
	worksheet.merge_range('K1:M1','Bordes', estilos['titulo_gr'])
	worksheet.write('K2', 'Color', estilos['titulo_min'])
	worksheet.write('L2', 'Contorno', estilos['titulo_min'])
	worksheet.write('M2', 'Grosor', estilos['titulo_min'])
	worksheet.merge_range('N1:N2','Nota al pie', estilos['titulo_gr'])
	worksheet.merge_range('O1:O2','Pie de pagina', estilos['titulo_gr'])
	worksheet.merge_range('P1:P2','Letra Capital', estilos['titulo_gr'])
	worksheet.merge_range('Q1:Q2','Columnas', estilos['titulo_gr'])
	worksheet.merge_range('R1:R2','Saltos', estilos['titulo_gr'])
	worksheet.merge_range('S1:S2','Vinetas', estilos['titulo_gr'])
	worksheet.merge_range('T1:T2','Total', estilos['titulo_gr'])
	#Registros
	fil = 2
	num = 1
	for r in notas:
		col = 0
		worksheet.write(fil, col, num, estilos['normal'])
		worksheet.write(fil, col+1, unicode(smart_str(r['archivo']), 'utf-8'), estilos['normal'])
		col=2
		for n in range(len(r['nota_tc'])):
			worksheet.write(fil, col+n, r['nota_tc'][n], estilos['normal'])
		col=col+n+1
		for n in range(len(r['notas_formato'])):
			worksheet.write(fil, col+n, r['notas_formato'][n], estilos['normal'])
		col=col+n+1
		for n in range(len(r['notas_bordes'])):
			worksheet.write(fil, col+n, r['notas_bordes'][n], estilos['normal'])
		col=col+n+1
		worksheet.write(fil, col, r['nota_notapie'], estilos['normal'])
		worksheet.write(fil, col+1, r['nota_piepagina'], estilos['normal'])
		worksheet.write(fil, col+2, r['nota_letracap'], estilos['normal'])
		worksheet.write(fil, col+3, r['nota_columnas'], estilos['normal'])
		worksheet.write(fil, col+4, r['nota_saltos'], estilos['normal'])
		worksheet.write(fil, col+5, r['nota_vinetas'], estilos['normal'])
		worksheet.write(fil, col+6, r['total'], estilos['normal_neg'])
		fil = fil+1
		num = num+1
	workbook.close()
	output.seek(0)
	return output.read()
	
def exportar_excel(request,token):
    #token = request.GET['token']
    notas = get_notas_tmp(token)
    response = HttpResponse(get_excel(notas), 'application/vnd.ms-excel')
    response['Content-Disposition'] = 'attachment;filename=NotasWord.xlsx'
    return response

def generar_token():
	import string
	allchar = string.digits #+ string.ascii_letters 
	token = "".join(choice(allchar) for x in range(randint(10, 10)))
	return token
	
def guardar_notas_tmp(arr,token):
	text = json.dumps(arr)
	tmp_path = tempfile.gettempdir()
	doc_path = tmp_path+"/"+token+".txt"
	doc = open(doc_path,'w')
	doc.write(text.encode('utf-8'))
	doc.close()

def get_notas_tmp(token):
	tmp_path = tempfile.gettempdir()
	doc_path = tmp_path+"/"+token+".txt"
	doc = open(doc_path,'r')
	return json.loads(doc.read())