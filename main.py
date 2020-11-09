#!/usr/bin/python
# coding=utf-8
# -*- coding: utf-8 -*-

from selenium import webdriver
from selenium.webdriver.support.select import Select
from selenium.webdriver.common.proxy import *
from datetime import date

import xlrd
import xlsxwriter
import os
import time
import datetime 

links = []

workbook = xlsxwriter.Workbook(''.join([os.getcwd(),'\\carros.xlsx']))
worksheet = workbook.add_worksheet()

worksheet.write(1,0,'Carro')
worksheet.write(1,1,u'Ano')
worksheet.write(1,2,u'Preço')
worksheet.write(1,3,u'Combustível')
worksheet.write(1,4,u'IPVA')
worksheet.write(1,5,u'Seguro')
worksheet.write(1,6,u'Revisões')
worksheet.write(1,7,u'Procedência')
worksheet.write(1,8,u'Garantia')
worksheet.write(1,9,u'Configuração')
worksheet.write(1,10,u'Porte')
worksheet.write(1,11,u'Lugares')
worksheet.write(1,12,u'Portas')
worksheet.write(1,13,u'Geração')
worksheet.write(1,14,u'Plataforma')
worksheet.write(1,15,u'Índice CNW')
worksheet.write(1,16,u'Ranking CNW')
worksheet.write(1,17,u'Nota do Leitor')

worksheet.write(0,18,u'Motor')

worksheet.write(1,18,u'Instalação')
worksheet.write(1,19,u'Aspiração')
worksheet.write(1,20,u'Disposição')
worksheet.write(1,21,u'Alimentação')
worksheet.write(1,22,u'Cilindros')
worksheet.write(1,23,u'Comando de Válvulas')
worksheet.write(1,24,u'Tuchos')
worksheet.write(1,25,u'Variação do Comando')
worksheet.write(1,26,u'Válvulas por Cilindro')
worksheet.write(1,27,u'Diâmetro dos Cilindros')
worksheet.write(1,28,u'Razão de Compressão')
worksheet.write(1,29,u'Curso dos Pistões')
worksheet.write(1,30,u'Cilindrada')
worksheet.write(1,31,u'Pontência Máxima')
worksheet.write(1,32,u'Código do Motor')
worksheet.write(1,33,u'Torque Máximo')
worksheet.write(1,34,u'Peso/Potência')
worksheet.write(1,35,u'Toque Específico')
worksheet.write(1,36,u'Peso/Torque')
worksheet.write(1,37,u'Potência Específica')
worksheet.write(1,38,u'Rotação Máxima')

worksheet.write(0,39,u'Transmissão')

worksheet.write(1,39,u'Tração')
worksheet.write(1,40,u'Câmbio')
worksheet.write(1,41,u'Código do Câmbio')
worksheet.write(1,42,u'Acoplamento')

worksheet.write(0,43,u'Suspensão')

worksheet.write(1,43,u'Dianteira')
worksheet.write(1,44,u'Traseira')
worksheet.write(1,45,u'Elemento Elástico')
worksheet.write(1,46,u'Elemento Elástico')

worksheet.write(0,47,u'Freios')

worksheet.write(1,47,u'Dianteiros')
worksheet.write(1,48,u'Traseiros')

worksheet.write(0,49,u'Direção')

worksheet.write(1,49,u'Assistência')
worksheet.write(1,50,u'Diâmetro Mínimo de giro')

worksheet.write(0,51,u'Pneus')

worksheet.write(1,51,u'Dianteiros')
worksheet.write(1,52,u'Traseiros')
worksheet.write(1,53,u'Altura do flanco')
worksheet.write(1,54,u'Altura do flanco')

worksheet.write(0,55,u'Dimensões')

worksheet.write(1,55,u'Comprimento')
worksheet.write(1,56,u'Largura')
worksheet.write(1,57,u'Distância entre-eixos')
worksheet.write(1,58,u'Altura')
worksheet.write(1,59,u'Bitola Dianteira')
worksheet.write(1,60,u'Bitola Traseira')
worksheet.write(1,61,u'Porta-malas')
worksheet.write(1,62,u'Tanque de Combustível')
worksheet.write(1,63,u'Peso')
worksheet.write(1,64,u'Carga Útil')
worksheet.write(1,65,u'Vão Livre do Solo')

worksheet.write(0,66,u'Aerodinâmica')

worksheet.write(1,66,u'Área Frontal')
worksheet.write(1,67,u'Coeficiente de Arrasto')
worksheet.write(1,68,u'Área frontal corrigida')

worksheet.write(1,69,u'Desempenho')

worksheet.write(1,69,u'Velocidade Máxima')
worksheet.write(1,70,u'Aceleração 0-100 km/h')

worksheet.write(0,71,u'Consumo')

worksheet.write(1,71,u'Urbano')
worksheet.write(1,73,u'Rodoviário')

worksheet.write(0,75,u'Autonomia')

worksheet.write(1,75,u'Urbana')
worksheet.write(1,77,u'Rodoviária')

worksheet.write(1,79,u'Segurança')
worksheet.write(1,80,u'Conforto')
worksheet.write(1,81,u'Infotenimento')

driver = webdriver.Firefox()
driver.get('https://www.carrosnaweb.com.br/catalogo.asp')
 
def main():

    while driver.find_element_by_xpath('//table[@width="100%"]//tbody//tr//td//font[@face="arial"]//a//b[contains(text(),"Próxima")]'):
        table()
        time.sleep(15)

        paginacao = driver.find_element_by_xpath('//table[@width="100%" and @border="0"]//tbody//tr//td[@align="right"]//font[@size="2" and @face="Arial" and @color="darkred"]').text 
        pagina = int(paginacao.replace('Página ','').replace(' de 1527',''))
        ultimapagina = int(paginacao[paginacao.find(' de ')+4:])
        
        print(' Página de {} de {}'.format(pagina,ultimapagina))

        if pagina > 5:
            break

        driver.find_element_by_xpath('//table[@width="100%"]//tbody//tr//td//font[@face="arial"]//a//b[contains(text(),"Próxima")]').click()
        
def table():
    for i,link in enumerate(driver.find_elements_by_xpath('//table[@width="770"]//tbody//tr//td//a')):
        if link.get_attribute('href').find('https://www.carrosnaweb.com.br/fichadetalhe.asp?codigo=') > -1:
            links.append(link.get_attribute('href'))
 
def excel(linha,link):
    driver.get(link)

    table = driver.find_elements_by_xpath('//table[@border="0" and @cellspacing="1" and @cellpadding="3" and @width="100%" ]//tbody//tr//td')
    
    worksheet.write(linha,0,table[4].text)  # Nome do Carro
    worksheet.write(linha,1,table[7].text)  # Ano
    worksheet.write(linha,2,table[9].text)  # Preço
    worksheet.write(linha,3,table[11].text) # Combustível
    worksheet.write(linha,4,table[13].text) # IPVA
    worksheet.write(linha,5,table[15].text) # Seguro
    worksheet.write(linha,6,table[17].text) # Revisões
    worksheet.write(linha,7,table[19].text) # Procedência
    worksheet.write(linha,8,table[21].text) # Garantia
    worksheet.write(linha,9,table[23].text) # Configuração
    worksheet.write(linha,10,table[25].text) # Porte
    worksheet.write(linha,11,table[27].text) # Lugares
    worksheet.write(linha,12,table[29].text) # Portas
    worksheet.write(linha,13,table[31].text) # Geração
    worksheet.write(linha,14,table[33].text) # Plataforma
    worksheet.write(linha,15,table[35].text) # Índice CNW
    worksheet.write(linha,16,table[37].text) # Ranking CNW
    worksheet.write(linha,17,table[39].text[0:table[39].text.find('Avalie')]) # Nota do Leitor
    
    #Motor
    worksheet.write(linha,18,table[43].text) # Instalação
    worksheet.write(linha,19,table[45].text) # Aspiração
    worksheet.write(linha,20,table[47].text) # Disposição
    worksheet.write(linha,21,table[49].text) # Alimentação
    worksheet.write(linha,22,table[51].text) # Cilindros
    worksheet.write(linha,23,table[53].text) # Comando de Válvulas
    worksheet.write(linha,24,table[55].text) # Tuchos
    worksheet.write(linha,25,table[57].text) # Variação do Comando
    worksheet.write(linha,26,table[59].text) # Válvulas por Cilindro
    worksheet.write(linha,27,table[61].text) # Diâmetro dos Cilindros
    worksheet.write(linha,28,table[63].text) # Razão de Compressão
    worksheet.write(linha,29,table[65].text) # Curso dos Pistões
    worksheet.write(linha,30,table[67].text) # Cilindrada
    worksheet.write(linha,31,table[69].text) # Pontência Máxima
    worksheet.write(linha,32,table[71].text) # Código do motor
    worksheet.write(linha,33,table[73].text) # Torque máximo
    worksheet.write(linha,34,table[75].text) # Peso/Potência
    worksheet.write(linha,35,table[77].text) # Torque Específico
    worksheet.write(linha,36,table[79].text) # Peso / Torque
    worksheet.write(linha,37,table[81].text) #Potência Especifica
    worksheet.write(linha,38,table[83].text) # Rotação Máxima  

    #Transmissão

    worksheet.write(linha,39,table[87].text) # Tração
    worksheet.write(linha,40,table[89].text) # Câmbio
    worksheet.write(linha,41,table[91].text) # Código do Câmbio
    worksheet.write(linha,42,table[93].text) # Acoplamento

    #Suspensão

    worksheet.write(linha,43,table[97].text) # Dianteira
    worksheet.write(linha,44,table[99].text) # Elemento Elástico
    worksheet.write(linha,45,table[101].text) # Traseira
    worksheet.write(linha,46,table[103].text) # Elemento Elástico

    #Freios

    worksheet.write(linha,47,table[107].text) # Disco Ventilado
    # worksheet.write(linha,48,table[109].text) # Traseiros

    #Direção

    worksheet.write(linha,49,table[113].text) # Assistência
    worksheet.write(linha,50,table[115].text) # Diâmetro mínimo de giro

    #Pneus

    worksheet.write(linha,51,table[119].text) # Dianteiros
    worksheet.write(linha,52,table[121].text) # Altura do Flanco
    worksheet.write(linha,53,table[123].text) # Traseiros
    worksheet.write(linha,54,table[125].text) # Altura do Flanco

    #Dimensões

    worksheet.write(linha,55,table[129].text) # Comprimento
    worksheet.write(linha,56,table[131].text) # Largura
    worksheet.write(linha,57,table[133].text) # Distância entre-eixos
    worksheet.write(linha,58,table[135].text) # Altura
    worksheet.write(linha,59,table[137].text) # Bitola Dianteira
    worksheet.write(linha,60,table[139].text) # Bitola Traseira
    worksheet.write(linha,61,table[141].text) # Porta-malas
    worksheet.write(linha,62,table[143].text) # Tanque de Combustível
    worksheet.write(linha,63,table[145].text) # Peso
    worksheet.write(linha,64,table[147].text) # Carga Útil
    worksheet.write(linha,65,table[149].text) # Vão livre do Solo
    
    #Aerodinâmica

    worksheet.write(linha,66,table[153].text) # Área Frontal (A)
    worksheet.write(linha,67,table[155].text) # Coeficiente de arrasto (Cx)
    worksheet.write(linha,68,table[157].text) # Área Frontal Corrigida

    #Desempenho

    worksheet.write(linha,69,table[161].text) # Velocidade Máxima
    worksheet.write(linha,70,table[163].text) # Aceleração 0-100 km/h

    #Consumo

    worksheet.write(linha,71,table[167].text) # Urbano
    worksheet.write(linha,72,table[171].text)
    worksheet.write(linha,73,table[169].text) # rodoviário
    worksheet.write(linha,74,table[173].text)
    
    #Autonomia

    worksheet.write(linha,75,table[177].text) # Urbana
    worksheet.write(linha,76,table[181].text)
    worksheet.write(linha,77,table[179].text) # Rodoviária
    worksheet.write(linha,78,table[183].text)

    #Segurança/Conforto/Infotenimento

    valor = {
        'seguranca'     : False,
        'conforto'      : False,
        'infotenimento' : False
    }

    seguranca = ''
    conforto = ''
    infotenimento = ''
     
    item = ''

    for i, tr in enumerate(driver.find_elements_by_xpath('//td[@colspan="6" and @align="right"]//table[@border="0" and @width="92%" ]//tbody//tr//td[@bgcolor="#ffffff"]')):
        if tr.text:
            if tr.text.strip() == 'Segurança':
                valor['seguranca'] = True
            elif tr.text.strip() == 'Conforto':
                valor['conforto'] = True
            elif tr.text.strip() == 'Infotenimento':
                valor['infotenimento'] = True 

            img = tr.find_element_by_xpath('//img[@title="Equipamento de série"]').get_attribute('src').replace('https://www.carrosnaweb.com.br/imgsite/','').replace('.gif','')

            if img == 'verde':
                item = 'Equipamento de Série'        
            else:
                item = 'Equipamento Opcional'

            if valor['seguranca'] == True and valor['conforto'] == False and valor['infotenimento'] == False and tr.text.strip() != u'Segurança':
                seguranca = '{} \n {} : {}'.format(seguranca.strip(),item,tr.text)
            elif valor['seguranca'] == True and valor['conforto'] == True and valor['infotenimento'] == False and tr.text.strip() != u'Conforto':
                conforto = '{} \n {} : {}'.format(conforto.strip(),item,tr.text)
            elif valor['seguranca'] == True and valor['conforto'] == True and valor['infotenimento'] == True and tr.text.strip() != u'Infotenimento' :
                infotenimento = '{} \n {} : {}'.format(infotenimento.strip(),item,tr.text)

    worksheet.write(linha,79,seguranca)
    worksheet.write(linha,80,conforto)
    worksheet.write(linha,81,infotenimento)

if __name__ == "__main__":
    main()
    # excel(2,'https://www.carrosnaweb.com.br/fichadetalhe.asp?codigo=12837')
    for i, link in  enumerate(links):
        print('Estamos no {} link de {} links '.format(i,len(links)))
        excel(i + 2,link) 

workbook.close()
