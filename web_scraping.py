import scrapy                                                                  
from scrapy.crawler import CrawlerProcess
import urllib.parse as urlparse
import builtwith
import re
from urllib.request import Request, urlopen
import urllib.robotparser as robotparser
import time
import lxml.html
import urllib
import pprint
import logging
import pandas as pd



#Se crea una clase como parte del proceso de la generación del spider
class prevac_scrapping(scrapy.Spider):
    #Se da un nombre al spide generado
    name = 'prevac_scrapping'
    #Se ingresa la dirección de los URLs a los que se les va a aplicar el Scraping
    start_urls = ['http://bisigsa.fac.mil.co/DBxtra.NET/LogIn.aspx',]
    #Esta función es la inicial, en la cual se realice el log_in, para ingresar a la plataforma
    def parse(self, response):
    #Se determinan los datos de acceso a la plataforma. Los nombres de los campos fueron usados verificando
    #la pagina web.
        data = {'UserEmail':'DESOP',
                'UserPass':'DESOP'}
    #Se genera un request de Scrapy, espeficiando el id de la forma que contiene el ingreso., se pasa tambien
    #como argumento los datos y el valor del boton que envía los datos. Finalmente se usa el callback
    #que basicamente envia la respuesta a la función parce_tasks.
        yield scrapy.FormRequest.from_response(response,formxpath='//*[@id="form1"]', formdata = data,
                                                clickdata={"value":"Ingresar"}, callback= self.parce_tasks)
    #Una vez se ingresa a la plataforma, se genera otro Request, redireccionado a una web, la cual contiene
    #la información que queremos descargar. Posterior dicha respuesta es enviada a la función parce_task_cacom5
    def parce_tasks(self, response):
        url = "http://bisigsa.fac.mil.co/DBxtra.NET/DataGrid.aspx?ID=1675&Parameters=true&ShowValues=110000000&Param1=2020&Param2='CACOM-5'&Param3=&Param4=PREVENCION&Param5=&Param6=&Param7=&Param8=&Param9="
        yield scrapy.Request(url = url, callback= self.parce_task_cacom5) 
    #Finalmente se recibe la respuesta, e inspeccionando la pagina web, se identifica la clase de la tabla que contiene
    #la información.
    def parce_task_cacom5(self, response):
        tables = response.xpath('//tr[@class="dxgvDataRow"]')
    #Se crea una variable lista, la cual va a adjuntar la información por columna. Se crea una lista ya que es reconocida
    #por pandas al momento de crear el dataframe.
        info = []
        for table in tables:
            info.append(table.xpath('.//text()').getall())
    #Se crea una lista con el nombre de todas las columnas requeridas para el procesamiento de los datos. Las columnas Y y X son
    #generadas por información adicional generada por la web.
        lista = ['X','ID','SEQUENCE_NUM','ACT_NAME','INICIO_PLAN','FIN_PLAN','TYPE_NAME','UNIDAD','PRIORIDAD',
                'RESPONSABLE_ACTIVIDAD','ESTADO_ACTIVIDAD','VALOR_ACTIVIDAD','PROGRESO_ACTIVIDAD','PORCENTAJE_ACTIVIDAD',
                'PROGRAMA_PREVENCION','PLAN_ID','PLAN_NAME','PLAN_YEAR','PLAN_TYPE_NAME','PLAN_STATUS','PROGRAM_NAME','PRIORIDAD_ACTIVIDAD',
                'ASOCIADO_A','PROGRAMA','Y']
    #Se crea un dataframde de nombre tasks que integra la información de todas las columnas.
        tasks = pd.DataFrame(info, columns = lista) 
    #Se implementa el metodo to_csv con el fin de pasar la información del dataframe a un archivo csv.
        tasks.to_csv(r'C:\Users\User\Google Drive\CACOM5_DESOP_PREVAC\Actividades_PREVAC\Tareas\2020\REFERENCIA_PLAN_ACCION.csv')
    #Finalmente se crea un yield para imprimir la información descargada.
        yield {
            'info':info,
            'tables':tables,
            
        }

    


process = CrawlerProcess()
process.crawl(prevac_scrapping)
process.start(stop_after_crawl=False)