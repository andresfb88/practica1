import smtplib
import pandas as pd
from datetime import datetime
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
import xlwt
from web_scraping 


data = r'C:\Users\User\Google Drive\CACOM5_DESOP_PREVAC\Actividades_PREVAC\Tareas\2020\REPORTE_PLAN_ACCION.xls'

users = {
    'Juan Camilo Agudelo Restrepo': 'juan.agudelor@fac.mil.co',
    'Abdon Estibenson Uribe Taborda': '',
    'Ricardo Augusto Ortiz Ramírez': '',
    'Andres Felipe Bello Zapata': 'andres.bello@fac.mil.co',
    'Victor Alfonso Lopez Salguero': 'victor.lopez@fac.mil.co',
    'Javier Jimenez Gaona': 'javier.jimenezg@fac.mil.co'
}

def create_list():
    plan_act_info = pd.read_excel(data, sheet_name = 0)
    date =  datetime.strptime(plan_act_info['INICIO_PLAN'][0], '%Y-%m-%d')
    pend_days = []
    for i in range(plan_act_info.shape[0]):
        date =  datetime.strptime(plan_act_info['FIN_PLAN'][i], '%Y-%m-%d')
        pend_days.append((date - datetime.today()).days)
    plan_act_info['DIAS_PEND'] = pend_days
    plan_act_info.to_excel(r'C:\Users\User\Google Drive\CACOM5_DESOP_PREVAC\Actividades_PREVAC\Tareas\2020\REPORTE_PLAN_ACCION.xls')
    return plan_act_info

def list_compar():
    ref = r'C:\Users\User\Google Drive\CACOM5_DESOP_PREVAC\Actividades_PREVAC\Tareas\2020\REFERENCIA_PLAN_ACCION.xls'
    elimi = []
    new = []
    info = {}
    act = pd.read_excel(data, sheet_name = 0)
    ref = pd.read_excel(ref, sheet_name = 0)
    act_act = list(set(act['SEQUENCE_NUM'].tolist()))
    act_ref = list(set(ref['SEQUENCE_NUM'].tolist()))
    for refe in act_ref:
        if refe not in act_act:
            elimi.append(refe)
    for actu in act_act:
        if actu not in act_ref:
            new.append(actu)
    if(len(new)>0):
        act = act.set_index('SEQUENCE_NUM')
        act = act.loc[new]
        act = act.reset_index('SEQUENCE_NUM')
        info['new'] = [act]
        people = set(act['RESPONSABLE_ACTIVIDAD'].tolist())
        for person in people:
            person_tasks = act[act.RESPONSABLE_ACTIVIDAD == str(person)]
            info['new'].append([person,person_tasks])

    if(len(elimi)>0):
        ref = ref.set_index('SEQUENCE_NUM')
        ref = ref.loc[elimi]
        ref = ref.reset_index()
        info['elimi'] = [ref]
        people = set(ref['RESPONSABLE_ACTIVIDAD'].tolist())
        for person in people:
            person_tasks = ref[ref.RESPONSABLE_ACTIVIDAD == str(person)]
            info['elimi'].append([person,person_tasks])

    if(len(info) > 0):
        return info
    else:
        return "none"

def report():
    plan_act_info = create_list()
    plan_act_info = plan_act_info[plan_act_info.ESTADO_ACTIVIDAD != 'Cumplida']
    date =  datetime.strptime(plan_act_info['INICIO_PLAN'][0], '%Y-%m-%d')
    list_report = {}
    list_30, list_15, list_8, list_5, list_4, list_3, list_2, list_1, pend_days = [],[],[],[],[],[],[],[],[]
    for i in range(plan_act_info.shape[0]):
        date =  datetime.strptime(plan_act_info['FIN_PLAN'][i], '%Y-%m-%d')
        pend_days.append(abs(datetime.today() - date).days)
        if((date - datetime.today()).days == 30):
            list_30.append(i)
        if((date - datetime.today()).days == 15): 
            list_15.append(i)
        if((date - datetime.today()).days == 8):
            list_8.append(i)
        if((date - datetime.today()).days == 5):
            list_5.append(i)
        if((date - datetime.today()).days == 4):
            list_4.append(i)
        if((date - datetime.today()).days == 3):
            list_3.append(i)
        if((date - datetime.today()).days == 2):
            list_2.append(i)
        if((date - datetime.today()).days == 1):
            list_1.append(i)

    if (len(list_30) > 0):
        list_report[30] = [plan_act_info.loc[list_30]]
    if (len(list_15) > 0):
        list_report[15] = [plan_act_info.loc[list_15]]
    if (len(list_8) > 0):
        list_report[8] = [plan_act_info.loc[list_8]]
    if (len(list_5) > 0):
        list_report[5] = [plan_act_info.loc[list_5]]
    if (len(list_4) > 0):
        list_report[4] = [plan_act_info.loc[list_4]]
    if (len(list_3) > 0):
        list_report[3] = [ plan_act_info.loc[list_3]]
    if (len(list_2) > 0):
        list_report[2] = [plan_act_info.loc[list_2]]
    if (len(list_1) > 0):
        list_report[1] = [plan_act_info.loc[list_1]]

    for key in list_report.keys():
        table = list_report[key][0]
        respon = set(table['RESPONSABLE_ACTIVIDAD'].tolist())
        for person in respon:
            person_tasks = table[table.RESPONSABLE_ACTIVIDAD == str(person)]
            list_report[key].append([person,person_tasks])
    if (len(list_report) > 0):
        return list_report
    else:
        return 'taskno' 

def send_mail(lista_data):

    if(lista_data[0] == ''):
        dataframe = lista_data[1]
        dataframe = dataframe[['SEQUENCE_NUM','RESPONSABLE_ACTIVIDAD','DIAS_PEND','FIN_PLAN','ACT_NAME']]
        dataframe.set_index('SEQUENCE_NUM')
        with smtplib.SMTP('smtp.gmail.com', 587) as smtp:
            smtp.ehlo()
            smtp.starttls()
            smtp.ehlo()
            smtp.login('andresfb88@gmail.com', 'altos6-401')
            msg = MIMEMultipart()
            html = """
                <html>
                    <h3>Buen dia</h3>
                    <body>
                    <h3>Las tareas pendientes a cumplir dentro de los siguientes {0} dias son:
                    <br>
                    <br>
                    {1}
                    <br>
                    <br>
                    <h3>Adjunto al presente correo encontrará el listado general de las actividades si requiere mayor información</h3>
                    </body>
                </html>
            """.format(lista_data[2],dataframe.to_html())
            table = MIMEText(html, 'html')
            att = MIMEApplication(open(data, 'rb').read())
            msg.attach(att)
            msg.attach(table)
            att = MIMEApplication(open(data, 'rb').read())
            filename = 'Informacion General.xls'
            att.add_header('Content-Disposition', "attachment; filename= %s" % filename)
            msg.attach(att)
            msg['Subject'] = 'Tareas pendientes {0} dias'.format(lista_data[2])
            # smtp.sendmail('andresfb88@gmail.com',['juan.agudelor@fac.mil.co','victor.lopez@fac.mil.co','andres.bello@fac.mil.co'], msg.as_string())       

    elif(lista_data[0] == 'new_gen'):
        dataframe = lista_data[1]
        dataframe = dataframe[['SEQUENCE_NUM','RESPONSABLE_ACTIVIDAD','DIAS_PEND','FIN_PLAN']]
        dataframe.set_index('SEQUENCE_NUM')
        with smtplib.SMTP('smtp.gmail.com', 587) as smtp:
            smtp.ehlo()
            smtp.starttls()
            smtp.ehlo()
            smtp.login('andresfb88@gmail.com', 'altos6-401')
            msg = MIMEMultipart()
            html = """
                <html>
                    <h3>Buen dia</h3>
                    <body>
                    <h3>Se ha(n) creado la(s) siguiente(s) tarea(s):
                    <br>
                    <br>
                    {0}
                    <br>
                    <h3>Adjunto al presente correo encontrará el listado general de las actividades si requiere mayor información</h3>
                    </body>
                </html>
            """.format(dataframe.to_html())
            table = MIMEText(html, 'html')
            msg.attach(table)
            att = MIMEApplication(open(data, 'rb').read())
            filename = 'Informacion General.xls'
            att.add_header('Content-Disposition', "attachment; filename= %s" % filename)
            msg.attach(att)
            msg['Subject'] = 'Nueva(s) actividades(s)'
            # smtp.sendmail('andresfb88@gmail.com',['juan.agudelor@fac.mil.co','victor.lopez@fac.mil.co','andres.bello@fac.mil.co'], msg.as_string())    

    elif(lista_data[0] == 'elim_gen'):
        dataframe = lista_data[1]
        dataframe = dataframe[['SEQUENCE_NUM','RESPONSABLE_ACTIVIDAD','DIAS_PEND','FIN_PLAN']]
        dataframe.set_index('SEQUENCE_NUM')
        with smtplib.SMTP('smtp.gmail.com', 587) as smtp:
            smtp.ehlo()
            smtp.starttls()
            smtp.ehlo()
            smtp.login('andresfb88@gmail.com', 'altos6-401')
            msg = MIMEMultipart()
            html = """
                <html>
                    <h3>Buen dia</h3>
                    <body>
                    <h3>Se ha(n) eliminado la(s) siguiente(s) tarea(s):
                    <br>
                    <br>
                    {0}
                    <br>
                    </body>
                </html>
            """.format(dataframe.to_html())
            table = MIMEText(html, 'html')
            msg.attach(table)
            msg['Subject'] = 'Actividades(s) Eliminada(s)'
            # smtp.sendmail('andresfb88@gmail.com',['juan.agudelor@fac.mil.co','victor.lopez@fac.mil.co','andres.bello@fac.mil.co'], msg.as_string())    

    elif(lista_data[0] == 'new_esp'):
        dataframe = lista_data[2]
        dataframe = dataframe[['SEQUENCE_NUM','DIAS_PEND','FIN_PLAN','DIAS_PEND','ACT_NAME']]
        dataframe.set_index('SEQUENCE_NUM')
        name = lista_data[1].split(' ')
        with smtplib.SMTP('smtp.gmail.com', 587) as smtp:
            smtp.ehlo()
            smtp.starttls()
            smtp.ehlo()
            smtp.login('andresfb88@gmail.com', 'altos6-401')
            msg = MIMEMultipart()
            html = """
                <html>
                    <h3>Buen dia {0} {1} </h3>
                    <body>
                    <h3>Se ha(n) creado la(s) siguiente(s) tarea(s):</h3>
                    <br>
                    <br>
                    {2}
                    <br>
                    <br>
                    <h3>Adjunto al presente correo encontrará el listado general de las actividades si requiere mayor información</h3>
                    </body>
                </html>
            """.format(name[0], name[1],dataframe.to_html())
            table = MIMEText(html, 'html')
            msg.attach(table)
            att = MIMEApplication(open(data, 'rb').read())
            filename = 'Informacion General.xls'
            att.add_header('Content-Disposition', "attachment; filename= %s" % filename)
            msg.attach(att)
            msg['Subject'] = 'Nueva(s) actividades(s)'
            # smtp.sendmail('andresfb88@gmail.com',users[str(lista_data[1])], msg.as_string())       

    elif(lista_data[0] == 'elim_esp'):
        dataframe = lista_data[2]
        dataframe = dataframe[['SEQUENCE_NUM','DIAS_PEND','FIN_PLAN','DIAS_PEND','ACT_NAME']]
        dataframe.set_index('SEQUENCE_NUM')
        name = lista_data[1].split(' ')
        with smtplib.SMTP('smtp.gmail.com', 587) as smtp:
            smtp.ehlo()
            smtp.starttls()
            smtp.ehlo()
            smtp.login('andresfb88@gmail.com', 'altos6-401')
            msg = MIMEMultipart()
            html = """
                <html>
                    <h3>Buen dia {0} {1} </h3>
                    <body>
                    <h3>Se ha(n) creado la(s) siguiente(s) tarea(s):</h3>
                    <br>
                    <br>
                    {2}
                    <br>
                    <br>
                    <h3>Adjunto al presente correo encontrará el listado general de las actividades si requiere mayor información</h3>
                    </body>
                </html>
            """.format(name[0], name[1],dataframe.to_html())
            table = MIMEText(html, 'html')
            msg.attach(table)
            msg['Subject'] = 'Actividades(s) Eliminada(s)'
            # smtp.sendmail('andresfb88@gmail.com',users[str(lista_data[1])], msg.as_string())       

    elif(lista_data[0] == 'tare_esp'):
        dataframe = lista_data[2]
        dataframe = dataframe[['SEQUENCE_NUM','DIAS_PEND','FIN_PLAN','DIAS_PEND','ACT_NAME']]
        dataframe.set_index('SEQUENCE_NUM')
        name = lista_data[1].split(' ')
        with smtplib.SMTP('smtp.gmail.com', 587) as smtp:
            smtp.ehlo()
            smtp.starttls()
            smtp.ehlo()
            smtp.login('andresfb88@gmail.com', 'altos6-401')
            msg = MIMEMultipart()
            html = """
                <html>
                    <h3>Buen dia {0} {1} </h3>
                    <body>
                    <h3>Las tareas pendientes a cumplir dentro de los siguientes {2} dias son:</h3>
                    <br>
                    <br>
                    {3}
                    <br>
                    <br>
                    <h3>Adjunto al presente correo encontrará el listado general de las actividades si requiere mayor información</h3>
                    </body>
                </html>
            """.format(name[0], name[1], lista_data[3],dataframe.to_html())
            table = MIMEText(html, 'html')
            msg.attach(table)
            att = MIMEApplication(open(data, 'rb').read())
            filename = 'Informacion General.xls'
            att.add_header('Content-Disposition', "attachment; filename= %s" % filename)
            msg.attach(att)
            msg['Subject'] = 'Tareas pendientes {0} dias'.format(lista_data[3])
            # smtp.sendmail('andresfb88@gmail.com',users[str(lista_data[1])], msg.as_string())       

def execute():
    web_scraping()
    info = report()
    if (info != 'taskno'):
        for key in info.keys():
            send_mail(['',info[key][0],key])
            for item in range(len(info[key])-1):
                send_mail(['tare_esp',info[key][item+1][0],info[key][item+1][1],key])

    new_elim = list_compar()
    if(new_elim != 'none'):
        for key in new_elim.keys():
            if(key == 'new'):
                send_mail(['new_gen',new_elim[key][0],key])
                for item in range(len(new_elim[key])-1):
                    send_mail(['new_esp',new_elim[key][item+1][0],new_elim[key][item+1][1]])
            else:
                send_mail(['elim_gen',new_elim[key][0],key])
                for item in range(len(new_elim[key])-1):
                    send_mail(['elim_esp',new_elim[key][item+1][0],new_elim[key][item+1][1]])

if __name__ == "__main__":
    execute()
    
 


                

        
         

