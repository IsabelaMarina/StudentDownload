#%% Selenium
#import selenium
#import numpy as np
import pandas as pd
import openpyxl #Tablas de Excel
import xlsxwriter
import time     #Delays
from selenium import webdriver
from selenium.webdriver.common.keys import Keys

#%% Info de Selenium
#https://towardsdatascience.com/controlling-the-web-with-python-6fceb22c5f08
#https://selenium-python.readthedocs.io/navigating.html
#https://www.geeksforgeeks.org/interacting-with-webpage-selenium-python/

#Drivers (Edge):
#https://developer.microsoft.com/en-us/microsoft-edge/tools/webdriver/

#Esto para los dropdowns
#https://www.toolsqa.com/selenium-webdriver/dropdown-in-selenium/

#%% Esta función quita los números del inicio de cada línea
#Los números pueden ser de uno o dos dígitos, así que este man busca hasta el primer espacio y lo remueve
def firstNumber(line):
    fnn = ""
    for i in line:
        fnn = fnn + i
        if (i.isspace()):
            break
    return line.removeprefix(fnn)


#%% Obtener el Código de la clase y Materia
def obtieneNombre(actualNombre):
    p = actualNombre
    p1 = p.split()
    if (p1[4]=='ELECTRICIDAD'):
        p2 = 'Elec'
    elif (p1[4]=='MECÁNICA'):
        p2 = 'Meca'
    else:
        p2 = p1[4]

    codeCourse = p2 + " " + p1[2].removesuffix(':')
    return(codeCourse)

    
#%% Función de creación de las tablas
def obtieneTablas(nombres):
    # %%Crear tabla a partir de lo importado
    kt = nombres.splitlines()   #Toda la tabla queda en un solo string, pero tiene \n caracteres
                                #splitlines divide usando esos caracteres
    k = kt[3:]  #Los primeros tres strings son los títulos y no los necesito

    #Inicializar vectores (no sé si python necesita esto, lo hago porsiaca)
    names = []
    lastnames = []
    codes = []

    for linea in k:
        lmod = firstNumber(linea)       #Esta función quita los números del inicio de cada línea
        ln = lmod.find("T000") - 1      #Encuentra la posición del código
        lc = lmod.find(", ")            #Poisición de la coma que separa nombre de apellido

        lname = lmod[lc+2:ln]           #Selecciona nombre (aparece después del apellido)
        lfin = lmod[:lc]                #Selecciona apellido
        names.append(lname)             #Colecciona en el vector respectivo
        lastnames.append(lfin)
        #print(lname)
        #print(lfin)

        lcode = lmod[ln+1:ln+10]        #Selecciona código T000
        codes.append(lcode)
        #print(lcode)

    #Creación de diccionario y DataFrame a partir de diccionario
    #tlista = {"Grupo":list[0,len(k)-1],"Nombre":names,"Apellido":lastnames,"Código":codes}
    tlista = {"Nombre":names,"Apellido":lastnames,"Código":codes}
    lista = pd.DataFrame(tlista)
    #print(lista)


    #%%
    #Agregar las carreras

    ings = []

    for i in range(0,len(lista)):
    #i = 3
        print(i)
        stud = lista.Apellido.loc[i] + ", " + lista.Nombre.loc[i]
        stu = driver.find_element(by="link text",value=stud)        #ESTO A VECES NO DA CLIC
        stu.click()
        #driver.execute_script('$("a:contains('+stud+')").click()')  #VOY POR AQUÍ: No le dio clic :(


        InfoAl = driver.find_element(by="link text",value='Información de Alumno')
        InfoAl.click()

        #Obtener la info
        carri = driver.find_element_by_xpath('//*[@id="contentHolder"]/div[3]/table[3]/tbody/tr[2]/td')
        carr = carri.text.title()
        #print(carr)
        ings.append(carr)
        driver.execute_script("window.history.go(-2)")  #Regresa (botón "atrás") dos páginas

    #print(ings)

    #%%
    lista2 = lista
    #carreras = {"Carrera":ings}
    lista3 = lista2.assign(Carrera = ings)  #Agrega la columna de carreras respectivas
    print(lista3)
    return(lista3)

# %% Abrir el navegador (Abre Edge estable)
driver = webdriver.Edge('msedgedriver_last.exe')
driver.get('https://www.utb.edu.co/mi-utb/')

# %% Abrir página de Banner (abre nueva pestaña)
banner = driver.find_element(by="link text",value='Banner')
banner.click()

#%% Cambia a la nueva pestaña, y abre "Servicios a Docentes"
driver.switch_to.window(driver.window_handles[-1])
docentesbutton = driver.find_element_by_id("bmenu--P_FacMainMnu___UID2")
docentesbutton.click()
# %% Abre "Resumen Lista de Clase"
# Al correrlo completo (no por celdas) esa hp cosa no encuentra el vergajo botón
#además, hay que poner un caso para seleccionar cada NRC por aparte
time.sleep(3)
listabutton = driver.find_element_by_xpath('/html/body/div[19]/div[3]/div[2]/div[2]/div/div[8]')
listabutton.click()
#/html/body/div[19]/div[3]/div[2]/div[2]/div/div[8]


#%% Selecciona el periodo (siempre aparece la primera vez que se abre Banner)
#También obtiene una lista de los elementos del selector, y la selección se hace por coincidencia con el texto (de esa manera se pueden seleccionar otros luego, o seleccionarlo aunque cambie de posición)

p2 = driver.find_element_by_xpath('//*[@id="contentHolder"]/div[2]/form/table/tbody/tr[1]/td[2]/select')
opciones = p2.find_elements_by_css_selector("option")
ops = []
for i in range(1,len(opciones)):
    ops.append(opciones[i].text)
#print(ops)
#ind = ops.index('PRIMER PERIODO 2021 PREGRADO')+2
ind = ops.index('SEGUNDO PERIODO 2021 PREGRADO')+2  #HARDCODEADO PARA ESTE SEMESTRE

periodo = driver.find_element_by_xpath('//*[@id="contentHolder"]/div[2]/form/table/tbody/tr[1]/td[2]/select/option['+str(ind)+']')
periodo.click()


# %% Se da clic a "Enviar" (y va a la siguiente página)
#Debería ir en la celda anterior, pero la dejaré aparte por ahora
enviarperiodo = driver.find_element_by_id("id____UID5")
enviarperiodo.click()

# %% Selección de NRC's
#Después le puedo poner un for bacano para que vaya a todos los NRC's
p3 = driver.find_element_by_xpath('//*[@id="contentHolder"]/div[2]/form/table/tbody/tr[1]/td[2]/select')
opciones3 = p3.find_elements_by_css_selector("option")
alltables = []      #Aquí se guardarán TODAS LAS LISTAS
allcursos = []
ops = []
for i in range(0,len(opciones3)):
    ops.append(opciones3[i].text)
print(ops)

#%% EL MAXIFOR
#Este GIGANTESCO FOR hace la lectura de las tablas y la lectura de las carreras de cada muchacho
for j in range(0,len(ops)):
    #ind = ops.index('CBAS F02A H1: FÍSICA MECÁNICA, 1423 (0)')+2
    ind = j+1

    #Selección de lista desplegable
    clase = driver.find_element_by_xpath('//*[@id="contentHolder"]/div[2]/form/table/tbody/tr[1]/td[2]/select/option['+str(ind)+']')
    clase.click()

    #Botón
    enviarNRC = driver.find_element_by_id("id____UID5")  #Quién hps creó está página
    enviarNRC.click()

    #writer = pd.ExcelWriter("ListaPrueba2.xlsx",engine='xlsxwriter')   #Esto no funcionó

    try:
        #Se intenta este bloque

        # Selección de la tabla de nombres de estudiantes
        # La mía es la tercera tabla (ese es el [2] del final)
        print("Intentando:",ops[j])
        tablaEst = driver.find_elements_by_class_name("datadisplaytable")[2]
        nombres = tablaEst.text
    except:
        #Este bloque se ejecuta si se da una excepción
        print("No encontré nada en",ops[j])
    else:
        #Este bloque se ejecuta si no hay excepciones
        print("Sí hay estudiantes en",ops[j])
        tabla = obtieneTablas(nombres)
        curso = obtieneNombre(ops[j])
        #tabla.to_excel(writer,sheet_name=curso,index=False)
        alltables.append(tabla) #Guarda todas las tablas en una lista para guardarlas a Excel luego
        allcursos.append(curso) #Nombres para usarlos para las hojas del Excel
    finally:
        #Este bloque siempre se ejecuta
        print("Fin de búsqueda")
        #print("Hola, siempre me ejecutaré\nAdemás, en este bloque está la instrucción para ir atrás")
        driver.execute_script("window.history.go(-1)")
    
#writer.save()       #No escribe a Excel :(

#%% Guardado a Excel

#Este fue el método que me funcionó para guardar todas las tablas por pestañas en el mismo archivo
with pd.ExcelWriter('ListaEstudiantes.xlsx') as writer1:
    for k in range(0,len(alltables)):
        alltables[k].to_excel(writer1, sheet_name=allcursos[k], index=False)




# %%
