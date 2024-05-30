import datetime
import logging
import math as mt
import pandas as pd
import time
import platform
import os
import sys

lista_todos = []
lista_s1= []
lista_s2= []
lista_s3= []
lista_s4= []
lista_s5= []
lista_s6= []
lista_s7= []
lista_s8= []
lista_s9= []
lista_s10= []

materias_s1 = [["Algebra y trigonometria",3,1],["Calculo diferencial",3,1],["Geometria vectorial y analitica",3,1],["Vivamos la universidad",1,1],["Ingles I",1,1],["Lectoescritura",3,1],["Introducción a la ing. industrial",1,1]]
materias_s2 = [["Gestion de las organizaciones",3,2],["Habilidades gerenciales",3,2],["Algebra lineal",3,2],["Calculo integral",3,2],["Descubriendo la fisica",3,2],["Ingles II",1,2]]
materias_s3 = [["Gestión Contable",3,3],["Física Mecánica",3,3],["Inglés III",1,3],["Algoritmia y ProgramaciÓn",3,3],["Probabilidad e Inferencia EstadÍstica",3,3],["Teoria General de Sistemas",3,3]]
materias_s4 = [["Ingenieria economica",3,4],["Electiva en fisica",3,4],["Ingles IV",1,4],["Diseño de experimentos y analisis de regresion",3,4],["Optimizacion",3,4],["Gestion de metodos y tiempos",4,4]]
materias_s5 = [["Gestión Financiera",3,5],["Laboratorio Integrado de Física",1,5],["Inglés V",1,5],["Formación Ciudadana y Constitucional",1,5],["Dinámica de Sistemas",3,5],["Muestreo y Series de Tiempo",3,1],["Procesos Estocásticos y Análisis de Decisión",3,5],["Gestión por Procesos",3,5]]
materias_s6 = [["Gestión Tecnológica",3,6],["Legislación",3,6],["Electiva en Humanidades I",3,6],["Inglés VI",1,6],["Simulación Discreta",3,6],["Formulación de Proyectos de Investigación",3,6],["Normalización y Control de la Calidad",3,6]]
materias_s7 = [["Formulación y Evaluación de Proyectos de Inversión",3,7],["Emprendimiento",2,3],["Electiva en Humanidades II",3,7],["Énfasis Profesional I",3,7],["Electiva Complementaria I",3,7],["Diseño de Sistemas Productivos",3,7]]
materias_s8 = [["Gestion de proyectos",3,8],["Electiva en Humanidades III",3,8],["Enfasis profesional II",3,8],["Electiva complementaria II",3,8],["Administracion de la produccion y del servicio",3,8]]
materias_s9 = [["Electiva en Humanidades IV",3,9],["Énfasis Profesional III",3,9],["Electiva Complementaria III",3,9],["Gestión de la Cadena de Abastecimiento",3,9],["Ingeniería del Mejoramiento Continuo",3,9]]
materias_s10 = [["Practica profesional",12,10]]

codigosm_s1 = []
codigosm_s2 = []
codigosm_s3 = []
codigosm_s4 = []
codigosm_s5 = []
codigosm_s6 = []
codigosm_s7 = []
codigosm_s8 = []
codigosm_s9 = []
codigosm_s10 = []

HTD_s1 = []
HTD_s2 = []
HTD_s3 = []
HTD_s4 = []
HTD_s5 = []
HTD_s6 = []
HTD_s7 = []
HTD_s8 = []
HTD_s9 = []
HTD_s10 = []

HTI_s1 = []
HTI_s2 = []
HTI_s3 = []
HTI_s4 = []
HTI_s5 = []
HTI_s6 = []
HTI_s7 = []
HTI_s8 = []
HTI_s9 = []
HTI_s10 = []

limite_semestre = [30,30,30,25,25,25,20,20,20,10]

def generar_codigos(lista_grande):
    codigos = []
    for elemento in lista_grande:
        codigo = elemento[0][:4] + str(elemento[1]) + str(elemento[2])
        codigos.append(codigo)
    return codigos

def horasdetrabajodocente(materias):
    HTD = []
    for materia  in materias:
        if materia[1] == 4:
            HTD.append(96)
        elif materia[1] == 3:
            HTD.append(64)
        elif materia[1] == 2:
            HTD.append(32)
        elif materia[1] == 12:
          HTD.append(288)  
        else: 
            HTD.append(16)
    return HTD

def horasdetrabajoindependiente(materias):
    HTI = []
    for materia  in materias:
        if materia[1] == 4:
            HTI.append(120)
        elif materia[1] == 3:
            HTI.append(80)
        elif materia[1] == 2:
            HTI.append(64)
        elif materia[1] == 12:
            HTI.append(360)
        else: 
            HTI.append(32)
    return HTI

df = pd.read_excel('Estudiantes.xlsx')

lista_todos = df.values.tolist()

for estudiante in lista_todos:
    if estudiante[1] == 1:
       lista_s1.append(estudiante)
    elif estudiante[1] == 2:
       lista_s2.append(estudiante)
    elif estudiante[1] == 3:
        lista_s3.append(estudiante)
    elif estudiante[1] == 4:
        lista_s4.append(estudiante)
    elif estudiante[1] == 5:
        lista_s5.append(estudiante)
    elif estudiante[1] == 6:
       lista_s6.append(estudiante)
    elif estudiante[1] == 7:
        lista_s7.append(estudiante)
    elif estudiante[1] == 8:
        lista_s8.append(estudiante)
    elif estudiante[1] == 9:
        lista_s9.append(estudiante)
    else:
        lista_s10.append(estudiante)

def grupos(codigos, materias,TCA):
    lista_g = [[] for _ in range(len(materias))]
    for i in codigos:
        for j in range(TCA):
            if len(materias) == 8:
                if len(lista_g[0]) < TCA:
                    lista_g[0].append(i + str(j+1))
                elif len(lista_g[1]) < TCA:
                    lista_g[1].append(i + str(j+1))
                elif len(lista_g[2]) < TCA:
                    lista_g[2].append(i + str(j+1))
                elif len(lista_g[3]) < TCA:
                    lista_g[3].append(i + str(j+1))
                elif len(lista_g[4]) < TCA:
                    lista_g[4].append(i + str(j+1))
                elif len(lista_g[5]) < TCA:
                    lista_g[5].append(i + str(j+1))
                elif len(lista_g[6]) < TCA:
                    lista_g[6].append(i + str(j+1))
                else:
                    lista_g[7].append(i + str(j+1))
            elif len(materias) == 7:
                if len(lista_g[0]) < TCA:
                    lista_g[0].append(i + str(j+1))
                elif len(lista_g[1]) < TCA:
                    lista_g[1].append(i + str(j+1))
                elif len(lista_g[2]) < TCA:
                    lista_g[2].append(i + str(j+1))
                elif len(lista_g[3]) < TCA:
                    lista_g[3].append(i + str(j+1))
                elif len(lista_g[4]) < TCA:
                    lista_g[4].append(i + str(j+1))
                elif len(lista_g[5]) < TCA:
                    lista_g[5].append(i + str(j+1))
                else:
                    lista_g[6].append(i + str(j+1))
            elif len(materias) == 6:
                if len(lista_g[0]) < TCA:
                    lista_g[0].append(i + str(j+1))
                elif len(lista_g[1]) < TCA:
                    lista_g[1].append(i + str(j+1))
                elif len(lista_g[2]) < TCA:
                    lista_g[2].append(i + str(j+1))
                elif len(lista_g[3]) < TCA:
                    lista_g[3].append(i + str(j+1))
                elif len(lista_g[4]) < TCA:
                    lista_g[4].append(i + str(j+1))
                else:
                    lista_g[5].append(i + str(j+1))
            elif len(materias) == 5:
                if len(lista_g[0]) < TCA:
                    lista_g[0].append(i + str(j+1))
                elif len(lista_g[1]) < TCA:
                    lista_g[1].append(i + str(j+1))
                elif len(lista_g[2]) < TCA:
                    lista_g[2].append(i + str(j+1))
                elif len(lista_g[3]) < TCA:
                    lista_g[3].append(i + str(j+1))
                else:
                    lista_g[4].append(i + str(j+1))
            elif len(materias) == 4:
                if len(lista_g[0]) < TCA:
                    lista_g[0].append(i + str(j+1))
                elif len(lista_g[1]) < TCA:
                    lista_g[1].append(i + str(j+1))
                elif len(lista_g[2]) < TCA:
                    lista_g[2].append(i + str(j+1))
                else:
                    lista_g[3].append(i + str(j+1))
            elif len(materias) == 3:
                if len(lista_g[0]) < TCA:
                    lista_g[0].append(i + str(j+1))
                elif len(lista_g[1]) < TCA:
                    lista_g[1].append(i + str(j+1))
                else:
                    lista_g[2].append(i + str(j+1))
            elif len(materias) == 2:
                if len(lista_g[0]) < TCA:
                    lista_g[0].append(i + str(j+1))
                else:
                    lista_g[1].append(i + str(j+1))
            elif len(materias) == 1:
                if len(lista_g[0]) < TCA:
                    lista_g[0].append(i + str(j+1))
            
    return lista_g
 
def listas_estudiantes(materia, codigo, lista_de_codigos, TCA, lista_estudiantes, limite, j, HTD, HTI, NTE): 
    gruposqq = [[] for _ in range(TCA)]
    
    fecha = datetime.date.today().strftime('%Y-%m-%d')
    for x in range(TCA):      
        gruposqq[x].append(materia[0])
        gruposqq[x].append(materia[1])
        gruposqq[x].append(materia[2])
        gruposqq[x].append(codigo)
        gruposqq[x].append(HTD[j])
        gruposqq[x].append(HTI[j])
        gruposqq[x].append(NTE)
        gruposqq[x].append(lista_de_codigos[x]) 
        gruposqq[x].append(TCA)
        gruposqq[x].append(fecha)
    for e in lista_estudiantes:
        if len(gruposqq[0]) < limite+10:
            gruposqq[0].append(e[0:3:2])
        elif len(gruposqq[1]) < limite+10:
            gruposqq[1].append(e[0:3:2])
        elif len(gruposqq[2]) < limite+10:
            gruposqq[2].append(e[0:3:2])
        elif len(gruposqq[3]) < limite+10:
            gruposqq[3].append(e[0:3:2])
        elif len(gruposqq[4]) < limite+10:
            gruposqq[4].append(e[0:3:2])
        elif len(gruposqq[5]) < limite+10:
            gruposqq[5].append(e[0:3:2])
        else:
            gruposqq[6].append(e[0:3:2])
    return gruposqq

def exportar(lista, ruta):
    df = pd.DataFrame(lista)
    archivo_excel = ruta
    df.to_excel(archivo_excel, index=False, header=False)
    
def exportarcsv(lista, ruta):
    df = pd.DataFrame(lista)
    archivo_csv = ruta
    df.to_csv(archivo_csv, index=False, header=False)

def nombrearchivo(CA, NombreCurso, CantEst, CC):
    nombre = ''
    nombre += str(CA) + '-'
    NombreCurso = NombreCurso.upper()
    nombre += NombreCurso + '-'
    nombre += str(CantEst) + '-'
    nombre += str(CC)
    return nombre
    

def nombres(codigos, materias, nte, cc,n,c):    
    n1 = nombrearchivo(codigos[n], materias[n][0], nte, cc[0][c][-1])  
    return n1

def dcd(materia, codigo, TCA, HTD, HTI, NTE): 
    gruposqq = [[]]
    gruposqq[0] = ["Nombre de materia", "Creditos", "Semestre", "Código asignatura", "HTD", "HTI", "Total estudiantes", "Total cursos"]
    i = 0
    for ma in materia:
        gruposqq.append([])
        gruposqq[i+1].append(ma[0])
        gruposqq[i+1].append(ma[1])
        gruposqq[i+1].append(ma[2])
        gruposqq[i+1].append(codigo[i])
        gruposqq[i+1].append(HTD[i])
        gruposqq[i+1].append(HTI[i])
        gruposqq[i+1].append(NTE)
        gruposqq[i+1].append(TCA)  
        i += 1
    return gruposqq
  
cuevana = 1 
print(f'{"Inicio del proceso":>10}')
inicio = time.perf_counter() #Inicio contador de ejecucion
hoy = datetime.date.today().strftime('%Y%m%d') #Captura de fecha de ejecucion
nombre_archivo_log = f"log_{hoy}.log" # Inicializacion del log
#Configuracion de almacenamiento y niveles del log
logging.basicConfig(filename=nombre_archivo_log, level=logging.INFO,
                    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s')
DirectorioActual = os.getcwd()
textemp = f'El directorio actual de trabajo es: \n\t--> {DirectorioActual} \nEsta carpeta contendrá los archivos del trabajo final'
print(textemp)
logging.info(textemp)
logging.info(f"El programa fue ejecutado por {os.getlogin()} en el siguiente sistema operativo: {platform.platform()} y el la siguiente plataforma: {sys.version}")
logging.info(f"El {cuevana} procedimiento por SemesTrack fue: Crear encabezado")
tiempo_transcurrido = time.perf_counter() - inicio
logging.info(f"El procedimiento {cuevana} demoro: {tiempo_transcurrido} segundos")
cuevana += 1
inicio = time.perf_counter()

logging.info(f"Procedimiento {cuevana}: Creando codigos de materias")
codigosm_s1 = generar_codigos(materias_s1)
codigosm_s2 = generar_codigos(materias_s2)
codigosm_s3 = generar_codigos(materias_s3)
codigosm_s4 = generar_codigos(materias_s4)
codigosm_s5 = generar_codigos(materias_s5)
codigosm_s6 = generar_codigos(materias_s6)
codigosm_s7 = generar_codigos(materias_s7)
codigosm_s8 = generar_codigos(materias_s8)
codigosm_s9 = generar_codigos(materias_s9)
codigosm_s10 = generar_codigos(materias_s10)

tiempo_transcurrido = time.perf_counter() - inicio
logging.info(f"El procedimiento {cuevana} demoro: {tiempo_transcurrido} segundos")
cuevana += 1
inicio = time.perf_counter()

logging.info(f"Procedimiento {cuevana}: Creando horas de trabajo docente")
HTD_s1 = horasdetrabajodocente(materias_s1)
HTD_s2 = horasdetrabajodocente(materias_s2)
HTD_s3 = horasdetrabajodocente(materias_s3)
HTD_s4 = horasdetrabajodocente(materias_s4)
HTD_s5 = horasdetrabajodocente(materias_s5)
HTD_s6 = horasdetrabajodocente(materias_s6)
HTD_s7 = horasdetrabajodocente(materias_s7)
HTD_s8 = horasdetrabajodocente(materias_s8)
HTD_s9 = horasdetrabajodocente(materias_s9)
HTD_s10 = horasdetrabajodocente(materias_s10)

tiempo_transcurrido = time.perf_counter() - inicio
logging.info(f"El procedimiento {cuevana} demoro: {tiempo_transcurrido} segundos")
cuevana += 1
inicio = time.perf_counter()

logging.info(f"Procedimiento {cuevana}: Creando horas de trabajo independiente")
HTI_s1 = horasdetrabajoindependiente(materias_s1)
HTI_s2 = horasdetrabajoindependiente(materias_s2)
HTI_s3 = horasdetrabajoindependiente(materias_s3)
HTI_s4 = horasdetrabajoindependiente(materias_s4)
HTI_s5 = horasdetrabajoindependiente(materias_s5)
HTI_s6 = horasdetrabajoindependiente(materias_s6)
HTI_s7 = horasdetrabajoindependiente(materias_s7)
HTI_s8 = horasdetrabajoindependiente(materias_s8)
HTI_s9 = horasdetrabajoindependiente(materias_s9)
HTI_s10 = horasdetrabajoindependiente(materias_s10)

tiempo_transcurrido = time.perf_counter() - inicio
logging.info(f"El procedimiento {cuevana} demoro: {tiempo_transcurrido} segundos")
cuevana += 1
inicio = time.perf_counter()

logging.info(f"Procedimiento {cuevana}: Calculando la cantidad de estudiantes por semestre")
NTE_s1 = len(lista_s1)
NTE_s2 = len(lista_s2)
NTE_s3 = len(lista_s3)
NTE_s4 = len(lista_s4)
NTE_s5 = len(lista_s5)
NTE_s6 = len(lista_s6)
NTE_s7 = len(lista_s7)
NTE_s8 = len(lista_s8)
NTE_s9 = len(lista_s9)
NTE_s10 = len(lista_s10)

tiempo_transcurrido = time.perf_counter() - inicio
logging.info(f"El procedimiento {cuevana} demoro: {tiempo_transcurrido} segundos")
cuevana += 1
inicio = time.perf_counter()

logging.info(f"Procedimiento {cuevana}: Calculando el total de cursos asignados por semestre")
TCA_s1 = mt.ceil(NTE_s1/limite_semestre[0])
TCA_s2 = mt.ceil(NTE_s2/limite_semestre[1])
TCA_s3 = mt.ceil(NTE_s3/limite_semestre[2])
TCA_s4 = mt.ceil(NTE_s4/limite_semestre[3])
TCA_s5 = mt.ceil(NTE_s5/limite_semestre[4])
TCA_s6 = mt.ceil(NTE_s6/limite_semestre[5])
TCA_s7 = mt.ceil(NTE_s7/limite_semestre[6])
TCA_s8 = mt.ceil(NTE_s8/limite_semestre[7])
TCA_s9 = mt.ceil(NTE_s9/limite_semestre[8])
TCA_s10 = mt.ceil(NTE_s10/limite_semestre[9])

tiempo_transcurrido = time.perf_counter() - inicio
logging.info(f"El procedimiento {cuevana} demoro: {tiempo_transcurrido} segundos")
cuevana += 1
inicio = time.perf_counter()

logging.info(f"Procedimiento {cuevana}: Creando los codigos de curso por semestre")
CC_s1 = grupos(codigosm_s1, materias_s1, TCA_s1)
CC_s2 = grupos(codigosm_s2, materias_s2, TCA_s2)
CC_s3 = grupos(codigosm_s3, materias_s3, TCA_s3)
CC_s4 = grupos(codigosm_s4, materias_s4, TCA_s4)
CC_s5 = grupos(codigosm_s5, materias_s5, TCA_s5)
CC_s6 = grupos(codigosm_s6, materias_s6, TCA_s6)
CC_s7 = grupos(codigosm_s7, materias_s7, TCA_s7)
CC_s8 = grupos(codigosm_s8, materias_s8, TCA_s8)
CC_s9 = grupos(codigosm_s9, materias_s9, TCA_s9)
CC_s10 = grupos(codigosm_s10, materias_s10,  TCA_s10)

tiempo_transcurrido = time.perf_counter() - inicio
logging.info(f"El procedimiento {cuevana} demoro: {tiempo_transcurrido} segundos")
cuevana += 1
inicio = time.perf_counter()

logging.info(f"Procedimiento {cuevana}: Creando listas para cada materia por semestre")
#Semestre 1
listas_AyT = listas_estudiantes(materias_s1[0], codigosm_s1[0], CC_s1[0], TCA_s1, lista_s1, limite_semestre[0],0,HTD_s1,HTI_s1,NTE_s1)
listas_CD = listas_estudiantes(materias_s1[1], codigosm_s1[1], CC_s1[1], TCA_s1, lista_s1, limite_semestre[0],1,HTD_s1,HTI_s1,NTE_s1)
listas_GV = listas_estudiantes(materias_s1[2], codigosm_s1[2], CC_s1[2], TCA_s1, lista_s1, limite_semestre[0],2,HTD_s1,HTI_s1,NTE_s1)
listas_VU = listas_estudiantes(materias_s1[3], codigosm_s1[3], CC_s1[3], TCA_s1, lista_s1, limite_semestre[0],3,HTD_s1,HTI_s1,NTE_s1)
listas_Ins1 = listas_estudiantes(materias_s1[4], codigosm_s1[4], CC_s1[4], TCA_s1, lista_s1, limite_semestre[0],4,HTD_s1,HTI_s1,NTE_s1)
listas_LT = listas_estudiantes(materias_s1[5], codigosm_s1[5], CC_s1[5], TCA_s1, lista_s1, limite_semestre[0],5,HTD_s1,HTI_s1,NTE_s1)
listas_II = listas_estudiantes(materias_s1[6], codigosm_s1[6], CC_s1[6], TCA_s1, lista_s1, limite_semestre[0],6,HTD_s1,HTI_s1,NTE_s1)

#Semestre 2
listas_GO = listas_estudiantes(materias_s2[0], codigosm_s2[0], CC_s2[0], TCA_s2, lista_s2, limite_semestre[1],0,HTD_s2,HTI_s2,NTE_s2)
listas_HG = listas_estudiantes(materias_s2[1], codigosm_s2[1], CC_s2[1], TCA_s2, lista_s2, limite_semestre[1],1,HTD_s2,HTI_s2,NTE_s2)
listas_AL = listas_estudiantes(materias_s2[2], codigosm_s2[2], CC_s2[2], TCA_s2, lista_s2, limite_semestre[1],2,HTD_s2,HTI_s2,NTE_s2)
listas_CI = listas_estudiantes(materias_s2[3], codigosm_s2[3], CC_s2[3], TCA_s2, lista_s2, limite_semestre[1],3,HTD_s2,HTI_s2,NTE_s2)
listas_DF = listas_estudiantes(materias_s2[4], codigosm_s2[4], CC_s2[4], TCA_s2, lista_s2, limite_semestre[1],4,HTD_s2,HTI_s2,NTE_s2)
listas_Ins2 = listas_estudiantes(materias_s2[5], codigosm_s2[5], CC_s2[5], TCA_s2, lista_s2, limite_semestre[1],5,HTD_s2,HTI_s2,NTE_s2)

#Semestre 3
listas_GC = listas_estudiantes(materias_s3[0], codigosm_s3[0], CC_s3[0], TCA_s3, lista_s3, limite_semestre[2],0,HTD_s3,HTI_s3,NTE_s3)
listas_FM = listas_estudiantes(materias_s3[1], codigosm_s3[1], CC_s3[1], TCA_s3, lista_s3, limite_semestre[2],1,HTD_s3,HTI_s3,NTE_s3)
listas_InIII = listas_estudiantes(materias_s3[2], codigosm_s3[2], CC_s3[2], TCA_s3, lista_s3, limite_semestre[2],2,HTD_s3,HTI_s3,NTE_s3)
listas_AyP = listas_estudiantes(materias_s3[3], codigosm_s3[3], CC_s3[3], TCA_s3, lista_s3, limite_semestre[2],3,HTD_s3,HTI_s3,NTE_s3)
listas_PeIE = listas_estudiantes(materias_s3[4], codigosm_s3[4], CC_s3[4], TCA_s3, lista_s3, limite_semestre[2],4,HTD_s3,HTI_s3,NTE_s3)
listas_TGS = listas_estudiantes(materias_s3[5], codigosm_s3[5], CC_s3[5], TCA_s3, lista_s3, limite_semestre[2],5,HTD_s3,HTI_s3,NTE_s3)

#Semestre 4
listas_IE = listas_estudiantes(materias_s4[0], codigosm_s4[0], CC_s4[0], TCA_s4, lista_s4, limite_semestre[3],0,HTD_s4,HTI_s4,NTE_s4)
listas_EF = listas_estudiantes(materias_s4[1], codigosm_s4[1], CC_s4[1], TCA_s4, lista_s4, limite_semestre[3],1,HTD_s4,HTI_s4,NTE_s4)
listas_InIV = listas_estudiantes(materias_s4[2], codigosm_s4[2], CC_s4[2], TCA_s4, lista_s4, limite_semestre[3],2,HTD_s4,HTI_s4,NTE_s4)
listas_DEyAR = listas_estudiantes(materias_s4[3], codigosm_s4[3], CC_s4[3], TCA_s4, lista_s4, limite_semestre[3],3,HTD_s4,HTI_s4,NTE_s4)
listas_OP = listas_estudiantes(materias_s4[4], codigosm_s4[4], CC_s4[4], TCA_s4, lista_s4, limite_semestre[3],4,HTD_s4,HTI_s4,NTE_s4)
listas_GMyT = listas_estudiantes(materias_s4[5], codigosm_s4[5], CC_s4[5], TCA_s4, lista_s4, limite_semestre[3],5,HTD_s4,HTI_s4,NTE_s4)

#Semestre 5
listas_GF = listas_estudiantes(materias_s5[0], codigosm_s5[0], CC_s5[0], TCA_s5, lista_s5, limite_semestre[4],0,HTD_s5,HTI_s5,NTE_s5)
listas_LIF = listas_estudiantes(materias_s5[1], codigosm_s5[1], CC_s5[1], TCA_s5, lista_s5, limite_semestre[4],1,HTD_s5,HTI_s5,NTE_s5)
listas_InV = listas_estudiantes(materias_s5[2], codigosm_s5[2], CC_s5[2], TCA_s5, lista_s5, limite_semestre[4],2,HTD_s5,HTI_s5,NTE_s5)
listas_FCC = listas_estudiantes(materias_s5[3], codigosm_s5[3], CC_s5[3], TCA_s5, lista_s5, limite_semestre[4],3,HTD_s5,HTI_s5,NTE_s5)
listas_DS = listas_estudiantes(materias_s5[4], codigosm_s5[4], CC_s5[4], TCA_s5, lista_s5, limite_semestre[4],4,HTD_s5,HTI_s5,NTE_s5)
listas_MST = listas_estudiantes(materias_s5[5], codigosm_s5[5], CC_s5[5], TCA_s5, lista_s5, limite_semestre[4],5,HTD_s5,HTI_s5,NTE_s5)
listas_PEAD = listas_estudiantes(materias_s5[6], codigosm_s5[6], CC_s5[6], TCA_s5, lista_s5, limite_semestre[4],6,HTD_s5,HTI_s5,NTE_s5)
listas_GP = listas_estudiantes(materias_s5[7], codigosm_s5[7], CC_s5[7], TCA_s5, lista_s5, limite_semestre[4],7,HTD_s5,HTI_s5,NTE_s5)

#Semestre 6
listas_GT = listas_estudiantes(materias_s6[0], codigosm_s6[0], CC_s6[0], TCA_s6, lista_s6, limite_semestre[5],0,HTD_s6,HTI_s6,NTE_s6)
listas_L = listas_estudiantes(materias_s6[1], codigosm_s6[1], CC_s6[1], TCA_s6, lista_s6, limite_semestre[5],1,HTD_s6,HTI_s5,NTE_s6)
listas_EHI = listas_estudiantes(materias_s6[2], codigosm_s6[2], CC_s6[2], TCA_s6, lista_s6, limite_semestre[5],2,HTD_s6,HTI_s5,NTE_s6)
listas_IVI = listas_estudiantes(materias_s6[3], codigosm_s6[3], CC_s6[3], TCA_s6, lista_s6, limite_semestre[5],3,HTD_s6,HTI_s5,NTE_s6)
listas_SD = listas_estudiantes(materias_s6[4], codigosm_s6[4], CC_s6[4], TCA_s6, lista_s6, limite_semestre[5],4,HTD_s6,HTI_s5,NTE_s6)
listas_FPI = listas_estudiantes(materias_s6[5], codigosm_s6[5], CC_s6[5], TCA_s6, lista_s6, limite_semestre[5],5,HTD_s6,HTI_s5,NTE_s6)
listas_NCC = listas_estudiantes(materias_s6[6], codigosm_s6[6], CC_s6[6], TCA_s6, lista_s6, limite_semestre[5],6,HTD_s6,HTI_s5,NTE_s6)

#Semestre 7
listas_FEPI = listas_estudiantes(materias_s7[0], codigosm_s7[0], CC_s7[0], TCA_s7, lista_s7, limite_semestre[6],0,HTD_s7,HTI_s7,NTE_s7)
listas_E = listas_estudiantes(materias_s7[1], codigosm_s7[1], CC_s7[1], TCA_s7, lista_s7, limite_semestre[6],1,HTD_s7,HTI_s7,NTE_s7)
listas_EHII = listas_estudiantes(materias_s7[2], codigosm_s7[2], CC_s7[2], TCA_s7, lista_s7, limite_semestre[6],2,HTD_s7,HTI_s7,NTE_s7)
listas_EPI = listas_estudiantes(materias_s7[3], codigosm_s7[3], CC_s7[3], TCA_s7, lista_s7, limite_semestre[6],3,HTD_s7,HTI_s7,NTE_s7)
listas_ECI = listas_estudiantes(materias_s7[4], codigosm_s7[4], CC_s7[4], TCA_s7, lista_s7, limite_semestre[6],4,HTD_s7,HTI_s7,NTE_s7)
listas_DSP = listas_estudiantes(materias_s7[5], codigosm_s7[5], CC_s7[5], TCA_s7, lista_s7, limite_semestre[6],5,HTD_s7,HTI_s7,NTE_s7)

#Semestre 8
listas_GP = listas_estudiantes(materias_s8[0], codigosm_s8[0], CC_s8[0], TCA_s8, lista_s8, limite_semestre[7],0,HTD_s8,HTI_s8,NTE_s8)
listas_EHIII = listas_estudiantes(materias_s8[1], codigosm_s8[1], CC_s8[1], TCA_s8, lista_s8, limite_semestre[7],1,HTD_s8,HTI_s8,NTE_s8)
listas_EPII = listas_estudiantes(materias_s8[2], codigosm_s8[2], CC_s8[2], TCA_s8, lista_s8, limite_semestre[7],2,HTD_s8,HTI_s8,NTE_s8)
listas_ECII = listas_estudiantes(materias_s8[3], codigosm_s8[3], CC_s8[3], TCA_s8, lista_s8, limite_semestre[7],3,HTD_s8,HTI_s8,NTE_s8)
listas_APS = listas_estudiantes(materias_s8[4], codigosm_s8[4], CC_s8[4], TCA_s8, lista_s8, limite_semestre[7],4,HTD_s8,HTI_s8,NTE_s8)

#Semestre 9
listas_EHIV = listas_estudiantes(materias_s9[0], codigosm_s9[0], CC_s9[0], TCA_s9, lista_s9, limite_semestre[8],0,HTD_s9,HTI_s9,NTE_s9)
listas_EPIII = listas_estudiantes(materias_s9[1], codigosm_s9[1], CC_s9[1], TCA_s9, lista_s9, limite_semestre[8],1,HTD_s9,HTI_s9,NTE_s9)
listas_ECIII = listas_estudiantes(materias_s9[2], codigosm_s9[2], CC_s9[2], TCA_s9, lista_s9, limite_semestre[8],2,HTD_s9,HTI_s9,NTE_s9)
listas_GCA = listas_estudiantes(materias_s9[3], codigosm_s9[3], CC_s9[3], TCA_s9, lista_s9, limite_semestre[8],3,HTD_s9,HTI_s9,NTE_s9)
listas_IMC = listas_estudiantes(materias_s9[4], codigosm_s9[4], CC_s9[4], TCA_s9, lista_s9, limite_semestre[8],4,HTD_s9,HTI_s9,NTE_s9)

#Semestre 10
listas_PP = listas_estudiantes(materias_s10[0], codigosm_s10[0], CC_s10[0], TCA_s10, lista_s10, limite_semestre[9],0,HTD_s10,HTI_s10,NTE_s10)

tiempo_transcurrido = time.perf_counter() - inicio
logging.info(f"El procedimiento {cuevana} demoro: {tiempo_transcurrido} segundos")
cuevana += 1
inicio = time.perf_counter()

logging.info(f"Procedimiento {cuevana}: Exportando listas a excel y a su respectiva carpeta")
#Gracas a Deus:
for h in range(TCA_s1):
    exportar(listas_AyT[h], f'S1/Algebra y Trigonometria/Archivo Excel/{nombres(codigosm_s1,materias_s1,NTE_s1,CC_s1,0,h)}.xlsx')
    exportar(listas_CD[h], f'S1/Calculo Diferencial/Archivo Excel/{nombres(codigosm_s1,materias_s1,NTE_s1,CC_s1,1,h)}.xlsx')
    exportar(listas_GV[h], f'S1/Geometria Vectorial y Analitica/Archivo Excel/{nombres(codigosm_s1,materias_s1,NTE_s1,CC_s1,2,h)}.xlsx')
    exportar(listas_VU[h], f'S1/Vivamos la Universidad/Archivo Excel/{nombres(codigosm_s1,materias_s1,NTE_s1,CC_s1,3,h)}.xlsx')
    exportar(listas_Ins1[h], f'S1/Ingles I/Archivo Excel/{nombres(codigosm_s1,materias_s1,NTE_s1,CC_s1,4,h)}.xlsx')
    exportar(listas_LT[h], f'S1/Lectoescritura/Archivo Excel/{nombres(codigosm_s1,materias_s1,NTE_s1,CC_s1,5,h)}.xlsx')
    exportar(listas_II[h], f'S1/Introduccion a la Ingenieria Industrial/Archivo Excel/{nombres(codigosm_s1,materias_s1,NTE_s1,CC_s1,6,h)}.xlsx')

for h in range(TCA_s2):
    exportar(listas_GO[h], f'S2/Gestion de las Organizaciones/Archivo Excel/{nombres(codigosm_s2,materias_s2,NTE_s2,CC_s2,0,h)}.xlsx')
    exportar(listas_HG[h], f'S2/Habilidades Gerenciales/Archivo Excel/{nombres(codigosm_s2,materias_s2,NTE_s2,CC_s2,1,h)}.xlsx')
    exportar(listas_AL[h], f'S2/Algebra Lineal/Archivo Excel/{nombres(codigosm_s2,materias_s2,NTE_s2,CC_s2,2,h)}.xlsx')
    exportar(listas_CI[h], f'S2/Calculo Integral/Archivo Excel/{nombres(codigosm_s2,materias_s2,NTE_s2,CC_s2,3,h)}.xlsx')
    exportar(listas_DF[h], f'S2/Descubriendo la Fisica/Archivo Excel/{nombres(codigosm_s2,materias_s2,NTE_s2,CC_s2,4,h)}.xlsx')
    exportar(listas_Ins2[h], f'S2/Ingles II/Archivo Excel/{nombres(codigosm_s2,materias_s2,NTE_s2,CC_s2,5,h)}.xlsx')

for h in range(TCA_s3):
    exportar(listas_GC[h], f'S3/Gestion Contable/Archivo Excel/{nombres(codigosm_s3,materias_s3,NTE_s3,CC_s3,0,h)}.xlsx')
    exportar(listas_FM[h], f'S3/Fisica Mecanica/Archivo Excel/{nombres(codigosm_s3,materias_s3,NTE_s3,CC_s3,1,h)}.xlsx')
    exportar(listas_InIII[h], f'S3/Ingles III/Archivo Excel/{nombres(codigosm_s3,materias_s3,NTE_s3,CC_s3,2,h)}.xlsx')
    exportar(listas_AyP[h], f'S3/Algoritmia y Programacion/Archivo Excel/{nombres(codigosm_s3,materias_s3,NTE_s3,CC_s3,3,h)}.xlsx')
    exportar(listas_PeIE[h], f'S3/Probabilidad e Inferencia Estadistica/Archivo Excel/{nombres(codigosm_s3,materias_s3,NTE_s3,CC_s3,4,h)}.xlsx')
    exportar(listas_TGS[h], f'S3/Teoria General de Sistemas/Archivo Excel/{nombres(codigosm_s3,materias_s3,NTE_s3,CC_s3,5,h)}.xlsx')
    
for h in range(TCA_s4):
    exportar(listas_IE[h], f'S4/Ingenieria Economica/Archivo excel/{nombres(codigosm_s4,materias_s4,NTE_s4,CC_s4,0,h)}.xlsx')
    exportar(listas_EF[h], f'S4/Electiva en Fisica/Archivo excel/{nombres(codigosm_s4,materias_s4,NTE_s4,CC_s4,1,h)}.xlsx')
    exportar(listas_InIV[h], f'S4/Ingles IV/Archivo excel/{nombres(codigosm_s4,materias_s4,NTE_s4,CC_s4,2,h)}.xlsx')
    exportar(listas_DEyAR[h], f'S4/Diseño de Experimentos y Analisis de Regresion/Archivo excel/{nombres(codigosm_s4,materias_s4,NTE_s4,CC_s4,3,h)}.xlsx')
    exportar(listas_OP[h], f'S4/Optimizacion/Archivo excel/{nombres(codigosm_s4,materias_s4,NTE_s4,CC_s4,4,h)}.xlsx')
    exportar(listas_GMyT[h], f'S4/Gestion de Metodos y Tiempos/Archivo excel/{nombres(codigosm_s4,materias_s4,NTE_s4,CC_s4,5,h)}.xlsx')

for h in range(TCA_s5):
    exportar(listas_GF[h], f'S5/Gestion Financiera/Archivo Excel/{nombres(codigosm_s5,materias_s5,NTE_s5,CC_s5,0,h)}.xlsx')
    exportar(listas_LIF[h], f'S5/Laboratorio Integrado de Fisica/Archivo Excel/{nombres(codigosm_s5,materias_s5,NTE_s5,CC_s5,1,h)}.xlsx')
    exportar(listas_InIV[h], f'S5/Ingles V/Archivo Excel/{nombres(codigosm_s5,materias_s5,NTE_s5,CC_s5,2,h)}.xlsx')
    exportar(listas_FCC[h], f'S5/Formacion Ciudadana y Constitucional/Archivo Excel/{nombres(codigosm_s5,materias_s5,NTE_s5,CC_s5,3,h)}.xlsx')
    exportar(listas_DS[h], f'S5/Dinamica de Sistemas/Archivo Excel/{nombres(codigosm_s5,materias_s5,NTE_s5,CC_s5,4,h)}.xlsx')
    exportar(listas_MST[h], f'S5/Muestreo y Series de Tiempo/Archivo Excel/{nombres(codigosm_s5,materias_s5,NTE_s5,CC_s5,5,h)}.xlsx')
    exportar(listas_PEAD[h], f'S5/Procesos Estocasticos y Analisis De Decision/Archivo Excel/{nombres(codigosm_s5,materias_s5,NTE_s5,CC_s5,6,h)}.xlsx')
    exportar(listas_GP[h], f'S5/Gestion por Procesos/Archivo Excel/{nombres(codigosm_s5,materias_s5,NTE_s5,CC_s5,7,h)}.xlsx')

for h in range(TCA_s6):
    exportar(listas_GT[h], f'S6/Gestion Tecnologica/Archivo Excel/{nombres(codigosm_s6,materias_s6,NTE_s6,CC_s6,0,h)}.xlsx')
    exportar(listas_L[h], f'S6/Legislacion/Archivo Excel/{nombres(codigosm_s6,materias_s6,NTE_s6,CC_s6,1,h)}.xlsx')
    exportar(listas_EHI[h], f'S6/Electiva en Humanidades I/{nombres(codigosm_s6,materias_s6,NTE_s6,CC_s6,2,h)}.xlsx')
    exportar(listas_IVI[h], f'S6/Ingles VI/Archivo Excel/{nombres(codigosm_s6,materias_s6,NTE_s6,CC_s6,3,h)}.xlsx')
    exportar(listas_SD[h], f'S6/Simulacion Discreta/Archivo Excel/{nombres(codigosm_s6,materias_s6,NTE_s6,CC_s6,4,h)}.xlsx')
    exportar(listas_FPI[h], f'S6/Formulacion de Proyectos de Investigacion/Archivo Excel/{nombres(codigosm_s6,materias_s6,NTE_s6,CC_s6,5,h)}.xlsx')
    exportar(listas_NCC[h], f'S6/Normalizacion y Control de la Calidad/Archivo Excel/{nombres(codigosm_s6,materias_s6,NTE_s6,CC_s6,6,h)}.xlsx')

for h in range(TCA_s7):
    exportar(listas_FEPI[h], f'S7/Formulacion y Evaluacion de Proyectos De Inversion/Archivo Excel/{nombres(codigosm_s7,materias_s7,NTE_s7,CC_s7,0,h)}.xlsx')
    exportar(listas_E[h], f'S7/Emprendimiento/Archivo Excel/{nombres(codigosm_s7,materias_s7,NTE_s7,CC_s7,1,h)}.xlsx')
    exportar(listas_EHII[h], f'S7/Electiva en Humanidades II/{nombres(codigosm_s7,materias_s7,NTE_s7,CC_s7,2,h)}.xlsx')
    exportar(listas_EPI[h], f'S7/Enfasis Profesional I/Archivo Excel/{nombres(codigosm_s7,materias_s7,NTE_s7,CC_s7,3,h)}.xlsx')
    exportar(listas_ECI[h], f'S7/Electiva Complementaria I/Archivo Excel/{nombres(codigosm_s7,materias_s7,NTE_s7,CC_s7,4,h)}.xlsx')
    exportar(listas_DSP[h], f'S7/Diseño de Sistemas Productivos/Archivo Excel/{nombres(codigosm_s7,materias_s7,NTE_s7,CC_s7,5,h)}.xlsx')

for h in range(TCA_s8):
    exportar(listas_GP[h], f'S8/Gestion de Proyectos/Archivo Excel/{nombres(codigosm_s8,materias_s8,NTE_s8,CC_s8,0,h)}.xlsx')
    exportar(listas_EHIII[h], f'S8/Electiva en Humanidades III/Archivo Excel/{nombres(codigosm_s8,materias_s8,NTE_s8,CC_s8,1,h)}.xlsx')
    exportar(listas_EPII[h], f'S8/Enfasis Profesional II/Archivo Excel/{nombres(codigosm_s8,materias_s8,NTE_s8,CC_s8,2,h)}.xlsx')
    exportar(listas_ECII[h], f'S8/Electiva Complementaria II/Archivo Excel/{nombres(codigosm_s8,materias_s8,NTE_s8,CC_s8,3,h)}.xlsx')
    exportar(listas_APS[h], f'S8/Administracion de la Produccion y del Servicio/Archivo Excel/{nombres(codigosm_s8,materias_s8,NTE_s8,CC_s8,4,h)}.xlsx')

for h in range(TCA_s9):
    exportar(listas_EHIV[h], f'S9/Electiva en Humanidades IV/Archivo Excel/{nombres(codigosm_s9,materias_s9,NTE_s9,CC_s9,0,h)}.xlsx')
    exportar(listas_EPIII[h], f'S9/Enfasis Profesional III/Archivo Excel/{nombres(codigosm_s9,materias_s9,NTE_s9,CC_s9,1,h)}.xlsx')
    exportar(listas_ECIII[h], f'S9/Electiva Complementaria III/Archivo Excel/{nombres(codigosm_s9,materias_s9,NTE_s9,CC_s9,2,h)}.xlsx')
    exportar(listas_GCA[h], f'S9/Gestion de la Cadena de Abastecimiento/Archivo Excel/{nombres(codigosm_s9,materias_s9,NTE_s9,CC_s9,3,h)}.xlsx')
    exportar(listas_IMC[h], f'S9/Ingenieria del Mejoramiento Continuo/Archivo Excel/{nombres(codigosm_s9,materias_s9,NTE_s9,CC_s9,4,h)}.xlsx')

for h in range(TCA_s10):
    exportar(listas_PP[h], f'S10/Practica Profesional/Archivo Excel/{nombres(codigosm_s10,materias_s10,NTE_s10,CC_s10,0,h)}.xlsx')

tiempo_transcurrido = time.perf_counter() - inicio
logging.info(f"El procedimiento {cuevana} demoro: {tiempo_transcurrido} segundos")
cuevana += 1
inicio = time.perf_counter()

logging.info(f"Procedimiento {cuevana}: Exportando listas a csv y a su respectiva carpeta")
for h in range(TCA_s1):
    exportarcsv(listas_AyT[h], f'S1/Algebra y Trigonometria/Archivo CSV/{nombres(codigosm_s1,materias_s1,NTE_s1,CC_s1,0,h)}.csv')
    exportarcsv(listas_CD[h], f'S1/Calculo Diferencial/Archivo CSV/{nombres(codigosm_s1,materias_s1,NTE_s1,CC_s1,1,h)}.csv')
    exportarcsv(listas_GV[h], f'S1/Geometria Vectorial y Analitica/Archivo CSV/{nombres(codigosm_s1,materias_s1,NTE_s1,CC_s1,2,h)}.csv')
    exportarcsv(listas_VU[h], f'S1/Vivamos la Universidad/Archivo CSV/{nombres(codigosm_s1,materias_s1,NTE_s1,CC_s1,3,h)}.csv')
    exportarcsv(listas_Ins1[h], f'S1/Ingles I/Archivo CSV/{nombres(codigosm_s1,materias_s1,NTE_s1,CC_s1,4,h)}.csv')
    exportarcsv(listas_LT[h], f'S1/Lectoescritura/Archivo CSV/{nombres(codigosm_s1,materias_s1,NTE_s1,CC_s1,5,h)}.csv')
    exportarcsv(listas_II[h], f'S1/Introduccion a la Ingenieria Industrial/Archivo CSV/{nombres(codigosm_s1,materias_s1,NTE_s1,CC_s1,6,h)}.csv')

for h in range(TCA_s2):
    exportarcsv(listas_GO[h], f'S2/Gestion de las Organizaciones/Archivo CSV/{nombres(codigosm_s2,materias_s2,NTE_s2,CC_s2,0,h)}.csv')
    exportarcsv(listas_HG[h], f'S2/Habilidades Gerenciales/Archivo CSV/{nombres(codigosm_s2,materias_s2,NTE_s2,CC_s2,1,h)}.csv')
    exportarcsv(listas_AL[h], f'S2/Algebra Lineal/Archivo CSV/{nombres(codigosm_s2,materias_s2,NTE_s2,CC_s2,2,h)}.csv')
    exportarcsv(listas_CI[h], f'S2/Calculo Integral/Archivo CSV/{nombres(codigosm_s2,materias_s2,NTE_s2,CC_s2,3,h)}.csv')
    exportarcsv(listas_DF[h], f'S2/Descubriendo la Fisica/Archivo CSV/{nombres(codigosm_s2,materias_s2,NTE_s2,CC_s2,4,h)}.csv')
    exportarcsv(listas_Ins2[h], f'S2/Ingles II/Archivo CSV/{nombres(codigosm_s2,materias_s2,NTE_s2,CC_s2,5,h)}.csv')

for h in range(TCA_s3):
    exportarcsv(listas_GC[h], f'S3/Gestion Contable/Archivo CSV/{nombres(codigosm_s3,materias_s3,NTE_s3,CC_s3,0,h)}.csv')
    exportarcsv(listas_FM[h], f'S3/Fisica Mecanica/Archivo CSV/{nombres(codigosm_s3,materias_s3,NTE_s3,CC_s3,1,h)}.csv')
    exportarcsv(listas_InIII[h], f'S3/Ingles III/Archivo CSV/{nombres(codigosm_s3,materias_s3,NTE_s3,CC_s3,2,h)}.csv')
    exportarcsv(listas_AyP[h], f'S3/Algoritmia y Programacion/Archivo CSV/{nombres(codigosm_s3,materias_s3,NTE_s3,CC_s3,3,h)}.csv')
    exportarcsv(listas_PeIE[h], f'S3/Probabilidad e Inferencia Estadistica/Archivo CSV/{nombres(codigosm_s3,materias_s3,NTE_s3,CC_s3,4,h)}.csv')
    exportarcsv(listas_TGS[h], f'S3/Teoria General de Sistemas/Archivo CSV/{nombres(codigosm_s3,materias_s3,NTE_s3,CC_s3,5,h)}.csv')
    
for h in range(TCA_s4):
    exportarcsv(listas_IE[h], f'S4/Ingenieria Economica/Archivo CSV/{nombres(codigosm_s4,materias_s4,NTE_s4,CC_s4,0,h)}.csv')
    exportarcsv(listas_EF[h], f'S4/Electiva en Fisica/Archivo CSV/{nombres(codigosm_s4,materias_s4,NTE_s4,CC_s4,1,h)}.csv')
    exportarcsv(listas_InIV[h], f'S4/Ingles IV/Archivo CSV/{nombres(codigosm_s4,materias_s4,NTE_s4,CC_s4,2,h)}.csv')
    exportarcsv(listas_DEyAR[h], f'S4/Diseño de Experimentos y Analisis de Regresion/Archivo CSV/{nombres(codigosm_s4,materias_s4,NTE_s4,CC_s4,3,h)}.csv')
    exportarcsv(listas_OP[h], f'S4/Optimizacion/Archivo CSV/{nombres(codigosm_s4,materias_s4,NTE_s4,CC_s4,4,h)}.csv')
    exportarcsv(listas_GMyT[h], f'S4/Gestion de Metodos y Tiempos/Archivo CSV/{nombres(codigosm_s4,materias_s4,NTE_s4,CC_s4,5,h)}.csv')

for h in range(TCA_s5):
    exportarcsv(listas_GF[h], f'S5/Gestion Financiera/Archivo CSV/{nombres(codigosm_s5,materias_s5,NTE_s5,CC_s5,0,h)}.csv')
    exportarcsv(listas_LIF[h], f'S5/Laboratorio Integrado de Fisica/Archivo CSV/{nombres(codigosm_s5,materias_s5,NTE_s5,CC_s5,1,h)}.csv')
    exportarcsv(listas_InIV[h], f'S5/Ingles V/Archivo CSV/{nombres(codigosm_s5,materias_s5,NTE_s5,CC_s5,2,h)}.csv')
    exportarcsv(listas_FCC[h], f'S5/Formacion Ciudadana y Constitucional/Archivo CSV/{nombres(codigosm_s5,materias_s5,NTE_s5,CC_s5,3,h)}.csv')
    exportarcsv(listas_DS[h], f'S5/Dinamica de Sistemas/Archivo CSV/{nombres(codigosm_s5,materias_s5,NTE_s5,CC_s5,4,h)}.csv')
    exportarcsv(listas_MST[h], f'S5/Muestreo y Series de Tiempo/Archivo CSV/{nombres(codigosm_s5,materias_s5,NTE_s5,CC_s5,5,h)}.csv')
    exportarcsv(listas_PEAD[h], f'S5/Procesos Estocasticos y Analisis De Decision/Archivo CSV/{nombres(codigosm_s5,materias_s5,NTE_s5,CC_s5,6,h)}.csv')
    exportarcsv(listas_GP[h], f'S5/Gestion por Procesos/Archivo CSV/{nombres(codigosm_s5,materias_s5,NTE_s5,CC_s5,7,h)}.csv')

for h in range(TCA_s6):
    exportarcsv(listas_GT[h], f'S6/Gestion Tecnologica/Archivo CSV/{nombres(codigosm_s6,materias_s6,NTE_s6,CC_s6,0,h)}.csv')
    exportarcsv(listas_L[h], f'S6/Legislacion/Archivo CSV/{nombres(codigosm_s6,materias_s6,NTE_s6,CC_s6,1,h)}.csv')
    exportarcsv(listas_EHI[h], f'S6/Electiva en Humanidades I/{nombres(codigosm_s6,materias_s6,NTE_s6,CC_s6,2,h)}.csv')
    exportarcsv(listas_IVI[h], f'S6/Ingles VI/Archivo CSV/{nombres(codigosm_s6,materias_s6,NTE_s6,CC_s6,3,h)}.csv')
    exportarcsv(listas_SD[h], f'S6/Simulacion Discreta/Archivo CSV/{nombres(codigosm_s6,materias_s6,NTE_s6,CC_s6,4,h)}.csv')
    exportarcsv(listas_FPI[h], f'S6/Formulacion de Proyectos de Investigacion/Archivo CSV/{nombres(codigosm_s6,materias_s6,NTE_s6,CC_s6,5,h)}.csv')
    exportarcsv(listas_NCC[h], f'S6/Normalizacion y Control de la Calidad/Archivo CSV/{nombres(codigosm_s6,materias_s6,NTE_s6,CC_s6,6,h)}.csv')

for h in range(TCA_s7):
    exportarcsv(listas_FEPI[h], f'S7/Formulacion y Evaluacion de Proyectos De Inversion/Archivo CSV/{nombres(codigosm_s7,materias_s7,NTE_s7,CC_s7,0,h)}.csv')
    exportarcsv(listas_E[h], f'S7/Emprendimiento/Archivo CSV/{nombres(codigosm_s7,materias_s7,NTE_s7,CC_s7,1,h)}.csv')
    exportarcsv(listas_EHII[h], f'S7/Electiva en Humanidades II/{nombres(codigosm_s7,materias_s7,NTE_s7,CC_s7,2,h)}.csv')
    exportarcsv(listas_EPI[h], f'S7/Enfasis Profesional I/Archivo CSV/{nombres(codigosm_s7,materias_s7,NTE_s7,CC_s7,3,h)}.csv')
    exportarcsv(listas_ECI[h], f'S7/Electiva Complementaria I/Archivo CSV/{nombres(codigosm_s7,materias_s7,NTE_s7,CC_s7,4,h)}.csv')
    exportarcsv(listas_DSP[h], f'S7/Diseño de Sistemas Productivos/Archivo CSV/{nombres(codigosm_s7,materias_s7,NTE_s7,CC_s7,5,h)}.csv')

for h in range(TCA_s8):
    exportarcsv(listas_GP[h], f'S8/Gestion de Proyectos/Archivo CSV/{nombres(codigosm_s8,materias_s8,NTE_s8,CC_s8,0,h)}.csv')
    exportarcsv(listas_EHIII[h], f'S8/Electiva en Humanidades III/Archivo CSV/{nombres(codigosm_s8,materias_s8,NTE_s8,CC_s8,1,h)}.csv')
    exportarcsv(listas_EPII[h], f'S8/Enfasis Profesional II/Archivo CSV/{nombres(codigosm_s8,materias_s8,NTE_s8,CC_s8,2,h)}.csv')
    exportarcsv(listas_ECII[h], f'S8/Electiva Complementaria II/Archivo CSV/{nombres(codigosm_s8,materias_s8,NTE_s8,CC_s8,3,h)}.csv')
    exportarcsv(listas_APS[h], f'S8/Administracion de la Produccion y del Servicio/Archivo CSV/{nombres(codigosm_s8,materias_s8,NTE_s8,CC_s8,4,h)}.csv')

for h in range(TCA_s9):
    exportarcsv(listas_EHIV[h], f'S9/Electiva en Humanidades IV/Archivo CSV/{nombres(codigosm_s9,materias_s9,NTE_s9,CC_s9,0,h)}.csv')
    exportarcsv(listas_EPIII[h], f'S9/Enfasis Profesional III/Archivo CSV/{nombres(codigosm_s9,materias_s9,NTE_s9,CC_s9,1,h)}.csv')
    exportarcsv(listas_ECIII[h], f'S9/Electiva Complementaria III/Archivo CSV/{nombres(codigosm_s9,materias_s9,NTE_s9,CC_s9,2,h)}.csv')
    exportarcsv(listas_GCA[h], f'S9/Gestion de la Cadena de Abastecimiento/Archivo CSV/{nombres(codigosm_s9,materias_s9,NTE_s9,CC_s9,3,h)}.csv')
    exportarcsv(listas_IMC[h], f'S9/Ingenieria del Mejoramiento Continuo/Archivo CSV/{nombres(codigosm_s9,materias_s9,NTE_s9,CC_s9,4,h)}.csv')

for h in range(TCA_s10):
    exportarcsv(listas_PP[h], f'S10/Practica Profesional/Archivo CSV/{nombres(codigosm_s10,materias_s10,NTE_s10,CC_s10,0,h)}.csv')

tiempo_transcurrido = time.perf_counter() - inicio
logging.info(f"El procedimiento {cuevana} demoro: {tiempo_transcurrido} segundos")
cuevana += 1
inicio = time.perf_counter()

logging.info(f"Procedimiento {cuevana}: Exportando documento docente")
exportar(dcd(materias_s1, codigosm_s1, TCA_s1, HTD_s1, HTI_s1, NTE_s1), "Documentos Docente/Documento-Docente-S1.xlsx")

exportar(dcd(materias_s2, codigosm_s2, TCA_s2, HTD_s2, HTI_s2, NTE_s2), "Documentos Docente/Documento-Docente-S2.xlsx")

exportar(dcd(materias_s3, codigosm_s3, TCA_s3, HTD_s3, HTI_s3, NTE_s3), "Documentos Docente/Documento-Docente-S3.xlsx")

exportar(dcd(materias_s4, codigosm_s4, TCA_s4, HTD_s4, HTI_s4, NTE_s4), "Documentos Docente/Documento-Docente-S4.xlsx")

exportar(dcd(materias_s5, codigosm_s5, TCA_s5, HTD_s5, HTI_s5, NTE_s5), "Documentos Docente/Documento-Docente-S5.xlsx")

exportar(dcd(materias_s6, codigosm_s6, TCA_s6, HTD_s6, HTI_s6, NTE_s6), "Documentos Docente/Documento-Docente-S6.xlsx")

exportar(dcd(materias_s7, codigosm_s7, TCA_s7, HTD_s7, HTI_s7, NTE_s7), "Documentos Docente/Documento-Docente-S7.xlsx")

exportar(dcd(materias_s8, codigosm_s8, TCA_s8, HTD_s8, HTI_s8, NTE_s8), "Documentos Docente/Documento-Docente-S8.xlsx")

exportar(dcd(materias_s9, codigosm_s9, TCA_s9, HTD_s9, HTI_s9, NTE_s9), "Documentos Docente/Documento-Docente-S9.xlsx")

exportar(dcd(materias_s10, codigosm_s10, TCA_s10, HTD_s10, HTI_s10, NTE_s10), "Documentos Docente/Documento-Docente-S10.xlsx")
cuevana += 1

logging.info(f"Procedimiento {cuevana}: Finalizado el proceso, tuvo {cuevana} procedimientos")

tiempo_transcurrido = time.perf_counter() - inicio
logging.info(f"El procedimiento {cuevana} demoro: {tiempo_transcurrido} segundos")
cuevana = 1