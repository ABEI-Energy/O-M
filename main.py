import datetime as dt
import io
import locale as lc
import pandas as pd
import streamlit as st
from functions import *
import streamlit_toggle as tog
from docx.shared import Cm
from zipfile import ZipFile
import matplotlib.pyplot as plt

if 'disable_opt' not in st.session_state:
    st.session_state.disable_opt = False

#Set the language for datetime
lc.setlocale(lc.LC_ALL,'es_ES.UTF-8')
month = dt.datetime.now().strftime("%B %Y")
time_doc = dt.datetime.now().strftime("%y%m")

class pict:
    def __init__(self, name, file):
        self.name = name
        self.file = file

def normalize(string):
    return str(round(float(string.replace(",", ".")),2))

def normalize2(string):
    return str(string.replace(",", "."))

st.set_page_config(layout="wide")

if 'finalCheck' not in st.session_state:
    st.session_state['finalCheck'] = None
    st.session_state['calculo_pr'] = None
    st.session_state['calculo_disponibilidad'] = None
    st.session_state['calculo_plantilla'] = None
    st.session_state['flag_T1_aux'] = None
    st.session_state['flag_T1_aux2'] = None
    st.session_state['create_document'] = None
    st.session_state['flagZip'] = None
    st.session_state['picsDone'] = None
    st.session_state['tablesDone'] = None
    st.session_state['wordsDone'] = None
    st.session_state['documentDone'] = None
    st.session_state['tablerDone'] = None
    st.session_state['days'] = 0
    st.session_state['month'] = None
    st.session_state['generarDocumento'] = None

    st.session_state['accum_PR1'] = None
    st.session_state['avail_accum_1'] = None
    st.session_state['unavEnergLoss'] = None

    st.session_state.visibility = 'visible'


    # st.session_state['plantilla'] = False
    # st.session_state['disponibilidad'] = False
    # st.session_state['pr'] = False


"""
# O&M doc maker
"""

# Time to read the documents

st.divider()

coly, colx = st.columns(2)
with colx:
    uploadedFiles = st.file_uploader("Drag here the .xlsx files (Calculo disponibilidad, Plantilla informe cliente y Calculo PR).", accept_multiple_files = True)

with coly:

    st.caption("Example files")
    with open("Resources/O&M report files.zip", "rb") as fp:
        btn = st.download_button(
            label="Download example files",
            data=fp,
            file_name="O&M report files.zip",
            mime="application/zip"
        )

    accum_PR1 = st.text_input("PR acumulado (Año 1)", label_visibility=st.session_state.visibility, key = 'accum_PR1')
    avail_accum_1 = st.text_input("Disponibilidad acumulada (Año 1)", label_visibility=st.session_state.visibility, key = 'avail_accum_1')
    unavEnergLoss = st.text_input("Energía perdida estimada por indisponibilidad", label_visibility=st.session_state.visibility, key = 'unavEnergLoss')


colz, cola = st.columns(2)

if uploadedFiles and st.session_state.accum_PR1 and st.session_state.avail_accum_1 and st.session_state.unavEnergLoss:

    # Ordenamos la entrada de los archivos en el bucle, porque si no está df

    uploadedFilesOrd = [None]*3
    
    for uploadedFile in uploadedFiles:
        if 'plantilla' in uploadedFile.name.lower():
            uploadedFilesOrd[0] = uploadedFile
        elif 'disponibilidad' in uploadedFile.name.lower():
            uploadedFilesOrd[1] = uploadedFile
        elif 'pr' in uploadedFile.name.lower():
            uploadedFilesOrd[2] = uploadedFile

    uploadedFiles = uploadedFilesOrd

    for uploadedFile in uploadedFiles:

        if uploadedFile is None:
            next
        elif ((uploadedFile.name.endswith('xlsm')) or (uploadedFile.name.endswith('xlsx'))):
        
            if ('plantilla' in uploadedFile.name.lower()):#& (st.session_state.plantilla == False)
                st.session_state['calculo_plantilla'] = True
                excelPlantilla = uploadedFile
                df_flagT1_aux, df_flagT2, df_flagT3, df_flagT4, dict_portada = excel_reader(excelPlantilla)

            elif ('disponibilidad' in uploadedFile.name.lower()) & (st.session_state.flag_T1_aux):#& (st.session_state.disponibilidad == False)
                st.session_state['calculo_disponibilidad'] = True
                excelDisponibilidad = uploadedFile
                df_flagT1_aux = excel_reader(excelDisponibilidad, df_flagT1_aux)

            elif ('pr' in uploadedFile.name.lower()) & (st.session_state.flag_T1_aux2): # & (st.session_state.pr == False)
                st.session_state['calculo_pr'] = True
                excelPR = uploadedFile
                df_flagT1 = excel_reader(excelPR, df_flagT1_aux)
    else:
        next

    # Edit and round(2) the float numbers

    df_flagT1.iloc[:,1:] = df_flagT1.iloc[:,1:].astype(float).round(2)
    df_flagT2.iloc[:,1:] = df_flagT2.iloc[:,1:].astype(float).round(2)
    df_flagT3.iloc[:,1:] = df_flagT3.iloc[:,1:].astype(float).round(2)
    df_flagT4.iloc[:,1:] = df_flagT4.iloc[:,1:].astype(float).round(2)

    dict_portada.update({'accumPR1':accum_PR1})
    dict_portada.update({'monthAcc1':avail_accum_1})
    dict_portada.update({'unavEnergLoss':unavEnergLoss})

    for elem in dict_portada:
        if type(dict_portada[elem]) != str:
            dict_portada[elem] = round(dict_portada[elem],3)

if st.session_state.tablesDone and uploadedFiles:

    picDict = {}

    # SET Production

    fig_SET_prod, ax = plt.subplots(figsize=(15, 6))
    ax.plot(df_flagT1['Fecha'].iloc[:-1],df_flagT1['SET'].iloc[:-1])
    ax.set_xlabel('Día')
    ax.set_ylabel('Producción SET (kWh)')
    ax.set_title('Producción SET (kWh)')
    
    plt.xticks(df_flagT1['Fecha'].iloc[:-1].astype(float))

    fig_io_SET_prod = io.BytesIO()
    fig_SET_prod.savefig(fig_io_SET_prod, format = 'png')
    fig_io_SET_prod.seek(0)
    
    picDict[pict('SetProduction', fig_io_SET_prod).name] = pict(uploadedFile.name, fig_io_SET_prod).file


    # CTS Production

    fig_CTS_prod, ax = plt.subplots(figsize=(15, 6))
    ax.plot(df_flagT2.iloc[:-1])
    ax.set_xlabel('Día')
    ax.set_ylabel('Producción CTs (kWh)')
    ax.set_title('Producción CTs (kWh)')
    plt.xticks(df_flagT1['Fecha'].iloc[:-1].astype(float))
    ax.legend(df_flagT2.columns[1:])

    fig_io_CTS_prod = io.BytesIO()
    fig_CTS_prod.savefig(fig_io_CTS_prod, format = 'png')
    fig_io_CTS_prod.seek(0)
    
    picDict[pict('CTSProduction', fig_io_CTS_prod).name] = pict(uploadedFile.name, fig_io_CTS_prod).file


    # Coplanar & Horizontal radiation

    fig_Coplanar_Horiz_rad, ax = plt.subplots(figsize=(15, 6))
    df_T3_aux = df_flagT3.iloc[:-1,[0,5,6]]
    ax.plot(df_T3_aux['Fecha'],df_T3_aux.iloc[:,1:])
    ax.set_xlabel('Día')
    ax.set_ylabel('Radiación Coplanar y Horizontal (Wh/(m$^2$))')
    ax.set_title('Radiación Coplanar y Horizontal (Wh/(m$^2$))')
    plt.xticks(df_T3_aux['Fecha'].astype(float))
    ax.legend(df_T3_aux.columns[1:])

    fig_io_Coplanar_Horiz_rad = io.BytesIO()
    fig_Coplanar_Horiz_rad.savefig(fig_io_Coplanar_Horiz_rad, format = 'png')
    fig_io_Coplanar_Horiz_rad.seek(0)
    
    picDict[pict('CopHorizRadiation', fig_io_Coplanar_Horiz_rad).name] = pict(uploadedFile.name, fig_io_Coplanar_Horiz_rad).file   


    # Temperatures

    fig_Temperatures, ax = plt.subplots(figsize=(15, 6))
    # df_T4_aux = pd.concat([df_flagT3['Fecha'],df_flagT4], axis = 1) #ya está cortada para que no coja la última
    df_T4_aux = df_flagT4 #ya está cortada para que no coja la última
    df_T4_aux.reset_index(inplace = True, drop = True)
    df_T4_aux = df_T4_aux.iloc[:-1]
    ax.plot(df_T4_aux['Fecha'],df_T4_aux.iloc[:,1:])
    ax.set_xlabel('Día')
    ax.set_ylabel('Temperatura (°C)')
    ax.set_title('Temperatura (°C)')
    plt.xticks(df_T4_aux['Fecha'].astype(float))
    ax.legend(df_T4_aux.columns[1:])

    fig_io_Temperatures = io.BytesIO()
    fig_Temperatures.savefig(fig_io_Temperatures, format = 'png')
    fig_io_Temperatures.seek(0)
    
    picDict[pict('Temperatures', fig_io_Temperatures).name] = pict(uploadedFile.name, fig_io_Temperatures).file   

    
    # PR

    fig_PR, ax = plt.subplots(figsize=(15, 6))
    ax.plot(df_flagT1['Fecha'].iloc[:-1], df_flagT1.iloc[:-1,-1])
    ax.set_xlabel('Día')
    ax.set_ylabel(f'PR {st.session_state.month}')
    ax.set_title(f'PR {st.session_state.month}')
    plt.xticks(df_flagT1['Fecha'].iloc[:-1].astype(float))

    fig_io_PR = io.BytesIO()
    fig_PR.savefig(fig_io_PR, format = 'png')
    fig_io_PR.seek(0)
    
    picDict[pict('PR', fig_io_PR).name] = pict(uploadedFile.name, fig_io_PR).file


    # Availability

    fig_Availability, ax = plt.subplots(figsize=(15, 6))
    ax.plot(df_flagT1['Fecha'].iloc[:-1], df_flagT1.iloc[:-1,-2])
    ax.set_xlabel('Día')
    ax.set_ylabel(f'Disponibilidad {st.session_state.month}')
    ax.set_title(f'Disponibilidad {st.session_state.month}')
    plt.xticks(df_flagT1['Fecha'].iloc[:-1].astype(float))

    fig_io_Availability = io.BytesIO()
    fig_Availability.savefig(fig_io_Availability, format = 'png')
    fig_io_Availability.seek(0)
    
    picDict[pict('Availability', fig_io_Availability).name] = pict(uploadedFile.name, fig_io_Availability).file

    # We got everything ready to implement in the document
    st.session_state.create_document = True

if st.session_state.create_document and uploadedFiles:
    
    st.button("Generar documento", key = 'generarDocumento')
    doc_file = duplicateDoc()

if st.session_state.generarDocumento:

    with st.status("Preparando archivo", expanded=True) as status:

        nameWord = time_doc + ' - Informe Cliente'+ " PSFV Cartuja" + ".doc"
        st.write("Preparando gráficas")

        insert_image_in_cell(doc_file, picDict)

        st.write('Preparando tablas')
        docWriter(doc_file, dict_portada)
        docTabler(doc_file, df_flagT1, df_flagT2, df_flagT3, df_flagT4)
        doc_modelo_bio = io.BytesIO()

        st.write("Guardando archivo")
        doc_file.save(doc_modelo_bio)
        doc_modelo_bio.seek(0)
        if st.session_state.picsDone and st.session_state.wordsDone and st.session_state.tablesDone and st.session_state.tablerDone: 
            st.session_state.documentDone = True
            st.session_state.picsDone = False
            st.session_state.wordsDone = False
            st.session_state.tablesDone = False
            st.session_state.tablerDone = False

        status.update(label="Archivo completado")
  
if st.session_state.documentDone:
    btn = st.download_button(
            label="Descarga archivos",
            data=doc_modelo_bio,
            file_name=nameWord,
            mime="application/docx"
        )
    st.session_state.documentDone = False
