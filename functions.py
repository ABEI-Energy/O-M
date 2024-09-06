'''

This script contains the functions to be called from streamlit.py

'''
import pandas as pd
import numpy as np
import locale as lc
import datetime as dt
import streamlit as st
import requests
import utm
from docx import Document
from PIL import Image
from docx.shared import Cm
import io

lc.setlocale(lc.LC_ALL,'en_US.UTF-8')

month = dt.datetime.now().strftime("%B %Y").capitalize()
year = month.split(' ')[-1]


def excel_reader(directory, df_aux = None):

    if st.session_state.calculo_plantilla:

        xls = pd.ExcelFile(directory)

        hoja_portada = pd.read_excel(xls, 'Portada', index_col=None, header = None)
        hoja_tablas_plantilla = pd.read_excel(xls, 'Tablas', index_col=None, header = None)


        # Cogemos sólo los datos de portada que nos interesan

        df_portada = pd.DataFrame(hoja_portada[263:])
        df_portada.dropna(axis = 1, how = 'all', inplace = True)
        df_portada.dropna(axis = 0, how = 'all', inplace = True)
        df_portada.dropna(axis = 1, thresh = 10, inplace = True)

        df_portada = df_portada.iloc[:,0:-1]
        df_portada.dropna(axis = 0, how = 'any', inplace = True)
        df_portada.reset_index(inplace = True, drop = True)
        df_portada.columns = pd.RangeIndex(len(df_portada.columns))

        dict_portada = {}

        dict_portada['degTable'] = df_portada.iloc[3,1]*100
        dict_portada['mtProdTable'] = df_portada.iloc[4,1]
        dict_portada['btProdTable'] = df_portada.iloc[5,1]
        dict_portada['mtAggTable'] = df_portada.iloc[6,1]
        dict_portada['copRadTable'] = df_portada.iloc[7,1]
        dict_portada['mtAggcopFilRadTable'] = df_portada.iloc[8,1]
        dict_portada['horRadTable'] = df_portada.iloc[9,1]
        dict_portada['monthPRTable'] = df_portada.iloc[10,1]
        dict_portada['monthAvailTable '] = df_portada.iloc[13,1] 

        # Get the tables.

        df_portada = hoja_tablas_plantilla.apply(lambda x: x.str.strip())
        end_tables = df_portada.loc[df_portada.iloc[:,0]=='TOTAL'].index #Empezar, la primera siempre en el mismo lado, y la segunda el primer index +2 
        df_portada = hoja_tablas_plantilla

        # We put in df_flagT1 the data for the first table, and we will build there.
        df_flagT1_aux = df_portada

        # Bloque 1
        df_flagT1_aux = df_portada

        df_flagT1_aux = df_flagT1_aux.iloc[:,:3]
        df_flagT1_aux = df_flagT1_aux.iloc[3:(end_tables[0]+1)]
        df_flagT1_aux.reset_index(inplace = True, drop = True)
        df_flagT1_aux.columns = ['Fecha', 'SET', 'Total Invs']
        df_flagT1_aux.name = 'PR.'
        
        st.session_state.flag_T1_aux = True

        # Bloque 2 va a flagT2
        df_flagT2 = df_portada
        cols_ct = df_flagT2.iloc[1].dropna()
        lenCT = len(df_flagT2.iloc[1].dropna())

        df_flagT2 = df_flagT2.iloc[:,3:3+lenCT]
        df_flagT2 = df_flagT2.iloc[3:(end_tables[0]+1)]
        df_flagT2.reset_index(inplace = True, drop = True)
        df_flagT2.columns = cols_ct
        df_flagT2.name = 'CTs.'


        # flagT3 inversores
        df_flagT3 = df_portada
        cols_inv = df_flagT3.iloc[(end_tables[0]+2),:8]

        df_flagT3 = df_flagT3.iloc[:,:8] # Columnas
        df_flagT3 = df_flagT3.iloc[(end_tables[0]+6):(end_tables[1])] # Filas
        df_flagT3.reset_index(inplace = True, drop = True)
        df_flagT3.columns = cols_inv
        df_flagT3.name = 'Irrad.'


        # flagT4 temperatura
        df_flagT4 = df_portada
        cols_temp = df_flagT4.iloc[(end_tables[0]+2),9:]

        df_flagT4 = df_flagT4.iloc[:,9:] # Columnas
        df_flagT4 = df_flagT4.iloc[(end_tables[0]+6):(end_tables[1])] # Filas
        df_flagT4.reset_index(inplace = True, drop = True)
        df_flagT4.columns = cols_temp
        df_flagT4.name = 'Temp.'

        st.session_state.calculo_plantilla = False

        return df_flagT1_aux, df_flagT2, df_flagT3, df_flagT4, dict_portada
        
    elif st.session_state.calculo_disponibilidad:

        df_flagT1_aux = df_aux

        xls = pd.ExcelFile(directory)

        hoja_calculo_disp = pd.read_excel(xls, 'Cálculo Disp. 24 (corr)', index_col=None, header = None)

        # Ver cuántos días tiene el mes, esa columna varía en longitud.
        end_tables = hoja_calculo_disp.loc[hoja_calculo_disp.iloc[:,42]=='TOTAL'].index[0] #Empezar, la primera siempre en el mismo lado, y la segunda el primer index +2 
        df_calculo_disp = hoja_calculo_disp.iloc[52:end_tables + 1,42:44].reset_index(drop = True)
        df_calculo_disp.columns = ['Fecha', 'Availability']

        df_flagT1_aux = df_flagT1_aux.merge(df_calculo_disp,on = 'Fecha')
        df_flagT1_aux.iloc[-1,-1] = hoja_calculo_disp.iloc[1,14]

        df_flagT1_aux['Availability'] = df_flagT1_aux['Availability']*100

        st.session_state.calculo_disponibilidad = False
        st.session_state.flag_T1_aux1 = False

        st.session_state.flag_T1_aux2 = True

        return df_flagT1_aux        

    elif st.session_state.calculo_pr:

        df_flag_T1_aux2 = df_aux

        xls = pd.ExcelFile(directory)

        hoja_calculo_disp = pd.read_excel(xls, 'Cálculo Disp. 24 (corr)', index_col=None, header = None)

        # Ver cuántos días tiene el mes, esa columna varía en longitud.
        end_tables = hoja_calculo_disp.loc[hoja_calculo_disp.iloc[:,42]=='TOTAL'].index[0] #Empezar, la primera siempre en el mismo lado, y la segunda el primer index +2 
        df_calculo_disp = hoja_calculo_disp.iloc[52:end_tables + 1,42:44].reset_index(drop = True)
        df_calculo_disp.columns = ['Fecha', 'Availability']

        df_flagT1_aux = df_flagT1_aux.merge(df_calculo_disp,on = 'Fecha')
        df_flagT1_aux.iloc[-1,-1] = hoja_calculo_disp.iloc[1,14]

        df_flagT1_aux['Availability'] = df_flagT1_aux['Availability']*100

        st.session_state.calculo_disponibilidad = False
        st.session_state.df_flag_T1_aux2 = False

        st.session_state.tablesDone = True        

        st.session_state.calculo_pr = False
        st.session_state.flag_T1_aux2 = False



        pass
