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


def excel_reader(directory):


    
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



        pass
   

        st.session_state.calculo_plantilla = False
