import datetime as datetime
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
from google.oauth2 import service_account
import win32com.client as win32
import pandas as pd
import numpy as np

import gspread
import openpyxl as xl
import json

def update_data():
    xlApp = win32.Dispatch('Excel.Application')
    # cambiar ruta por la ruta donde este alojado el excel con las consultas
    wb = xlApp.Workbooks.Open('path.../w_file.p')

    
def save_to_gsheets(df):         
    # Generar key.json ver README.md 
    client = gspread.service_account('key.json')
    sheet = client.open_by_url('URL spreadsheet')
    worksheet = sheet.worksheet('NAME spreadsheet')

    # Se converte el tipo de las columnas que sean datetime a string
    for column in df.columns[df.dtypes == 'datetime64[ns]']:
        df[column] = df[column].astype(str)

    
    worksheet.update(
        [df.columns.values.tolist()] +
        # se reemplazan los valores no deseados [ NaN, NaT, Errors.. ] por strings vacíos    
         df.fillna('')
         .replace(0,'')
         .replace('NaT','')
         .replace('Error 11','')
         .replace('Error 6','')
         .values.tolist() )

# cambiar ruta por la ruta donde este alojado el excel con las consultas
df1 = pd.read_excel('path.../w_file.py', engine='openpyxl')    

if __name__ == "__main__":
    update_data()
    save_to_gsheets(df1)


