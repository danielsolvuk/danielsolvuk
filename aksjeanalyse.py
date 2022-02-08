import os.path
import pandas as pd
import numpy as np
import requests
import json

if os.path.isfile('./export.csv'):
    df_1 = pd.read_csv('./export.csv', encoding='ISO-8859-1', 
    sep=";",
    on_bad_lines='skip')
    
    df_1_new = df_1.rename(columns={'symbol': 'Ticker','name': 'Navn', 'date': 'Dato'})

    df_2 = pd.read_excel('./Investtech_Aksjestigning.xlsx')

    for i, rowi in df_1_new.iterrows():
        for j, rowj in df_2.iterrows():
            if (rowi['Ticker'] == rowj['Ticker']):
                df_1_new.concat([rowj['Børs']])
    #merge_data = pd.merge(df_2, df_1_new, on=['Ticker', 'Navn'], how='inner')

    #print(merge_data)

    #test.to_excel("test.xlsx", sheet_name="Sheet", index=True)

