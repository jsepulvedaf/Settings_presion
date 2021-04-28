# -*- coding: utf-8 -*-
"""
Created on Sun Mar 14 16:46:54 2021

@author: JSEPULVEDA-PC
""" 

import streamlit as st
import datetime

import pandas as pd
import plotly
import plotly.express as px
from typing import Dict
import openpyxl
from openpyxl import Workbook
from openpyxl.chart import (
    LineChart,
    Reference,
)
from openpyxl.chart.axis import DateAxis

def main():
    
     
    




        st.title("APLICACION ESTABLECER SETTINGS DE REGULACION")
        st.text("by @jsepulvedaf")
        st.subheader("Por favor subir el archivo teniendo en cuenta que las columnas deben tener el siguiente orden y los encabezados asi:  fecha, Caudal PE PS PC PM")
        
        #lecutra y carggue 
        
        data_file = st.file_uploader("Suba el archivo",type=['xlsx'])
        df= pd.read_excel (data_file)
        
        sec =st.text_input("Entre el Sector")
        sector=sec.upper()     
        vlr=st.text_input("Valor presion minima PC")
        vlr_min=float(vlr)
        
       
        
        table = pd.pivot_table(df, values=['Caudal', 'PE','PS','PC','PM'], index=[df['fecha'].dt.hour])
        
        
        
        
        
        
        
        
        table['delta_p']=table['PC']-vlr_min
        table['Parametro']=table['PS']- table['delta_p']
        
        
        table1=round(table.iloc[0:4,6].max(),0)
        table2=round(table.iloc[4:8,6].max(),0)
        table3=round(table.iloc[8:12,6].max(),0)
        table4=round(table.iloc[12:16,6].max(),0)
        table5=round(table.iloc[16:19,6].max(),0)
        table6=round(table.iloc[19:23,6].max(),0)
        
        table['set_proy']=0
        
        
        table.iloc[0,7]=table1
        table.iloc[1,7]=table1
        table.iloc[2,7]=table1
        table.iloc[3,7]=table1
        
            
        
        table.iloc[4,7]=table2
        table.iloc[5,7]=table2
        table.iloc[6,7]=table2    
        table.iloc[7,7]=table2
        
        table.iloc[8,7]=table3
        table.iloc[9,7]=table3
        table.iloc[10,7]=table3       
        table.iloc[11,7]=table3
           
        
        table.iloc[12,7]=table4
        table.iloc[13,7]=table4
        table.iloc[14,7]=table4
        table.iloc[15,7]=table4
           
            
        table.iloc[16,7]=table5
        table.iloc[17,7]=table5
        table.iloc[18,7]=table5   
        table.iloc[19,7]=table5
        
        
        table.iloc[20,7]=table6
        table.iloc[21,7]=table6
        table.iloc[22,7]=table6
        table.iloc[23,7]=table6
        
        # table.iloc[table.groupby('C1').agg(max_ = ('C3', lambda data: data.idxmax())).max_]
        
        
        table['error']=table['set_proy']-table['Parametro']
        table['set_final']=0
        
        for n in range(0,24,1):
            
            if table.iloc[n,8]<0  :
                table.iloc[n,9]=(table.iloc[n,7])+1
                
            else:
                table.iloc[n,9]=table.iloc[n,7]
                
            if table.iloc[n,8]> 3  :
                
                table.iloc[n,9]=(table.iloc[n-1,7])  
            
               
                
        
        
        st.dataframe(table)
        
       
       
        nombre_columna = table.columns.tolist()    
        seleccion=st.multiselect("seleeciones los campos", nombre_columna)
        selected = table[seleccion]
        
        fig = px.line(selected, title ='Grafico: '+str(seleccion) )  
        #fig.show()
        fig.write_html("Resumen_settings.html")
            
        
        st.plotly_chart(fig, use_container_width=True)
    
        
        file=  'settings_final_'+sector+'.xlsx'    
        boton=st.button("exportar XLSX")        
               
        if boton :
            table.to_excel(file)
            st.write("archivo guardado")
       
                      
        wb = openpyxl.load_workbook(file)
        sheet =wb.sheetnames
       
        ws=wb.active
        
        c1 = LineChart()
        c1.title = "Settings Regulación  "+sector
        c1.style = 13
        c1.y_axis.title = 'P m.c.a'
        c1.x_axis.title = 'Hora'

        data = Reference(ws, min_col=8, min_row=1, max_col=11, max_row=25)
        c1.add_data(data, titles_from_data=True)

        s1 = c1.series[0]
        s1.marker.symbol = "triangle"
        s1.marker.graphicalProperties.solidFill = "7E3F00" # Marker filling
        s1.marker.graphicalProperties.line.solidFill = "7E3F00" # Marker outline
        
        s1.graphicalProperties.line.noFill = True
        
        s2 = c1.series[1]
        s2.graphicalProperties.line.solidFill = "00AAAA"
        s2.graphicalProperties.line.solidFill = "sysDot"
        s2.graphicalProperties.line.solidFill= "7B241C " # width in EMUs
        
        s2 = c1.series[2]
        s2.smooth = True # Make the line smooth

        
        
        ws.add_chart(c1, "N3")
        wb.save(file)
        

    
if __name__ == '__main__':
    	main()    
    