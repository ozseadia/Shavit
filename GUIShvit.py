# -*- coding: utf-8 -*-
"""
Created on Wed Jul  8 09:13:18 2020

@author: OzSea
"""
import sys
import streamlit as st
#from streamlit import caching
import tkinter as tk
from tkinter import filedialog
import Report1 as RE
import ZivAv
import Landa
import Landa1
import Kornit1
import Kornit
import Kinetics
import time
import threading
import pandas as pd
Title1=st.empty()
Header1=st.empty()
Line1=st.empty()
Line2=st.empty()
Line3=st.empty()
global file_path_C , file_path_QT
file_path_C=[]
file_path_QT=[]





#global file_path
#file_path=list()
#SN=[]
#*****************************************************************************
#GUI by streamlit

#allow_output_mutation=True
#@st.cache_data()
def select_file(n):
    global file_path
    FileName = RE.SelectFile()
    #st.write(file_path)
    #st.write('fdgfdg %d' %(3)+'jshadjhak')
    return(FileName)



def KornitR(QT,C,Customer,file_path_C):
    Kornit.main(QT,C,Customer)
 
def KornitR1(QT,C,Customer,file_path_C):
    Kornit1.main(QT,C,Customer,file_path_C)    
    
def LandaR(QT,C,Customer,file_path_C):
    Landa1.main(QT,C,Customer,file_path_C)
         
def Kinetics1(QT,C,Customer,file_path_C):
    Kinetics.main(QT,C,Customer,file_path_C)
def Zivav(QT,C,Customer,file_path_C):
    ZivAv.main(QT,C,Customer,file_path_C)    


switcher = {
        'Landa': LandaR,
        'Kornit': KornitR1,
        'Elbit' : KornitR,
        'Kinetics':Kinetics1,
        'ZivAv':Zivav
    }    

#@st.cache() 
def main():
    global file_path_C , file_path_QT ,C,QT

    file_path_C=''
    #global file_path   
    #Title1.title("<h1 style='text-align: Right; color: red;'>דוח הזמנות פתוחות</h1>")
    st.markdown("<h1 style='text-align: Right; color: Black;'>דוח הזמנות פתוחות</h1>", unsafe_allow_html=True)
    #st.markdown("<h2 style='text-align: Right; color: Black;'>דוח הזמנות פתוחות</h1>", unsafe_allow_html=True)
    Customer=st.selectbox('Select Customer',('','Landa', 'Kornit', 'Kinetics','Elbit','ZivAv'))
    
    
    #N=st.text_input("Serial Number")
    #st.button('add')
    col1 ,col2 = st.columns(2)
    with col1:
        FileNameC = st.file_uploader("בחר דוח לקוח", type=["xlsx"])
        if FileNameC is not None: # Read the Excel file df = pd.read_excel(uploaded_file)
            C=pd.read_excel(FileNameC)
    with col2:
        FileNameQ = st.file_uploader("בחר דוח שביט", type=["xlsx"])
        if FileNameQ is not None: # Read the Excel file df = pd.read_excel(uploaded_file)
            QT=pd.read_excel(FileNameQ)    
    # if st.checkbox('בחר דוח לקוח'):
    #     file_path_C=select_file(1)
    
    #     C=RE.LoadData(file_path_C)
    
        
            
    # if st.checkbox('בחר דוח שביט'):
    #     file_path_QT=select_file(2)
    #     QT=RE.LoadData(file_path_QT)
    if FileNameQ and FileNameC:
        if st.button('הכן דוח'):
            
            # Get the function from switcher dictionary
            
            func = switcher.get(Customer)
            # Execute the function
            func(QT,C,Customer,file_path_C)
            #RE.main(QT,C,Customer)
            #Kornit.main(QT,C,Customer)

    if st.button('Clear all'):
        #caching.clear_cache()
        pass
        #st.checkbox("select report file", value = False)
        #st.checkbox('select Test case config file',value=False)
        
    uploaded_file = st.file_uploader("Choose a XLSX file", type="xlsx")

    if uploaded_file:
        df = pd.read_excel(uploaded_file)
    
        st.dataframe(df)
        st.table(df)
        
main()      