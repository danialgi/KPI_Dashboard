import streamlit as st
import pandas as pd
from sqlalchemy import create_engine, Column, Integer, String, ForeignKey
from sqlalchemy.orm import sessionmaker, aliased
from sqlalchemy.ext.declarative import declarative_base
from sqlalchemy.sql import join
from urllib.parse import quote_plus
import plotly.express as px
import webbrowser as wb
import openpyxl
from openpyxl import load_workbook
from openpyxl.styles import Alignment
import os
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import numpy as np
from io import BytesIO
import calendar
from datetime import datetime, date, time

st.cache_data.clear()
today_date = datetime.now().strftime('%Y-%m-%d')
current_datetime = datetime.now()

st.set_page_config(page_title="KPI Dashboard", page_icon="🚚", layout="wide")
st.write("🚚 Genuine Inside (M) Sdn. Bhd.")
st.title("KPI Dashboard 📊")
"_________________________________________________________"

weight_file = 'WMS Item_Weights.xlsx'
df_weight = pd.read_excel(weight_file, engine='openpyxl')
df_weight = df_weight.drop_duplicates(subset=['Product'], keep='first')

def assign_points(x):
    if x < 1:
        return 1
    elif 1 <= x < 3:
        return 2
    elif 3 <= x < 10:
        return 4
    elif x >= 10:
        return 8
    else:
        return None

@st.cache_data
def load_data(query,_engine):
    df = pd.read_sql(query, con=engine)
    df.columns = range(df.shape[1])
    #df
    return df

@st.cache_data
def dfs_to_excel(df_list, sheet_list):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        for dataframe, sheet in zip(df_list, sheet_list):
            dataframe.to_excel(writer, sheet_name=sheet, index=False)
    output.seek(0)
    return output

def handle_leading_zeros(x):
    if x.startswith('0'):
        return x.lstrip('0')
    else:
        return x

def get_weight(df, product_column, df_weight):
    df['Product'] = df[product_column].astype(str)
    df_weight['Product'] = df_weight['Product'].astype(str)
    df['Product'] = df[product_column].apply(handle_leading_zeros)
    df_weight['Product'] = df_weight['Product'].apply(handle_leading_zeros)
    df['Product'] = df['Product'].str.strip()

    df_merge = pd.merge(df, df_weight, on='Product', how='outer', indicator=True)
    df_both = df_merge[df_merge['_merge'] == 'both']
    df_left = df_merge[df_merge['_merge'] == 'left_only']
    return df_both, df_left

def empty(df_left, id, cust, code, name):
    df_empty  = df_left[[id, cust, code, name]].copy()
    df_empty = df_empty.drop_duplicates(subset=[code], keep='first')
    df_empty.columns = ['Cust ID', 'Cust Name' , 'Item Code', 'Item Name']
    df_empty.reset_index(inplace=True)
    df_empty = df_empty .drop(['index'], axis=1)
    return df_empty

def points(df_both, quantity_column):
    df_both['weight']=df_both['Weight(kg)']*df_both[quantity_column]
    df_both['pts'] = df_both['weight'].apply(assign_points)
    return df_both

def bar_chart(df_bar, stack_list, title, line_value):
    fig = px.bar(df_bar, x='Name', y=stack_list, title=title)
    fig.add_trace(go.Scatter(x=df_bar['Name'], y=df_bar['Total'], mode='text', text=df_bar['Total'].astype(int), textposition='top center'))
    fig.add_hline(y=line_value, line_dash="dash", line_color="white")
    fig.update_layout(barmode='stack', xaxis={'categoryorder': 'total descending'},
            legend_title_text='',
            height=425,
            legend=dict(y=1.1, orientation='h')
            )
    fig.update_xaxes(title_text='')
    fig.update_yaxes(title_text='Points')
    st.plotly_chart(fig, use_container_width=True)

username = "brp_user"
password = "brp!@#456"
hostname = "192.168.2.4"
port = "1433"
dbname = "SHWMSDBV2"
encoded_password = quote_plus(password)
engine = create_engine(f"mssql+pymssql://{username}:{encoded_password}@{hostname}:{port}/{dbname}")
session = sessionmaker(bind=engine)()

st.sidebar.write("Select Date: ")
with st.sidebar.form(key='filter_form'):
    start_date = st.date_input("Start", pd.to_datetime(today_date))
    end_date = st.date_input("End", pd.to_datetime(today_date))
    number_days = (end_date - start_date).days + 1
    start_datetime = datetime.combine(start_date, datetime.min.time())
    end_datetime = datetime.combine(end_date, time.max.replace(microsecond=0))

    query_receive = f"SELECT A.*, B.user_name, C.ItemCode, C.ItemDesc1, D.Custcode, D.Custname FROM tblReceivingDetails AS A JOIN uvw_userlogin AS B ON A.LastUpdatedBy= B.user_id JOIN tblMasterItem AS C ON A.ItemID = C.ItemID JOIN tblMasterOwner AS D ON C.CustID=D.Custid WHERE A.DateLastUpdated BETWEEN '{start_datetime}' AND '{end_datetime}';"
    query_putaway = f"SELECT A.*, B.user_name, C.ItemCode, C.ItemDesc1, D.Custcode, D.Custname FROM tblPutawayDetails AS A JOIN uvw_userlogin AS B ON A.LastUpdatedBy= B.user_id JOIN tblMasterItem AS C ON A.ItemID = C.ItemID JOIN tblMasterOwner AS D ON C.CustID=D.Custid WHERE A.DateLastUpdated BETWEEN '{start_datetime}' AND '{end_datetime}';"
    query_pick = f"SELECT A.*, B.ItemCode, B.ItemDesc1, D.Custcode, D.Custname, C.user_name FROM tblPickingDetailsBatch AS A JOIN tblMasterItem AS B ON A.ItemID = B.ItemID JOIN tblMasterOwner AS D ON B.CustID=D.Custid JOIN uvw_userlogin AS C ON A.LastUpdatedBy= C.user_id WHERE A.DateLastUpdated BETWEEN '{start_datetime}' AND '{end_datetime}';"
    query_sort = f"SELECT A.*, B.ItemCode, B.ItemDesc1, D.Custcode, D.Custname, C.user_name FROM tblPickingDetails AS A JOIN tblMasterItem AS B ON A.ItemID = B.ItemID JOIN tblMasterOwner AS D ON B.CustID=D.Custid JOIN uvw_userlogin AS C ON A.LastUpdatedBy= C.user_id WHERE A.DateLastUpdated BETWEEN '{start_datetime}' AND '{end_datetime}';"
    query_pack = f"SELECT A.*, B.ItemCode, B.ItemDesc1, C.user_name, D.Custcode, D.Custname FROM tblPackingItemDetails AS A JOIN tblMasterItem AS B ON A.ItemID = B.ItemID JOIN uvw_userlogin AS C ON A.LastUpdatedBy= C.user_id JOIN tblMasterOwner AS D ON B.CustID=D.Custid WHERE A.DateLastUpdated BETWEEN '{start_datetime}' AND '{end_datetime}';"
    query_load = f"SELECT A.*, B.ItemCode, B.ItemDesc1, C.user_name, D.Custcode, D.Custname FROM tblShipmentDetails AS A JOIN tblMasterItem AS B ON A.ItemID = B.ItemID JOIN uvw_userlogin AS C ON A.LastUpdatedBy= C.user_id JOIN tblMasterOwner AS D ON B.CustID=D.Custid WHERE A.DateLastUpdated BETWEEN '{start_datetime}' AND '{end_datetime}';"
    submitted = st.form_submit_button('Filter')

def check_db_connection(host, user, password, database):
    try:
        connection = mysql.connector.connect(
            host=host,
            user=user,
            password=password,
            database=database
        )
        cursor = connection.cursor()
        cursor.execute("SELECT 1")  # Execute a simple query
        cursor.close()
        connection.close()
        return True
    except mysql.connector.Error:
        return False

if check_db_connection(hostname, username, password, dbname):
    print("Database connection is active.")
else:
    print("Database connection failed.")

