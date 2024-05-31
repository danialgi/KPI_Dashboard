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

today_date = datetime.now().strftime('%Y-%m-%d')
current_datetime = datetime.now()

st.set_page_config(page_title="KPI Dashboard", page_icon="🚚", layout="wide")
st.write("🚚 Genuine Inside (M) Sdn. Bhd.")
st.title("KPI Dashboard 📊")
"_________________________________________________________"

weight_file = r'C:\Users\Danial Azrai\Downloads\WMS Item_Weights.xlsx'
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

df2=load_data(query_receive,engine)
df3=load_data(query_putaway,engine)
df1=load_data(query_pick,engine)
df_sort=load_data(query_sort,engine)
df_pack=load_data(query_pack,engine)
df_load=load_data(query_load,engine)

df2 = df2[df2[3].str.contains('BRP')]
df3 = df3[df3[4].str.contains('BRP')]
unique_st_df2 = df2[3].nunique()
unique_st_df3 = df3[4].nunique()

df_both2, df_left2= get_weight(df2, 51, df_weight)
df_both3, df_left3= get_weight(df3, 36, df_weight)
df_both1, df_left1= get_weight(df1, 45, df_weight)
df_both_sort, df_left_sort= get_weight(df_sort, 49, df_weight)
df_both_pack, df_left_pack= get_weight(df_pack, 15, df_weight)
df_both_load, df_left_load= get_weight(df_load, 22, df_weight)

df_empty2=empty(df_left2, 53, 54, 51, 52)
df_empty3=empty(df_left3, 38, 39, 36, 37)
df_empty1=empty(df_left1, 47, 48, 45, 46)
df_empty_sort=empty(df_left_sort, 51, 52, 49, 50)
df_empty_pack=empty(df_left_pack, 18, 19, 15, 16)
df_empty_load=empty(df_left_load, 25, 26, 22, 23)
df_empty= pd.concat([df_empty2, df_empty3, df_empty1, df_empty_sort, df_empty_pack, df_empty_load], ignore_index=True)
df_empty = df_empty.drop_duplicates(subset=['Item Code'], keep='first')

df_both1=points(df_both1, 11)
df_both_sort=points(df_both_sort, 10)
df_both_pack=points(df_both_pack, 6)
df_both_load=points(df_both_load, 10)

df_pick = df_both1 [[22,2,46,11,49,'weight','pts', 47]].copy()
df_pick.columns = ['Date', 'ID', 'Item', 'Quantity', 'Name', 'Weight','Pick', 'CustID']
#df_pick['Pick'] = df_pick.apply(lambda row: row['Pick'] * 0.5 if 'BRPZUC001' in row['CustID'] else row['Pick'], axis=1)
#df_pick.drop(columns=['CustID'], inplace=True)

df_sort = df_both_sort [[19,2,50,10,53,'weight','pts', 51]].copy()
df_sort.columns = ['Date', 'ID', 'Item', 'Quantity', 'Name', 'Weight','Sort', 'CustID']
#df_sort['Sort'] = df_sort.apply(lambda row: row['Sort'] * 0.5 if 'BRPZUC001' in row['CustID'] else row['Sort'], axis=1)
#df_sort.drop(columns=['CustID'], inplace=True)

df_pack = df_both_pack [[12,2,16,6,17,'weight','pts', 18]].copy()
df_pack.columns = ['Date', 'ID', 'Item', 'Quantity', 'Name', 'Weight','Pack', 'CustID']
df_pack['Pack'] = df_pack.apply(lambda row: row['Pack'] * 0.5 if 'BRPZUC001' in row['CustID'] else row['Pack'], axis=1)
df_pack.drop(columns=['CustID'], inplace=True)

df_load = df_both_load [[18,2,23,10,24,'weight','pts']].copy()
df_load.columns = ['Date', 'ID', 'Item', 'Quantity', 'Name', 'Weight','Load']
################################################################
#query = f"SELECT A.*, B.ItemCode, B.ItemDesc1, C.user_name, D.user_name, E.user_name, F.user_name  FROM tblPPSControl AS A JOIN tblMasterItem AS B ON A.ItemID = B.ItemID JOIN uvw_userlogin AS C ON A.BatchPickedBy= C.user_id JOIN uvw_userlogin AS D ON A.PickedBy= D.user_id JOIN uvw_userlogin AS E ON A.PackedBy= E.user_id JOIN uvw_userlogin AS F ON A.ShippedBy= F.user_id WHERE A.dateLastupdated BETWEEN '{start_datetime}' AND '{end_datetime}';"

#df=load_data(query,engine)
#df_both, df_left= get_weight(df, 55, df_weight)
#df_empty1=empty(df_left, 55, 56)
#df_both=points(df_both, 17)

#df_pick = df_both [[33,6,56,17,57,'weight','pts']].copy()
#df_pick.columns = ['Date', 'OrderID', 'Item', 'Quantity', 'Name', 'Weight','Pick']

#df_sort = df_both [[35,6,56,17,58,'weight','pts']].copy()
#df_sort.columns = ['Date', 'OrderID', 'Item', 'Quantity', 'Name', 'Weight','Sort']

#df_pack = df_both [[41,6,56,17,59,'weight','pts']].copy()
#df_pack.columns = ['Date', 'OrderID', 'Item', 'Quantity', 'Name', 'Weight','Pack']
#df_pack['Pack'] = df_pack.apply(lambda row: row['Pack'] * 0.5 if 'BRPZUC001' in row['OrderID'] else row['Pack'], axis=1)

#df_load = df_both [[43,6,56,17,60,'weight','pts']].copy()
#df_load.columns = ['Date', 'OrderID', 'Item', 'Quantity', 'Name', 'Weight','Load']

#df_chart= pd.concat([df_pick, df_sort, df_pack, df_load], ignore_index=True)
#df_bar = df_chart.groupby('Name', as_index=False).agg({'Pick': 'sum', 'Pack': 'sum', 'Sort': 'sum', 'Load': 'sum'})
#df_bar['Total'] = df_bar['Pick'] + df_bar['Sort'] + df_bar['Pack'] + df_bar['Load']
#stack_list=['Pick', 'Sort', 'Pack', 'Load']
###################################################################
Graph1, Graph2= st.columns(2)
with Graph1:
    df_chart_receive= df2[[21,3,52,7]].copy()
    df_chart_putaway= df3[[16,4,37,17]].copy()
    df_chart_receive.columns = ['Date','ID','Item', 'Quantity']
    df_chart_putaway.columns = ['Date','ID','Item', 'Quantity']
    values = [unique_st_df2, unique_st_df3]
    labels = ['Receive', 'Putaway']
    colors = ['lightgreen', 'darkgreen']

    fig1= px.bar(x=labels, y=values, color=labels, color_discrete_map={label: color for label, color in zip(labels, colors)}, labels={'x': '', 'y': 'Storage'}, title='📥Receive & Putaway')
    fig1.add_trace(go.Scatter(x=labels, y=values, mode='text', text=values, textposition='top center'))
    fig1.update_layout(height=415, showlegend=False)
    st.plotly_chart(fig1, use_container_width=True)

with Graph2:
     df_chart_pack= pd.concat([df_pick, df_sort, df_pack], ignore_index=True)
     df_bar = df_chart_pack.groupby('Name', as_index=False).agg({'Pick': 'sum', 'Pack': 'sum', 'Sort': 'sum'})
     df_bar['Total'] = df_bar['Pick'] + df_bar['Sort'] + df_bar['Pack']
     only=['Mu Mu Lwin','Myo Ma', 'Maung Oo', 'Shwe Win', 'Aung Soe Lin', 'Win Than Htay']
     df_bar = df_bar[df_bar['Name'].isin(only)]
     fig = px.bar(df_bar, x='Name', y=['Pick', 'Sort', 'Pack'], title='🎁Packing')
     fig.add_trace(go.Scatter(x=df_bar['Name'], y=df_bar['Total'], mode='text', text=df_bar['Total'].astype(int), textposition='top center'))
     #packing_kpi=number_days*600
     #fig.add_hline(y=15600, line_dash="dash", line_color="white")
     fig.update_layout(barmode='stack', xaxis={'categoryorder': 'total descending'}, height=415, showlegend=False)
     fig.update_xaxes(title_text='')
     fig.update_yaxes(title_text='Points')
     #fig.update_traces(marker_color='lightpink')
     st.plotly_chart(fig, use_container_width=True)

Graph3, Graph4 = st.columns(2)

df_chart_picksort= pd.concat([df_pick, df_sort, df_pack], ignore_index=True)
df_bar = df_chart_picksort.groupby('Name', as_index=False).agg({'Pick': 'sum', 'Pack': 'sum', 'Sort': 'sum'})
df_bar['Total'] = df_bar['Pick'] + df_bar['Sort'] + df_bar['Pack']
exclude=['Mu Mu Lwin','Myo Ma', 'Maung Oo', 'Shwe Win', 'Aung Soe Lin', 'Win Than Htay']
df_bar = df_bar[~df_bar['Name'].isin(exclude)]

fig = px.bar(df_bar, x='Name', y=['Pick', 'Sort', 'Pack'], title='🧺Pick & Sort')
fig.add_trace(go.Scatter(x=df_bar['Name'], y=df_bar['Total'], mode='text', text=df_bar['Total'].astype(int), textposition='top center'))
#picksort_kpi=number_days*1000
#fig.add_hline(y=26000, line_dash="dash", line_color="white")
fig.update_layout(barmode='stack', xaxis={'categoryorder': 'total descending'},
        legend_title_text='',
        height=500,
        legend=dict(y=1.1, orientation='h')
        )
fig.update_xaxes(title_text='')
fig.update_yaxes(title_text='Points')
st.plotly_chart(fig, use_container_width=True)
##############################################################

colA, _ ,colB = st.columns([1,5,1])
with colA:
    missing_product=st.toggle(f"Missing Weights ({len(df_empty)})", value=False)
if missing_product:
     df_empty

with colB:
    current_datetime

df_chart_picksort.rename(columns={'Pick': 'Pick(pts)', 'Sort': 'Sort(pts)', 'Pack': 'Pack(pts)', 'Load': 'Load(pts)'}, inplace=True)
df_chart_pack.rename(columns={'Pick': 'Pick(pts)', 'Sort': 'Sort(pts)', 'Pack': 'Pack(pts)', 'Load': 'Load(pts)'}, inplace=True)
excel_file = dfs_to_excel([df_chart_receive, df_chart_putaway, df_chart_picksort, df_chart_pack], ['Receiving', 'Putaway', 'Pick & Sort', 'Packing'])
st.sidebar.write("#")
st.sidebar.download_button(
            label="Export to Excel",
            data=excel_file,
            file_name=f"KPI_{today_date}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
