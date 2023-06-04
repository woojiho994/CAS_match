# -*- coding: utf-8 -*-
"""
Created on Sun Jun  4 15:55:12 2023

@author: wooji
"""
import re
import streamlit as st
import numpy as np
import pandas as pd
import streamlit as st
import pandas as pd
from io import StringIO
import pdfplumber
import docx2pdf
st.title("MSDS报告CAS号提取程序")

#%%
def pdf_path(path):
    dirs = os.listdir(path)
    files = []
    for file in dirs:
        files.append(file)
    return(files)
#%% 函数二、打开pdf文件,输出每一页pdf中的所有文字
def openpdf(path):
    with pdfplumber.open(path) as pdf:
        # pdf = pdfplumber.open(path)
        item = []
        for page in pdf.pages:
            text = page.extract_text()
            item.append(text)
        # item = [''.join(i) for i in item]
        item = ';'.join(item).strip('')
    return item
#%% 上传文件区域
# uploaded_files = st.file_uploader('请上传MSDS报告',type=['pdf'],accept_multiple_files = True)
#%% 函数三、将目标CAS号，和pdf中的内容进行比对。返回什么？
def extract(text,cas):
    pattern = re.compile(cas,re.S)
    r_list = pattern.findall(text)
    return r_list
#%%
data = pd.DataFrame(columns=['CAS查询情况','匹配结果'])
uploaded_file = st.file_uploader("请上传pdf文件",accept_multiple_files=True)
if uploaded_file == []:
    st.stop()
else:
    # openpdf(uploaded_file)
    cas = r'[0-9]+-[0-9][0-9]-[0-9][^0-9]' 
    # st.write(extract(openpdf(uploaded_file),cas))
    for file in range(len(uploaded_file)):  
        # st.write(uploaded_file[file])
        if uploaded_file[file].name[-4:] == 'docx':
        st.write(uploaded_file[file])
        pythoncom.CoInitialize()
        docx2pdf.convert(uploaded_file[file].name)
        text = openpdf(uploaded_file[file].name[:-4]+'pdf')   
        else:    
            text = openpdf(uploaded_file[file])
        cas_extract = extract(text,cas)
        if cas_extract != []:
            for item in range(len(cas_extract)):
                cas_iso = cas_extract[item]
                cas_iso = cas_iso[0:len(cas_iso)-1]
                cas_set = pd.Series({uploaded_file[file].name:cas_iso})
                data = pd.concat([data,cas_set],axis=0)              
    data_reset_index = data.reset_index(drop=False)  
    #修改列名
    data_rename = data_reset_index.rename(columns={'index':'pdf名称',0:'CAS号提取'})    
    #去除重复行             
    data_output = data_rename.drop_duplicates()  #subset='pdf名称'可以查看是不是所有文件都包含在表格里
    # data_output.to_excel('5.28test-加入doc和docx.xlsx')
    # data_output[['pdf名称','CAS号提取']]
    target_data_base = pd.read_excel('102-104物质清单.xlsx',sheet_name='基102-3960种',index_col=0)
    target_data_pri = pd.read_excel('102-104物质清单.xlsx',sheet_name='基103-40种',index_col=0)
    target_data_key = pd.read_excel('102-104物质清单.xlsx',sheet_name='基104-14种',index_col=0)
    target_cas_base = target_data_base['CAS']
    target_cas_pri = target_data_pri['CAS']
    target_cas_key = target_data_key['CAS']

#%%
    for row in data_output.index:
        # print(data_output.loc[row]['CAS号提取'])
        for b in target_cas_base:
            if data_output.loc[row]['CAS号提取'] == b:
                data_output.loc[row]['匹配结果'] = '3960种'
        for j in target_cas_pri:
            if data_output.loc[row]['CAS号提取'] == j:
                data_output.loc[row]['匹配结果'] = '优评优控物质'
        for p in target_cas_key:
            if data_output.loc[row]['CAS号提取'] == p:
                data_output.loc[row]['匹配结果'] = '重点管控物质'
    
 
    data_final = data_output[['pdf名称','匹配结果','CAS号提取']]
    data_final


# #%%
# st.write("Here's our first attempt at using data to create a table:")
# st.write(pd.DataFrame({
#     'first column': [1, 2, 3, 4],
#     'second column': [10, 20, 30, 40]
# }))
# df = pd.DataFrame({
#   'first column': [5, 6, 7, 8],
#   'second column': [50, 60, 70, 80]
# })
# map_data = pd.DataFrame(
#     np.random.randn(1000, 2) / [50, 50] + [23.07,113.27],
#     columns=['lat', 'lon'])

# st.map(map_data)
# if st.checkbox('Show dataframe'):
#     chart_data = pd.DataFrame(
#        np.random.randn(20, 3),
#        columns=['a', 'b', 'c'])
#     chart_data
# option = st.selectbox(
#     'Which number do you like best?',
#      df['first column'])

# 'You selected: ', option
