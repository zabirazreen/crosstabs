import streamlit as st
import pandas as pd
import datetime as dt
import math
from io import BytesIO
from pyxlsb import open_workbook as open_xlsb
from PIL import Image
import matplotlib.pyplot as plt
from colour import Color

hide_st_style = """
            <style>
            #MainMenu {visibility: hidden;}
            footer {visibility: hidden;}
            header {visibility: hidden;}
            </style>
            """

st.markdown(hide_st_style, unsafe_allow_html=True)

def single_choice_crosstab(df, q, column =None, value='weight', column_seq=None, row_seq=None):
    if row_seq != None:
        row_labels = row_seq + ["Grand Total"]
    else:
        row_labels = list(dict(df[q].value_counts()).keys()) + ["Grand Total"]
    df_ct = pd.DataFrame({q:row_labels})
    if column_seq != None:
        column_seq = column_seq + ['Grand Total']
    else:
        column_seq = list(df[column].unique()) + ['Grand Total']
    for demo in column_seq:
        temp = []
        for row in df_ct[q]:
            if row != 'Grand Total':
                if demo != 'Grand Total':
                    total_sum = df[df[column] == demo][value].sum()
                    temp_df = df[(df[column] == demo) & (df[q] == row)]
                    temp.append(round(temp_df[value].sum()/total_sum, 4))
                else:
                    temp_df = df[df[q] == row]
                    temp.append(round(temp_df[value].sum()/df[value].sum(), 4))
            else:
                temp.append(1)
        df_ct[demo] = temp
    if row_seq == None:
        df_ct = pd.concat([df_ct[:-1].sort_values('Grand Total', ascending = False),df_ct[-1:]])
    return df_ct

def multi_choice_crosstab(df, q, column, value='weight', column_seq=None):
    if column_seq != None:
        column_seq =  column_seq + ['Grand Total']
    else:
        column_seq = list(df[column].unique())
        column_seq.sort()
        column_seq = column_seq + ['Grand Total']
    demo_dict = {}
    for demo in column_seq:
        ans_dict = {}
        if demo == 'Grand Total':
            demo_df = df
        else:
            demo_df = df[df[column] == demo]
            
        for i in demo_df.index:
            answer = str(demo_df[q][i])
            if answer != 'nan':
                answer = answer.split(', ')
                total_weight = 0
                for ans in answer:
                    total_weight = df[value][i]
                    if ans not in ans_dict:
                        ans_dict[ans] = df[value][i]
                    else:
                        ans_dict[ans] += df[value][i]
                    
        for key, val in ans_dict.items():
            ans_dict[key] = round(val/sum(list(demo_df[value])),4)   
        ans_dict = dict(sorted(ans_dict.items(), key=lambda x: x[1], reverse=True))
        if demo == 'Grand Total':
            row_labels = list(ans_dict.keys())
            gt = list(ans_dict.values())
        else:
            demo_dict[demo] = ans_dict
    result = pd.DataFrame({q:row_labels})
    for demo in demo_dict:
        temp = []
        for row in row_labels:
            if row in demo_dict[demo]:
                temp.append(demo_dict[demo][row])
            else:
                temp.append(0.0000)
        result[demo] = temp
    result['Grand Total'] = gt
    return result
    
image = Image.open('invoke_logo.jpg')
st.title('Crosstabs Generator')
st.image(image)

st.subheader("Upload Survey responses (csv/xlsx)")
df = st.file_uploader("Please ensure the data are cleaned and weighted (if need to be) prior to uploading.")
if df:
    df_name = df.name
    if df_name[-3:] == 'csv':
        df = pd.read_csv(df, na_filter = False)
    else:
        df = pd.read_excel(df, na_filter = False)
    
    weight = st.selectbox('Select weight column', ['', 'Unweighted'] + list(df.columns))
    if weight != '':
        demos = st.multiselect('Choose the demograhic(s) you want to build the crosstabs across', list(df.columns))
        if len(demos) > 0:
            score = 0
            col_seqs = {}
            for demo in demos:
                st.subheader('Column: ' + demo)
                col_seq = st.multiselect('Please arrange ALL values in order', list(df[demo].unique()), key = demo)
                col_seqs[demo] = col_seq
                if len(col_seq) == df[demo].nunique():
                    score += 1
            if score == len(demos):
                first = st.selectbox('Select the first question of the survey', [''] + list(df.columns))
                if first != '':
                    first_idx = list(df.columns).index(first)
                    last = st.selectbox('Select the last question of the survey', [''] + list(df.columns)[first_idx + 1:])
                    if last != '':
                        last_idx = list(df.columns).index(last)
                        st.subheader('Number of questions to build the crosstab on: ' + str(last_idx - first_idx + 1))
                        q_ls= [df.columns[x] for x in range(first_idx, last_idx + 1)]
                        multi = st.multiselect('Choose mutiple answers question(s), if any', list(df.columns)[first_idx: last_idx + 1])
                        button = st.button('Generate Crosstabs')
                        if button:
                            with st.spinner('Building crosstabs...'):
                                output = BytesIO()
                                writer = pd.ExcelWriter(output, engine='xlsxwriter')
                                df.to_excel(writer, index=False, sheet_name= 'data')
                                for demo in demos:
                                    start = 1
                                    for q in q_ls:
                                        if q in multi:
                                            table = multi_choice_crosstab(df, q, demo, value= weight, column_seq= col_seqs[demo])
                                        else:
                                            table = single_choice_crosstab(df, q, demo, value= weight, column_seq= col_seqs[demo])

                                        table.to_excel(writer, index=False, sheet_name=demo, startrow = start)
                                        start = start + len(table) + 3
                                        workbook = writer.book
                                        worksheet = writer.sheets[demo]
                            
                            writer.save()
                            df_xlsx = output.getvalue()
                            df_name = df_name[:df_name.find('.')]
                            st.balloons()
                            st.header('Crosstabs ready for download!')
                            st.download_button(label='📥 Download', data=df_xlsx, file_name= df_name + '-crosstabs.xlsx')
                                



                    




