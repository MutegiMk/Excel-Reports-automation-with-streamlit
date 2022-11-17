import streamlit as st
import pandas as pd
import plotly.express as px
import base64
from io import StringIO, BytesIO


def generate_excel_download_link(df):
    towrite = BytesIO()
    df.to_excel(towrite, encoding="utf-8", index=False, header=True)  # write to BytesIO buffer
    towrite.seek(0)  # reset pointer
    b64 = base64.b64encode(towrite.read()).decode()
    href = f'<a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}" download="data_download.xlsx">Download Excel File</a>'
    return st.markdown(href, unsafe_allow_html=True)

def generate_html_download_link(fig):
    towrite = StringIO()
    fig.write_html(towrite, include_plotlyjs="cdn")
    towrite = BytesIO(towrite.getvalue().encode())
    b64 = base64.b64encode(towrite.read()).decode()
    href = f'<a href="data:text/html;charset=utf-8;base64, {b64}" download="plot.html">Download Plot</a>'
    return st.markdown(href, unsafe_allow_html=True)


st.set_page_config(page_title= 'Excel Plotting')
st.title('Excel Plotter 📊')
st.subheader('Upload your Excel file')

uploaded_file = st.file_uploader('Choose your XLSX file', type='xlsx')
if uploaded_file:
    st.markdown('---')
    df = pd.read_excel(uploaded_file, engine='openpyxl')
    st.dataframe(df)

    groupby_column = st.selectbox(
        'Select column to analyze',
        ('Activity Date','Bakery Item',"Dee's Location",'Activity Type','Inventory date'),
    )

    # Grouping dataframe
    output_columns = ['Quantity']
    df_grouped = df.groupby(by=[groupby_column], as_index=False)[output_columns].sum()
    #st.dataframe(df_grouped)

    # Plotting Dataframe
    fig = px.bar(
        df_grouped,
        x = groupby_column,
        y = 'Quantity',
        color_continuous_scale=['red','yellow','green'],
        template='plotly_dark',
        title = f'<b> Quantity by {groupby_column}</b>'
    )

    st.plotly_chart(fig)

    # Downloadng
    st.subheader('Downloads:')
    generate_excel_download_link(df_grouped)
    generate_html_download_link(fig)

