import streamlit as st
import pandas as pd
from io import BytesIO

data = pd.read_excel("testt.xlsx")
data.head()


data["media"] = (data["TOTAL_VENDAS"] / 3).round(2)

def to_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='dados')
    return output.getvalue()

excel_data = to_excel(data)

st.download_button(
    label="Baixar Excel",
    data=excel_data,
    file_name="dados.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)