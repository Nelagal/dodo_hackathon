import streamlit as st
import pandas as pd

st.title("Dodo Hackathon")
st.write("Ура, Streamlit работает!")

# Тест pandas
df = pd.DataFrame({
    'Ингредиент': ['Тесто', 'Сыр', 'Пепперони'],
    'Количество': [1, 2, 3]
})
st.dataframe(df)
st.write("Ура, PyCharm подружился с GitHub!")