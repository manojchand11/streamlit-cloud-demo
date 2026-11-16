import streamlit as st
import pandas as pd
import plotly.express as px

st.set_page_config(page_title="Cloud Dashboard", layout="wide")

st.title("India Energy Analytics")
st.write("My first cloud-hosted Streamlit dashboard.")

df = pd.DataFrame({
    "Month": ["Jan", "Feb", "Mar", "Apr", "May", "Jun"],
    "Demand": [180000, 190000, 195000, 205000, 215000, 225000],
    "Generation": [175000, 188000, 192000, 201000, 212000, 220000]
})

c1, c2 = st.columns(2)
c1.metric("Average Demand", "201,667 MW")
c2.metric("Average Generation", "198,000 MW")

fig = px.line(df, x="Month", y=["Demand", "Generation"], markers=True)
st.plotly_chart(fig, use_container_width=True)

st.success("Successfully deployed on Streamlit Cloud!")
