import streamlit as st
from src.prediction import Insurance_Prediction
st.title("Insurance Premium Prediction")
st.write("Description about your project ")
Age = st.number_input("Enter Age")
Annual_Income_LPA = st.number_input("Enter Annual Income in LPA")
Policy_Term_Years = st.number_input("Enter Policy Term in Years")
Sum_Assured_Lakhs = st.number_input("Enter Sum Assured in Lakhs")

if st.button("Predict"):
    model = Insurance_Prediction()
    result = model.prediction(Age,Annual_Income_LPA,Policy_Term_Years,Sum_Assured_Lakhs)
    st.success(result)