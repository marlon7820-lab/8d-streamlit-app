import streamlit as st

st.title("Test App")
st.write("If you see this, the app is running correctly.")

name = st.text_input("Enter your name:")
if st.button("Say Hello"):
    st.write(f"Hello, {name}!")
