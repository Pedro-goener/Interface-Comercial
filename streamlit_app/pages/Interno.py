import streamlit as st
from utils.interacao_db import db_config,load_and_prepare_data
import os
if not st.session_state.get('login_status', False) or not st.session_state.get('admin',False):
    st.warning('Página para acompanhamento interno')
    st.stop()
from PIL import Image
import io
import datetime
#Encontra diretório atual
current_dir = os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))
#Achando o caminho do icone
icon_path = os.path.join('.streamlit','Logo_azul_quadrada.PNG')
st.set_page_config(page_title="Acompanhamento", layout="wide",page_icon=icon_path)
#Achando o caminho da logo
img_path = os.path.join('.streamlit','Logo_goener_colorida.png')
img = Image.open(img_path)
# Redimensionar a imagem (alterando a altura e mantendo a proporção)
img = img.resize((int(img.width * (50 / img.height)), 50))  # Altura = 30 pixels
# Exibir a imagem redimensionada
st.image(img)
st.title('Acompanhamento de propostas')
st.subheader('Propostas')
df_propostas = load_and_prepare_data(db_config,'SELECT * FROM propostas')
st.dataframe(df_propostas)
