import streamlit as st
from PIL import Image
import os
import sys
from utils import *
import io
# Adicione o diretório 'utils' ao caminho de importação
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '../utils')))
#Encontra diretório atual
current_dir = os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))
#Achando o caminho do icone
icon_path = os.path.join('.streamlit','Logo_azul_quadrada.PNG')
st.set_page_config(page_title="Proposta Comercial", layout="wide",page_icon=icon_path)
#Achando o caminho da logo
img_path = os.path.join('.streamlit','Logo_goener_colorida.png')
img = Image.open(img_path)
# Redimensionar a imagem (alterando a altura e mantendo a proporção)
img = img.resize((int(img.width * (50 / img.height)), 50))  # Altura = 30 pixels
# Exibir a imagem redimensionada
st.image(img)
st.title('Proposta Comercial')
nome_cliente = st.text_input('Nome do cliente')

# Criando duas colunas para os inputs numéricos
col1, col2 = st.columns(2)

with col1:
    desconto = st.number_input('Desconto',step=5)
    custo_disponibilidade = st.number_input('Custo de Disponibilidade',step=10)

with col2:
    consumo = st.number_input('Consumo (Kwh)',step=100)
    fidelidade = st.number_input('Fidelidade',step=1)

# Salvando as informações em um dicionário
infos = {
    'nome_cliente': nome_cliente,
    'desconto': desconto,
    'consumo': consumo,
    'custo_disponibilidade': custo_disponibilidade,
    'fidelidade': fidelidade
}

create_button = st.button('Criar Apresentação')

if create_button:
    #Abertura do arquivo
    with open('APRESENTAÇÃO_GD.pptx','rb') as f:
        pptx_bytes = f.read()

    pptx_buffer = io.BytesIO(pptx_bytes)
    pptx_file = powerpoint_edit(infos,pptx_buffer)
    st.success('Apresentação Criada com Sucesso!')
    #Disponibilizar powerpoint para Download
    st.download_button(
        label="Baixar Apresentação Power Point",
        data=pptx_file,
        file_name=f"Apresentacao_{infos['nome_cliente']}.pptx",
        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
    )
    # #Conversão para pdf
    # pdf_file_name = f"Apresentacao_{infos['nome_cliente']}.pdf"
    # convert_ppt_to_pdf(f"Apresentacao_{infos['nome_cliente']}.pptx",f"Apresentacao_{infos['nome_cliente']}.pdf")
    # st.success('Apresentação em PDF criada com sucesso!')
    # #Disponibilizar pdf para download
    # with open(pdf_file_name,'rb') as f:
    #     pdf_data = f.read()
    # st.download_button(
    #     label='Baixar Apresentação em PDF',
    #     data= pdf_data,
    #     file_name = pdf_file_name,
    #     mime='application/pdf'
    # )