import streamlit as st 
import pdfplumber
import pandas as pd
import os
from io import BytesIO
from reportlab.lib.pagesizes import letter, landscape
from reportlab.pdfgen import canvas
from reportlab.lib.units import inch
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase import pdfmetrics
from reportlab.platypus import Table

st.set_page_config(
    layout="wide", 
    page_title="Auditar Engenharia"
)

local_atividade = st.text_input(label= "Informe o Local de realização das Atividades")

with st.sidebar:
    st.header("CONVERSOR DE ATIVIDADES")
    arquivo_atividades = st.file_uploader(
        label="Selecione o Arquivo com as Atividades",
        type=("pdf", "xlsx")
    )

possiveis_colunas = {
    "Atividade": ["ITEM", "Nome da Tarefa", "Nome da Atividade", "Atividade", "Tarefa"],
    "Data Inicio": ["Início", "INÍCIO", "Data de Inicio", "Data Inicial", "Data a ser Iniciada"],
    "Data Termino": ["Término", "TÉRMINO", "Data de Termino", "Data Final", "Data a ser Concluida", "Data Conclusao"]
}

def renomear_colunas(df_arquivo, possiveis_colunas):
    nova_coluna = {}
    for nome_padrao, nomes_alternativos in possiveis_colunas.items():
        for nome in nomes_alternativos:
            if nome in df_arquivo.columns:
                nova_coluna[nome] = nome_padrao
                break
    return df_arquivo.rename(columns=nova_coluna)

def extrair_datas_do_arquivo(caminho_arquivo):
    atividades, data_inicio, data_termino = [], [], []
    with pdfplumber.open(caminho_arquivo) as pdf:
        for pagina in pdf.pages:
            try:
                tabela = pagina.extract_table()
                if tabela:
                    df_arquivo = pd.DataFrame(tabela[1:], columns=tabela[0])
                    df_arquivo = renomear_colunas(df_arquivo, possiveis_colunas)
                    if 'Atividade' in df_arquivo.columns and 'Data Inicio' in df_arquivo.columns and 'Data Termino' in df_arquivo.columns:
                        atividades.extend(df_arquivo['Atividade'].dropna().tolist())
                        data_inicio.extend(df_arquivo['Data Inicio'].dropna().tolist())
                        data_termino.extend(df_arquivo['Data Termino'].dropna().tolist())
            except Exception as e:
                st.error(f"Erro ao Processar a Página: {str(e)}")
        return atividades, data_inicio, data_termino

def calcular_dias_de_atividade(df_arquivo):
    df_arquivo['Data Inicio'] = pd.to_datetime(df_arquivo['Data Inicio'], errors="coerce", dayfirst=True)
    df_arquivo['Data Termino'] = pd.to_datetime(df_arquivo['Data Termino'], errors="coerce", dayfirst=True)
    df_arquivo['Dias de Atividade'] = (df_arquivo['Data Termino'] - df_arquivo['Data Inicio']).dt.days + 1

    df_final = pd.DataFrame()
    for _, row in df_arquivo.iterrows():
        df_temporario = pd.DataFrame({
            'Atividade': [row['Atividade']] * row['Dias de Atividade'],
            'Atividade Foi Realizada': ['Sim (    ) | Não (    )'] * row['Dias de Atividade'],
            'Data de Execução da Atividade': [row['Data Inicio'] + pd.Timedelta(days=i) for i in range(row['Dias de Atividade'])],
        })
        df_final = pd.concat([df_final, df_temporario], ignore_index=True)

    df_final['Data de Execução da Atividade'] = df_final['Data de Execução da Atividade'].dt.strftime('%d/%m/%Y')
    return df_final

# def desenhar_regua(pdf):
    #pdf.drawString(100,810, 'x100')
    #pdf.drawString(200,810, 'x200')
    #pdf.drawString(300,810, 'x300')
    #pdf.drawString(400,810, 'x400')
    #pdf.drawString(500,810, 'x500')
  
    #pdf.drawString(10,100, 'y100')
    #pdf.drawString(10,200, 'y200')
    #pdf.drawString(10,300, 'y300')
    #pdf.drawString(10,400, 'y400')
    #pdf.drawString(10,500, 'y500')
    #pdf.drawString(10,600, 'y600')
    #pdf.drawString(10,700, 'y700')
    #pdf.drawString(10,800, 'y800')


# Configurações DO PDF
fileName = f'Atividades_{local_atividade}.pdf'
documentTitle = 'Atividades Convertidas'
title = 'Atividades Convertidas'
subTitle = 'Atividades' 

def dataframe_para_pdf(df_final):
    buffer = BytesIO()
    pdf = canvas.Canvas(buffer)
    # desenhar_regua(pdf)

    # DEFINIÇÕES PDF   
    pdf.setTitle(documentTitle)   

    data_de_execucao = df_final['Data de Execução da Atividade'].unique()

    for data in data_de_execucao:
        df_data = df_final[df_final['Data de Execução da Atividade'] == data]

        pdf.setFont("Helvetica-Bold", 16)
        pdf.drawCentredString(300, 770, title) 

        pdf.setFont("Helvetica-Bold", 10)
        pdf.drawString(25, 725, 'LOCAL:')
        pdf.setFont('Helvetica', 10)
        pdf.drawString(100, 725, local_atividade)

        pdf.setFont("Helvetica-Bold", 10)
        pdf.drawString(25,705, 'DATA:')
        pdf.setFont('Helvetica', 10)
        pdf.drawString(100, 705, data)

        pdf.setFont("Helvetica-Bold", 10)
        pdf.drawString(25,680, 'EFETIVO: ')

        pdf.setFont('Helvetica', 10)
        pdf.drawString(25, 660, 'Arquiteto: ')
        pdf.line(100, 660, 400, 660)

        pdf.drawString(25, 645, 'Engenheiro: ')
        pdf.line(100, 645, 400, 645)

        pdf.drawString(25, 630, 'Estagiário: ')
        pdf.line(100, 630, 400, 630)

        pdf.drawString(25, 615, 'Encarregado: ')
        pdf.line(100, 615, 400, 615)

        pdf.drawString(25, 600, 'Pedreiro: ')
        pdf.line(100, 600, 400, 600)

        pdf.drawString(25, 585, 'Servente: ')
        pdf.line(100, 585, 400, 585)

        pdf.drawString(25, 570, 'Outros: ')
        pdf.line(100, 570, 400, 570)

        pdf.setFont("Helvetica-Bold", 10)
        pdf.drawString(25,545, 'CHEGADA DE MATERIAL: ')

        pdf.line(25, 530, 400, 530)
        pdf.line(25, 515, 400, 515)
        pdf.line(25, 500, 400, 500)
        pdf.line(25, 485, 400, 485)
        pdf.line(25, 470, 400, 470)

        pdf.setFont("Helvetica-Bold", 10)
        pdf.drawString(25, 450, 'REGISTRO DE OCORRÊNCIAS / PONTOS DE ATENÇÃO: ')

        pdf.line(25, 435, 400, 435)
        pdf.line(25, 420, 400, 420)
        pdf.line(25, 405, 400, 405)
        pdf.line(25, 390, 400, 390)
        pdf.line(25, 375, 400, 375)

        pdf.setFont("Helvetica-Bold", 14)
        pdf.setFillColorRGB(0, 0, 255)
        pdf.drawString(25,335, 'ATIVIDADE')

        pdf.setFont("Helvetica-Bold", 14)
        pdf.drawString(350, 335, 'ATIVIDADE FOI REALIZADA')

        # ESCREVER AS ATIVIDADES E A COLUNA "ATIVIDADE FOI REALIZADA"
        pdf.setFillColorRGB(0, 0, 0)
        y_posicao = 315
        for _, row in df_data.iterrows():
            pdf.setFont("Helvetica", 10)
            pdf.drawString(25, y_posicao, row['Atividade'])
            pdf.drawString(400, y_posicao, row['Atividade Foi Realizada'])
            y_posicao -= 20
            if y_posicao < 50:
                pdf.showPage()
                y_posicao = 770
                
        pdf.setFont("Helvetica-Bold", 10)
        pdf.drawCentredString(300, 95, 'ATENÇÃO: ')
        pdf.drawCentredString(300, 80, 'Enviar registro Fotográfico para o WhatsApp acessando o QR Code abaixo:')
        image = 'qrlogo.png'
        pdf.drawImage(image, 275, 25, 50, 50)

        pdf.showPage()  # NOVA PÁGINA PARA CADA DATA DE EXECUÇÃO

    pdf.save()
    buffer.seek(0)
    return buffer

if arquivo_atividades is not None:
    if arquivo_atividades.type == 'application/pdf':
        with st.spinner('Processando o Arquivo PDF...'):
            with open('temp.pdf', 'wb') as arquivo_temporario:
                arquivo_temporario.write(arquivo_atividades.read())

            atividades, data_inicio, data_termino = extrair_datas_do_arquivo('temp.pdf')
            os.remove("temp.pdf")

            if atividades and data_inicio and data_termino:
                dias_de_atividade = pd.DataFrame({
                    "Atividade": atividades,
                    "Data Inicio": data_inicio,
                    "Data Termino": data_termino
                })
                dias_de_atividade = calcular_dias_de_atividade(dias_de_atividade)

                pdf_final = dataframe_para_pdf(dias_de_atividade)
                
                if local_atividade.strip() == "":
                    st.warning = "Informe o Local de realização das Atividades"
                else:
                    
                    st.download_button(
                        label = 'Download do Arquivo PDF',
                        data = pdf_final,
                        file_name = fileName,
                        mime = 'application/pdf'
                    )
            else:
                st.error('As colunas Atividade, Data Início e/ou Data Término não foram encontradas no PDF.')

    elif arquivo_atividades.type == 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet':
        with st.spinner('Processando o Arquivo EXCEL...'):
            dias_de_atividade = pd.read_excel(arquivo_atividades)
            dias_de_atividade = renomear_colunas(dias_de_atividade, possiveis_colunas)
            if 'Atividade' in dias_de_atividade.columns and 'Data Inicio' in dias_de_atividade.columns and 'Data Termino' in dias_de_atividade.columns:
                dias_de_atividade = calcular_dias_de_atividade(dias_de_atividade)

                pdf_final = dataframe_para_pdf(dias_de_atividade)

                if local_atividade.strip() == "":
                    st.warning = ("Informe o Local de realização das Atividades")
                else:
                    st.download_button(
                        label = 'Download do Arquivo PDF',
                        data = pdf_final,
                        file_name = fileName,
                        mime = 'application/pdf'
                    )
            else:
                st.error('As colunas Atividade, Data Início e/ou Data Término não foram encontradas no Excel.')

st.write("##")
st.write("Desenvolvido por CMB Capital")
st.write("© 2024 CMB Capital. Todos os direitos reservados.")