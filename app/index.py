import streamlit as st 
import pdfplumber
import pandas as pd
import os
from io import BytesIO
from reportlab.lib.pagesizes import letter, landscape, A4
from reportlab.pdfgen import canvas
from reportlab.lib.units import inch
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase import pdfmetrics
from reportlab.platypus import SimpleDocTemplate, TableStyle, Table
from reportlab.lib import colors

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
            'Atividade Foi Realizada': ['Sim (    ) - Não (    )'] * row['Dias de Atividade'],
            'Percentual Concluído': "" * row['Dias de Atividade'],
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
title = 'Relatório Diário de Obra'
subTitle = 'Atividades' 

def dataframe_para_pdf(df_final):
    buffer = BytesIO()
    pdf = canvas.Canvas(buffer, pagesize= A4)
    # desenhar_regua(pdf)

    # DEFINIÇÕES PDF   
    pdf.setTitle(documentTitle)   

    data_de_execucao = df_final['Data de Execução da Atividade'].unique()

    # efetivo = [['arquiteto','Engenheiro','Estagiário','Encarregado','Pedreiro','Servente','Outros']]

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
        pdf.drawString(25,685, 'CLIMA:')
        pdf.line(100, 685, 150, 685)

        # Desenha a seção de "EFETIVO"
        pdf.drawString(25, 660, 'EFETIVO: ')

        # Linhas para a tabela de "EFETIVO"
        pdf.line(15, 655, 580, 655)     # Linha horizontal superior
        pdf.line(15, 655, 15, 610)      # Linha vertical esquerda
        pdf.line(580, 655, 580, 610)    # Linha vertical direita
        pdf.line(15, 610, 580, 610)     # Linha horizontal inferior
        pdf.line(15, 630, 580, 630)     # Linha horizontal interna

        # Colunas e títulos para a tabela de "EFETIVO"
        pdf.drawString(25, 640, 'ARQUITETO')
        pdf.line(90, 655, 90, 610)

        pdf.drawString(100, 640, 'ENGENHEIRO')
        pdf.line(175, 655, 175, 610)

        pdf.drawString(185, 640, 'ESTAGIÁRIO')
        pdf.line(255, 655, 255, 610)

        pdf.drawCentredString(310, 640, 'ENCARREGADO')
        pdf.line(360, 655, 360, 610)

        pdf.drawString(370, 640, 'PEDREIRO')
        pdf.line(435, 655, 435, 610)

        pdf.drawString(450, 640, 'SERVENTE')
        pdf.line(515, 655, 515, 610)

        pdf.drawString(525, 640, 'OUTROS')

        # Desenha a seção de "Equipamentos" abaixo de "EFETIVO"
        pdf.drawString(25, 580, 'EQUIPAMENTOS: ')

        # Linhas para a tabela de "Equipamentos"
        pdf.line(15, 575, 580, 575)     # Linha horizontal superior
        pdf.line(15, 575, 15, 530)      # Linha vertical esquerda
        pdf.line(580, 575, 580, 530)    # Linha vertical direita
        pdf.line(15, 530, 580, 530)     # Linha horizontal inferior
        pdf.line(15, 550, 580, 550)     # Linha horizontal interna

        # Colunas para a tabela de "Equipamentos"
        pdf.line(90, 575, 90, 530)
        pdf.line(175, 575, 175, 530)
        pdf.line(255, 575, 255, 530)
        pdf.line(360, 575, 360, 530)
        pdf.line(435, 575, 435, 530)
        pdf.line(515, 575, 515, 530)

        # Ajuste das coordenadas Y para "CHEGADA DE MATERIAL"
        material_title_y = 500  # Título "CHEGADA DE MATERIAL" 20 unidades abaixo da última linha de "EQUIPAMENTOS"
        material_y_top = material_title_y - 5  # Parte superior da tabela de "CHEGADA DE MATERIAL", 15 unidades abaixo do título
        material_y_bottom = material_y_top - 120  # Altura da tabela

        pdf.setFont("Helvetica-Bold", 10)
        pdf.drawString(25, material_title_y, 'CHEGADA DE MATERIAL: ')

        # Linhas para a tabela de "CHEGADA DE MATERIAL"
        pdf.line(15, material_y_top, 580, material_y_top)     # Linha superior
        pdf.line(15, material_y_top, 15, material_y_bottom)   # Linha esquerda
        pdf.line(580, material_y_top, 580, material_y_bottom) # Linha direita
        pdf.line(15, material_y_bottom, 580, material_y_bottom) # Linha inferior

        # Mover para uma nova página
        pdf.showPage()

        # Desenhar a seção "REGISTRO DE OCORRÊNCIAS / PONTOS DE ATENÇÃO" no início da nova página
        pdf.setFont("Helvetica-Bold", 10)
        pdf.drawString(25, 800, 'REGISTRO DE OCORRÊNCIAS / PONTOS DE ATENÇÃO: ')

        # Ajuste das coordenadas Y para a tabela
        ocorrencias_y_top = 790  # Parte superior da tabela de "REGISTRO DE OCORRÊNCIAS / PONTOS DE ATENÇÃO"
        ocorrencias_y_bottom = ocorrencias_y_top - 120  # Altura da tabela

        pdf.line(15, ocorrencias_y_top, 580, ocorrencias_y_top)     # Linha superior
        pdf.line(15, ocorrencias_y_top, 15, ocorrencias_y_bottom)   # Linha esquerda
        pdf.line(580, ocorrencias_y_top, 580, ocorrencias_y_bottom) # Linha direita
        pdf.line(15, ocorrencias_y_bottom, 580, ocorrencias_y_bottom) # Linha inferior

        def draw_header():
            pdf.setFont("Helvetica-Bold", 10)
            pdf.drawString(25, ocorrencias_y_bottom - 30, 'ATIVIDADE')
            pdf.drawString(300, ocorrencias_y_bottom - 30, 'ATIVIDADE FOI REALIZADA')
            pdf.drawString(450, ocorrencias_y_bottom - 30, 'PERCENTUAL CONCLUÍDO')
            
        # Iniciar a nova página e desenhar o cabeçalho
        draw_header()

        # Escrevendo as atividades e a coluna "ATIVIDADE FOI REALIZADA"
        pdf.setFillColorRGB(0, 0, 0)
        y_posicao = ocorrencias_y_bottom - 60
        for _, row in df_data.iterrows():
            pdf.setFont("Helvetica", 8)
            pdf.drawString(25, y_posicao - 3, row['Atividade'])
            pdf.drawString(335, y_posicao, row['Atividade Foi Realizada'])
            pdf.drawString(450, y_posicao, row['Percentual Concluído'])

            pdf.line(300, y_posicao + 10, 300, y_posicao - 10)      # 1° Linha vertical
            pdf.line(440, y_posicao + 10, 440, y_posicao - 10)      # 2° Linha vertical
            pdf.line(15, y_posicao - 10, 580, y_posicao - 10)       # Linhas internas

            y_posicao -= 20
            if y_posicao < 50:
                pdf.showPage()
                draw_header()  # Desenha o cabeçalho na nova página
                y_posicao = 770

        pdf.setFont("Helvetica-Bold", 10)
        pdf.line(25, 70, 400, 70)
        pdf.drawString(25, 60, 'ASSINATURA RESPONSÁVEL PREENCHIMENTO')
        pdf.drawString(25, 30, 'ATENÇÃO: Enviar registro Fotográfico para o WhatsApp acessando o QR Code:')
        image = 'qrlogo.png'
        pdf.drawImage(image, 525, 25, 50, 50)


        pdf.showPage()  # NOVA PÁGINA PARA CADA DATA DE EXECUÇÃO
        pdf.setFont("Helvetica-Bold", 16)
        pdf.drawCentredString(300, 770, f"{title} (Continuação)") 

        pdf.setFont("Helvetica-Bold", 10)
        pdf.drawString(25, 725, 'LOCAL:')
        pdf.setFont('Helvetica', 10)
        pdf.drawString(100, 725, local_atividade)

        pdf.setFont("Helvetica-Bold", 10)
        pdf.drawString(25,705, 'DATA:')
        pdf.setFont('Helvetica', 10)
        pdf.drawString(100, 705, data)

        pdf.setFont("Helvetica-Bold", 10)
        pdf.drawString(25,685, 'ATIVIDADES REALIZADAS NÃO LISTADAS ANTERIORMENTE:')

        def draw_header():
            pdf.setFont("Helvetica-Bold", 10)
            pdf.drawString(25, 655, 'ATIVIDADE')
            pdf.drawString(450, 655, 'PERCENTUAL CONCLUÍDO')
            
        # Iniciar a nova página e desenhar o cabeçalho
        draw_header()

        # Escrevendo as atividades e a coluna "ATIVIDADE FOI REALIZADA"
        pdf.setFillColorRGB(0, 0, 0)
        y_posicao = 640
        pdf.setFont("Helvetica", 8)
        pdf.line(440, y_posicao + 10, 440, y_posicao - 100)      # 2° Linha vertical

        pdf.line(15, y_posicao - 10, 580, y_posicao - 10)        # 1° Linha interna
        pdf.line(15, y_posicao - 30, 580, y_posicao - 30)        # 2° Linha interna
        pdf.line(15, y_posicao - 50, 580, y_posicao - 50)        # 3° Linha interna
        pdf.line(15, y_posicao - 70, 580, y_posicao - 70)        # 4° Linha interna
        pdf.line(15, y_posicao - 90, 580, y_posicao - 90)        # 5° Linha interna

        y_posicao -= 20
        if y_posicao < 50:
            pdf.showPage()
            draw_header()  # Desenha o cabeçalho na nova página
            y_posicao = 770

        pdf.showPage()

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