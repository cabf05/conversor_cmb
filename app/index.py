import streamlit as st 
import pdfplumber
import pandas as pd
import os
from io import BytesIO
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas

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
    "Atividade": ["ITEM", "Nome da Tarefa", "Nome da Atividade", "Atividade", "Tarefa", "Serviço"],
    "Data Inicio": ["Início", "INÍCIO", "Data de Inicio", "Data Inicial", "Data a ser Iniciada"],
    "Data Termino": ["Término", "TÉRMINO", "Data de Termino", "Data Final", "Data a ser Concluida", "Data Conclusao", "Fim"]
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
    # Converta as colunas 'Data Inicio' e 'Data Termino' para datetime, tratando erros
    df_arquivo['Data Inicio'] = pd.to_datetime(df_arquivo['Data Inicio'], errors="coerce", dayfirst=True)
    df_arquivo['Data Termino'] = pd.to_datetime(df_arquivo['Data Termino'], errors="coerce", dayfirst=True)
    
    # Calcule os dias de atividade, lidando com valores NaT
    df_arquivo['Dias de Atividade'] = (df_arquivo['Data Termino'] - df_arquivo['Data Inicio']).dt.days + 1
    
    # Substitua NaN por 0 ou outro valor padrão, se necessário
    df_arquivo['Dias de Atividade'].fillna(0, inplace=True)
    
    df_final = pd.DataFrame()
    for _, row in df_arquivo.iterrows():
        try:
            dias_de_atividade = int(row['Dias de Atividade'])  # Certifique-se de que é um inteiro
        except ValueError:
            dias_de_atividade = 0  # Ou outro valor padrão, se necessário
        
        if dias_de_atividade > 0:
            df_temporario = pd.DataFrame({
                'Atividade': [row['Atividade']] * dias_de_atividade,
                'Atividade Foi Realizada': ['Sim (    ) - Não (    )'] * dias_de_atividade,
                'Percentual Concluído': [""] * dias_de_atividade,
                'Data de Execução da Atividade': [row['Data Inicio'] + pd.Timedelta(days=i) for i in range(dias_de_atividade)],
            })
            df_final = pd.concat([df_final, df_temporario], ignore_index=True)
    
    df_final['Data de Execução da Atividade'] = df_final['Data de Execução da Atividade'].dt.strftime('%d/%m/%Y')
    return df_final



# Configurações DO PDF
fileName = f'Atividades_{local_atividade}.pdf'
documentTitle = 'Atividades Convertidas'
title = 'Relatório Diário de Obra'
subTitle = 'Atividades' 

def dataframe_para_pdf(df_final):
    buffer = BytesIO()
    pdf = canvas.Canvas(buffer, pagesize= A4)

    # DEFINIÇÕES PDF   
    pdf.setTitle(documentTitle)   

    data_de_execucao = df_final['Data de Execução da Atividade'].unique()

    # efetivo = [['arquiteto','Engenheiro','Estagiário','Encarregado','Pedreiro','Servente','Outros']]

    for data in data_de_execucao:
        df_data = df_final[df_final['Data de Execução da Atividade'] == data]

        pdf.setFont("Helvetica-Bold", 16)
        pdf.drawCentredString(300, 790, title) 

        pdf.setFont("Helvetica-Bold", 10)
        pdf.drawString(25, 750, 'LOCAL:')
        pdf.setFont('Helvetica', 10)
        pdf.drawString(100, 750, local_atividade)

        pdf.setFont("Helvetica-Bold", 10)
        pdf.drawString(25,735, 'DATA:')
        pdf.setFont('Helvetica', 10)
        pdf.drawString(100, 735, data)

        pdf.setFont("Helvetica-Bold", 10)
        pdf.drawString(25, 720, 'CLIMA:')
        pdf.line(100, 720, 150, 720)

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
        pdf.drawString(25, 575, 'EQUIPAMENTOS: ')

        # Linhas para a tabela de "Equipamentos"
        pdf.line(15, 570, 580, 570)     # Linha horizontal superior
        pdf.line(15, 570, 15, 525)      # Linha vertical esquerda
        pdf.line(580, 570, 580, 525)    # Linha vertical direita
        pdf.line(15, 525, 580, 525)     # Linha horizontal inferior
        pdf.line(15, 545, 580, 545)     # Linha horizontal interna

        # Colunas para a tabela de "Equipamentos"
        pdf.line(90, 570, 90, 525)
        pdf.line(175, 570, 175, 525)
        pdf.line(255, 570, 255, 525)
        pdf.line(360, 570, 360, 525)
        pdf.line(435, 570, 435, 525)
        pdf.line(515, 570, 515, 525)

        def draw_header():
            pdf.setFont("Helvetica-Bold", 10)
            pdf.drawString(25, 490, 'ATIVIDADES PENDENTES DIAS ANTERIORES')
            pdf.drawString(445, 490, 'PERCENTUAL CONCLUÍDO')
            
        # Iniciar a nova página e desenhar o cabeçalho
        draw_header()
        
        pdf.setFillColorRGB(0, 0, 0)
        y_posicao = 470
        topo = y_posicao - 10
        fundo = y_posicao - 90

        pdf.setFont("Helvetica", 8)
        pdf.line(15, 480, 580, 480)                                 # Linha Cabeçalho

        pdf.line(15, y_posicao + 10, 15, y_posicao - 90)            # 1° Linha vertical
        pdf.line(440, y_posicao + 10, 440, y_posicao - 90)          # 2° Linha vertical
        pdf.line(580, y_posicao + 10, 580, y_posicao - 90)          # 3° Linha vertical


        pdf.line(15, topo, 580, topo)                     # 1° Linha interna
        pdf.line(15, topo - 20, 580, topo - 20)           # 2° Linha interna
        pdf.line(15, topo - 40, 580, topo - 40)           # 3° Linha interna
        pdf.line(15, topo - 60, 580, topo - 60)           # 4° Linha interna
        pdf.line(15, fundo, 580, fundo)                   # 5° Linha interna

        def draw_header(pdf, y_position):
            pdf.setFont("Helvetica-Bold", 10)
            pdf.drawString(25, y_position, 'ATIVIDADE')
            pdf.drawString(303, y_position, 'ATIVIDADE FOI REALIZADA')
            pdf.drawString(445, y_position, 'PERCENTUAL CONCLUÍDO')
            
        # Iniciar a nova página e desenhar o cabeçalho
        draw_header(pdf, fundo - 30)

        # Escrevendo as atividades e a coluna "ATIVIDADE FOI REALIZADA"
        pdf.setFillColorRGB(0, 0, 0)
        y_posicao = fundo - 50
        pdf.line(15, y_posicao + 10, 580, y_posicao + 10)           # Linha Cabeçalho

        for _, row in df_data.iterrows():
            pdf.setFont("Helvetica", 8)
            pdf.drawString(25, y_posicao - 3, row['Atividade'])
            pdf.drawString(335, y_posicao, row['Atividade Foi Realizada'])
            pdf.drawString(445, y_posicao, row['Percentual Concluído'])

            pdf.line(15, y_posicao + 10, 15, y_posicao - 10)        # 1° Linha vertical
            pdf.line(300, y_posicao + 10, 300, y_posicao - 10)      # 2° Linha vertical
            pdf.line(440, y_posicao + 10, 440, y_posicao - 10)      # 3° Linha vertical
            pdf.line(580, y_posicao + 10, 580, y_posicao - 10)      # 4° Linha vertical
            pdf.line(15, y_posicao - 10, 580, y_posicao - 10)       # Linhas internas

            y_posicao -= 20
            if y_posicao < 50:
                pdf.showPage()
                draw_header(pdf, 790)  
                y_posicao = 770
                pdf.line(15, y_posicao + 10, 580, y_posicao + 10)           # Linha Cabeçalho

        pdf.showPage()  

        pdf.setFont("Helvetica-Bold", 16)
        pdf.drawCentredString(300, 790, f"{title} (Continuação)") 

        pdf.setFont("Helvetica-Bold", 10)
        pdf.drawString(25, 750, 'LOCAL:')
        pdf.setFont('Helvetica', 10)
        pdf.drawString(100, 750, local_atividade)

        pdf.setFont("Helvetica-Bold", 10)
        pdf.drawString(25,735, 'DATA:')
        pdf.setFont('Helvetica', 10)
        pdf.drawString(100, 735, data)

        # Desenhar a seção "REGISTRO DE OCORRÊNCIAS / PONTOS DE ATENÇÃO" no início da nova página
        pdf.setFont("Helvetica-Bold", 10)
        pdf.drawString(25, 660, 'REGISTRO DE OCORRÊNCIAS / PONTOS DE ATENÇÃO: ')

        # Ajuste das coordenadas Y para a tabela
        ocorrencias_y_top = 655  # Parte superior da tabela de "REGISTRO DE OCORRÊNCIAS / PONTOS DE ATENÇÃO"
        ocorrencias_y_bottom = ocorrencias_y_top - 120  # Altura da tabela

        pdf.line(15, ocorrencias_y_top, 580, ocorrencias_y_top)     # Linha superior
        pdf.line(15, ocorrencias_y_top, 15, ocorrencias_y_bottom)   # Linha esquerda
        pdf.line(580, ocorrencias_y_top, 580, ocorrencias_y_bottom) # Linha direita
        pdf.line(15, ocorrencias_y_bottom, 580, ocorrencias_y_bottom) # Linha inferior

        # Ajuste das coordenadas Y para "CHEGADA DE MATERIAL"
        pdf.setFont("Helvetica-Bold", 10)
        pdf.drawString(25, 510, 'CHEGADA DE MATERIAL: ')

        material_y_top = 505  # Linha superior da tabela de "CHEGADA DE MATERIAL"
        material_y_bottom = material_y_top - 120  # Altura da tabela

        # Linhas para a tabela de "CHEGADA DE MATERIAL"
        pdf.line(15, material_y_top, 580, material_y_top)               # Linha superior
        pdf.line(15, material_y_top, 15, material_y_bottom)             # Linha esquerda
        pdf.line(580, material_y_top, 580, material_y_bottom)           # Linha direita
        pdf.line(15, material_y_bottom, 580, material_y_bottom)         # Linha inferior

        def draw_header():
            pdf.setFont("Helvetica-Bold", 10)
            pdf.drawString(25, 360, 'ATIVIDADES REALIZADAS NÃO LISTADAS ANTERIORMENTE')
            pdf.drawString(445, 360, 'PERCENTUAL CONCLUÍDO')
            
        # Iniciar a nova página e desenhar o cabeçalho
        draw_header()

        pdf.setFillColorRGB(0, 0, 0)
        y_posicao = 345
        pdf.setFont("Helvetica", 8)
        pdf.line(15, 355, 580, 355)                                 # Linha Cabeçalho

        pdf.line(15, y_posicao + 10, 15, y_posicao - 90)            # 1° Linha vertical
        pdf.line(440, y_posicao + 10, 440, y_posicao - 90)          # 2° Linha vertical
        pdf.line(580, y_posicao + 10, 580, y_posicao - 90)          # 2° Linha vertical

        pdf.line(15, y_posicao - 10, 580, y_posicao - 10)           # 1° Linha interna
        pdf.line(15, y_posicao - 30, 580, y_posicao - 30)           # 2° Linha interna
        pdf.line(15, y_posicao - 50, 580, y_posicao - 50)           # 3° Linha interna
        pdf.line(15, y_posicao - 70, 580, y_posicao - 70)           # 4° Linha interna
        pdf.line(15, y_posicao - 90, 580, y_posicao - 90)           # 5° Linha interna

        pdf.setFont("Helvetica-Bold", 10)
        pdf.line(25, 70, 400, 70)
        pdf.drawString(25, 60, 'ASSINATURA RESPONSÁVEL PREENCHIMENTO')
        pdf.drawString(25, 30, 'ATENÇÃO: Enviar registro Fotográfico para o WhatsApp acessando o QR Code:')
        image = 'qrlogo.png'
        pdf.drawImage(image, 525, 25, 50, 50)

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
                        label = 'Baixar o Relatório de Atividades',
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
                        label = 'Baixar o Relatório de Atividades',
                        data = pdf_final,
                        file_name = fileName,
                        mime = 'application/pdf'
                    )
            else:
                st.error('As colunas Atividade, Data Início e/ou Data Término não foram encontradas no Excel.')

st.write("##")
st.write("Desenvolvido por CMB Capital")
st.write("© 2024 CMB Capital. Todos os direitos reservados.")