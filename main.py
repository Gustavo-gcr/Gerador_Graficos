import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt

# Configuração da página do Streamlit
st.set_page_config(layout="wide")

# Título do aplicativo
st.title('Análise de Categorias, Atendentes e Situações')

# Upload do arquivo Excel
uploaded_file = st.file_uploader("Escolha o arquivo Excel", type="xlsx")

if uploaded_file is not None:
    # Leitura das abas do arquivo Excel
    xls = pd.ExcelFile(uploaded_file)
    
    # Leitura da aba "Comparativo" sem pular linhas
    comparativo_df = pd.read_excel(xls, sheet_name='Comparativo')
    
    # Leitura da aba "Worksheet" pulando as 5 primeiras linhas
    worksheet_df = pd.read_excel(xls, sheet_name='Worksheet', header=5)
    
    # Verificação se as colunas necessárias existem no DataFrame da aba "Worksheet"
    required_columns = ['Categoria', 'Atendente', 'Origem do Chamado', 'Última Situação']
    if all(col in worksheet_df.columns for col in required_columns):
        # Criação das abas
        tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs(["Análise por Categoria", "Análise por Atendente", "Painel do Atendente", "Situação", "Treinamento", "Comparativo"])
        
        with tab1:
            # Contagem das ocorrências de cada categoria
            category_counts = worksheet_df['Categoria'].value_counts()
            
            # Criação de colunas para layout
            col1, col2 = st.columns(2)
            
            # Exibição da tabela de contagem
            col1.subheader('Contagem de Categorias')
            col1.write(category_counts)
            
            # Criação do gráfico
            col2.subheader('Gráfico de Categorias')
            fig, ax = plt.subplots()
            bars = category_counts.plot(kind='bar', ax=ax)
            ax.set_xlabel('Categoria')
            ax.set_ylabel('Contagem')
            
            # Adicionando os valores no topo de cada barra
            for bar in bars.patches:
                ax.annotate(format(bar.get_height(), '.0f'), 
                            (bar.get_x() + bar.get_width() / 2, bar.get_height()), 
                            ha='center', va='center', 
                            size=10, xytext=(0, 8), 
                            textcoords='offset points')
                
            col2.pyplot(fig)
        
        with tab2:
            # Contagem das ocorrências de cada atendente
            attendant_counts = worksheet_df['Atendente'].value_counts()
            
            # Criação de colunas para layout
            col3, col4 = st.columns(2)
            
            # Exibição da tabela de contagem
            col3.subheader('Contagem de Atendentes')
            col3.write(attendant_counts)
            
            # Criação do gráfico
            col4.subheader('Gráfico de Atendentes')
            fig, ax = plt.subplots()
            bars = attendant_counts.plot(kind='bar', ax=ax)
            ax.set_xlabel('Atendente')
            ax.set_ylabel('Contagem')
            
            # Adicionando os valores no topo de cada barra
            for bar in bars.patches:
                ax.annotate(format(bar.get_height(), '.0f'), 
                            (bar.get_x() + bar.get_width() / 2, bar.get_height()), 
                            ha='center', va='center', 
                            size=10, xytext=(0, 8), 
                            textcoords='offset points')
                
            col4.pyplot(fig)
        
        with tab3:
            # Filtrando os dados pela coluna "Origem do Chamado" para "Painel do Atendente"
            painel_df = worksheet_df[worksheet_df['Origem do Chamado'] == 'Painel do Atendente']
            
            # Contagem das ocorrências de cada atendente no Painel do Atendente
            painel_attendant_counts = painel_df['Atendente'].value_counts()
            
            # Criação de colunas para layout
            col5, col6 = st.columns(2)
            
            # Exibição da tabela de contagem
            col5.subheader('Contagem de Atendentes (Painel do Atendente)')
            col5.write(painel_attendant_counts)
            
            # Criação do gráfico
            col6.subheader('Gráfico de Atendentes (Painel do Atendente)')
            fig, ax = plt.subplots()
            bars = painel_attendant_counts.plot(kind='bar', ax=ax)
            ax.set_xlabel('Atendente')
            ax.set_ylabel('Contagem')
            
            # Adicionando os valores no topo de cada barra
            for bar in bars.patches:
                ax.annotate(format(bar.get_height(), '.0f'), 
                            (bar.get_x() + bar.get_width() / 2, bar.get_height()), 
                            ha='center', va='center', 
                            size=10, xytext=(0, 8), 
                            textcoords='offset points')
                
            col6.pyplot(fig)
        
        with tab4:
            # Contagem das ocorrências de cada situação
            situacao_counts = worksheet_df['Última Situação'].value_counts()
            
            # Criação de colunas para layout
            col7, col8 = st.columns(2)
            
            # Exibição da tabela de contagem
            col7.subheader('Contagem por Última Situação')
            col7.write(situacao_counts)
            
            # Criação do gráfico
            col8.subheader('Gráfico de Última Situação')
            fig, ax = plt.subplots()
            bars = situacao_counts.plot(kind='bar', ax=ax)
            ax.set_xlabel('Última Situação')
            ax.set_ylabel('Contagem')
            
            # Adicionando os valores no topo de cada barra
            for bar in bars.patches:
                ax.annotate(format(bar.get_height(), '.0f'), 
                            (bar.get_x() + bar.get_width() / 2, bar.get_height()), 
                            ha='center', va='center', 
                            size=10, xytext=(0, 8), 
                            textcoords='offset points')
                
            col8.pyplot(fig)

        with tab5:
            # Filtrando os dados pela coluna "Categoria" para "Treinamento"
            treinamento_df = worksheet_df[worksheet_df['Categoria'] == 'Treinamento']
            
            # Contagem das ocorrências de cada atendente no Treinamento
            treinamento_attendant_counts = treinamento_df['Atendente'].value_counts()
            
            # Criação de colunas para layout
            col9, col10 = st.columns(2)
            
            # Exibição da tabela de contagem
            col9.subheader('Contagem de Atendentes (Treinamento)')
            col9.write(treinamento_attendant_counts)
            
            # Criação do gráfico
            col10.subheader('Gráfico de Atendentes (Treinamento)')
            fig, ax = plt.subplots()
            bars = treinamento_attendant_counts.plot(kind='bar', ax=ax)
            ax.set_xlabel('Atendente')
            ax.set_ylabel('Contagem')
            
            # Adicionando os valores no topo de cada barra
            for bar in bars.patches:
                ax.annotate(format(bar.get_height(), '.0f'), 
                            (bar.get_x() + bar.get_width() / 2, bar.get_height()), 
                            ha='center', va='center', 
                            size=10, xytext=(0, 8), 
                            textcoords='offset points')
                
            col10.pyplot(fig)
        
        with tab6:
            # Exibição dos dados da aba "Comparativo"
            col10, col11 = st.columns(2)
            
            with col10:
                st.subheader('Dados da aba Comparativo')
                st.write(comparativo_df)
            
            with col11:
                # Verificação se as colunas '2023' e '2024' estão presentes
                if 'ano 2023' in comparativo_df.columns and 'ano 2024' in comparativo_df.columns:
                    # Plotar os gráficos para os dados de 2023 e 2024
                    st.subheader('Comparativo de Chamados 2023 vs 2024')
                    
                    fig, ax = plt.subplots(figsize=(10, 6))
                    comparativo_df.plot(x='Mês', y=['ano 2023', 'ano 2024'], kind='bar', ax=ax)
                    ax.set_xlabel('Mês')
                    ax.set_ylabel('Quantidade de Chamados')
                    ax.set_title('Comparativo de Chamados por Mês: 2023 vs 2024')
                    
                    # Adicionando os valores no topo de cada barra
                    for container in ax.containers:
                        ax.bar_label(container, label_type='edge', padding=3)
                    
                    st.pyplot(fig)
                else:
                    st.error('As colunas "2023" e/ou "2024" não foram encontradas na aba "Comparativo".')
    else:
        st.error('As colunas "Categoria", "Atendente", "Origem do Chamado", "Última Situação" e/ou "Data do Chamado" não foram encontradas no arquivo Excel.')
else:
    st.info('Por favor, carregue um arquivo Excel para começar.')