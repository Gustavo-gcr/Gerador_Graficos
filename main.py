import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import plotly.express as px
import plotly.graph_objects as go

# Configuração da página do Streamlit
st.set_page_config(layout="wide")

# Título do aplicativo
st.title('Relatório Mensal TomTicket')

# Upload do arquivo Excel
uploaded_file = st.file_uploader("Escolha o arquivo Excel", type="xlsx")

if uploaded_file is not None:
    # Leitura das abas do arquivo Excel
    xls = pd.ExcelFile(uploaded_file)
    
    
    # Leitura da aba "Comparativo" sem pular linhas
    comparativo_df = pd.read_excel(xls, sheet_name='Comparativo')
    
    comparativo_detalhado_df = pd.read_excel(xls, sheet_name='Comparativo Detalhado')
    
    # Leitura da aba "Worksheet" pulando as 5 primeiras linhas
    worksheet_df = pd.read_excel(xls, sheet_name='Worksheet', header=5)

    # Leitura da aba "Comparativo Detalhado"
    comparativo_detalhado_df = pd.read_excel(xls, sheet_name='Comparativo Detalhado')

    # Verificação se as colunas necessárias existem no DataFrame da aba "Worksheet"
    required_columns = ['Categoria', 'Atendente', 'Origem do Chamado', 'Última Situação']
    if all(col in worksheet_df.columns for col in required_columns):
        # Criação das abas
        tab1, tab2, tab3, tab9, tab4, tab5, tab6, tab7, tab8  = st.tabs(["Análise por Categoria", "Análise por Atendente", "Painel do Atendente","Detalhamento de Atendente", "Situação", "Treinamento", "Comparativo Total Anual", "Comparativo Mês Anual", "Comparativo Categorias Anual"])
        
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
            col6.subheader('Gráfico de Atendentes')
            fig, ax = plt.subplots()
            bars = painel_attendant_counts.plot(kind='bar', ax=ax, color='skyblue', edgecolor='black')
            ax.set_xlabel('Atendente')
            ax.set_ylabel('Contagem')
            ax.set_title('Quantidade de Atendimentos criados pelo proprio atendente')

            # Adicionando os valores no topo de cada barra
            for bar in bars.patches:
                ax.annotate(format(bar.get_height(), '.0f'), 
                            (bar.get_x() + bar.get_width() / 2, bar.get_height()), 
                            ha='center', va='center', 
                            size=10, xytext=(0, 8), 
                            textcoords='offset points')
                        
            col6.pyplot(fig)

            # Mini planilha com categorias por atendente
            st.subheader("Categorias por Atendente")
            categories_by_attendant = painel_df.groupby(['Atendente', 'Categoria']).size().unstack(fill_value=0)
            st.write(categories_by_attendant)

        with tab9:
            st.subheader('Detalhamento de chamados criados por Atendente ')

            # Filtrando os dados pela coluna "Origem do Chamado" para incluir apenas "Painel do Atendente"
            painel_atendente_df = worksheet_df[worksheet_df['Origem do Chamado'] == 'Painel do Atendente']
            
            # Contagem de categorias, mas não vamos exibir isso
            contagem_categorias = painel_atendente_df.groupby(['Atendente', 'Categoria']).size().reset_index(name='Quantidade')

            # Seletor para o atendente específico
            selected_attendant = st.selectbox("Selecione o atendente para visualizar detalhes", contagem_categorias['Atendente'].unique())
            
            # Filtrando os detalhes dos chamados para o atendente selecionado e origem "Painel do Atendente"
            chamados_detalhados = painel_atendente_df[painel_atendente_df['Atendente'] == selected_attendant]

            # Exibindo os detalhes do chamado
            st.write(f"Chamados detalhados para o atendente {selected_attendant} no Painel do Atendente:")
            st.write(chamados_detalhados[['Protocolo', 'Assunto', 'Categoria', 'Última Situação']])

       
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

        with tab7:
            # Criação do gráfico de comparação
            st.subheader('Comparativo de Chamados por Mês entre Anos')
            
            meses = comparativo_detalhado_df['Mês'].unique()
            anos = comparativo_detalhado_df['Ano'].unique()
            
            # Seletor de mês
            selected_month = st.selectbox('Selecione o mês', meses)
            
            fig = go.Figure()
            for ano in anos:
                df_ano = comparativo_detalhado_df[comparativo_detalhado_df['Ano'] == ano]
                df_mes = df_ano[df_ano['Mês'] == selected_month]
                if not df_mes.empty:
                    total_chamados = df_mes.iloc[0, 2:].sum()
                    fig.add_trace(go.Bar(
                        x=[selected_month],
                        y=[total_chamados],
                        name=f"{selected_month} {ano}",
                        width=0.3  # Ajuste da largura da barra
                    ))
            
            fig.update_layout(
                title='Comparativo de Chamados por Mês entre Anos',
                xaxis_title='Mês',
                yaxis_title='Total de Chamados',
                barmode='group'
            )
            
            # Criação de colunas para layout
            col1, col2, col3 = st.columns([1.5, 0.75, 0.75])
            
            with col1:
                st.plotly_chart(fig)
            
            with col2:
                if selected_month:
                    st.subheader(f'(2023)')
                    
                    # Contagem de categorias para 2023
                    categoria_counts_2023 = comparativo_detalhado_df[(comparativo_detalhado_df['Mês'] == selected_month) & (comparativo_detalhado_df['Ano'] == 2023)].sum(numeric_only=True)
                    
                    # Exibição das contagens detalhadas em caixas
                    st.write(categoria_counts_2023)
            with col3:
                if selected_month:
                    st.subheader(f'(2024)')
                    
                    # Contagem de categorias para 2024
                    categoria_counts_2024 = comparativo_detalhado_df[(comparativo_detalhado_df['Mês'] == selected_month) & (comparativo_detalhado_df['Ano'] == 2024)].sum(numeric_only=True)
                    
                    # Exibição das contagens detalhadas em caixas
                    st.write(categoria_counts_2024)
                else:
                    st.error("A aba 'Comparativo Detalhado' não foi encontrada no arquivo Excel.")
        with tab8:
            st.subheader('Análise de Categorias')

            # Verificar se 'Categoria' pode ser representado pelas colunas do DataFrame
            categorias = comparativo_detalhado_df.columns[3:]  # Ajuste o índice conforme necessário para excluir 'Mês' e 'Ano'
            
            # Seletor de categorias (várias opções)
            selected_categories = st.multiselect('Selecione as categorias', categorias)
            
            # Seletor de ano(s)
            selected_years = [2023, 2024]  # Definido para os anos 2023 e 2024

            if selected_categories:
                # Filtrar os dados para as categorias selecionadas
                df_categoria = comparativo_detalhado_df[['Mês', 'Ano'] + selected_categories]
                
                # Verificar se os dados têm os nomes dos meses como strings e convertê-los se necessário
                if df_categoria['Mês'].dtype == 'object':
                    month_mapping = {
                        'janeiro': 1, 'fevereiro': 2, 'março': 3, 'abril': 4, 'maio': 5, 'junho': 6,
                        'julho': 7, 'agosto': 8, 'setembro': 9, 'outubro': 10, 'novembro': 11, 'dezembro': 12
                    }
                    df_categoria['Mês'] = df_categoria['Mês'].map(month_mapping)
                
                # Transformar o DataFrame em formato longo para plotagem
                df_melted = df_categoria.melt(id_vars=['Mês', 'Ano'], value_vars=selected_categories, var_name='Categoria', value_name='Quantidade')
                
                # Filtrar os dados para os anos selecionados
                df_melted = df_melted[df_melted['Ano'].isin(selected_years)]

                # Criar uma nova coluna para combinar mês e ano
                df_melted['Mês_Ano'] = df_melted.apply(lambda row: f"{row['Mês']:02d}/{row['Ano']}", axis=1)

                # Definir uma lista de cores
                colors = [
                     '#A6CEE3', '#1F78B4', '#BSDF8A', '#33A02C', '#FB9A99', '#E31A1C', '#FDBF6F', '#FF7F00', '#CAB2D6', '#6A3D9A',
    '#FFFF99', '#B15928', '#FFB3B3', '#BEBADA', '#FB8072', '#80B1D3', '#FDB462', '#B3DE69', '#FCCDE5', '#D9D9D9',
    '#BC80BD', '#CCEBC5', '#FFED6F', '#9E0142', '#D53E4F', '#F46D43', '#FDAE61', '#FEE08B', '#E6F598', '#ABDDA4',
    '#66C2A5', '#3288BD', '#5E4FA2', '#9E9AC8', '#B3CDE3', '#CCEBC5', '#DECBE4', '#FED9A6', '#FFFFCC', '#E5D8BD',
    '#FDDAEC', '#F2F2F2', '#B2182B', '#D6604D', '#F4A582', '#FDDBC7', '#GCRG10','#D1E5F0', '#92C5DE', '#4393C3', '#2166AC'
                ]
                
                # Criar gráfico de barras empilhadas (Evolução mensal)
                fig = go.Figure()

                # Adicionar traços de barras empilhadas para cada categoria e ano
                for i, categoria in enumerate(selected_categories):
                    for ano in selected_years:
                        df_categoria_ano = df_melted[(df_melted['Categoria'] == categoria) & (df_melted['Ano'] == ano)]
                        fig.add_trace(go.Bar(
                            x=df_categoria_ano['Mês_Ano'],
                            y=df_categoria_ano['Quantidade'],
                            name=f'{categoria} {ano}',
                            text=df_categoria_ano['Quantidade'],
                            textposition='auto',
                            marker=dict(color=colors[i % len(colors)])
                        ))
                
                # Configurar layout do gráfico de barras empilhadas
                fig.update_layout(
                    title='Evolução das Categorias por Mês (2023-2024)',
                    xaxis_title='Mês/Ano',
                    yaxis_title='Quantidade',
                    barmode='stack',  # Empilha as barras
                    xaxis=dict(
                        tickmode='array',
                        tickvals=[f"{month:02d}/2023" for month in range(1, 13)] + [f"{month:02d}/2024" for month in range(1, 13)],
                        ticktext=['janeiro 2023', 'fevereiro 2023', 'março 2023', 'abril 2023', 'maio 2023', 'junho 2023',
                                'julho 2023', 'agosto 2023', 'setembro 2023', 'outubro 2023', 'novembro 2023', 'dezembro 2023',
                                'janeiro 2024', 'fevereiro 2024', 'março 2024', 'abril 2024', 'maio 2024', 'junho 2024',
                                'julho 2024', 'agosto 2024', 'setembro 2024', 'outubro 2024', 'novembro 2024', 'dezembro 2024']
                    ),
                    legend_title='Categoria e Ano'
                )

                st.plotly_chart(fig)
                
                total_ano_categoria = df_melted.groupby(['Ano', 'Categoria'])['Quantidade'].sum().unstack().fillna(0)

                # Criar gráfico de barras (Totais anuais por categoria)
                fig_totais_categoria = go.Figure()

                # Adicionar traços para cada categoria
                for categoria in selected_categories:
                    if categoria in total_ano_categoria.columns:
                        fig_totais_categoria.add_trace(go.Bar(
                            x=total_ano_categoria.index,
                            y=total_ano_categoria[categoria],
                            name=categoria,
                            text=total_ano_categoria[categoria],
                            textposition='auto'
                        ))

                # Configurar layout do gráfico de totais anuais por categoria
                fig_totais_categoria.update_layout(
                    title='Comparação Total por Categoria (2023 vs 2024)',
                    xaxis_title='Ano',
                    yaxis_title='Quantidade Total',
                    barmode='group',
                    legend_title='Categoria'
                )

                st.plotly_chart(fig_totais_categoria)
            else:
                st.info('Por favor, selecione pelo menos uma categoria.')
else:
         st.warning('Por favor, faça o upload de um arquivo Excel.')
