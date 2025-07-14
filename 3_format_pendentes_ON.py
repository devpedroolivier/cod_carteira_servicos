### IMPORTING LIBRARIES ###

import os
import shutil
import pandas as pd
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.worksheet.table import Table, TableStyleInfo
import logging
import tempfile

# Configuração de Log - Esse é o gatilho para envio pelo Power Automate
log_folder = r"C:\Users\poliveira.eficien.SBSP\Desktop\cod_carteira_servicos\carteira_servicos_MBP\log"
error_folder = r"C:\Users\poliveira.eficien.SBSP\Desktop\cod_carteira_servicos\carteira_servicos_MBP\Perror"

# Create the folders if they do not exist
os.makedirs(log_folder, exist_ok=True)
os.makedirs(error_folder, exist_ok=True)

log_file = os.path.join(log_folder, "pendentes.log")
error_log_file = os.path.join(error_folder, "error.log")

try:
    # Configuring the logging for successful execution
    logging.basicConfig(filename=log_file, level=logging.INFO,
                        format='%(asctime)s - %(levelname)s - %(message)s')

    ### READING THE PENDING FILES ###

    # Set the path where your files are located
    path = r"C:\Users\poliveira.eficien.SBSP\Desktop\cod_carteira_servicos\pendentes"

    # Create an empty list to hold the dataframes
    dfs = []

    # Mapping of ATC values to Polo
    atc_to_polo = {
        # 'Bragança': ['225', '403', '472', '516', '531', '534', '668', '746'],
        'Freguesia': ['922'],
        'Gopoúva': ['70'],
        'Pimentas': ['71'],
        'Extremo Norte': ['239', '241', '311', '312', '433'],
        'Pirituba': ['918'],
        'Santana': ['916', '927']
    }
    # Bragança foi removida da ON.

    # Function to determine Polo based on ATC value
    def determine_polo(atc_value):
        for polo, atc_values in atc_to_polo.items():
            if atc_value in atc_values:
                return polo
        return None  # Return None if no match is found


    # Iterate over each file in the directory
    for filename in os.listdir(path):
        # Check if the file is an .xlsx file
        if filename.endswith('.xlsx'):
            # Full path to the file
            file_path = os.path.join(path, filename)

            # Read the Excel file into a pandas DataFrame
            df = pd.read_excel(file_path, dtype=object)

            # Add a new column to the DataFrame with 'Polo'
            df['Polo'] = df['ATC'].astype(str).apply(determine_polo)

            # Append the DataFrame to the list
            dfs.append(df)

    # Concatenate all dataframes in the list into a single dataframe
    df = pd.concat(dfs, ignore_index=True)

    ### FILTERING ###

    # Filtrar as seguintes famílias:
    filtro_familia = ['ABASTECIMENTO', 'CAVALETE', 'RAMAL ESGOTO', 'REDE DE ESGOTO',
                    'REPOSIÇÃO', 'VAZAMENTO DE AGUA', 'VAZAMENTO DE ÁGUA', 'RAMAL DE AGUA', 'RAMAL DE ÁGUA',
                    'SERV COMPLEMENTAR', 'QUALIDADE DE ÁGUA', 'QUALIDADE DE AGUA', 'OUTROS SERVIÇOS DE REPOSIÇÃO',
                    'DESOBSTRUÇÃO', 'CONSERTO DE ESGOTO', 'REDE DE ÁGUA', 'REDE DE AGUA']
    # Adicionado 'OUTROS SERVIÇOS DE REPOSIÇÃO', 'DESOBSTRUÇÃO', 'CONSERTO DE ESGOTO'.
    # Adicionado 'REDE DE ÁGUA' para classificar os serviços das Equipes de VRP - ONOA.
    df = df[df['Família'].isin(filtro_familia)]

    # Planejável, Planejada, Suspensa, Reprogramável, Validada, Despachada, OS Admitida, Em Execução
    filtro_status_op = ['Planejável', 'Planejada', 'Validada', 'Reprogramável',
                        'Despachada', 'Suspensa', 'Em Execução', 'OS Admitida']
    df = df[df['Status da Operação'].isin(filtro_status_op)]

    ### EXTRACT TIME ###
    # Extract time from 'Data de Competência' and 'Prazo de Execução - Início', 
    # Compare both values 
    # If is different then use 'Data Inserção' to calculate df['Horas']

    # Extract time from 'Data de Competência'
    df['tempoCompetencia'] = pd.to_datetime(df['Data de Competência'], format='%d/%m/%Y %H:%M').dt.strftime('%H:%M')
    df['prazoExecInicio'] = pd.to_datetime(df['Prazo de Execução - Início'], format='%d/%m/%Y %H:%M').dt.strftime('%H:%M')


    ### CALCULATING TIME ###

    # Convert 'Data de Competência' to datetime and then localize to Brasília's timezone
    df['Data de Competência'] = pd.to_datetime(df['Data de Competência'], format='%d/%m/%Y %H:%M').dt.tz_localize(
        'America/Sao_Paulo')

    # Convert 'Data Inserção'
    df['Data Inserção'] = pd.to_datetime(df['Data Inserção'], format='%d/%m/%Y %H:%M').dt.tz_localize(
        'America/Sao_Paulo')

    # Get the current time in Brasília's timezone
    now_in_brasilia = pd.Timestamp.now(tz='America/Sao_Paulo')

    # Creating a function to calculate the time diff
    def calculate_hours(row):
        if row['tempoCompetencia'] == row['prazoExecInicio']:
            return round((now_in_brasilia - row['Data de Competência']).total_seconds() / 3600, 2)
        else:
            return round((now_in_brasilia - row['Data Inserção']).total_seconds() / 3600, 2)

    # Apply the function to calculate the 'Horas' column
    df['Horas'] = df.apply(calculate_hours, axis=1)

    # Drop the columns 'tempoCompetencia' and 'prazoExecInicio'
    df = df.drop(columns=['tempoCompetencia', 'prazoExecInicio'])

    # Função para classificar os grupos de carteira:
    def classify(row):
        if row['TSS'] in ['CARRO TANQUE GRATUITO', 'CARRO TANQUE VENDA', 'FALTA DE ÁGUA GERAL',
                        'FECHAR VÁLVULA DE REDE ÁGUA', 'FECHAR VÁLVULA DE REDE DE ÁGUA P/TESTE',
                        'FECHAMENTO RISCO DE SINISTRO', 'MUITA PRESSÃO DE ÁGUA',
                        'POUCA PRESSÃO DE ÁGUA GERAL',
                        'RETOMAR FECHAMENTO DE VÁLVULA', 'MANUTENÇÃO EM INSTALAÇÕES DE ÁGUA SABESP']:
            return 'ABASTECIMENTO'
        # 'FALTA DE ÁGUA LOCAL' e 'POUCA PRESSÃO DE ÁGUA LOCAL' movido para 'CAVALETE'
        # 'QUALIDADE DE ÁGUA - CHEIRO OP' movido para 'OUTROS', 'QUALIDADE DE ÁGUA - COR OP' movido para 'OUTROS',
        # 'QUALIDADE DE ÁGUA - GOSTO OP' movido para 'OUTROS'
        elif row['TSS'] in ['ABRIR VALVULA DE REDE DE AGUA', 'ABRIR VALVULA DE REDE DE AGUA PARA TESTE',
                        'CONSERTAR VÁLVULA REDUTORA DE PRESSÃO', 'INSTALAR VÁLVULA REDUTORA DE PRESSÃO',
                        'RETIRAR VÁLVULA REDUTORA DE PRESSÃO', 'TROCAR VÁLVULA REDUTORA DE PRESSÃO',
                        'RETIRAR LOGGER', 'INSTALAR LOGGER']:
            return 'EQUIPES VRP'
        # Criado categoria nova para identificar equipes de VRP - ONOA.
        elif row['TSS'] in ['CAVALETE QUEBRADO', 'CAVALETE VAZANDO', 'REGISTRO DO CAVALETE INVERTIDO',
                            'REGISTRO DO CAVALETE QUEBRADO', 'REGISTRO DO CAVALETE VAZANDO', 'FALTA DE ÁGUA LOCAL',
                            'POUCA PRESSÃO DE ÁGUA LOCAL']:
            return 'CAVALETE'
        elif row['TSS'] in ['ARREBENTADO DE REDE DE ÁGUA', 'HIDRANTE VAZANDO', 'VAZAMENTO DE ÁGUA COM INFILTRAÇÃO',
                            'VAZAMENTO DE ÁGUA EM RAMAL ABANDONADO', 'VAZAMENTO DE ÁGUA LEITO PAVIMENTADO',
                            'VAZAMENTO DE ÁGUA LEITO TERRA', 'VAZAMENTO DE ÁGUA NA VÁLVULA',
                            'VAZAMENTO DE ÁGUA NO PASSEIO']:
            return 'VAZAMENTO DE ÁGUA'
        elif row['TSS'] in ['TROCAR RAMAL DE ÁGUA - VAZ NÃO VISIVEL', 'VAZAMENTO DE ÁGUA NÃO VISÍVEL CAVALETE',
                            'VAZAMENTO DE ÁGUA NÃO VISÍVEL RAMAL', 'VAZAMENTO DE ÁGUA NÃO VISÍVEL REDE']:
            return 'VAZAMENTO NÃO VISÍVEL'
        elif row['TSS'] in ['DESOBSTRUIR RAMAL DE ESGOTO', 'DESOBSTRUIR RETORNO PARA IMOVEL',
                            'DESOBSTRUIR REDE DE ESGOTO']:
            return 'DESOBSTRUÇÕES'
        elif row['TSS'] in ['CONSERTAR RAMAL DE ESGOTO', 'TROCAR RAMAL DE ESGOTO']:
            return 'CRE; TRE'
        elif row['TSS'] in ['CONSERTAR REDE DE ESGOTO']:
            return 'CONSERTO DE COLETOR'
        elif row['TSS'] in ['REPOR ASFALTO A FRIO', 'REPOR ASFALTO A FRIO INV', 'REPOR PASSEIO ADJACENTE CIMENTADO',
                            'REPOR PASSEIO ADJACENTE CIMENTADO INV', 'REPOR PASSEIO OPOSTO CIMENTADO',
                            'REPOR PASSEIO OPOSTO CIMENTADO INV', 'REPOR PISO INTERNO CIMENTADO',
                            'REPOR PISO INTERNO CIMENTADO INV', 'RETIRAR ENTULHO']:
            return 'PA CIM; PI CIM; RET ENTULHO; BASE P/CAPA (12H)'
        elif row['TSS'] in ['REPOR PASSEIO ADJACENTE ESPECIAL', 'REPOR PASSEIO ADJACENTE ESPECIAL INV',
                            'REPOR PASSEIO OPOSTO ESPECIAL', 'REPOR PASSEIO OPOSTO ESPECIAL INV',
                            'REPOR PISO INTERNO ESPECIAL', 'REPOR PISO INTERNO ESPECIAL INV']:
            return 'PA ESP; PI ESP'
        elif row['TSS'] in ['REPOR BLOQUETE INV', 'REPOR BLOQUETE', 'REPOR CONCRETO', 'REPOR GRAMA', 'REPOR GRAMA INV',
                            'REPOR CONCRETO INV', 'REPOR GUIA', 'REPOR PARALELO', 'REPOR PARALELO INV',
                            'REPOR PAREDE/MURO',
                            'REPOR PAREDE/MURO INV', 'REPOR GUIA INV', 'REPOR SARJETA', 'REPOR SARJETA INV',
                            'TROCAR SOLO',
                            'TROCAR SOLO INV']:
            return 'GUIA; SARJETA; CONCRETO; BLOQUETE; MURO; GRAMA; PARALELO; TROCA DE SOLO'
        elif row['TSS'] in ['RECOMPOR SINALIZAÇÃO VIARIA HORIZONTAL', 'RECOMPOR SINALIZAÇÃO VIARIA HORIZONT INV',
                            'REPOR ASFALTO', 'REPOR ASFALTO INV', 'REPOR CAPA ASFALTICA',
                            'REPOR CAPA ASFALTICA ECOLOGICA',
                            'REPOR CAPA ASFALTICA INV', 'REPOR CAPA ASFALTICA ECOLOGICA INV',
                            'REPOR ASFALTO ECOLOGICO INV']:
            return 'CAPA ASF; CAPA ASF ECOL; ASFALTO; ASFALTO ECOL; SINALIZAÇÃO'
        elif row['TSS'] in ['ATERRAR VALA', 'ATERRAR VALA INV']:
            return 'ATERROS DE VALA'
        else:
            return 'OUTROS'  # define a default value for when none of the conditions is met

    # Create a column named 'tipo' to apply the function above
    df['tipo'] = df.apply(classify, axis=1)

    ### DATA FRAME ONOA  - ABASTECIMENTO ###

    # Define all Polos and required TSS categories
    all_polos = ['Extremo Norte', 'Freguesia', 'Santana', 'Pirituba', 'Pimentas', 'Gopoúva']
    # Removido 'Bragança'
    all_tss_categories = [
    #    'ABRIR VÁLVULA DE REDE DE ÁGUA',
    #    'ABRIR VALVULA DE REDE DE AGUA',
    #    'ABRIR VÁLVULA DE REDE DE ÁGUA PARA TESTE',
    #    'ABRIR VALVULA DE REDE DE AGUA PARA TESTE'
        'FALTA DE ÁGUA GERAL',
        'FECHAR VÁLVULA DE REDE ÁGUA',
        'FECHAR VÁLVULA DE REDE ÁGUA P/TESTE',
        'FECHAMENTO RISCO DE SINISTRO',
        'MUITA PRESSÃO DE ÁGUA',
        'POUCA PRESSÃO DE ÁGUA GERAL',
        'RETOMAR FECHAMENTO DE VÁLVULA',
        'QUALIDADE DE ÁGUA - CHEIRO OP',
        'QUALIDADE DE ÁGUA - COR OP',
        'QUALIDADE DE ÁGUA - GOSTO OP'
    ]

    # Filter for 'ABASTECIMENTO' and required TSS categories
    df_filtered_abastecimento = df[(df['tipo'] == 'ABASTECIMENTO') & (df['TSS'].isin(all_tss_categories))]

    # Create a temporary column for categorizing 'Horas'
    df_filtered_abastecimento['Categoria'] = pd.cut(
        df_filtered_abastecimento['Horas'],
        bins=[-1, 12, 24, 48, 72, 96, 120, 144, 168, 192, 216, 240, float('inf')],
        labels=['Até 12h', 'Até 24h', 'Até 48h', 'Até 72h', 'Até 96h', 'Até 120h',
                'Até 144h', 'Até 168h', 'Até 192h', 'Até 216h', 'Até 240h', '> 240h']
    )

    # Determine if each row is on time or late using the categorized column
    on_time_categories = ['Até 12h', 'Até 24h']  # Categories that count as on-time
    df_filtered_abastecimento['Prazo'] = df_filtered_abastecimento['Categoria'].apply(
        lambda x: 'Dentro do prazo' if x in on_time_categories else 'Fora do prazo'
    )

    # Create pivot table using the categorized 'Categoria' column
    df_pivot = pd.pivot_table(
        df_filtered_abastecimento,
        index='Polo',
        columns='Categoria',
        aggfunc='size',
        fill_value=0
    )

    # Ensure all expected columns are present
    required_columns = ['Até 12h', 'Até 24h', 'Até 48h', 'Até 72h', 'Até 96h', 'Até 120h',
                        'Até 144h', 'Até 168h', 'Até 192h', 'Até 216h', 'Até 240h', '> 240h']
    for column in required_columns:
        if column not in df_pivot.columns:
            df_pivot[column] = 0

    # Reindex to ensure all Polos are present (add missing Polos with 0 values)
    df_pivot = df_pivot.reindex(all_polos, fill_value=0)

    # Calculate "No Prazo" based on rules
    df_pivot['No Prazo'] = df_pivot[['Até 12h', 'Até 24h']].sum(axis=1)

    # Add "Total" column
    df_pivot['Total'] = df_pivot[required_columns].sum(axis=1)

    # Calculate "Taxa" (percentage of on-time tasks)
    df_pivot['Taxa'] = (df_pivot['No Prazo'] / df_pivot['Total']) * 100

    # Handle cases where 'Total' is 0 by setting 'Taxa' to 100%
    df_pivot.loc[df_pivot['Total'] == 0, 'Taxa'] = 100.0

    # Format "Taxa" as a percentage string
    df_pivot['Taxa'] = df_pivot['Taxa'].round(2).astype(str) + '%'

    # Reorder columns for clarity
    df_pivot = df_pivot[required_columns + ['Total', 'No Prazo', 'Taxa']]

    ### DATA FRAME ONOA - BY-PASS DE VRPS ###


    # Define all Polos and required TSS categories
    all_polos_bypass = ['Extremo Norte', 'Freguesia', 'Santana', 'Pirituba', 'Pimentas', 'Gopoúva']
    all_tss_categories_bypass = [
        'ABRIR VALVULA DE REDE DE AGUA', 'ABRIR VALVULA DE REDE DE AGUA PARA TESTE',
        'CONSERTAR VÁLVULA REDUTORA DE PRESSÃO', 'INSTALAR VÁLVULA REDUTORA DE PRESSÃO',
        'RETIRAR VÁLVULA REDUTORA DE PRESSÃO', 'TROCAR VÁLVULA REDUTORA DE PRESSÃO',
        'RETIRAR LOGGER', 'INSTALAR LOGGER'
    ]

    # Filter for 'ABASTECIMENTO' and required TSS categories
    df_filtered_bypass = df[(df['tipo'] == 'EQUIPES VRP') & (df['TSS'].isin(all_tss_categories_bypass))]

    # Create a temporary column for categorizing 'Horas'
    df_filtered_bypass['Categoria'] = pd.cut(
        df_filtered_bypass['Horas'],
        bins=[-1, 12, 24, 48, 72, 96, 120, 144, 168, 192, 216, 240, float('inf')],
        labels=['Até 12h', 'Até 24h', 'Até 48h', 'Até 72h', 'Até 96h', 'Até 120h',
                'Até 144h', 'Até 168h', 'Até 192h', 'Até 216h', 'Até 240h', '> 240h']
    )

    # Determine if each row is on time or late using the categorized column
    on_time_categories_bypass = ['Até 12h', 'Até 24h']
    df_filtered_bypass['Prazo'] = df_filtered_bypass['Categoria'].apply(
        lambda x: 'Dentro do prazo' if x in on_time_categories_bypass else 'Fora do prazo'
    )

    # Create pivot table using the categorized 'Categoria' column
    df_pivot_bypass = pd.pivot_table(
        df_filtered_bypass,
        index='Polo',
        columns='Categoria',
        aggfunc='size',
        fill_value=0
    )

    # Ensure all expected columns are present
    required_columns_bypass = ['Até 12h', 'Até 24h', 'Até 48h', 'Até 72h', 'Até 96h', 'Até 120h',
                        'Até 144h', 'Até 168h', 'Até 192h', 'Até 216h', 'Até 240h', '> 240h']
    for column in required_columns_bypass:
        if column not in df_pivot_bypass.columns:
            df_pivot_bypass[column] = 0

    # Reindex to ensure all Polos are present (add missing Polos with 0 values)
    df_pivot_bypass = df_pivot_bypass.reindex(all_polos_bypass, fill_value=0)

    # Calculate "No Prazo" based on rules
    df_pivot_bypass['No Prazo'] = df_pivot_bypass[['Até 12h', 'Até 24h']].sum(axis=1)

    # Add "Total" column
    df_pivot_bypass['Total'] = df_pivot_bypass[required_columns_bypass].sum(axis=1)

    # Calculate "Taxa" (percentage of on-time tasks)
    df_pivot_bypass['Taxa'] = (df_pivot_bypass['No Prazo'] / df_pivot_bypass['Total']) * 100

    # Handle cases where 'Total' is 0 by setting 'Taxa' to 100%
    df_pivot_bypass.loc[df_pivot_bypass['Total'] == 0, 'Taxa'] = 100.0

    # Format "Taxa" as a percentage string
    df_pivot_bypass['Taxa'] = df_pivot_bypass['Taxa'].round(2).astype(str) + '%'

    # Reorder columns for clarity
    df_pivot_bypass = df_pivot_bypass[required_columns_bypass + ['Total', 'No Prazo', 'Taxa']]

    # Create the "Equipes VRP" subtotal row
    subtotal = df_pivot_bypass.sum()
    subtotal['Taxa'] = (subtotal['No Prazo'] / subtotal['Total']) * 100
    subtotal['Taxa'] = f"{subtotal['Taxa']:.2f}%"

    # Append the subtotal row
    df_pivot_bypass.loc['Equipes VRP'] = subtotal

    ### ACOPLANDO LINHA EQUIPES VRP AO DATAFRAME ABASTECIMENTO ###

    # Extract the "Equipes VRP" row from df_pivot_bypass
    equipes_vrp_row = df_pivot_bypass.loc['Equipes VRP']

    # Convert the row to a DataFrame
    equipes_vrp_df = equipes_vrp_row.to_frame().T

    # Append the "Equipes VRP" row to df_pivot using pd.concat
    df_pivot = pd.concat([df_pivot, equipes_vrp_df])

    ### CONCATENANDO DATAFRAME ABASTECIMENTO + EQUIPES VRP

    # Concatenate the two dataframes
    df_filtered_abastecimento = pd.concat([df_filtered_abastecimento, df_filtered_bypass])

    ### TABELA ABASTECIMENTO FORA DO PRAZO ### - N8N
    # OS, TSS, Endereço, Horas

    df_n8n_abastecimento = df_filtered_abastecimento.copy()

    # Filtering orders expired
    df_n8n_abastecimento = df_n8n_abastecimento[df_n8n_abastecimento['Prazo'] == 'Fora do prazo']

    # Extracting only the desired columns
    df_n8n_abastecimento = df_n8n_abastecimento.drop(['Unidade Executante', 'Etapa', 'Prioridade', 'Fornecimento - Titular', 'PDE',
                                                 'Município', 'Número', 'Complemento', 'Bairro', 'ATC', 'SF', 'ATO', 'AS',
                                                 'Iniciativa', 'Data Inserção', 'Data de Competência', 'Data Agendada',
                                                'Prazo de Execução - Início', 'Prazo de Execução - Fim', 'Tempo Residual', 
                                                'Status da OS', 'Status da Operação', 'Data Fim Execução', 'Família', 'Causa Resultado',
                                               'Formulário de Baixa', 'Fornecedor', 'Contrato', 'Notificação de Esgoto', 'Categoria'], axis=1)

    # Exporting as .xlsx
    df_n8n_abastecimento.to_excel(r"C:\Users\poliveira.eficien.SBSP\Desktop\cod_carteira_servicos\carteira_servicos_MBP\n8n\abastecimento_fora_do_prazo.xlsx", index=False)


    ### HIGHLIGHT ONOA ###

    # Define the cells to highlight (first two columns of the pivot table)
    def generate_highlight_cells(df):
        # Get the dimensions of the DataFrame
        rows, cols = df.shape

        # Highlight first two columns for all rows
        highlight_cells = [(row, col) for row in range(rows) for col in range(2)]

        return highlight_cells


    # Define the highlighting function
    def apply_highlight(x):
        # Highlighting color for the first two columns
        color = 'background-color: #CEF2D3'

        # Initialize a DataFrame with the same shape as `x`, filled with empty strings
        df_styler = pd.DataFrame('', index=x.index, columns=x.columns)

        # Get the cells to highlight
        highlight_cells = generate_highlight_cells(x)

        # Apply the highlight to the specified cells
        for cell in highlight_cells:
            if 0 <= cell[0] < len(df_styler) and 0 <= cell[1] < len(df_styler.columns):  # Ensure cell indices are valid
                df_styler.iloc[cell] = color
        return df_styler


    # Apply the highlighting and export as HTML
    styled_df = df_pivot.style.apply(apply_highlight, axis=None)

    ### EXPORTANDO ONOA ###

    # Define export paths
    html_output_path = r"C:\Users\poliveira.eficien.SBSP\Desktop\cod_carteira_servicos\carteira_servicos_MBP\styled_htmls\ONOA.txt"
    excel_output_path = r"C:\Users\poliveira.eficien.SBSP\Desktop\cod_carteira_servicos\carteira_servicos_MBP\temp_file\df_abastecimento.xlsx"

    # Ensure directories exist
    os.makedirs(os.path.dirname(html_output_path), exist_ok=True)
    os.makedirs(os.path.dirname(excel_output_path), exist_ok=True)

    # Make datetime columns timezone-unaware
    datetime_columns = [
        'Data Inserção', 'Data de Competência', 'Data Agendada',
        'Prazo de Execução - Início', 'Prazo de Execução - Fim', 'Data Fim Execução'
    ]
    for col in datetime_columns:
        if col in df_filtered_abastecimento.columns:
            df_filtered_abastecimento[col] = pd.to_datetime(df_filtered_abastecimento[col], errors='coerce').dt.tz_localize(None)

    # Export pivot table as an HTML file
    with open(html_output_path, "w", encoding="utf-8") as f:
        f.write(styled_df.to_html())

    # Export filtered data as an Excel file
    df_filtered_abastecimento.to_excel(excel_output_path, index=False, engine="openpyxl")

    print(f"HTML ONOA enviado para: {html_output_path}")
    print(f"Anexo ONOA enviado para: {excel_output_path}")

    ## DEMAIS CARTEIRAS ###

    def create_pivot(df_subset):
        # Create pivot table
        df_pivot = pd.pivot_table(df_subset, index='tipo', columns='Horas', aggfunc='size', fill_value=0)

        # Add missing categories
        all_types = ['CAVALETE', 'VAZAMENTO DE ÁGUA', 'VAZAMENTO NÃO VISÍVEL', 'DESOBSTRUÇÕES',
                    'CRE; TRE', 'CONSERTO DE COLETOR', 'PA CIM; PI CIM; RET ENTULHO; BASE P/CAPA (12H)',
                    'PA ESP; PI ESP', 'GUIA; SARJETA; CONCRETO; BLOQUETE; MURO; GRAMA; PARALELO; TROCA DE SOLO',
                    'CAPA ASF; CAPA ASF ECOL; ASFALTO; ASFALTO ECOL; SINALIZAÇÃO', 'ATERROS DE VALA']
        for category in all_types:
            if category not in df_pivot.index:
                df_pivot.loc[category, :] = 0

        # Reorder the pivot table to follow all_types list order
        df_pivot = df_pivot.reindex(all_types)

        # Make sure required columns are present
        required_columns = ['Até 12h', 'Até 24h', 'Até 48h', 'Até 72h', 'Até 96h', 'Até 120h',
                            'Até 144h', 'Até 168h', 'Até 192h', 'Até 216h', 'Até 240h', '> 240h']

        for column in required_columns:
            if column not in df_pivot.columns:
                df_pivot[column] = 0

        # Convert all values to integer
        df_pivot = df_pivot.astype(int)

        # Remove 'OUTROS' category
        if 'OUTROS' in df_pivot.index:
            df_pivot = df_pivot.drop('OUTROS')

        # Pivot table considering the required columns and add total
        df_pivot = df_pivot[required_columns]
        df_pivot['Total'] = df_pivot.sum(axis=1)
        df_pivot.loc['Subtotal'] = df_pivot.sum()

        return df_pivot


    # Create a dictionary to hold dataframes
    dfs = {}

    # Iterate over each unique value in the 'Unidade Executante' column
    for unidade in df['Unidade Executante'].unique():
        # Create a subset dataframe for the current 'Unidade Executante'
        df_subset = df[df['Unidade Executante'] == unidade]

        # Add a new column 'Categoria' with the classification based on 'Horas'
        df_subset['Categoria'] = pd.cut(df_subset['Horas'],
                                        bins=[-1, 12, 24, 48, 72, 96, 120, 144, 168, 192, 216, 240, float('inf')],
                                        labels=['Até 12h', 'Até 24h', 'Até 48h', 'Até 72h', 'Até 96h', 'Até 120h',
                                                'Até 144h', 'Até 168h', 'Até 192h', 'Até 216h', 'Até 240h', '> 240h']
                                        )

        # Add the categorized column back to the original DataFrame
        df.loc[df_subset.index, 'Categoria'] = df_subset['Categoria']

        # Group the 'Horas' column (this will overwrite 'Horas' with categorized values)
        df_subset['Horas'] = pd.cut(df_subset['Horas'],
                                    bins=[-1, 12, 24, 48, 72, 96, 120, 144, 168, 192, 216, 240, float('inf')],
                                    labels=['Até 12h', 'Até 24h', 'Até 48h', 'Até 72h', 'Até 96h', 'Até 120h',
                                            'Até 144h', 'Até 168h', 'Até 192h', 'Até 216h', 'Até 240h', '> 240h']
                                    )

        # Check if the 'Unidade Executante' is 'DIV MANUT SERV OPE':
        if 'DIV MANUT SERV OPE' in unidade:
            unit_name = unidade.split('DIV MANUT SERV OPE')[1].strip().lower().replace(' ', '_')
        elif 'ONOA - DIV OPERAÇÃO DE ÁGUA NORTE' in unidade:
            unit_name = unidade.split('DIV')[1].strip().lower().replace(' ', '_')
        else:
            unit_name = unidade.lower().replace(' ', '_')

        # Create pivot table for the entire subset and save to dictionary with the unit name as the key
        dfs[unit_name] = create_pivot(df_subset)

    # Columns with date/datetime values
    columns_to_change = ['Data Inserção', 'Data de Competência', 'Prazo de Execução - Início',
                        'Prazo de Execução - Fim', 'Data Fim Execução']

    # Ensure columns are properly converted to datetime
    for column in columns_to_change:
        if column in df.columns:  # Check if the column exists
            # Convert to datetime, handle errors, and remove timezone if applicable
            df[column] = pd.to_datetime(df[column], errors='coerce')  # Convert to datetime
            if pd.api.types.is_datetime64_any_dtype(df[column]):
                df[column] = df[column].dt.tz_localize(None)  # Remove timezone only for datetime columns

    # File paths
    output_file = r"C:\Users\poliveira.eficien.SBSP\Desktop\cod_carteira_servicos\temp_file\carteira.xlsx"
    destination_file = r"C:\Users\poliveira.eficien.SBSP\Desktop\cod_carteira_servicos\carteira_servicos_MBP\temp_file\carteira.xlsx"

    # Export the dataframe to an Excel file
    df.to_excel(output_file, index=False)  # Save to the first location

    # Copy the file to the second location
    shutil.copy(output_file, destination_file)

    # Define the sum column for each tipo
    sum_columns = {
        "CAVALETE": ['Até 12h', 'Até 24h'],
        "VAZAMENTO DE ÁGUA": ['Até 12h', 'Até 24h'],
        "VAZAMENTO NÃO VISÍVEL": ['Até 12h', 'Até 24h', 'Até 48h', 'Até 72h', 'Até 96h'],
        "DESOBSTRUÇÕES": ['Até 12h', 'Até 24h'],
        "CRE; TRE": ['Até 12h', 'Até 24h', 'Até 48h'],
        "CONSERTO DE COLETOR": ['Até 12h', 'Até 24h', 'Até 48h', 'Até 72h', 'Até 96h'],
        "PA CIM; PI CIM; RET ENTULHO; BASE P/CAPA (12H)": ['Até 12h'],
        "PA ESP; PI ESP": ['Até 12h', 'Até 24h', 'Até 48h', 'Até 72h', 'Até 96h', 'Até 120h',
                        'Até 144h', 'Até 168h'],
        "GUIA; SARJETA; CONCRETO; BLOQUETE; MURO; GRAMA; PARALELO; TROCA DE SOLO": ['Até 12h', 'Até 24h', 'Até 48h',
                                                                                    'Até 72h'],
        "CAPA ASF; CAPA ASF ECOL; ASFALTO; ASFALTO ECOL; SINALIZAÇÃO": ['Até 12h', 'Até 24h', 'Até 48h', 'Até 72h'],
        "ATERROS DE VALA": ['Até 12h', 'Até 24h', 'Até 48h', 'Até 72h']
    }

    # Go through all dataframes in the dictionary
    for df_name, df in dfs.items():
        # Apply a function to calculate the sum based on the 'tipo' value
        df['No prazo'] = df.apply(lambda row: row[sum_columns[row.name]].sum() if row.name != 'Subtotal' else 0, axis=1)

        # Compute the subtotal for the 'No prazo' column
        df.loc['Subtotal', 'No prazo'] = df.loc[df.index != 'Subtotal', 'No prazo'].sum()

        # Create a 'Taxa' column
        df['Taxa'] = df['No prazo'] / df['Total'] * 100

        # Fill NaN values in the 'Taxa' column with 100.0
        df['Taxa'] = df['Taxa'].fillna(100.0)

        # Round the 'Taxa' column to two decimal places
        df['Taxa'] = df['Taxa'].round(2)

        # Convert 'Taxa' to a string format with two decimal places
        df['Taxa'] = df['Taxa'].map('{:.2f}'.format)

    # Export each dataframe in dfs to a csv file
    # Caminho base para salvar os CSVs dos polos
    csv_output_dir = os.path.join(
        "C:/Users/poliveira.eficien.SBSP/Desktop/cod_carteira_servicos",
        "carteira_servicos_MBP", "polos"
    )

    # Garante que a pasta existe
    os.makedirs(csv_output_dir, exist_ok=True)

    # Exporta cada DataFrame do dicionário para CSV
    for key, df_export in dfs.items():
        csv_path = os.path.join(csv_output_dir, f"{key}.csv")
        df_export.to_csv(csv_path, index=True, header=True)
        print(f"✅ CSV exportado: {csv_path}")


    # Define the cells to highlight
    highlight_cells = [(x, y) for x in range(0, 11) for y in [0]] + \
                    [(x, y) for x in range(0, 6) for y in [1]] + [(x, y) for x in range(7, 11) for y in [1]] + \
                    [(x, y) for x in [2, 4, 5, 7, 8, 9, 10] for y in [2]] + \
                    [(x, y) for x in [2, 5, 7, 8, 9, 10] for y in [3]] + \
                    [(x, y) for x in [2, 5, 7] for y in [4]] + \
                    [(x, y) for x in [7] for y in [5, 6, 7]]


    # Define the highlighting function
    def apply_highlight(x):
        color = 'background-color: #CEF2D3'
        df_styler = pd.DataFrame('', index=x.index, columns=x.columns)
        for cell in highlight_cells:
            df_styler.iloc[cell] = color
        return df_styler


    # Loop through each dataframe in the dictionary
    for name, df in dfs.items():
        # Replace problematic characters
        df.index = df.index.to_series().apply(
            lambda x: x.replace('Ç', '&Ccedil;').replace('Ã', '&Atilde;').replace('Á', '&Aacute;').replace('Í',
                                                                                                        '&Iacute;').replace(
                'Õ', '&Otilde;'))

        # Apply the highlighting function
        df = df.style.apply(apply_highlight, axis=None)

        # Export as HTML
        html = df.to_html()

        # Path to save the file
        # Path to save the file
        # Define o caminho para salvar o arquivo
        path = os.path.join(
            "C:/Users/poliveira.eficien.SBSP/Desktop/cod_carteira_servicos",
            "carteira_servicos_MBP", "styled_htmls", f"{name}.txt"
        )

        # Garante que a pasta existe
        os.makedirs(os.path.dirname(path), exist_ok=True)

        # Informa no console
        print(f"📄 HTML {name} sendo exportado para: {path}")

        # Salva o HTML estilizado no arquivo
        with open(path, "w", encoding="utf-8") as file:
            file.write(html)



        # Reading the base to get all orders expired
        df_n8n_demais_carteiras = pd.read_excel(destination_file)

        # Function to classify 'Prazo'
        def classify_prazo(tipo, categoria):
            if tipo in sum_columns and categoria in sum_columns[tipo]:
                return "Dentro do prazo"
            else:
                return "Fora do prazo"

        # Apply classification to the dataframe
        df_n8n_demais_carteiras["Prazo"] = df_n8n_demais_carteiras.apply(lambda row: classify_prazo(row["tipo"], row["Categoria"]), axis=1)

        # Filtering orders expired
        df_n8n_demais_carteiras = df_n8n_demais_carteiras[df_n8n_demais_carteiras['Prazo'] == 'Fora do prazo']

        # Filtering out 'Outros' tipo
        df_n8n_demais_carteiras = df_n8n_demais_carteiras[df_n8n_demais_carteiras["tipo"].isin(sum_columns.keys())]

        # Extracting only the desired columns
        df_n8n_demais_carteiras = df_n8n_demais_carteiras.drop(['Unidade Executante', 'Etapa', 'Prioridade', 'Fornecimento - Titular', 'PDE',
                                                 'Município', 'Número', 'Complemento', 'Bairro', 'ATC', 'SF', 'ATO', 'AS',
                                                 'Iniciativa', 'Data Inserção', 'Data de Competência', 'Data Agendada',
                                                'Prazo de Execução - Início', 'Prazo de Execução - Fim', 'Tempo Residual', 
                                                'Status da OS', 'Status da Operação', 'Data Fim Execução', 'Família', 'Causa Resultado',
                                               'Formulário de Baixa', 'Fornecedor', 'Contrato', 'Notificação de Esgoto', 'Categoria'], axis=1)
        
        # Exporting each table

        # Define the export folder
       # Pasta de exportação
        export_folder = r"C:\Users\poliveira.eficien.SBSP\Desktop\cod_carteira_servicos\carteira_servicos_MBP\n8n"
        os.makedirs(export_folder, exist_ok=True)

        # Verifica se há registros para exportar
        polos_disponiveis = df_n8n_demais_carteiras["Polo"].dropna().unique()

        if len(polos_disponiveis) == 0:
            print("⚠️ Nenhum polo encontrado com ordens fora do prazo para exportação.")
        else:
            for polo in polos_disponiveis:
                try:
                    polo_filename = polo.replace(" ", "_").lower()
                    file_path = os.path.join(export_folder, f"{polo_filename}_fora_do_prazo.xlsx")

                    df_polo = df_n8n_demais_carteiras[df_n8n_demais_carteiras["Polo"] == polo]

                    # Cria arquivo vazio se não houver registros, com as mesmas colunas
                    if df_polo.empty:
                        df_polo = pd.DataFrame(columns=df_n8n_demais_carteiras.columns)
                        print(f"⚠️ Polo '{polo}' não possui registros fora do prazo. Arquivo vazio será gerado.")

                    # Exporta
                    df_polo.to_excel(file_path, index=False)
                    
                    # Confirma exportação
                    if os.path.exists(file_path):
                        print(f"✅ Exportado: {file_path}")
                    else:
                        print(f"❌ ERRO: {file_path} não foi criado.")
                except Exception as e:
                    print(f"❌ ERRO ao exportar '{polo}': {str(e)}")


        print("Export completed successfully!")


    logging.info("Code executed successfully.")

except Exception as e:
    # If an error occurs, configure the error logger and log the error
    logging.basicConfig(filename=error_log_file, level=logging.ERROR,
                        format='%(asctime)s - %(levelname)s - %(message)s')

    logging.error(f"An error occurred: {str(e)}")


