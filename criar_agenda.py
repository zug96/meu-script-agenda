# -*- coding: utf-8 -*-
import pandas as pd
import numpy as np
import argparse
import sys
from ics import Calendar, Event
from datetime import datetime, timedelta, time
import pytz # Import para timezone
import sys

# --- Configurações --- (Manter igual)
CSV_FILENAME = 'agenda_data.csv'
SHEET_NAME = 'AgendaSemanal'

# --- Funções ---

# CORRIGIDA: Função carregar_dados com parse de datetime
def carregar_dados(csv_path):
    """Carrega os dados da agenda a partir de um arquivo CSV e tenta parsear horário."""
    print(f"Carregando dados de '{csv_path}'...")
    try:
        df = pd.read_csv(csv_path, encoding='utf-8')
        # Tratar valores NaN (células vazias) como strings vazias ANTES do loop
        df = df.fillna('')
        print("Dados carregados com sucesso.")
    except FileNotFoundError:
        print(f"Erro Crítico: Arquivo '{csv_path}' não encontrado.")
        print("Certifique-se de que o arquivo está na mesma pasta que o script.")
        sys.exit(1) # Sai do script com código de erro
    except Exception as e:
        print(f"Erro Crítico ao ler o arquivo CSV: {e}")
        sys.exit(1)

    # Tentar parsear o horário de início
    print("Tentando parsear horários de início...")
    start_times = []
    # Loop para processar cada horário no DataFrame
    for horario_str in df['Horario']:
        # --- Bloco try...except CORRETAMENTE INDENTADO ---
        try:
            # Tenta pegar a primeira parte 'hh:mm'
            start_str = horario_str.split('-')[0].strip()
            # Converte para um objeto time
            time_obj = datetime.strptime(start_str, '%H:%M').time()
            start_times.append(time_obj)
        except (ValueError, AttributeError, IndexError):
            # Se não for formato hh:mm - hh:mm, ou der erro, adiciona None
            start_times.append(None)
        # --- Fim do bloco try...except ---

    # --- Linhas EXECUTADAS APÓS o loop for terminar ---
    # Adiciona a nova coluna ao DataFrame (com indentação correta, fora do loop)
    df['StartTimeObject'] = start_times
    print("Tentativa de parse do horário de início concluída.")

    # Retorna o DataFrame modificado (com indentação correta, fora do loop)
    return df

# --- Função configurar_formatos --- (Manter igual a versão anterior)
def configurar_formatos(workbook):
    """Define e retorna um dicionário com os formatos de célula para o Excel."""
    print("Configurando formatos de célula...")
    # Formato do Cabeçalho
    header_format = workbook.add_format({
        'bold': True, 'text_wrap': True, 'valign': 'top',
        'fg_color': '#D7E4BC', 'border': 1, 'align': 'center'
    })
    # Formatos de Célula Base
    cell_format_base = workbook.add_format({'border': 1, 'text_wrap': True, 'valign': 'top'})
    cell_format_dia_hora = workbook.add_format({'border': 1, 'text_wrap': True, 'valign': 'top', 'align': 'center'})

    # Formatos de Cores por Atividade
    color_formats = {
        'Estudo Linux': workbook.add_format({'bg_color': '#FFFFCC', 'border': 1, 'text_wrap': True, 'valign': 'top'}),
        'AULA PEDAGOGIA': workbook.add_format({'bg_color': '#CCFFFF', 'border': 1, 'text_wrap': True, 'valign': 'top'}),
        'Estágio': workbook.add_format({'bg_color': '#CCFFFF', 'border': 1, 'text_wrap': True, 'valign': 'top'}),
        'VÔLEI': workbook.add_format({'bg_color': '#CCFFCC', 'border': 1, 'text_wrap': True, 'valign': 'top'}),
        'Caminhada': workbook.add_format({'bg_color': '#CCFFCC', 'border': 1, 'text_wrap': True, 'valign': 'top'}),
        'Afazeres': workbook.add_format({'bg_color': '#FFDDCC', 'border': 1, 'text_wrap': True, 'valign': 'top'}),
        'Flexível': workbook.add_format({'bg_color': '#E0E0E0', 'border': 1, 'text_wrap': True, 'valign': 'top'}),
        'Tempo Livre': workbook.add_format({'bg_color': '#F0F0F0', 'border': 1, 'text_wrap': True, 'valign': 'top'}),
        'Descanso': workbook.add_format({'bg_color': '#F0F0F0', 'border': 1, 'text_wrap': True, 'valign': 'top'}),
    }
    print("Formatos definidos.")
    return {
        'header': header_format,
        'base': cell_format_base,
        'dia_hora': cell_format_dia_hora,
        'colors': color_formats
    }

def aplicar_formatacao_e_escrever_dados(worksheet, workbook, df, formatos):
    """Aplica formatação (COM TÍTULO) e escreve os dados."""
    print("Aplicando formatação (com título) e escrevendo dados...")

    # --- Escrever Título ---
    titulo_format = workbook.add_format({'bold': True, 'font_size': 14, 'align': 'center', 'valign': 'vcenter'})
    # Mescla células da primeira linha para o título (A1 até D1)
    worksheet.merge_range('A1:D1', f"Agenda da Semana ({SHEET_NAME})", titulo_format)
    # Definir altura da primeira linha (opcional)
    worksheet.set_row(0, 30) # Linha 0 (primeira linha) com altura 30

    # --- Escrever Cabeçalho (AGORA NA LINHA 2, índice 1) ---
    for col_num, value in enumerate(df.columns.values):
        worksheet.write(1, col_num, value, formatos['header']) # Mudou para linha 1

    # --- Aplicar Formato e Cores às Células de Dados (A PARTIR DA LINHA 3, índice 2) ---
    for row_num in range(len(df)):
        # ... (lógica para determinar bg_color_to_apply continua igual) ...
         activity_text = str(df.iloc[row_num, 2]).lower()
         bg_color_to_apply = None
         for keyword, fmt in formatos['colors'].items():
              if hasattr(fmt, 'bg_color') and keyword.lower() in activity_text:
                  bg_color_to_apply = fmt.bg_color
                  break

         # Aplicar formato a cada célula da linha (AGORA EM row_num + 2)
         for col_num in range(len(df.columns)):
             cell_value = df.iloc[row_num, col_num]
             if col_num < 2:
                 final_props = {'border': 1, 'text_wrap': True, 'valign': 'top', 'align': 'center'}
             else:
                 final_props = {'border': 1, 'text_wrap': True, 'valign': 'top'}
             if bg_color_to_apply:
                 final_props['bg_color'] = bg_color_to_apply
             final_cell_format = workbook.add_format(final_props)
             worksheet.write(row_num + 2, col_num, cell_value, final_cell_format) # Mudou para row_num + 2

    # --- Ajustar Largura das Colunas ---
    worksheet.set_column('A:A', 20) # Dia
    worksheet.set_column('B:B', 18) # Horário
    worksheet.set_column('C:C', 60) # Atividade
    worksheet.set_column('D:D', 70) # Observações
    print("Largura das colunas (aumentada)  ajustada.")

    # --- Mesclar Células da Coluna 'Dia' (A PARTIR DA LINHA 3, índice 2) ---
    print("Mesclando células da coluna 'Dia'...")
    merge_start_row = 2 # Começa na linha 2 (onde estão os primeiros dados)
    for i in range(1, len(df)):
        current_day = df.iloc[i, 0]
        previous_day = df.iloc[i-1, 0]
        day_props = {'border': 1, 'text_wrap': True, 'valign': 'top', 'align': 'center'}

        if current_day == previous_day:
            if i == len(df) - 1 or df.iloc[i+1, 0] != current_day:
                if merge_start_row <= i + 1: # Ajuste na condição de mesclagem (+1 por causa do startrow)
                    # ... (lógica para determinar cor e criar final_merge_format igual antes) ...
                    activity_text_merge = str(df.iloc[merge_start_row - 2, 2]).lower() # Ajuste índice df
                    # ...(resto da lógica de cor)...
                    final_merge_format = workbook.add_format(day_props) # Recalcular com cor
                    worksheet.merge_range(merge_start_row, 0, i + 2, 0, current_day, final_merge_format) # Ajuste nos índices de linha (+1)
                merge_start_row = i + 3 # Ajuste (+1)
        else:
            if merge_start_row == i + 1: # Ajuste (+1)
                 # ... (lógica para determinar cor e criar final_prev_format igual antes) ...
                 activity_text_prev = str(df.iloc[i-1, 2]).lower()
                 # ...(resto da lógica de cor)...
                 final_prev_format = workbook.add_format(day_props) # Recalcular com cor
                 worksheet.write(i + 1, 0, previous_day, final_prev_format) # Ajuste linha (+1)
            merge_start_row = i + 2 # Ajuste (+1)

    if merge_start_row == len(df) + 1: # Ajuste (+1)
        # ... (lógica para determinar cor e criar final_last_format igual antes) ...
        activity_text_last = str(df.iloc[len(df)-1, 2]).lower()
         # ...(resto da lógica de cor)...
        final_last_format = workbook.add_format(day_props) # Recalcular com cor
        worksheet.write(len(df) + 1, 0, df.iloc[len(df)-1, 0], final_last_format) # Ajuste linha (+1)
    print("Mesclagem concluída.")

    # --- Congelar Painel Superior (AGORA CONGELA ABAIXO DO TÍTULO) ---
    worksheet.freeze_panes(2, 0) # Congela a partir da linha 2 (abaixo do cabeçalho)
    print("Painel abaixo do título congelado.")
# --- Função Principal ---
# Mantenha os imports e as definições de constantes/outras funções iguais

# --- Nova Função para Legenda ---
def escrever_legenda(worksheet, workbook, formatos, num_data_rows):
    """Escreve a legenda de cores abaixo dos dados da agenda."""
    print("Escrevendo legenda de cores...")
    # Determina a linha inicial para a legenda
    # Linha 0: Título; Linha 1: Cabeçalho; Linhas 2 a (num_data_rows + 1): Dados.
    # Última linha de dados tem índice num_data_rows + 1.
    # Deixaremos uma linha em branco (índice num_data_rows + 2).
    legend_start_row = num_data_rows + 3 # Legenda começa na linha com índice num_data_rows + 3

    # Formato para o título da legenda
    legend_title_format = workbook.add_format({'bold': True, 'font_size': 10})
    # Formato para o texto da legenda (alinhado ao topo se a célula for alta)
    legend_text_format = workbook.add_format({'valign': 'top'})

    # Escrever título da legenda (mesclado em 2 colunas, A e B)
    try:
         worksheet.merge_range(legend_start_row, 0, legend_start_row, 1, "Legenda de Cores:", legend_title_format)
    except Exception as merge_err:
         print(f"  Aviso: Não foi possível mesclar células para título da legenda - {merge_err}")
         worksheet.write(legend_start_row, 0, "Legenda de Cores:", legend_title_format)


    current_row = legend_start_row + 1
    processed_colors = {} # Dicionário para agrupar keywords pela cor

    # Ajustar largura das colunas específicas da legenda
    worksheet.set_column(1, 1, 35) # Coluna B (descrição) mais larga que antes

    # Agrupar keywords por cor para evitar repetição na legenda
    for keyword, fmt in formatos['colors'].items():
        if hasattr(fmt, 'bg_color'):
            color_hex = fmt.bg_color # Pega a cor definida no formato
            if color_hex not in processed_colors:
                processed_colors[color_hex] = {'format': fmt, 'keywords': []}
            processed_colors[color_hex]['keywords'].append(keyword)

    # Escrever cada item da legenda (cor + descrição agrupada)
    for color_hex, data in processed_colors.items():
        color_format_for_swatch = data['format']
        # Junta as keywords que usam a mesma cor
        description = " / ".join(data['keywords'])

        try:
            # Escreve uma célula em branco APENAS com a cor de fundo
            worksheet.write_blank(current_row, 0, None, color_format_for_swatch)
            # Escreve a descrição na coluna ao lado
            worksheet.write(current_row, 1, description, legend_text_format)
            current_row += 1
        except Exception as write_err:
             print(f"  Erro ao escrever linha da legenda para {description}: {write_err}")


    print("Legenda de cores escrita.")

# Garanta estes imports no topo do arquivo:
from ics import Calendar, Event
from datetime import datetime, timedelta, time
import pytz # Import para timezone
import sys

def gerar_ics(df, ics_filename="agenda.ics"):
    """Gera um arquivo .ics com timestamps Timezone-Aware (CORRIGIDO PARA TIMEZONE)."""
    print(f"Gerando arquivo de calendário '{ics_filename}' com Timezone...")
    c = Calendar()
    year = 2025 # Ano base

    # --- Definir o Timezone Correto ---
    try:
        timezone = pytz.timezone("America/Sao_Paulo")
        print(f"Usando Timezone: {timezone}")
    except pytz.UnknownTimeZoneError:
        print("Erro Crítico: Timezone 'America/Sao_Paulo' não encontrado pelo pytz.")
        print("Verifique se a biblioteca pytz está instalada e atualizada.")
        sys.exit(1)
    # ------------------------------------

    for index, row in df.iterrows():
        e = Event()
        e.name = row.get('Atividade', 'Evento Sem Nome')
        e.description = str(row.get('Observacoes', ''))

        try:
            # Extrair dia e mês (lógica anterior)
            dia_str_completo = str(row.get('Dia', ''))
            if '(' in dia_str_completo and '/' in dia_str_completo and ')' in dia_str_completo:
                date_part_str = dia_str_completo.split('(')[1].split(')')[0]
                day_str, month_str = date_part_str.split('/')
                day = int(day_str)
                month = int(month_str)
            else:
                raise ValueError(f"Formato de 'Dia' inesperado: {dia_str_completo}")

            start_time_obj = row.get('StartTimeObject')

            if isinstance(start_time_obj, time):
                # Cria datetime NAIVE primeiro
                naive_begin_dt = datetime(year, month, day, start_time_obj.hour, start_time_obj.minute)
                # TORNAR AWARE usando o timezone definido
                e.begin = timezone.localize(naive_begin_dt) # Associa o fuso de SP

                # Lógica para fim ou duração (com timezone)
                try:
                    horario_str_completo = str(row.get('Horario', ''))
                    if '-' in horario_str_completo:
                        end_str = horario_str_completo.split('-')[1].strip()
                        end_time_obj = datetime.strptime(end_str, '%H:%M').time()
                        naive_end_dt = datetime(year, month, day, end_time_obj.hour, end_time_obj.minute)
                        # Lidar com eventos que passam da meia-noite (simples)
                        if naive_end_dt.time() <= start_time_obj:
                             naive_end_dt += timedelta(days=1)
                        # TORNAR AWARE
                        e.end = timezone.localize(naive_end_dt) # Associa o fuso de SP ao fim
                    else:
                        e.duration = timedelta(hours=1) # Duração não precisa de fuso explícito
                except:
                    e.duration = timedelta(hours=1)

            else:
                # Evento do dia inteiro (DATE não tem fuso horário)
                date_obj = datetime(year, month, day).date()
                e.begin = date_obj
                e.make_all_day()

            c.events.add(e)

        # Captura erros durante o parsing ou criação da data/evento
        except (ValueError, IndexError, TypeError, AttributeError) as parse_error:
            print(f"  Aviso: Não foi possível processar evento para .ics na linha {index+1}: {row.get('Atividade','N/A')} ({parse_error})")
        except Exception as general_error:
             print(f"  Erro inesperado ao processar linha {index+1} para .ics: {general_error}")


    # Escrever no arquivo .ics (igual antes)
    try:
        # Usar 'with' garante fechamento correto do arquivo
        with open(ics_filename, 'w', encoding='utf-8') as f:
            f.writelines(c.serialize_iter())
        print("Arquivo .ics gerado com sucesso.")
    except Exception as write_error:
        print(f"  Erro ao escrever arquivo .ics: {write_error}")

def main():
    """Função principal que orquestra a criação da planilha."""
    print("--- Iniciando Geração da Agenda ---")

    # --- Configuração do Argparse ---
    parser = argparse.ArgumentParser(description='Gera uma planilha Excel formatada da agenda semanal a partir de um CSV.')
    parser.add_argument('-o', '--output',
                        default='agenda_semana_formatada.xlsx', # Nome padrão
                        help='Nome do arquivo Excel de saída (ex: minha_agenda.xlsx)')
    parser.add_argument('-c', '--csv',
                        default='agenda_data.csv', # Nome padrão do CSV
                        help='Nome do arquivo CSV de entrada (ex: dados.csv)')
    args = parser.parse_args()

    # Usa os nomes de arquivo dos argumentos
    csv_input_filename = args.csv
    excel_output_filename = args.output
    # --------------------------------

    # Carregar dados
    df_agenda = carregar_dados(csv_input_filename) # Usa o nome do CSV do argumento

    # Criar e usar o 'escritor' Excel com 'with' para garantir fechamento
    try:
        with pd.ExcelWriter(excel_output_filename, engine='xlsxwriter') as excel_writer:
            print(f"Escrevendo dados na planilha '{SHEET_NAME}'...")
            # Escrever DataFrame no Excel (isto cria a planilha com nome SHEET_NAME)
            df_agenda.to_excel(excel_writer, index=False, sheet_name=SHEET_NAME, startrow=2, header=False)

            # Obter objetos workbook e worksheet para formatação
            workbook = excel_writer.book

            # --- CORREÇÃO APLICADA AQUI ---
            # Em vez de usar excel_writer.sheets[SHEET_NAME], pegamos a primeira planilha da lista do workbook.
            # Isso é mais robusto caso o dicionário 'sheets' não esteja populado imediatamente.
            worksheets_list = workbook.worksheets()
            if worksheets_list:
                worksheet = worksheets_list[0]
                print(f"Planilha '{worksheet.name}' obtida para formatação.") # Verifica o nome obtido
            else:
                # Se df.to_excel falhou em criar a planilha (improvável sem erro anterior)
                print(f"Erro Crítico: Nenhuma planilha foi encontrada no workbook após a escrita inicial.")
                sys.exit(1)
            # --- FIM DA CORREÇÃO ---

            # Configurar os formatos
            formatos = configurar_formatos(workbook)

            # Aplicar formatação e escrever os dados formatados
            aplicar_formatacao_e_escrever_dados(worksheet, workbook, df_agenda, formatos)

            escrever_legenda(worksheet, workbook, formatos, len(df_agenda))

            # O 'with' cuida do save/close automaticamente ao sair do bloco
            print("Arquivo Excel sendo finalizado...")

        print(f"--- Geração da Agenda Concluída: '{excel_output_filename}' ---") # Mostra o nome usado

    except PermissionError as e:
         print(f"Erro de Permissão ao tentar escrever '{excel_output_filename}': {e}")
         print("Verifique se o arquivo não está aberto no Excel e se você tem permissão de escrita na pasta.")
         sys.exit(1)
    except KeyError as e:
         print(f"Erro Inesperado de Chave: {e}. A planilha '{SHEET_NAME}' não foi encontrada como esperado.")
         sys.exit(1)
    except Exception as e:
        print(f"Ocorreu um erro inesperado durante a escrita do Excel: {e}")
        sys.exit(1)

   # --- Gerar ICS ---
    gerar_ics(df_agenda, "minha_agenda.ics") # Chama a nova função

    print(f"--- Geração da Agenda Concluída ---")

# --- Ponto de Entrada --- (Manter if __name__ == "__main__": ...)
# Lembre-se de ter os imports necessários (pandas, sys, argparse, etc.)
if __name__ == "__main__":
    import pandas as pd
    import numpy as np
    import sys
    import argparse
    from datetime import datetime, timedelta # Se já estiver usando datetime
    from ics import Calendar, Event # Se já estiver usando ics
    main()