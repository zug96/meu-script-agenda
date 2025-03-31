# -*- coding: utf-8 -*-
import pandas as pd
import numpy as np
import argparse
import sys # Importar sys se ainda não estiver globalmente main()
import sys # Para usar sys.exit()

# --- Configurações ---
CSV_FILENAME = 'agenda_data.csv'
SHEET_NAME = 'AgendaSemanal'

# --- Funções ---

def carregar_dados(csv_path):
    """Carrega os dados da agenda a partir de um arquivo CSV."""
    print(f"Carregando dados de '{csv_path}'...")
    try:
        df = pd.read_csv(csv_path, encoding='utf-8')
        # Tratar valores NaN (células vazias) como strings vazias
        df = df.fillna('')
        print("Dados carregados com sucesso.")
        return df
    except FileNotFoundError:
        print(f"Erro Crítico: Arquivo '{csv_path}' não encontrado.")
        print("Certifique-se de que o arquivo está na mesma pasta que o script.")
        sys.exit(1) # Sai do script com código de erro
    except Exception as e:
        print(f"Erro Crítico ao ler o arquivo CSV: {e}")
        sys.exit(1)

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

    # Formatos de Cores por Atividade (ajuste cores/palavras-chave conforme necessário)
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
        # Adicione mais categorias/cores se precisar
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

# --- Ponto de Entrada do Script ---
if __name__ == "__main__":
    # Imports necessários AQUI DENTRO se não estiverem no topo global
    import argparse
    import sys
    # Chama a função principal para iniciar o processo
    main()   