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
    """Aplica formatação (incluindo cor de linha inteira CORRIGIDO) e escreve os dados."""
    print("Aplicando formatação (com cor de linha) e escrevendo dados...")

    # --- Escrever Cabeçalho ---
    for col_num, value in enumerate(df.columns.values):
        worksheet.write(0, col_num, value, formatos['header'])

    # --- Aplicar Formato e Cores às Células de Dados (LINHA INTEIRA - CORRIGIDO) ---
    for row_num in range(len(df)):
        # Determinar a cor de fundo para a linha atual
        activity_text = str(df.iloc[row_num, 2]).lower() # Coluna 'Atividade'
        bg_color_to_apply = None # Nenhuma cor de fundo inicialmente
        # Tenta encontrar uma cor correspondente à atividade
        for keyword, fmt in formatos['colors'].items():
            # Acessamos diretamente a propriedade bg_color definida no formato de cor
            # Se essa propriedade existir no objeto fmt (deve existir como definimos)
            if hasattr(fmt, 'bg_color') and keyword.lower() in activity_text:
                bg_color_to_apply = fmt.bg_color
                break # Usa a primeira cor encontrada

        # Aplicar formato a cada célula da linha
        for col_num in range(len(df.columns)):
            cell_value = df.iloc[row_num, col_num]

            # 1. Definir propriedades base para a célula (como um dicionário)
            if col_num < 2: # Colunas Dia e Horario
                final_props = {'border': 1, 'text_wrap': True, 'valign': 'top', 'align': 'center'}
            else: # Colunas Atividade e Observacoes
                final_props = {'border': 1, 'text_wrap': True, 'valign': 'top'}

            # 2. Adicionar cor de fundo se aplicável
            if bg_color_to_apply:
                final_props['bg_color'] = bg_color_to_apply

            # 3. Criar o formato final para ESTA célula específica
            final_cell_format = workbook.add_format(final_props)

            # 4. Escrever a célula com o formato final
            worksheet.write(row_num + 1, col_num, cell_value, final_cell_format)

    # --- Ajustar Largura das Colunas --- (Igual antes)
    worksheet.set_column('A:A', 18)
    worksheet.set_column('B:B', 15)
    worksheet.set_column('C:C', 40)
    worksheet.set_column('D:D', 50)
    print("Largura das colunas ajustada.")

    # --- Mesclar Células da Coluna 'Dia' --- (Lógica revisada para usar 'final_props')
    print("Mesclando células da coluna 'Dia'...")
    merge_start_row = 1
    for i in range(1, len(df)):
        current_day = df.iloc[i, 0]
        previous_day = df.iloc[i-1, 0]

        # Define as propriedades base para a célula do Dia (será usada na mesclagem ou escrita)
        day_props = {'border': 1, 'text_wrap': True, 'valign': 'top', 'align': 'center'}

        if current_day == previous_day:
             # Se é o fim de uma sequência ou o fim do dataframe
            if i == len(df) - 1 or df.iloc[i+1, 0] != current_day:
                if merge_start_row <= i: # Mescla apenas se houver > 1 linha
                    # Determina a cor da primeira linha da mesclagem para aplicar a toda a área mesclada
                    activity_text_merge = str(df.iloc[merge_start_row - 1, 2]).lower()
                    bg_color_merge = None
                    for keyword, fmt in formatos['colors'].items():
                        if hasattr(fmt, 'bg_color') and keyword.lower() in activity_text_merge:
                            bg_color_merge = fmt.bg_color
                            break
                    # Adiciona a cor às propriedades base do Dia
                    if bg_color_merge:
                        day_props['bg_color'] = bg_color_merge
                    # Cria o formato final para a mesclagem
                    final_merge_format = workbook.add_format(day_props)
                    worksheet.merge_range(merge_start_row, 0, i + 1, 0, current_day, final_merge_format)
                merge_start_row = i + 2 # Próxima linha (índice i+1) começa nova seq, então start = i+2
        else:
             # A linha anterior (i-1) terminou uma sequência (ou era única)
             # Se ela não iniciou a sequência que acabou de ser mesclada, escreve-a individualmente
             if merge_start_row == i: # Se a linha anterior era o início da sequencia atual
                 activity_text_prev = str(df.iloc[i-1, 2]).lower()
                 bg_color_prev = None
                 for keyword, fmt in formatos['colors'].items():
                      if hasattr(fmt, 'bg_color') and keyword.lower() in activity_text_prev:
                          bg_color_prev = fmt.bg_color
                          break
                 # Adiciona cor às propriedades base e cria formato final
                 if bg_color_prev:
                     day_props['bg_color'] = bg_color_prev
                 final_prev_format = workbook.add_format(day_props)
                 worksheet.write(i, 0, previous_day, final_prev_format) # Escreve a linha i (dados da linha i-1 do df)
             merge_start_row = i + 1 # Linha atual i (índice i) inicia nova sequência

    # Trata caso da última linha ser única
    if merge_start_row == len(df) : # Se a última linha iniciou uma nova sequência
          activity_text_last = str(df.iloc[len(df)-1, 2]).lower()
          bg_color_last = None
          for keyword, fmt in formatos['colors'].items():
              if hasattr(fmt, 'bg_color') and keyword.lower() in activity_text_last:
                  bg_color_last = fmt.bg_color
                  break
          last_props = {'border': 1, 'text_wrap': True, 'valign': 'top', 'align': 'center'}
          if bg_color_last:
              last_props['bg_color'] = bg_color_last
          final_last_format = workbook.add_format(last_props)
          worksheet.write(len(df), 0, df.iloc[len(df)-1, 0], final_last_format) # Escreve linha len(df) (dados da linha len(df)-1)
    print("Mesclagem (com cores) concluída.")


    # --- Congelar Painel Superior --- (Igual antes)
    worksheet.freeze_panes(1, 0)
    print("Painel superior congelado.")
def salvar_planilha(writer):
    """Fecha o writer e salva o arquivo Excel."""
    print("Salvando arquivo Excel...")
    try:
        writer.close()
        print("Arquivo salvo com sucesso!")
    except Exception as e:
        print(f"Erro ao salvar o arquivo Excel: {e}")

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
            df_agenda.to_excel(excel_writer, index=False, sheet_name=SHEET_NAME, startrow=1, header=False)

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