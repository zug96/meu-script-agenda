import pandas as pd
import numpy as np # Necessário para verificar NaNs ao mesclar

# Carregar dados do arquivo CSV
try:
    df = pd.read_csv('agenda_data.csv', encoding='utf-8')
    # Tratar valores NaN (células vazias) como strings vazias, especialmente para Observacoes
    df['Observacoes'] = df['Observacoes'].fillna('')
except FileNotFoundError:
    print("Erro: Arquivo 'agenda_data.csv' não encontrado.")
    print("Certifique-se de que o arquivo está na mesma pasta que o script.")
    exit() # Sai do script se o arquivo não for encontrado
except Exception as e:
    print(f"Erro ao ler o arquivo CSV: {e}")
    exit()

# Dados da agenda 
data = [
    # Segunda-feira, 31/03
    {'Dia': 'Segunda (31/03)', 'Horário': '06:45 - 07:15', 'Atividade': 'Caminhada com pai/cachorras', 'Observações': 'Ajustada (30 min). Conversar com pai.'},
    {'Dia': 'Segunda (31/03)', 'Horário': '07:15 - 08:30', 'Atividade': 'Café da Manhã / Organização', 'Observações': 'Preparar material estudo Linux.'},
    {'Dia': 'Segunda (31/03)', 'Horário': '08:30 - 12:00', 'Atividade': 'Estudo Linux (DIO)', 'Observações': 'Bloco Focado. Pausas curtas.'},
    {'Dia': 'Segunda (31/03)', 'Horário': '12:00 - 13:30', 'Atividade': 'Almoço / Descanso', 'Observações': 'Desconectar.'},
    {'Dia': 'Segunda (31/03)', 'Horário': '13:30 - 15:30', 'Atividade': 'Estudo Linux (DIO) ou Afazeres', 'Observações': 'Flexível: Mais estudo ou Buscar guias médicas (verificar horários).'},
    {'Dia': 'Segunda (31/03)', 'Horário': '15:30 - 18:00', 'Atividade': 'Tempo Livre / Descanso', 'Observações': 'Evitar obrigações.'},
    {'Dia': 'Segunda (31/03)', 'Horário': '18:00 em diante', 'Atividade': 'Tempo Livre / Tarefas Leves / Jantar', 'Observações': 'Relaxar, planejar brevemente a Terça.'},

    # Terça-feira, 01/04
    {'Dia': 'Terça (01/04)', 'Horário': '06:45 - 07:15', 'Atividade': 'Caminhada com pai/cachorras', 'Observações': 'Ajustada (30 min).'},
    {'Dia': 'Terça (01/04)', 'Horário': '07:15 - 08:00', 'Atividade': 'Café da Manhã / Preparação Aula', 'Observações': ''},
    {'Dia': 'Terça (01/04)', 'Horário': '08:00 - 12:00', 'Atividade': 'AULA PEDAGOGIA (Escola e Currículo)', 'Observações': 'Compromisso Fixo.'},
    {'Dia': 'Terça (01/04)', 'Horário': '12:00 - 13:30', 'Atividade': 'Almoço / Descanso', 'Observações': ''},
    {'Dia': 'Terça (01/04)', 'Horário': '13:30 - 16:00', 'Atividade': 'Estudo Linux (DIO)', 'Observações': 'Bloco Focado.'},
    {'Dia': 'Terça (01/04)', 'Horário': '16:00 - 19:00', 'Atividade': 'Tempo Livre / Descanso / Preparo Vôlei', 'Observações': ''},
    {'Dia': 'Terça (01/04)', 'Horário': '19:00 - 19:30', 'Atividade': 'Deslocamento / Preparação Vôlei', 'Observações': ''},
    {'Dia': 'Terça (01/04)', 'Horário': '19:30 - 21:30', 'Atividade': 'VÔLEI (SESC)', 'Observações': 'Compromisso Fixo. Aproveitar!'},
    {'Dia': 'Terça (01/04)', 'Horário': '21:30 em diante', 'Atividade': 'Retorno / Banho / Jantar Leve / Relaxar', 'Observações': ''},

    # Quarta-feira, 02/04
    {'Dia': 'Quarta (02/04)', 'Horário': '06:45 - 07:15', 'Atividade': 'Caminhada com pai/cachorras', 'Observações': 'Ajustada (30 min).'},
    {'Dia': 'Quarta (02/04)', 'Horário': '07:15 - 08:30', 'Atividade': 'Café da Manhã / Organização', 'Observações': ''},
    {'Dia': 'Quarta (02/04)', 'Horário': '08:30 - 12:00', 'Atividade': 'Estudo Linux (DIO)', 'Observações': 'Bloco Focado.'},
    {'Dia': 'Quarta (02/04)', 'Horário': '12:00 - 13:30', 'Atividade': 'Almoço / Descanso', 'Observações': ''},
    {'Dia': 'Quarta (02/04)', 'Horário': '13:30 - 15:30', 'Atividade': 'Afazeres / Flexível', 'Observações': 'Sugestão: Buscar guias restantes E/OU Aplicar Injeção. Ou estudo/descanso.'},
    {'Dia': 'Quarta (02/04)', 'Horário': '15:30 - 18:00', 'Atividade': 'Tempo Livre / Descanso', 'Observações': 'Descanso ativo ou passivo.'},
    {'Dia': 'Quarta (02/04)', 'Horário': '18:00 em diante', 'Atividade': 'Tempo Livre / Tarefas Leves / Jantar', 'Observações': 'Relaxar.'},

    # Quinta-feira, 03/04
    {'Dia': 'Quinta (03/04)', 'Horário': '06:45 - 07:15', 'Atividade': 'Caminhada com pai/cachorras', 'Observações': 'Ajustada (30 min).'},
    {'Dia': 'Quinta (03/04)', 'Horário': '07:15 - 08:00', 'Atividade': 'Café da Manhã / Preparação Aula', 'Observações': ''},
    {'Dia': 'Quinta (03/04)', 'Horário': '08:00 - 10:00', 'Atividade': 'AULA PEDAGOGIA (Estágio Alfabetização)', 'Observações': 'Compromisso Fixo.'},
    {'Dia': 'Quinta (03/04)', 'Horário': '10:00 - 12:00', 'Atividade': 'Pausa Pós-Aula / Estudo Leve Linux', 'Observações': 'Organizar notas ou iniciar estudo.'},
    {'Dia': 'Quinta (03/04)', 'Horário': '12:00 - 13:30', 'Atividade': 'Almoço / Descanso', 'Observações': ''},
    {'Dia': 'Quinta (03/04)', 'Horário': '13:30 - 16:00', 'Atividade': 'Estudo Linux (DIO)', 'Observações': 'Bloco Focado.'},
    {'Dia': 'Quinta (03/04)', 'Horário': '16:00 - 19:00', 'Atividade': 'Tempo Livre / Descanso / Preparo Vôlei', 'Observações': ''},
    {'Dia': 'Quinta (03/04)', 'Horário': '19:00 - 19:30', 'Atividade': 'Deslocamento / Preparação Vôlei', 'Observações': ''},
    {'Dia': 'Quinta (03/04)', 'Horário': '19:30 - 21:30', 'Atividade': 'VÔLEI (SESC)', 'Observações': 'Compromisso Fixo. Divirta-se!'},
    {'Dia': 'Quinta (03/04)', 'Horário': '21:30 em diante', 'Atividade': 'Retorno / Banho / Jantar Leve / Relaxar', 'Observações': ''},

    # Sexta-feira, 04/04
    {'Dia': 'Sexta (04/04)', 'Horário': '06:45 - 07:15', 'Atividade': 'Caminhada com pai/cachorras', 'Observações': 'Ajustada (30 min).'},
    {'Dia': 'Sexta (04/04)', 'Horário': '07:15 - 08:30', 'Atividade': 'Café da Manhã / Planejamento', 'Observações': 'Planejar dia e fds.'},
    {'Dia': 'Sexta (04/04)', 'Horário': '08:30 - 12:00', 'Atividade': 'Estudo Linux (DIO)', 'Observações': 'Bloco Focado. Concluir etapa/módulo.'},
    {'Dia': 'Sexta (04/04)', 'Horário': '12:00 - 13:30', 'Atividade': 'Almoço / Descanso', 'Observações': ''},
    {'Dia': 'Sexta (04/04)', 'Horário': '13:30 - 16:00', 'Atividade': 'Revisão Estudo / Afazeres Finais / Descanso', 'Observações': 'Flexível. Resolver pendências ou iniciar descanso.'},
    {'Dia': 'Sexta (04/04)', 'Horário': '16:00 em diante', 'Atividade': 'INÍCIO FIM DE SEMANA', 'Observações': 'Desconectar, Lazer, Relaxar.'},

    # Sábado, 05/04
    {'Dia': 'Sábado (05/04)', 'Horário': 'Manhã', 'Atividade': 'Flexível', 'Observações': 'Opções: Caminhada, Estudo Leve (sem pressão), Tarefas Domésticas. Acordar com calma.'},
    {'Dia': 'Sábado (05/04)', 'Horário': 'Tarde', 'Atividade': 'Almoço / Lazer / Descanso', 'Observações': 'Prioridade: Recarregar. Passeio, hobby, descanso.'},
    {'Dia': 'Sábado (05/04)', 'Horário': 'Noite', 'Atividade': 'Jantar / Lazer / Relaxar', 'Observações': 'Preparar para Domingo tranquilo.'}
]

# Criar DataFrame
df = pd.DataFrame(data)

# Nome do arquivo Excel
excel_filename = 'agenda_semana_31mar_05abr_detalhada.xlsx' # Novo nome

# Criar um Excel writer usando XlsxWriter
writer = pd.ExcelWriter(excel_filename, engine='xlsxwriter')
df.to_excel(writer, index=False, sheet_name='AgendaSemanal', startrow=1, header=False) # Escreve os dados a partir da linha 2, sem cabeçalho do pandas

# Obter os objetos workbook e worksheet
workbook  = writer.book
worksheet = writer.sheets['AgendaSemanal']

# --- Definir Formatos ---

# Formato do Cabeçalho
header_format = workbook.add_format({
    'bold': True, 'text_wrap': True, 'valign': 'top',
    'fg_color': '#D7E4BC', 'border': 1, 'align': 'center'
})

# Formatos de Célula Base (com borda e quebra de texto)
cell_format_base = workbook.add_format({'border': 1, 'text_wrap': True, 'valign': 'top'})
cell_format_dia_hora = workbook.add_format({'border': 1, 'text_wrap': True, 'valign': 'top', 'align': 'center'}) # Centralizado para Dia/Hora

# Formatos de Cores por Atividade (adicione ou ajuste cores/palavras-chave)
color_formats = {
    'Estudo Linux': workbook.add_format({'bg_color': '#FFFFCC', 'border': 1, 'text_wrap': True, 'valign': 'top'}), # Amarelo claro
    'AULA PEDAGOGIA': workbook.add_format({'bg_color': '#CCFFFF', 'border': 1, 'text_wrap': True, 'valign': 'top'}), # Azul claro
    'Estágio': workbook.add_format({'bg_color': '#CCFFFF', 'border': 1, 'text_wrap': True, 'valign': 'top'}), # Azul claro
    'VÔLEI': workbook.add_format({'bg_color': '#CCFFCC', 'border': 1, 'text_wrap': True, 'valign': 'top'}), # Verde claro
    'Caminhada': workbook.add_format({'bg_color': '#CCFFCC', 'border': 1, 'text_wrap': True, 'valign': 'top'}), # Verde claro
    'Afazeres': workbook.add_format({'bg_color': '#FFDDCC', 'border': 1, 'text_wrap': True, 'valign': 'top'}), # Laranja claro
    'Flexível': workbook.add_format({'bg_color': '#E0E0E0', 'border': 1, 'text_wrap': True, 'valign': 'top'}), # Cinza claro (Sábado)
    'Tempo Livre': workbook.add_format({'bg_color': '#F0F0F0', 'border': 1, 'text_wrap': True, 'valign': 'top'}), # Cinza bem claro
    'Descanso': workbook.add_format({'bg_color': '#F0F0F0', 'border': 1, 'text_wrap': True, 'valign': 'top'}), # Cinza bem claro
}

# --- Aplicar Formatação e Cores ---

# Escrever o cabeçalho com o formato definido
for col_num, value in enumerate(df.columns.values):
    worksheet.write(0, col_num, value, header_format)

# Aplicar formato e cores às células de dados
for row_num in range(len(df)):
    # Identificar a cor baseada na atividade (case-insensitive)
    activity_text = str(df.iloc[row_num, 2]).lower() # Coluna 'Atividade' é a índice 2
    chosen_format = cell_format_base # Formato padrão
    for keyword, fmt in color_formats.items():
        if keyword.lower() in activity_text:
            chosen_format = fmt
            break # Usa o primeiro formato encontrado

    # Aplicar formato às colunas (exceto Dia e Hora que terão formato centralizado)
    worksheet.write(row_num + 1, 2, df.iloc[row_num, 2], chosen_format) # Atividade
    worksheet.write(row_num + 1, 3, df.iloc[row_num, 3], chosen_format) # Observações

    # Formato centralizado para Dia e Hora
    worksheet.write(row_num + 1, 0, df.iloc[row_num, 0], cell_format_dia_hora) # Dia
    worksheet.write(row_num + 1, 1, df.iloc[row_num, 1], cell_format_dia_hora) # Horário

# --- Ajustar Largura das Colunas ---
worksheet.set_column('A:A', 18) # Dia
worksheet.set_column('B:B', 15) # Horário
worksheet.set_column('C:C', 40) # Atividade
worksheet.set_column('D:D', 50) # Observações

# --- Mesclar Células da Coluna 'Dia' ---
merge_start_row = 1 # Linha inicial dos dados (abaixo do cabeçalho)
for i in range(1, len(df)):
    if df.iloc[i, 0] == df.iloc[i-1, 0]: # Se o dia atual é igual ao anterior
        # Se chegou ao fim da sequencia igual ou o proximo é diferente
        if i == len(df) - 1 or df.iloc[i+1, 0] != df.iloc[i, 0]:
            # Mescla da linha inicial da sequencia até a linha atual 'i'
            worksheet.merge_range(merge_start_row, 0, i + 1, 0, df.iloc[i, 0], cell_format_dia_hora)
            merge_start_row = i + 2 # Próxima linha será o início da nova sequencia
    else:
        # Se o dia anterior era único (ou fim de uma sequência curta)
        if merge_start_row == i + 1: # Verifica se precisa mesclar apenas a linha anterior
             worksheet.write(i, 0, df.iloc[i-1, 0], cell_format_dia_hora) # Reescreve com formato correto se não mesclou
        merge_start_row = i + 1 # Começa nova sequencia
# Trata caso da última linha ser única
if merge_start_row == len(df) + 1:
    worksheet.write(len(df), 0, df.iloc[len(df)-1, 0], cell_format_dia_hora)


# --- Congelar Painel Superior ---
worksheet.freeze_panes(1, 0) # Congela a linha 1 (cabeçalho)

# --- Fechar o Writer ---
writer.close()

print(f"Arquivo Excel super formatado '{excel_filename}' criado com sucesso!")