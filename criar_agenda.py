import pandas as pd

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
excel_filename = 'agenda_semana_31mar_05abr_formatada.xlsx' # 

# Criar um Excel writer usando XlsxWriter como engine.
writer = pd.ExcelWriter(excel_filename, engine='xlsxwriter')

# Converter o dataframe para um objeto XlsxWriter Excel.
# index=False para não escrever o índice do dataframe no Excel.
# sheet_name para nomear a aba da planilha.
df.to_excel(writer, index=False, sheet_name='AgendaSemanal')

# Obter os objetos workbook e worksheet do XlsxWriter.
workbook  = writer.book
worksheet = writer.sheets['AgendaSemanal']

# --- Início da Formatação ---

# Definir formato para o cabeçalho (Negrito)
header_format = workbook.add_format({
    'bold': True,
    'text_wrap': True,
    'valign': 'top', # Alinhamento vertical no topo
    'fg_color': '#D7E4BC', # Uma cor de fundo suave para o cabeçalho
    'border': 1})

# Definir formato para as células de dados (Quebra de texto)
cell_format = workbook.add_format({'text_wrap': True, 'valign': 'top'})

# Aplicar o formato de cabeçalho
for col_num, value in enumerate(df.columns.values):
    worksheet.write(0, col_num, value, header_format)

# Definir a largura das colunas (ajuste conforme necessário)
# Formato: worksheet.set_column('ColunaInicial:ColunaFinal', Largura)
worksheet.set_column('A:A', 18, cell_format) # Coluna Dia
worksheet.set_column('B:B', 15, cell_format) # Coluna Horário
worksheet.set_column('C:C', 40, cell_format) # Coluna Atividade (mais larga)
worksheet.set_column('D:D', 50, cell_format) # Coluna Observações (mais larga)

# (Opcional) Poderia adicionar mais formatações aqui:
# - Cores diferentes para tipos de atividade (condicional formatting)
# - Congelar painéis (worksheet.freeze_panes(1, 0)) # Congela a primeira linha

# --- Fim da Formatação ---

# Fechar o Pandas Excel writer e gerar o arquivo Excel.
writer.close() # Antes era save(), mas close() é recomendado para XlsxWriter

print(f"Arquivo Excel formatado '{excel_filename}' criado com sucesso!")