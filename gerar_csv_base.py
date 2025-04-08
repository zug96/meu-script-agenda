# -*- coding: utf-8 -*-
import csv
from datetime import datetime, date, timedelta
import sys
import locale # Import locale para nomes de dias da semana

# --- Configuração de Localização (Para Nomes dos Dias em Português) ---
try:
    # Tenta definir para Português do Brasil no Linux/macOS
    locale.setlocale(locale.LC_TIME, 'pt_BR.UTF-8') 
except locale.Error:
    try:
        # Tenta definir para Português do Brasil no Windows
        locale.setlocale(locale.LC_TIME, 'Portuguese_Brazil.1252')
    except locale.Error:
        print("Aviso: Não foi possível configurar a localização para Português (pt_BR). Nomes dos dias podem aparecer em inglês.")
        # Fallback manual se a localização falhar
        DIAS_SEMANA_PT_FALLBACK = { 0: "Segunda", 1: "Terça", 2: "Quarta", 3: "Quinta", 4: "Sexta", 5: "Sábado", 6: "Domingo"}


# --- Definição dos Eventos Recorrentes ---
# (Ajuste os horários/nomes/observações conforme sua rotina FINAL)
EVENTOS_RECORRENTES = {
    # Weekday: 0=Segunda, 1=Terça, ..., 6=Domingo
    0: [ # Segunda-feira
          {'Horario': '06:45 - 07:15', 'Atividade': '[EDITAR] Caminhada?', 'Observacoes': 'FOLGA'},
          {'Horario': '08:00 - 11:30', 'Atividade': '[EDITAR] Bloco Estudo Principal (Linux/Pedagogia)', 'Observacoes': 'Definir Foco'},
          {'Horario': '12:30 - 14:00', 'Atividade': '[EDITAR] Bloco Estudo Pedagogia', 'Observacoes': 'Leituras/Tarefas'},
          {'Horario': '15:30 - 16:30', 'Atividade': 'Terapia de Casal', 'Observacoes': 'Compromisso Fixo'},
          {'Horario': '21:30 - 22:15', 'Atividade': '[EDITAR] Estudo Leve Pedagogia?', 'Observacoes': 'Revisão/Leitura Rápida'},
        ],
    1: [ # Terça-feira
          {'Horario': '08:00 - 10:00', 'Atividade': 'AULA: Estágio Alfabetização', 'Observacoes': 'UFSCar - Compromisso Fixo'},
          {'Horario': '11:40 - 18:00', 'Atividade': 'Trabalho', 'Observacoes': 'Expediente'},
          {'Horario': '19:10 - 19:30', 'Atividade': 'Lanche Pré-Vôlei c/ Namorada', 'Observacoes': 'Encontrar antes'},
          {'Horario': '19:30 - 21:30', 'Atividade': 'Vôlei', 'Observacoes': 'Compromisso Fixo'},
        ],
    2: [ # Quarta-feira
          {'Horario': '07:30 - 10:30', 'Atividade': 'Estágio na Escola (Inserção)', 'Observacoes': 'UFSCar - Compromisso Fixo'},
          {'Horario': '11:40 - 18:00', 'Atividade': 'Trabalho', 'Observacoes': 'Expediente'},
          {'Horario': '19:15 - 20:45', 'Atividade': '[EDITAR] Bloco Estudo (Linux/Pedagogia)', 'Observacoes': 'Definir Foco'},
          {'Horario': '20:45 - 21:45', 'Atividade': '[EDITAR] Bloco Estudo (Pedagogia)', 'Observacoes': 'Definir Foco'},
        ],
    3: [ # Quinta-feira
          {'Horario': '08:00 - 10:30', 'Atividade': 'AULA: Escola e Currículo', 'Observacoes': 'UFSCar - Compromisso Fixo (Sair 10:30)'},
          {'Horario': '11:40 - 18:00', 'Atividade': 'Trabalho', 'Observacoes': 'Expediente'},
          {'Horario': '19:10 - 19:30', 'Atividade': 'Lanche Pré-Vôlei c/ Namorada', 'Observacoes': 'Encontrar antes'},
          {'Horario': '19:30 - 21:30', 'Atividade': 'Vôlei', 'Observacoes': 'Compromisso Fixo'},
        ],
    4: [ # Sexta-feira
          {'Horario': '07:30 - 08:15', 'Atividade': '[EDITAR] Caminhada?', 'Observacoes': ''},
          {'Horario': '08:15 - 10:00', 'Atividade': '[EDITAR] Bloco Estudo (Linux/Pedagogia)', 'Observacoes': 'Definir Foco'},
          {'Horario': '10:00 - 10:30', 'Atividade': '[EDITAR] Bloco Estudo (Pedagogia)', 'Observacoes': 'Definir Foco'},
          {'Horario': '11:40 - 18:00', 'Atividade': 'Trabalho', 'Observacoes': 'Expediente'},
          {'Horario': '19:30 onwards', 'Atividade': '[EDITAR] Tempo Livre com Namorada', 'Observacoes': 'Definir Atividade'},
        ],
    5: [ # Sábado
          {'Horario': 'Manhã', 'Atividade': '[EDITAR] Caminhada? Estudo?', 'Observacoes': 'Verificar Escala (Folga=Sim / Trab=Não)'},
          {'Horario': '11:40 - 18:00', 'Atividade': 'Trabalho', 'Observacoes': 'SE FOR 1º/3º SÁBADO (Confirmar!)'},
          {'Horario': 'Tarde', 'Atividade': '[EDITAR] Estudo Pedagogia?', 'Observacoes': 'SE FOR FOLGA (2º/4º Sáb)'},
          {'Horario': 'Noite', 'Atividade': '[EDITAR] Tempo Livre com Namorada', 'Observacoes': 'Definir Atividade'},
        ],
    6: [ # Domingo
          {'Horario': 'Manhã', 'Atividade': '[EDITAR] Caminhada? Estudo?', 'Observacoes': 'Verificar Escala (Folga=Sim / Trab=Não)'},
          {'Horario': '11:40 - 18:00', 'Atividade': 'Trabalho', 'Observacoes': 'SE FOR 1º/3º DOMINGO (Confirmar!)'},
          {'Horario': 'Noite', 'Atividade': '[EDITAR] Tempo Livre com Namorada', 'Observacoes': 'Definir Atividade'},
          {'Horario': 'Noite', 'Atividade': 'Planejar Semana', 'Observacoes': 'Importante!'},
        ],
}

def get_dia_nome(dia_num):
    """Retorna o nome do dia em português, tentando via locale ou fallback."""
    try:
        # Tenta gerar nome pelo locale configurado
        dt_temp = date(2024, 1, 1) + timedelta(days=dia_num - 0) # Gera uma data para obter o nome do dia
        nome = dt_temp.strftime('%A').capitalize()
        # Workaround para possível problema de capitalização no Windows locale
        if locale.getlocale(locale.LC_TIME)[0] and 'Portuguese' in locale.getlocale(locale.LC_TIME)[0] and len(nome) > 0 :
             nome = nome[0].upper() + nome[1:]
        return nome if nome else DIAS_SEMANA_PT_FALLBACK.get(dia_num) # Garante retorno se strftime falhar
    except:
         # Usa fallback manual se locale falhar
        return DIAS_SEMANA_PT_FALLBACK.get(dia_num, "")


def gerar_agenda_semana(data_inicio_str):
    """Gera as linhas do CSV para a semana iniciando na data fornecida."""
    try:
        data_inicio = datetime.strptime(data_inicio_str, '%Y-%m-%d').date()
    except ValueError:
        print("Erro: Formato de data inválido. Use AAAA-MM-DD.")
        return None # Retorna None em caso de erro

    linhas_agenda = []
    print(f"Gerando agenda base para a semana de {data_inicio_str}...")

    for i in range(7):
        data_atual = data_inicio + timedelta(days=i)
        dia_semana_num = data_atual.weekday() # Segunda = 0, Domingo = 6
        dia_nome_pt = get_dia_nome(dia_semana_num)
        dia_formatado = f"{dia_nome_pt} ({data_atual.strftime('%d/%m')})"

        # Adiciona eventos recorrentes do dia
        eventos_do_dia = EVENTOS_RECORRENTES.get(dia_semana_num, [])
        for evento in eventos_do_dia:
             # Adiciona linha formatada
            linhas_agenda.append([
                dia_formatado,
                evento.get('Horario', ''),
                evento.get('Atividade', ''),
                evento.get('Observacoes', '')
            ])

    # Ordena por data e depois tenta ordenar por hora (melhor esforço)
    def sort_key(row):
        try:
            # Extrai data da coluna 'Dia'
            date_part_str = row[0].split('(')[1].split(')')[0]
            # Cria um objeto date, inferindo o ano (pode dar problema na virada do ano)
            date_part = datetime.strptime(date_part_str, '%d/%m').date().replace(year=data_inicio.year)
            # Heurística simples para virada de ano (se a data gerada for muito antes da data de início)
            if (data_inicio - date_part).days > 180: # Mais de ~6 meses de diferença
                 date_part = date_part.replace(year=data_inicio.year + 1)
            elif (date_part - data_inicio).days > 180: # Data gerada muito depois
                 date_part = date_part.replace(year=data_inicio.year -1) # Improvável, mas seguro ter

        except:
            date_part = data_inicio # Usa data de início como fallback

        try:
            # Extrai hora de início (se possível) da coluna 'Horario'
            time_part = datetime.strptime(row[1].split('-')[0].strip(), '%H:%M').time()
        except:
            # Define ordem para Manhã/Tarde/Noite ou horários inválidos/vazios
            horario_lower = str(row[1]).lower().strip()
            if 'manhã' in horario_lower or not horario_lower : # Manhã ou Vazio = Manhã
                time_part = datetime.min.time().replace(hour=8) # Manhã ~8h
            elif 'tarde' in horario_lower:
                time_part = datetime.min.time().replace(hour=14) # Tarde ~14h
            elif 'noite' in horario_lower or 'onwards' in horario_lower:
                time_part = datetime.min.time().replace(hour=20) # Noite ~20h
            else:
                 time_part = datetime.min.time().replace(hour=12) # Meio-dia como fallback

        return (date_part, time_part) # Ordena primeiro por data, depois por hora estimada

    try:
        linhas_agenda.sort(key=sort_key)
    except Exception as e:
        print(f"Aviso: Ocorreu um erro ao tentar ordenar os eventos: {e}")
        print("A ordem dos eventos no CSV pode não estar ideal.")

    return linhas_agenda

def escrever_csv(linhas, nome_arquivo):
    """Escreve as linhas em um arquivo CSV."""
    try:
        with open(nome_arquivo, 'w', newline='', encoding='utf-8') as csvfile:
            writer = csv.writer(csvfile)
            writer.writerow(['Dia', 'Horario', 'Atividade', 'Observacoes']) # Cabeçalho
            writer.writerows(linhas)
        print(f"Arquivo base '{nome_arquivo}' gerado com sucesso!")
        print("--> ABRA este arquivo e EDITE os placeholders [EDITAR] e confirme a escala de Sáb/Dom antes de usá-lo!")
    except Exception as e:
        print(f"Erro ao escrever o arquivo CSV: {e}")
        sys.exit(1)

# --- Execução Principal ---
if __name__ == "__main__":
    if len(sys.argv) != 2:
        print("Erro: Forneça a data de início da semana (geralmente uma Segunda) no formato AAAA-MM-DD.")
        print("Uso: python3 gerar_csv_base.py AAAA-MM-DD")
        sys.exit(1)

    data_inicio_input = sys.argv[1]
    linhas_geradas = gerar_agenda_semana(data_inicio_input)

    if linhas_geradas is not None:
        # Cria nome do arquivo de saída
        nome_arquivo_saida = f"agenda_base_{data_inicio_input}.csv"
        escrever_csv(linhas_geradas, nome_arquivo_saida)
