# -*- coding: utf-8 -*-
import csv
from datetime import datetime, date, timedelta
import sys
import locale

# --- Configuração de Localização ---
try:
    locale.setlocale(locale.LC_TIME, 'pt_BR.UTF-8')
except locale.Error:
    try:
        locale.setlocale(locale.LC_TIME, 'Portuguese_Brazil.1252')
    except locale.Error:
        print("Aviso: Não foi possível configurar a localização para Português (pt_BR). Nomes dos dias podem aparecer em inglês.")
        DIAS_SEMANA_PT_FALLBACK = { 0: "Segunda", 1: "Terça", 2: "Quarta", 3: "Quinta", 4: "Sexta", 5: "Sábado", 6: "Domingo"}

# --- Constantes da Rotina ---
DATA_RETORNO_TRABALHO = date(2025, 4, 15)
DATA_BASE_TERAPIA_CASAL = date(2025, 4, 7) # Uma segunda com Terapia de Casal

# --- Definição Base de Eventos (Tratados com Lógica) ---
EVENTOS_FIXOS = {
    # Uni - Adicione aqui outros eventos fixos se houver em outros dias
    1: [{'Horario': '08:00 - 10:00', 'Atividade': 'AULA: Estágio Alfabetização', 'Observacoes': 'UFSCar - Compromisso Fixo'}], # Terça
    2: [{'Horario': '07:30 - 10:30', 'Atividade': 'Estágio na Escola (Inserção)', 'Observacoes': 'UFSCar - Compromisso Fixo'}], # Quarta
    3: [{'Horario': '08:00 - 10:30', 'Atividade': 'AULA: Escola e Currículo', 'Observacoes': 'UFSCar - Compromisso Fixo (Sair 10:30)'}], # Quinta
}

# Adiciona Vôlei separadamente para evitar erro de definição
VOLEI_EVENTOS = [
    {'Horario': '19:10 - 19:30', 'Atividade': 'Lanche Pré-Vôlei c/ Namorada', 'Observacoes': 'Encontrar antes'},
    {'Horario': '19:30 - 21:30', 'Atividade': 'Vôlei', 'Observacoes': 'Compromisso Fixo'}
]
EVENTOS_FIXOS[1] = EVENTOS_FIXOS.get(1, []) + VOLEI_EVENTOS # Adiciona Vôlei na Terça (dia 1)
EVENTOS_FIXOS[3] = EVENTOS_FIXOS.get(3, []) + VOLEI_EVENTOS # Adiciona Vôlei na Quinta (dia 3)

# --- Placeholders para edição manual (simplificados) ---
# (O dicionário PLACEHOLDERS continua igual como estava antes)
PLACEHOLDERS = {
    0: [{'Horario': 'Manhã', 'Atividade': '[EDITAR] Caminhada? Estudo?', 'Observacoes': 'FOLGA'},
        {'Horario': 'Tarde', 'Atividade': '[EDITAR] Afazeres?', 'Observacoes': 'FOLGA'}], # Terapia será adicionada por lógica
    4: [{'Horario': 'Manhã', 'Atividade': '[EDITAR] Caminhada? Estudo?', 'Observacoes': ''},
        {'Horario': 'Noite', 'Atividade': '[EDITAR] Tempo Livre com Namorada?', 'Observacoes': ''}],
    5: [{'Horario': 'Manhã', 'Atividade': '[EDITAR] Caminhada? Estudo?', 'Observacoes': 'Verificar Escala Trab/Folga'},
        {'Horario': 'Tarde', 'Atividade': '[EDITAR] Estudo Pedagogia?', 'Observacoes': 'SE FOLGA (2º/4º Sáb)'},
        {'Horario': 'Noite', 'Atividade': '[EDITAR] Tempo Livre com Namorada?', 'Observacoes': ''}],
    6: [{'Horario': 'Manhã', 'Atividade': '[EDITAR] Caminhada? Estudo?', 'Observacoes': 'Verificar Escala Trab/Folga'},
        {'Horario': 'Noite', 'Atividade': '[EDITAR] Tempo Livre com Namorada?', 'Observacoes': ''},
        {'Horario': 'Noite', 'Atividade': 'Planejar Semana Seguinte', 'Observacoes': 'Importante!'}],
}

# --- O resto do script (locale setup, get_dia_nome, gerar_agenda_semana, etc.) continua igual ---    


def get_dia_nome(dia_num):
    """Retorna o nome do dia em português, tentando via locale ou fallback."""
    try:
        dt_temp = date(2024, 1, 1) + timedelta(days=dia_num)
        nome = dt_temp.strftime('%A').capitalize()
        if locale.getlocale(locale.LC_TIME)[0] and 'Portuguese' in locale.getlocale(locale.LC_TIME)[0] and len(nome) > 0 :
             nome = nome[0].upper() + nome[1:]
        return nome if nome else DIAS_SEMANA_PT_FALLBACK.get(dia_num)
    except:
        return DIAS_SEMANA_PT_FALLBACK.get(dia_num, "")

def gerar_agenda_semana(data_inicio_str):
    """Gera as linhas do CSV para a semana iniciando na data fornecida."""
    try:
        data_inicio = datetime.strptime(data_inicio_str, '%Y-%m-%d').date()
    except ValueError:
        print("Erro: Formato de data inválido. Use AAAA-MM-DD.")
        return None

    linhas_agenda_dia = {} # Dicionário para agrupar por dia antes de ordenar

    print(f"Gerando agenda base para a semana de {data_inicio_str}...")

    for i in range(7):
        data_atual = data_inicio + timedelta(days=i)
        dia_semana_num = data_atual.weekday()
        dia_nome_pt = get_dia_nome(dia_semana_num)
        dia_formatado = f"{dia_nome_pt} ({data_atual.strftime('%d/%m')})"

        eventos_do_dia_atual = []

        # 1. Adiciona Placeholders do dia
        eventos_do_dia_atual.extend(PLACEHOLDERS.get(dia_semana_num, []))

        # 2. Adiciona Eventos Fixos (Aulas, Vôlei)
        eventos_do_dia_atual.extend(EVENTOS_FIXOS.get(dia_semana_num, []))

        # 3. Adiciona Terapias (Segunda-feira)
        if dia_semana_num == 0: # Segunda
            semana_relativa = (data_atual - DATA_BASE_TERAPIA_CASAL).days // 7
            if semana_relativa % 2 == 0: # Semanas da Terapia de Casal (0, 2, 4...)
                eventos_do_dia_atual.append({'Horario': '15:30 - 16:30', 'Atividade': 'Terapia de Casal', 'Observacoes': 'Compromisso Fixo (Quinzenal)'})
            else: # Semanas da Terapia Individual (1, 3, 5...)
                # Assumindo duração de 1h para Terapia Individual
                eventos_do_dia_atual.append({'Horario': '11:15 - 12:15', 'Atividade': 'Terapia Individual', 'Observacoes': 'Compromisso Fixo (Quinzenal)'})

        # 4. Adiciona Trabalho (Condicional)
        if data_atual >= DATA_RETORNO_TRABALHO:
            if 1 <= dia_semana_num <= 4: # Terça a Sexta
                eventos_do_dia_atual.append({'Horario': '11:40 - 18:00', 'Atividade': 'Trabalho', 'Observacoes': 'Expediente'})
            elif dia_semana_num in [5, 6]: # Sábado ou Domingo
                # Calcula se é 1º ou 3º fim de semana do mês
                dia_do_mes = data_atual.day
                semana_no_mes = (dia_do_mes - 1) // 7 + 1
                if semana_no_mes in [1, 3]:
                    eventos_do_dia_atual.append({'Horario': '11:40 - 18:00', 'Atividade': 'Trabalho', 'Observacoes': f'{semana_no_mes}º Fim de Semana - Expediente'})


        # Formata as linhas para o CSV final
        for evento in eventos_do_dia_atual:
             if dia_formatado not in linhas_agenda_dia:
                 linhas_agenda_dia[dia_formatado] = []
             linhas_agenda_dia[dia_formatado].append([
                dia_formatado,
                evento.get('Horario', ''),
                evento.get('Atividade', ''),
                evento.get('Observacoes', '')
             ])


    # Ordena os eventos dentro de cada dia e depois junta tudo
    linhas_finais_ordenadas = []
    for i in range(7): # Garante a ordem dos dias da semana
        data_referencia = data_inicio + timedelta(days=i)
        dia_chave = f"{get_dia_nome(data_referencia.weekday())} ({data_referencia.strftime('%d/%m')})"

        if dia_chave in linhas_agenda_dia:
            # Ordena eventos do dia pela hora de início (melhor esforço)
            def sort_key_time(row):
                try:
                    time_part = datetime.strptime(row[1].split('-')[0].strip(), '%H:%M').time()
                except:
                    horario_lower = str(row[1]).lower().strip()
                    if 'manhã' in horario_lower or not horario_lower : time_part = datetime.min.time().replace(hour=8)
                    elif 'tarde' in horario_lower: time_part = datetime.min.time().replace(hour=14)
                    elif 'noite' in horario_lower or 'onwards' in horario_lower: time_part = datetime.min.time().replace(hour=20)
                    else: time_part = datetime.min.time().replace(hour=12)
                return time_part

            try:
                linhas_agenda_dia[dia_chave].sort(key=sort_key_time)
            except Exception as e:
                 print(f"Aviso: Erro ao ordenar eventos para {dia_chave}: {e}")

            linhas_finais_ordenadas.extend(linhas_agenda_dia[dia_chave])

    return linhas_finais_ordenadas

def escrever_csv(linhas, nome_arquivo):
    """Escreve as linhas em um arquivo CSV."""
    # (Função escrever_csv continua igual à versão anterior)
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
        nome_arquivo_saida = f"agenda_base_{data_inicio_input}.csv"
        escrever_csv(linhas_geradas, nome_arquivo_saida)
