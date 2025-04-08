#!/bin/bash

# Script para automatizar a atualização da agenda e geração do ICS

# Verifica se os argumentos foram passados
if [ -z "$1" ] || [ -z "$2" ]; then
  echo "Erro: Forneça o nome do arquivo CSV de entrada e o nome do arquivo ICS de saída."
  echo "Uso: ./atualizar_agenda.sh <arquivo_csv_entrada.csv> <arquivo_ics_saida.ics>"
  exit 1
fi

CSV_FILE=$1
ICS_FILE=$2
COMMIT_MESSAGE="Data: Atualiza agenda via script ($CSV_FILE)"

echo "--- Iniciando atualização da agenda ---"

# 1. (Opcional) Commitar alterações no CSV
echo "Adicionando $CSV_FILE ao Git..."
git add "$CSV_FILE"
echo "Fazendo commit..."
# Verifica se há algo para commitar antes de tentar
if git diff --staged --quiet; then
  echo "Nenhuma alteração no CSV para commitar."
else
  git commit -m "$COMMIT_MESSAGE"
  echo "Enviando para o GitHub (opcional)..."
  git push origin feature/ics-argument # Ou a branch que estiver usando
fi

# 2. Ativar ambiente virtual
echo "Ativando ambiente virtual..."
source .venv/bin/activate
if [ $? -ne 0 ]; then
    echo "Erro ao ativar o ambiente virtual .venv!"
    exit 1
fi
echo "Ambiente virtual ativado."

# 3. Executar script Python para gerar ICS
echo "Executando script Python para gerar $ICS_FILE a partir de $CSV_FILE..."
python3 criar_agenda.py -c "$CSV_FILE" --ics "$ICS_FILE"
if [ $? -ne 0 ]; then
    echo "Erro ao executar o script criar_agenda.py!"
    deactivate # Desativa mesmo se deu erro no python
    exit 1
fi
echo "Script Python executado com sucesso."

# 4. Desativar ambiente virtual
echo "Desativando ambiente virtual..."
deactivate
echo "Ambiente virtual desativado."

echo "--- Atualização da agenda concluída ---"
echo "Arquivo gerado: $ICS_FILE"

exit 0
