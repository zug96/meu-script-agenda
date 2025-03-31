# Projeto Agenda Semanal em Excel e ICS

## Descrição Curta

Este projeto consiste em um script Python que lê uma agenda semanal a partir de um arquivo CSV e gera automaticamente dois arquivos de saída:
1.  Uma planilha Excel (`.xlsx`) bem formatada, com cores, mesclagem de células e legenda.
2.  Um arquivo de calendário (`.ics`) que pode ser importado em aplicativos como Google Agenda, Outlook Calendar, etc.

## Motivação

Este projeto foi desenvolvido com dois objetivos principais:
* Criar uma ferramenta personalizada para organizar e visualizar a agenda semanal de forma prática.
* Aplicar e aprofundar conhecimentos em Python, manipulação de dados, Git, GitHub e o fluxo de trabalho de desenvolvimento, como parte do aprendizado contínuo e especificamente para o Desafio de Projeto "Contribuindo em um Projeto Open Source no GitHub" do Santander Bootcamp em parceria com a DIO.

## Nota sobre o Desenvolvimento: 
Este projeto foi desenvolvido em um processo interativo e colaborativo com a Inteligência Artificial Gemini (Modelo Experimental do Google). Como iniciante nos tópicos abordados, utilizei a IA como ferramenta de aprendizado para obter explicações, gerar exemplos de código, refatorar estruturas e depurar erros, enquanto realizava a implementação, teste e gerenciamento de versão.

## Funcionalidades

* Leitura da agenda a partir de um arquivo `agenda_data.csv` customizável.
* Geração de planilha Excel (`.xlsx`) com formatação avançada:
    * Colunas com largura ajustada.
    * Quebra de texto automática.
    * Cabeçalho destacado (negrito, cor de fundo).
    * Coloração da linha inteira baseada no tipo de atividade.
    * Mesclagem vertical das células de dia para melhor visualização.
    * Linha de título na planilha.
    * Legenda de cores explicativa no final da planilha.
    * Painel de cabeçalho congelado.
* Geração de arquivo de calendário `.ics` compatível com diversos aplicativos.
    * Converte horários definidos (`HH:MM - HH:MM`) em eventos com início e fim.
    * Trata entradas sem horário específico (ex: "Manhã") como eventos de dia inteiro.
    * Inclui título e descrição do evento.
* Nomes dos arquivos de entrada (CSV) e saída (Excel, ICS) configuráveis via argumentos de linha de comando (`--csv`, `--output`, `--ics`). *(Verificar se adicionamos o argumento para o ICS)*
* Código Python organizado em funções para melhor leitura e manutenção.
* Uso de `.gitignore` para manter o repositório limpo.

## Tecnologias Utilizadas

* **Linguagem:** Python 3
* **Bibliotecas Principais:**
    * Pandas (Leitura de CSV, manipulação de dados)
    * XlsxWriter (Escrita e formatação avançada de arquivos Excel `.xlsx`)
    * ics.py (Criação de arquivos de calendário `.ics`)
    * Argparse (Processamento de argumentos de linha de comando)
    * NumPy (Dependência do Pandas)
    * datetime (Manipulação de datas e horas)
* **Controle de Versão:** Git
* **Plataforma:** GitHub

## Como Usar

### Pré-requisitos

* Python 3 instalado (disponível em [python.org](https://python.org/))
* Pip (gerenciador de pacotes do Python, geralmente vem junto com a instalação)

### Instalação

1.  **Clone o repositório:**
    ```bash
    git clone https://github.com/zug96/meu-script-agenda
    cd MeusScripts
    ```
2.  **Instale as dependências:**
    ```bash
    pip install pandas xlsxwriter ics.py numpy
    ```
    *(Opcional, mas recomendado: Crie um arquivo `requirements.txt` com essas bibliotecas listadas e rode `pip install -r requirements.txt`)*

### Configuração

1.  **Crie o arquivo `agenda_data.csv`** na mesma pasta do script `criar_agenda.py`.
2.  **Estrutura do CSV:** O arquivo deve ter as colunas `Dia`,`Horario`,`Atividade`,`Observacoes`.
    * `Dia`: Texto descrevendo o dia, preferencialmente incluindo a data no formato `(dd/mm)` para a correta geração do `.ics`. Ex: `Segunda (31/03)`.
    * `Horario`: O horário da atividade. Use o formato `HH:MM - HH:MM` para eventos com duração definida (ex: `08:30 - 12:00`). Para outros, use textos como `Manhã`, `Tarde`, `Noite`, etc. (serão tratados como eventos de dia inteiro no `.ics`).
    * `Atividade`: O nome/título do evento/tarefa. Palavras-chave aqui (como 'Estudo Linux', 'VÔLEI') são usadas para a coloração no Excel.
    * `Observacoes`: Detalhes adicionais sobre a atividade.
3.  **Preencha o `agenda_data.csv`** com os dados da sua semana. Veja o exemplo fornecido no histórico da nossa conversa ou crie o seu.

### Execução

Execute o script Python através do seu terminal (Git Bash, Prompt de Comando, etc.):

* **Para usar os nomes padrão** (`agenda_data.csv` para entrada, `agenda_semana_formatada.xlsx` e `agenda.ics` para saída):
    ```bash
    python criar_agenda.py
    ```
* **Para especificar nomes de arquivos:**
    ```bash
    python criar_agenda.py -c meu_arquivo_agenda.csv -o minha_planilha.xlsx --ics meu_calendario.ics
    ```
    *(Verificar/Adicionar o argumento `--ics` no `argparse` se ainda não foi feito)*
* **Para ver a ajuda:**
    ```bash
    python criar_agenda.py --help
    ```

### Saída

O script gerará os seguintes arquivos na mesma pasta:
* Um arquivo Excel (`.xlsx`) com a agenda formatada.
* Um arquivo de calendário (`.ics`) pronto para importação.

## Jornada de Desenvolvimento (Resumo das Etapas)

Este projeto evoluiu bastante desde a ideia inicial! Passamos por diversas fases, gerenciadas com Git e GitHub:

1.  **Script Inicial:** Criação básica da estrutura para gerar um Excel simples.
2.  **Formatação Avançada (`formatacao_avancada` branch):** Introdução do `XlsxWriter`, adição de cores, mesclagem de células 'Dia', congelamento de painel, ajuste de largura de colunas.
3.  **Refatoração e Dados Externos (`refatoracao_geral` branch):** O código foi reorganizado em funções para melhor clareza e os dados da agenda foram movidos para um arquivo `agenda_data.csv`. Foi praticado o fluxo de Pull Request no GitHub para integrar essas mudanças na `main`.
4.  **Melhorias Visuais (`feature/titulo_excel` branch):** Adição de uma linha de título na planilha Excel e ajuste fino das larguras de coluna. (Integrado via PR).
5.  **Integração `datetime` (`feature/datetime_parse` branch):** O script passou a tentar converter os horários do CSV para objetos `datetime.time`, preparando para funcionalidades futuras. (Integrado via PR).
6.  **Exportação para Calendário (`feature/ics_export` branch):** Adição da funcionalidade para gerar um arquivo `.ics`, permitindo a importação da agenda em outros aplicativos, utilizando a biblioteca `ics.py`. (Integrado via PR).
7.  **Argumentos de Linha de Comando (`feature/output_filename` branch):** Implementação do `argparse` para permitir a customização dos nomes dos arquivos de entrada/saída. (Integrado via PR).
8.  **Documentação (Este README):** Criação deste arquivo para documentar o projeto.

## Próximos Passos (Ideias Futuras)

* Implementar lógica para eventos recorrentes.
* Tratamento mais robusto de datas/horas e fusos horários no `.ics`.
* Desenvolver uma Interface Gráfica (GUI) simples para gerenciar a agenda.
* Adicionar testes automatizados.

## Autor

* Gustavo Correa Campana - https://github.com/zug96/

## Licença

Este projeto está licenciado sob a Licença MIT - veja o arquivo LICENSE.md para detalhes.