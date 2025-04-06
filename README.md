# Projeto Agenda Semanal em Excel e ICS

## Descrição Curta

Este projeto consiste em um script Python que lê uma agenda semanal a partir de um arquivo CSV e gera automaticamente dois arquivos de saída:
1.  Uma planilha Excel (`.xlsx`) bem formatada, com cores, mesclagem de células e legenda.
2.  Um arquivo de calendário (`.ics`) que pode ser importado em aplicativos como Google Agenda, Outlook Calendar, etc., com nome de arquivo configurável.

## Motivação

Este projeto foi desenvolvido com dois objetivos principais:
* Criar uma ferramenta personalizada para organizar e visualizar a agenda semanal de forma prática, integrando compromissos de trabalho, estudo e vida pessoal.
* Aplicar e aprofundar conhecimentos em Python, manipulação de dados, Git, GitHub e o fluxo de trabalho de desenvolvimento, como parte do aprendizado contínuo e especificamente para o Desafio de Projeto "Contribuindo em um Projeto Open Source no GitHub" do Santander Bootcamp em parceria com a DIO.

## Nota sobre o Desenvolvimento:
Este projeto foi desenvolvido em um processo interativo e colaborativo com a Inteligência Artificial Gemini (Modelo Experimental do Google). Como iniciante nos tópicos abordados, utilizei a IA como ferramenta de aprendizado para obter explicações, gerar exemplos de código, refatorar estruturas e depurar erros, enquanto realizava a implementação, teste e gerenciamento de versão.

## Funcionalidades

* Leitura da agenda a partir de um arquivo `.csv` customizável (ex: `agenda_data.csv`, `agenda_google_abril.csv`).
* Geração de planilha Excel (`.xlsx`) com formatação avançada:
    * Colunas com largura ajustada.
    * Quebra de texto automática.
    * Cabeçalho destacado (negrito, cor de fundo).
    * Coloração da linha inteira baseada no tipo de atividade (configurável no script).
    * Mesclagem vertical das células de dia para melhor visualização.
    * Linha de título na planilha.
    * Legenda de cores explicativa no final da planilha.
    * Painel de cabeçalho congelado.
* Geração de arquivo de calendário `.ics` compatível com diversos aplicativos:
    * Converte horários definidos (`HH:MM - HH:MM`) em eventos com início e fim.
    * Trata entradas sem horário específico (ex: "Manhã") como eventos de dia inteiro.
    * Inclui título (Atividade) e descrição (Observacoes) do evento.
    * Utiliza a biblioteca `pytz` para gerar eventos com timezone correto (configurado por padrão para `America/Sao_Paulo` no script - pode necessitar ajuste para outros fusos).
    * Permite configurar o nome do arquivo `.ics` de saída via argumento de linha de comando (`--ics`).
* Nomes dos arquivos de entrada (CSV) e saída (Excel, ICS) configuráveis via argumentos de linha de comando (`--csv`, `--output`, `--ics`).
* Código Python organizado em funções para melhor leitura e manutenção.
* Uso de ambiente virtual (`venv`) recomendado para gerenciamento de dependências.
* Uso de `.gitignore` para manter o repositório limpo (ignorando `.venv`, etc.).

## Tecnologias Utilizadas

* **Linguagem:** Python 3
* **Bibliotecas Principais:**
    * **Pandas:** Para leitura e manipulação eficiente dos dados do arquivo CSV.
    * **XlsxWriter:** Para criar o arquivo Excel `.xlsx` com formatações avançadas.
    * **ics.py:** Para gerar o arquivo de calendário no formato padrão `.ics`.
    * **Argparse:** Para permitir a configuração de nomes de arquivos via linha de comando.
    * **pytz:** Para garantir que os horários no arquivo `.ics` tenham o fuso horário correto.
    * **NumPy:** Dependência do Pandas.
    * **datetime:** Para manipulação interna de datas e horas.
* **Controle de Versão:** Git
* **Plataforma:** GitHub

## Como Usar

### Pré-requisitos

* Python 3 instalado (disponível em [python.org](https://python.org/))
* Pip (gerenciador de pacotes do Python, geralmente vem junto com a instalação)
* Git instalado (disponível em [git-scm.com](https://git-scm.com/))

### Instalação

1.  **Clone o repositório:**
    ```bash
    # Substitua <url_do_seu_repo> pela URL SSH ou HTTPS do seu repositório
    git clone <url_do_seu_repo> MeusScripts
    cd MeusScripts
    ```
2.  **(Recomendado) Crie e Ative um Ambiente Virtual:** Isso isola as dependências do projeto e evita conflitos com o Python do sistema, especialmente em Linux. Veja a seção "Solução de Problemas Comuns" caso encontre erros ao instalar dependências sem um venv.
    ```bash
    # Cria o ambiente virtual (pasta .venv)
    python3 -m venv .venv
    # Ativa o ambiente (Linux/macOS)
    source .venv/bin/activate
    # No Windows use: .venv\Scripts\activate
    ```
    *Lembre-se de ativar o ambiente (`source .venv/bin/activate`) sempre que for trabalhar no projeto em um novo terminal.*
3.  **Instale as dependências (com o ambiente virtual ativo):**
    ```bash
    pip install -r requirements.txt
    ```
    *Adicione `.venv/` ao seu arquivo `.gitignore` (`echo ".venv/" >> .gitignore`) para não enviar essa pasta ao GitHub.*

### Configuração

1.  **Crie seus arquivos `.csv`** na pasta `MeusScripts` (ex: `minha_agenda_detalhada.csv`, `minha_agenda_google.csv`). Você pode ter múltiplos arquivos para diferentes níveis de detalhe ou períodos.
2.  **Estrutura do CSV:** O arquivo deve ter **exatamente 4 colunas** com o cabeçalho `Dia`,`Horario`,`Atividade`,`Observacoes`.
    * `Dia`: Texto descrevendo o dia, **essencial** incluir a data no formato `(dd/mm)` (ex: `Segunda (07/04)`) para a correta geração do `.ics`.
    * `Horario`: Formato `HH:MM - HH:MM` para eventos com duração (ex: `08:30 - 12:00`), ou textos como `Manhã`, `Tarde` para eventos de dia inteiro no `.ics`.
    * `Atividade`: Título do evento. Palavras-chave podem ser usadas para coloração no Excel (configurar no script).
    * `Observacoes`: Detalhes adicionais. **Importante:** Se o texto da observação contiver vírgulas, coloque todo o texto da observação **entre aspas duplas** (`"assim, com vírgula"`) para evitar erros na leitura do arquivo.
3.  **Preencha os CSVs** com os dados da sua semana.

### Execução

Execute o script Python (com o ambiente virtual ativo) através do seu terminal:

* **Para usar os nomes padrão** (`agenda_data.csv` para entrada, `agenda_semana_formatada.xlsx` e `minha_agenda.ics` para saída - *verifique os defaults no script*):
    ```bash
    python3 criar_agenda.py
    ```
* **Para especificar nomes de arquivos (Recomendado):** Use os argumentos `-c` (CSV de entrada), `-o` (Excel de saída, opcional) e `--ics` (ICS de saída).
    ```bash
    # Exemplo gerando ICS detalhado e Excel correspondente
    python3 criar_agenda.py -c minha_agenda_detalhada.csv -o planilha_detalhada.xlsx --ics calendario_detalhado.ics

    # Exemplo gerando ICS simplificado para Google Agenda (sem gerar Excel)
    # Útil para ter uma versão mais limpa para importar em calendários
    python3 criar_agenda.py -c minha_agenda_google.csv --ics calendario_google.ics
    ```
* **Para ver a ajuda:**
    ```bash
    python3 criar_agenda.py --help
    ```

### Saída

O script gerará os seguintes arquivos na mesma pasta (com os nomes que você especificar):
* Um arquivo Excel (`.xlsx`) formatado (se o argumento `-o` for usado ou se for o default).
* Um arquivo de calendário (`.ics`) pronto para importação.

### Utilizando em Máquina Virtual (Transferência de Arquivos)

Se você estiver executando este script dentro de uma máquina virtual Linux (como Ubuntu no VirtualBox) e precisar usar os arquivos gerados (como o `.ics` ou `.xlsx`) no seu sistema operacional principal (Host, como o Windows), será necessário transferi-los da VM para o Host. Algumas maneiras comuns de fazer isso são:

* **Clientes SFTP Gráficos (Recomendado):** Use programas como [WinSCP](https://winscp.net/) ou [FileZilla Client](https://filezilla-project.org/) no seu Windows para conectar-se à VM via SFTP (usando o IP da VM e suas credenciais de login) e transferir arquivos com uma interface gráfica similar ao Windows Explorer.
* **Comando `scp` (Secure Copy):** Se o seu sistema Host (Windows 10/11) tiver o cliente OpenSSH instalado, você pode usar o comando `scp` diretamente no Prompt de Comando ou PowerShell para copiar arquivos da VM. Exemplo (execute no Windows):
    ```bash
    # Substitua <user>@<ip_vm> pelo seu usuário e IP da VM
    # Substitua /caminho/completo/no/windows pela pasta de destino no Windows
    scp <user>@<ip_vm>:/home/<user>/MeusScripts/calendario_google.ics C:\caminho\completo\no\windows\
    ```
* **Pastas Compartilhadas (VirtualBox/VMware):** Se você tiver os "Adicionais para Convidado" (Guest Additions) instalados na VM e tiver configurado uma Pasta Compartilhada nas configurações da VM, essa pasta aparecerá dentro do Linux (geralmente em `/media/sf_nome_da_pasta`) permitindo a cópia direta entre os sistemas.

Lembre-se de obter o endereço IP da sua VM (geralmente com o comando `ip addr show` dentro da VM) para usar os métodos SFTP ou `scp`.

## Solução de Problemas Comuns

Durante a configuração e uso, alguns problemas podem surgir:

* **Erro `externally-managed-environment` ao usar `pip install`:** Este erro em sistemas Linux recentes (Ubuntu/Debian) indica que você não deve instalar pacotes Python globalmente com `pip`. A solução é **sempre** usar um ambiente virtual. Crie-o com `python3 -m venv .venv` e ative-o com `source .venv/bin/activate` antes de instalar dependências ou rodar o script.
* **Erro ao ler CSV (`Expected X fields... saw Y`):** Verifique o arquivo `.csv` na linha indicada pelo erro. Geralmente é causado por uma vírgula extra que cria colunas a mais, ou por texto na coluna `Observacoes` (ou outra) que contém uma vírgula mas **não** está entre aspas duplas (`"`). Corrija a formatação da linha para ter exatamente 4 colunas separadas por 3 vírgulas.
* **Erro `Authentication failed` no `git push` via HTTPS:** O GitHub não aceita mais senhas para autenticação HTTPS. Mude a URL do seu repositório remoto para usar SSH: `git remote set-url origin git@github.com:<seu_usuario>/<seu_repo>.git` e certifique-se de ter configurado chaves SSH na sua conta GitHub e na sua máquina.

## Jornada de Desenvolvimento (Resumo das Etapas)

Este projeto evoluiu bastante desde a ideia inicial! Passamos por diversas fases, gerenciadas com Git e GitHub:

1.  **Script Inicial:** Criação básica da estrutura para gerar um Excel simples.
2.  **Formatação Avançada (`formatacao_avancada` branch):** Introdução do `XlsxWriter`, adição de cores, mesclagem de células 'Dia', congelamento de painel, ajuste de largura de colunas.
3.  **Refatoração e Dados Externos (`refatoracao_geral` branch):** O código foi reorganizado em funções, dados movidos para `agenda_data.csv`, praticado fluxo de Pull Request.
4.  **Melhorias Visuais (`feature/titulo_excel` branch):** Adição de título na planilha Excel. (Integrado via PR).
5.  **Integração `datetime` (`feature/datetime_parse` branch):** Tentativa de converter horários do CSV para objetos `datetime.time`. (Integrado via PR).
6.  **Exportação para Calendário (`feature/ics_export` branch):** Adição da funcionalidade para gerar arquivo `.ics` com `ics.py`. (Integrado via PR).
7.  **Argumentos de Linha de Comando (`feature/output_filename` branch):** Implementação do `argparse` para customizar nomes dos arquivos de entrada/saída Excel. (Integrado via PR).
8.  **Argumento para Saída ICS (`feature/ics-argument` branch):** Adição do argumento `--ics` para permitir nomear o arquivo de calendário `.ics` gerado.
9.  **Documentação (Este README):** Criação e atualização deste arquivo para documentar o projeto, incluindo notas sobre uso em VM e solução de problemas comuns.

## Próximos Passos (Ideias Futuras)

* Implementar lógica para eventos recorrentes no `.ics`.
* Tratamento mais robusto de datas que cruzam o ano no `.ics`.
* Desenvolver uma Interface Gráfica (GUI) simples.
* Adicionar testes automatizados.
* Permitir configuração de cores/formatos do Excel via arquivo externo.

## Autor

* Gustavo Correa Campana - https://github.com/zug96/

## Licença

Este projeto está licenciado sob a Licença MIT - veja o arquivo LICENSE.md para detalhes.
