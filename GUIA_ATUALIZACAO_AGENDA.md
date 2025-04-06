# Guia Rápido: Atualizando a Agenda Semanal/Mensal

Este guia resume os passos essenciais para atualizar sua agenda para semanas ou meses futuros usando o script `criar_agenda.py` e os arquivos CSV. Execute estes passos principalmente dentro da sua VM Ubuntu via Putty (ou no seu ambiente Linux configurado).

1.  **Planeje a Nova Semana/Mês:**
    * Defina seus compromissos fixos (Trabalho, Aulas, Terapia, Vôlei).
    * Verifique os fins de semana de trabalho/folga corretos para o período.
    * Consulte os cronogramas da faculdade para tarefas e leituras.
    * Aloque blocos de estudo (lembre-se da transição Linux -> Pedagogia após ~27/04).
    * Inclua tempo pessoal e com a namorada.

2.  **Conecte à VM e Navegue:**
    * Acesse sua VM Ubuntu via Putty (ou terminal equivalente).
    * Navegue até a pasta do seu projeto:
        ```bash
        cd ~/MeusScripts 
        # (Ajuste o caminho se necessário)
        ```

3.  **Edite o Arquivo CSV:**
    * Escolha o arquivo CSV que contém a agenda a ser atualizada (provavelmente o simplificado, ex: `agenda_google_abril.csv`).
    * **Opção A:** Edite o arquivo existente:
        ```bash
        nano agenda_google_abril.csv 
        ```
        Adicione as linhas para o novo período, **certificando-se do formato** `Dia (dd/mm),HH:MM - HH:MM,Atividade,"Observacao com virgula"`. Apague as semanas antigas se desejar.
    * **Opção B:** Crie um novo arquivo para o mês (recomendado para histórico):
        ```bash
        cp agenda_google_abril.csv agenda_google_maio.csv 
        nano agenda_google_maio.csv 
        ```
        Edite o novo arquivo com os dados do novo mês.
    * Salve as alterações no `nano` (`Ctrl+O`, `Enter`) e saia (`Ctrl+X`).

4.  **(Opcional, mas recomendado) Registre a Mudança no Git:**
    * Verifique o arquivo modificado/novo:
        ```bash
        git status
        ```
    * Adicione o arquivo ao stage:
        ```bash
        # Use o nome do arquivo que você editou/criou
        git add nome_do_arquivo.csv 
        ```
    * Faça o commit:
        ```bash
        git commit -m "Data: Atualiza agenda para [Semana/Mês X]"
        ```
    * Envie para o GitHub (opcional, para backup):
        ```bash
        git push
        ```

5.  **Ative o Ambiente Virtual:**
    * **Essencial** para que o script encontre as bibliotecas corretas.
        ```bash
        source .venv/bin/activate 
        ```
    * Verifique se `(.venv)` aparece no início do prompt.

6.  **Execute o Script para Gerar o Novo ICS:**
    * Rode `python3`, especificando o **CSV atualizado** (`-c`) e o **nome do arquivo ICS** que você quer gerar (`--ics`).
        ```bash
        # Exemplo: usando o arquivo de Maio para gerar um ICS de Maio
        python3 criar_agenda.py -c agenda_google_maio.csv --ics calendario_google_maio.ics 

        # Ou, se editou o de Abril e quer gerar/substituir o ICS de Abril:
        # python3 criar_agenda.py -c agenda_google_abril.csv --ics calendario_google.ics 
        ```
    * Verifique se o script roda sem erros e confirma a geração do arquivo `.ics`.

7.  **Desative o Ambiente Virtual:**
    ```bash
    deactivate
    ```

8.  **Transfira o Novo Arquivo `.ics` para o Windows:**
    * Use seu método preferido (`scp` no Prompt/PowerShell do Windows, WinSCP, FileZilla) para copiar o arquivo `.ics` recém-gerado (ex: `calendario_google_maio.ics`) da pasta `~/MeusScripts` na VM para o seu computador Windows (ex: pasta Downloads).

9.  **Importe no Google Agenda:**
    * No Google Agenda (versão web), vá em Configurações (⚙️) > Importar e exportar.
    * Clique em "Selecionar arquivo do seu computador" e escolha o arquivo `.ics` que você acabou de transferir.
    * Escolha a agenda de destino e clique em "Importar".
    * **Atenção:** Verifique como o Google lida com eventos existentes para evitar duplicatas. Pode ser útil limpar eventos antigos do período correspondente antes de importar, ou usar um calendário temporário para importar e verificar.

---
