# Script Automatizado de Relatório de Compartilhamento do Drive

Este guia detalha todos os passos necessários para configurar e utilizar o script que audita os arquivos compartilhados no seu Google Drive.

## O que o Script Faz?

Este script foi projetado para ser uma solução "**configure e esqueça**". Seu objetivo é analisar todos os arquivos e pastas de sua propriedade no Google Drive para identificar quais estão compartilhados.

* **Execução em Ciclos**: Para evitar o limite de tempo de execução do Google (6 minutos), o script trabalha em ciclos curtos, processando uma parte dos seus arquivos de cada vez.
* **Automação Completa**: Ele cria e apaga seus próprios acionadores (triggers) para se autoexecutar em segundo plano.
* **Relatório Final Anexado**: Ao final de um ciclo completo de verificação, ele envia um único e-mail com o relatório final em um arquivo `.zip` anexado, e então se agenda para recomeçar o processo em 7 dias.

## Passo 1: Configuração Inicial (Pré-requisitos)

Antes de colar o código, duas configurações são essenciais no seu projeto do Apps Script.

### A. Habilitar a API Avançada do Drive

O script precisa de acesso rápido à sua lista de arquivos, o que é feito pela API do Drive.

1.  No editor do Apps Script, clique em **Serviços** no menu à esquerda (ícone de **+**).
2.  Na lista que aparece, encontre **Drive API** e clique nela.
3.  Clique no botão **Adicionar**.

### B. Configurar o Arquivo de Manifesto (`appsscript.json`)

Este arquivo define as permissões que o script solicitará.

1.  Clique no ícone de **Configurações do projeto** (a engrenagem ⚙️) no menu à esquerda.
2.  Marque a caixa de seleção "**Mostrar o arquivo de manifesto 'appsscript.json' no editor**".
3.  Volte para o **Editor** (ícone </>). Você verá um novo arquivo chamado `appsscript.json`. Clique nele.
4.  Apague todo o conteúdo que estiver lá e cole o seguinte código:
    ```json
    {
      "timeZone": "America/Sao_Paulo",
      "dependencies": {
        "enabledAdvancedServices": [{
          "userSymbol": "Drive",
          "serviceId": "drive",
          "version": "v2"
        }]
      },
      "exceptionLogging": "STACKDRIVER",
      "runtimeVersion": "V8",
      "oauthScopes": [
        "[https://www.googleapis.com/auth/drive.readonly](https://www.googleapis.com/auth/drive.readonly)",
        "[https://www.googleapis.com/auth/script.send_mail](https://www.googleapis.com/auth/script.send_mail)",
        "[https://www.googleapis.com/auth/userinfo.email](https://www.googleapis.com/auth/userinfo.email)",
        "[https://www.googleapis.com/auth/script.scriptapp](https://www.googleapis.com/auth/script.scriptapp)"
      ]
    }
    ```
5.  Salve o arquivo.

## Passo 2: Instalação e Primeira Execução

Com a configuração pronta, agora vamos instalar e iniciar o script.

1.  **Cole o Código Principal**: Abra o arquivo `Código.gs` e cole o código completo do script fornecido no Canvas.
2.  **Salve o Projeto**: Clique no ícone de disquete para salvar.
3.  **Execute Manualmente (Apenas Uma Vez)**:
    1.  Na barra de ferramentas do editor, ao lado do botão "Depurar", clique no menu suspenso de funções.
    2.  Selecione a função `gerenciarCicloDeRelatorio`.
    3.  Clique em **Executar**.
4.  **Autorize as Permissões**:
    1.  Uma janela pop-up solicitará sua permissão. Clique em **Revisar permissões**.
    2.  Escolha sua conta do Google.
    3.  Você verá um aviso de "O Google não verificou este app". Clique em **Avançado**