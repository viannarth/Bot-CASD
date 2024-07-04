# Manual de utilização do código de envio automático de mensagens
## Pré-requisitos
1. **Python**: Certifique-se de ter o Python 3.7+ e instalado no seu computador.


## Autenticação do usuário no Google Cloud Console
1. **Acesse o Google Cloud Console**: Vá para [Google Cloud Console](https://console.cloud.google.com/).
   
2. **Crie um novo projeto**:
   - Clique em "Select a project" (ou "Selecione um projeto") no topo da página.
   - Clique em "New Project" (ou "Novo Projeto").
   - Dê um nome ao projeto e clique em "Create" (ou "Criar").
     
3. **Habilite a API do Google Drive**:
   - No painel do projeto, vá para "APIs & Services" (ou "APIs e Serviços").
   - Clique em "Library" (ou "Biblioteca").
   - Procure por "Google Drive API" e clique em "Enable" (ou "Habilitar").
     
4. **Configure a Tela de Consentimento OAuth**:
   - No painel do projeto, acessando-se o menu lateral, vá para "APIs & Services" > "OAuth consent screen" (ou "APIs e Serviços" > "Tela de consentimento OAuth").
   - Selecione "External" (ou "Externo") e clique em "Create" (ou "Criar").
   - Preencha os campos obrigatórios e salve. As informações concedidas, como o nome, não serão relevantes a fins do programa. Não é necessário adicionar nenhum escopo manualmente.
     
5. **Crie Credenciais OAuth**:
   - No painel do projeto, acessando-se o menu lateral, vá para "APIs & Services" > "Credentials" (ou "APIs e Serviços" > "Credenciais").
   - Clique em "Create Credentials" (ou "Criar Credenciais") e selecione "OAuth client ID" (ou "ID de cliente OAuth").
   - Configure o tipo de aplicação como "Desktop App" (ou "Aplicativo de Desktop").
   - Clique em "Create" (ou "Criar").
   - Faça o download do arquivo JSON. Este arquivo contém suas credenciais do cliente OAuth 2.0.


## Configuração do ambiente
1. Crie uma pasta em um local de preferência chamado "bot_casd" ou semelhante, a fins de organização.
   
2. Adicione à pasta os arquivos `bot.py`, `download_planilha.py` e `requirements.txt`.
   
3. Renomeie o arquivo JSON baixado na seção anterior para `client_secrets.json` e o mova para a pasta criada.
   
4. Abra o terminal na pasta criada. Isso pode ser feita ao clicar com o botão direito sobre a pasta e clicar em "Open in Terminal" (ou "Abrir no Terminal") ou semelhantes.
   
5. *(Opcional)* Antes de adicionar as bibliotecas, recomenda-se a criação de um *virtual environment* para evitar conflitos de versões das dependências com outros projetos. Isso pode ser feito pelas seguintes linhas de código no terminal:
   ```python
   python -m venv env # env pode ser substituído pelo nome à sua escolha
   source env/bin/activate # No Windows, use `env\Scripts\activate`
   ```
   Você verá o nome do seu ambiente virtual aparecer no prompt de comando, indicando que ele está ativado.
   
6. No terminal, digite a seguinte linha de código:
   ```python
   pip install -r requirements.txt
   ```


## Primeira utilização do código para download da planilha
1. **Inserir ID da planilha**: no Google Drive, acesse a planilha que deseja baixar (você não precisa ser o proprietário da planilha) e, no URL, identifique o ID da planilha.
   Exemplo: para a URL **`https://docs.google.com/spreadsheets/d/1A2B3C4D5E6F7G8H9I0J1K2L3M4N5O6P7Q8R9S0T1U2V/edit`**, o ID é **`1A2B3C4D5E6F7G8H9I0J1K2L3M4N5O6P7Q8R9S0T1U2V`**.
  
2. **Executar o script**: Rode o script, por meio da linha de código:
   ```python
   python download_planilha.py
   ```
   
3. **Autenticação pela web**: Ao rodar o script, você será redirecionado para uma tela de login com sua Conta Google para autorize o acesso à API do Google Drive.
   **OBS.**: Caso o Google sinalize que o app não foi verificado, prossiga em "Advanced" (ou "Avançado") e "Go to `nome do app` (unsafe)" (ou "Ir para `nome do app` (não seguro)) para continuar.
   
5. **Verificação**: Se tudo estiver configurado corretamente, você verá a mensagem "O download foi realizado com sucesso" no terminal e o arquivo Controle de Presença 2024.xlsx será baixado na mesma pasta do script.
   
6. **Possíveis erros**: Caso o procedimento dê errado, haverá uma mensagem de erro no terminal. Avalie-a e tente corrigir a partir dela. Em geral, atente-se para os seguintes detalhes:
   - Verifique se o arquivo client_secrets.json está corretamente configurado e localizado na mesma pasta do script.
   - Verifique as configurações de OAuth no Google Cloud Console.
   - Verifique se o ID da planilha está correto.
   - Certifique-se de que você tem permissão para acessar o arquivo no Google Drive.
  

## Utilização do bot
Após o download da planilha com o nome `Controle de Presença 2024.xlsx` na pasta do script, resta a breve configuração inicial do script `bot.py`. Antes de rodar o script, modifique a semana e o bimestre atuais no início do código, nas linhas 8 e 9:
```python
BIMESTRE = # Insira o bimestre atual aqui
SEMANA = # Insira a semana atual aqui
```
Depois, basta o rodar o script do `bot.py`, digitando a seguinte linha de código no terminal:
```python
python bot.py
```
**Nota**: Na primeira utilização, haverá um período de tempo de 60 s para login inicial do usuário no WhatsApp Web, caso não tenha feito antes. Após o login inicial, pode-se comentar as linhas 19 e 20 do código para retirar esse tempo inicial e agilizar o processo. O comentário pode ser feito ao se inserir `#` no início de cada linha ou ao se selecionarem as linhas e pressionar `ctrl`+`/`.
