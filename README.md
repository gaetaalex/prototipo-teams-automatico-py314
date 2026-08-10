# Controle automático pelo Teams — Python 3.14

Aplicação local que autentica cada operador no Microsoft Entra, localiza o destinatário pelo `employeeId`, envia a mensagem pelo Microsoft Graph e atualiza a planilha após a confirmação da Microsoft.

Esta variante foi preparada para **CPython 3.14 convencional de 64 bits no Windows**. Ela recusa deliberadamente uma instalação free-threaded para evitar misturar ambientes diferentes durante a homologação.

## Segurança

- Nunca inclua o arquivo `.env` no Git ou em chamados.
- Nunca utilize dados reais de clientes sem autorização.
- Planilhas, tokens, ambientes virtuais e arquivos operacionais são ignorados pelo Git.
- Em ambiente corporativo, utilize somente dependências e repositórios homologados pela TI.

## Dados necessários da TI

- `TENANT_ID`
- `CLIENT_ID`
- `CLIENT_SECRET`
- URI de redirecionamento cadastrada no Entra
- Permissões delegadas: `User.Read`, `User.Read.All`, `Chat.Create` e `ChatMessage.Send`
- Consentimento administrativo quando exigido pela política da organização

Copie `.env.example` para `.env` e preencha somente na máquina autorizada.

## Instalação normal

Execute:

```bat
iniciar.bat
```

O script localiza o Python 3.14, cria `.venv`, instala as dependências e inicia o endereço `http://127.0.0.1:5056`.

## Instalação em rede sem acesso ao PyPI

Em uma máquina Windows com internet, execute:

```bat
preparar_pacotes_offline.bat
```

Isso cria `pacotes_offline` com os arquivos `.whl` destinados ao CPython 3.14 convencional de 64 bits. Transfira a pasta completa pelos meios aprovados pela empresa. Na máquina corporativa, execute `iniciar.bat`; ele dará preferência aos pacotes locais.

Se a transferência de pacotes externos não for permitida, solicite à TI a URL do repositório Python corporativo e a lista de versões homologadas.

## Verificação rápida

Depois da instalação:

```bat
.venv\Scripts\python.exe -c "import flask,openpyxl,msal,requests,dotenv,cryptography; print('Dependencias OK')"
```

## Planilha reconhecida

- `MATRICULA`, utilizada como `employeeId`
- `USUARIO` ou `NOME`
- `TEXTO DE CONTATO`
- `TEXTO SEGUNDO CONTATO`
- `STATUS ENVIO`
- `NUMERO DE CONTATO`
- `ULTIMA DATA DE CONTATO`

Qualquer conteúdo em `STATUS ENVIO` bloqueia um novo envio. O segundo texto é usado quando `NUMERO DE CONTATO` já é 1.
