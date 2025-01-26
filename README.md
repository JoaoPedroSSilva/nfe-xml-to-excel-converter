# nfe-xml-to-excel-converter
Ferramenta para converter arquivos XML de NF-e (Nota Fiscal Eletrônica) para uma planilha Excel consolidada.

## Recursos
- **Upload múltiplo:** Envie um ou mais arquivos XML de NF-e de uma vez.
- **Conversão automática:** Extrai dados relevantes e organiza em uma planilha Excel formatada.
- **Download imediato:** Baixe o arquivo Excel gerado diretamente no navegador.

## Tech Stack
- **Front-End**: HTML, CSS, JavaScript, Bootstrap.
- **Back-End**: C# (.NET 6+), ClosedXML.

## Como executar

### Pré-requisitos
- Instale o [SDK do .NET](https://dotnet.microsoft.com/download).
- Instale o Live Server no Visual Studio Code para rodar o front-end.

### Back-End
**1. Abra o projeto no Visual Studio 2022:**
- Navegue até a pasta do back-end e abra o arquivo de solução no Visual Studio.
- Compile o projeto pressionando Ctrl + Shift + B.

**2. Execute o servidor:**
- Pressione F5 para iniciar o servidor.
- O servidor estará acessível em https://localhost:7075.

### Front-End
**1. Navegue até a pasta do front-end:**
- Abra a pasta no Visual Studio Code.

**2. Inicie o Live Server:**
- Clique com o botão direito no arquivo index.html e selecione "Open with Live Server".
- Certifique-se de que o Live Server esteja configurado para usar 127.0.0.1 ou localhost.

### Testes
- Carregue um ou mais arquivos XML de NF-e pelo front-end.
- Certifique-se de que o download da planilha Excel ocorre corretamente.
- Valide os dados na planilha gerada.

## Planejamentos Futuros
- Implementar suporte para diferentes layouts de XML.
- Melhorar o design do front-end.
- Criar documentação completa da API.
