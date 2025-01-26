# nfe-xml-to-excel-converter
Ferramenta para converter arquivos XML de NF-e (Nota Fiscal Eletrônica) para uma planilha Excel consolidada.

## Recursos
- Upload de um ou múltiplos arquivos XML de NF-e.
- Conversão automática dos dados relevantes em uma planilha Excel formatada.
- Download instantâneo da planilha gerada.

## Tech Stack
- **Front-End**: HTML, CSS, JavaScript, Bootstrap.
- **Back-End**: C# (.NET 6+), ClosedXML.

## Como executar

### Pré-requisitos
- Instale o [SDK do .NET](https://dotnet.microsoft.com/download).
- (Opcional) Instale um servidor local para o front-end, como o [Live Server](https://marketplace.visualstudio.com/items?itemName=ritwickdey.LiveServer).

### Back-End
1. Navegue até a pasta do projeto:
    ```bash
    cd back-end
    ```
2. Compile e execute:
    ```bash
    dotnet run
    ```
3. O servidor estará disponível em `https://localhost:7075`.

### Front-End
1. Navegue até a pasta do front-end.
2. Abra `index.html` diretamente em um navegador ou use um servidor local como o Live Server.

### Testes
- Carregue um ou mais arquivos XML de NF-e pelo front-end.
- Certifique-se de que o download da planilha Excel ocorre corretamente.
- Valide os dados na planilha gerada.

## Planejamentos Futuros
- Implementar suporte para diferentes layouts de XML.
- Melhorar o design do front-end.
- Criar documentação completa da API.
