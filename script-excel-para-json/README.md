# README.md

# Script para Converter Planilhas Excel em JSON

Este projeto contém um script que lê planilhas do arquivo `prescricoes.xlsx`, converte os dados em formato JSON e salva os arquivos JSON gerados na pasta `output`.

## Estrutura do Projeto

- **src/index.ts**: Ponto de entrada do script, contém a lógica para leitura e conversão dos dados.
- **src/types/index.ts**: Exporta interfaces que definem a estrutura dos dados lidos do Excel e do JSON resultante.
- **input/prescricoes.xlsx**: Arquivo Excel que contém as planilhas a serem processadas.
- **output/**: Pasta onde os arquivos JSON gerados serão armazenados.
- **tsconfig.json**: Configuração do TypeScript.
- **package.json**: Configuração do npm, incluindo dependências e scripts.

## Como Executar o Script

1. Certifique-se de ter o Node.js e o npm instalados.
2. Clone o repositório ou baixe os arquivos do projeto.
3. Navegue até a pasta do projeto no terminal.
4. Instale as dependências com o comando:
   ```
   npm install
   ```
5. Execute o script com o comando:
   ```
   npm start
   ```

Os arquivos JSON gerados estarão disponíveis na pasta `output`.

## Dependências

- [xlsx](https://www.npmjs.com/package/xlsx): Para ler arquivos Excel.
- [fs](https://nodejs.org/api/fs.html): Para manipulação de arquivos no sistema.

## Contribuições

Contribuições são bem-vindas! Sinta-se à vontade para abrir issues ou pull requests.