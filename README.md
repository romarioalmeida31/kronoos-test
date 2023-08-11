# Teste de vaga fullstack *KRONOOS*

## Manipulação de Dados de CSV e Conversão para Array

- Os dados são fornecidos em formato CSV.
- Utilizaremos a biblioteca `fs` (File System) para ler o arquivo CSV e a biblioteca `csv-parser` para processar os dados e convertê-los em um array de objetos JavaScript.

## Conversão de Dados para Moeda Real Brasileira

- Valores monetários, como `vlTotal`, `vlPresta`, `vlMora`, etc., precisam ser formatados como moeda brasileira (BRL).
- Utilizaremos a biblioteca `intl` do JavaScript para formatar os valores numéricos como moeda BRL, incluindo o símbolo de real (R$), separador de milhar e precisão de duas casas decimais.

## Validação de CPF ou CNPJ

- Implementaremos uma função para validar o campo `nrCpfCnpj` e verificar se ele é um CPF ou CNPJ válido, seguindo as regras de validação apropriadas para cada formato.

## Validação de Valor Total e Prestações

- Dividiremos o valor de `vlTotal` pela quantidade de prestações (`qtPrestacoes`).
- Verificaremos se o resultado dessa divisão é igual ao valor de `vlPresta` para cada prestação, garantindo que os cálculos estejam corretos e consistentes.

## Conversão de Datas para o Tipo Date

- Os campos `dtContrato` e `dtVctPre` estão no formato `YYYYMMDD`.
- Utilizaremos o JavaScript para converter esses campos em objetos do tipo `Date`, permitindo manipulações e formatações mais adequadas.

Certifique-se de testar cada etapa do processo para garantir que os dados sejam manipulados, convertidos e validados corretamente de acordo com os cenários especificados.