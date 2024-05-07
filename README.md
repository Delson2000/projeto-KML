# Controle de Contratos - Geração de Tabelas HTML em Arquivos KML

Este script Python automatiza o processo de geração de tabelas HTML contendo informações sobre contratos e os insere em arquivos KML para visualização em ferramentas de geolocalização, como o Google Earth.

## Funcionalidades

- **Leitura de Planilhas de Controle de Contratos**: O script lê planilhas Excel que contêm informações sobre contratos, incluindo obras em execução, obras concluídas, licitações, serviços contínuos e obras paralisadas.

- **Processamento de Dados**: Os dados são processados para limpeza e formatação, incluindo a conversão de datas e valores monetários para o formato adequado.

- **Geração de Tabelas HTML**: Com base nos dados dos contratos, o script gera tabelas HTML detalhadas contendo informações relevantes, como objeto do contrato, contratada, valor do contrato, datas de início e término, responsáveis, fiscais, gestor e status do contrato.

- **Incorporação em Arquivos KML**: As tabelas HTML são incorporadas em arquivos KML, que são formatos de arquivo usados para exibir dados geográficos em aplicativos como o Google Earth. As tabelas são inseridas como descrições em pontos específicos do mapa, fornecendo uma representação visual dos contratos em suas localizações geográficas.

## Como Usar

1. **Preparação dos Dados**: Certifique-se de que os dados dos contratos estão organizados em planilhas Excel, com informações relevantes devidamente preenchidas.

2. **Configuração do Ambiente**: Instale os requisitos necessários, como pandas e pywin32

3. **Execução do Script**: Execute o script Python `main.py`, que processará os dados dos contratos, gerará tabelas HTML e as incorporará em arquivos KML.

4. **Visualização dos Resultados**: Abra os arquivos KML gerados em ferramentas de visualização de mapas, como o Google Earth, para visualizar as tabelas HTML incorporadas em pontos específicos do mapa, representando os contratos.


