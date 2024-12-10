Processador de Planilhas de Estoque
Este projeto é um script em Python para processamento e transformação de planilhas de estoque de veículos.
O objetivo principal é padronizar os dados de entrada e gerar um arquivo Excel formatado, com colunas renomeadas, valores tratados e estrutura compatível com requisitos específicos.

Funcionalidades
Leitura da planilha de entrada: Carrega uma planilha Excel com os dados de veículos e processa a partir de um cabeçalho.
Renomeação de colunas: Mapeia os nomes das colunas de entrada para nomes padrão.
Transformações específicas:
Ajustes em valores de colunas, como COMBUSTÍVEL e TIPO.
Tratamento de valores nulos e adição de colunas padrão, como ESTADO_CONVERSACAO.
Divisão da coluna MODELO em duas: MODELO e COMPLEMENTO.
Sincronização de valores entre colunas, como ANO FAB. e ANO MOD.
Geração de planilha formatada: Salva os dados processados em um novo arquivo Excel com timestamp no nome.


Como Usar
Pré-requisitos
Python 3.7+
Instale as dependências:
bash
Copiar código
pip install pandas openpyxl pytz
Execução
Prepare o arquivo de entrada:

Crie ou obtenha a planilha de entrada chamada planilha_origem.xlsx.
Certifique-se de que a primeira linha contém os cabeçalhos.
Execute o script:

bash
Copiar código
python nome_do_script.py
O script:

Carrega a planilha planilha_origem.xlsx.
Processa e salva um arquivo de saída no formato ESTOQUE_YYYYMMDD_HHMMSS.xlsx.
Resultado:

O arquivo de saída será gerado na mesma pasta do script.
