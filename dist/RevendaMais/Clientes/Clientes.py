import os
from datetime import datetime
import sys
import pandas as pd


def mapear_colunas(df):
    """
    Mapeia e renomeia as colunas conforme o formato desejado.
    """
    mapping = {
        "cpf_cnpj": "CPFCNPJ",
        "pessoa": "Pessoa",
        "sexo": "Sexo",
        "nome": "Nome Completo",
        "telefone_celular": "Telefone1",
        "telefone_residencial": "Telefone2",
        "telefone_comercial": "Telefone3",
        "rg": "IE/RG",
        "ie": "IE/RG",
        "data_nascimento": "Data Nascimento",
        "data_cadastro": "Data Cadastro",
        "email": "Email",
        "cep": "Cep",
        "rua": "Rua",
        "numero": "Numero",
        "bairro": "Bairro",
        "estado": "UF",
        "cidade": "Cidade",
        "complemento": "Complemento",
        "apelido": "Apelido",
        "cargo": "Fornecedor",
        "grupos": "Grupos",
        "nome_mae": "Nome Mae",
        "nome_pai": "Nome Pai",
        "data_ultima_compra": "Data Última Compra",
        "quantidade_veic_comprados": "Quantidade Veículos Comprados",
        "clie_cod": "Código Cliente",
        "reve_cod": "Código Revenda",
    }

    # Rename the DataFrame columns based on the mapping
    df = df.rename(columns=mapping)
    return df


def processar_planilha(input_file, output_file):
    """
    Processa a planilha conforme as colunas desejadas.
    Apaga colunas desnecessárias e salva no formato desejado.
    """
    try:
        # Ler a planilha de entrada
        df = pd.read_excel(input_file)

        # Aplicar o mapeamento de colunas
        df_mapeado = mapear_colunas(df)

        # Selecionar apenas as colunas necessárias
        colunas_desejadas = [
            "Código Cliente", "Pessoa", "Sexo", "Nome Completo", "Apelido", "CPFCNPJ", "Email", 
            "Cep", "Rua", "Numero", "Complemento", "Bairro", "Cidade", "UF", 
            "Data Nascimento", "IE/RG", "Telefone1", "Telefone2", "Telefone3", 
            "Tipo Telefone 1", "Tipo Telefone 2", "Tipo Telefone 3", "Fornecedor"
        ]
        
        # Apenas manter as colunas desejadas
        df_final = df_mapeado[colunas_desejadas]

        # Tratar valores nulos, se necessário
        df_final.fillna("", inplace=True)

        # Configurar a saída no Excel com formatação
        with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
            df_final.to_excel(writer, index=False, sheet_name='Planilha Final')
            workbook = writer.book
            worksheet = writer.sheets['Planilha Final']

            # Configuração para largura automática para melhor visualização
            for i, col in enumerate(df_final.columns):
                worksheet.set_column(i, i, 20)

        print(f"Arquivo salvo em: {output_file}")
        
    except Exception as e:
        raise Exception(f"Erro ao processar a planilha: {e}")


if __name__ == "__main__":
    if len(sys.argv) < 3:
        print("Por favor, forneça uma planilha e o caminho de saída.")
        sys.exit(1)

    input_file = sys.argv[1]
    output_file = sys.argv[2]

    try:
        processar_planilha(input_file, output_file)
    except Exception as error:
        print(error)
