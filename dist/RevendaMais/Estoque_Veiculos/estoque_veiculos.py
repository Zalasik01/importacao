import os
from datetime import datetime
import sys
import pandas as pd


def mapear_colunas(df, mapping):
    """
    Aplica o mapeamento de colunas a um DataFrame específico.
    """
    return df.rename(columns=mapping)  # Renomeia conforme o mapeamento


def processar_planilhas(input_file_1, input_file_2=None, output_file=None):
    """
    Processa as duas planilhas fornecidas, faz o mapeamento correto,
    mescla os dados e salva no formato final no caminho especificado.
    """
    try:
        print("Iniciando processamento...")

        # Carregar planilhas
        print(f"Lendo planilha 1: {input_file_1}")
        df1 = pd.read_excel(input_file_1)
        print(f"Planilha 1 carregada com {len(df1)} linhas.")

        if input_file_2:
            print(f"Lendo planilha 2: {input_file_2}")
            df2 = pd.read_excel(input_file_2)
            print(f"Planilha 2 carregada com {len(df2)} linhas.")
        else:
            df2 = None
            print("Nenhuma planilha 2 fornecida.")

        # Mapear colunas da planilha 1
        mapping_1 = {
            "POSICAO ESTOQUE": "STATUS",
            "DATA COMPRA": "DATA DE ENTRADA",
            "DATA VENDA": "DATA E HORA DE SAIDA",
            "MARCA": "MARCA",
            "MODELO": "MODELO",
            "CHASSI": "CHASSI",
            "RENAVAM": "RENAVAM",
            "ANO MODELO": "ANO MOD",
            "COR": "COR",
            "COMBUSTIVEL": "COMBUSTIVEL",
            "PLACA": "PLACA",
            "TIPO": "TIPO",
            "VALOR COMPRA": "VALOR COMPRA",
            "VALOR VENDA": "VALOR DE VENDA",
            "KM": "KM",
            "FORNECEDOR": "NOME PROPRIETARIO ENTRADA",
            "CPF/CNPJ FORNECEDOR": "CPF/CNPJ PROPRIETARIO ENTRADA"
        }

        # Mapear colunas da planilha 2
        mapping_2 = {
            "DATA VENDA": "DATA E HORA DE SAIDA",
            "MARCA": "MARCA",
            "MODELO": "MODELO",
            "CHASSI": "CHASSI",
            "RENAVAM": "RENAVAM",
            "ANO MODELO": "ANO MOD",
            "COR": "COR",
            "COMBUSTIVEL": "COMBUSTIVEL",
            "PLACA": "PLACA",
            "TIPO": "TIPO",
            "VALOR COMPRA": "VALOR COMPRA",
            "VALOR VENDA": "VALOR DE VENDA",
            "CLIENTE": "NOME PROPRIETARIO ENTRADA",
            "CPF/CNPJ CLIENTE": "CPF/CNPJ PROPRIETARIO ENTRADA",
            "KM": "KM",
            "POSICAO ESTOQUE": "STATUS"
        }

        # Aplicar os mapeamentos
        df1_mapeado = mapear_colunas(df1, mapping_1)
        df2_mapeado = mapear_colunas(df2, mapping_2) if df2 is not None else None

        # Mesclar os dados
        if df2_mapeado is not None:
            df_combinado = pd.concat([df1_mapeado, df2_mapeado], ignore_index=True)

        else:
            df_combinado = df1_mapeado

        # Processar e substituir valores importantes
        df_combinado["COMBUSTIVEL"] = df_combinado["COMBUSTIVEL"].str.strip().str.upper().replace({
            "FLEX": "ALCOOL/GASOLINA",
        })

        df_combinado["TIPO"] = df_combinado["TIPO"].str.strip().str.upper().replace({
            "PROPRIO": "PROPRIO",
            "CONSIGNADO": "TERCEIRO_CONSIGNADO",
        })

        # Estrutura final
        colunas_final = [
            "ID", "STATUS", "DATA DE ENTRADA", "DATA E HORA DE SAIDA", "MARCA", "MODELO",
            "COMPLEMENTO", "CHASSI", "RENAVAM", "NUMERO MOTOR", "ANO FAB.", "ANO MOD", "COR",
            "COMBUSTIVEL", "PLACA", "TIPO", "VALOR COMPRA", "VALOR A VISTA", "VALOR DE VENDA",
            "NOME PROPRIETARIO ENTRADA", "CPF/CNPJ PROPRIETARIO ENTRADA", "KM", "PORTAS",
            "CAMBIO", "CNPJ REVENDA", "ESTADO_CONVERSACAO"
        ]
        df_final = pd.DataFrame(columns=colunas_final)

        # Preenchendo DataFrame final
        for col in colunas_final:
            df_final[col] = df_combinado.get(col, None)

        df_final["ESTADO_CONVERSACAO"] = "USADO"
        df_final["VALOR A VISTA"] = df_final.apply(
            lambda row: row["VALOR DE VENDA"] if pd.isnull(row["VALOR A VISTA"]) else row["VALOR A VISTA"],
            axis=1
        )
        df_final["ANO FAB."] = df_final["ANO MOD"]

        # Dividindo MODELO em MODELO e COMPLEMENTO
        df_final[["MODELO", "COMPLEMENTO"]] = df_final["MODELO"].str.split(n=1, expand=True)

        # Salvar no Excel
        with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
            df_final.to_excel(writer, index=False, sheet_name='Planilha Final')
            workbook = writer.book
            worksheet = writer.sheets['Planilha Final']

        print(f"Arquivo gerado com sucesso em: {output_file}")

    except Exception as e:
        raise Exception(f"Erro ao processar as planilhas: {e}")


if __name__ == "__main__":
    if len(sys.argv) < 3:
        print("Por favor, forneça pelo menos uma planilha de entrada e o caminho de saída.")
        print("Uso: python script.py <input_file_1> [<input_file_2>] <output_file>")
        sys.exit(1)

    input_file_1 = sys.argv[1]
    input_file_2 = sys.argv[2] if len(sys.argv) > 3 else None
    output_file = sys.argv[-1]

    try:
        if not output_file.endswith('.xlsx'):
            raise ValueError("O arquivo de saída deve ter a extensão .xlsx")

        processar_planilhas(input_file_1, input_file_2, output_file)
    except Exception as error:
        print(f"Erro encontrado: {error}")
