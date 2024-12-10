import os
from datetime import datetime
import sys
import pandas as pd


def mapear_colunas(df, mapping):
    """
    Aplica o mapeamento de colunas a um DataFrame específico.
    """
    df = df.rename(columns=mapping)  # Renomeia conforme o mapeamento
    return df


def processar_planilhas(input_file_1, input_file_2=None, output_file=None):
    """
    Processa as duas planilhas fornecidas, faz o mapeamento correto,
    mescla os dados e salva no formato final no caminho especificado.
    """
    try:
        # Ler as planilhas
        df1 = pd.read_excel(input_file_1)
        df2 = pd.read_excel(input_file_2) if input_file_2 else None

        # Mapear colunas da planilha 1
        mapping_1 = {
            "POSICAO ESTOQUE": "STATUS",
            "DATA COMPRA": "DATA DE ENTRADA",
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
            "VENDEDOR": "NOME PROPRIETARIO ENTRADA",
            "CPF/CNPJ CLIENTE": "CPF/CNPJ PROPRIETARIO ENTRADA",
            "KM": "KM",
            "POSICAO ESTOQUE": "STATUS"
        }

        # Aplicar os mapeamentos às planilhas
        df1_mapeado = mapear_colunas(df1, mapping_1)
        df2_mapeado = mapear_colunas(df2, mapping_2) if df2 is not None else None

        # Mesclar os dados
        if df2_mapeado is not None:
            df_combinado = pd.concat([df1_mapeado, df2_mapeado], ignore_index=True)
        else:
            df_combinado = df1_mapeado

        # Obter a data/hora atual
        data_atual = datetime.now()

        # Atualizar onde 'DATA E HORA DE SAIDA' estiver null e STATUS == 1 com a data atual
        df_combinado["DATA E HORA DE SAIDA"] = pd.to_datetime(df_combinado["DATA E HORA DE SAIDA"])
        df_combinado["DATA E HORA DE SAIDA"] = df_combinado.apply(
            lambda row: data_atual if pd.isnull(row["DATA E HORA DE SAIDA"]) and row["STATUS"] == 1 else row["DATA E HORA DE SAIDA"],
            axis=1
        )

        # Criar estrutura final de colunas
        colunas_final = [
            "ID", "STATUS", "DATA DE ENTRADA", "DATA E HORA DE SAIDA", "MARCA", "MODELO",
            "COMPLEMENTO", "CHASSI", "RENAVAM", "NUMERO MOTOR", "ANO FAB.", "ANO MOD", "COR",
            "COMBUSTIVEL", "PLACA", "TIPO", "VALOR COMPRA", "VALOR A VISTA", "VALOR DE VENDA",
            "NOME PROPRIETARIO ENTRADA", "CPF/CNPJ PROPRIETARIO ENTRADA", "KM", "PORTAS",
            "CAMBIO", "CNPJ REVENDA", "ESTADO_CONVERSACAO"
        ]

        # Mapear e ajustar o DataFrame para conter apenas as colunas finais esperadas
        df_final = pd.DataFrame(columns=colunas_final)

        # Mapear os dados ao DataFrame final
        df_final["STATUS"] = df_combinado.get("STATUS", None)
        df_final["DATA DE ENTRADA"] = df_combinado.get("DATA DE ENTRADA", None)
        df_final["DATA E HORA DE SAIDA"] = df_combinado.get("DATA E HORA DE SAIDA", None)
        df_final["MARCA"] = df_combinado.get("MARCA", None)
        df_final["MODELO"] = df_combinado.get("MODELO", None)
        df_final["CHASSI"] = df_combinado.get("CHASSI", None)
        df_final["RENAVAM"] = df_combinado.get("RENAVAM", None)
        df_final["ANO MOD"] = df_combinado.get("ANO MOD", None)
        df_final["COR"] = df_combinado.get("COR", None)
        df_final["COMBUSTIVEL"] = df_combinado.get("COMBUSTIVEL", None)
        df_final["PLACA"] = df_combinado.get("PLACA", None)
        df_final["TIPO"] = df_combinado.get("TIPO", None)
        df_final["ESTADO_CONVERSACAO"] = "USADO"

        # Configurar formatação para saída no Excel
        with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
            df_final.to_excel(writer, index=False, sheet_name='Planilha Final')
            workbook = writer.book
            worksheet = writer.sheets['Planilha Final']

            # Identificar a coluna da data/hora
            if "DATA E HORA DE SAIDA" in df_final.columns:
                col_idx = df_final.columns.get_loc("DATA E HORA DE SAIDA") + 1
                worksheet.set_column(col_idx, col_idx, 20)  # Define largura da coluna
                # Formato para data e hora no LibreOffice
                format_data_hora = workbook.add_format({'num_format': 'dd/mm/yyyy hh:mm:ss'})
                worksheet.set_column(col_idx, col_idx, None, format_data_hora)


        print(f"Arquivo salvo em: {output_file}")
        
    except Exception as e:
        raise Exception(f"Erro ao processar as planilhas: {e}")


if __name__ == "__main__":
    if len(sys.argv) < 3:
        print("Por favor, forneça pelo menos duas planilhas e o caminho de saída.")
        sys.exit(1)

    input_file_1 = sys.argv[1]
    input_file_2 = sys.argv[2]
    output_file = sys.argv[3]

    try:
        processar_planilhas(input_file_1, input_file_2, output_file)
    except Exception as error:
        print(error)