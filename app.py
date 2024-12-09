import pandas as pd
from datetime import datetime
import pytz  # Para lidar com o fuso horário


def process_planilha(input_file):
    # Capturando a data e hora no formato correto de Brasília
    now = datetime.now(pytz.timezone("America/Sao_Paulo"))
    timestamp = now.strftime("%Y%m%d_%H%M%S")  # Formato: AAAAMMDD_HHMMSS
    output_file = f"ESTOQUE_{timestamp}.xlsx"

    # Ler a planilha e considerar a primeira linha como cabeçalho
    df = pd.read_excel(input_file)

    # Remover espaços em branco nas colunas
    df.columns = df.columns.str.strip()

    # Ajustar VALOR COMPRA, VALOR DE VENDA e VALOR A VISTA antes da remoção
    if all(col in df.columns for col in ["O", "P", "Q", "R"]):
        df["VALOR COMPRA"] = pd.to_numeric(df["O"], errors="coerce").fillna(0).map(lambda x: f"{x:.2f}")
        df["VALOR DE VENDA"] = pd.to_numeric(df["P"], errors="coerce").fillna(0).map(lambda x: f"{x:.2f}")
        df["VALOR A VISTA"] = pd.to_numeric(df["Q"], errors="coerce").fillna(0).map(lambda x: f"{x:.2f}")
    else:
        print("Alguma das colunas O, P, Q ou R não foi encontrada no DataFrame.")

    # Mapear as colunas da planilha de origem para destino
    colunas_destino = {
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
        "FORNECEDOR": "NOME PROPRIETARIO ENTRADA",
        "CPF/CNPJ FORNECEDOR": "CPF/CNPJ PROPRIETARIO ENTRADA",
        "KM": "KM"
    }

    # Renomear colunas
    df_renomeado = df.rename(columns=colunas_destino)

    # Tratamento COMBUSTIVEL
    if "COMBUSTIVEL" in df.columns:
        df_renomeado["COMBUSTIVEL"] = df["COMBUSTIVEL"].replace({
            "FLEX": "ALCOOL/GASOLINA"
        })
    else:
        print("A coluna 'COMBUSTIVEL' não foi encontrada no DataFrame.")

    # Tratamento TIPO da coluna 'TIPO'
    if "TIPO" in df.columns:
        df_renomeado["TIPO"] = df["TIPO"].replace({
            "Consignado": "TERCEIRO_CONSIGNADO",
            "Proprio": "PROPRIO"
        })
    else:
        print("A coluna 'TIPO' não foi encontrada no DataFrame.")

    # Listar colunas finais desejadas
    colunas_final = [
        "ID", "STATUS", "DATA DE ENTRADA", "DATA E HORA DE SAIDA", "MARCA",
        "MODELO", "COMPLEMENTO", "CHASSI", "RENAVAM", "NUMERO MOTOR",
        "ANO FAB.", "ANO MOD", "COR", "COMBUSTIVEL", "PLACA", "TIPO",
        "VALOR COMPRA", "VALOR A VISTA", "VALOR DE VENDA",
        "NOME PROPRIETARIO ENTRADA", "CPF/CNPJ PROPRIETARIO ENTRADA",
        "KM", "PORTAS", "CAMBIO", "CNPJ REVENDA", "ESTADO_CONVERSACAO"
    ]

    # Criar DataFrame final com as colunas desejadas
    df_final = pd.DataFrame(columns=colunas_final)
    for coluna in colunas_final:
        if coluna in df_renomeado.columns:
            df_final[coluna] = df_renomeado[coluna]
        else:
            df_final[coluna] = None

    # Tratar valores null na coluna "ESTADO_CONVERSACAO"
    df_final["ESTADO_CONVERSACAO"] = df_final["ESTADO_CONVERSACAO"].fillna("USADO")
    # Tratar valores da coluna "ANO FAB." para serem iguais aos valores de "ANO MOD"
    df_final["ANO FAB."] = df_final["ANO MOD"]
    # Tratar valores da coluna "VALOR A VISTA" para ser igual ao "VALOR DE VENDA"
    df_final["VALOR A VISTA"] = df_final["VALOR DE VENDA"]
    # Dividir a coluna MODELO em MODELO e COMPLEMENTO
    df_final[["MODELO", "COMPLEMENTO"]] = df_final["MODELO"].str.split(n=1, expand=True)
    # Ajustar colunas para o formato numérico (removendo R$ e vírgulas)
    for col in ["VALOR COMPRA", "VALOR A VISTA", "VALOR DE VENDA"]:
        df_final[col] = df_final[col].str.replace(r"[^\d]", "", regex=True).astype(float)

    # Salvar o resultado na planilha destino
    df_final.to_excel(output_file, index=False)
    print(f"Arquivo salvo como: {output_file}")


# Chamada da função apenas com o caminho de origem
input_file = "planilha_origem.xlsx"
process_planilha(input_file)