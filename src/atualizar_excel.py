import os
import pandas as pd
from utils import log, log_erro
from openpyxl import load_workbook


# ----------------------------------------------------------
# Carregar Excel
# ----------------------------------------------------------
def carregar_excel(nome_excel):
    """Carrega o Excel ou cria um novo DataFrame se não existir."""
    if os.path.exists(nome_excel):
        try:
            df_produtos = pd.read_excel(nome_excel, parse_dates=["ultima_compra"])
            return df_produtos
        except Exception as e:
            log_erro(f"Erro ao ler o arquivo Excel: {e}")
            return None
    else:
        colunas = [
            "codigo",
            "nome_produto",
            "quantidade_total_comprada",
            "preco_medio",
            "ultimo_preco",
            "ultima_compra",
            "chave_nota"
        ]
        return pd.DataFrame(columns=colunas)


# ----------------------------------------------------------
# Verificar duplicidade da nota
# ----------------------------------------------------------
def nota_ja_processada(df_produtos, chave_nota):
    """Retorna True se a chave da nota já estiver registrada."""
    if "chave_nota" not in df_produtos.columns:
        return False
    return chave_nota in df_produtos["chave_nota"].values


# ----------------------------------------------------------
# Salvar aba principal
# ----------------------------------------------------------
def salvar_excel(df_produtos, nome_excel):
    """Salva o DataFrame no arquivo Excel."""
    try:
        df_produtos.to_excel(nome_excel, index=False)
        log(f"Excel atualizado: {nome_excel}")
    except Exception as e:
        log_erro(f"Erro ao salvar o Excel: {e}")


# ----------------------------------------------------------
# Atualizar aba produtos_base (consulta)
# ----------------------------------------------------------
def atualizar_aba_produtos_base(df_produtos, nome_excel="produtos.xlsx"):
    """
    Cria ou atualiza a aba 'produtos_base' com:
    codigo, nome_produto, ultimo_preco, preco_medio, preco_venda (manual)
    """

    df_base = df_produtos[[
        "codigo",
        "nome_produto",
        "ultimo_preco",
        "preco_medio"
    ]].copy()

    # Ordena por nome
    df_base = df_base.sort_values(by="nome_produto", ascending=True)

    aba = "produtos_base"

    try:
        book = load_workbook(nome_excel)

        # Se a aba já existir, preservar preco_venda
        if aba in book.sheetnames:
            df_antigo = pd.read_excel(nome_excel, sheet_name=aba)

            # Se preco_venda já existir, manter
            if "preco_venda" in df_antigo.columns:
                df_base = df_base.merge(
                    df_antigo[["codigo", "preco_venda"]],
                    on="codigo",
                    how="left"
                )
            else:
                df_base["preco_venda"] = None
        else:
            df_base["preco_venda"] = None

        # Salvar aba
        with pd.ExcelWriter(nome_excel, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
            df_base.to_excel(writer, sheet_name=aba, index=False)

    except FileNotFoundError:
        # Se o arquivo não existir ainda
        df_base["preco_venda"] = None
        df_base.to_excel(nome_excel, sheet_name=aba, index=False)


# ----------------------------------------------------------
# Atualizar Excel principal (controle de compras)
# ----------------------------------------------------------
def atualizar_excel_compras(novos_produtos, chave_nota, nome_excel="produtos.xlsx"):
    """
    Atualiza o Excel com informações das notas:
    - quantidade total comprada
    - média ponderada
    - último preço
    - última compra
    - chave da nota
    """

    df_produtos = carregar_excel(nome_excel)
    if df_produtos is None:
        return

    # Evitar duplicidade
    if nota_ja_processada(df_produtos, chave_nota):
        log(f"Nota já registrada. Ignorando importação: {chave_nota}")
        return

    # Processar cada produto da nota
    for item in novos_produtos:
        codigo = item["codigo"]
        quantidade_nova = item["quantidade"]
        preco_novo = item["preco_unitario"]
        data_nova = item["data_compra"]

        produto_existente = df_produtos[df_produtos["codigo"] == codigo]

        if not produto_existente.empty:
            indice = produto_existente.index[0]

            quantidade_antiga = df_produtos.loc[indice, "quantidade_total_comprada"]
            preco_medio_antigo = df_produtos.loc[indice, "preco_medio"]

            quantidade_total = quantidade_antiga + quantidade_nova

            # média ponderada
            novo_preco_medio = (
                (quantidade_antiga * preco_medio_antigo) +
                (quantidade_nova * preco_novo)
            ) / quantidade_total

            # Atualizar
            df_produtos.loc[indice, "quantidade_total_comprada"] = quantidade_total
            df_produtos.loc[indice, "preco_medio"] = novo_preco_medio

            # Atualizar último preço/data
            if pd.isna(df_produtos.loc[indice, "ultima_compra"]) or data_nova > df_produtos.loc[indice, "ultima_compra"]:
                df_produtos.loc[indice, "ultimo_preco"] = preco_novo
                df_produtos.loc[indice, "ultima_compra"] = data_nova

        else:
            # Criar novo produto
            df_produtos.loc[len(df_produtos)] = {
                "codigo": codigo,
                "nome_produto": item["nome_produto"],
                "quantidade_total_comprada": quantidade_nova,
                "preco_medio": preco_novo,
                "ultimo_preco": preco_novo,
                "ultima_compra": data_nova,
                "chave_nota": chave_nota
            }

    # ------------------------------------------------------
    # ORDENAR A ABA PRINCIPAL POR NOME
    # ------------------------------------------------------
    df_produtos = df_produtos.sort_values(by="nome_produto", ascending=True)

    # Salvar aba principal
    salvar_excel(df_produtos, nome_excel)

    # Atualizar aba secundária
    atualizar_aba_produtos_base(df_produtos, nome_excel)
