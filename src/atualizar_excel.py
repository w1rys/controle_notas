import os
import pandas as pd
from utils import log, log_erro
from openpyxl import load_workbook

# ===========================================================
# CLASSIFICAÇÃO DE PRODUTOS
# ===========================================================
def classificar_produto(nome_produto, emitente):
    nome = nome_produto.upper().strip()
    emit = emitente.upper().strip()

    # Shimano Blue Cycle
    if emit == "BLUE CYCLE & FISHING DISTRIBUIDORA SA":
        return "Shimano - Blue Cycle"

    # Bicicletas
    prefixos_bike = ["BIC.", "BICI", "BICICLETA"]
    if any(nome.startswith(p) for p in prefixos_bike):
        return "Bicicletas"

    # Acessórios
    acessorios = [
        "BOLSA", "BOMBA", "BUZINA", "CADEIRINHA", "CALÇA",
        "CAMISA", "CAPACETE", "GARRAFA", "CESTA", "LUVA",
        "MANOPLA", "PARALAMA", "RODA LATERAL", "SELIM"
    ]
    if any(a in nome for a in acessorios):
        return "Acessórios"

    return "Peças"


# ===========================================================
# GERAR ABAS DE CATEGORIAS
# ===========================================================
def atualizar_abas_categorias(df_produtos, nome_excel="produtos.xlsx"):
    categorias = ["Bicicletas", "Peças", "Acessórios", "Shimano - Blue Cycle"]

    try:
        load_workbook(nome_excel)
    except:
        return

    for categoria in categorias:
        df_cat = df_produtos[df_produtos["categoria"] == categoria]

        if df_cat.empty:
            continue

        df_final = df_cat[[
            "codigo",
            "nome_produto",
            "ultima_compra",
            "ultimo_preco",
            "penultimo_preco"
        ]]

        with pd.ExcelWriter(nome_excel, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
            df_final.to_excel(writer, sheet_name=categoria, index=False)

        book = load_workbook(nome_excel)
        ws = book[categoria]

        for row in range(2, ws.max_row + 1):
            ws.cell(row=row, column=3).number_format = "DD/MM/YYYY"

        book.save(nome_excel)

    log("[INFO] Abas de categorias atualizadas.")


# ===========================================================
# CARREGAR ARQUIVO EXCEL
# ===========================================================
def carregar_excel(nome_excel):
    colunas_base = [
        "codigo",
        "nome_produto",
        "quantidade_total_comprada",
        "ultimo_preco",
        "penultimo_preco",
        "ultima_compra",
        "chave_nota",
        "categoria"
    ]

    if not os.path.exists(nome_excel):
        df = pd.DataFrame(columns=colunas_base)
        df.to_excel(nome_excel, sheet_name="Compras", index=False)
        return df

    df = pd.read_excel(nome_excel, sheet_name="Compras", parse_dates=["ultima_compra"])

    for col in colunas_base:
        if col not in df.columns:
            df[col] = None

    return df


# ===========================================================
# SALVAR ABA COMPRAS
# ===========================================================
def salvar_excel(df_produtos, nome_excel):
    try:
        df_produtos["ultima_compra"] = pd.to_datetime(df_produtos["ultima_compra"], errors="coerce")

        with pd.ExcelWriter(
            nome_excel, engine="openpyxl", mode="a", if_sheet_exists="replace"
        ) as writer:
            df_produtos.to_excel(writer, sheet_name="Compras", index=False)

    except Exception as e:
        log_erro(f"Erro ao salvar Excel: {e}")


# ===========================================================
# GERAR ABA "PRODUTOS"
# ===========================================================
def atualizar_aba_produtos(df_produtos, nome_excel="produtos.xlsx"):
    from openpyxl import load_workbook

    try:
        # Normaliza datas, removendo timezone e hora
        df_produtos["ultima_compra"] = (
            pd.to_datetime(df_produtos["ultima_compra"], errors="coerce")
            .dt.tz_localize(None)
            .dt.normalize()
        )

        # Preenche N/A
        df_produtos = df_produtos.fillna("")

        # Base padronizada
        df_base = df_produtos[[
            "codigo",
            "nome_produto",
            "ultima_compra",
            "ultimo_preco",
            "penultimo_preco"
        ]]

        # Ordena alfabeticamente
        df_base = df_base.sort_values(by="nome_produto")

        # ---- CORREÇÃO DO BUG ----
        # Carrega Excel para obter preco_venda (se existir)
        try:
            book = load_workbook(nome_excel)
            if "Produtos" in book.sheetnames:
                antigo = pd.read_excel(nome_excel, sheet_name="Produtos")

                # Se preco_venda existe, fazemos merge sem perder produtos
                if "preco_venda" in antigo.columns:
                    df_base = df_base.merge(
                        antigo[["codigo", "preco_venda"]],
                        on="codigo",
                        how="left"  # <-- mantém TODOS os produtos do novo df
                    )
                else:
                    df_base["preco_venda"] = ""

            else:
                df_base["preco_venda"] = ""

        except Exception:
            df_base["preco_venda"] = ""

        # Remove duplicados caso existam
        df_base = df_base.drop_duplicates(subset=["codigo"])

        # --- SALVA A ABA PRODUTOS ---
        with pd.ExcelWriter(
            nome_excel,
            engine="openpyxl",
            mode="a",
            if_sheet_exists="replace"
        ) as writer:
            df_base.to_excel(writer, sheet_name="Produtos", index=False)

        # --- FORMATA A COLUNA DE DATA ---
        book = load_workbook(nome_excel)
        ws = book["Produtos"]

        # Coluna "ultima_compra" é a 3ª coluna
        for row in range(2, ws.max_row + 1):
            ws.cell(row=row, column=3).number_format = "DD/MM/YYYY"

        book.save(nome_excel)

    except Exception as e:
        log_erro(f"Erro ao atualizar aba Produtos: {e}")

# ===========================================================
# FUNÇÃO PRINCIPAL
# ===========================================================
def atualizar_excel_compras(novos_produtos, chave_nota, nome_excel="produtos.xlsx"):

    df = carregar_excel(nome_excel)

    if chave_nota in df["chave_nota"].astype(str).values:
        log(f"[INFO] Nota {chave_nota} já registrada.")
        return

    for item in novos_produtos:

        codigo = item["codigo"]
        qtd_nova = item["quantidade"]
        preco_novo = item["preco_unitario"]
        data_nova = pd.to_datetime(item["data_compra"]).tz_localize(None)
        emitente = item["emitente"]

        categoria = classificar_produto(item["nome_produto"], emitente)

        existente = df[df["codigo"] == codigo]

        if not existente.empty:
            idx = existente.index[0]

            df.loc[idx, "quantidade_total_comprada"] += qtd_nova

            data_atual = df.loc[idx, "ultima_compra"]
            data_atual = pd.to_datetime(data_atual) if not pd.isna(data_atual) else None

            if data_atual is None or data_nova > data_atual:
                df.loc[idx, "penultimo_preco"] = df.loc[idx, "ultimo_preco"]
                df.loc[idx, "ultimo_preco"] = preco_novo
                df.loc[idx, "ultima_compra"] = data_nova

        else:
            df.loc[len(df)] = {
                "codigo": codigo,
                "nome_produto": item["nome_produto"],
                "quantidade_total_comprada": qtd_nova,
                "ultimo_preco": preco_novo,
                "penultimo_preco": preco_novo,
                "ultima_compra": data_nova,
                "chave_nota": chave_nota,
                "categoria": categoria
            }

    df = df.sort_values(by="nome_produto")

    salvar_excel(df, nome_excel)
    atualizar_aba_produtos(df, nome_excel)
    atualizar_abas_categorias(df, nome_excel)
