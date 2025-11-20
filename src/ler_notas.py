import os
import xmltodict
import pandas as pd
from datetime import datetime
from pandas.api.types import DatetimeTZDtype

# ----------------------------------------------------------
# Extrair data da nota
# ----------------------------------------------------------
def extrair_data_nota(data_xml):
    if not data_xml:
        return None
    try:
        data = data_xml.replace("Z", "")
        dt = datetime.fromisoformat(data)
        # Remover timezone se existir
        if dt.tzinfo is not None:
            dt = dt.replace(tzinfo=None)
        return dt
    except:
        try:
            return datetime.strptime(data_xml, "%Y-%m-%d")
        except:
            return None

# ----------------------------------------------------------
# Extrair chave da NF-e
# ----------------------------------------------------------
def extrair_chave_nota(inf):
    try:
        id_nota = inf.get("@Id")  # Ex: "NFe35123456789..."
        if id_nota and id_nota.startswith("NFe"):
            return id_nota.replace("NFe", "")
    except:
        pass
    return None


# ----------------------------------------------------------
# Ler XML da nota e extrair produtos
# ----------------------------------------------------------
def ler_xml(path_xml):
    with open(path_xml, encoding="utf-8") as f:
        xml = f.read()
    data = xmltodict.parse(xml)

    try:
        inf = data["nfeProc"]["NFe"]["infNFe"]
    except:
        inf = data["NFe"]["infNFe"]

    # Extrair chave da nota
    chave_acesso = extrair_chave_nota(inf)

    # Extrair data
    data_emissao = extrair_data_nota(
        inf.get("ide", {}).get("dhEmi") or
        inf.get("ide", {}).get("dEmi")
    )

    # Itens
    itens = inf["det"]
    if isinstance(itens, dict):
        itens = [itens]

    produtos_extraidos = []
    for item in itens:
        prod = item["prod"]
        produtos_extraidos.append({
            "codigo": prod.get("cProd", ""),
            "nome_produto": prod.get("xProd", ""),
            "quantidade": float(prod.get("qCom", 0)),
            "preco_unitario": float(prod.get("vUnCom", 0)),
            "data_compra": data_emissao,
            "chave_nota": chave_acesso
        })

    return produtos_extraidos, chave_acesso


# ----------------------------------------------------------
# Atualizar Excel (controle de compras, sem estoque)
# ----------------------------------------------------------
def atualizar_excel_compras(novos_produtos, chave_nota, nome_excel="produtos.xlsx"):

    # Carregar Excel
    if os.path.exists(nome_excel):
        df_produtos = pd.read_excel(nome_excel, parse_dates=["ultima_compra"])
    else:
        df_produtos = pd.DataFrame(columns=[
            "codigo", "nome_produto",
            "quantidade_total_comprada",
            "preco_medio",
            "ultimo_preco",
            "ultima_compra",
            "chave_nota"
        ])

    # Verificar duplicidade pela chave da nota
    if chave_nota in df_produtos["chave_nota"].values:
        print("Nota já registrada. Ignorando:", chave_nota)
        return

    for item in novos_produtos:
        codigo = item["codigo"]
        quantidade_nova = item["quantidade"]
        preco_novo = item["preco_unitario"]
        data_nova = item["data_compra"]

        if codigo in df_produtos["codigo"].values:
            indice = df_produtos[df_produtos["codigo"] == codigo].index[0]

            quantidade_antiga = df_produtos.loc[indice, "quantidade_total_comprada"]
            preco_medio_antigo = df_produtos.loc[indice, "preco_medio"]

            novo_preco_medio = (
                (quantidade_antiga * preco_medio_antigo) +
                (quantidade_nova * preco_novo)
            ) / (quantidade_antiga + quantidade_nova)

            df_produtos.loc[indice, "quantidade_total_comprada"] = quantidade_antiga + quantidade_nova
            df_produtos.loc[indice, "preco_medio"] = novo_preco_medio

            if pd.isna(df_produtos.loc[indice, "ultima_compra"]) or data_nova > df_produtos.loc[indice, "ultima_compra"]:
                df_produtos.loc[indice, "ultimo_preco"] = preco_novo
                df_produtos.loc[indice, "ultima_compra"] = data_nova

        else:
            df_produtos.loc[len(df_produtos)] = {
                "codigo": item["codigo"],
                "nome_produto": item["nome_produto"],
                "quantidade_total_comprada": quantidade_nova,
                "preco_medio": preco_novo,
                "ultimo_preco": preco_novo,
                "ultima_compra": data_nova,
                "chave_nota": chave_nota
            }

    # 🔥 Remover timezone das datas ANTES de salvar
    for col in df_produtos.columns:
        if isinstance(df_produtos[col].dtype, DatetimeTZDtype):
            df_produtos[col] = df_produtos[col].dt.tz_localize(None)

    df_produtos.to_excel(nome_excel, index=False)
    print("Excel atualizado:", nome_excel)

# ----------------------------------------------------------
# Ler todas as notas da pasta
# ----------------------------------------------------------
def ler_todas_notas(pasta):
    for arquivo in os.listdir(pasta):
        if arquivo.endswith(".xml"):
            caminho = os.path.join(pasta, arquivo)
            print("Lendo:", arquivo)

            produtos, chave = ler_xml(caminho)
            if chave:
                atualizar_excel_compras(produtos, chave)
            else:
                print("Chave da nota não encontrada. Nota ignorada.")


if __name__ == "__main__":
    pasta_notas = "notas"
    ler_todas_notas(pasta_notas)
