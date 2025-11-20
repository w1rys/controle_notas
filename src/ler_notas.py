import os
import xmltodict
from utils import (
    log,
    log_erro,
    extrair_data_nota,
    extrair_chave_nota,
    validar_nota_nfe
)
from atualizar_excel import atualizar_excel_compras


# ----------------------------------------------------------
# Ler uma nota XML e extrair produtos
# ----------------------------------------------------------
def ler_xml(path_xml):
    """
    Lê um XML de NF-e, extrai:
    - chave da nota
    - data da compra
    - itens
    E retorna: (lista_produtos, chave_nota)
    """

    # Ler arquivo XML
    try:
        with open(path_xml, encoding="utf-8") as f:
            xml = f.read()
    except Exception as e:
        log_erro(f"Erro ao abrir arquivo {path_xml}: {e}")
        return [], None

    # Interpretar XML
    try:
        data = xmltodict.parse(xml)
    except Exception as e:
        log_erro(f"Erro ao interpretar XML {path_xml}: {e}")
        return [], None

    # Validar estrutura
    if not validar_nota_nfe(data):
        log_erro(f"Arquivo ignorado (não é NF-e): {path_xml}")
        return [], None

    # Acessar bloco principal da NF-e
    try:
        inf = data["nfeProc"]["NFe"]["infNFe"]
    except:
        try:
            inf = data["NFe"]["infNFe"]
        except:
            log_erro(f"NF-e com estrutura inválida: {path_xml}")
            return [], None

    # ===============================
    # Extrair chave da nota
    # ===============================
    chave = extrair_chave_nota(inf)
    if not chave:
        log_erro(f"Chave da nota não encontrada: {path_xml}")
        return [], None

    # ===============================
    # Extrair data de emissão
    # ===============================
    data_emissao = extrair_data_nota(
        inf.get("ide", {}).get("dhEmi") or inf.get("ide", {}).get("dEmi")
    )

    # ===============================
    # Extrair lista de itens
    # ===============================
    itens = inf["det"]
    if isinstance(itens, dict):  # Se tiver apenas 1 item
        itens = [itens]

    produtos = []

    for item in itens:
        prod = item["prod"]

        produtos.append({
            "codigo": str(prod.get("cProd", "")).strip(),
            "nome_produto": prod.get("xProd", "").strip(),
            "quantidade": float(prod.get("qCom", 0)),
            "preco_unitario": float(prod.get("vUnCom", 0)),
            "data_compra": data_emissao,
            "chave_nota": chave
        })

    log(f"Nota processada: {os.path.basename(path_xml)}")

    return produtos, chave


# ----------------------------------------------------------
# Ler toda a pasta "notas" (modo manual, sem monitor)
# ----------------------------------------------------------
def ler_todas_notas(pasta="notas"):
    for arquivo in os.listdir(pasta):
        if arquivo.endswith(".xml"):
            caminho = os.path.join(pasta, arquivo)

            log(f"Lendo {arquivo}...")
            produtos, chave = ler_xml(caminho)

            if chave:
                atualizar_excel_compras(produtos, chave)
            else:
                log_erro(f"Nota ignorada: {arquivo}")


if __name__ == "__main__":
    ler_todas_notas("notas")
