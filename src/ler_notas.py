import os
import re
import xmltodict
from utils import (
    log,
    log_erro,
    extrair_data_nota,
    extrair_chave_nota,
    validar_nota_nfe
)
from atualizar_excel import atualizar_excel_compras


def limpar_nome_emitente(nome):
    # remove caracteres especiais e deixa só letras/numeros/espaço
    nome_limpo = re.sub(r"[^A-Za-z0-9 ]+", "", nome).strip().upper()

    # separa por espaço
    partes = nome_limpo.split()

    # pega somente as DUAS primeiras palavras
    if len(partes) >= 2:
        nome_duas = f"{partes[0]}_{partes[1]}"
    else:
        nome_duas = partes[0]

    return nome_duas


def ler_xml(path_xml):

    try:
        with open(path_xml, encoding="utf-8") as f:
            xml = f.read()
    except Exception as e:
        log_erro(f"Erro ao abrir arquivo {path_xml}: {e}")
        return [], None

    try:
        data = xmltodict.parse(xml)
    except Exception as e:
        log_erro(f"Erro ao interpretar XML {path_xml}: {e}")
        return [], None

    # validar xml
    if not validar_nota_nfe(data):
        log_erro(f"Arquivo ignorado (não é NF-e): {path_xml}")
        return [], None

    inf = data["nfeProc"]["NFe"]["infNFe"]
    chave = extrair_chave_nota(inf)

    if not chave:
        log_erro(f"Chave não encontrada: {path_xml}")
        return [], None

    data_emissao = extrair_data_nota(
        inf["ide"].get("dhEmi") or inf["ide"].get("dEmi")
    )

    emitente = inf["emit"]["xNome"].strip()
    emitente_norm = limpar_nome_emitente(emitente)

    # itens
    itens = inf["det"]
    if isinstance(itens, dict):
        itens = [itens]

    produtos = []

    for item in itens:
        prod = item["prod"]

        codigo_prod = str(prod.get("cProd", "")).strip()

        # código final: DUAS PRIMEIRAS PALAVRAS DO FORNECEDOR
        codigo = f"{emitente_norm}-{codigo_prod}"

        produtos.append({
            "codigo": codigo,
            "nome_produto": prod.get("xProd", "").strip(),
            "quantidade": float(prod.get("qCom", 0)),
            "preco_unitario": float(prod.get("vUnCom", 0)),
            "data_compra": data_emissao,
            "chave_nota": chave,
            "emitente": emitente
        })

    log(f"Nota processada: {os.path.basename(path_xml)}")
    return produtos, chave


def ler_todas_notas(pasta="notas"):
    for arquivo in os.listdir(pasta):
        if arquivo.endswith(".xml"):
            caminho = os.path.join(pasta, arquivo)
            produtos, chave = ler_xml(caminho)

            if chave:
                atualizar_excel_compras(produtos, chave)
            else:
                log_erro(f"Nota ignorada: {arquivo}")