from datetime import datetime


def log(msg):
    """Exibe mensagens padronizadas no terminal."""
    print(f"[INFO] {msg}")


def log_erro(msg):
    """Exibe mensagens de erro padronizadas no terminal."""
    print(f"[ERRO] {msg}")


def extrair_data_nota(data_xml):
    """
    Recebe dhEmi ou dEmi e converte para datetime (sempre timezone-naive).
    Suporta:
    - Formato moderno: 2024-01-15T12:22:00-03:00
    - Formato antigo: 2024-01-15
    """
    if not data_xml:
        return None

    try:
        # Formato ISO moderno
        if "T" in data_xml:
            dt = datetime.fromisoformat(data_xml.replace("Z", ""))
            # remover timezone se tiver
            if getattr(dt, "tzinfo", None) is not None:
                try:
                    # se for pandas Timestamp compatível, usar astimezone antes de dropar tzinfo
                    dt = dt.astimezone(tz=None)
                except Exception:
                    pass
                dt = dt.replace(tzinfo=None)
            return dt
        # Formato simples AAAA-MM-DD
        return datetime.strptime(data_xml, "%Y-%m-%d")
    except Exception:
        log_erro(f"Falha ao converter data: {data_xml}")
        return None



def extrair_chave_nota(inf):
    """
    Extrai a chave da NF-e a partir de:
    <infNFe Id="NFe35123456789012345678901234567890123456789012">
    """
    try:
        id_nota = inf.get("@Id")
        if id_nota and id_nota.startswith("NFe"):
            return id_nota.replace("NFe", "")
    except:
        pass
    return None


def validar_nota_nfe(xml_dict):
    """
    Verifica se o XML contém a estrutura mínima de NF-e.
    Retorna True/False.
    """
    try:
        if "nfeProc" in xml_dict:
            if "NFe" in xml_dict["nfeProc"]:
                return True

        if "NFe" in xml_dict:
            return True

        return False
    except:
        return False