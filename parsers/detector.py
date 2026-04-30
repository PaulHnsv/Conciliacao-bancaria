import os

_EXTENSOES = {
    ".ofx": "ofx",
    ".ofc": "ofx",
    ".csv": "csv",
    ".xls": "excel",
    ".xlsx": "excel",
    ".pdf": "pdf",
}

_ASSINATURAS_CONTEUDO = [
    (b"OFXHEADER", "ofx"),
    (b"<OFX>", "ofx"),
    (b"%PDF", "pdf"),
]

_EXTENSOES_TEXTO = {"", ".txt", ".dat", ".tsv"}


def detect_format(filepath: str) -> str:
    """Detecta o formato do arquivo de extrato bancario."""
    ext = os.path.splitext(filepath)[1].lower()

    if ext in _EXTENSOES:
        return _EXTENSOES[ext]

    if ext not in _EXTENSOES_TEXTO:
        raise ValueError(
            "Formato não suportado para o arquivo: " + os.path.basename(filepath)
            + "\nFormatos aceitos: .ofx, .ofc, .csv, .xls, .xlsx, .pdf"
        )

    try:
        with open(filepath, "rb") as f:
            head = f.read(512)
        for assinatura, formato in _ASSINATURAS_CONTEUDO:
            if assinatura in head:
                return formato
        head.decode("utf-8")
        return "csv"
    except UnicodeDecodeError:
        pass

    raise ValueError(
        "Formato não suportado para o arquivo: " + os.path.basename(filepath)
        + "\nFormatos aceitos: .ofx, .ofc, .csv, .xls, .xlsx, .pdf"
    )
