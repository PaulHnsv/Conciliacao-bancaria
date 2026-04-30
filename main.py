import os
import sys
import subprocess
from datetime import datetime
from typing import Optional


def _selecionar_arquivo() -> Optional[str]:
    """Abre dialogo de selecao de arquivo nativo via tkinter."""
    try:
        import tkinter as tk
        from tkinter import filedialog
        root = tk.Tk()
        root.withdraw()
        root.attributes("-topmost", True)
        path = filedialog.askopenfilename(
            title="Selecionar extrato bancario",
            filetypes=[
                ("Extratos bancarios", "*.ofx *.ofc *.csv *.xlsx *.xls *.pdf"),
                ("OFX / OFC", "*.ofx *.ofc"),
                ("CSV", "*.csv"),
                ("Excel", "*.xlsx *.xls"),
                ("PDF", "*.pdf"),
                ("Todos os arquivos", "*.*"),
            ]
        )
        root.destroy()
        return path if path else None
    except Exception as e:
        print(f"Erro ao abrir seletor de arquivo: {e}")
        return None


def _gerar_nome_saida(input_path: str) -> str:
    pasta = os.path.dirname(os.path.abspath(input_path))
    timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M")
    return os.path.join(pasta, f"Conciliacao_{timestamp}.xlsx")


def processar_extrato(input_path: str, output_path: Optional[str] = None) -> str:
    """Processa arquivo de extrato e gera relatorio Excel."""
    from parsers.detector import detect_format
    from parsers.ofx_parser import parse_ofx
    from parsers.csv_parser import parse_csv
    from parsers.excel_parser import parse_excel
    from parsers.pdf_parser import parse_pdf
    from report.excel_report import gerar_relatorio

    formato = detect_format(input_path)
    _parsers = {
        "ofx":   parse_ofx,
        "csv":   parse_csv,
        "excel": parse_excel,
        "pdf":   parse_pdf,
    }
    print(f"Formato detectado: {formato.upper()}")
    transacoes = _parsers[formato](input_path)
    print(f"Transacoes encontradas: {len(transacoes)}")

    if output_path is None:
        output_path = _gerar_nome_saida(input_path)

    gerar_relatorio(transacoes, output_path)
    print(f"Relatorio gerado: {output_path}")
    return output_path


def main():
    if len(sys.argv) > 1:
        input_path = sys.argv[1]
    else:
        input_path = _selecionar_arquivo()
        if not input_path:
            print("Nenhum arquivo selecionado. Encerrando.")
            sys.exit(0)
    try:
        output = processar_extrato(input_path)
        if sys.platform == "win32":
            os.startfile(output)
        else:
            subprocess.run(["xdg-open", output])
    except Exception as e:
        print(f"\nErro ao processar o extrato: {e}")
        input("\nPressione ENTER para fechar...")
        sys.exit(1)


if __name__ == "__main__":
    main()
