import os
import sys
import subprocess
from datetime import datetime
from typing import Optional


def _selecionar_modo() -> Optional[str]:
    """Janela de escolha entre modo Simples e Detalhado."""
    try:
        import tkinter as tk
        from tkinter import ttk

        resultado = {"modo": None}

        root = tk.Tk()
        root.title("Conciliacao Bancaria")
        root.resizable(False, False)
        root.attributes("-topmost", True)

        # Centralizar na tela
        root.update_idletasks()
        w, h = 420, 260
        x = (root.winfo_screenwidth()  - w) // 2
        y = (root.winfo_screenheight() - h) // 2
        root.geometry(f"{w}x{h}+{x}+{y}")
        root.configure(bg="#1F3864")

        tk.Label(root, text="Conciliacao Bancaria",
                 font=("Segoe UI", 14, "bold"), fg="white", bg="#1F3864").pack(pady=(18, 4))
        tk.Label(root, text="Escolha o tipo de relatorio que deseja gerar:",
                 font=("Segoe UI", 10), fg="#BDD7EE", bg="#1F3864").pack(pady=(0, 16))

        frame = tk.Frame(root, bg="#1F3864")
        frame.pack(padx=20, fill="x")

        def escolher(modo):
            resultado["modo"] = modo
            root.destroy()

        btn_simples = tk.Button(
            frame, text="  Simples\n(Resumo executivo)", font=("Segoe UI", 10),
            bg="#2E75B6", fg="white", relief="flat", cursor="hand2",
            activebackground="#1F3864", activeforeground="white",
            width=16, height=3,
            command=lambda: escolher("simples")
        )
        btn_simples.pack(side="left", padx=(0, 10))

        btn_det = tk.Button(
            frame, text="  Detalhado\n(Extrato + Analise)", font=("Segoe UI", 10),
            bg="#1E7145", fg="white", relief="flat", cursor="hand2",
            activebackground="#1F3864", activeforeground="white",
            width=16, height=3,
            command=lambda: escolher("detalhado")
        )
        btn_det.pack(side="left")

        tk.Label(root, text="Simples: visao rapida  |  Detalhado: extrato completo + alertas de categoria",
                 font=("Segoe UI", 8), fg="#8EAADC", bg="#1F3864").pack(pady=(14, 0))

        root.protocol("WM_DELETE_WINDOW", root.destroy)
        root.mainloop()
        return resultado["modo"]
    except Exception as e:
        print(f"Erro ao abrir janela de modo: {e}")
        return "detalhado"


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


def _gerar_nome_saida(input_path: str, modo: str) -> str:
    pasta = os.path.dirname(os.path.abspath(input_path))
    timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M")
    sufixo = "simples" if modo == "simples" else "detalhado"
    return os.path.join(pasta, f"Conciliacao_{sufixo}_{timestamp}.xlsx")


def processar_extrato(input_path: str, output_path: Optional[str] = None,
                      modo: str = "detalhado") -> str:
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
    print(f"Modo do relatorio: {modo.upper()}")

    if output_path is None:
        output_path = _gerar_nome_saida(input_path, modo)

    gerar_relatorio(transacoes, output_path, modo=modo)
    print(f"Relatorio gerado: {output_path}")
    return output_path


def main():
    modo = "detalhado"
    input_path = None

    # Argumentos de linha de comando: main.py [arquivo] [--simples|--detalhado]
    args = sys.argv[1:]
    for arg in args:
        if arg in ("--simples", "-s"):
            modo = "simples"
        elif arg in ("--detalhado", "-d"):
            modo = "detalhado"
        elif not arg.startswith("-"):
            input_path = arg

    # Sem argumentos: mostra janela de selecao de modo primeiro
    if not args:
        modo_selecionado = _selecionar_modo()
        if modo_selecionado is None:
            print("Operacao cancelada.")
            sys.exit(0)
        modo = modo_selecionado

    if input_path is None:
        input_path = _selecionar_arquivo()
        if not input_path:
            print("Nenhum arquivo selecionado. Encerrando.")
            sys.exit(0)

    try:
        output = processar_extrato(input_path, modo=modo)
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
