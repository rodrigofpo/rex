"""
Interface de Linha de Comando (CLI) unificada para o REX.
"""

import sys
import argparse
import time
from pathlib import Path
from rex import __version__
from rex.core.engine import EDSExtractorEngine


def run_extract(args):
    docx_path = Path(args.docx)
    if not docx_path.is_file():
        print(f"Erro: O arquivo '{docx_path}' não foi encontrado.")
        sys.exit(1)

    output_dir = Path(args.output) if args.output else docx_path.parent

    print("=" * 65)
    print(f"  REX v{__version__} — MEV-EDS REPORT EXTRACTOR")
    print("=" * 65)
    print(f"Arquivo de entrada: {docx_path.name} ({docx_path.stat().st_size / (1024*1024):.1f} MB)")
    print(f"Diretório de saída: {output_dir.resolve()}")
    print("-" * 65)

    t0 = time.time()
    engine = EDSExtractorEngine()

    try:
        df_dados, df_amostras = engine.process(docx_path, output_dir=output_dir)
        elapsed = time.time() - t0

        print("\n" + "=" * 65)
        print("  RESUMO DA EXTRAÇÃO")
        print("=" * 65)
        print(f"Tempo total gasto:        {elapsed:.2f} segundos")
        print(f"Total de Amostras Lidas:   {len(df_amostras)}")
        print(f"Total de Linhas Químicas:  {len(df_dados)}")

        if not df_dados.empty:
            elementos = sorted(df_dados['Elemento'].unique().tolist())
            pontos = df_dados['Ponto'].nunique()
            print(f"Elementos Detectados:     {', '.join(elementos)}")
            print(f"Total de Pontos Únicos:   {pontos}")
            print("\nPrévia dos primeiros registros:")
            print(df_dados[['Amostra', 'Serie', 'Ponto', 'Elemento', 'PercentualPeso', 'PercentualAtomico']].head(10).to_string(index=False))

        print("\n[OK] Extração finalizada com sucesso!")

    except Exception as e:
        print(f"\n[ERRO CRÍTICO] Falha durante a extração: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)


def run_web(args):
    from rex.ui.web import launch_web_app
    launch_web_app(port=args.port, open_browser=not args.no_browser)


def run_desktop(args):
    try:
        from rex.ui.desktop import launch_desktop_app
        launch_desktop_app()
    except ImportError as e:
        if "tkinter" in str(e).lower():
            print("\n[ERRO] O módulo 'tkinter' não está instalado no seu sistema.")
            print("  • No Fedora Linux: execute 'sudo dnf install python3-tkinter'")
            print("  • No Ubuntu/Debian: execute 'sudo apt install python3-tk'")
            print("  • Dica: Você também pode usar a interface Web com 'rex web' (não precisa de tkinter)!\n")
            sys.exit(1)
        raise e


def main():
    parser = argparse.ArgumentParser(
        prog="rex",
        description="REX — MEV-EDS Report Extractor: Extração e estruturação determinística de laudos de microscopia."
    )
    parser.add_argument("-v", "--version", action="version", version=f"rex {__version__}")

    subparsers = parser.add_subparsers(dest="command", help="Comando a ser executado")

    # Comando 'extract'
    parser_extract = subparsers.add_parser("extract", help="Extrair dados de um laudo .docx")
    parser_extract.add_argument("docx", help="Caminho para o arquivo de laudo .docx")
    parser_extract.add_argument("-o", "--output", help="Diretório de saída para o Excel e CSV (opcional)")
    parser_extract.set_defaults(func=run_extract)

    # Comando 'web'
    parser_web = subparsers.add_parser("web", help="Iniciar a interface Web local no navegador")
    parser_web.add_argument("-p", "--port", type=int, default=8085, help="Porta local do servidor (padrão: 8085)")
    parser_web.add_argument("--no-browser", action="store_true", help="Não abrir o navegador automaticamente")
    parser_web.set_defaults(func=run_web)

    # Comando 'gui' / 'desktop'
    parser_gui = subparsers.add_parser("gui", help="Iniciar a interface gráfica Desktop (CustomTkinter)")
    parser_gui.set_defaults(func=run_desktop)

    args = parser.parse_args()

    if not args.command:
        parser.print_help()
        sys.exit(0)

    args.func(args)


if __name__ == "__main__":
    main()
