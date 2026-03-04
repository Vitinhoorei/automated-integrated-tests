import os
import sys
sys.path.insert(0, os.path.dirname(__file__))

import argparse
from runner import run_excel_tests

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--file", required=True)
    parser.add_argument(
        "--sheet",
        default="ALL",
        help='Nome da aba, lista separada por vírgula, ou "ALL" para todas.',
    )
    parser.add_argument("--no-copy", action="store_true")
    args = parser.parse_args()

    out = run_excel_tests(args.file, args.sheet, make_copy=not args.no_copy)
    print(f"\n✅ Finalizado. Planilha processada em: {out}")

if __name__ == "__main__":
    main()
