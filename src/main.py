import os
import sys
import argparse
from src.runner import run_excel_tests

root = os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))
if root not in sys.path:
    sys.path.insert(0, root)

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--file", required=True)
    parser.add_argument("--sheet", default="ALL")
    parser.add_argument("--no-copy", action="store_true")
    args = parser.parse_args()

    out = run_excel_tests(args.file, args.sheet, make_copy=not args.no_copy)
    print(f"\n✅ Finalizado. Planilha processada em: {out}")

if __name__ == "__main__":
    main()