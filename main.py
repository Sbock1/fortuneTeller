import argparse
import APIandDatabaseLib as gil

def main():
    parser = argparse.ArgumentParser(description="FullStackTradingApp")
    parser.add_argument('--mode', choices=['cli', 'auto'], default='cli', help='Start mode')
    args = parser.parse_args()

    if args.mode == 'cli':
        # Beispiel: CLI-Modus (hier könntest du ein CLI-Menü aufrufen)
        print("Starte CLI-Modus...")
        # z.B. symbol_list = gil.load_stock_symbol_list()
        # print(symbol_list)
        # Hier könntest du weitere CLI-Interaktionen einbauen
    elif args.mode == 'auto':
        # Beispiel: Automatischer Ablauf (API-Abruf, DB-Speicherung, Analyse)
        print("Starte automatischen Ablauf...")
        symbol_list = gil.loadOneColumnRowDataAsList("Stock_symbols_list.xlsx", "Overview", "A")()
        print(f"Geladene Symbole: {symbol_list}")
        # Hier weitere automatische Schritte einbauen

if __name__ == "__main__":
    main()
