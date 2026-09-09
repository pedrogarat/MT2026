# -*- coding: utf-8 -*-
"""
main.py
Punto de entrada principal para o Asistente do Estudio Básico de Seguridade e Saúde.
Executa a interface web por defecto, ou o modo consola con --cli.
"""

import sys
import os

CURRENT_DIR = os.path.dirname(os.path.abspath(__file__))
if CURRENT_DIR not in sys.path:
    sys.path.insert(0, CURRENT_DIR)

def main():
    if "--cli" in sys.argv:
        from cli import run_cli
        run_cli()
    elif "--test" in sys.argv:
        from ebss_processor import DEFAULT_DATA, generate_ebss
        print("Executando test de xeración...")
        test_data = dict(DEFAULT_DATA)
        test_data["tipo_obra"] = "PROBA AUTOMATIZADA DE XERACIÓN EBSS"
        test_data["nome_ficheiro_saida"] = "TEST_AUTOMATIZADO_EBSS.docx"
        res = generate_ebss(test_data)
        print("Resultado:", res)
    else:
        from app import start_server
        start_server()

if __name__ == "__main__":
    main()
