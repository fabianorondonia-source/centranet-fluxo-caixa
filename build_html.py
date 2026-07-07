#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Embute dados.json em index_template.html antes da criptografia StatiCrypt.
Gera _build/index.html (NÃO criptografado — só existe temporariamente no
runner do GitHub Actions, nunca é commitado).
"""
import json, os, sys

TEMPLATE_PATH = "index_template.html"
DATA_PATH = "dados.json"
OUT_DIR = "_build"
OUT_PATH = os.path.join(OUT_DIR, "index.html")

def main():
    if not os.path.exists(DATA_PATH):
        print(f"ERRO: {DATA_PATH} não encontrado. Rode atualizar_dados.py primeiro.", file=sys.stderr)
        sys.exit(1)

    with open(TEMPLATE_PATH, "r", encoding="utf-8") as f:
        html = f.read()
    with open(DATA_PATH, "r", encoding="utf-8") as f:
        dados = json.load(f)

    marker = "<!-- DADOS_PLACEHOLDER -->"
    if marker not in html:
        print("ERRO: marcador DADOS_PLACEHOLDER não encontrado no template.", file=sys.stderr)
        sys.exit(1)

    injecao = "window.DADOS_EMBUTIDOS = " + json.dumps(dados, ensure_ascii=False, default=str) + ";"
    html = html.replace(marker, injecao)

    os.makedirs(OUT_DIR, exist_ok=True)
    with open(OUT_PATH, "w", encoding="utf-8") as f:
        f.write(html)

    print(f"OK — dados embutidos em {OUT_PATH}")

if __name__ == "__main__":
    main()
