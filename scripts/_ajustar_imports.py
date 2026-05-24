"""
_ajustar_imports.py — Script auxiliar pra ajustar imports após reorganização.

Roda pela primeira vez com `--dry-run` pra ver o que mudaria, depois sem o
flag pra aplicar de fato.

Uso:
    python scripts/_ajustar_imports.py --dry-run   # só mostra
    python scripts/_ajustar_imports.py             # aplica
"""

from __future__ import annotations

import argparse
import re
import sys
from pathlib import Path

# ── Mapa de substituições (regex pattern → replacement) ───────────────
# A ordem importa: padrões mais específicos vêm ANTES dos genéricos.
# Usa lookbehind/lookahead pra evitar substituir dentro de strings/comentários
# quando possível.

SUBSTITUICOES: list[tuple[re.Pattern, str, str]] = [
    # ── from X import ... (mais específico, tem que vir antes) ──
    (re.compile(r"^(\s*)from extracao_api import "),
     r"\1from extratores.Condominios.extracao_api import ",
     "from extracao_api import"),

    (re.compile(r"^(\s*)from fluxo_pw_new import "),
     r"\1from extratores.Condominios.fluxo_pw_new import ",
     "from fluxo_pw_new import"),

    (re.compile(r"^(\s*)from webapp\."),
     r"\1from ui.web.webapp.",
     "from webapp."),

    (re.compile(r"^(\s*)from janela import "),
     r"\1from ui.desktop.janela import ",
     "from janela import"),

    # ── import X as Y ──
    (re.compile(r"^(\s*)import auxiliares as (\w+)"),
     r"\1from auxiliares import utils as \2",
     "import auxiliares as ..."),

    (re.compile(r"^(\s*)import sensiveis as (\w+)"),
     r"\1from auxiliares import sensiveis as \2",
     "import sensiveis as ..."),

    (re.compile(r"^(\s*)import Biptu_hibrido as (\w+)"),
     r"\1from extratores.Prefeitura import Biptu_hibrido as \2",
     "import Biptu_hibrido as ..."),

    (re.compile(r"^(\s*)import extracao_api as (\w+)"),
     r"\1from extratores.Condominios import extracao_api as \2",
     "import extracao_api as ..."),

    (re.compile(r"^(\s*)import taxabombeiros as (\w+)"),
     r"\1from extratores.Bombeiros._legado import taxabombeiros as \2",
     "import taxabombeiros as ..."),

    (re.compile(r"^(\s*)import messagebox as (\w+)"),
     r"\1from auxiliares import messagebox as \2",
     "import messagebox as ..."),

    (re.compile(r"^(\s*)import boletos as (\w+)"),
     r"\1from core import boletos as \2",
     "import boletos as ..."),

    # ── import X (sem alias) ──
    (re.compile(r"^(\s*)import boletos\b"),
     r"\1from core import boletos",
     "import boletos"),

    (re.compile(r"^(\s*)import web\b(?!app)"),  # evita pegar 'webapp' ou 'webdriver'
     r"\1from core import web",
     "import web"),

    (re.compile(r"^(\s*)import extracao\b(?!_api)"),  # não pega extracao_api
     r"\1from core import extracao",
     "import extracao"),

    (re.compile(r"^(\s*)import messagebox\b"),
     r"\1from auxiliares import messagebox",
     "import messagebox"),

    (re.compile(r"^(\s*)import aux_patches\b"),
     r"\1from auxiliares import aux_patches",
     "import aux_patches"),

    (re.compile(r"^(\s*)import condominios\b"),
     r"\1from extratores.Condominios import condominios",
     "import condominios"),

    (re.compile(r"^(\s*)import extracao_api\b"),
     r"\1from extratores.Condominios import extracao_api",
     "import extracao_api"),

    (re.compile(r"^(\s*)import fluxo_pw_new\b"),
     r"\1from extratores.Condominios import fluxo_pw_new",
     "import fluxo_pw_new"),

    (re.compile(r"^(\s*)import inventario_api\b"),
     r"\1from extratores.Condominios import inventario_api",
     "import inventario_api"),

    (re.compile(r"^(\s*)import Biptu\b(?!_)"),  # não pega Biptu_hibrido/Biptu_pw
     r"\1from extratores.Prefeitura import Biptu",
     "import Biptu"),

    (re.compile(r"^(\s*)import Biptu_hibrido\b"),
     r"\1from extratores.Prefeitura import Biptu_hibrido",
     "import Biptu_hibrido"),

    (re.compile(r"^(\s*)import taxabombeiros\b"),
     r"\1from extratores.Bombeiros._legado import taxabombeiros",
     "import taxabombeiros"),
]


# Pastas onde vou procurar arquivos .py
PASTAS_ALVO = ["extratores", "ui", "core", "auxiliares", "scripts"]

# Arquivos que NÃO devem ser modificados (caso especial: o próprio _ajustar_imports)
EXCLUIR = {"_ajustar_imports.py"}


def processar_arquivo(caminho: Path, dry_run: bool) -> list[tuple[int, str, str, str]]:
    """Aplica todas as substituições no arquivo. Retorna lista de (linha, antes, depois, regra)."""
    try:
        original = caminho.read_text(encoding="utf-8")
    except UnicodeDecodeError:
        try:
            original = caminho.read_text(encoding="latin-1")
        except Exception:
            print(f"  [PULAR] {caminho}: falha de encoding")
            return []

    mudancas = []
    novo_conteudo_linhas = []
    for num_linha, linha in enumerate(original.splitlines(keepends=True), start=1):
        nova = linha
        for padrao, replacement, descricao in SUBSTITUICOES:
            resultado = padrao.sub(replacement, nova)
            if resultado != nova:
                mudancas.append((num_linha, nova.rstrip("\n"), resultado.rstrip("\n"), descricao))
                nova = resultado
        novo_conteudo_linhas.append(nova)

    if mudancas and not dry_run:
        caminho.write_text("".join(novo_conteudo_linhas), encoding="utf-8")

    return mudancas


def main():
    parser = argparse.ArgumentParser(description="Ajusta imports após reorganização")
    parser.add_argument("--dry-run", action="store_true", help="Só mostra o que mudaria")
    args = parser.parse_args()

    raiz = Path(__file__).resolve().parent.parent
    total_mudancas = 0
    arquivos_afetados = 0

    for pasta_nome in PASTAS_ALVO:
        pasta = raiz / pasta_nome
        if not pasta.exists():
            continue
        for arquivo in sorted(pasta.rglob("*.py")):
            if arquivo.name in EXCLUIR:
                continue
            mudancas = processar_arquivo(arquivo, args.dry_run)
            if mudancas:
                arquivos_afetados += 1
                total_mudancas += len(mudancas)
                rel = arquivo.relative_to(raiz)
                print(f"\n📄 {rel}")
                for num_linha, antes, depois, regra in mudancas:
                    print(f"  L{num_linha:>4} [{regra}]")
                    print(f"       - {antes}")
                    print(f"       + {depois}")

    modo = "DRY-RUN (nenhuma alteração feita)" if args.dry_run else "APLICADO"
    print(f"\n{'═' * 60}")
    print(f"  Modo: {modo}")
    print(f"  Arquivos afetados: {arquivos_afetados}")
    print(f"  Total de substituições: {total_mudancas}")
    print(f"{'═' * 60}")


if __name__ == "__main__":
    main()
