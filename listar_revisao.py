"""
listar_revisao.py — Gera lista simples de Condomínios e Unidades (Scai + API)
                    por administradora, sem fuzzy match, pro Elias revisar
                    visualmente e decidir o que ajustar no cadastro.

Lê:
  • Inventario_Condominios_*.xlsx + Inventario_Condominios_x_Unidades_*.xlsx
    (gerados por inventario_api.py)
  • Tabela Inquilin_New no Scai_local.WMB

Produz: logs/Lista_Revisao_<ts>.xlsx com QUATRO abas separadas:
  • Condominios_Scai → Adm | Nome | CodCliente | LoginScai
  • Condominios_API  → Adm | Nome | Servidor | UsuarioAPI | ID_Condominio
  • Unidades_Scai    → Adm | Condominio | Unidade | CodCliente | LoginScai
  • Unidades_API     → Adm | Condominio | Unidade | Servidor | UsuarioAPI | ID_Unidade

Cada linha aparece uma única vez (deduplicado). Ordenado por Adm + Nome
para que pares "quase iguais" fiquem adjacentes em ordem alfabética dentro
de cada aba.

Uso:
  py listar_revisao.py             # todas as admins
  py listar_revisao.py CIPA BCF    # filtra por administradora
"""

from __future__ import annotations

import sys
import datetime
from pathlib import Path

import pandas as pd
from rich.console import Console

from dotenv import load_dotenv
load_dotenv()

import auxiliares as aux


console = Console(force_terminal=True, color_system="truecolor")

ROOT       = Path(__file__).parent
LOGS_DIR   = ROOT / "logs"
BANCO_PATH = r"C:\Users\efart\OneDrive\Área de Trabalho\Scai_local.WMB"


# ═══════════════════════════════════════════════════════════════════════════════
#  Leitura dos dados
# ═══════════════════════════════════════════════════════════════════════════════

def _excel_mais_recente(padrao: str) -> Path | None:
    cands = sorted(LOGS_DIR.glob(padrao), reverse=True)
    return cands[0] if cands else None


def _ler_inventario() -> tuple[pd.DataFrame, pd.DataFrame, str]:
    arq_c = _excel_mais_recente("Inventario_Condominios_[0-9]*.xlsx")
    arq_u = _excel_mais_recente("Inventario_Condominios_x_Unidades_*.xlsx")
    if not arq_c or not arq_u:
        raise FileNotFoundError(
            f"Não encontrei Inventario_*.xlsx em {LOGS_DIR}.\n"
            "Rode 'python inventario_api.py' antes."
        )
    console.print(f"  API condomínios: [dim]{arq_c.name}[/dim]")
    console.print(f"  API unidades:    [dim]{arq_u.name}[/dim]")
    df_c = pd.read_excel(arq_c, sheet_name="Condominios")
    df_u = pd.read_excel(arq_u, sheet_name="Condominios_x_Unidades")
    ts = arq_c.stem.replace("Inventario_Condominios_", "")
    return df_c, df_u, ts


def _ler_scai(admins_filtro: list | None) -> pd.DataFrame:
    # MS Access: alias com mesmo nome do campo dá referência circular.
    sql = """
        SELECT TRIM(NomeAdm)         AS Adm,
               TRIM(NomeCond)        AS Condo,
               TRIM(Unidade)         AS Unid,
               TRIM(LogAdm)          AS LoginScai,
               LEFT(TRIM(Codigo), 4) AS CodCli
          FROM Inquilin_New
         WHERE Situa NOT IN ('E','F','K','V')
               AND TRIM(NomeAdm)  <> ''
               AND TRIM(NomeCond) <> ''
    """
    banco = aux.Banco(BANCO_PATH)
    rows = banco.consultar(sql)
    df = pd.DataFrame(rows, columns=["Adm", "Condo", "Unid", "LogAdm", "CodCliente"])
    if admins_filtro:
        df = df[df["Adm"].isin(admins_filtro)]
    return df


# ═══════════════════════════════════════════════════════════════════════════════
#  Construção das listas
# ═══════════════════════════════════════════════════════════════════════════════

def _agg_unique(series) -> str:
    """Junta valores únicos não-vazios em uma string separada por vírgula."""
    vals = sorted({str(x).strip() for x in series if str(x).strip()})
    return ", ".join(vals)


def condos_scai(df_scai: pd.DataFrame) -> pd.DataFrame:
    """Lista de condomínios cadastrados — um por (Adm, Nome)."""
    if df_scai.empty:
        return pd.DataFrame(columns=["Adm", "Nome", "CodCliente", "LoginScai"])
    df = (df_scai.groupby(["Adm", "Condo"], dropna=False)
                 .agg(CodCliente=("CodCliente", _agg_unique),
                      LoginScai=("LogAdm",      _agg_unique))
                 .reset_index()
                 .rename(columns={"Condo": "Nome"}))
    return df[["Adm", "Nome", "CodCliente", "LoginScai"]] \
              .sort_values(["Adm", "Nome"]).reset_index(drop=True)


def condos_api(df_api: pd.DataFrame) -> pd.DataFrame:
    """Lista de condomínios da API — um por (Adm, Nome)."""
    if df_api.empty:
        return pd.DataFrame(columns=["Adm", "Nome", "Servidor", "UsuarioAPI", "ID_Condominio"])
    df = (df_api.groupby(["Administradora", "Condominio"], dropna=False)
                .agg(Servidor=("Servidor",           _agg_unique),
                     UsuarioAPI=("Usuario",          _agg_unique),
                     ID_Condominio=("ID_Condominio", _agg_unique))
                .reset_index()
                .rename(columns={"Administradora": "Adm", "Condominio": "Nome"}))
    return df[["Adm", "Nome", "Servidor", "UsuarioAPI", "ID_Condominio"]] \
              .sort_values(["Adm", "Nome"]).reset_index(drop=True)


def unids_scai(df_scai: pd.DataFrame) -> pd.DataFrame:
    """Lista de unidades cadastradas — uma por (Adm, Condominio, Unidade)."""
    if df_scai.empty:
        return pd.DataFrame(columns=["Adm", "Condominio", "Unidade", "CodCliente", "LoginScai"])
    s = df_scai[df_scai["Unid"].fillna("").astype(str).str.strip() != ""].copy()
    df = (s.groupby(["Adm", "Condo", "Unid"], dropna=False)
           .agg(CodCliente=("CodCliente", _agg_unique),
                LoginScai=("LogAdm",      _agg_unique))
           .reset_index()
           .rename(columns={"Condo": "Condominio", "Unid": "Unidade"}))
    return df[["Adm", "Condominio", "Unidade", "CodCliente", "LoginScai"]] \
              .sort_values(["Adm", "Condominio", "Unidade"]).reset_index(drop=True)


def unids_api(df_api_unids: pd.DataFrame) -> pd.DataFrame:
    """Lista de unidades da API — uma por (Adm, Condominio, Unidade)."""
    if df_api_unids.empty:
        return pd.DataFrame(columns=["Adm", "Condominio", "Unidade",
                                     "Servidor", "UsuarioAPI", "ID_Unidade"])
    a = df_api_unids[df_api_unids["Unidade"].fillna("").astype(str).str.strip() != ""].copy()
    df = (a.groupby(["Administradora", "Condominio", "Unidade"], dropna=False)
           .agg(Servidor=("Servidor",     _agg_unique),
                UsuarioAPI=("Usuario",    _agg_unique),
                ID_Unidade=("ID_Unidade", _agg_unique))
           .reset_index()
           .rename(columns={"Administradora": "Adm"}))
    return df[["Adm", "Condominio", "Unidade", "Servidor", "UsuarioAPI", "ID_Unidade"]] \
              .sort_values(["Adm", "Condominio", "Unidade"]).reset_index(drop=True)


# ═══════════════════════════════════════════════════════════════════════════════
#  Execução
# ═══════════════════════════════════════════════════════════════════════════════

def main():
    admins_filtro = [a.strip() for a in sys.argv[1:] if a.strip()] or None

    console.print("[bold cyan]ListarRevisao[/bold cyan] — lendo inventário gerado...")
    df_api_condos, df_api_unids, _ts = _ler_inventario()
    console.print(f"  {len(df_api_condos)} linhas de condomínios na API")
    console.print(f"  {len(df_api_unids)} linhas de unidades na API")

    # Recolhe admins que apareceram no inventário
    admins_api = set(df_api_condos["Administradora"].dropna().unique()) \
                 | set(df_api_unids["Administradora"].dropna().unique())

    console.print("[bold cyan]→[/bold cyan] lendo cadastro Scai...")
    df_scai = _ler_scai(admins_filtro)
    console.print(f"  {len(df_scai)} linhas em Inquilin_New "
                  f"(filtro={admins_filtro or 'nenhum'})")

    # Se filtro foi passado, restringe API também
    if admins_filtro:
        df_api_condos = df_api_condos[df_api_condos["Administradora"].isin(admins_filtro)]
        df_api_unids  = df_api_unids[df_api_unids["Administradora"].isin(admins_filtro)]

    console.print("[bold cyan]→[/bold cyan] montando listas...")
    df_cs = condos_scai(df_scai)
    df_ca = condos_api(df_api_condos)
    df_us = unids_scai(df_scai)
    df_ua = unids_api(df_api_unids)
    console.print(f"  Condominios_Scai: {len(df_cs):>4} linhas")
    console.print(f"  Condominios_API:  {len(df_ca):>4} linhas")
    console.print(f"  Unidades_Scai:    {len(df_us):>4} linhas")
    console.print(f"  Unidades_API:     {len(df_ua):>4} linhas")

    ts = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    arq = LOGS_DIR / f"Lista_Revisao_{ts}.xlsx"
    with pd.ExcelWriter(arq, engine="openpyxl") as w:
        df_cs.to_excel(w, sheet_name="Condominios_Scai", index=False)
        df_ca.to_excel(w, sheet_name="Condominios_API",  index=False)
        df_us.to_excel(w, sheet_name="Unidades_Scai",    index=False)
        df_ua.to_excel(w, sheet_name="Unidades_API",     index=False)

    console.print()
    console.print("[bold]" + "=" * 70 + "[/bold]")
    console.print(f"[cyan]Arquivo gerado:[/cyan] {arq}")
    console.print(
        "[yellow]Use as abas em pares (Condominios_Scai vs Condominios_API; "
        "Unidades_Scai vs Unidades_API). Filtre por Adm no Excel pra "
        "isolar uma administradora de cada vez na comparação.[/yellow]"
    )


if __name__ == "__main__":
    main()
