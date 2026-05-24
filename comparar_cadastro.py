"""
comparar_cadastro.py — Compara cadastro Scai (Inquilin_New) com inventário da API.

Lê os dois Excels mais recentes em logs/ (gerados por inventario_api.py) e cruza
com o cadastro interno, usando fuzzy matching pra identificar erros de
nomenclatura/digitação.

Saída: logs/Divergencias_Cadastro_<ts>.xlsx, com abas:
  • Condo_QUASE_IGUAL    → score 85-99 (ALVO — ajustar Scai pra virar igual)
  • Condo_REVISAR        → score 60-84 (talvez sejam o mesmo, com diferenças grandes)
  • Condo_SO_NO_SCAI     → cadastrado mas API não enxerga
  • Condo_SO_NA_API      → API enxerga mas Scai não tem
  • Condo_ADM_SEM_INVENTARIO → admins do Scai sem inventário gerado
  • Condo_IGUAL          → já estão alinhados (não precisa ação)
  • Unid_QUASE_IGUAL / Unid_REVISAR / Unid_SO_NA_API / Unid_IGUAL

Uso (padrão filtra apenas credenciais "multi-seleção"):
  py comparar_cadastro.py             # só credenciais com lista (>1 condo OU >1 unid)
  py comparar_cadastro.py CIPA        # idem, filtrado por administradora
  py comparar_cadastro.py --todos     # inclui credenciais de 1-condo + 1-unid também

"Multi-seleção" = ao logar com a credencial, o portal exibe uma LISTA com
mais de 1 item (condomínio ou unidade) — i.e. o operador precisa selecionar.
"""

from __future__ import annotations

import sys
import re
import unicodedata
import datetime
from pathlib import Path

import pandas as pd
from rich.console import Console
from rapidfuzz import fuzz, process

from dotenv import load_dotenv
load_dotenv()

import auxiliares as aux


console = Console(force_terminal=True, color_system="truecolor")

ROOT       = Path(__file__).parent
LOGS_DIR   = ROOT / "logs"
BANCO_PATH = r"C:\Users\efart\OneDrive\Área de Trabalho\Scai_local.WMB"

# Limites de similaridade (rapidfuzz WRatio: 0-100)
SCORE_IGUAL   = 100   # bate caractere por caractere (após normalização)
SCORE_QUASE   = 85    # quase igual — provavelmente o mesmo, com typo/abreviação
SCORE_REVISAR = 60    # candidato fraco — revisar manualmente


# ═══════════════════════════════════════════════════════════════════════════════
#  Normalização e classificação
# ═══════════════════════════════════════════════════════════════════════════════

def _normalizar(s) -> str:
    """Uppercase, sem acentos, sem prefixo de código tipo '0711 - ', sem pontuação."""
    if s is None:
        return ""
    s = str(s).strip().upper()
    s = "".join(c for c in unicodedata.normalize("NFD", s)
                if unicodedata.category(c) != "Mn")
    # Strip prefixos como "0711 - ", "1234-", "01 - " — IDs vindos da API
    s = re.sub(r"^\d+\s*[-–]\s*", "", s)
    s = re.sub(r"[^A-Z0-9 ]+", " ", s)
    s = re.sub(r"\s+", " ", s).strip()
    return s


def _classificar(score: float) -> str:
    if score >= SCORE_IGUAL:    return "IGUAL"
    if score >= SCORE_QUASE:    return "QUASE_IGUAL"
    if score >= SCORE_REVISAR:  return "REVISAR"
    return "SEM_MATCH"


# ═══════════════════════════════════════════════════════════════════════════════
#  Leitura dos dados
# ═══════════════════════════════════════════════════════════════════════════════

def _encontrar_excel_mais_recente(padrao: str) -> Path | None:
    """Aceita um glob específico — ex: 'Inventario_Condominios_[0-9]*.xlsx'."""
    cands = sorted(LOGS_DIR.glob(padrao), reverse=True)
    return cands[0] if cands else None


def _ler_inventario() -> tuple[pd.DataFrame, pd.DataFrame, str]:
    # Timestamp começa com dígito → diferencia 'Inventario_Condominios_2026...'
    # de 'Inventario_Condominios_x_Unidades_2026...'
    arq_condos = _encontrar_excel_mais_recente("Inventario_Condominios_[0-9]*.xlsx")
    arq_unids  = _encontrar_excel_mais_recente("Inventario_Condominios_x_Unidades_*.xlsx")
    if not arq_condos or not arq_unids:
        raise FileNotFoundError(
            f"Não encontrei Inventario_*.xlsx em {LOGS_DIR}.\n"
            "Rode 'python inventario_api.py' antes."
        )
    console.print(f"  API condomínios: [dim]{arq_condos.name}[/dim]")
    console.print(f"  API unidades:    [dim]{arq_unids.name}[/dim]")
    df_c = pd.read_excel(arq_condos, sheet_name="Condominios")
    df_u = pd.read_excel(arq_unids,  sheet_name="Condominios_x_Unidades")
    ts = arq_condos.stem.replace("Inventario_Condominios_", "")
    return df_c, df_u, ts


def _filtrar_multi_selecao(df_condos: pd.DataFrame,
                            df_unids: pd.DataFrame
                            ) -> tuple[pd.DataFrame, pd.DataFrame]:
    """
    Mantém só as credenciais (Servidor, Administradora, Usuario) que oferecem
    "multi-seleção": >1 condomínio acessível OU >1 unidade em algum condomínio.
    """
    chave = ["Servidor", "Administradora", "Usuario"]

    # Credenciais com >1 condomínio acessível
    if df_condos.empty:
        multi_c = pd.DataFrame(columns=chave)
    else:
        c = (df_condos.groupby(chave)["Condominio"].nunique()
                       .reset_index(name="n"))
        multi_c = c[c["n"] > 1][chave]

    # Credenciais com algum condomínio onde >1 unidade
    if df_unids.empty:
        multi_u = pd.DataFrame(columns=chave)
    else:
        u = (df_unids.groupby(chave + ["Condominio"])["Unidade"].nunique()
                      .reset_index(name="n"))
        multi_u = u[u["n"] > 1][chave].drop_duplicates()

    creds = pd.concat([multi_c, multi_u]).drop_duplicates()
    if creds.empty:
        return df_condos.iloc[0:0], df_unids.iloc[0:0]

    df_c2 = df_condos.merge(creds, on=chave, how="inner")
    df_u2 = (df_unids.merge(creds, on=chave, how="inner")
             if not df_unids.empty else df_unids)
    return df_c2, df_u2


def _ler_scai(admins_filtro: list | None) -> pd.DataFrame:
    # MS Access: alias não pode ter o mesmo nome do campo (referência circular).
    # CodCliente = LEFT(TRIM(Codigo),4) — mesma definição do helper_upload.py.
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


def _logins_scai_por_condo(df_scai: pd.DataFrame) -> dict:
    """Mapa (Adm, Condo) → string com logins distintos (separados por vírgula)."""
    g = (df_scai.dropna(subset=["LogAdm"])
                 .groupby(["Adm", "Condo"])["LogAdm"]
                 .agg(lambda s: ", ".join(sorted({str(x).strip() for x in s if str(x).strip()}))))
    return g.to_dict()


def _logins_scai_por_unidade(df_scai: pd.DataFrame) -> dict:
    """Mapa (Adm, Condo, Unid) → string com logins distintos (separados por vírgula)."""
    g = (df_scai.dropna(subset=["LogAdm"])
                 .groupby(["Adm", "Condo", "Unid"])["LogAdm"]
                 .agg(lambda s: ", ".join(sorted({str(x).strip() for x in s if str(x).strip()}))))
    return g.to_dict()


def _codclientes_por_condo(df_scai: pd.DataFrame) -> dict:
    """Mapa (Adm, Condo) → string com códigos de cliente distintos."""
    g = (df_scai.dropna(subset=["CodCliente"])
                 .groupby(["Adm", "Condo"])["CodCliente"]
                 .agg(lambda s: ", ".join(sorted({str(x).strip() for x in s if str(x).strip()}))))
    return g.to_dict()


def _codclientes_por_unidade(df_scai: pd.DataFrame) -> dict:
    """Mapa (Adm, Condo, Unid) → string com códigos de cliente distintos."""
    g = (df_scai.dropna(subset=["CodCliente"])
                 .groupby(["Adm", "Condo", "Unid"])["CodCliente"]
                 .agg(lambda s: ", ".join(sorted({str(x).strip() for x in s if str(x).strip()}))))
    return g.to_dict()


# ═══════════════════════════════════════════════════════════════════════════════
#  Comparação fuzzy
# ═══════════════════════════════════════════════════════════════════════════════

def comparar_condos(df_scai: pd.DataFrame, df_api: pd.DataFrame) -> pd.DataFrame:
    """
    Para cada administradora em comum, faz best-match Scai ↔ API.
    Retorna DataFrame:
      Adm | NomeScai | NomeAPI | Score | Classificacao
      | LoginScai | ID_Condominio | Servidor | UsuarioAPI
    """
    resultados = []
    admins_scai = set(df_scai["Adm"].unique())
    admins_api  = set(df_api["Administradora"].unique())
    admins_ambos = sorted(admins_scai & admins_api)
    mapa_log_scai = _logins_scai_por_condo(df_scai)
    mapa_cod_scai = _codclientes_por_condo(df_scai)

    console.print(f"  {len(admins_ambos)} administradora(s) em comum")

    for adm in admins_ambos:
        scai_nomes = sorted(df_scai[df_scai["Adm"] == adm]["Condo"].unique())
        api_subset = df_api[df_api["Administradora"] == adm]
        api_nomes  = sorted(api_subset["Condominio"].dropna().unique())

        if not api_nomes:
            for nome_scai in scai_nomes:
                resultados.append(dict(
                    Adm=adm, CodCliente=mapa_cod_scai.get((adm, nome_scai), ""),
                    NomeScai=nome_scai, NomeAPI="",
                    Score=0, Classificacao="SO_NO_SCAI",
                    LoginScai=mapa_log_scai.get((adm, nome_scai), ""),
                    ID_Condominio="", Servidor="", UsuarioAPI="",
                ))
            continue

        api_norm = {_normalizar(n): n for n in api_nomes}
        api_norm_keys = list(api_norm.keys())

        usados_api = set()
        for nome_scai in scai_nomes:
            chave = _normalizar(nome_scai)
            best = process.extractOne(chave, api_norm_keys, scorer=fuzz.WRatio)
            if best is None:
                continue
            api_key, score, _ = best
            nome_api = api_norm[api_key]
            clf = _classificar(score)

            # Match fraco (score < SCORE_REVISAR) = condômino cadastrado mas
            # nenhum nome decente correspondente na API → "só no Scai".
            if clf == "SEM_MATCH":
                resultados.append(dict(
                    Adm=adm, CodCliente=mapa_cod_scai.get((adm, nome_scai), ""),
                    NomeScai=nome_scai, NomeAPI="",
                    Score=round(score, 1), Classificacao="SO_NO_SCAI",
                    LoginScai=mapa_log_scai.get((adm, nome_scai), ""),
                    ID_Condominio="", Servidor="", UsuarioAPI="",
                ))
                continue

            # Detalhes da linha API correspondente (1ª ocorrência basta)
            row_api = api_subset[api_subset["Condominio"] == nome_api].iloc[0]
            resultados.append(dict(
                Adm=adm,
                CodCliente=mapa_cod_scai.get((adm, nome_scai), ""),
                NomeScai=nome_scai,
                NomeAPI=nome_api,
                Score=round(score, 1),
                Classificacao=clf,
                LoginScai=mapa_log_scai.get((adm, nome_scai), ""),
                ID_Condominio=row_api.get("ID_Condominio", ""),
                Servidor=row_api.get("Servidor", ""),
                UsuarioAPI=row_api.get("Usuario", ""),
            ))
            usados_api.add(nome_api)

        # Condomínios na API que ninguém puxou
        for nome_api in api_nomes:
            if nome_api in usados_api:
                continue
            row_api = api_subset[api_subset["Condominio"] == nome_api].iloc[0]
            resultados.append(dict(
                Adm=adm,
                CodCliente="",
                NomeScai="",
                NomeAPI=nome_api,
                Score=0,
                Classificacao="SO_NA_API",
                LoginScai="",
                ID_Condominio=row_api.get("ID_Condominio", ""),
                Servidor=row_api.get("Servidor", ""),
                UsuarioAPI=row_api.get("Usuario", ""),
            ))

    # Admins só no Scai
    for adm in sorted(admins_scai - admins_api):
        for nome_scai in sorted(df_scai[df_scai["Adm"] == adm]["Condo"].unique()):
            resultados.append(dict(
                Adm=adm,
                CodCliente=mapa_cod_scai.get((adm, nome_scai), ""),
                NomeScai=nome_scai, NomeAPI="",
                Score=0, Classificacao="ADM_SEM_INVENTARIO",
                LoginScai=mapa_log_scai.get((adm, nome_scai), ""),
                ID_Condominio="", Servidor="", UsuarioAPI="",
            ))

    return pd.DataFrame(resultados)


def comparar_unidades(df_scai: pd.DataFrame,
                      df_api_unids: pd.DataFrame,
                      df_match_condos: pd.DataFrame) -> pd.DataFrame:
    """Para cada par de condomínio com match razoável, cruza as unidades."""
    resultados = []
    pares = df_match_condos[df_match_condos["Classificacao"].isin(
        ["IGUAL", "QUASE_IGUAL", "REVISAR"]
    )]
    mapa_log_unid = _logins_scai_por_unidade(df_scai)
    mapa_cod_unid = _codclientes_por_unidade(df_scai)

    console.print(f"  {len(pares)} pares de condomínio com match razoável")

    for _, par in pares.iterrows():
        adm = par["Adm"]
        condo_scai = par["NomeScai"]
        condo_api  = par["NomeAPI"]
        if not condo_scai or not condo_api:
            continue

        scai_unids = sorted(
            df_scai[(df_scai["Adm"] == adm) & (df_scai["Condo"] == condo_scai)]
                   ["Unid"].dropna().unique()
        )
        api_unids = sorted(
            df_api_unids[(df_api_unids["Administradora"] == adm)
                         & (df_api_unids["Condominio"] == condo_api)]
                        ["Unidade"].dropna().unique()
        )
        scai_unids = [u for u in scai_unids if str(u).strip()]
        api_unids  = [u for u in api_unids  if str(u).strip()]
        if not scai_unids or not api_unids:
            continue

        api_norm = {_normalizar(u): u for u in api_unids}
        api_keys = list(api_norm.keys())

        usados = set()
        for u_scai in scai_unids:
            chave = _normalizar(u_scai)
            best = process.extractOne(chave, api_keys, scorer=fuzz.WRatio)
            if best is None:
                continue
            api_key, score, _ = best
            u_api = api_norm[api_key]
            clf = _classificar(score)

            # Match fraco = unidade cadastrada mas sem correspondente decente
            # na API daquele condomínio
            if clf == "SEM_MATCH":
                resultados.append(dict(
                    Adm=adm,
                    CodCliente=mapa_cod_unid.get((adm, condo_scai, u_scai), ""),
                    CondoScai=condo_scai, CondoAPI=condo_api,
                    UnidadeScai=u_scai, UnidadeAPI="",
                    Score=round(score, 1), Classificacao="SO_NO_SCAI",
                    LoginScai=mapa_log_unid.get((adm, condo_scai, u_scai), ""),
                ))
                continue

            resultados.append(dict(
                Adm=adm,
                CodCliente=mapa_cod_unid.get((adm, condo_scai, u_scai), ""),
                CondoScai=condo_scai,
                CondoAPI=condo_api,
                UnidadeScai=u_scai,
                UnidadeAPI=u_api,
                Score=round(score, 1),
                Classificacao=clf,
                LoginScai=mapa_log_unid.get((adm, condo_scai, u_scai), ""),
            ))
            usados.add(u_api)

        for u_api in api_unids:
            if u_api in usados:
                continue
            resultados.append(dict(
                Adm=adm,
                CodCliente="",
                CondoScai=condo_scai,
                CondoAPI=condo_api,
                UnidadeScai="",
                UnidadeAPI=u_api,
                Score=0,
                Classificacao="SO_NA_API",
                LoginScai="",
            ))

    return pd.DataFrame(resultados)


# ═══════════════════════════════════════════════════════════════════════════════
#  Execução
# ═══════════════════════════════════════════════════════════════════════════════

def main():
    args = [a.strip() for a in sys.argv[1:] if a.strip()]
    incluir_todos = False
    if "--todos" in args:
        incluir_todos = True
        args = [a for a in args if a != "--todos"]
    admins_filtro = args or None

    console.print("[bold cyan]CompararCadastro[/bold cyan] — lendo inventário gerado...")
    df_api_condos, df_api_unids, _ts = _ler_inventario()
    console.print(f"  {len(df_api_condos)} linhas de condomínios na API")
    console.print(f"  {len(df_api_unids)} linhas de unidades na API")

    if not incluir_todos:
        n_c_antes = len(df_api_condos)
        n_u_antes = len(df_api_unids)
        df_api_condos, df_api_unids = _filtrar_multi_selecao(df_api_condos, df_api_unids)
        console.print(
            f"[yellow]Filtro multi-seleção:[/yellow] mantidas "
            f"{len(df_api_condos)}/{n_c_antes} linhas de condo e "
            f"{len(df_api_unids)}/{n_u_antes} linhas de unidade "
            f"(credenciais que veem lista com >1 item)"
        )
    else:
        console.print("[dim]Modo --todos: sem filtro de multi-seleção.[/dim]")

    console.print("[bold cyan]→[/bold cyan] lendo cadastro Scai...")
    df_scai = _ler_scai(admins_filtro)
    console.print(f"  {len(df_scai)} linhas em Inquilin_New "
                  f"(filtro={admins_filtro or 'nenhum'})")

    console.print("[bold cyan]→[/bold cyan] cruzando condomínios (fuzzy)...")
    df_match_condos = comparar_condos(df_scai, df_api_condos)
    console.print(f"  {len(df_match_condos)} pares avaliados")

    console.print("[bold cyan]→[/bold cyan] cruzando unidades (fuzzy)...")
    df_match_unids = comparar_unidades(df_scai, df_api_unids, df_match_condos)
    console.print(f"  {len(df_match_unids)} pares avaliados")

    # ── Saída ──────────────────────────────────────────────────────────────────
    ts = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    arq = LOGS_DIR / f"Divergencias_Cadastro_{ts}.xlsx"

    abas_condos = ["QUASE_IGUAL", "REVISAR", "SO_NO_SCAI",
                   "SO_NA_API", "ADM_SEM_INVENTARIO", "IGUAL"]
    abas_unids  = ["QUASE_IGUAL", "REVISAR", "SO_NO_SCAI", "SO_NA_API", "IGUAL"]

    with pd.ExcelWriter(arq, engine="openpyxl") as w:
        for clf in abas_condos:
            sub = df_match_condos[df_match_condos["Classificacao"] == clf]
            if not sub.empty:
                aba = f"Condo_{clf}"[:31]
                sub.sort_values(["Adm","Score","NomeScai"], ascending=[True,False,True]
                                ).to_excel(w, sheet_name=aba, index=False)
        for clf in abas_unids:
            sub = df_match_unids[df_match_unids["Classificacao"] == clf]
            if not sub.empty:
                aba = f"Unid_{clf}"[:31]
                sub.sort_values(["Adm","CondoScai","Score","UnidadeScai"],
                                ascending=[True,True,False,True]
                                ).to_excel(w, sheet_name=aba, index=False)

    # ── Resumo ─────────────────────────────────────────────────────────────────
    console.print()
    console.print("[bold]" + "=" * 70 + "[/bold]")
    console.print("[bold green]RESUMO de Condomínios:[/bold green]")
    for clf in abas_condos:
        n = (df_match_condos["Classificacao"] == clf).sum()
        cor = "yellow" if clf == "QUASE_IGUAL" else "white"
        console.print(f"  [{cor}]{clf:<22}[/{cor}] {n}")
    console.print("[bold green]RESUMO de Unidades:[/bold green]")
    for clf in abas_unids:
        n = (df_match_unids["Classificacao"] == clf).sum()
        cor = "yellow" if clf == "QUASE_IGUAL" else "white"
        console.print(f"  [{cor}]{clf:<22}[/{cor}] {n}")
    console.print("[bold]" + "=" * 70 + "[/bold]")
    console.print(f"\n[cyan]Arquivo gerado:[/cyan] {arq}")
    console.print(
        "\n[yellow]Foque primeiro na aba 'Condo_QUASE_IGUAL': "
        "score 85-99 = forte candidato a padronização no Scai "
        "(o nome do site é a fonte da verdade).[/yellow]"
    )


if __name__ == "__main__":
    main()
