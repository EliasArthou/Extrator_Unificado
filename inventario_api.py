"""
inventario_api.py — Inventário de condomínios e unidades por credencial.

Loga UMA VEZ em cada credencial (NomeAdm + LogAdm + SenAdm) cadastrada no
banco e dump tudo o que cada credencial enxerga no portal — independente de
existir boleto em aberto. Serve como "fonte da verdade" para padronizar o
cadastro interno com o nome usado na API.

Saídas (em logs/):
  • Inventario_Condominios_<timestamp>.xlsx        → 1 linha por (credencial × condomínio)
  • Inventario_Condominios_x_Unidades_<timestamp>.xlsx → 1 linha por (credencial × condomínio × unidade)

Famílias suportadas: Superlogica, LiveFacilities, Condomob, CIPA, Webware.
Protel e Nacional não expõem listagem dedicada e ficam de fora.

Uso:
  py inventario_api.py                          # todas as administradoras, todas as credenciais
  py inventario_api.py CIPA                     # filtra por administradora
  py inventario_api.py CIPA "QUALITY HOUSE"     # várias administradoras
  py inventario_api.py --multiplos              # SÓ logins múltiplos (uma credencial usada por >1 condomínio)
  py inventario_api.py --multiplos CIPA         # logins múltiplos + filtro por administradora
"""

from __future__ import annotations

import sys
import os
import datetime
from pathlib import Path
from urllib.parse import urlparse, parse_qs

import pandas as pd
from rich.console import Console

from dotenv import load_dotenv
load_dotenv()

# Reaproveita as classes e helpers do extrator principal
from extracao_api import (
    MAPA_SITES,
    _familia_from_url,
    _base_url,
    FAMILIAS_SUPORTADAS,
    SessaoSuperlogica,
    SessaoLiveFacilities,
    SessaoCondomob,
    SessaoCipa,
    SessaoWebware,
)
import auxiliares as aux


console = Console(force_terminal=True, color_system="truecolor")

ROOT       = Path(__file__).parent
LOGS_DIR   = ROOT / "logs"
LOGS_DIR.mkdir(parents=True, exist_ok=True)
BANCO_PATH = r"C:\Users\efart\OneDrive\Área de Trabalho\Scai_local.WMB"

# Famílias que conseguimos inventariar via listagem nativa
FAMILIAS_INVENTARIO = {"superlogica", "livefacilities", "cipa", "condomob", "webware"}

# Nome amigável do "Servidor do Serviço" por família
SERVIDOR_LABEL = {
    "superlogica":    "Superlogica",
    "livefacilities": "LiveFacilities",
    "cipa":           "CIPA",
    "condomob":       "Condomob",
    "webware":        "Webware",
}


# ═══════════════════════════════════════════════════════════════════════════════
#  Coleta de credenciais do banco
# ═══════════════════════════════════════════════════════════════════════════════

def coletar_credenciais(filtro: list | None, apenas_multiplos: bool = False) -> list[tuple]:
    """Lê do banco as triplas distintas (NomeAdm, LogAdm, SenAdm) ativas.

    Se `apenas_multiplos=True`, devolve apenas as credenciais que atendem
    a definição de "Login Múltiplo" usada no resto do sistema:
    mesma (NomeAdm, LogAdm, SenAdm) cobrindo >1 CodCliente (LEFT(Codigo,4))
    distinto na tabela Inquilin_New.
    """
    if filtro:
        admins_alvo = [a.strip() for a in filtro]
    else:
        admins_alvo = [
            nome for nome, site in MAPA_SITES.items()
            if _familia_from_url(site) in FAMILIAS_INVENTARIO
        ]

    if not admins_alvo:
        return []

    formatado = ", ".join(f"'{a.replace(chr(39), chr(39)*2)}'" for a in admins_alvo)

    if apenas_multiplos:
        # Agrupa por credencial e mantém só as que servem >1 condomínio.
        sql = f"""
            SELECT T.AdmTrim AS Nome,
                   T.LogTrim AS Login,
                   T.SenTrim AS Senha
              FROM (SELECT DISTINCT
                           TRIM(NomeAdm) AS AdmTrim,
                           TRIM(LogAdm)  AS LogTrim,
                           TRIM(SenAdm)  AS SenTrim,
                           LEFT(TRIM(Codigo),4) AS CodCliente
                      FROM Inquilin_New
                     WHERE Situa NOT IN ('E','F','K','V')
                           AND TRIM(NomeAdm) <> ''
                           AND TRIM(LogAdm)  <> ''
                           AND TRIM(SenAdm)  <> ''
                           AND TRIM(NomeAdm) IN ({formatado})) AS T
             GROUP BY T.AdmTrim, T.LogTrim, T.SenTrim
            HAVING COUNT(T.CodCliente) > 1
             ORDER BY T.AdmTrim, T.LogTrim
        """
    else:
        sql = f"""
            SELECT DISTINCT
                   TRIM(NomeAdm) AS Nome,
                   TRIM(LogAdm)  AS Login,
                   TRIM(SenAdm)  AS Senha
              FROM Inquilin_New
             WHERE Situa NOT IN ('E','F','K','V')
                   AND TRIM(NomeAdm) <> ''
                   AND TRIM(LogAdm)  <> ''
                   AND TRIM(SenAdm)  <> ''
                   AND TRIM(NomeAdm) IN ({formatado})
             ORDER BY TRIM(NomeAdm), TRIM(LogAdm)
        """

    banco = aux.Banco(BANCO_PATH)
    return banco.consultar(sql)


# ═══════════════════════════════════════════════════════════════════════════════
#  Listagem por família
# ═══════════════════════════════════════════════════════════════════════════════

def _row_condo(servidor, adm, usuario, condo, id_condo) -> dict:
    return {
        "Servidor":      servidor,
        "Administradora": adm,
        "Usuario":       usuario,
        "Condominio":    condo,
        "ID_Condominio": str(id_condo) if id_condo is not None else "",
    }


def _row_unid(servidor, adm, usuario, condo, id_condo, unidade, id_unidade) -> dict:
    return {
        "Servidor":      servidor,
        "Administradora": adm,
        "Usuario":       usuario,
        "Condominio":    condo,
        "ID_Condominio": str(id_condo) if id_condo is not None else "",
        "Unidade":       unidade,
        "ID_Unidade":    str(id_unidade) if id_unidade is not None else "",
    }


def _inventario_superlogica(site, nome_adm, login, senha):
    # O construtor já lida com 'licenca' embutida na query string,
    # ou subdomínio (qualityhouse.superlogica.net) — ambos os formatos.
    sessao = SessaoSuperlogica(site)
    sessao.login(login, senha)
    raw = sessao.listar_condominios()  # dict {id: nome | dict}

    condos = []
    for cid, info in (raw or {}).items():
        if isinstance(info, dict):
            nome = info.get("nome") or info.get("name") or info.get("st_nome_cond") or str(info)
        else:
            nome = str(info)
        condos.append(_row_condo(SERVIDOR_LABEL["superlogica"], nome_adm, login, nome, cid))
    return condos, []  # Superlogica não tem nível "unidade" separado


def _inventario_livefacilities(site, nome_adm, login, senha):
    sessao = SessaoLiveFacilities(_base_url(site))
    sessao.login(login, senha)
    unids = sessao.listar_unidades()

    condos_dedup = {}
    unidades = []
    for u in unids:
        cnome = (u.get("condo") or "").strip()
        bloco   = (u.get("bloco")   or "").strip()
        unidade = (u.get("unidade") or "").strip()
        rotulo_unid = " ".join(p for p in [bloco, unidade] if p).strip() or unidade

        if cnome and cnome not in condos_dedup:
            condos_dedup[cnome] = True

        unidades.append(_row_unid(
            SERVIDOR_LABEL["livefacilities"], nome_adm, login,
            cnome, "", rotulo_unid, "",
        ))

    condos = [
        _row_condo(SERVIDOR_LABEL["livefacilities"], nome_adm, login, cnome, "")
        for cnome in condos_dedup
    ]
    return condos, unidades


def _inventario_cipa(site, nome_adm, login, senha):
    sessao = SessaoCipa()
    sessao.login(login, senha)
    sessao.carregar_perfil()
    sessao.carregar_condominios()

    condos = [
        _row_condo(
            SERVIDOR_LABEL["cipa"], nome_adm, login,
            sessao._condo_names.get(cid, ""), cid,
        )
        for cid in sessao._condo_names
    ]

    unidades = []
    for u in (sessao._units or []):
        cid    = u.get("condominium_id", "")
        cnome  = sessao._condo_names.get(cid, "")
        rotulo = " ".join(str(x) for x in [u.get("code", ""), u.get("name", "")] if x)
        unidades.append(_row_unid(
            SERVIDOR_LABEL["cipa"], nome_adm, login,
            cnome, cid, rotulo.strip(), u.get("id", ""),
        ))
    return condos, unidades


def _inventario_condomob(site, nome_adm, login, senha):
    qs = parse_qs(urlparse(site).query)
    adm_id = qs.get("administradora", [None])[0]
    if not adm_id:
        raise ValueError("URL Condomob sem 'administradora_id'")

    sessao = SessaoCondomob(adm_id)
    sessao.login(login, senha)

    condos_api = sessao.listar_condominios() or []
    condos, unidades = [], []
    for c in condos_api:
        cid   = c.get("id", "")
        cnome = c.get("nome", c.get("name", ""))
        condos.append(_row_condo(
            SERVIDOR_LABEL["condomob"], nome_adm, login, cnome, cid,
        ))
        try:
            unids = sessao.listar_unidades(cid) or []
        except Exception:
            unids = []
        for u in unids:
            unidades.append(_row_unid(
                SERVIDOR_LABEL["condomob"], nome_adm, login,
                cnome, cid, u.get("identificador", ""), u.get("id", ""),
            ))
    return condos, unidades


def _inventario_webware(site, nome_adm, login, senha):
    sessao = SessaoWebware(site)
    sessao.login(login, senha)
    cards = sessao.listar_cards_condominios() or []

    # Em Webware cada card é tipicamente um condomínio acessível.
    condos = [
        _row_condo(SERVIDOR_LABEL["webware"], nome_adm, login,
                   c.get("titulo", ""), "")
        for c in cards
    ]
    # Webware não expõe lista de unidades separada via cards.
    return condos, []


_DISPATCH = {
    "superlogica":    _inventario_superlogica,
    "livefacilities": _inventario_livefacilities,
    "cipa":           _inventario_cipa,
    "condomob":       _inventario_condomob,
    "webware":        _inventario_webware,
}


def coletar_inventario(nome_adm, login, senha) -> tuple[list, list, str]:
    """Roteia para a família correta. Retorna (condos, unidades, status)."""
    site = MAPA_SITES.get(nome_adm, "")
    if not site:
        return [], [], "sem_site"

    fam = _familia_from_url(site)
    handler = _DISPATCH.get(fam)
    if not handler:
        return [], [], f"familia_nao_suportada:{fam}"

    try:
        condos, unidades = handler(site, nome_adm, login, senha)
        return condos, unidades, "ok"
    except Exception as e:
        return [], [], f"erro:{type(e).__name__}:{e}"


# ═══════════════════════════════════════════════════════════════════════════════
#  Execução principal
# ═══════════════════════════════════════════════════════════════════════════════

def main():
    args = [a.strip() for a in sys.argv[1:] if a.strip()]
    apenas_multiplos = False
    if "--multiplos" in args or "-m" in args:
        apenas_multiplos = True
        args = [a for a in args if a not in ("--multiplos", "-m")]
    filtro = args or None

    console.print("[bold cyan]InventarioAPI[/bold cyan] — coletando credenciais do banco...")
    if apenas_multiplos:
        console.print("  [yellow]Modo --multiplos: somente credenciais que servem >1 condomínio[/yellow]")
    creds = coletar_credenciais(filtro, apenas_multiplos=apenas_multiplos)
    console.print(f"  {len(creds)} credencial(is) distinta(s) encontrada(s)")

    if not creds:
        console.print("[yellow]Nenhuma credencial para processar.[/yellow]")
        return

    todos_condos: list[dict] = []
    todas_unids:  list[dict] = []
    erros:        list[dict] = []

    for idx, (nome_adm, login, senha) in enumerate(creds, 1):
        site = MAPA_SITES.get(nome_adm, "")
        fam  = _familia_from_url(site) if site else "?"
        prefixo = f"[{idx:>3}/{len(creds)}] {nome_adm:<35} ({fam:<14}) {login}"
        console.print(prefixo + " ... ", end="")

        condos, unids, status = coletar_inventario(nome_adm, login, senha)

        if status == "ok":
            console.print(f"[green]OK[/green] (condos={len(condos)}, unids={len(unids)})")
            todos_condos.extend(condos)
            todas_unids.extend(unids)
        else:
            console.print(f"[red]{status}[/red]")
            erros.append({
                "Administradora": nome_adm,
                "Servidor":       SERVIDOR_LABEL.get(fam, fam),
                "Usuario":        login,
                "Status":         status,
            })

    # ── Saída ──────────────────────────────────────────────────────────────────
    ts = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    arq_condos = LOGS_DIR / f"Inventario_Condominios_{ts}.xlsx"
    arq_unids  = LOGS_DIR / f"Inventario_Condominios_x_Unidades_{ts}.xlsx"

    df_condos = pd.DataFrame(todos_condos, columns=[
        "Servidor", "Administradora", "Usuario", "Condominio", "ID_Condominio",
    ])
    df_unids = pd.DataFrame(todas_unids, columns=[
        "Servidor", "Administradora", "Usuario",
        "Condominio", "ID_Condominio", "Unidade", "ID_Unidade",
    ])
    df_erros = pd.DataFrame(erros, columns=[
        "Administradora", "Servidor", "Usuario", "Status",
    ])

    # Arquivo 1 — Condomínios
    with pd.ExcelWriter(arq_condos, engine="openpyxl") as w:
        df_condos.to_excel(w, sheet_name="Condominios", index=False)
        if not df_erros.empty:
            df_erros.to_excel(w, sheet_name="Erros", index=False)

    # Arquivo 2 — Condomínios × Unidades
    with pd.ExcelWriter(arq_unids, engine="openpyxl") as w:
        df_unids.to_excel(w, sheet_name="Condominios_x_Unidades", index=False)
        if not df_erros.empty:
            df_erros.to_excel(w, sheet_name="Erros", index=False)

    # ── Resumo ─────────────────────────────────────────────────────────────────
    console.print()
    console.print(f"[bold]{'='*70}[/bold]")
    console.print(f"[bold green]RESUMO[/bold green]  "
                  f"Condomínios=[bold]{len(df_condos)}[/bold]  "
                  f"Unidades=[bold]{len(df_unids)}[/bold]  "
                  f"Credenciais com erro=[bold red]{len(df_erros)}[/bold red]")
    console.print(f"[bold]{'='*70}[/bold]")
    console.print(f"\n[cyan]Condomínios:[/cyan]            {arq_condos}")
    console.print(f"[cyan]Condomínios × Unidades:[/cyan]  {arq_unids}")


if __name__ == "__main__":
    main()
