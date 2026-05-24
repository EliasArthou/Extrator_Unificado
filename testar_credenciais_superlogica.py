"""
testar_credenciais_superlogica.py — Diagnóstico standalone.

Para cada administradora superlogica que aparece como "em manutenção" no log
do extracao_api.py, pega a primeira credencial não-vazia do banco e tenta
logar via HTTP. Reporta o resultado real para distinguir:

  • Login OK              → credencial válida (a "manutenção" é mensagem
                             de fato do portal e está aberto em outras horas)
  • "em manutenção"        → mesma mensagem ambígua (memory: pode mascarar
                             credencial inválida)
  • E-mail/senha inválido  → credencial inválida confirmada
  • Outro erro             → reporta literal

NÃO altera nada no projeto, só lê e tenta login.
"""

import os
import sys
import traceback

# Garante que importa do projeto corrente
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from dotenv import load_dotenv
load_dotenv()

import auxiliares as aux
import extracao_api as api


# ── Administradoras a testar (substring case-insensitive do NomeAdm) ─────────
ADMINS_ALVO = [
    "QUALITY HOUSE",
    "PROTEST",
    "PRED RIO",
    "LUMEN",
    "IRIGON",
]


def _mascarar(s: str) -> str:
    if not s:
        return "(vazio)"
    if len(s) <= 3:
        return "*" * len(s)
    return s[:2] + "*" * (len(s) - 2)


def _testar_credencial(nome_adm: str, site: str, usuario: str, senha: str):
    """Tenta login e retorna (icone, status_curto, detalhe)."""
    try:
        sessao = api.SessaoSuperlogica(site)
        sessao.login(usuario, senha)
    except ValueError as e:
        msg = str(e).lower()
        if "manutenção" in msg or "manutencao" in msg:
            return ("🔧", "EM MANUTENÇÃO (ambíguo)", str(e))
        if "e-mail inválido" in msg or "email inválido" in msg:
            return ("❌", "EMAIL INVÁLIDO", str(e))
        if "e-mail ou senha" in msg or "senha invalido" in msg or "senha inválido" in msg:
            return ("❌", "SENHA INVÁLIDA", str(e))
        return ("⚠", "OUTRO ERRO DE LOGIN", str(e))
    except Exception as e:
        return ("🔌", "ERRO DE REDE/HTTP", f"{type(e).__name__}: {e}")

    # Login OK — tenta listar condomínios
    try:
        condos = sessao.listar_condominios()
        return ("✅", "LOGIN OK", f"{len(condos)} condomínio(s) acessíveis")
    except Exception as e:
        return ("✅", "LOGIN OK (mas erro ao listar)", f"{type(e).__name__}: {e}")


def main():
    print("=" * 78)
    print(" Teste de credenciais — bases superlogica 'em manutenção'")
    print("=" * 78)

    caminho_banco = os.path.join(aux.caminhoprojeto(), "Scai.WMB")
    banco = aux.Banco(caminho_banco)

    for filtro in ADMINS_ALVO:
        print(f"\n─── {filtro!r} " + "─" * (70 - len(filtro)))

        # Busca 1 credencial não-vazia para qualquer NomeAdm que contenha o filtro
        sql = f"""
            SELECT DISTINCT
                   TRIM(NomeAdm) AS Nome,
                   TRIM(LogAdm)  AS Login_,
                   TRIM(SenAdm)  AS Senha_
              FROM Inquilin_New
             WHERE UCase(TRIM(NomeAdm)) LIKE '%{filtro.upper()}%'
               AND TRIM(LogAdm) <> ''
               AND TRIM(SenAdm) <> ''
        """
        try:
            linhas = banco.consultar(sql)
        except Exception as e:
            print(f"  ❌ Erro SQL: {e}")
            continue

        if not linhas:
            print(f"  ⚠ Sem credenciais preenchidas (LogAdm/SenAdm vazios) "
                  f"para NomeAdm contendo '{filtro}'")
            continue

        # Agrupa por NomeAdm exato
        por_nome: dict = {}
        for nome, login, senha in linhas:
            por_nome.setdefault(nome, []).append((login, senha))

        for nome_adm, creds in por_nome.items():
            site = api.MAPA_SITES.get(nome_adm)
            print(f"\n  • Administradora: {nome_adm!r}")
            print(f"    Site no .env:   {site or '(não mapeado)'}")
            if not site:
                print(f"    Status:          ⚠ Pulado — NomeAdm não está no MAPA_SITES")
                continue

            # Testa a 1ª credencial
            login_str, senha_str = creds[0]
            print(f"    Login testado:  {login_str}")
            print(f"    Senha testada:  {_mascarar(senha_str)} ({len(senha_str)} chars)")
            print(f"    Total creds disponíveis para esta adm: {len(creds)}")

            icone, status, detalhe = _testar_credencial(
                nome_adm, site, login_str, senha_str
            )
            print(f"    Resultado:      {icone} {status}")
            print(f"    Detalhe:        {detalhe}")

    print("\n" + "=" * 78)
    print(" Concluído.")
    print("=" * 78)


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\nInterrompido.")
        sys.exit(130)
    except Exception:
        traceback.print_exc()
        sys.exit(1)
