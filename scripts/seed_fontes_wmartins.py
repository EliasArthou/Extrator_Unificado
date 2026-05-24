"""
Seed: cria as fontes de dados da W Martins no banco.
Executar uma vez: python seed_fontes_wmartins.py

Pré-requisito: a empresa W Martins já deve existir no banco.
"""

import os
import sys
sys.path.insert(0, os.path.dirname(__file__))

from dotenv import load_dotenv
load_dotenv()

from ui.web.webapp.models import SessionLocal, Empresa, FonteDado

db = SessionLocal()

# Buscar empresa W Martins
empresa = db.query(Empresa).filter(Empresa.nome == "W Martins").first()
if not empresa:
    print("ERRO: Empresa 'W Martins' não encontrada no banco.")
    print("Crie-a primeiro pelo painel admin em /admin/empresas")
    sys.exit(1)

print(f"Empresa encontrada: {empresa.nome} (id={empresa.id})")

# Fontes a criar
FONTES = [
    {
        "tipo_extracao": "Prefeitura__IPTU",
        "nome_fonte": "IPTU",
        "colunas_esperadas": '["Codigo", "IPTU"]',
        "sql_select": (
            "SELECT DISTINCT * FROM ("
            "SELECT Codigo, Prop.IPTU FROM Prop WHERE LTRIM(RTRIM(IPTU))<>'' AND Prop.IPTU <>'00000000' AND Prop.IPTU <> '88888888' "
            "UNION ALL "
            "SELECT Codigo, Prop.IPTU1 FROM Prop WHERE LTRIM(RTRIM(IPTU1))<>'' AND Prop.IPTU1 <>'00000000' AND Prop.IPTU1 <> '88888888' "
            "UNION ALL "
            "SELECT Codigo, Prop.IPTU2 FROM Prop WHERE LTRIM(RTRIM(IPTU2))<>'' AND Prop.IPTU2 <>'00000000' AND Prop.IPTU2 <> '88888888' "
            "UNION ALL "
            "SELECT Codigo, Prop.IPTU3 FROM Prop WHERE LTRIM(RTRIM(IPTU3))<>'' AND Prop.IPTU3 <>'00000000' AND Prop.IPTU3 <> '88888888' "
            "UNION ALL "
            "SELECT Codigo, Prop.IPTU4 FROM Prop WHERE LTRIM(RTRIM(IPTU4))<>'' AND Prop.IPTU4 <>'00000000' AND Prop.IPTU4 <> '88888888' "
            "UNION ALL "
            "SELECT Codigo, Prop.IPTU5 FROM Prop WHERE LTRIM(RTRIM(IPTU5))<>'' AND Prop.IPTU5 <>'00000000' AND Prop.IPTU5 <> '88888888' "
            "UNION ALL "
            "SELECT Codigo, Prop.IPTU6 FROM Prop WHERE LTRIM(RTRIM(IPTU6))<>'' AND Prop.IPTU6 <>'00000000' AND Prop.IPTU6 <> '88888888' "
            "UNION ALL "
            "SELECT Codigo, Prop.IPTU7 FROM Prop WHERE LTRIM(RTRIM(IPTU7))<>'' AND Prop.IPTU7 <>'00000000' AND Prop.IPTU7 <> '88888888' "
            "UNION ALL "
            "SELECT Codigo, Prop.IPTU8 FROM Prop WHERE LTRIM(RTRIM(IPTU8))<>'' AND Prop.IPTU8 <>'00000000' AND Prop.IPTU8 <> '88888888' "
            "UNION ALL "
            "SELECT Codigo, Prop.IPTU9 FROM Prop WHERE LTRIM(RTRIM(IPTU9))<>'' AND Prop.IPTU9 <>'00000000' AND Prop.IPTU9 <> '88888888' "
            "UNION ALL "
            "SELECT Codigo, Prop.IPTU10 FROM Prop WHERE LTRIM(RTRIM(IPTU10))<>'' AND Prop.IPTU10 <>'00000000' AND Prop.IPTU10 <> '88888888') AS IPTUs"
        ),
    },
    {
        "tipo_extracao": "Prefeitura__Nada_Consta",
        "nome_fonte": "Nada Consta",
        "colunas_esperadas": '["Codigo", "IPTU"]',
        "sql_select": (
            "SELECT DISTINCT * FROM ("
            "SELECT Codigo, Prop.IPTU FROM Prop WHERE LTRIM(RTRIM(IPTU))<>'' AND Prop.IPTU <>'00000000' AND Prop.IPTU <> '88888888' "
            "UNION ALL "
            "SELECT Codigo, Prop.IPTU1 FROM Prop WHERE LTRIM(RTRIM(IPTU1))<>'' AND Prop.IPTU1 <>'00000000' AND Prop.IPTU1 <> '88888888' "
            "UNION ALL "
            "SELECT Codigo, Prop.IPTU2 FROM Prop WHERE LTRIM(RTRIM(IPTU2))<>'' AND Prop.IPTU2 <>'00000000' AND Prop.IPTU2 <> '88888888' "
            "UNION ALL "
            "SELECT Codigo, Prop.IPTU3 FROM Prop WHERE LTRIM(RTRIM(IPTU3))<>'' AND Prop.IPTU3 <>'00000000' AND Prop.IPTU3 <> '88888888' "
            "UNION ALL "
            "SELECT Codigo, Prop.IPTU4 FROM Prop WHERE LTRIM(RTRIM(IPTU4))<>'' AND Prop.IPTU4 <>'00000000' AND Prop.IPTU4 <> '88888888' "
            "UNION ALL "
            "SELECT Codigo, Prop.IPTU5 FROM Prop WHERE LTRIM(RTRIM(IPTU5))<>'' AND Prop.IPTU5 <>'00000000' AND Prop.IPTU5 <> '88888888' "
            "UNION ALL "
            "SELECT Codigo, Prop.IPTU6 FROM Prop WHERE LTRIM(RTRIM(IPTU6))<>'' AND Prop.IPTU6 <>'00000000' AND Prop.IPTU6 <> '88888888' "
            "UNION ALL "
            "SELECT Codigo, Prop.IPTU7 FROM Prop WHERE LTRIM(RTRIM(IPTU7))<>'' AND Prop.IPTU7 <>'00000000' AND Prop.IPTU7 <> '88888888' "
            "UNION ALL "
            "SELECT Codigo, Prop.IPTU8 FROM Prop WHERE LTRIM(RTRIM(IPTU8))<>'' AND Prop.IPTU8 <>'00000000' AND Prop.IPTU8 <> '88888888' "
            "UNION ALL "
            "SELECT Codigo, Prop.IPTU9 FROM Prop WHERE LTRIM(RTRIM(IPTU9))<>'' AND Prop.IPTU9 <>'00000000' AND Prop.IPTU9 <> '88888888' "
            "UNION ALL "
            "SELECT Codigo, Prop.IPTU10 FROM Prop WHERE LTRIM(RTRIM(IPTU10))<>'' AND Prop.IPTU10 <>'00000000' AND Prop.IPTU10 <> '88888888') AS IPTUs"
        ),
    },
    {
        "tipo_extracao": "Prefeitura__Certidao_Negativa",
        "nome_fonte": "Certidão Negativa",
        "colunas_esperadas": '["Codigo", "IPTU"]',
        "sql_select": (
            "SELECT DISTINCT * FROM ("
            "SELECT Codigo, Prop.IPTU FROM Prop WHERE LTRIM(RTRIM(IPTU))<>'' AND Prop.IPTU <>'00000000' AND Prop.IPTU <> '88888888' "
            "UNION ALL "
            "SELECT Codigo, Prop.IPTU1 FROM Prop WHERE LTRIM(RTRIM(IPTU1))<>'' AND Prop.IPTU1 <>'00000000' AND Prop.IPTU1 <> '88888888' "
            "UNION ALL "
            "SELECT Codigo, Prop.IPTU2 FROM Prop WHERE LTRIM(RTRIM(IPTU2))<>'' AND Prop.IPTU2 <>'00000000' AND Prop.IPTU2 <> '88888888' "
            "UNION ALL "
            "SELECT Codigo, Prop.IPTU3 FROM Prop WHERE LTRIM(RTRIM(IPTU3))<>'' AND Prop.IPTU3 <>'00000000' AND Prop.IPTU3 <> '88888888' "
            "UNION ALL "
            "SELECT Codigo, Prop.IPTU4 FROM Prop WHERE LTRIM(RTRIM(IPTU4))<>'' AND Prop.IPTU4 <>'00000000' AND Prop.IPTU4 <> '88888888' "
            "UNION ALL "
            "SELECT Codigo, Prop.IPTU5 FROM Prop WHERE LTRIM(RTRIM(IPTU5))<>'' AND Prop.IPTU5 <>'00000000' AND Prop.IPTU5 <> '88888888' "
            "UNION ALL "
            "SELECT Codigo, Prop.IPTU6 FROM Prop WHERE LTRIM(RTRIM(IPTU6))<>'' AND Prop.IPTU6 <>'00000000' AND Prop.IPTU6 <> '88888888' "
            "UNION ALL "
            "SELECT Codigo, Prop.IPTU7 FROM Prop WHERE LTRIM(RTRIM(IPTU7))<>'' AND Prop.IPTU7 <>'00000000' AND Prop.IPTU7 <> '88888888' "
            "UNION ALL "
            "SELECT Codigo, Prop.IPTU8 FROM Prop WHERE LTRIM(RTRIM(IPTU8))<>'' AND Prop.IPTU8 <>'00000000' AND Prop.IPTU8 <> '88888888' "
            "UNION ALL "
            "SELECT Codigo, Prop.IPTU9 FROM Prop WHERE LTRIM(RTRIM(IPTU9))<>'' AND Prop.IPTU9 <>'00000000' AND Prop.IPTU9 <> '88888888' "
            "UNION ALL "
            "SELECT Codigo, Prop.IPTU10 FROM Prop WHERE LTRIM(RTRIM(IPTU10))<>'' AND Prop.IPTU10 <>'00000000' AND Prop.IPTU10 <> '88888888') AS IPTUs"
        ),
    },
    {
        "tipo_extracao": "Prefeitura__IPTU_Faltante",
        "nome_fonte": "IPTU Faltante",
        "colunas_esperadas": '["Codigo", "IPTU", "Nr Cod IPTU"]',
        "sql_select": (
            "SELECT * FROM ("
            "SELECT Codigo, Prop.IPTU, 1 AS [Nr Cod IPTU] FROM Prop WHERE LTRIM(RTRIM(IPTU))<>'' AND Prop.IPTU <>'00000000' AND Prop.IPTU <> '88888888' AND Prop.IPTU NOT IN (SELECT DISTINCT [Cod IPTUs] FROM [Codigos IPTUs]) "
            "UNION ALL "
            "SELECT Codigo, Prop.IPTU1, 2 AS [Nr Cod IPTU] FROM Prop WHERE LTRIM(RTRIM(IPTU1))<>'' AND Prop.IPTU1 <>'00000000' AND Prop.IPTU1 <> '88888888' AND Prop.IPTU1 NOT IN (SELECT DISTINCT [Cod IPTUs] FROM [Codigos IPTUs]) "
            "UNION ALL "
            "SELECT Codigo, Prop.IPTU2, 3 AS [Nr Cod IPTU] FROM Prop WHERE LTRIM(RTRIM(IPTU2))<>'' AND Prop.IPTU2 <>'00000000' AND Prop.IPTU2 <> '88888888' AND Prop.IPTU2 NOT IN (SELECT DISTINCT [Cod IPTUs] FROM [Codigos IPTUs]) "
            "UNION ALL "
            "SELECT Codigo, Prop.IPTU3, 4 AS [Nr Cod IPTU] FROM Prop WHERE LTRIM(RTRIM(IPTU3))<>'' AND Prop.IPTU3 <>'00000000' AND Prop.IPTU3 <> '88888888' AND Prop.IPTU3 NOT IN (SELECT DISTINCT [Cod IPTUs] FROM [Codigos IPTUs]) "
            "UNION ALL "
            "SELECT Codigo, Prop.IPTU4, 5 AS [Nr Cod IPTU] FROM Prop WHERE LTRIM(RTRIM(IPTU4))<>'' AND Prop.IPTU4 <>'00000000' AND Prop.IPTU4 <> '88888888' AND Prop.IPTU4 NOT IN (SELECT DISTINCT [Cod IPTUs] FROM [Codigos IPTUs]) "
            "UNION ALL "
            "SELECT Codigo, Prop.IPTU5, 6 AS [Nr Cod IPTU] FROM Prop WHERE LTRIM(RTRIM(IPTU5))<>'' AND Prop.IPTU5 <>'00000000' AND Prop.IPTU5 <> '88888888' AND Prop.IPTU5 NOT IN (SELECT DISTINCT [Cod IPTUs] FROM [Codigos IPTUs]) "
            "UNION ALL "
            "SELECT Codigo, Prop.IPTU6, 7 AS [Nr Cod IPTU] FROM Prop WHERE LTRIM(RTRIM(IPTU6))<>'' AND Prop.IPTU6 <>'00000000' AND Prop.IPTU6 <> '88888888' AND Prop.IPTU6 NOT IN (SELECT DISTINCT [Cod IPTUs] FROM [Codigos IPTUs]) "
            "UNION ALL "
            "SELECT Codigo, Prop.IPTU7, 8 AS [Nr Cod IPTU] FROM Prop WHERE LTRIM(RTRIM(IPTU7))<>'' AND Prop.IPTU7 <>'00000000' AND Prop.IPTU7 <> '88888888' AND Prop.IPTU7 NOT IN (SELECT DISTINCT [Cod IPTUs] FROM [Codigos IPTUs]) "
            "UNION ALL "
            "SELECT Codigo, Prop.IPTU8, 9 AS [Nr Cod IPTU] FROM Prop WHERE LTRIM(RTRIM(IPTU8))<>'' AND Prop.IPTU8 <>'00000000' AND Prop.IPTU8 <> '88888888' AND Prop.IPTU8 NOT IN (SELECT DISTINCT [Cod IPTUs] FROM [Codigos IPTUs]) "
            "UNION ALL "
            "SELECT Codigo, Prop.IPTU9, 10 AS [Nr Cod IPTU] FROM Prop WHERE LTRIM(RTRIM(IPTU9))<>'' AND Prop.IPTU9 <>'00000000' AND Prop.IPTU9 <> '88888888' AND Prop.IPTU9 NOT IN (SELECT DISTINCT [Cod IPTUs] FROM [Codigos IPTUs]) "
            "UNION ALL "
            "SELECT Codigo, Prop.IPTU10, 11 AS [Nr Cod IPTU] FROM Prop WHERE LTRIM(RTRIM(IPTU10))<>'' AND Prop.IPTU10 <>'00000000' AND Prop.IPTU10 <> '88888888' AND Prop.IPTU10 NOT IN (SELECT DISTINCT [Cod IPTUs] FROM [Codigos IPTUs])) AS Total"
        ),
    },
    {
        "tipo_extracao": "Bombeiros",
        "nome_fonte": "Bombeiros CBMERJ",
        "colunas_esperadas": '["Codigo", "CBM", "IPTU", "Cidade"]',
        "sql_select": (
            "SELECT DISTINCT Codigo, CBMERJ AS CBM, IPTU, Cidade FROM ("
            "SELECT Codigo, Prop.CBMERJ, Prop.IPTU, Prop.cidade AS Cidade FROM Prop WHERE LTRIM(RTRIM(CBMERJ))<>'' AND Prop.CBMERJ <>'00000000' AND Prop.CBMERJ <> '88888888' "
            "UNION ALL "
            "SELECT Codigo, Prop.CBMERJ1, Prop.IPTU1, Prop.cidade FROM Prop WHERE LTRIM(RTRIM(CBMERJ1))<>'' AND Prop.CBMERJ1 <>'00000000' AND Prop.CBMERJ1 <> '88888888' "
            "UNION ALL "
            "SELECT Codigo, Prop.CBMERJ2, Prop.IPTU2, Prop.cidade FROM Prop WHERE LTRIM(RTRIM(CBMERJ2))<>'' AND Prop.CBMERJ2 <>'00000000' AND Prop.CBMERJ2 <> '88888888' "
            "UNION ALL "
            "SELECT Codigo, Prop.CBMERJ3, Prop.IPTU3, Prop.cidade FROM Prop WHERE LTRIM(RTRIM(CBMERJ3))<>'' AND Prop.CBMERJ3 <>'00000000' AND Prop.CBMERJ3 <> '88888888' "
            "UNION ALL "
            "SELECT Codigo, Prop.CBMERJ4, Prop.IPTU4, Prop.cidade FROM Prop WHERE LTRIM(RTRIM(CBMERJ4))<>'' AND Prop.CBMERJ4 <>'00000000' AND Prop.CBMERJ4 <> '88888888' "
            "UNION ALL "
            "SELECT Codigo, Prop.CBMERJ5, Prop.IPTU5, Prop.cidade FROM Prop WHERE LTRIM(RTRIM(CBMERJ5))<>'' AND Prop.CBMERJ5 <>'00000000' AND Prop.CBMERJ5 <> '88888888' "
            "UNION ALL "
            "SELECT Codigo, Prop.CBMERJ6, Prop.IPTU6, Prop.cidade FROM Prop WHERE LTRIM(RTRIM(CBMERJ6))<>'' AND Prop.CBMERJ6 <>'00000000' AND Prop.CBMERJ6 <> '88888888' "
            "UNION ALL "
            "SELECT Codigo, Prop.CBMERJ7, Prop.IPTU7, Prop.cidade FROM Prop WHERE LTRIM(RTRIM(CBMERJ7))<>'' AND Prop.CBMERJ7 <>'00000000' AND Prop.CBMERJ7 <> '88888888' "
            "UNION ALL "
            "SELECT Codigo, Prop.CBMERJ8, Prop.IPTU8, Prop.cidade FROM Prop WHERE LTRIM(RTRIM(CBMERJ8))<>'' AND Prop.CBMERJ8 <>'00000000' AND Prop.CBMERJ8 <> '88888888' "
            "UNION ALL "
            "SELECT Codigo, Prop.CBMERJ9, Prop.IPTU9, Prop.cidade FROM Prop WHERE LTRIM(RTRIM(CBMERJ9))<>'' AND Prop.CBMERJ9 <>'00000000' AND Prop.CBMERJ9 <> '88888888' "
            "UNION ALL "
            "SELECT Codigo, Prop.CBMERJ10, Prop.IPTU10, Prop.cidade FROM Prop WHERE LTRIM(RTRIM(CBMERJ10))<>'' AND Prop.CBMERJ10 <>'00000000' AND Prop.CBMERJ10 <> '88888888') AS CBMERJs"
        ),
    },
    {
        "tipo_extracao": "Condominios",
        "nome_fonte": "Condomínios",
        "colunas_esperadas": '["Codigo", "LogAdm", "SenAdm", "NomeAdm", "Condominio", "Unidade", "LoginMultiplo"]',
        "sql_select": (
            "SELECT LEFT(TRIM(Tabela.CODIGO),4) AS Codigo, Tabela.LogAdm, Tabela.SenAdm, Tabela.NomeAdm, "
            "Tabela.NomeCond AS Condominio, Tabela.Unidade, Consulta.LoginMultiplo, '' AS Col8, '' AS Col9, "
            "False AS Col10, '' AS Col11, False AS Col12 "
            "FROM Inquilin_New Tabela "
            "LEFT JOIN "
            "(SELECT T.NomeAdm, T.LogAdm, T.SenAdm, Count(T.CodCliente) > 1 AS LoginMultiplo "
            "FROM (SELECT DISTINCT NomeAdm, LogAdm, SenAdm, LEFT(TRIM(Codigo),4) AS CodCliente FROM Inquilin_New) AS T "
            "GROUP BY T.NomeAdm, T.LogAdm, T.SenAdm "
            "HAVING (((T.NomeAdm)<>''))) Consulta "
            "ON Tabela.NomeAdm = Consulta.NomeAdm AND Tabela.LogAdm = Consulta.LogAdm AND Tabela.SenAdm = Consulta.SenAdm "
            "WHERE Tabela.Situa NOT IN ('E', 'F', 'K', 'V') "
            "AND LEFT(TRIM(Tabela.CODIGO),4) NOT IN (SELECT DISTINCT Cliente FROM BoletosCondominios)"
        ),
    },
]

# Inserir fontes (pula se já existir o tipo para esta empresa)
criadas = 0
existentes = 0

for f in FONTES:
    existente = db.query(FonteDado).filter(
        FonteDado.empresa_id == empresa.id,
        FonteDado.tipo_extracao == f["tipo_extracao"],
        FonteDado.ativo == True,
    ).first()

    if existente:
        # Atualizar nome_fonte se mudou
        if existente.nome_fonte != f["nome_fonte"]:
            existente.nome_fonte = f["nome_fonte"]
            print(f"  [UPD] {f['tipo_extracao']} — nome atualizado para '{f['nome_fonte']}'")
        else:
            print(f"  [SKIP] {f['tipo_extracao']} já existe (id={existente.id})")
        existentes += 1
        continue

    nova = FonteDado(
        empresa_id=empresa.id,
        tipo_extracao=f["tipo_extracao"],
        nome_fonte=f["nome_fonte"],
        sql_select=f["sql_select"],
        colunas_esperadas=f["colunas_esperadas"],
    )
    db.add(nova)
    print(f"  [OK] {f['tipo_extracao']} — {f['nome_fonte']}")
    criadas += 1

db.commit()
db.close()

print(f"\nResultado: {criadas} fontes criadas, {existentes} já existiam.")
print("As fontes estão disponíveis em /admin/empresas → W Martins → Fontes de Dados")
