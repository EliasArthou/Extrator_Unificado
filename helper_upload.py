"""
helper_upload.py — Helper Desktop para ExtratorUnificado Web
Conecta no banco local (.WMB/Access), executa o SELECT apropriado,
e envia o resultado como CSV para o webapp via API.

Requisitos: pip install pyodbc requests
Uso: python helper_upload.py
"""

import csv
import io
import os
import sys
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

try:
    import pyodbc
except ImportError:
    print("ERRO: pyodbc não instalado. Execute: pip install pyodbc")
    sys.exit(1)

try:
    import requests
except ImportError:
    print("ERRO: requests não instalado. Execute: pip install requests")
    sys.exit(1)


# ═══════════════════════════════════════════════════════════════════
# SQL Queries (mesmas do sensiveis.py)
# ═══════════════════════════════════════════════════════════════════

QUERIES = {
    "Prefeitura - IPTU Completo": {
        "sql": (
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
        "colunas": ["Codigo", "IPTU"],
        "tipo_webapp": "Prefeitura",
        "subtipo_webapp": "IPTU",
    },
    "Prefeitura - IPTU Faltante": {
        "sql": (
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
        "colunas": ["Codigo", "IPTU", "Nr Cod IPTU"],
        "tipo_webapp": "Prefeitura",
        "subtipo_webapp": "IPTU",
    },
    "Bombeiros (CBMERJ)": {
        "sql": (
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
        "colunas": ["Codigo", "CBM", "IPTU", "Cidade"],
        "tipo_webapp": "Bombeiros",
        "subtipo_webapp": "",
    },
    "Condomínios": {
        "sql": (
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
        "colunas": ["Codigo", "LogAdm", "SenAdm", "NomeAdm", "Condominio", "Unidade",
                     "LoginMultiplo", "Col8", "Col9", "Col10", "Col11", "Col12"],
        "tipo_webapp": "Condomínios",
        "subtipo_webapp": "",
    },
}


# ═══════════════════════════════════════════════════════════════════
# Configuração padrão
# ═══════════════════════════════════════════════════════════════════

CONFIG_FILE = os.path.join(os.path.expanduser("~"), ".extrator_helper.ini")
DEFAULT_SERVER = "http://127.0.0.1:8000"


def salvar_config(dados: dict):
    """Salva configurações em arquivo INI simples."""
    try:
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            for k, v in dados.items():
                f.write(f"{k}={v}\n")
    except Exception:
        pass


def carregar_config() -> dict:
    """Carrega configurações salvas."""
    config = {"servidor": DEFAULT_SERVER, "ultimo_arquivo": "", "ultimo_tipo": ""}
    try:
        with open(CONFIG_FILE, "r", encoding="utf-8") as f:
            for line in f:
                if "=" in line:
                    k, v = line.strip().split("=", 1)
                    config[k] = v
    except FileNotFoundError:
        pass
    return config


# ═══════════════════════════════════════════════════════════════════
# Funções de banco e upload
# ═══════════════════════════════════════════════════════════════════

def conectar_access(caminho_wmb: str, senha: str) -> pyodbc.Connection:
    """Conecta ao banco Access (.wmb)."""
    conn_str = (
        r"Driver={Microsoft Access Driver (*.mdb, *.accdb)};"
        f"Dbq={caminho_wmb};"
        f"Pwd={senha};"
    )
    return pyodbc.connect(conn_str)


def executar_select(conn: pyodbc.Connection, query_info: dict) -> tuple:
    """Executa SELECT e retorna (colunas, linhas)."""
    cursor = conn.cursor()
    cursor.execute(query_info["sql"])
    rows = cursor.fetchall()
    colunas = query_info["colunas"]
    return colunas, rows


def rows_para_csv(colunas: list, rows: list) -> str:
    """Converte resultado em string CSV."""
    buf = io.StringIO()
    writer = csv.writer(buf)
    writer.writerow(colunas)
    for row in rows:
        writer.writerow([str(v).strip() if v is not None else "" for v in row])
    return buf.getvalue()


def enviar_para_servidor(servidor: str, email: str, senha_web: str,
                         csv_data: str, query_info: dict) -> dict:
    """Faz login no webapp e envia CSV via API."""
    session = requests.Session()

    # 1. Login pra pegar token
    resp = session.post(f"{servidor}/login", data={
        "email": email,
        "senha": senha_web,
        "csrf_token": "",
    }, allow_redirects=False)

    # Pegar token do cookie
    token = session.cookies.get("access_token", "")
    if not token:
        # Tentar pegar do header set-cookie
        for cookie in resp.cookies:
            if cookie.name == "access_token":
                token = cookie.value
                break

    if not token:
        raise Exception("Falha no login. Verifique e-mail e senha do webapp.")

    # 2. Enviar CSV
    csv_bytes = csv_data.encode("utf-8")
    files = {"arquivo_dados": ("dados.csv", csv_bytes, "text/csv")}
    data = {
        "tipo": query_info["tipo_webapp"],
        "subtipo": query_info.get("subtipo_webapp", ""),
        "csrf_token": "",
    }

    resp2 = session.post(
        f"{servidor}/api/upload-dados",
        files=files,
        data=data,
        cookies={"access_token": token},
    )

    if resp2.status_code == 200:
        return resp2.json()
    else:
        raise Exception(f"Erro no upload: {resp2.status_code} — {resp2.text[:200]}")


# ═══════════════════════════════════════════════════════════════════
# Interface Gráfica
# ═══════════════════════════════════════════════════════════════════

class HelperApp:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("ExtratorUnificado — Enviar Dados")
        self.root.geometry("540x620")
        self.root.resizable(False, False)
        self.root.configure(bg="#1a2e44")

        self.config = carregar_config()
        self._build_ui()

    def _build_ui(self):
        # ── Header ──
        header = tk.Frame(self.root, bg="#1a2e44", pady=12)
        header.pack(fill="x")

        tk.Label(header, text="ExtratorUnificado", font=("Segoe UI", 16, "bold"),
                 fg="#fff", bg="#1a2e44").pack()
        tk.Label(header, text="Enviar dados do banco local para o servidor",
                 font=("Segoe UI", 9), fg="#8b9db5", bg="#1a2e44").pack()

        # ── Card principal ──
        card = tk.Frame(self.root, bg="#fff", padx=24, pady=20)
        card.pack(fill="both", expand=True, padx=16, pady=(0, 16))

        row = 0

        # Arquivo WMB
        tk.Label(card, text="Arquivo do Banco (.WMB)", font=("Segoe UI", 9, "bold"),
                 bg="#fff", fg="#2c3e50", anchor="w").grid(row=row, column=0, columnspan=2, sticky="w", pady=(0, 4))
        row += 1

        frame_file = tk.Frame(card, bg="#fff")
        frame_file.grid(row=row, column=0, columnspan=2, sticky="ew", pady=(0, 12))
        frame_file.columnconfigure(0, weight=1)

        self.var_arquivo = tk.StringVar(value=self.config.get("ultimo_arquivo", ""))
        entry_file = tk.Entry(frame_file, textvariable=self.var_arquivo, font=("Segoe UI", 9),
                              state="readonly", readonlybackground="#f8f9fa")
        entry_file.grid(row=0, column=0, sticky="ew", ipady=4)

        btn_browse = tk.Button(frame_file, text="Selecionar...", font=("Segoe UI", 9),
                               bg="#2d5a8c", fg="#fff", relief="flat", padx=12, cursor="hand2",
                               command=self._selecionar_arquivo)
        btn_browse.grid(row=0, column=1, padx=(8, 0))
        row += 1

        # Senha do banco
        tk.Label(card, text="Senha do Banco", font=("Segoe UI", 9, "bold"),
                 bg="#fff", fg="#2c3e50", anchor="w").grid(row=row, column=0, columnspan=2, sticky="w", pady=(0, 4))
        row += 1

        self.var_senha_banco = tk.StringVar()
        tk.Entry(card, textvariable=self.var_senha_banco, show="*", font=("Segoe UI", 9)
                 ).grid(row=row, column=0, columnspan=2, sticky="ew", ipady=4, pady=(0, 12))
        row += 1

        # Tipo de extração
        tk.Label(card, text="Tipo de Extração", font=("Segoe UI", 9, "bold"),
                 bg="#fff", fg="#2c3e50", anchor="w").grid(row=row, column=0, columnspan=2, sticky="w", pady=(0, 4))
        row += 1

        self.var_tipo = tk.StringVar(value=self.config.get("ultimo_tipo", ""))
        combo_tipo = ttk.Combobox(card, textvariable=self.var_tipo, state="readonly",
                                  values=list(QUERIES.keys()), font=("Segoe UI", 9))
        combo_tipo.grid(row=row, column=0, columnspan=2, sticky="ew", ipady=2, pady=(0, 16))
        if not self.var_tipo.get() and QUERIES:
            combo_tipo.current(0)
        row += 1

        # ── Separador ──
        sep = tk.Frame(card, bg="#e2e8f0", height=1)
        sep.grid(row=row, column=0, columnspan=2, sticky="ew", pady=(0, 16))
        row += 1

        # Servidor
        tk.Label(card, text="Servidor", font=("Segoe UI", 9, "bold"),
                 bg="#fff", fg="#2c3e50", anchor="w").grid(row=row, column=0, columnspan=2, sticky="w", pady=(0, 4))
        row += 1

        self.var_servidor = tk.StringVar(value=self.config.get("servidor", DEFAULT_SERVER))
        tk.Entry(card, textvariable=self.var_servidor, font=("Segoe UI", 9)
                 ).grid(row=row, column=0, columnspan=2, sticky="ew", ipady=4, pady=(0, 12))
        row += 1

        # Login do webapp
        tk.Label(card, text="E-mail (webapp)", font=("Segoe UI", 9, "bold"),
                 bg="#fff", fg="#2c3e50", anchor="w").grid(row=row, column=0, sticky="w", pady=(0, 4))
        tk.Label(card, text="Senha (webapp)", font=("Segoe UI", 9, "bold"),
                 bg="#fff", fg="#2c3e50", anchor="w").grid(row=row, column=1, sticky="w", padx=(8, 0), pady=(0, 4))
        row += 1

        self.var_email_web = tk.StringVar(value=self.config.get("email_web", ""))
        self.var_senha_web = tk.StringVar()

        tk.Entry(card, textvariable=self.var_email_web, font=("Segoe UI", 9)
                 ).grid(row=row, column=0, sticky="ew", ipady=4, pady=(0, 16))
        tk.Entry(card, textvariable=self.var_senha_web, show="*", font=("Segoe UI", 9)
                 ).grid(row=row, column=1, sticky="ew", ipady=4, padx=(8, 0), pady=(0, 16))
        row += 1

        card.columnconfigure(0, weight=1)
        card.columnconfigure(1, weight=1)

        # ── Status ──
        self.var_status = tk.StringVar(value="Pronto.")
        tk.Label(card, textvariable=self.var_status, font=("Segoe UI", 8),
                 fg="#6c7a89", bg="#fff", anchor="w").grid(row=row, column=0, columnspan=2, sticky="w", pady=(0, 8))
        row += 1

        # ── Botão Enviar ──
        self.btn_enviar = tk.Button(
            card, text="Extrair e Enviar", font=("Segoe UI", 11, "bold"),
            bg="#2d5a8c", fg="#fff", relief="flat", cursor="hand2",
            padx=20, pady=8, command=self._iniciar_processo,
        )
        self.btn_enviar.grid(row=row, column=0, columnspan=2, sticky="ew", ipady=4)

    def _selecionar_arquivo(self):
        caminho = filedialog.askopenfilename(
            title="Selecionar banco de dados",
            filetypes=[
                ("Banco Scai", "*.wmb;*.WMB"),
                ("Access Database", "*.mdb;*.accdb"),
                ("Todos os arquivos", "*.*"),
            ],
        )
        if caminho:
            self.var_arquivo.set(caminho)

    def _set_status(self, msg, cor="#6c7a89"):
        self.var_status.set(msg)

    def _iniciar_processo(self):
        # Validações
        if not self.var_arquivo.get():
            messagebox.showwarning("Atenção", "Selecione o arquivo do banco (.WMB).")
            return
        if not self.var_senha_banco.get():
            messagebox.showwarning("Atenção", "Digite a senha do banco.")
            return
        if not self.var_tipo.get():
            messagebox.showwarning("Atenção", "Selecione o tipo de extração.")
            return
        if not self.var_email_web.get() or not self.var_senha_web.get():
            messagebox.showwarning("Atenção", "Preencha e-mail e senha do webapp.")
            return

        self.btn_enviar.configure(state="disabled", text="Processando...")
        self._set_status("Conectando ao banco...")

        # Rodar em thread pra não travar a UI
        threading.Thread(target=self._processo_completo, daemon=True).start()

    def _processo_completo(self):
        try:
            query_info = QUERIES[self.var_tipo.get()]

            # 1. Conectar ao banco
            self.root.after(0, self._set_status, "Conectando ao banco...")
            conn = conectar_access(self.var_arquivo.get(), self.var_senha_banco.get())

            # 2. Executar SELECT
            self.root.after(0, self._set_status, "Executando consulta...")
            colunas, rows = executar_select(conn, query_info)
            conn.close()

            total = len(rows)
            self.root.after(0, self._set_status, f"Consulta OK: {total} registros. Preparando CSV...")

            if total == 0:
                self.root.after(0, lambda: messagebox.showinfo("Resultado", "A consulta retornou 0 registros. Nada a enviar."))
                self.root.after(0, self._resetar_botao)
                return

            # 3. Converter pra CSV
            csv_data = rows_para_csv(colunas, rows)
            tamanho_kb = len(csv_data.encode("utf-8")) / 1024
            self.root.after(0, self._set_status, f"{total} registros ({tamanho_kb:.0f} KB). Enviando...")

            # 4. Enviar pro servidor
            resultado = enviar_para_servidor(
                servidor=self.var_servidor.get(),
                email=self.var_email_web.get(),
                senha_web=self.var_senha_web.get(),
                csv_data=csv_data,
                query_info=query_info,
            )

            # 5. Salvar config
            salvar_config({
                "servidor": self.var_servidor.get(),
                "ultimo_arquivo": self.var_arquivo.get(),
                "ultimo_tipo": self.var_tipo.get(),
                "email_web": self.var_email_web.get(),
            })

            # 6. Sucesso
            msg = f"Enviado com sucesso!\n\n{total} registros enviados ao servidor."
            if "job_id" in resultado:
                msg += f"\n\nJob ID: {resultado['job_id']}"
                msg += "\nAcompanhe o progresso no Dashboard do webapp."

            self.root.after(0, lambda: messagebox.showinfo("Sucesso", msg))
            self.root.after(0, self._set_status, f"Enviado! {total} registros.")

        except pyodbc.Error as e:
            erro = str(e)
            if "Not a valid password" in erro or "invalid password" in erro.lower():
                self.root.after(0, lambda: messagebox.showerror("Erro", "Senha do banco incorreta."))
            else:
                self.root.after(0, lambda: messagebox.showerror("Erro de Banco", f"Erro ao acessar o banco:\n\n{erro[:300]}"))
            self.root.after(0, self._set_status, "Erro na conexão com o banco.")

        except requests.ConnectionError:
            self.root.after(0, lambda: messagebox.showerror(
                "Erro de Conexão",
                f"Não foi possível conectar ao servidor:\n{self.var_servidor.get()}\n\nVerifique se o webapp está rodando.",
            ))
            self.root.after(0, self._set_status, "Erro: servidor indisponível.")

        except Exception as e:
            self.root.after(0, lambda: messagebox.showerror("Erro", str(e)[:400]))
            self.root.after(0, self._set_status, "Erro.")

        finally:
            self.root.after(0, self._resetar_botao)

    def _resetar_botao(self):
        self.btn_enviar.configure(state="normal", text="Extrair e Enviar")

    def run(self):
        self.root.mainloop()


if __name__ == "__main__":
    app = HelperApp()
    app.run()
