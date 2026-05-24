"""
models.py - Modelos SQLAlchemy para autenticacao e tracking de uso.

Banco: PostgreSQL (Neon) — configuração via .env (DATABASE_URL obrigatória).

Arquitetura multi-tenant:
    Empresa → Usuarios → Execucoes
    Empresa → FontesDados (SELECTs configuráveis para o helper)

Limites de extração:
    1. Empresa tem limite_diario (franquia do pacote contratado)
    2. Usuário tem limite_adicional (cobrança extra quando empresa esgota)
    3. Extração consome primeiro da empresa; esgotou, desconta do adicional do usuário

Configuracao via .env:
    DATABASE_URL=postgresql://user:pass@host/dbname?sslmode=require
"""

import datetime
import os
import secrets
from sqlalchemy import (
    create_engine, Column, Integer, String, Boolean,
    DateTime, Float, ForeignKey, Text, Index,
)
from sqlalchemy.orm import declarative_base, sessionmaker, relationship
from dotenv import load_dotenv

load_dotenv()

# DATABASE_URL é OBRIGATÓRIA — sem fallback silencioso pra SQLite
# (o fallback antigo mascarava perda de dados quando .env não era carregada)
DATABASE_URL = os.getenv("DATABASE_URL")

if not DATABASE_URL:
    raise RuntimeError(
        "DATABASE_URL não configurada. Defina na .env a string de conexão do "
        "PostgreSQL (ex: Neon) antes de subir o webapp.\n"
        "Exemplo: DATABASE_URL=postgresql://user:pass@host/dbname?sslmode=require"
    )

_PG_SCHEMA = "extrator"  # Schema dedicado no PostgreSQL

engine = create_engine(
    DATABASE_URL,
    echo=False,
    future=True,
    pool_pre_ping=True,
    pool_size=5,
    max_overflow=10,
)

SessionLocal = sessionmaker(bind=engine, expire_on_commit=False)
Base = declarative_base()

# Quando usando PostgreSQL com schema dedicado, informamos ao SQLAlchemy
# para que todo SQL gerado (DDL e DML) use "extrator.tabela".
# Neon pooler não aceita search_path via options/SET, então o schema
# precisa ser explícito nos metadados.
if _PG_SCHEMA:
    Base.metadata.schema = _PG_SCHEMA


# ── Empresa (tenant) ────────────────────────────────────────────────────────

class Empresa(Base):
    __tablename__ = "empresas"

    id = Column(Integer, primary_key=True, autoincrement=True)
    nome = Column(String(200), nullable=False)
    cnpj = Column(String(20), nullable=True, unique=True)
    ativo = Column(Boolean, default=True, nullable=False, server_default="true")
    chave_api = Column(String(64), unique=True, nullable=False,
                       default=lambda: secrets.token_urlsafe(32))
    limite_diario = Column(Integer, default=500)
    criada_em = Column(DateTime, default=datetime.datetime.utcnow)
    atualizada_em = Column(DateTime, default=datetime.datetime.utcnow,
                           onupdate=datetime.datetime.utcnow)

    # Relacionamentos
    usuarios = relationship("Usuario", back_populates="empresa")
    fontes = relationship("FonteDado", back_populates="empresa",
                          order_by="FonteDado.tipo_extracao")
    execucoes = relationship("Execucao", back_populates="empresa")

    def __repr__(self):
        return f"<Empresa {self.id}: {self.nome}>"


# ── Fonte de Dados (SELECT configurável por empresa) ────────────────────────

class FonteDado(Base):
    __tablename__ = "fontes_dados"

    id = Column(Integer, primary_key=True, autoincrement=True)
    empresa_id = Column(Integer, ForeignKey("empresas.id"), nullable=False)
    tipo_extracao = Column(String(50), nullable=False)  # "Prefeitura__IPTU", "Bombeiros", etc.
    nome_fonte = Column(String(120), nullable=False)     # Nome amigável: "IPTU Completo"
    sql_select = Column(Text, nullable=False)            # SELECT SQL para pyodbc
    colunas_esperadas = Column(Text, default="[]")       # JSON array: ["Codigo","IPTU"]
    ativo = Column(Boolean, default=True)
    versao = Column(Integer, default=1)                  # Incrementa a cada edição
    criada_em = Column(DateTime, default=datetime.datetime.utcnow)
    atualizada_em = Column(DateTime, default=datetime.datetime.utcnow,
                           onupdate=datetime.datetime.utcnow)

    # Relacionamentos
    empresa = relationship("Empresa", back_populates="fontes")

    # Índice composto: uma fonte por tipo por empresa
    __table_args__ = (
        Index("ix_fonte_empresa_tipo", "empresa_id", "tipo_extracao"),
    )

    def __repr__(self):
        return f"<FonteDado {self.id}: {self.nome_fonte} ({self.tipo_extracao})>"


# ── Usuário ──────────────────────────────────────────────────────────────────

class Usuario(Base):
    __tablename__ = "usuarios"

    id = Column(Integer, primary_key=True, autoincrement=True)
    nome = Column(String(120), nullable=False)
    email = Column(String(200), unique=True, nullable=False)
    senha_hash = Column(String(256), nullable=False)
    ativo = Column(Boolean, default=True)
    is_admin = Column(Boolean, default=False)
    pode_upload_banco = Column(Boolean, default=False)
    somente_local = Column(Boolean, default=False)
    excluido = Column(Boolean, default=False)

    # Multi-tenant: NULL = admin global, preenchido = usuário da empresa
    empresa_id = Column(Integer, ForeignKey("empresas.id"), nullable=True)

    # Limites: limite_diario é legado (compatibilidade); novo sistema usa empresa
    # limite_adicional = extrações extras pagas (só consome quando empresa esgota)
    limite_diario = Column(Integer, default=200)
    limite_adicional = Column(Integer, default=0)

    criado_em = Column(DateTime, default=datetime.datetime.utcnow)

    # Relacionamentos
    empresa = relationship("Empresa", back_populates="usuarios")
    execucoes = relationship("Execucao", back_populates="usuario")

    __table_args__ = (
        Index("ix_usuario_empresa", "empresa_id"),
    )


# ── Execução ─────────────────────────────────────────────────────────────────

class Execucao(Base):
    __tablename__ = "execucoes"

    id = Column(Integer, primary_key=True, autoincrement=True)
    usuario_id = Column(Integer, ForeignKey("usuarios.id"), nullable=False)
    empresa_id = Column(Integer, ForeignKey("empresas.id"), nullable=True)
    tipo = Column(String(50), nullable=False)
    subtipo = Column(String(50), default="")
    total_registros = Column(Integer, default=0)
    registros_ok = Column(Integer, default=0)
    registros_erro = Column(Integer, default=0)
    status = Column(String(20), default="pendente")
    mensagem = Column(Text, default="")
    iniciado_em = Column(DateTime, default=datetime.datetime.utcnow)
    finalizado_em = Column(DateTime, nullable=True)
    duracao_seg = Column(Float, default=0.0)
    log_texto = Column(Text, default="")
    erros_detalhados = Column(Text, default="[]")

    # Origem: "web" ou "helper" — para relatórios
    origem = Column(String(20), default="web")
    # Se usou adicional do usuário (para cobrança)
    usou_adicional = Column(Boolean, default=False)

    # Relacionamentos
    usuario = relationship("Usuario", back_populates="execucoes")
    empresa = relationship("Empresa", back_populates="execucoes")

    __table_args__ = (
        Index("ix_execucao_empresa", "empresa_id", "iniciado_em"),
        Index("ix_execucao_usuario", "usuario_id", "iniciado_em"),
    )


def init_db():
    """Cria as tabelas se nao existirem e aplica migrações.
    Tolerante a falhas — se o DB não permitir CREATE TABLE, loga e continua."""
    from sqlalchemy import inspect
    insp = inspect(engine)
    tabelas_existentes = insp.get_table_names(schema=_PG_SCHEMA)

    # Tentar criar tabelas que ainda não existem, uma por uma
    for table in Base.metadata.sorted_tables:
        if table.name not in tabelas_existentes:
            try:
                table.create(engine)
                print(f"[WEBAPP] Tabela '{table.name}' criada com sucesso.")
            except Exception as e:
                print(f"[WEBAPP] AVISO: Não foi possível criar tabela '{table.name}': {e}")
                print(f"[WEBAPP]   Crie manualmente no banco.")

    _migrar_colunas()


def _migrar_colunas():
    """Adiciona colunas novas em tabelas existentes (SQLite não faz via create_all).
    Tolerante a falhas — se o DB não permitir ALTER TABLE, loga aviso e continua."""
    from sqlalchemy import inspect, text
    insp = inspect(engine)
    tabelas = insp.get_table_names(schema=_PG_SCHEMA)

    # Prefixo de schema para ALTER TABLE (PostgreSQL usa "extrator.", SQLite não)
    _sp = f"{_PG_SCHEMA}." if _PG_SCHEMA else ""

    # ── Migrações na tabela usuarios ──
    if "usuarios" in tabelas:
        colunas = [c["name"] for c in insp.get_columns("usuarios", schema=_PG_SCHEMA)]
        migracoes_usuarios = {
            "pode_upload_banco": f"ALTER TABLE {_sp}usuarios ADD COLUMN pode_upload_banco BOOLEAN DEFAULT FALSE",
            "somente_local": f"ALTER TABLE {_sp}usuarios ADD COLUMN somente_local BOOLEAN DEFAULT FALSE",
            "excluido": f"ALTER TABLE {_sp}usuarios ADD COLUMN excluido BOOLEAN DEFAULT FALSE",
            "empresa_id": f"ALTER TABLE {_sp}usuarios ADD COLUMN empresa_id INTEGER REFERENCES {_sp}empresas(id)",
            "limite_adicional": f"ALTER TABLE {_sp}usuarios ADD COLUMN limite_adicional INTEGER DEFAULT 0",
        }
        _executar_migracoes("usuarios", colunas, migracoes_usuarios)

    # ── Corrigir registros com ativo=NULL (empresas e fontes) ──
    if "empresas" in tabelas:
        try:
            with engine.begin() as conn:
                result = conn.execute(text(
                    f"UPDATE {_sp}empresas SET ativo = TRUE WHERE ativo IS NULL"
                ))
                if result.rowcount:
                    print(f"[WEBAPP] {result.rowcount} empresa(s) com ativo=NULL corrigida(s) para TRUE.")
        except Exception:
            pass

    # ── Migrações na tabela execucoes ──
    if "execucoes" in tabelas:
        colunas = [c["name"] for c in insp.get_columns("execucoes", schema=_PG_SCHEMA)]
        migracoes_execucoes = {
            "empresa_id": f"ALTER TABLE {_sp}execucoes ADD COLUMN empresa_id INTEGER REFERENCES {_sp}empresas(id)",
            "origem": f"ALTER TABLE {_sp}execucoes ADD COLUMN origem VARCHAR(20) DEFAULT 'web'",
            "usou_adicional": f"ALTER TABLE {_sp}execucoes ADD COLUMN usou_adicional BOOLEAN DEFAULT FALSE",
        }
        _executar_migracoes("execucoes", colunas, migracoes_execucoes)


def _executar_migracoes(tabela, colunas_existentes, migracoes):
    """Executa migrações ALTER TABLE de forma tolerante a falhas."""
    from sqlalchemy import text
    for col_name, ddl in migracoes.items():
        if col_name not in colunas_existentes:
            try:
                with engine.begin() as conn:
                    conn.execute(text(ddl))
                print(f"[WEBAPP] {tabela}.{col_name} adicionada com sucesso.")
            except Exception as e:
                print(f"[WEBAPP] AVISO: Não foi possível adicionar {tabela}.{col_name}: {e}")
                print(f"[WEBAPP]   Execute manualmente: {ddl};")


def get_db():
    """Gerador para dependency injection no FastAPI."""
    db = SessionLocal()
    try:
        yield db
    finally:
        db.close()
