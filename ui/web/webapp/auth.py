"""
auth.py — Autenticação e autorização com JWT + CSRF.
Suporte multi-tenant: limites por empresa + adicional por usuário.
"""

import os
import datetime
import secrets
import time
from collections import defaultdict
from typing import Optional, NamedTuple

from passlib.context import CryptContext
from jose import JWTError, jwt
from fastapi import Depends, HTTPException, Request, status
from fastapi.responses import RedirectResponse
from sqlalchemy.orm import Session

from .models import get_db, Usuario, Empresa, Execucao

# ── Configuração ──────────────────────────────────────────────────────────────
SECRET_KEY = os.getenv("WEBAPP_SECRET_KEY", "chave-secreta-trocar-em-producao-!@#2026")
ALGORITHM = "HS256"
TOKEN_EXPIRE_HOURS = 2  # Reduzido de 12h para 2h

SENHA_MIN_LENGTH = 8  # Mínimo de 8 caracteres

pwd_context = CryptContext(schemes=["bcrypt"], deprecated="auto")

# ── Rate limiting em memória ─────────────────────────────────────────────────
# IP → lista de timestamps de tentativas falhas
_login_attempts: dict[str, list[float]] = defaultdict(list)
_LOGIN_MAX_ATTEMPTS = 5   # máximo de tentativas
_LOGIN_WINDOW_SECS = 300  # janela de 5 minutos


def verificar_rate_limit(ip: str) -> bool:
    """Retorna True se o IP está bloqueado. Limpa tentativas antigas."""
    agora = time.time()
    _login_attempts[ip] = [t for t in _login_attempts[ip] if agora - t < _LOGIN_WINDOW_SECS]
    return len(_login_attempts[ip]) >= _LOGIN_MAX_ATTEMPTS


def registrar_tentativa_falha(ip: str):
    """Registra uma tentativa falha de login."""
    _login_attempts[ip].append(time.time())


def limpar_tentativas(ip: str):
    """Limpa tentativas após login bem-sucedido."""
    _login_attempts.pop(ip, None)


# ── CSRF Token ──────────────────────────────────────────────────────────────
def gerar_csrf_token() -> str:
    """Gera um token CSRF criptograficamente seguro."""
    return secrets.token_urlsafe(32)


def verificar_csrf_token(request_token: str, session_token: str) -> bool:
    """Compara tokens CSRF de forma segura (timing-safe)."""
    if not request_token or not session_token:
        return False
    return secrets.compare_digest(request_token, session_token)


# ── Helpers de senha ──────────────────────────────────────────────────────────

def hash_senha(senha: str) -> str:
    return pwd_context.hash(senha)


def verificar_senha(senha_plain: str, senha_hash: str) -> bool:
    return pwd_context.verify(senha_plain, senha_hash)


# ── JWT ───────────────────────────────────────────────────────────────────────

def criar_token(usuario_id: int, email: str, is_admin: bool = False) -> str:
    payload = {
        "sub": str(usuario_id),
        "email": email,
        "admin": is_admin,
        "iat": datetime.datetime.utcnow(),
        "exp": datetime.datetime.utcnow() + datetime.timedelta(hours=TOKEN_EXPIRE_HOURS),
        "jti": secrets.token_urlsafe(16),  # ID único do token
    }
    return jwt.encode(payload, SECRET_KEY, algorithm=ALGORITHM)


def decodificar_token(token: str) -> Optional[dict]:
    try:
        return jwt.decode(token, SECRET_KEY, algorithms=[ALGORITHM])
    except JWTError:
        return None


# ── Dependência FastAPI ───────────────────────────────────────────────────────

def _is_localhost(request: Request) -> bool:
    """Verifica se a requisição vem de localhost (127.0.0.1 ou ::1)."""
    client_ip = request.client.host if request.client else ""
    return client_ip in ("127.0.0.1", "::1", "localhost")


def get_usuario_logado(request: Request, db: Session = Depends(get_db)) -> Usuario:
    """Extrai usuário do cookie de sessão. Redireciona ao login se inválido."""
    token = request.cookies.get("access_token")
    if not token:
        raise HTTPException(
            status_code=status.HTTP_307_TEMPORARY_REDIRECT,
            headers={"Location": "/login"},
        )

    payload = decodificar_token(token)
    if payload is None:
        raise HTTPException(
            status_code=status.HTTP_307_TEMPORARY_REDIRECT,
            headers={"Location": "/login"},
        )

    usuario = db.query(Usuario).filter(Usuario.id == int(payload["sub"])).first()
    if not usuario or not usuario.ativo or usuario.excluido:
        raise HTTPException(
            status_code=status.HTTP_307_TEMPORARY_REDIRECT,
            headers={"Location": "/login"},
        )

    # Usuário restrito a localhost — bloquear acesso remoto
    if usuario.somente_local and not _is_localhost(request):
        raise HTTPException(
            status_code=status.HTTP_307_TEMPORARY_REDIRECT,
            headers={"Location": "/login"},
        )

    return usuario


def get_admin(usuario: Usuario = Depends(get_usuario_logado)) -> Usuario:
    """Garante que o usuário logado é admin."""
    if not usuario.is_admin:
        raise HTTPException(status_code=403, detail="Acesso negado: somente administradores.")
    return usuario


# ── Helpers de uso ────────────────────────────────────────────────────────────

class StatusLimite(NamedTuple):
    """Resultado da verificação de limites."""
    pode_executar: bool
    usando_adicional: bool      # True = empresa esgotou, usando extra do usuário
    uso_empresa_hoje: int       # Total de extrações da empresa hoje
    limite_empresa: int         # Limite diário da empresa
    uso_adicional_hoje: int     # Extrações adicionais deste usuário hoje
    limite_adicional: int       # Limite adicional do usuário
    mensagem: str               # Mensagem para exibir ao usuário


def contar_extracoes_hoje(db: Session, usuario_id: int) -> int:
    """Retorna quantas extrações o usuário fez hoje (compatibilidade)."""
    hoje = datetime.date.today()
    inicio = datetime.datetime.combine(hoje, datetime.time.min)
    return (
        db.query(Execucao)
        .filter(
            Execucao.usuario_id == usuario_id,
            Execucao.iniciado_em >= inicio,
        )
        .count()
    )


def contar_extracoes_empresa_hoje(db: Session, empresa_id: int) -> int:
    """Retorna total de extrações da empresa hoje (todos os usuários)."""
    hoje = datetime.date.today()
    inicio = datetime.datetime.combine(hoje, datetime.time.min)
    return (
        db.query(Execucao)
        .filter(
            Execucao.empresa_id == empresa_id,
            Execucao.iniciado_em >= inicio,
        )
        .count()
    )


def contar_adicionais_usuario_hoje(db: Session, usuario_id: int) -> int:
    """Retorna quantas extrações adicionais o usuário usou hoje."""
    hoje = datetime.date.today()
    inicio = datetime.datetime.combine(hoje, datetime.time.min)
    return (
        db.query(Execucao)
        .filter(
            Execucao.usuario_id == usuario_id,
            Execucao.usou_adicional == True,
            Execucao.iniciado_em >= inicio,
        )
        .count()
    )


def verificar_limite(db: Session, usuario: Usuario) -> StatusLimite:
    """Verifica se o usuário pode executar, considerando limites empresa + adicional.

    Lógica:
    1. Se não tem empresa → usa limite_diario legado do próprio usuário
    2. Se empresa ainda tem saldo → OK, consome da empresa
    3. Se empresa esgotou e usuário tem adicional → OK com aviso de cobrança extra
    4. Se empresa esgotou e sem adicional → bloqueado
    """
    # Sem empresa = admin global ou legado → usa limite_diario do usuário
    if not usuario.empresa_id:
        uso = contar_extracoes_hoje(db, usuario.id)
        limite = usuario.limite_diario
        pode = uso < limite
        return StatusLimite(
            pode_executar=pode,
            usando_adicional=False,
            uso_empresa_hoje=uso,
            limite_empresa=limite,
            uso_adicional_hoje=0,
            limite_adicional=0,
            mensagem="" if pode else "Limite diário atingido.",
        )

    # Com empresa → verificar franquia da empresa
    uso_empresa = contar_extracoes_empresa_hoje(db, usuario.empresa_id)
    empresa = db.query(Empresa).filter(Empresa.id == usuario.empresa_id).first()

    if not empresa or not empresa.ativo:
        return StatusLimite(
            pode_executar=False, usando_adicional=False,
            uso_empresa_hoje=0, limite_empresa=0,
            uso_adicional_hoje=0, limite_adicional=0,
            mensagem="Empresa inativa.",
        )

    limite_empresa = empresa.limite_diario

    # Empresa ainda tem saldo
    if uso_empresa < limite_empresa:
        return StatusLimite(
            pode_executar=True,
            usando_adicional=False,
            uso_empresa_hoje=uso_empresa,
            limite_empresa=limite_empresa,
            uso_adicional_hoje=0,
            limite_adicional=usuario.limite_adicional,
            mensagem="",
        )

    # Empresa esgotou → verificar adicional do usuário
    uso_adicional = contar_adicionais_usuario_hoje(db, usuario.id)
    limite_adicional = usuario.limite_adicional

    if limite_adicional > 0 and uso_adicional < limite_adicional:
        restante = limite_adicional - uso_adicional
        return StatusLimite(
            pode_executar=True,
            usando_adicional=True,
            uso_empresa_hoje=uso_empresa,
            limite_empresa=limite_empresa,
            uso_adicional_hoje=uso_adicional,
            limite_adicional=limite_adicional,
            mensagem=f"Limite da empresa atingido. Usando extrações adicionais ({restante} restantes). Cobrança adicional será aplicada.",
        )

    # Sem saldo nenhum
    return StatusLimite(
        pode_executar=False,
        usando_adicional=False,
        uso_empresa_hoje=uso_empresa,
        limite_empresa=limite_empresa,
        uso_adicional_hoje=uso_adicional,
        limite_adicional=limite_adicional,
        mensagem="Limite diário da empresa atingido." + (
            "" if limite_adicional == 0 else " Extrações adicionais também esgotadas."
        ),
    )
