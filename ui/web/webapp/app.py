"""
app.py — Aplicação principal FastAPI do ExtratorUnificado Web.

Para rodar:
    cd ExtratorUnificado
    uvicorn webapp.app:app --host 0.0.0.0 --port 8000 --reload
"""

import os
import secrets
import sys
from pathlib import Path

from fastapi import FastAPI, Request
from fastapi.staticfiles import StaticFiles
from fastapi.responses import RedirectResponse
from starlette.middleware.base import BaseHTTPMiddleware

# Garante que a raiz do projeto está no path
_project_root = str(Path(__file__).resolve().parent.parent)
if _project_root not in sys.path:
    sys.path.insert(0, _project_root)

from dotenv import load_dotenv
load_dotenv()

from .models import init_db, SessionLocal, Usuario
from .auth import hash_senha
from .routes.auth_routes import router as auth_router
from .routes.extraction import router as extraction_router
from .routes.admin_routes import router as admin_router


# ── Security headers middleware ─────────────────────────────────────────────

class SecurityHeadersMiddleware(BaseHTTPMiddleware):
    async def dispatch(self, request: Request, call_next):
        response = await call_next(request)
        response.headers["X-Content-Type-Options"] = "nosniff"
        response.headers["X-Frame-Options"] = "DENY"
        response.headers["X-XSS-Protection"] = "1; mode=block"
        response.headers["Referrer-Policy"] = "strict-origin-when-cross-origin"
        response.headers["Permissions-Policy"] = "camera=(), microphone=(), geolocation=()"
        # CSP: permitir Bootstrap CDN + inline styles (Bootstrap precisa)
        response.headers["Content-Security-Policy"] = (
            "default-src 'self'; "
            "script-src 'self' 'unsafe-inline'; "
            "style-src 'self' 'unsafe-inline'; "
            "font-src 'self'; "
            "img-src 'self' data:; "
            "connect-src 'self'; "
            "frame-ancestors 'none'"
        )
        return response


# ── Cria app ──────────────────────────────────────────────────────────────────

app = FastAPI(
    title="ExtratorUnificado Web",
    description="Interface web para extração de boletos (IPTU, Bombeiros, Condomínios)",
    version="1.0.0",
    docs_url=None,      # Desabilitar Swagger UI em produção
    redoc_url=None,      # Desabilitar ReDoc em produção
    openapi_url=None,    # Desabilitar OpenAPI schema
)

# Adicionar middleware de segurança
app.add_middleware(SecurityHeadersMiddleware)

# Servir arquivos estáticos (CSS, JS)
_static_dir = str(Path(__file__).resolve().parent / "static")
app.mount("/static", StaticFiles(directory=_static_dir), name="static")

# Registrar rotas
app.include_router(auth_router)
app.include_router(extraction_router)
app.include_router(admin_router)


# ── Startup ───────────────────────────────────────────────────────────────────

@app.on_event("startup")
def startup():
    init_db()
    try:
        _criar_admin_padrao()
    except Exception as e:
        print(f"[WEBAPP] AVISO: Não foi possível verificar admin padrão: {e}")
        print("[WEBAPP]   Provavelmente faltam colunas no banco. Execute os SQLs indicados acima.")


def _criar_admin_padrao():
    """Cria usuário admin padrão se não existir nenhum.
    Gera senha aleatória e exibe apenas uma vez no console."""
    db = SessionLocal()
    try:
        admin = db.query(Usuario).filter(Usuario.is_admin == True).first()
        if not admin:
            senha_temp = secrets.token_urlsafe(12)
            admin = Usuario(
                nome="Administrador",
                email="admin@extrator.local",
                senha_hash=hash_senha(senha_temp),
                is_admin=True,
                pode_upload_banco=True,
                ativo=True,
                limite_diario=9999,
            )
            db.add(admin)
            db.commit()
            print("=" * 60)
            print("[WEBAPP] Admin padrão criado!")
            print(f"  E-mail: admin@extrator.local")
            print(f"  Senha:  {senha_temp}")
            print("  >>> TROQUE ESTA SENHA IMEDIATAMENTE <<<")
            print("=" * 60)
    finally:
        db.close()


# ── Handler para redirect em caso de rota não autenticada ─────────────────────

@app.exception_handler(307)
async def redirect_handler(request: Request, exc):
    return RedirectResponse(url="/login")
