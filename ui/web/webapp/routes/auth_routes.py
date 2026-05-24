"""
auth_routes.py - Rotas de login, logout e gerenciamento de usuarios.
Inclui: rate limiting, CSRF, validação de senha forte.
"""

from fastapi import APIRouter, Depends, Request, Form
from fastapi.responses import HTMLResponse, RedirectResponse
from fastapi.templating import Jinja2Templates
from sqlalchemy.orm import Session

from ..models import get_db, Usuario, Empresa
from ..auth import (
    hash_senha, verificar_senha, criar_token,
    get_usuario_logado, get_admin, _is_localhost,
    verificar_rate_limit, registrar_tentativa_falha, limpar_tentativas,
    gerar_csrf_token, verificar_csrf_token, SENHA_MIN_LENGTH,
    TOKEN_EXPIRE_HOURS, contar_extracoes_hoje,
)

router = APIRouter()
from pathlib import Path as _Path
templates = Jinja2Templates(directory=str(_Path(__file__).resolve().parent.parent / "templates"))


# -- Login --

@router.get("/login", response_class=HTMLResponse)
async def pagina_login(request: Request):
    csrf_token = gerar_csrf_token()
    response = templates.TemplateResponse(
        request, "login.html", {"erro": "", "csrf_token": csrf_token}
    )
    response.set_cookie(
        key="csrf_token", value=csrf_token,
        httponly=True, samesite="strict", max_age=600,
    )
    return response


@router.post("/login")
async def fazer_login(
    request: Request,
    email: str = Form(...),
    senha: str = Form(...),
    csrf_token: str = Form(""),
    db: Session = Depends(get_db),
):
    # Verificar CSRF
    cookie_csrf = request.cookies.get("csrf_token", "")
    if not verificar_csrf_token(csrf_token, cookie_csrf):
        return templates.TemplateResponse(
            request, "login.html",
            {"erro": "Requisição inválida. Recarregue a página.", "csrf_token": gerar_csrf_token()},
            status_code=403,
        )

    # Rate limiting por IP
    client_ip = request.client.host if request.client else "unknown"
    if verificar_rate_limit(client_ip):
        return templates.TemplateResponse(
            request, "login.html",
            {"erro": "Muitas tentativas. Aguarde 5 minutos.", "csrf_token": gerar_csrf_token()},
            status_code=429,
        )

    usuario = db.query(Usuario).filter(Usuario.email == email).first()
    if not usuario or not verificar_senha(senha, usuario.senha_hash):
        registrar_tentativa_falha(client_ip)
        novo_csrf = gerar_csrf_token()
        response = templates.TemplateResponse(
            request, "login.html",
            {"erro": "E-mail ou senha incorretos.", "csrf_token": novo_csrf},
            status_code=401,
        )
        response.set_cookie(
            key="csrf_token", value=novo_csrf,
            httponly=True, samesite="strict", max_age=600,
        )
        return response

    if usuario.excluido or not usuario.ativo:
        novo_csrf = gerar_csrf_token()
        response = templates.TemplateResponse(
            request, "login.html",
            {"erro": "Conta desativada. Contate o administrador.", "csrf_token": novo_csrf},
            status_code=403,
        )
        response.set_cookie(
            key="csrf_token", value=novo_csrf,
            httponly=True, samesite="strict", max_age=600,
        )
        return response

    # Usuário restrito a localhost
    if usuario.somente_local and not _is_localhost(request):
        novo_csrf = gerar_csrf_token()
        response = templates.TemplateResponse(
            request, "login.html",
            {"erro": "Este usuário só pode acessar via localhost.", "csrf_token": novo_csrf},
            status_code=403,
        )
        response.set_cookie(
            key="csrf_token", value=novo_csrf,
            httponly=True, samesite="strict", max_age=600,
        )
        return response

    limpar_tentativas(client_ip)
    token = criar_token(usuario.id, usuario.email, usuario.is_admin)
    response = RedirectResponse(url="/", status_code=303)
    response.set_cookie(
        key="access_token", value=token,
        httponly=True,
        max_age=TOKEN_EXPIRE_HOURS * 3600,
        samesite="lax",
        secure=True,  # Só envia via HTTPS
    )
    response.delete_cookie("csrf_token")
    return response


@router.get("/logout")
async def fazer_logout():
    response = RedirectResponse(url="/login", status_code=303)
    response.delete_cookie("access_token")
    response.delete_cookie("csrf_token")
    return response


# -- Gerenciamento de usuarios (admin) --

@router.get("/admin/usuarios", response_class=HTMLResponse)
async def listar_usuarios(
    request: Request,
    usuario: Usuario = Depends(get_admin),
    db: Session = Depends(get_db),
):
    csrf_token = gerar_csrf_token()
    usuarios = db.query(Usuario).filter(Usuario.excluido != True).order_by(Usuario.nome).all()
    empresas = db.query(Empresa).order_by(Empresa.nome).all()
    uso_hoje = contar_extracoes_hoje(db, usuario.id)
    response = templates.TemplateResponse(
        request, "admin_usuarios.html",
        {"usuario": usuario, "usuarios": usuarios, "empresas": empresas,
         "csrf_token": csrf_token, "uso_hoje": uso_hoje},
    )
    response.set_cookie(
        key="csrf_token", value=csrf_token,
        httponly=True, samesite="strict", max_age=3600,
    )
    return response


@router.post("/admin/usuarios/criar")
async def criar_usuario(
    request: Request,
    nome: str = Form(...),
    email: str = Form(...),
    senha: str = Form(...),
    is_admin: bool = Form(False),
    pode_upload_banco: bool = Form(False),
    somente_local: bool = Form(False),
    limite_diario: int = Form(200),
    empresa_id: int = Form(0),
    csrf_token: str = Form(""),
    db: Session = Depends(get_db),
    _admin: Usuario = Depends(get_admin),
):
    # Verificar CSRF
    cookie_csrf = request.cookies.get("csrf_token", "")
    if not verificar_csrf_token(csrf_token, cookie_csrf):
        return RedirectResponse(url="/admin/usuarios", status_code=303)

    # Helper para recarregar a página com erro
    def _responder_erro(msg):
        usuarios = db.query(Usuario).filter(Usuario.excluido != True).order_by(Usuario.nome).all()
        empresas = db.query(Empresa).order_by(Empresa.nome).all()
        novo_csrf = gerar_csrf_token()
        response = templates.TemplateResponse(
            request, "admin_usuarios.html",
            {"usuario": _admin, "usuarios": usuarios, "empresas": empresas,
             "erro": msg, "csrf_token": novo_csrf, "uso_hoje": contar_extracoes_hoje(db, _admin.id)},
        )
        response.set_cookie(
            key="csrf_token", value=novo_csrf,
            httponly=True, samesite="strict", max_age=3600,
        )
        return response

    # Validar senha forte
    if len(senha) < SENHA_MIN_LENGTH:
        return _responder_erro(f"Senha deve ter no mínimo {SENHA_MIN_LENGTH} caracteres.")

    existente = db.query(Usuario).filter(Usuario.email == email).first()
    if existente:
        return _responder_erro(f"E-mail {email} já cadastrado.")

    novo = Usuario(
        nome=nome,
        email=email,
        senha_hash=hash_senha(senha),
        is_admin=is_admin,
        pode_upload_banco=pode_upload_banco or is_admin,
        somente_local=somente_local,
        limite_diario=limite_diario,
        empresa_id=empresa_id if empresa_id else None,
    )
    db.add(novo)
    db.commit()
    return RedirectResponse(url="/admin/usuarios", status_code=303)


@router.post("/admin/usuarios/{user_id}/toggle")
async def toggle_usuario(
    request: Request,
    user_id: int,
    csrf_token: str = Form(""),
    db: Session = Depends(get_db),
    _admin: Usuario = Depends(get_admin),
):
    # Verificar CSRF
    cookie_csrf = request.cookies.get("csrf_token", "")
    if not verificar_csrf_token(csrf_token, cookie_csrf):
        return RedirectResponse(url="/admin/usuarios", status_code=303)

    user = db.query(Usuario).get(user_id)
    if user and user.id != _admin.id:
        user.ativo = not user.ativo
        db.commit()
    return RedirectResponse(url="/admin/usuarios", status_code=303)


@router.post("/admin/usuarios/{user_id}/toggle-admin")
async def toggle_admin(
    request: Request,
    user_id: int,
    csrf_token: str = Form(""),
    db: Session = Depends(get_db),
    _admin: Usuario = Depends(get_admin),
):
    """Alterna permissão de administrador para o usuário."""
    cookie_csrf = request.cookies.get("csrf_token", "")
    if not verificar_csrf_token(csrf_token, cookie_csrf):
        return RedirectResponse(url="/admin/usuarios", status_code=303)

    user = db.query(Usuario).get(user_id)
    if user and user.id != _admin.id:  # Não altera a si mesmo
        user.is_admin = not user.is_admin
        if user.is_admin:
            user.pode_upload_banco = True  # Admin sempre tem acesso ao banco
        db.commit()
    return RedirectResponse(url="/admin/usuarios", status_code=303)


@router.post("/admin/usuarios/{user_id}/toggle-banco")
async def toggle_upload_banco(
    request: Request,
    user_id: int,
    csrf_token: str = Form(""),
    db: Session = Depends(get_db),
    _admin: Usuario = Depends(get_admin),
):
    """Alterna permissão de upload de banco (.WMB) para o usuário."""
    cookie_csrf = request.cookies.get("csrf_token", "")
    if not verificar_csrf_token(csrf_token, cookie_csrf):
        return RedirectResponse(url="/admin/usuarios", status_code=303)

    user = db.query(Usuario).get(user_id)
    if user and not user.is_admin:  # Não altera admin (sempre tem acesso)
        user.pode_upload_banco = not user.pode_upload_banco
        db.commit()
    return RedirectResponse(url="/admin/usuarios", status_code=303)


@router.post("/admin/usuarios/{user_id}/toggle-local")
async def toggle_somente_local(
    request: Request,
    user_id: int,
    csrf_token: str = Form(""),
    db: Session = Depends(get_db),
    _admin: Usuario = Depends(get_admin),
):
    """Alterna restrição de acesso somente via localhost."""
    cookie_csrf = request.cookies.get("csrf_token", "")
    if not verificar_csrf_token(csrf_token, cookie_csrf):
        return RedirectResponse(url="/admin/usuarios", status_code=303)

    user = db.query(Usuario).get(user_id)
    if user and user.id != _admin.id:  # Não altera o próprio admin
        user.somente_local = not user.somente_local
        db.commit()
    return RedirectResponse(url="/admin/usuarios", status_code=303)


@router.post("/admin/usuarios/{user_id}/limite")
async def alterar_limite(
    request: Request,
    user_id: int,
    limite_diario: int = Form(...),
    csrf_token: str = Form(""),
    db: Session = Depends(get_db),
    _admin: Usuario = Depends(get_admin),
):
    # Verificar CSRF
    cookie_csrf = request.cookies.get("csrf_token", "")
    if not verificar_csrf_token(csrf_token, cookie_csrf):
        return RedirectResponse(url="/admin/usuarios", status_code=303)

    user = db.query(Usuario).get(user_id)
    if user:
        user.limite_diario = max(0, limite_diario)  # Evitar valores negativos
        db.commit()
    return RedirectResponse(url="/admin/usuarios", status_code=303)


@router.post("/admin/usuarios/{user_id}/empresa")
async def alterar_empresa(
    request: Request,
    user_id: int,
    empresa_id: int = Form(0),
    csrf_token: str = Form(""),
    db: Session = Depends(get_db),
    _admin: Usuario = Depends(get_admin),
):
    """Altera a empresa vinculada ao usuário."""
    cookie_csrf = request.cookies.get("csrf_token", "")
    if not verificar_csrf_token(csrf_token, cookie_csrf):
        return RedirectResponse(url="/admin/usuarios", status_code=303)

    user = db.query(Usuario).get(user_id)
    if user:
        user.empresa_id = empresa_id if empresa_id else None
        db.commit()
    return RedirectResponse(url="/admin/usuarios", status_code=303)


@router.post("/admin/usuarios/{user_id}/excluir")
async def excluir_usuario(
    request: Request,
    user_id: int,
    csrf_token: str = Form(""),
    db: Session = Depends(get_db),
    _admin: Usuario = Depends(get_admin),
):
    """Soft delete — marca como excluído, some da listagem mas mantém no banco."""
    cookie_csrf = request.cookies.get("csrf_token", "")
    if not verificar_csrf_token(csrf_token, cookie_csrf):
        return RedirectResponse(url="/admin/usuarios", status_code=303)

    user = db.query(Usuario).get(user_id)
    if user and user.id != _admin.id:  # Não pode excluir a si mesmo
        user.excluido = True
        user.ativo = False  # Também desativa para garantir bloqueio
        db.commit()
    return RedirectResponse(url="/admin/usuarios", status_code=303)
