"""
admin_routes.py — Rotas de administração de Empresas e Fontes de Dados.
Acesso restrito a admins globais (is_admin=True).
"""

import json
import secrets

from fastapi import APIRouter, Depends, Request, Form
from fastapi.responses import HTMLResponse, RedirectResponse
from fastapi.templating import Jinja2Templates
from sqlalchemy.orm import Session

from ..models import get_db, Empresa, FonteDado, Usuario
from ..auth import (
    get_admin, gerar_csrf_token, verificar_csrf_token, contar_extracoes_hoje,
)

router = APIRouter(prefix="/admin")

from pathlib import Path as _Path
templates = Jinja2Templates(directory=str(_Path(__file__).resolve().parent.parent / "templates"))


# ═══════════════════════════════════════════════════════════════════════════════
# EMPRESAS
# ═══════════════════════════════════════════════════════════════════════════════

@router.get("/empresas", response_class=HTMLResponse)
async def listar_empresas(
    request: Request,
    usuario: Usuario = Depends(get_admin),
    db: Session = Depends(get_db),
):
    csrf_token = gerar_csrf_token()
    empresas = db.query(Empresa).order_by(Empresa.nome).all()
    uso_hoje = contar_extracoes_hoje(db, usuario.id)

    # Contar usuários e fontes por empresa
    for emp in empresas:
        emp._num_usuarios = db.query(Usuario).filter(
            Usuario.empresa_id == emp.id, Usuario.excluido != True
        ).count()
        emp._num_fontes = db.query(FonteDado).filter(
            FonteDado.empresa_id == emp.id, FonteDado.ativo == True
        ).count()

    response = templates.TemplateResponse(
        request, "admin_empresas.html",
        {
            "usuario": usuario,
            "empresas": empresas,
            "csrf_token": csrf_token,
            "uso_hoje": uso_hoje,
        },
    )
    response.set_cookie(
        key="csrf_token", value=csrf_token,
        httponly=True, samesite="strict", max_age=3600,
    )
    return response


@router.post("/empresas/criar")
async def criar_empresa(
    request: Request,
    nome: str = Form(...),
    cnpj: str = Form(""),
    limite_diario: int = Form(500),
    csrf_token: str = Form(""),
    db: Session = Depends(get_db),
    _admin: Usuario = Depends(get_admin),
):
    cookie_csrf = request.cookies.get("csrf_token", "")
    if not verificar_csrf_token(csrf_token, cookie_csrf):
        return RedirectResponse(url="/admin/empresas", status_code=303)

    # Validar nome único
    existente = db.query(Empresa).filter(Empresa.nome == nome).first()
    if existente:
        # Redirecionar com erro (simplificado)
        return RedirectResponse(url="/admin/empresas", status_code=303)

    nova = Empresa(
        nome=nome,
        cnpj=cnpj.strip() or None,
        limite_diario=max(0, limite_diario),
        chave_api=secrets.token_urlsafe(32),
        ativo=True,
    )
    db.add(nova)
    db.commit()
    return RedirectResponse(url="/admin/empresas", status_code=303)


@router.post("/empresas/{empresa_id}/editar")
async def editar_empresa(
    request: Request,
    empresa_id: int,
    nome: str = Form(...),
    cnpj: str = Form(""),
    limite_diario: int = Form(500),
    csrf_token: str = Form(""),
    db: Session = Depends(get_db),
    _admin: Usuario = Depends(get_admin),
):
    cookie_csrf = request.cookies.get("csrf_token", "")
    if not verificar_csrf_token(csrf_token, cookie_csrf):
        return RedirectResponse(url="/admin/empresas", status_code=303)

    empresa = db.query(Empresa).get(empresa_id)
    if empresa:
        empresa.nome = nome
        empresa.cnpj = cnpj.strip() or None
        empresa.limite_diario = max(0, limite_diario)
        db.commit()
    return RedirectResponse(url="/admin/empresas", status_code=303)


@router.post("/empresas/{empresa_id}/toggle-ativo")
async def toggle_ativo_empresa(
    request: Request,
    empresa_id: int,
    csrf_token: str = Form(""),
    db: Session = Depends(get_db),
    _admin: Usuario = Depends(get_admin),
):
    cookie_csrf = request.cookies.get("csrf_token", "")
    if not verificar_csrf_token(csrf_token, cookie_csrf):
        return RedirectResponse(url="/admin/empresas", status_code=303)

    empresa = db.query(Empresa).get(empresa_id)
    if empresa:
        empresa.ativo = not empresa.ativo if empresa.ativo is not None else False
        db.commit()
    return RedirectResponse(url="/admin/empresas", status_code=303)


@router.post("/empresas/{empresa_id}/regenerar-chave")
async def regenerar_chave(
    request: Request,
    empresa_id: int,
    csrf_token: str = Form(""),
    db: Session = Depends(get_db),
    _admin: Usuario = Depends(get_admin),
):
    cookie_csrf = request.cookies.get("csrf_token", "")
    if not verificar_csrf_token(csrf_token, cookie_csrf):
        return RedirectResponse(url="/admin/empresas", status_code=303)

    empresa = db.query(Empresa).get(empresa_id)
    if empresa:
        empresa.chave_api = secrets.token_urlsafe(32)
        db.commit()
    return RedirectResponse(url=f"/admin/empresas/{empresa_id}/fontes", status_code=303)


# ═══════════════════════════════════════════════════════════════════════════════
# FONTES DE DADOS
# ═══════════════════════════════════════════════════════════════════════════════

@router.get("/empresas/{empresa_id}/fontes", response_class=HTMLResponse)
async def listar_fontes(
    request: Request,
    empresa_id: int,
    usuario: Usuario = Depends(get_admin),
    db: Session = Depends(get_db),
):
    empresa = db.query(Empresa).get(empresa_id)
    if not empresa:
        return RedirectResponse(url="/admin/empresas", status_code=303)

    csrf_token = gerar_csrf_token()
    fontes = db.query(FonteDado).filter(
        FonteDado.empresa_id == empresa_id,
        FonteDado.ativo == True,
    ).order_by(FonteDado.tipo_extracao).all()

    uso_hoje = contar_extracoes_hoje(db, usuario.id)

    response = templates.TemplateResponse(
        request, "admin_fontes.html",
        {
            "usuario": usuario,
            "empresa": empresa,
            "fontes": fontes,
            "csrf_token": csrf_token,
            "uso_hoje": uso_hoje,
        },
    )
    response.set_cookie(
        key="csrf_token", value=csrf_token,
        httponly=True, samesite="strict", max_age=3600,
    )
    return response


@router.post("/empresas/{empresa_id}/fontes/criar")
async def criar_fonte(
    request: Request,
    empresa_id: int,
    tipo_extracao: str = Form(...),
    nome_fonte: str = Form(...),
    sql_select: str = Form(...),
    colunas_esperadas: str = Form("[]"),
    csrf_token: str = Form(""),
    db: Session = Depends(get_db),
    _admin: Usuario = Depends(get_admin),
):
    cookie_csrf = request.cookies.get("csrf_token", "")
    if not verificar_csrf_token(csrf_token, cookie_csrf):
        return RedirectResponse(url=f"/admin/empresas/{empresa_id}/fontes", status_code=303)

    empresa = db.query(Empresa).get(empresa_id)
    if not empresa:
        return RedirectResponse(url="/admin/empresas", status_code=303)

    # Validar JSON das colunas
    try:
        cols = json.loads(colunas_esperadas)
        if not isinstance(cols, list):
            colunas_esperadas = "[]"
    except (json.JSONDecodeError, TypeError):
        colunas_esperadas = "[]"

    nova = FonteDado(
        empresa_id=empresa_id,
        tipo_extracao=tipo_extracao.strip(),
        nome_fonte=nome_fonte.strip(),
        sql_select=sql_select.strip(),
        colunas_esperadas=colunas_esperadas,
    )
    db.add(nova)
    db.commit()
    return RedirectResponse(url=f"/admin/empresas/{empresa_id}/fontes", status_code=303)


@router.post("/empresas/{empresa_id}/fontes/{fonte_id}/editar")
async def editar_fonte(
    request: Request,
    empresa_id: int,
    fonte_id: int,
    nome_fonte: str = Form(...),
    sql_select: str = Form(...),
    colunas_esperadas: str = Form("[]"),
    csrf_token: str = Form(""),
    db: Session = Depends(get_db),
    _admin: Usuario = Depends(get_admin),
):
    cookie_csrf = request.cookies.get("csrf_token", "")
    if not verificar_csrf_token(csrf_token, cookie_csrf):
        return RedirectResponse(url=f"/admin/empresas/{empresa_id}/fontes", status_code=303)

    fonte = db.query(FonteDado).filter(
        FonteDado.id == fonte_id,
        FonteDado.empresa_id == empresa_id,
    ).first()

    if fonte:
        fonte.nome_fonte = nome_fonte.strip()
        fonte.sql_select = sql_select.strip()
        try:
            cols = json.loads(colunas_esperadas)
            if isinstance(cols, list):
                fonte.colunas_esperadas = colunas_esperadas
        except (json.JSONDecodeError, TypeError):
            pass
        fonte.versao += 1  # Incrementar versão para o helper detectar mudanças
        db.commit()
    return RedirectResponse(url=f"/admin/empresas/{empresa_id}/fontes", status_code=303)


@router.post("/empresas/{empresa_id}/fontes/{fonte_id}/excluir")
async def excluir_fonte(
    request: Request,
    empresa_id: int,
    fonte_id: int,
    csrf_token: str = Form(""),
    db: Session = Depends(get_db),
    _admin: Usuario = Depends(get_admin),
):
    cookie_csrf = request.cookies.get("csrf_token", "")
    if not verificar_csrf_token(csrf_token, cookie_csrf):
        return RedirectResponse(url=f"/admin/empresas/{empresa_id}/fontes", status_code=303)

    fonte = db.query(FonteDado).filter(
        FonteDado.id == fonte_id,
        FonteDado.empresa_id == empresa_id,
    ).first()

    if fonte:
        fonte.ativo = False  # Soft delete
        db.commit()
    return RedirectResponse(url=f"/admin/empresas/{empresa_id}/fontes", status_code=303)
