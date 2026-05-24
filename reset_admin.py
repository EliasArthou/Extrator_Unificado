"""
reset_admin.py — Reseta a senha do administrador e exibe no console.
Uso: python reset_admin.py
"""
import secrets
from webapp.models import SessionLocal, Usuario
from webapp.auth import hash_senha

db = SessionLocal()
try:
    admin = db.query(Usuario).filter(Usuario.is_admin == True).first()
    if not admin:
        print("Nenhum admin encontrado no banco.")
    else:
        nova_senha = secrets.token_urlsafe(12)
        admin.senha_hash = hash_senha(nova_senha)
        db.commit()
        print("=" * 50)
        print(f"  Admin: {admin.email}")
        print(f"  Nova senha: {nova_senha}")
        print("  >>> TROQUE ESTA SENHA DEPOIS <<<")
        print("=" * 50)
finally:
    db.close()
